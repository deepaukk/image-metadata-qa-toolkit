"""
Image Metadata QA – Delivery Packager (GUI)

Delivery folder contains ONLY image + JSON pairs in:
  YYYYMMDD / participant / locale / country / category /

One combined batch metadata Excel is written OUTSIDE the date folder.
No Metadata_*.xlsx inside category folders.

S3 path pattern:
  s3://your-bucket/project/source-data/v1/{YYYYMMDD}/
"""

import os
import re
import shutil
import logging
import threading
import warnings
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime
from logging.handlers import TimedRotatingFileHandler
from logging import StreamHandler

import pandas as pd

warnings.filterwarnings(
    "ignore",
    message="Data Validation extension is not supported and will be removed",
)

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
os.makedirs(SCRIPT_DIR, exist_ok=True)

S3_BASE = "s3://your-bucket/project/source-data/v1"  # configure per project

total_files = 0
skipped_files = 0
skipped_records = []
processed_rows = []  # full metadata rows for batch export


def init_logger():
    logger = logging.getLogger("ImageQA.Packager")
    logger.setLevel(logging.DEBUG)
    if logger.handlers:
        return logger
    log_dir = os.path.join(SCRIPT_DIR, "logs")
    os.makedirs(log_dir, exist_ok=True)
    fmt = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
    fh = TimedRotatingFileHandler(
        os.path.join(log_dir, "Packager.log"), when="midnight", encoding="utf-8"
    )
    fh.setFormatter(fmt)
    logger.addHandler(fh)
    sh = StreamHandler()
    sh.setFormatter(fmt)
    logger.addHandler(sh)
    return logger


logger = init_logger()


def clean_value(value):
    if pd.isna(value):
        return ""
    return str(value).replace("\xa0", "").strip()


def sanitize_path_component(value, fallback="Unknown"):
    value = clean_value(value)
    if not value:
        return fallback
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", value)
    value = value.strip(" .")
    return value or fallback


def normalize_column_name(value):
    value = clean_value(value)
    return value.lower().replace("_", "").replace(" ", "").replace("-", "").strip()


def get_flexible_value(row, target_keys):
    row_columns = {normalize_column_name(c): c for c in row.index}
    for target in target_keys:
        key = normalize_column_name(target)
        if key in row_columns:
            value = clean_value(row[row_columns[key]])
            if value:
                return value
    return ""


def normalize_filename_for_match(filename):
    filename = clean_value(filename)
    if not filename:
        return ""
    return os.path.basename(filename).lower().strip()


def get_filename_stem(filename):
    filename = clean_value(filename)
    if not filename:
        return ""
    stem, _ = os.path.splitext(os.path.basename(filename))
    return stem.strip()


def is_valid_datetime(value):
    if not value:
        return True
    try:
        pd.to_datetime(value)
        return True
    except Exception:
        return False


def build_duplicate_key(filename, participant_id, category):
    return (
        f"{normalize_filename_for_match(filename)}|"
        f"{clean_value(participant_id).lower()}|"
        f"{clean_value(category).lower()}"
    )


def log_skipped(asset_id, filename, participant_id, reason, source_file):
    skipped_records.append({
        "Asset_ID": asset_id,
        "Filename": filename,
        "Participant ID": participant_id,
        "Reason": reason,
        "Source File": source_file,
        "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    })


def find_metadata_file(input_dir):
    for root, dirs, files in os.walk(input_dir):
        for filename in files:
            fl = filename.lower()
            if fl.startswith("metadata_participant_file") and fl.endswith((".xlsx", ".xls")):
                return os.path.join(root, filename)
    return ""


def build_image_index(input_dir):
    image_index = {}
    exts = {".jpg", ".jpeg", ".png", ".heic", ".bmp", ".tif", ".tiff", ".webp"}
    count = 0
    for root, dirs, files in os.walk(input_dir):
        for filename in files:
            if os.path.splitext(filename)[1].lower() not in exts:
                continue
            key = normalize_filename_for_match(filename)
            if not key:
                continue
            count += 1
            if key not in image_index:
                image_index[key] = os.path.join(root, filename)
    return image_index, count


def build_json_index(input_dir):
    json_index = {}
    count = 0
    for root, dirs, files in os.walk(input_dir):
        for filename in files:
            if os.path.splitext(filename)[1].lower() != ".json":
                continue
            key = normalize_filename_for_match(filename)
            if not key:
                continue
            count += 1
            if key not in json_index:
                json_index[key] = os.path.join(root, filename)
    return json_index, count


def process_data(input_dir, output_dir, metadata_file, date_folder, image_index, json_index, log_fn):
    global total_files, skipped_files, processed_rows
    total_files = 0
    skipped_files = 0
    skipped_records.clear()
    processed_rows = []
    seen_keys = set()

    df_meta = pd.read_excel(metadata_file)
    log_fn(f"Metadata rows: {len(df_meta)}")

    for row_number, row in df_meta.iterrows():
        asset_id = filename = participant_id = ""
        try:
            if row.isna().all():
                continue

            asset_id = get_flexible_value(row, ["asset_id", "assetid", "asset id"])
            filename = get_flexible_value(row, ["filename", "file_name", "file name"])
            participant_id = get_flexible_value(
                row, ["participant_id", "participantid", "participant"]
            )
            date_uploaded = get_flexible_value(row, ["dateuploaded", "date_uploaded", "date"])
            locale = get_flexible_value(
                row, ["locale", "languagecode", "language_code", "language"]
            )
            country = get_flexible_value(row, ["country"])
            category = get_flexible_value(row, ["category"])

            if date_uploaded and not is_valid_datetime(date_uploaded):
                log_skipped(asset_id, filename, participant_id, "Invalid date uploaded", metadata_file)
                skipped_files += 1
                continue
            if not participant_id:
                log_skipped(asset_id, filename, participant_id, "Missing Participant ID", metadata_file)
                skipped_files += 1
                continue
            if not filename:
                log_skipped(asset_id, filename, participant_id, "Missing Filename", metadata_file)
                skipped_files += 1
                continue

            normalized_filename = normalize_filename_for_match(filename)
            if not normalized_filename:
                log_skipped(asset_id, filename, participant_id, "Invalid Filename", metadata_file)
                skipped_files += 1
                continue

            dup_key = build_duplicate_key(filename, participant_id, category)
            if dup_key in seen_keys:
                log_skipped(asset_id, filename, participant_id, "Duplicate in this batch", metadata_file)
                skipped_files += 1
                continue

            image_path = image_index.get(normalized_filename)
            if not image_path:
                log_skipped(asset_id, filename, participant_id, "Image file not found", metadata_file)
                skipped_files += 1
                continue

            json_name = normalize_filename_for_match(get_filename_stem(filename) + ".json")
            json_path = json_index.get(json_name)
            if not json_path:
                log_skipped(asset_id, filename, participant_id, "Matching JSON file not found", metadata_file)
                skipped_files += 1
                continue

            # Delivery path: ONLY image + json (no metadata files inside)
            target_dir = os.path.join(
                output_dir,
                date_folder,
                sanitize_path_component(participant_id, "Unknown_Participant"),
                sanitize_path_component(locale, "Unknown_Locale"),
                sanitize_path_component(country, "Unknown_Country"),
                sanitize_path_component(category, "Uncategorized"),
            )
            os.makedirs(target_dir, exist_ok=True)

            dest_img = os.path.join(target_dir, os.path.basename(image_path))
            dest_json = os.path.join(target_dir, os.path.basename(json_path))
            if not os.path.exists(dest_img):
                shutil.copy2(image_path, dest_img)
            if not os.path.exists(dest_json):
                shutil.copy2(json_path, dest_json)

            # Keep row for single batch metadata (written outside date folder)
            processed_rows.append(row)
            seen_keys.add(dup_key)
            total_files += 1

        except Exception as e:
            skipped_files += 1
            log_skipped(
                asset_id, filename, participant_id, f"Unexpected error: {e}", metadata_file
            )
            logger.error(f"Row {row_number + 2}: {e}")


def write_batch_metadata(output_dir, date_folder, s3_path):
    """One combined metadata file OUTSIDE the delivery date folder."""
    if not processed_rows:
        return ""
    df = pd.DataFrame(processed_rows)
    df = df.drop_duplicates()
    df["data_drop_path"] = s3_path
    # Outside date folder: sits next to YYYYMMDD/, not inside it
    path = os.path.join(output_dir, f"{date_folder}_Metadata.xlsx")
    df.to_excel(path, index=False)
    return path


class PackagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Image Metadata QA – Delivery Packager")
        self.root.geometry("880x600")
        self.root.minsize(700, 480)
        try:
            ttk.Style().theme_use("clam")
        except Exception:
            pass
        self.setup_ui()

    def setup_ui(self):
        top = ttk.LabelFrame(self.root, text=" Setup ", padding=12)
        top.pack(fill="x", padx=12, pady=8)

        ttk.Label(top, text="Input folder (images + JSON + metadata):").grid(
            row=0, column=0, sticky="w", pady=4
        )
        self.input_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.input_var, width=55).grid(
            row=0, column=1, padx=5, sticky="ew"
        )
        ttk.Button(top, text="Browse…", command=self.browse_input).grid(row=0, column=2)

        ttk.Label(top, text="Output folder:").grid(row=1, column=0, sticky="w", pady=4)
        self.output_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.output_var, width=55).grid(
            row=1, column=1, padx=5, sticky="ew"
        )
        ttk.Button(top, text="Browse…", command=self.browse_output).grid(row=1, column=2)

        ttk.Label(top, text="Date folder (YYYYMMDD):").grid(row=2, column=0, sticky="w", pady=4)
        self.date_var = tk.StringVar(value=datetime.now().strftime("%Y%m%d"))
        ttk.Entry(top, textvariable=self.date_var, width=20).grid(
            row=2, column=1, sticky="w", padx=5
        )

        ttk.Label(top, text="S3 path (auto):").grid(row=3, column=0, sticky="w", pady=4)
        self.s3_var = tk.StringVar()
        self._update_s3()
        ttk.Entry(top, textvariable=self.s3_var, width=55, state="readonly").grid(
            row=3, column=1, padx=5, sticky="ew"
        )
        self.date_var.trace_add("write", lambda *_: self._update_s3())
        top.columnconfigure(1, weight=1)

        ttk.Label(
            self.root,
            text="Delivery folder = images + JSON only. One YYYYMMDD_Metadata.xlsx is written outside it.",
            foreground="#555",
        ).pack(anchor="w", padx=14)

        btn_row = ttk.Frame(self.root)
        btn_row.pack(fill="x", padx=12, pady=8)
        self.run_btn = ttk.Button(btn_row, text="▶ Run Packaging", command=self.start)
        self.run_btn.pack(side="left")
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(btn_row, textvariable=self.status_var).pack(side="left", padx=12)

        self.progress = ttk.Progressbar(self.root, mode="indeterminate")
        self.progress.pack(fill="x", padx=12, pady=4)

        log_frame = ttk.LabelFrame(self.root, text=" Log ", padding=8)
        log_frame.pack(fill="both", expand=True, padx=12, pady=8)
        self.log_text = tk.Text(log_frame, wrap="word", font=("Consolas", 10), height=16)
        sb = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.log_text.pack(side="left", fill="both", expand=True)
        self.log_text.tag_config("ok", foreground="#1e8e3e")
        self.log_text.tag_config("fail", foreground="#d93025")
        self.log_text.tag_config("info", foreground="#1a73e8")

    def _update_s3(self):
        d = self.date_var.get().strip() or "YYYYMMDD"
        self.s3_var.set(f"{S3_BASE}/{d}/")

    def log(self, msg, tag=None):
        self.log_text.insert(tk.END, msg + "\n", tag)
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def browse_input(self):
        p = filedialog.askdirectory(title="Select input folder")
        if p:
            self.input_var.set(p)

    def browse_output(self):
        p = filedialog.askdirectory(title="Select output folder")
        if p:
            self.output_var.set(p)

    def start(self):
        input_dir = self.input_var.get().strip()
        output_dir = self.output_var.get().strip()
        date_folder = self.date_var.get().strip()

        if not input_dir or not os.path.isdir(input_dir):
            messagebox.showwarning("Input", "Select a valid input folder.")
            return
        if not output_dir:
            messagebox.showwarning("Output", "Select an output folder.")
            return
        if not (len(date_folder) == 8 and date_folder.isdigit()):
            messagebox.showwarning("Date", "Date must be YYYYMMDD (8 digits).")
            return

        self.run_btn.config(state="disabled")
        self.log_text.delete("1.0", tk.END)
        self.progress.start(12)
        threading.Thread(
            target=self.run, args=(input_dir, output_dir, date_folder), daemon=True
        ).start()

    def run(self, input_dir, output_dir, date_folder):
        try:
            os.makedirs(output_dir, exist_ok=True)
            s3_path = f"{S3_BASE}/{date_folder}/"
            self.log(f"Input : {input_dir}", "info")
            self.log(f"Output: {output_dir}", "info")
            self.log(f"Date  : {date_folder}", "info")
            self.log(f"S3    : {s3_path}", "info")
            self.log("=" * 60)

            meta = find_metadata_file(input_dir)
            if not meta:
                self.log("ERROR: Metadata_Participant_File.xlsx not found.", "fail")
                messagebox.showerror(
                    "Missing metadata",
                    "Metadata_Participant_File.xlsx not found under input folder.",
                )
                return
            self.log(f"Source metadata: {meta}", "ok")

            self.status_var.set("Indexing…")
            image_index, img_count = build_image_index(input_dir)
            json_index, json_count = build_json_index(input_dir)
            self.log(f"Images: {img_count}  |  JSON: {json_count}", "info")

            self.status_var.set("Packaging…")
            process_data(
                input_dir,
                output_dir,
                meta,
                date_folder,
                image_index,
                json_index,
                lambda m: self.log(m),
            )

            # One batch metadata file OUTSIDE the date delivery folder
            batch_meta = write_batch_metadata(output_dir, date_folder, s3_path)
            if batch_meta:
                self.log(f"Batch metadata (outside delivery): {batch_meta}", "ok")

            if skipped_records:
                skip_path = os.path.join(output_dir, f"Skipped_Report_{date_folder}.xlsx")
                pd.DataFrame(skipped_records).to_excel(skip_path, index=False)
                self.log(f"Skipped report: {skip_path}", "info")

            self.log("=" * 60)
            self.log(f"Processed: {total_files}  |  Skipped: {skipped_files}", "ok")
            self.log(
                f"Delivery folder (images+json only): {output_dir}/{date_folder}/...",
                "info",
            )
            self.status_var.set("Done")
            messagebox.showinfo(
                "Packaging complete",
                f"Processed: {total_files}\nSkipped: {skipped_files}\n\n"
                f"Delivery: {output_dir}\\{date_folder}\\\n"
                f"Metadata: {date_folder}_Metadata.xlsx (outside date folder)",
            )
        except Exception as e:
            self.log(f"ERROR: {e}", "fail")
            self.status_var.set("Failed")
            messagebox.showerror("Error", str(e))
        finally:
            self.progress.stop()
            self.run_btn.config(state="normal")


if __name__ == "__main__":
    root = tk.Tk()
    PackagerApp(root)
    root.mainloop()
