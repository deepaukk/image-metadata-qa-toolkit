"""
Delete Rejected Assets – Image Metadata QA

Upload/paste a list of filenames (or Asset IDs / base names).
Deletes matching image + JSON pairs from a folder.

Does NOT touch other files. Asks for confirmation before deleting.
"""

import os
import csv
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".heic", ".heif", ".gif", ".bmp", ".tif", ".tiff", ".webp"}


def normalize_base(name: str) -> str:
    name = (name or "").strip().strip('"').strip("'")
    if not name:
        return ""
    name = os.path.basename(name)
    stem, ext = os.path.splitext(name)
    if ext.lower() in IMAGE_EXTS or ext.lower() == ".json":
        return stem
    return name  # already a base / asset id style


class DeleteApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Delete Rejected Assets – Image Metadata QA")
        self.root.geometry("780x560")
        self.root.minsize(640, 480)
        try:
            ttk.Style().theme_use("clam")
        except Exception:
            pass
        self.setup_ui()

    def setup_ui(self):
        top = ttk.LabelFrame(self.root, text=" Folder ", padding=10)
        top.pack(fill="x", padx=12, pady=8)

        ttk.Label(top, text="Folder containing images + JSON:").grid(row=0, column=0, sticky="w")
        self.folder_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.folder_var, width=55).grid(row=0, column=1, padx=5, sticky="ew")
        ttk.Button(top, text="Browse…", command=self.browse_folder).grid(row=0, column=2)
        self.recursive_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(top, text="Include subfolders", variable=self.recursive_var).grid(
            row=1, column=1, sticky="w", pady=4
        )
        top.columnconfigure(1, weight=1)

        mid = ttk.LabelFrame(
            self.root,
            text=" Rejected filenames (one per line — filename, base name, or Asset ID) ",
            padding=10,
        )
        mid.pack(fill="both", expand=True, padx=12, pady=4)

        btn_row = ttk.Frame(mid)
        btn_row.pack(fill="x", pady=(0, 6))
        ttk.Button(btn_row, text="Load from CSV / Excel / TXT", command=self.load_list_file).pack(side="left")
        ttk.Button(btn_row, text="Clear", command=lambda: self.list_text.delete("1.0", tk.END)).pack(
            side="left", padx=6
        )

        self.list_text = tk.Text(mid, wrap="word", font=("Consolas", 10), height=12)
        sb = ttk.Scrollbar(mid, orient="vertical", command=self.list_text.yview)
        self.list_text.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.list_text.pack(fill="both", expand=True)

        ttk.Label(
            self.root,
            text="Tip: You can paste the Filename column from the QA report (FAIL rows only).",
            foreground="#555",
        ).pack(anchor="w", padx=14)

        actions = ttk.Frame(self.root)
        actions.pack(fill="x", padx=12, pady=8)
        ttk.Button(actions, text="1. Preview matches", command=self.preview).pack(side="left")
        ttk.Button(actions, text="2. Delete matched files", command=self.delete).pack(side="left", padx=8)
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(actions, textvariable=self.status_var).pack(side="left", padx=10)

        log_frame = ttk.LabelFrame(self.root, text=" Log ", padding=8)
        log_frame.pack(fill="both", expand=True, padx=12, pady=8)
        self.log_text = tk.Text(log_frame, wrap="word", font=("Consolas", 9), height=10)
        lsb = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=lsb.set)
        lsb.pack(side="right", fill="y")
        self.log_text.pack(fill="both", expand=True)

    def log(self, msg):
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def browse_folder(self):
        p = filedialog.askdirectory(title="Select folder")
        if p:
            self.folder_var.set(p)

    def load_list_file(self):
        path = filedialog.askopenfilename(
            title="Select list file",
            filetypes=[
                ("CSV / Excel / Text", "*.csv *.xlsx *.xls *.txt"),
                ("All", "*.*"),
            ],
        )
        if not path:
            return
        names = []
        try:
            if path.lower().endswith((".xlsx", ".xls")):
                import pandas as pd

                df = pd.read_excel(path)
                # Prefer Filename column, else first column
                col = None
                for c in df.columns:
                    if str(c).lower().replace(" ", "") in ("filename", "file_name", "file"):
                        col = c
                        break
                if col is None:
                    col = df.columns[0]
                names = [str(x) for x in df[col].dropna().tolist()]
            elif path.lower().endswith(".csv"):
                with open(path, newline="", encoding="utf-8-sig") as f:
                    reader = csv.reader(f)
                    rows = list(reader)
                if not rows:
                    return
                header = [h.strip().lower() for h in rows[0]]
                idx = 0
                for i, h in enumerate(header):
                    if h.replace(" ", "") in ("filename", "file_name", "file"):
                        idx = i
                        break
                # if header looks like header, skip it
                start = 1 if any(x in header for x in ("filename", "file_name", "asset id", "asset_id")) else 0
                for row in rows[start:]:
                    if len(row) > idx and row[idx].strip():
                        names.append(row[idx].strip())
            else:
                with open(path, encoding="utf-8-sig") as f:
                    names = [ln.strip() for ln in f if ln.strip()]
        except Exception as e:
            messagebox.showerror("Load failed", str(e))
            return

        self.list_text.delete("1.0", tk.END)
        self.list_text.insert("1.0", "\n".join(names))
        self.status_var.set(f"Loaded {len(names)} names")

    def get_bases(self):
        raw = self.list_text.get("1.0", tk.END).splitlines()
        bases = []
        for line in raw:
            b = normalize_base(line)
            if b:
                bases.append(b)
        # unique preserve order
        seen = set()
        out = []
        for b in bases:
            key = b.lower()
            if key not in seen:
                seen.add(key)
                out.append(b)
        return out

    def index_folder(self, folder: Path, recursive: bool):
        """Map lowercase stem -> list of paths (images and jsons)."""
        index = {}
        paths = folder.rglob("*") if recursive else folder.iterdir()
        for p in paths:
            if not p.is_file():
                continue
            ext = p.suffix.lower()
            if ext not in IMAGE_EXTS and ext != ".json":
                continue
            stem = p.stem.lower()
            index.setdefault(stem, []).append(p)
        return index

    def find_matches(self):
        folder = self.folder_var.get().strip()
        if not folder or not Path(folder).is_dir():
            messagebox.showwarning("Folder", "Select a valid folder.")
            return None, None
        bases = self.get_bases()
        if not bases:
            messagebox.showwarning("List", "Paste or load at least one filename.")
            return None, None
        index = self.index_folder(Path(folder), self.recursive_var.get())
        matched_files = []  # paths to delete
        missing = []
        for b in bases:
            key = b.lower()
            # also try without common suffixes? keep simple: exact stem
            found = index.get(key, [])
            if not found:
                missing.append(b)
            else:
                matched_files.extend(found)
        return matched_files, missing

    def preview(self):
        self.log_text.delete("1.0", tk.END)
        matched, missing = self.find_matches()
        if matched is None:
            return
        self.log(f"Would delete {len(matched)} file(s):")
        for p in matched:
            self.log(f"  {p}")
        if missing:
            self.log(f"\nNo files found for {len(missing)} name(s):")
            for m in missing[:50]:
                self.log(f"  {m}")
            if len(missing) > 50:
                self.log(f"  ... and {len(missing) - 50} more")
        self.status_var.set(f"Preview: {len(matched)} files, {len(missing)} missing")

    def delete(self):
        matched, missing = self.find_matches()
        if matched is None:
            return
        if not matched:
            messagebox.showinfo("Nothing to delete", "No matching image/JSON files found.")
            return
        ok = messagebox.askyesno(
            "Confirm delete",
            f"Delete {len(matched)} file(s)?\n\nThis cannot be undone.",
        )
        if not ok:
            return

        self.log_text.delete("1.0", tk.END)
        deleted = 0
        errors = 0
        for p in matched:
            try:
                os.remove(p)
                self.log(f"Deleted: {p}")
                deleted += 1
            except Exception as e:
                self.log(f"ERROR {p}: {e}")
                errors += 1
        if missing:
            self.log(f"\nSkipped (not found): {len(missing)}")
        self.status_var.set(f"Deleted {deleted}, errors {errors}, not found {len(missing)}")
        messagebox.showinfo(
            "Done",
            f"Deleted: {deleted}\nErrors: {errors}\nNot found: {len(missing)}",
        )


if __name__ == "__main__":
    root = tk.Tk()
    DeleteApp(root)
    root.mainloop()
