"""
Delivery Intake Validator – Image Metadata QA
Runs on vendor date folders (flat image + JSON pairs).
Does NOT restructure folders. Does NOT modify the QA app.

Steps:
  1. Pair images and JSON by base filename
  2. Validate JSON against schema — category/sub_category must match schema enum exactly
  3. Check URL required/optional (from taxonomy)
  4. Check basic EXIF (camera make/model, etc.)
  5. Write one Excel report
"""

import os
import json
import time
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
from datetime import datetime

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

# Schema
try:
    from jsonschema import Draft202012Validator, exceptions as js_exceptions
    HAS_SCHEMA = True
except ImportError:
    HAS_SCHEMA = False

# Images / EXIF
from PIL import Image
from PIL.ExifTags import TAGS

Image.MAX_IMAGE_PIXELS = None
try:
    from pillow_heif import register_heif_opener
    register_heif_opener()
except Exception:
    pass

IMAGE_EXTENSIONS = {
    ".jpg", ".jpeg", ".png", ".heic", ".heif",
    ".gif", ".bmp", ".tif", ".tiff", ".webp",
}


def clean_excel_value(value):
    if value is None:
        return ""
    if not isinstance(value, str):
        return value
    return "".join(ch for ch in value if ch in "\t\n\r" or ord(ch) >= 32)


def normalize_key(text: str) -> str:
    return text.strip().replace("_", " ").lower()


class IntakeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Delivery Intake Validator – Image Metadata QA")
        self.root.geometry("900x640")
        self.root.minsize(720, 500)

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        self.schema = None
        self.validator = None
        self._enum_lookup = {}
        self.setup_ui()

    def setup_ui(self):
        top = ttk.LabelFrame(self.root, text=" Setup ", padding=12)
        top.pack(fill="x", padx=12, pady=8)

        ttk.Label(top, text="Delivery folder (vendor date folder):").grid(row=0, column=0, sticky="w", pady=4)
        self.folder_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.folder_var, width=55).grid(row=0, column=1, padx=5, sticky="ew")
        ttk.Button(top, text="Browse…", command=self.browse_folder).grid(row=0, column=2, padx=2)

        ttk.Label(top, text="JSON Schema file:").grid(row=1, column=0, sticky="w", pady=4)
        self.schema_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.schema_var, width=55).grid(row=1, column=1, padx=5, sticky="ew")
        ttk.Button(top, text="Browse…", command=self.browse_schema).grid(row=1, column=2, padx=2)

        self.recursive_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(top, text="Include subfolders", variable=self.recursive_var).grid(
            row=2, column=1, sticky="w", pady=4
        )
        top.columnconfigure(1, weight=1)

        opts = ttk.LabelFrame(self.root, text=" Checks ", padding=10)
        opts.pack(fill="x", padx=12, pady=4)
        self.do_pairs = tk.BooleanVar(value=True)
        self.do_schema = tk.BooleanVar(value=True)
        self.do_url = tk.BooleanVar(value=True)
        self.do_exif = tk.BooleanVar(value=True)
        ttk.Checkbutton(opts, text="1. Image ↔ JSON pairing & counts", variable=self.do_pairs).pack(anchor="w")
        ttk.Checkbutton(opts, text="2. JSON schema (exact category/sub_category names)", variable=self.do_schema).pack(anchor="w")
        ttk.Checkbutton(opts, text="3. URL required check (taxonomy)", variable=self.do_url).pack(anchor="w")
        ttk.Checkbutton(opts, text="4. EXIF (camera make / model)", variable=self.do_exif).pack(anchor="w")

        btn_row = ttk.Frame(self.root)
        btn_row.pack(fill="x", padx=12, pady=6)
        self.run_btn = ttk.Button(btn_row, text="▶ Run Intake Validation", command=self.start)
        self.run_btn.pack(side="left")
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(btn_row, textvariable=self.status_var).pack(side="left", padx=12)

        self.progress = ttk.Progressbar(self.root, mode="determinate")
        self.progress.pack(fill="x", padx=12, pady=4)

        log_frame = ttk.LabelFrame(self.root, text=" Log ", padding=8)
        log_frame.pack(fill="both", expand=True, padx=12, pady=8)
        self.log_text = tk.Text(log_frame, wrap="word", font=("Consolas", 10), height=18)
        sb = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.log_text.pack(side="left", fill="both", expand=True)
        self.log_text.tag_config("ok", foreground="#1e8e3e")
        self.log_text.tag_config("fail", foreground="#d93025")
        self.log_text.tag_config("info", foreground="#1a73e8")
        self.log_text.tag_config("warn", foreground="#b06000")

    def log(self, msg, tag=None):
        self.log_text.insert(tk.END, msg + "\n", tag)
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def browse_folder(self):
        path = filedialog.askdirectory(title="Select vendor delivery folder")
        if path:
            self.folder_var.set(path)

    def browse_schema(self):
        path = filedialog.askopenfilename(
            title="Select JSON Schema",
            filetypes=[("JSON", "*.json"), ("All", "*.*")],
        )
        if path:
            self.schema_var.set(path)

    def start(self):
        folder = self.folder_var.get().strip()
        schema = self.schema_var.get().strip()
        if not folder or not Path(folder).is_dir():
            messagebox.showwarning("Missing folder", "Select a valid delivery folder.")
            return
        if self.do_schema.get():
            if not HAS_SCHEMA:
                messagebox.showerror("Missing package", "Install jsonschema:\npip install jsonschema")
                return
            if not schema or not Path(schema).is_file():
                messagebox.showwarning("Missing schema", "Select the JSON schema file.")
                return
        self.run_btn.config(state="disabled")
        self.log_text.delete("1.0", tk.END)
        threading.Thread(target=self.run, args=(folder, schema), daemon=True).start()

    def load_schema(self, schema_path):
        with open(schema_path, "r", encoding="utf-8") as f:
            schema = json.load(f)
        Draft202012Validator.check_schema(schema)
        self.schema = schema
        self.validator = Draft202012Validator(schema)
        self._enum_lookup = {}
        for field in ("category", "sub_category"):
            enums = schema.get("properties", {}).get(field, {}).get("enum", [])
            self._enum_lookup[field] = {
                normalize_key(item): item for item in enums if isinstance(item, str)
            }

    def normalize_categories(self, obj):
        if isinstance(obj, dict):
            for key, value in list(obj.items()):
                if key in ("category", "sub_category") and isinstance(value, str):
                    lookup = self._enum_lookup.get(key, {})
                    norm = normalize_key(value)
                    if norm in lookup:
                        obj[key] = lookup[norm]
                else:
                    self.normalize_categories(value)
        elif isinstance(obj, list):
            for item in obj:
                self.normalize_categories(item)


    def load_taxonomy(self):
        """Load taxonomy.json for URL required/optional."""
        here = Path(__file__).resolve().parent
        candidates = [
            here / "taxonomy.json",
            here.parent / "qa_app" / "taxonomy.json",
            Path.cwd() / "taxonomy.json",
        ]
        for p in candidates:
            if p.is_file():
                with open(p, "r", encoding="utf-8") as f:
                    self.taxonomy = json.load(f)
                self.log(f"Taxonomy loaded: {p} ({len(self.taxonomy)} rows)", "info")
                return
        self.taxonomy = []
        self.log("Taxonomy.json not found — URL check will mark Unknown", "warn")

    def exact_schema_name(self, field, value):
        """Check value is exactly a schema enum member. Returns (is_exact, expected_or_'')."""
        if not isinstance(value, str) or not value.strip():
            return False, ""
        enums = self.schema.get("properties", {}).get(field, {}).get("enum", [])
        if value in enums:
            return True, value
        expected = self._enum_lookup.get(field, {}).get(normalize_key(value), "")
        return False, expected

    def lookup_url_requirement(self, category, sub_category):
        if not getattr(self, "taxonomy", None):
            return "Unknown"

        def n(s):
            s = (s or "").lower()
            for ch in ("_", "-", "&", "/", ","):
                s = s.replace(ch, " ")
            s = s.replace(" and ", " ")
            return " ".join(x for x in s.split() if x)

        nc, ns = n(category), n(sub_category)
        best, best_score = None, 0
        for t in self.taxonomy:
            tsub = n(t.get("sub_category", ""))
            tcat = n(t.get("category", ""))
            if not ns:
                continue
            score = 0
            if ns == tsub:
                score = 3
            elif ns in tsub or tsub in ns:
                score = 2
            else:
                continue
            if nc and (nc == tcat or nc in tcat or tcat in nc):
                score += 1
            if score > best_score:
                best_score = score
                best = t
        if best:
            return best.get("url_requirement") or "Unknown"
        return "Unknown"

    def url_is_present(self, url_val):
        if url_val is None:
            return False
        s = str(url_val).strip()
        if not s or s.lower() in ("n/a", "na", "none", "null", "nan"):
            return False
        return True

    def collect_files(self, folder, recursive):
        images = {}  # base -> path
        jsons = {}
        folder = Path(folder)
        if recursive:
            paths = list(folder.rglob("*"))
        else:
            paths = list(folder.iterdir())

        for p in paths:
            if not p.is_file():
                continue
            ext = p.suffix.lower()
            base = p.stem
            if ext in IMAGE_EXTENSIONS:
                # prefer first; note duplicates
                if base not in images:
                    images[base] = p
                else:
                    # keep list of dupes via side channel later
                    images[base + f"__dup__{p.name}"] = p
            elif ext == ".json":
                if base not in jsons:
                    jsons[base] = p
                else:
                    jsons[base + f"__dup__{p.name}"] = p
        return images, jsons

    def extract_exif_basic(self, image_path):
        try:
            img = Image.open(image_path)
            width, height = img.size
            make = model = software = ""
            has_exif = False
            try:
                raw = img.getexif()
                if raw:
                    has_exif = True
                    for tag_id, value in raw.items():
                        tag = TAGS.get(tag_id, str(tag_id))
                        if tag == "Make":
                            make = clean_excel_value(str(value))
                        elif tag == "Model":
                            model = clean_excel_value(str(value))
                        elif tag == "Software":
                            software = clean_excel_value(str(value))
            except Exception:
                pass
            return {
                "width": width,
                "height": height,
                "has_exif": has_exif,
                "camera_make": make,
                "camera_model": model,
                "software": software,
            }
        except Exception as e:
            return {
                "width": "",
                "height": "",
                "has_exif": False,
                "camera_make": "",
                "camera_model": "",
                "software": "",
                "error": str(e),
            }

    def run(self, folder, schema_path):
        try:
            self.status_var.set("Scanning…")
            start = time.time()
            recursive = self.recursive_var.get()

            if self.do_schema.get():
                try:
                    self.load_schema(schema_path)
                    self.log(f"Schema loaded: {Path(schema_path).name}", "info")
                except Exception as e:
                    self.log(f"Schema load failed: {e}", "fail")
                    self.run_btn.config(state="normal")
                    self.status_var.set("Failed")
                    return

            self.log(f"Scanning: {folder}", "info")
            images_map, jsons_map = self.collect_files(folder, recursive)

            # Separate real bases from dupe keys
            def real_bases(d):
                return {k: v for k, v in d.items() if "__dup__" not in k}

            images = real_bases(images_map)
            jsons = real_bases(jsons_map)
            img_dupes = [k for k in images_map if "__dup__" in k]
            json_dupes = [k for k in jsons_map if "__dup__" in k]

            img_bases = set(images.keys())
            json_bases = set(jsons.keys())
            matched = sorted(img_bases & json_bases)
            only_img = sorted(img_bases - json_bases)
            only_json = sorted(json_bases - img_bases)

            self.log("=" * 60)
            self.log(f"Images found : {len(images)}", "info")
            self.log(f"JSON found   : {len(jsons)}", "info")
            self.log(f"Matched pairs: {len(matched)}", "ok" if len(matched) else "warn")
            self.log(f"Image only   : {len(only_img)}", "fail" if only_img else "ok")
            self.log(f"JSON only    : {len(only_json)}", "fail" if only_json else "ok")
            if img_dupes or json_dupes:
                self.log(f"Duplicate image names: {len(img_dupes)} | Duplicate JSON names: {len(json_dupes)}", "warn")

            pair_rows = []
            schema_rows = []
            exif_rows = []
            summary_issues = []
            naming_rows = []
            url_rows = []
            naming_fail = 0
            url_missing = 0
            schema_pass = schema_fail = 0

            # Pairing detail
            if self.do_pairs.get():
                for base in matched:
                    pair_rows.append({
                        "Base Name": base,
                        "Status": "MATCHED",
                        "Image": images[base].name,
                        "JSON": jsons[base].name,
                        "Image Path": str(images[base]),
                        "JSON Path": str(jsons[base]),
                        "Issue": "",
                    })
                for base in only_img:
                    pair_rows.append({
                        "Base Name": base,
                        "Status": "IMAGE_ONLY",
                        "Image": images[base].name,
                        "JSON": "",
                        "Image Path": str(images[base]),
                        "JSON Path": "",
                        "Issue": "Missing JSON",
                    })
                    summary_issues.append({"Base Name": base, "Check": "Pairing", "Issue": "Missing JSON"})
                for base in only_json:
                    pair_rows.append({
                        "Base Name": base,
                        "Status": "JSON_ONLY",
                        "Image": "",
                        "JSON": jsons[base].name,
                        "Image Path": "",
                        "JSON Path": str(jsons[base]),
                        "Issue": "Missing image",
                    })
                    summary_issues.append({"Base Name": base, "Check": "Pairing", "Issue": "Missing image"})

            # Work list for schema + exif: matched pairs primarily; also JSON-only for schema
            work_json = [(b, jsons[b], images.get(b)) for b in matched]
            work_json += [(b, jsons[b], None) for b in only_json]

            total_steps = 0
            if self.do_schema.get() or self.do_url.get():
                total_steps += len(work_json)
            if self.do_exif.get():
                total_steps += len(matched) + len(only_img)
            self.progress["maximum"] = max(total_steps, 1)
            step = 0

            # Schema + exact naming + URL checks (single pass over JSON)
            schema_pass = schema_fail = 0
            naming_fail = 0
            url_rows = []
            url_missing = 0
            naming_rows = []

            if self.do_schema.get() or self.do_url.get():
                if self.do_schema.get():
                    try:
                        self.load_schema(schema_path)
                        self.log(f"Schema loaded: {Path(schema_path).name}", "info")
                    except Exception as e:
                        self.log(f"Schema load error: {e}", "fail")
                        raise
                if self.do_url.get():
                    self.load_taxonomy()

                self.status_var.set("Validating JSON / naming / URL…")
                for base, jpath, ipath in work_json:
                    step += 1
                    self.progress["value"] = step
                    try:
                        with open(jpath, "r", encoding="utf-8") as f:
                            data = json.load(f)

                        # --- Exact schema category / sub_category names ---
                        cat_raw = data.get("category", "") if isinstance(data, dict) else ""
                        sub_raw = data.get("sub_category", "") if isinstance(data, dict) else ""
                        naming_issues = []
                        if self.do_schema.get() and self.schema:
                            cat_ok, cat_exp = self.exact_schema_name("category", cat_raw)
                            sub_ok, sub_exp = self.exact_schema_name("sub_category", sub_raw)
                            if not cat_ok:
                                naming_issues.append(
                                    f"category must be schema enum exactly; got '{cat_raw}'"
                                    + (f" → use '{cat_exp}'" if cat_exp else "")
                                )
                            if not sub_ok:
                                naming_issues.append(
                                    f"sub_category must be schema enum exactly; got '{sub_raw}'"
                                    + (f" → use '{sub_exp}'" if sub_exp else "")
                                )
                            naming_rows.append({
                                "Base Name": base,
                                "JSON File": jpath.name,
                                "Category (submitted)": cat_raw,
                                "Category (schema)": cat_exp if not cat_ok else cat_raw,
                                "Sub-category (submitted)": sub_raw,
                                "Sub-category (schema)": sub_exp if not sub_ok else sub_raw,
                                "Status": "PASS" if not naming_issues else "FAIL",
                                "Issue": " | ".join(naming_issues),
                            })
                            if naming_issues:
                                naming_fail += 1
                                summary_issues.append({
                                    "Base Name": base,
                                    "Check": "Schema naming",
                                    "Issue": " | ".join(naming_issues)[:500],
                                })

                        # --- URL required (taxonomy) ---
                        if self.do_url.get():
                            url_val = data.get("url") if isinstance(data, dict) else None
                            req = self.lookup_url_requirement(str(cat_raw), str(sub_raw))
                            present = self.url_is_present(url_val)
                            if req == "Required" and not present:
                                url_status = "MISSING_URL"
                                url_missing += 1
                                summary_issues.append({
                                    "Base Name": base,
                                    "Check": "URL",
                                    "Issue": f"URL required for {cat_raw}/{sub_raw} but missing or N/A",
                                })
                            elif req == "Required" and present:
                                url_status = "OK"
                            elif req == "Optional":
                                url_status = "OK_OPTIONAL"
                            else:
                                url_status = "UNKNOWN_TAXONOMY"
                            url_rows.append({
                                "Base Name": base,
                                "JSON File": jpath.name,
                                "Category": cat_raw,
                                "Sub-category": sub_raw,
                                "URL_Requirement": req,
                                "URL_Value": "" if url_val is None else str(url_val)[:200],
                                "URL_Present": "Yes" if present else "No",
                                "Status": url_status,
                            })

                        # --- Full schema validation (normalize only for enum flexibility on other checks) ---
                        if self.do_schema.get():
                            data_for_schema = json.loads(json.dumps(data))  # copy
                            self.normalize_categories(data_for_schema)
                            errors = sorted(
                                self.validator.iter_errors(data_for_schema),
                                key=lambda e: list(e.path),
                            )
                            # Treat exact-name failures as schema FAIL as well
                            if naming_issues:
                                errors_msgs = naming_issues[:]
                                for err in errors:
                                    loc = " -> ".join(map(str, err.path)) if err.path else "root"
                                    errors_msgs.append(f"[{loc}] {err.message}")
                                schema_fail += 1
                                schema_rows.append({
                                    "Base Name": base,
                                    "JSON File": jpath.name,
                                    "Status": "FAIL",
                                    "Errors": " | ".join(errors_msgs),
                                })
                            elif not errors:
                                schema_pass += 1
                                schema_rows.append({
                                    "Base Name": base,
                                    "JSON File": jpath.name,
                                    "Status": "PASS",
                                    "Errors": "",
                                })
                            else:
                                schema_fail += 1
                                msgs = []
                                for err in errors:
                                    loc = " -> ".join(map(str, err.path)) if err.path else "root"
                                    msgs.append(f"[{loc}] {err.message}")
                                schema_rows.append({
                                    "Base Name": base,
                                    "JSON File": jpath.name,
                                    "Status": "FAIL",
                                    "Errors": " | ".join(msgs),
                                })
                                summary_issues.append({
                                    "Base Name": base,
                                    "Check": "Schema",
                                    "Issue": " | ".join(msgs)[:500],
                                })
                    except json.JSONDecodeError as e:
                        if self.do_schema.get():
                            schema_fail += 1
                            schema_rows.append({
                                "Base Name": base,
                                "JSON File": jpath.name,
                                "Status": "FAIL",
                                "Errors": f"Invalid JSON: {e}",
                            })
                            summary_issues.append({"Base Name": base, "Check": "Schema", "Issue": f"Invalid JSON: {e}"})
                    except Exception as e:
                        if self.do_schema.get():
                            schema_fail += 1
                            schema_rows.append({
                                "Base Name": base,
                                "JSON File": jpath.name,
                                "Status": "FAIL",
                                "Errors": str(e),
                            })
                            summary_issues.append({"Base Name": base, "Check": "Schema", "Issue": str(e)})

                if self.do_schema.get():
                    self.log(
                        f"Schema: {schema_pass} PASS | {schema_fail} FAIL "
                        f"(naming mismatches counted as FAIL: {naming_fail})",
                        "ok" if schema_fail == 0 else "fail",
                    )
                if self.do_url.get():
                    self.log(
                        f"URL: {url_missing} missing required | {len(url_rows) - url_missing} OK/optional/unknown",
                        "ok" if url_missing == 0 else "fail",
                    )

            # EXIF on images (matched + image-only)
            exif_ok = exif_flag = 0
            if self.do_exif.get():
                self.status_var.set("Checking EXIF…")
                img_list = [(b, images[b]) for b in matched] + [(b, images[b]) for b in only_img]
                for base, ipath in img_list:
                    step += 1
                    self.progress["value"] = step
                    info = self.extract_exif_basic(ipath)
                    issues = []
                    if not info.get("has_exif"):
                        issues.append("No EXIF")
                    if not info.get("camera_make"):
                        issues.append("Missing camera_make")
                    if not info.get("camera_model"):
                        issues.append("Missing camera_model")
                    soft = (info.get("software") or "").lower()
                    if any(x in soft for x in ("whatsapp", "instagram", "facebook", "snapchat", "telegram")):
                        issues.append(f"Suspicious software: {info.get('software')}")

                    status = "OK" if not issues else "FLAG"
                    if issues:
                        exif_flag += 1
                        summary_issues.append({
                            "Base Name": base,
                            "Check": "EXIF",
                            "Issue": "; ".join(issues),
                        })
                    else:
                        exif_ok += 1

                    exif_rows.append({
                        "Base Name": base,
                        "Image": ipath.name,
                        "Status": status,
                        "Width": info.get("width", ""),
                        "Height": info.get("height", ""),
                        "Has EXIF": "Yes" if info.get("has_exif") else "No",
                        "Camera Make": info.get("camera_make", ""),
                        "Camera Model": info.get("camera_model", ""),
                        "Software": info.get("software", ""),
                        "Issues": "; ".join(issues),
                    })
                self.log(f"EXIF: {exif_ok} OK | {exif_flag} flagged", "ok" if exif_flag == 0 else "warn")

            # Clean pairs = matched + schema pass (if schema run) + no critical EXIF missing make/model (soft flag only)
            schema_fail_bases = {
                r["Base Name"] for r in schema_rows if r.get("Status") == "FAIL"
            } if self.do_schema.get() else set()
            pair_fail_bases = set(only_img) | set(only_json)

            clean = []
            for base in matched:
                if base in schema_fail_bases or base in pair_fail_bases:
                    continue
                clean.append({
                    "Base Name": base,
                    "Image": images[base].name,
                    "JSON": jsons[base].name,
                    "Image Path": str(images[base]),
                    "JSON Path": str(jsons[base]),
                })

            elapsed = time.time() - start
            self.log("=" * 60)
            self.log(f"Clean pairs (ready for QA sample): {len(clean)}", "ok")
            self.log(f"Elapsed: {elapsed:.1f}s", "info")

            # Write report next to folder
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            out_name = f"Intake_Report_{Path(folder).name}_{ts}.xlsx"
            out_path = str(Path(folder) / out_name)

            summary_df = pd.DataFrame([
                {"Metric": "Images", "Value": len(images)},
                {"Metric": "JSON files", "Value": len(jsons)},
                {"Metric": "Matched pairs", "Value": len(matched)},
                {"Metric": "Image only (missing JSON)", "Value": len(only_img)},
                {"Metric": "JSON only (missing image)", "Value": len(only_json)},
                {"Metric": "Schema PASS", "Value": schema_pass if self.do_schema.get() else "N/A"},
                {"Metric": "Schema FAIL", "Value": schema_fail if self.do_schema.get() else "N/A"},
                {"Metric": "Naming FAIL (not exact schema enum)", "Value": naming_fail if self.do_schema.get() else "N/A"},
                {"Metric": "URL missing (required)", "Value": url_missing if self.do_url.get() else "N/A"},
                {"Metric": "EXIF OK", "Value": exif_ok if self.do_exif.get() else "N/A"},
                {"Metric": "EXIF flagged", "Value": exif_flag if self.do_exif.get() else "N/A"},
                {"Metric": "Clean pairs", "Value": len(clean)},
                {"Metric": "Folder", "Value": folder},
                {"Metric": "Report time", "Value": datetime.now().isoformat(timespec="seconds")},
            ])

            with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
                summary_df.to_excel(writer, sheet_name="Summary", index=False)
                if pair_rows:
                    pd.DataFrame(pair_rows).to_excel(writer, sheet_name="Pairing", index=False)
                if schema_rows:
                    pd.DataFrame(schema_rows).to_excel(writer, sheet_name="Schema", index=False)
                if naming_rows:
                    pd.DataFrame(naming_rows).to_excel(writer, sheet_name="Schema_Naming", index=False)
                if url_rows:
                    pd.DataFrame(url_rows).to_excel(writer, sheet_name="URL_Check", index=False)
                if exif_rows:
                    pd.DataFrame(exif_rows).to_excel(writer, sheet_name="EXIF", index=False)
                if summary_issues:
                    pd.DataFrame(summary_issues).to_excel(writer, sheet_name="All_Issues", index=False)
                if clean:
                    pd.DataFrame(clean).to_excel(writer, sheet_name="Clean_Pairs", index=False)

            self._autofit(out_path)
            self.log(f"Report saved:\n{out_path}", "info")
            self.status_var.set("Done")
            messagebox.showinfo("Intake complete", f"Report saved to:\n\n{out_path}")
        except Exception as e:
            self.log(f"ERROR: {e}", "fail")
            self.status_var.set("Failed")
            messagebox.showerror("Error", str(e))
        finally:
            self.run_btn.config(state="normal")
            self.progress["value"] = 0

    def _autofit(self, filename):
        try:
            wb = load_workbook(filename)
            header_fill = PatternFill("solid", fgColor="1F4E79")
            header_font = Font(bold=True, color="FFFFFF")
            for ws in wb.worksheets:
                for col in ws.columns:
                    max_length = 0
                    letter = col[0].column_letter
                    for cell in col:
                        if cell.value is not None:
                            max_length = max(max_length, len(str(cell.value)))
                    ws.column_dimensions[letter].width = min(max_length + 2, 80)
                for cell in ws[1]:
                    cell.font = header_font
                    cell.fill = header_fill
            wb.save(filename)
        except Exception:
            pass


if __name__ == "__main__":
    root = tk.Tk()
    IntakeApp(root)
    root.mainloop()
