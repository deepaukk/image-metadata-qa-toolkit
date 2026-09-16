"""
Local QA Viewer – Image Metadata QA
Two reviewers, count-based sampling (min 1 per locale), taxonomy guide, separate results.
"""

import streamlit as st
import pandas as pd
import json
import re
from pathlib import Path
from datetime import datetime
from PIL import Image

try:
    from pillow_heif import register_heif_opener
    register_heif_opener()
except Exception:
    pass

APP_TITLE = "Local QA Viewer"
BASE = Path(__file__).parent
RESULTS_DIR = BASE / "qa_results"
RESULTS_DIR.mkdir(exist_ok=True)
ERROR_CODES_FILE = BASE / "error_codes.json"
TAXONOMY_FILE = BASE / "taxonomy.json"
IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp", ".heic", ".heif", ".tif", ".tiff"}

st.set_page_config(page_title=APP_TITLE, layout="wide", initial_sidebar_state="expanded")


# ── helpers ──────────────────────────────────────────────────────────────────
def load_error_codes():
    if ERROR_CODES_FILE.exists():
        return json.loads(ERROR_CODES_FILE.read_text(encoding="utf-8"))
    return []


def load_taxonomy():
    if TAXONOMY_FILE.exists():
        return json.loads(TAXONOMY_FILE.read_text(encoding="utf-8"))
    return []


def find_column(df, candidates):
    cols = {c.lower().strip(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in cols:
            return cols[cand.lower()]
    for cand in candidates:
        for cl, orig in cols.items():
            if cand.lower() in cl:
                return orig
    return None


def build_image_index(folder: Path):
    index = {}
    if not folder or not folder.exists():
        return index
    for p in folder.rglob("*"):
        if p.is_file() and p.suffix.lower() in IMAGE_EXTENSIONS:
            stem = p.stem.lower()
            if stem not in index:
                index[stem] = p
            base = re.sub(r"[_\-]\d+$", "", stem)
            if base not in index:
                index[base] = p
    return index


def resolve_image(row, image_index, filename_col, asset_col):
    candidates = []
    if filename_col and pd.notna(row.get(filename_col)):
        stem = Path(str(row[filename_col]).strip()).stem.lower()
        candidates += [stem, re.sub(r"[_\-]\d+$", "", stem)]
    if asset_col and pd.notna(row.get(asset_col)):
        aid = str(row[asset_col]).strip().lower()
        candidates += [aid, Path(aid).stem.lower()]
    for c in candidates:
        if c in image_index:
            return image_index[c]
    return None


def sample_with_locale_coverage(df, target_n, locale_col, seed=42):
    """
    Pick target_n rows with at least 1 per locale when possible.
    Deterministic (seed).
    """
    n = len(df)
    if n == 0 or target_n <= 0:
        return df.iloc[0:0].copy()
    target_n = min(target_n, n)

    if not locale_col or locale_col not in df.columns:
        # fall back to systematic every-kth
        k = max(1, n // target_n)
        offset = seed % k
        idx = list(range(offset, n, k))[:target_n]
        return df.iloc[idx].copy().reset_index(drop=True)

    rng = pd.Series(range(n))
    # shuffle indices deterministically
    shuffled = df.sample(frac=1, random_state=seed)

    selected_idx = []
    seen_locales = set()

    # Phase 1: one per locale
    for i, row in shuffled.iterrows():
        loc = str(row[locale_col]) if pd.notna(row[locale_col]) else "UNKNOWN"
        if loc not in seen_locales:
            selected_idx.append(i)
            seen_locales.add(loc)
        if len(selected_idx) >= target_n:
            break

    # Phase 2: fill remaining toward target_n (systematic from remaining)
    if len(selected_idx) < target_n:
        remaining = [i for i in shuffled.index if i not in selected_idx]
        need = target_n - len(selected_idx)
        # take evenly spaced from remaining
        if remaining:
            step = max(1, len(remaining) // need)
            extra = remaining[::step][:need]
            selected_idx.extend(extra)

    selected_idx = selected_idx[:target_n]
    # preserve original order of appearance in full df for stable navigation
    selected_idx = sorted(selected_idx, key=lambda x: list(df.index).index(x) if x in df.index else 0)
    out = df.loc[selected_idx].copy().reset_index(drop=True)
    return out


def assign_reviewers(sample_df):
    """Alternate Reviewer 1 / Reviewer 2."""
    owners = ["Reviewer 1" if i % 2 == 0 else "Reviewer 2" for i in range(len(sample_df))]
    sample_df = sample_df.copy()
    sample_df["_qa_owner"] = owners
    return sample_df


def results_path(session_id: str) -> Path:
    return RESULTS_DIR / f"qa_results_{session_id}.csv"


def load_results(session_id: str) -> pd.DataFrame:
    path = results_path(session_id)
    cols = ["session_id", "batch_id", "asset_id", "participant_id", "locale", "filename",
            "qa_owner", "qa_result", "error_code", "qa_comments", "qa_date", "sample_index"]
    if path.exists():
        # Force text columns to string so Batch IDs like 20260817 are not int64
        dtype_map = {c: str for c in cols if c != "sample_index"}
        df = pd.read_csv(path, dtype=dtype_map, keep_default_na=False)
        for c in cols:
            if c not in df.columns:
                df[c] = "" if c != "sample_index" else 0
        # sample_index as int-safe
        df["sample_index"] = pd.to_numeric(df["sample_index"], errors="coerce").fillna(-1).astype(int)
        for c in cols:
            if c != "sample_index":
                df[c] = df[c].astype(str)
        return df
    return pd.DataFrame(columns=cols)


def save_result_row(session_id, batch_id, asset_id, participant_id, locale, filename,
                    qa_owner, qa_result, error_code, comments, sample_index):
    path = results_path(session_id)
    row = {
        "session_id": str(session_id or ""),
        "batch_id": str(batch_id or ""),
        "asset_id": str(asset_id or ""),
        "participant_id": str(participant_id or ""),
        "locale": str(locale or ""),
        "filename": str(filename or ""),
        "qa_owner": str(qa_owner or ""),
        "qa_result": str(qa_result or ""),
        "error_code": str(error_code or ""),
        "qa_comments": str(comments or ""),
        "qa_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "sample_index": int(sample_index),
    }
    if path.exists():
        df = load_results(session_id)
        mask = (df["sample_index"].astype(int) == int(sample_index)) & (df["qa_owner"].astype(str) == str(qa_owner))
        if mask.any():
            for k, v in row.items():
                df.loc[mask, k] = v
        else:
            df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    else:
        df = pd.DataFrame([row])
    # Ensure string columns stay strings before write
    for c in df.columns:
        if c != "sample_index":
            df[c] = df[c].astype(str)
    df.to_csv(path, index=False)
    return df


def _norm(s):
    """Normalize for fuzzy match: lower, underscores/hyphens/& -> space, collapse spaces."""
    s = (s or "").lower()
    for ch in ("_", "-", "&", "/", ","):
        s = s.replace(ch, " ")
    s = " ".join(s.split())
    return s


def _tokens(s):
    # drop very common short words that don't help matching
    stop = {"and", "or", "the", "of", "a", "an"}
    return {w for w in _norm(s).split() if len(w) > 1 and w not in stop}


def _token_overlap(a, b):
    wa, wb = _tokens(a), _tokens(b)
    if not wa or not wb:
        return 0.0
    return len(wa & wb) / max(len(wa), len(wb))


# Metadata often uses underscores; taxonomy uses spaces/&
# Map frequent metadata category strings → taxonomy category key words
CATEGORY_ALIASES = {
    "environmental natural elements": "environmental natural elements",
    "fine art museum": "fine art museum",
    "pets animals": "pets animals",
    "household items": "household items",
    "electronics devices": "electronics devices",
    "tools hardware": "tools hardware",
    "clothing accessories": "clothing accessories",
    "sports fitness": "sports fitness",
    "toys games": "toys games",
    "food beverage": "food beverage",
    "books media": "books media",
    "religious cultural": "religious cultural",
    "outdoor garden": "outdoor garden",
    "text documents": "text documents",
    "business storefronts": "business storefronts",
}


def taxonomy_lookup(taxonomy, category, sub_category):
    """Prefer exact/close sub-category match; fall back to category only if needed."""
    if not taxonomy:
        return []
    cat = _norm(category)
    sub = _norm(sub_category)
    cat = CATEGORY_ALIASES.get(cat, cat)

    sub_hits = []
    cat_hits = []

    for t in taxonomy:
        tcat = _norm(t.get("category", ""))
        tsub = _norm(t.get("sub_category", ""))

        cat_ok = (
            (not cat)
            or (cat == tcat)
            or (cat in tcat)
            or (tcat in cat)
            or _token_overlap(cat, tcat) >= 0.5
        )
        if not cat_ok:
            continue

        if sub:
            if sub == tsub or sub in tsub or tsub in sub:
                sub_hits.append((0, t))
            elif _token_overlap(sub, tsub) >= 0.75:
                sub_hits.append((1, t))
            else:
                cat_hits.append(t)
        else:
            cat_hits.append(t)

    if sub_hits:
        sub_hits.sort(key=lambda x: x[0])
        out, seen = [], set()
        for _, h in sub_hits:
            key = (h.get("category"), h.get("sub_category"))
            if key not in seen:
                seen.add(key)
                out.append(h)
        return out[:2]

    out, seen = [], set()
    for h in cat_hits:
        key = (h.get("category"), h.get("sub_category"))
        if key not in seen:
            seen.add(key)
            out.append(h)
    return out[:2]


# ── session defaults ─────────────────────────────────────────────────────────
defaults = {
    "meta_df": None,
    "sample_df": None,
    "image_index": {},
    "current_idx": 0,
    "session_id": datetime.now().strftime("%Y%m%d_%H%M%S"),
    "batch_id": "",
    "image_folder": "",
    "filename_col": None,
    "asset_col": None,
    "locale_col": None,
    "participant_col": None,
    "category_col": None,
    "subcat_col": None,
    "loaded": False,
    "reviewer": "Reviewer 1",
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ── sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("1. Setup")
    st.session_state.batch_id = st.text_input("Batch ID", value=st.session_state.batch_id or "")

    st.session_state.reviewer = st.radio(
        "I am",
        ["Reviewer 1", "Reviewer 2"],
        index=0 if st.session_state.reviewer == "Reviewer 1" else 1,
        horizontal=True,
    )

    meta_file = st.file_uploader("Metadata Excel / CSV", type=["xlsx", "xls", "csv"])
    image_folder_str = st.text_input(
        "Image folder path",
        value=st.session_state.image_folder,
        help=r"e.g. C:\Deliveries\20260813",
    )
    st.session_state.image_folder = image_folder_str

    st.markdown("**Sample size**")
    sample_mode = st.radio("Mode", ["Number of assets", "Percentage"], horizontal=True)
    if sample_mode == "Number of assets":
        sample_n = st.number_input("How many to QA", min_value=1, value=50, step=1)
        sample_pct = None
    else:
        sample_pct = st.selectbox("QA %", [5, 10, 15, 20], index=3)
        sample_n = None

    if st.button("Load & generate sample", type="primary"):
        if meta_file is None:
            st.error("Upload metadata.")
        elif not image_folder_str.strip():
            st.error("Enter image folder path.")
        else:
            folder = Path(image_folder_str.strip())
            if not folder.exists():
                st.error(f"Folder not found:\n{folder}")
            else:
                if meta_file.name.lower().endswith(".csv"):
                    df = pd.read_csv(meta_file)
                else:
                    df = pd.read_excel(meta_file)

                asset_col = find_column(df, ["asset_id", "Asset_ID", "Asset ID", "assetid", "UUID", "id"])
                filename_col = find_column(df, ["filename", "Filename", "file_name", "File Name", "image"])
                locale_col = find_column(df, ["locale", "Locale", "Language_Code", "language_code", "lang"])
                participant_col = find_column(df, ["participant_id", "Participant ID", "Participant_ID", "participant", "Participant"])
                category_col = find_column(df, ["category", "Category"])
                subcat_col = find_column(df, ["sub_category", "subcategory", "Sub Category", "Subcategory"])

                if not asset_col and not filename_col:
                    st.error("Need Asset ID or Filename column.")
                else:
                    if sample_n is None:
                        sample_n = max(1, int(round(len(df) * sample_pct / 100)))

                    sample = sample_with_locale_coverage(df, int(sample_n), locale_col, seed=42)
                    sample = assign_reviewers(sample)

                    st.session_state.meta_df = df
                    st.session_state.sample_df = sample
                    st.session_state.asset_col = asset_col
                    st.session_state.filename_col = filename_col
                    st.session_state.locale_col = locale_col
                    st.session_state.participant_col = participant_col
                    st.session_state.category_col = category_col
                    st.session_state.subcat_col = subcat_col
                    st.session_state.image_index = build_image_index(folder)
                    st.session_state.current_idx = 0
                    st.session_state.session_id = f"{st.session_state.batch_id or 'batch'}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                    st.session_state.loaded = True

                    locales = sample[locale_col].nunique() if locale_col else "?"
                    st.success(
                        f"{len(df)} total → sample {len(sample)} "
                        f"({locales} locales). "
                        f"Images indexed: {len(st.session_state.image_index)}"
                    )

    st.divider()
    st.header("2. Export report")
    if st.session_state.loaded:
        res = load_results(st.session_state.session_id)
        # Report columns requested
        report_cols = ["batch_id", "asset_id", "locale", "filename", "qa_owner", "qa_result", "qa_comments"]
        # include error_code as helpful extra before comments
        export_df = res.copy()
        if len(export_df):
            # rename for export
            out = pd.DataFrame({
                "Batch ID": export_df.get("batch_id", ""),
                "Asset ID": export_df.get("asset_id", ""),
                "Participant ID": export_df.get("participant_id", ""),
                "Locale": export_df.get("locale", ""),
                "Filename": export_df.get("filename", ""),
                "QA Owner": export_df.get("qa_owner", ""),
                "QA Result": export_df.get("qa_result", ""),
                "Error Code": export_df.get("error_code", ""),
                "QA Comments": export_df.get("qa_comments", ""),
            })
            csv_bytes = out.to_csv(index=False).encode("utf-8")
            st.download_button(
                "Download QA report CSV",
                data=csv_bytes,
                file_name=f"QA_Report_{st.session_state.batch_id or st.session_state.session_id}.csv",
                mime="text/csv",
            )
            st.caption(f"{len(out)} result(s)")
        else:
            st.caption("No results yet.")

    st.divider()
    with st.expander("Error codes"):
        for c in load_error_codes():
            st.text(f"{c['code']}  {c['reason']}")

# ── main ─────────────────────────────────────────────────────────────────────
st.title(APP_TITLE)
st.caption("Original metadata is never modified. Results are stored separately.")

if not st.session_state.loaded or st.session_state.sample_df is None or len(st.session_state.sample_df) == 0:
    st.info("← Load metadata + image folder, set sample size, then click **Load & generate sample**.")
    st.markdown("""
**Workflow**
1. Both reviewers load the **same** metadata file and folder (same sample is generated).
2. Choose **Reviewer 1** or **Reviewer 2** — each only sees their assigned assets.
3. Review image + metadata → PASS/FAIL → Save & Continue (or Skip).
4. Download QA report and share/merge with the other reviewer.
    """)
    st.stop()

full_sample = st.session_state.sample_df
# Filter to current reviewer
my_sample = full_sample[full_sample["_qa_owner"] == st.session_state.reviewer].reset_index(drop=True)
# Keep link to original sample_index in full list for stable saving
# Map: position in my_sample -> original index in full_sample
my_orig_indices = full_sample.index[full_sample["_qa_owner"] == st.session_state.reviewer].tolist()
# Actually sample_df was reset_index, so use integer positions
owner_mask = full_sample["_qa_owner"] == st.session_state.reviewer
full_positions = list(full_sample.index[owner_mask])  # positions in full_sample

if len(my_sample) == 0:
    st.warning(f"No assets assigned to {st.session_state.reviewer}.")
    st.stop()

n = len(my_sample)
idx = max(0, min(st.session_state.current_idx, n - 1))
st.session_state.current_idx = idx
row = my_sample.iloc[idx]
# original position in full sample (for result key)
full_pos = int(full_positions[idx]) if idx < len(full_positions) else idx

asset_col = st.session_state.asset_col
filename_col = st.session_state.filename_col
locale_col = st.session_state.locale_col
category_col = st.session_state.category_col
subcat_col = st.session_state.subcat_col

asset_id = str(row[asset_col]) if asset_col else ""
filename = str(row[filename_col]) if filename_col else ""
locale = str(row[locale_col]) if locale_col and pd.notna(row.get(locale_col)) else ""
participant_col = st.session_state.get("participant_col")
participant_id = str(row[participant_col]) if participant_col and pd.notna(row.get(participant_col)) else ""
category = str(row[category_col]) if category_col and pd.notna(row.get(category_col)) else ""
subcat = str(row[subcat_col]) if subcat_col and pd.notna(row.get(subcat_col)) else ""

img_path = resolve_image(row, st.session_state.image_index, filename_col, asset_col)

results = load_results(st.session_state.session_id)
my_results = results[results["qa_owner"] == st.session_state.reviewer] if len(results) else results
done = len(my_results)
passes = int((my_results["qa_result"] == "PASS").sum()) if len(my_results) else 0
fails = int((my_results["qa_result"] == "FAIL").sum()) if len(my_results) else 0

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("My progress", f"{done} / {n}")
c2.metric("PASS", passes)
c3.metric("FAIL", fails)
c4.metric("Current", f"{idx + 1} of {n}")
c5.metric("Reviewer", st.session_state.reviewer)

existing = my_results[my_results["sample_index"] == full_pos] if len(my_results) else pd.DataFrame()
prev = existing.iloc[0].to_dict() if len(existing) else {}

left, right = st.columns([1.15, 1])

with left:
    st.subheader("Image")
    if img_path and Path(img_path).exists():
        try:
            image = Image.open(img_path)
            st.image(image, use_container_width=True, caption=Path(img_path).name)
            st.caption(f"{image.size[0]}×{image.size[1]}")
        except Exception as e:
            st.error(f"Could not open image: {e}")
    else:
        st.error("⚠️ Image not found")
        st.write(f"Asset ID: `{asset_id}` · Filename: `{filename}`")

with right:
    st.subheader("Metadata")
    st.markdown(f"**Asset ID:** {asset_id}")
    st.markdown(f"**Participant ID:** {participant_id}")
    st.markdown(f"**Filename:** {filename}")
    st.markdown(f"**Locale:** {locale}")
    st.markdown(f"**Category:** {category}")
    st.markdown(f"**Sub-category:** {subcat}")

    preferred = [
        ("url", ["url", "URL"]),
        ("disambiguation_details", ["disambiguation_details", "Disambiguation details", "disambiguation"]),
        ("subject_name", ["subject_name", "Subject name", "subject"]),
        ("place_details", ["place_details", "Place details"]),
    ]
    shown = {asset_col, filename_col, locale_col, category_col, subcat_col}
    for label, cands in preferred:
        col = find_column(my_sample, cands)
        if col and col not in shown:
            val = row.get(col, "")
            if pd.notna(val) and str(val).strip():
                st.markdown(f"**{label}:** {val}")
            shown.add(col)

    with st.expander("All metadata fields"):
        for col in my_sample.columns:
            if str(col).startswith("_"):
                continue
            val = row.get(col, "")
            if pd.notna(val) and str(val).strip():
                st.text(f"{col}: {val}")

    # Taxonomy guidance
    st.subheader("Taxonomy guide")
    tax = load_taxonomy()
    hits = taxonomy_lookup(tax, category, subcat)
    if hits:
        for h in hits:
            req = h.get("url_requirement", "")
            badge = "🔴 URL Required" if req == "Required" else ("🟡 URL Optional" if req == "Optional" else req)
            st.markdown(f"**{h.get('category')}** / {h.get('sub_category')} → {badge}")
            with st.expander("Examples & notes", expanded=False):
                desc = (h.get("description") or "").strip()
                examples = (h.get("examples") or "").strip()
                notes = (h.get("notes") or "").strip()
                privacy = (h.get("privacy_note") or "").strip()
                if desc:
                    st.markdown(f"**Description:** {desc}")
                if examples:
                    st.markdown(f"**Examples:** {examples}")
                if notes:
                    st.markdown(f"**Notes:** {notes}")
                if privacy:
                    st.markdown(f"**Privacy:** {privacy}")
                if not any([desc, examples, notes, privacy]):
                    st.caption("No extra details stored for this row. Refer to the taxonomy PDF.")
    else:
        st.caption("No taxonomy match found for this category. Check the taxonomy PDF manually.")

    st.divider()
    st.subheader("QA decision")
    codes = load_error_codes()
    code_options = [""] + [f"{c['code']} — {c['reason']}" for c in codes]

    # Unique widget keys per asset so selections do NOT carry over on Next/Skip
    wkey = f"{st.session_state.session_id}_{full_pos}_{st.session_state.reviewer}"

    # Only pre-fill if this asset was already saved; otherwise start fresh
    prev_result = prev.get("qa_result", "") if prev else ""
    result_options = ["— select result —", "PASS", "FAIL"]
    if prev_result == "PASS":
        result_index = 1
    elif prev_result == "FAIL":
        result_index = 2
    else:
        result_index = 0
    result_choice = st.radio(
        "Result",
        result_options,
        index=result_index,
        horizontal=True,
        key=f"result_{wkey}",
    )
    qa_result = result_choice if result_choice in ("PASS", "FAIL") else ""

    default_code = prev.get("error_code", "") if prev else ""
    code_index = 0
    for i, opt in enumerate(code_options):
        if default_code and opt.startswith(str(default_code)):
            code_index = i
            break
    error_sel = st.selectbox(
        "Error code (if FAIL)",
        code_options,
        index=code_index,
        key=f"error_{wkey}",
    )
    error_code = error_sel.split(" — ")[0] if error_sel else ""
    comments = st.text_area(
        "Comments",
        value=prev.get("qa_comments", "") if prev else "",
        height=70,
        key=f"comments_{wkey}",
    )

    b1, b2, b3 = st.columns(3)
    with b1:
        back_clicked = st.button("⬅ Back", use_container_width=True)
    with b2:
        save_cont_clicked = st.button("💾 Save & Continue", use_container_width=True, type="primary")
    with b3:
        skip_clicked = st.button("Skip ➡", use_container_width=True)

    if back_clicked:
        st.session_state.current_idx = max(0, idx - 1)
        st.rerun()

    if skip_clicked:
        # Move on without saving — does not record PASS/FAIL
        st.session_state.current_idx = min(n - 1, idx + 1)
        st.rerun()

    if save_cont_clicked:
        if not qa_result:
            st.warning("Select PASS or FAIL before saving.")
        elif qa_result == "FAIL" and not error_code:
            st.warning("Select an error code for FAIL.")
        else:
            save_result_row(
                session_id=st.session_state.session_id,
                batch_id=st.session_state.batch_id,
                asset_id=asset_id,
                participant_id=participant_id,
                locale=locale,
                filename=filename,
                qa_owner=st.session_state.reviewer,
                qa_result=qa_result,
                error_code=error_code if qa_result == "FAIL" else "",
                comments=comments,
                sample_index=full_pos,
            )
            st.session_state.current_idx = min(n - 1, idx + 1)
            st.rerun()

st.caption("Back = previous asset · Save & Continue = save rating and go next · Skip = next without saving. Merge both reviewers' CSV reports when done.")
