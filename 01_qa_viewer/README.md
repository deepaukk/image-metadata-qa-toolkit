# QA Viewer (Streamlit)

Local visual QA tool for image + metadata deliveries.

# Local QA Viewer – Image Metadata QA

Offline QA tool for metadata + image review.

## Features

- Load Excel/CSV metadata + local image folder
- Auto-match images by filename / Asset ID
- **Two reviewers**: Reviewer 1 / Reviewer 2 get alternating assets
- Sample by **number of assets** or **percentage**
- **At least 1 asset per locale** in the sample
- Taxonomy guide (URL required / optional) next to metadata
- Error codes aligned to project list
- Export report: Batch ID, Asset ID, Locale, Filename, QA Owner, QA Result, Error Code, QA Comments
- Original metadata is never modified

## Setup (Windows)

```bat
cd qa_app
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

Or double-click `run_qa.bat`.

## Run

```bat
venv\Scripts\activate
streamlit run app.py
```

## Two-person workflow

1. Both open the app and load the **same** metadata file and image folder.
2. Both use the **same sample size** (e.g. 50 assets or 20%).
3. Person A selects **Reviewer 1** — only sees their half.
4. Person B selects **Reviewer 2** — only sees the other half.
5. Each downloads their QA report CSV when done.
6. Merge the two CSVs (or paste into Vendor QA Log) and send to vendor.

## Sampling

- Prefer **Number of assets** (e.g. 40, 50, 100).
- Algorithm: ensure ≥1 per locale, then fill to target with deterministic shuffle (seed 42).
- Same inputs → same sample → same Reviewer 1/2 split on both machines.
