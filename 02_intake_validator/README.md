# Delivery Intake Validator – Image Metadata QA

Run this on every vendor delivery **before** QA.

## Checks
1. **Image ↔ JSON pairing** – same base filename
2. **JSON schema** – full schema validation
3. **Exact schema names** – `category` / `sub_category` must match the schema enum **exactly**
   - Client confirmed: use schema names (e.g. `Pets and Animals`, not `Pets_Animals` or `Pets & Animals`)
4. **URL required** – from taxonomy; flags assets where URL is required but missing / N/A
5. **EXIF** – camera make / model flags

## How to run
```
run_intake.bat
```
or
```
python Delivery_Intake_Validator.py
```

Select:
- Vendor date folder (images + JSON)
- `project_schema.schema.json`

Keep `taxonomy.json` in the same folder as the script (already included).

## Report sheets
| Sheet | Purpose |
|-------|---------|
| Summary | Counts |
| Pairing | Image/JSON match status |
| Schema | Schema PASS/FAIL |
| Schema_Naming | Submitted vs required schema names |
| URL_Check | Required / optional / missing URL |
| EXIF | Camera metadata flags |
| All_Issues | Combined issues |
| Clean_Pairs | Ready for QA sampling |

## Where this sits in the workflow
1. **Intake Validator** (this tool) ← schema names + URL + pairing + EXIF
2. **QA app** – sample visual/metadata review
3. **Packager** – S3 folder structure for PASS assets only
