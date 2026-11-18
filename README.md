# Spend/Support Processing Portal

A simple web interface for running the existing Spend/Support workflow.

## Workflow

1. Upload the ADS CSV file and Variable Mapping Excel file.
2. Run Spend/Support Production.
3. Run Quality Control on the generated production file.
4. Generate the GZIP-compressed production file.
5. Download the production CSV, QC workbook, GZIP file, or all outputs in one ZIP.

The page displays a live workflow percentage while each stage is running. Since the
underlying scripts do not publish internal progress callbacks, the percentage is an
estimate across the current stage and advances to the next stage when that script ends.

## Processing scripts

- `spend_support_26(1).py` — Production
- `spend_support_qc_26.py` — Quality Control
- `csv_gz.py` — GZIP compression

The processing scripts are retained as supplied in the working package. The QC
transformation-week setting is 26 to match Production.

## Run locally

1. Install Python 3.10 or newer.
2. Install dependencies: `pip install -r requirements.txt`
3. Start the app: `streamlit run app.py`
4. Open the local URL shown in the terminal.

## Deploy

Deploy this folder to a Python-capable hosting platform with Streamlit support.
Use `app.py` as the entry point and install dependencies from `requirements.txt`.
The included `Dockerfile` can also be used to build a container.

```bash
docker build -t spend-support-portal .
docker run --rm -p 8501:8501 spend-support-portal
```

## Files

- `app.py` — Web interface and workflow orchestration
- `spend_support_26(1).py` — Production script
- `spend_support_qc_26.py` — QC script
- `csv_gz.py` — GZIP script
- `requirements.txt` — Python dependencies
- `Dockerfile` — Container deployment definition
