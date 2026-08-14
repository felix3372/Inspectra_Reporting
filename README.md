# Inspectra — QA Report Generator (standalone)

A standalone Streamlit app that generates QA reports from an uploaded Excel file
(`Qualified` / `Disqualified` sheets, plus an optional `Mapping` sheet that drives
the hierarchical **Custom Multi-Level Report**).

## Setup

```bash
pip install -r requirements.txt
```

## Run

```bash
streamlit run app.py
```

Then open http://localhost:8501 and upload your `.xlsx` / `.xlsm` file.

## Notes

- **No login** — the app opens straight to the uploader.
- **Upload only** — the network-path file picker was removed, so there are no
  shared-drive (`W:\`) or `utils/` dependencies.
- `.streamlit/config.toml` raises the upload limit to 1 GB.
- The optional `Mapping` sheet defines grouping levels (headers, left-to-right)
  and their valid parent→child combinations (merged cells = hierarchy). Rows in
  the exported Excel that fall outside the Mapping snapshot are highlighted in red.

## Project layout

```
Inspectra_Reporting/
  app.py                 # the Streamlit app (run this)
  QA_Report_Helper/      # report/data/export logic
  .streamlit/config.toml # upload-size config
  requirements.txt
  runtime.txt
```
