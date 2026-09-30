# Jaseer Virtual Office — Web / Render

Interactive browser version of the Virtual Office. The office itself is HTML/CSS/JS (not a static picture); Flask/Python is the accounting backend.

## Deploy to GitHub + Render
1. Create a PRIVATE GitHub repository and upload the contents of this folder (not the folder itself).
2. In Render: New > Blueprint and connect that repository. `render.yaml` supplies the build/start commands.
3. Do not commit real supermarket `.xls/.xlsx` files to GitHub. The included `data/Demo_Data.xlsx` is demo data only.
4. Optional LLM: add Render environment variables `LLM_API_URL`, `LLM_API_KEY`, `LLM_MODEL`. The URL must be an OpenAI-compatible `/chat/completions` endpoint. Without these variables, the app uses deterministic Python answers.
5. Google Drive: `GDRIVE_FOLDER_ID` is reserved for the real Drive connector. Do not put Drive credentials or API keys in GitHub. Add secrets only in Render Environment.

## Local test
```
pip install -r requirements.txt
python app.py
```
Open http://127.0.0.1:5000

## Current capabilities
- Interactive office rooms; click Sales, Inventory, Finance, Supplier, Audit, Reporting, or Jaseer.
- Staff status changes Idle > Working > Completed.
- Live Flow Manager.
- Chat with Python-calculated sales/inventory/finance/audit/supplier results.
- Advisor/investigation mode for natural questions.
- Optional cloud LLM for language/reasoning; Python remains authoritative for financial figures.
- Upload a smaller XLSX/XLS/CSV at runtime.
- Demo workbook is included so the app works immediately after deployment.

## Large Google Drive sales files
The UI and backend are ready for a server-side Drive sync/cache layer, but credentials are intentionally not bundled. Large files should be downloaded/processed on the server and only compact results returned to the browser. The real Drive downloader from the existing MTD notebook should be connected using private Render secrets before production use.
