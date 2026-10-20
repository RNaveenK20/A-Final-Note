# Farewell Website

A small, responsive farewell site built with HTML, CSS, vanilla JavaScript, Flask, and one Excel workbook.

## Run locally

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python app.py
```

Open `http://127.0.0.1:5000`. The workbook is created automatically at `data/FarewellData.xlsx`; alternatively run `python scripts/initialize_excel.py` first.

## Customise

Edit the `people` array in `static/js/app.js` to change colleague names, messages, order, or visibility. Keep those names aligned with `KNOWN_NAMES` in `app.py` and `NAMES` in `scripts/initialize_excel.py`.

## Excel storage and limits

The workbook has one `FarewellData` sheet with exactly `Name`, `Like`, and `Comment` columns. Likes and comments are saved on every request. An in-process lock serialises workbook reads/writes, suitable for a small single-process deployment only. It is not a high-traffic or multi-worker data store.

Each browser receives a random UUID in LocalStorage and the running server keeps a lightweight in-memory record to discourage repeat likes. This basic protection resets if browser storage is cleared, a different device/browser is used, or the server restarts; it is not authentication.

## Deployment

Deploy the Flask app as one process on a host with persistent writable storage for `data/FarewellData.xlsx`. Do not expose the `data/` folder as static/public files.
