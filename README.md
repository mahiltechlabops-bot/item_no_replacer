# Item No Replacer

A local Streamlit tool to match and replace item_No values in Excel A using records from Excel B.

## Setup (one time)

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

## Run

```powershell
.\.venv\Scripts\Activate.ps1
python -m streamlit run app.py
```

Opens at: http://localhost:8501

## How to use

1. Upload **Excel A** (your product list — only Inactive records are shown)
2. Upload **Excel B** (item master)
3. **Check a row** in the Excel A table to select it
4. Matching items from Excel B appear on the right (filtered by brand + subcategory)
5. Click **Apply** next to the item you want → item_No is updated
6. Use **↩ Undo / ↪ Redo** to go back and forth
7. Click **💾 Save** to download the updated Excel A file

## Notes

- Change log is saved to `change_history.json` in the same folder — persists across sessions
- Matching is case-insensitive and partial (e.g. "Medimix" matches "Medimix Green")
- Only Inactive records from Excel A are shown
