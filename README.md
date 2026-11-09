# Workday Accounting Journal EIB Ingest Studio & Agent

An automated enterprise ingestion engine and native Windows Desktop Studio (`WorkdayEIBStudio`) that transforms raw accounting journal entries (`.xlsx` / `.csv`) into official **Workday Enterprise Interface Builder (EIB)** workbooks (`v45.1` Two-Sheet and `v46.0`–`v48.0` Single-Sheet templates) with 100% debit/credit balance verification.

> **Standalone Windows Executable (`WorkdayEIBStudio.exe`)**
> The standalone executable is published separately on the [**GitHub Releases**](https://github.com/Git229KEV/CSV-to-EIB-Agent/releases) page.

---

## Recent Changes & Fixes Included

1. **Unified Two-Sheet (`v45.1`) Mapping for 2025 & 2026 Datasets (`run_agent.py`)**
   - Unified `process_journal_2025` and `process_journal_2026` to use the exact same two-sheet `Import Accounting Journal` (`v45.1`) relational mapping logic (`Import Accounting Journal` Header sheet + `Journal Entry Line Replacement` Lines sheet).
   - **Strict Blank Column Enforcement (Sheet 1)**: Enforced **Column E (`Accounting Journal`, Col 5)** and **Column U (`Record Quantity`, Col 21)** to remain strictly blank (`None`) across all data rows (Rows 6+), while populating **Column F (`Accounting Journal ID`, Col 6)** from `AccountingJournalID` / `JournalKey`.
   - **Strict `LineMemo` Blank Handling (Sheet 2)**: Any `NULL`, `NONE`, `NAN`, or empty values in `LineMemo` are written strictly as empty cells (`None`) in **Column T (`Memo`, Col 20)** with zero fallback to header memo.
   - **Spend Category Lookup**: Automatically maps 108 Ledger Accounts to their corresponding Workday Spend Categories (`Col Y` / Col 25) via `source/Spend Category Mapped Ledger Account.xlsx`.
   - **3-Digit Cost Center Formatting**: Formats Cost Center codes (`Col AA` / Col 27) as 3-digit strings (`000`, `070`, `090`, `200`) with Excel text format `@` to preserve leading zeros.
   - **Zero-Amount Preservation**: Preserves explicit `0.0` values on Debit (`Col J`) and Credit (`Col K`) lines.

2. **Enterprise Desktop Studio Enhancements (`desktop_app.py` & `frontend/index.html`)**
   - **Custom Output Filename Input**: Added user-configurable output filename input in Step 2 (with automatic `.xlsx` extension handling) alongside native Windows Folder Browser (`select_folder`) and File Browser (`select_file`) dialogs.
   - **Smart Source Discovery & Fallback Staging**: Enhanced `find_source_file` to resolve files across workspace `source/`, executable directory, PyInstaller `_MEIPASS`, and user `Downloads` / `Desktop` / `Documents` folders, plus base64 file staging (`stage_uploaded_file`) for HTML file input fallback.
   - **Lightweight `_PandasLite` Engine**: Replaced heavy `pandas`/`numpy` runtime dependencies with an internal `_PandasLite` (`openpyxl` + `csv`) reader, reducing the standalone executable size to ~17 MB and eliminating startup overhead.

---

## Project Structure

- `desktop_app.py` — Native Windows Desktop application (`pywebview` + Native IPC bridge, zero local network ports required).
- `run_agent.py` — Core CLI and library ingestion engine supporting Workday EIB `v45.1` (Two-Sheet) and `v46.0` / `v47.0` / `v48.0` (Single-Sheet) templates.
- `frontend/` — Enterprise Desktop UI (`index.html`, `journal_data.js`).
- `source/` — Raw accounting journal source files (2025, 2026, CSV) and `Spend Category Mapped Ledger Account.xlsx`.
- `target/` — Pristine Workday EIB templates (`v45.1`, `v46.0`, `v47.0`, `v48.0`).
- `output/` — Generated EIB workbooks organized by date (`output/YYYY-MM-DD/`).
- `WorkdayEIBStudio.spec` — PyInstaller specification for building `WorkdayEIBStudio.exe`.
- `HOW_IT_WORKS.md` & `agent.md` — Detailed field mapping and architecture documentation.

---

## Quick Start

### Option 1: Run Standalone Executable
Download **`WorkdayEIBStudio.exe`** from the [**Releases**](https://github.com/Git229KEV/CSV-to-EIB-Agent/releases) page and launch it directly on Windows.

### Option 2: Run from Python Source
```powershell
# Launch the Desktop GUI
python desktop_app.py

# Or run via CLI (interactive or batch)
python run_agent.py
python run_agent.py --all
```
