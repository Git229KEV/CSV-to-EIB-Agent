# Workday Accounting Journal EIB Ingest System — Architecture & Workflow Documentation

This document explains the end-to-end architecture, data transformation rules, execution flow, and user interfaces of the **Workday Accounting Journal EIB Ingest System (v45.1)**.

---

## 1. Executive Summary

The system automates the ingestion of raw accounting journal entries from source files (`Accounting_Journal_Raw_02_202601.xlsx` and `Accounting_Journal_RawAgg_01_202501.xlsx`) into official Workday Enterprise Interface Builder (EIB) two-sheet spreadsheets (`Import_Accounting_Journal_v45.0_LE01_2026_P2.xlsx` and `Import_Accounting_Journal_v45.0_LE01_2025_P1.xlsx`).

Key Features:
- **Restored 2025 Mapping Architecture**: 
  - Sheet 1 (`Import Accounting Journal`): Exactly **1 journal header row** (Row 6) for `01-2025-P1` (`Header Key* = 1`).
  - Sheet 2 (`Journal Entry Line Replacement`): Exactly **8 journal line rows** (Rows 6 to 13), all linked to `Header Key = 1`.
  - `Line Key` mapped from `ROW_ID` (1..8) and `Line Order` mapped from `LineOrder` (1..8).
- **2026 Mapping Architecture**:
  - Sheet 1 (`Import Accounting Journal`): **20 journal headers** (Rows 6 to 25, `Header Key* 1..20`).
  - Sheet 2 (`Journal Entry Line Replacement`): **222 journal lines** (Rows 6 to 227), with `Header Key` occurrences strictly preserved from source file.
  - `Line Key` mapped from `ROW_ID` and `Line Order` mapped from `LineOrder`.
- **Spend Category Mapping Lookup (`Spend Category Mapped Ledger Account.xlsx`)**:
  - The second sheet (`Journal Entry Line Replacement`) contains `Ledger Account` in Column F (Col 6) and `Spend Category` in Column Y (Col 25).
  - For every record, the system references `Spend Category Mapped Ledger Account.xlsx` (108 accounts) and maps the matching `Spend Category` (e.g. `74900 -> SC098`, `74800 -> SC048`, `77100 -> SC049`, `75400 -> SC043`, etc.) into Column 25.
- **Date-Wise Output Generation**:
  - Output files are automatically saved into a date-stamped folder: `output/YYYY-MM-DD/` (e.g., `output/2026-10-29/`).
  - File names strictly follow the Workday standard:
    - 2025 File: `Import_Accounting_Journal_v45.0_LE01_2025_P1.xlsx`
    - 2026 File: `Import_Accounting_Journal_v45.0_LE01_2026_P2.xlsx`
  - Synchronized root copies are also maintained at the workspace root for direct access.
- **Two-Sheet Relational Integrity**: Sheet 1 (`Import Accounting Journal`) contains journal headers linked via `Header Key*` to Sheet 2 (`Journal Entry Line Replacement`) containing line items linked via `Header Key`.
- **Zero Template Corruption**: Preserves Workday XML namespaces, macros, metadata, and validation headers in **Rows 1–5** in both sheets.
- **Field Alignments**:
  - **Sheet 1 Col 5 (`Col E`) & Col 6 (`Col F`)**: Both populated with `Accounting Journal ID` (`AccountingJournalID` / `JournalKey`).
  - **Sheet 2 Col 6 (`Col F`)**: `Ledger Account` mapped from `LedgerAccountReferenceID` (preserving alphanumeric subaccounts like `18000_01_02`).
  - **Sheet 2 Col 7 (`Col G`)**: `Account Set` mapped from `LedgerAccountReferenceID_ParentID` (`ACCOUNT_SET`).
  - **Cost Center (`Col AA` / Col 27)**: Formatted as a **3-digit number** (e.g. `'000'`, `'090'`, `'070'`, `'200'`) with leading zeros preserved and text format `@`.
  - **Zero Amount Preservation**: Source file zero amounts for debits and credits are populated as `0.0` (not left empty).
  - **Line Memo Null & Blank Integrity**: Cells with `"NULL"` or empty values under `LineMemo` in the source file are mapped strictly as blank (empty) cells under `Memo` (`Col T` / Col 20 in Sheet 2) with no fallback to `JournalEntryMemo`.

```
+-------------------------------------------------------+
|  Select Source Accounting File                        |
|  [1] Accounting_Journal_Raw_02_202601.xlsx            |
|  [2] Accounting_Journal_RawAgg_01_202501.xlsx         |
|  [3] Process BOTH 2025 and 2026 files                 |
+---------------------------+---------------------------+
                            |
                            v
+-------------------------------------------------------+
|  Spend Category Lookup Engine                         |
|  - Reads 'Spend Category Mapped Ledger Account.xlsx'  |
|  - Maps Ledger Account -> Spend Category              |
+---------------------------+---------------------------+
                            |
                            v
+-------------------------------------------------------+
|       Date-Wise Output Directory Creation             |
|       output/YYYY-MM-DD/                              |
|       - Import_Accounting_Journal_v45.0_LE01_2026_P2  |
|       - Import_Accounting_Journal_v45.0_LE01_2025_P1  |
+-------------------------------------------------------+
```

---

## 2. Workday EIB Spreadsheet Layout (v45.1 Two-Sheet)

### 2.1 Sheet 1: `Import Accounting Journal` (Header Sheet)

| Field Label | EIB Column | Excel Col | Source Field / Rule |
|:---|:---:|:---:|:---|
| **Header Key\*** | `B` | 2 | Sequential Integer (`1..N`) for each unique journal header |
| **Add Only** | `C` | 3 | `AddOnlyJournal` (`'Y'` or `'N'`, default `'Y'`) |
| **Accounting Journal** | **`E`** | **5** | **Populated with the same value as Column F (`Accounting Journal ID`)** |
| **Accounting Journal ID** | **`F`** | **6** | **`AccountingJournalID` / `JournalKey`** |
| **Company\*** | `L` | 12 | `CompanyReferenceID` (`LE01`) |
| **Currency\*** | `M` | 13 | `Currency` (`USD`) |
| **Ledger Type\*** | `N` | 14 | `LedgerType` (standardized to `Actuals`) |
| **Accounting Date\*** | `P` | 16 | `AccountingDate` (`YYYY-MM-DD`) |
| **Journal Source\*** | `Q` | 17 | `JournalSource` (`Manual_Journal`) |
| **Record Quantity** | `U` | 21 | `Record_Quantity` (`'N'`) |
| **Journal Entry Memo** | `V` | 22 | `JournalEntryMemo` |
| **External Reference ID** | `W` | 23 | `JournalExternalReferenceID` |
| **Auto Complete** | `AG` | 33 | Populated with `'Y'` for automated ingestion |

### 2.2 Sheet 2: `Journal Entry Line Replacement` (Lines Sheet)

| Field Label | EIB Column | Excel Col | Source Field / Rule |
|:---|:---:|:---:|:---|
| **Header Key** | `B` | 2 | **2026**: Count strictly matches source file occurrences<br>**2025**: All 8 lines link to `Header Key = 1` |
| **Line Key** | **`C`** | **3** | **Mapped directly from `ROW_ID` column** in source file |
| **Line Order** | **`D`** | **4** | **Mapped directly from `LineOrder` column** in source file |
| **Line Company** | `E` | 5 | `LineCompanyReferenceID` (`LE01`) |
| **Ledger Account** | **`F`** | **6** | **`LedgerAccountReferenceID` (e.g. `74900`, `74800`, `18000_01_02`)** |
| **Account Set** | **`G`** | **7** | **`LedgerAccountReferenceID_ParentID` (`ACCOUNT_SET`)** |
| **Debit Amount** | `J` | 10 | `DebitAmount` (**zeros in source populated as `0.0`**) |
| **Credit Amount** | `K` | 11 | `CreditAmount` (**zeros in source populated as `0.0`**) |
| **Currency** | `L` | 12 | `LineCurrency` (`USD`) |
| **Memo** | `T` | 20 | `LineMemo` (cells with `"NULL"` or empty values mapped as blank; no fallback to `JournalEntryMemo`) |
| **External Reference ID** | `U` | 21 | `JournalLineExternalReferenceID` |
| **Budget Date** | `V` | 22 | `BudgetDate` |
| **Revenue Category** | `W` | 23 | `Worktag_Revenue_Category_ID` |
| **Spend Category** | **`Y`** | **25** | **Looked up from `Spend Category Mapped Ledger Account.xlsx` based on `Ledger Account` (Col 6)** |
| **Cost Center** | **`AA`** | **27** | **Formatted as 3-digit number** (e.g. `'000'`, `'090'`, `'070'`, `'200'`) |
| **Balancing Worktag Affiliate** | `AC` | 29 | `Worktag_Business_Unit_ID` |

---

## 3. Comparison: 2025 vs 2026 Dataset Mapping

| Metric / Attribute | 2025 Dataset (`Accounting_Journal_RawAgg_01_202501.xlsx`) | 2026 Dataset (`Accounting_Journal_Raw_02_202601.xlsx`) |
|:---|:---|:---|
| **Target Output File** | `Import_Accounting_Journal_v45.0_LE01_2025_P1.xlsx` | `Import_Accounting_Journal_v45.0_LE01_2026_P2.xlsx` |
| **Date-Wise Directory** | `output/<YYYY-MM-DD>/` | `output/<YYYY-MM-DD>/` |
| **Sheet 1 (Headers)** | **1 Header Row** (Row 6, `01-2025-P1`) | **20 Header Rows** (Rows 6 to 25) |
| **Sheet 2 (Lines)** | **8 Line Rows** (Rows 6 to 13, all `HeaderKey = 1`) | **222 Line Rows** (Rows 6 to 227) |
| **Spend Categories Mapped** | **2 lines** (`74800 -> SC048`, `77100 -> SC049`) | **56 lines** (`74900 -> SC098`, `75400 -> SC043`, etc.) |
| **Line Key & Line Order** | Mapped from `ROW_ID` (1..8) & `LineOrder` (1..8) | Mapped from `ROW_ID` & `LineOrder` |
| **Financial Balance** | Total: **$96,086,590.02** (Variance: **$0.00**) | Total: **$9,367,000.88** (Variance: **$0.00**) |

---

## 4. Execution & Verification

### Interactive Terminal Run
```powershell
python run_agent.py
```
*Allows interactive selection of [1] 2026, [2] 2025, or [3] Process BOTH files simultaneously.*

### Automated Batch Run (Both 2025 and 2026)
```powershell
python run_agent.py --all
```

### Direct Single File Ingest
```powershell
# Ingest 2025:
python run_agent.py --source Accounting_Journal_RawAgg_01_202501.xlsx

# Ingest 2026:
python run_agent.py --source Accounting_Journal_Raw_02_202601.xlsx
```

### Interactive Web Studio
```powershell
python launch_studio.py
```
*Opens [http://localhost:8080](http://localhost:8080) for real-time verification, worktag summary, and spreadsheet downloads.*
