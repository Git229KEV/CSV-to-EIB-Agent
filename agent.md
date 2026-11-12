# Workday Accounting Journal EIB Ingest Agent Specification (v45.1)

## 1. Overview & Objective
This agent extracts accounting journal entries from raw Excel/CSV files and populates an official Workday EIB two-sheet workbook ([`Import_Accounting_Journal_v45.1.xlsx`](file:///c:/CSV%20to%20EIB/Import_Accounting_Journal_v45.1.xlsx)) directly in-place. Mapping is performed **individually per file** (one-by-one).

When executed from terminal (`python run_agent.py`), the agent prompts the user to select:
1. **Source File**: (e.g., `Accounting_Journal_Raw_02_202601.xlsx` or `Accounting_Journal_RawAgg_01_202501.xlsx`)
2. **Target EIB Template**: (e.g., `Import_Accounting_Journal_v45.1.xlsx`)

---

## 2. File Specifications
- **Available Source Files**:
  - [`Accounting_Journal_Raw_LE01_202601.xlsx`](file:///c:/CSV%20to%20EIB/Accounting_Journal_Raw_LE01_202601.xlsx) (20 journals, 222 line records)
  - [`Accounting_Journal_Raw_LE01_202501.xlsx`](file:///c:/CSV%20to%20EIB/Accounting_Journal_Raw_LE01_202501.xlsx) (1 journal, 8 line records)
  - [`Accounting_Journal_Raw_02_202601.xlsx`](file:///c:/CSV%20to%20EIB/Accounting_Journal_Raw_02_202601.xlsx) (20 journals, 222 line records)
  - [`Accounting_Journal_RawAgg_01_202501.xlsx`](file:///c:/CSV%20to%20EIB/Accounting_Journal_RawAgg_01_202501.xlsx) (1 journal, 8 line records)
  - [`FPEAST_HSTBAL_155_2023P05_20230711071408.csv`](file:///c:/CSV%20to%20EIB/FPEAST_HSTBAL_155_2023P05_20230711071408.csv) (189 line records)
- **Target Excel Workbook**: [`Import_Accounting_Journal_v45.1.xlsx`](file:///c:/CSV%20to%20EIB/Import_Accounting_Journal_v45.1.xlsx)
  - **Sheet 1**: `Import Accounting Journal` (Header level)
  - **Sheet 2**: `Journal Entry Line Replacement` (Line level)
  - **Header Rows**: Rows 1–5 preserved intact
  - **Data Injection**: Rows 6+

---

## 3. Transformation & Field Mapping Rules

### 3.1 Sheet 1: `Import Accounting Journal` (Header Sheet)
Rows 1–5 are metadata and header rows. Data rows start at Row 6.
- **Header Key\*** (`Col B` / Col 2): Sequential integer identifier (`1..N`) for each unique journal header.
- **Add Only** (`Col C` / Col 3): `AddOnlyJournal` (`'Y'` or `'N'`).
- **Accounting Journal** (`Col E` / Col 5): **Populated with the exact same value as Column F (`Accounting Journal ID`)**.
- **Accounting Journal ID** (`Col F` / Col 6): Mapped directly from `AccountingJournalID` (or `JournalKey`).
- **Company\*** (`Col L` / Col 12): `CompanyReferenceID` (`LE01`).
- **Currency\*** (`Col M` / Col 13): `Currency` (`USD`).
- **Ledger Type\*** (`Col N` / Col 14): `LedgerType` (standardized to `Actuals`).
- **Accounting Date\*** (`Col P` / Col 16): `AccountingDate` (formatted as date `YYYY-MM-DD`).
- **Journal Source\*** (`Col Q` / Col 17): `JournalSource` (`Manual_Journal`).
- **Record Quantity** (`Col U` / Col 21): `Record_Quantity` (`'N'`).
- **Journal Entry Memo** (`Col V` / Col 22): `JournalEntryMemo`.
- **External Reference ID** (`Col W` / Col 23): `JournalExternalReferenceID`.
- **Auto Complete** (`Col AG` / Col 33): Always populated with `'Y'`.

### 3.2 Sheet 2: `Journal Entry Line Replacement` (Lines Sheet)
Rows 1–5 are metadata and header rows. Data rows start at Row 6.
- **Header Key** (`Col B` / Col 2): **Count strictly matches the source file** (e.g. if HeaderKey 1 appeared twice in source, it appears twice here; if HeaderKey 3 appeared 26 times, it appears 26 times).
- **Line Key** (`Col C` / Col 3): **Mapped directly from the `ROW_ID` column** in the source file.
- **Line Order** (`Col D` / Col 4): **Mapped directly from the `LineOrder` column** in the source file.
- **Line Company** (`Col E` / Col 5): `LineCompanyReferenceID` (`LE01`).
- **Ledger Account** (`Col F` / Col 6): Mapped from `LedgerAccountReferenceID` (pure numbers stored as integer, alphanumeric subaccounts like `'18000_01_02'` stored as text).
- **Account Set** (`Col G` / Col 7): Mapped from `LedgerAccountReferenceID_ParentID` (`ACCOUNT_SET`).
- **Debit Amount** (`Col J` / Col 10): Numeric float. **If debit amount is zero in source file, it is populated as `0.0`**.
- **Credit Amount** (`Col K` / Col 11): Numeric float. **If credit amount is zero in source file, it is populated as `0.0`**.
- **Currency** (`Col L` / Col 12): `LineCurrency` (`USD`).
- **Memo** (`Col T` / Col 20): `LineMemo` (if empty or `"NULL"`, mapped strictly as blank with no value and no fallback to `JournalEntryMemo`).
- **External Reference ID** (`Col U` / Col 21): `JournalLineExternalReferenceID`.
- **Budget Date** (`Col V` / Col 22): `BudgetDate`.
- **Revenue Category** (`Col W` / Col 23): `Worktag_Revenue_Category_ID`.
- **Spend Category** (`Col Y` / Col 25): `Worktag_Spend_Category_ID`.
- **Cost Center** (`Col AA` / Col 27): **Formatted as a 3-digit number** (e.g., `'000'`, `'090'`, `'070'`, `'200'`) with text format `@` to preserve leading zeros.
- **Intercompany Affiliate** (`Col AC` / Col 29): `Worktag_Intercompany_Affiliate_ID` / `IntercompanyAffiliate` (blank if absent in source).
- **Balancing Worktag Affiliate** (`Col AE` / Col 31): `Worktag_Business_Unit_ID`.