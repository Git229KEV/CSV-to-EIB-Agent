---
name: workday-eib-mapper
description: Specialized domain skill for Claude AI Agent to autonomously analyze, map, and validate accounting journal schemas from arbitrary CSV/Excel source files to Workday EIB spreadsheet templates (v45.1 two-sheet, v46.0–v48.0 single-sheet).
---

# Workday Accounting Journal EIB Mapping Skill (Claude AI Agent)

This skill provides comprehensive instructions, accounting domain rules, field semantics, and validation logic for the Claude AI Agent to accurately map raw accounting files (CSV/Excel) to official Workday Enterprise Interface Builder (EIB) templates.

---

## 1. Workday EIB Template Architectures

### 1.1 Two-Sheet Architecture (`v45.1`)
- **Sheet 1 (`Import Accounting Journal`)**: Contains journal header records (Rows 1–5 metadata, Row 6+ data).
  - `Header Key*` (`Col B` / Col 2): Unique integer identifier (`1..N`) for each distinct journal header.
  - `Add Only` (`Col C` / Col 3): Defaults to `'Y'`.
  - `Accounting Journal` (`Col E` / Col 5): **Identical** to Column F (`Accounting Journal ID`).
  - `Accounting Journal ID` (`Col F` / Col 6): Journal identification string (e.g. `01-2026-P2-01`).
  - `Company*` (`Col L` / Col 12): Primary Company Reference ID (e.g. `LE01`).
  - `Currency*` (`Col M` / Col 13): Currency code (e.g. `USD`).
  - `Ledger Type*` (`Col N` / Col 14): Standardized to `Actuals`.
  - `Accounting Date*` (`Col P` / Col 16): Formatted as `YYYY-MM-DD`.
  - `Journal Source*` (`Col Q` / Col 17): Standardized to `Manual_Journal`.
  - `Record Quantity` (`Col U` / Col 21): Set to `'N'`.
  - `Journal Entry Memo` (`Col V` / Col 22): Header-level narrative.
  - `External Reference ID` (`Col W` / Col 23): Journal external reference.
  - `Auto Complete` (`Col AG` / Col 33): Set to `'Y'`.

- **Sheet 2 (`Journal Entry Line Replacement`)**: Contains journal line records (Rows 1–5 metadata, Row 6+ data).
  - `Header Key` (`Col B` / Col 2): Foreign key referencing Sheet 1. Count must strictly match source file distribution.
  - `Line Key` (`Col C` / Col 3): Mapped directly from source `ROW_ID` or unique line identifier.
  - `Line Order` (`Col D` / Col 4): Mapped directly from source `LineOrder` or line sequence.
  - `Line Company` (`Col E` / Col 5): Line-level company identifier (e.g. `LE01`).
  - `Ledger Account` (`Col F` / Col 6): General Ledger account number (pure numbers as int, alphanumeric subaccounts like `18000_01_02` as text).
  - `Account Set` (`Col G` / Col 7): Mapped from parent account set (e.g. `ACCOUNT_SET`).
  - `Debit Amount` (`Col J` / Col 10): Numeric float. If zero in source, populated as `0.0`.
  - `Credit Amount` (`Col K` / Col 11): Numeric float. If zero in source, populated as `0.0`.
  - `Currency` (`Col L` / Col 12): Line currency (e.g. `USD`).
  - `Memo` (`Col T` / Col 20): Mapped from `LineMemo`. Empty or `"NULL"` must remain strictly blank with no fallback to header memo.
  - `External Reference ID` (`Col U` / Col 21): Line-level external reference.
  - `Budget Date` (`Col V` / Col 22): Formatted as `YYYY-MM-DD`.
  - `Revenue Category` (`Col W` / Col 23): Worktag Revenue Category ID.
  - `Spend Category` (`Col Y` / Col 25): Master lookup from `spend category 1.xlsx` (155 accounts) based on Ledger Account, with fallback to source spend category.
  - `Cost Center` (`Col AA` / Col 27): **Strictly formatted as a 3-digit text** (e.g. `'000'`, `'090'`, `'070'`, `'200'`) with leading zeros preserved.
  - `Intercompany Affiliate` (`Col AC` / Col 29): Auto-detected across Workday & ERP column variants. **All underscores must be stripped with zero spaces added** (e.g. `'LE_04'` -> `'LE04'`).
  - `Balancing Worktag Affiliate` (`Col AE` / Col 31): Business Unit worktag.

### 1.2 Single-Sheet Architectures (`v46.0`, `v47.0`, `v48.0`)
- **Sheet (`Submit Accounting Journal`)**: Combined header and line records on a single sheet.
  - Standard columns align with the single-sheet schema (`Col B` Spreadsheet Key, `Col L` Ledger Account, `Col AP` Debit, `Col AQ` Credit, etc.).

---

## 2. ERP & Source Column Semantic Taxonomy

When analyzing arbitrary CSV or Excel files from any ERP system (SAP, Oracle, NetSuite, Microsoft Dynamics, QuickBooks, or custom GL extracts), Claude AI Agent must apply semantic recognition across these alias clusters:

| Core Field | Workday EIB Target | Semantic Aliases & ERP Headers |
|:---|:---|:---|
| `header_key` | Sheet 1 Col B / Sheet 2 Col B | `HeaderKey`, `JournalKey`, `VoucherNo`, `DocNumber`, `TransactionID`, `EntryNo`, `Batch_ID` |
| `line_key` | Sheet 2 Col C | `ROW_ID`, `LineKey`, `LineNumber`, `ItemNo`, `SeqNo`, `Line_Id` |
| `line_order` | Sheet 2 Col D | `LineOrder`, `LineSequence`, `Ordering`, `LineItem` |
| `company` | Sheet 1 Col L / Sheet 2 Col E | `CompanyReferenceID`, `LineCompanyReferenceID`, `LegalEntity`, `Entity`, `CompanyCode`, `OrgID` |
| `ledger_account` | Sheet 2 Col F | `LedgerAccountReferenceID`, `GLAccount`, `AccountCode`, `NominalCode`, `AccountNumber`, `Account` |
| `account_set` | Sheet 2 Col G | `LedgerAccountReferenceID_ParentID`, `AccountSet`, `ChartOfAccounts`, `COA`, `AccountGroup` |
| `debit` | Sheet 2 Col J | `DebitAmount`, `Debit`, `Dr`, `AmountDebit`, `DR_AMT`, `Debit_USD` |
| `credit` | Sheet 2 Col K | `CreditAmount`, `Credit`, `Cr`, `AmountCredit`, `CR_AMT`, `Credit_USD` |
| `currency` | Sheet 1 Col M / Sheet 2 Col L | `Currency`, `LineCurrency`, `Cur`, `CurrCode`, `Doc_Currency` |
| `memo` | Sheet 2 Col T | `LineMemo`, `Description`, `Narration`, `LineText`, `ItemMemo`, `Remarks` |
| `budget_date` | Sheet 2 Col V | `BudgetDate`, `Budget_Period`, `EffectiveDate`, `PostingDate` |
| `cost_center` | Sheet 2 Col AA | `CostCenter`, `Cost_Center`, `Department`, `Dept`, `CC`, `CostCentre`, `ProfitCenter` |
| `intercompany_affiliate` | Sheet 2 Col AC | `Affliate Company`, `Affliate_Company`, `Affiliate Company`, `Affiliate_Company`, `IntercompanyAffiliate`, `TradingPartner`, `PartnerEntity`, `InterCo` |
| `spend_category` | Sheet 2 Col Y | `SpendCategory`, `Worktag_Spend_Category_ID`, `Spend_Cat`, `ExpenseType` |
| `business_unit` | Sheet 2 Col AE | `BalancingWorktagAffiliate`, `BusinessUnit`, `BU`, `Worktag_Business_Unit_ID`, `Division` |
| `custom_org` | Sheet 2 Col AG | `CustomOrganization`, `Custom_Org`, `Worktag_Custom_Org` |

---

## 3. Special Transformation Rules

1. **Intercompany Affiliate Underscore Stripping**:
   - Values like `'LE_04'`, `'LE_01'`, `'AFF_02'` must be transformed by directly stripping the underscore: `'LE04'`, `'LE01'`, `'AFF02'`.
   - Never insert a space (do NOT convert `'LE_04'` to `'LE 04'`).

2. **Cost Center 3-Digit Formatting**:
   - Numeric or string cost centers under 3 digits must be left-padded with zeros (e.g. `'0'` -> `'000'`, `'90'` -> `'090'`, `'70'` -> `'070'`).
   - Stored with text format `@` to preserve leading zeros.

3. **Zero Amount Preservation**:
   - Zero debits or zero credits must be populated explicitly as `0.0`, never left blank.

4. **Line Memo Blank Handling**:
   - When source `LineMemo` is empty, whitespace, or literally `"NULL"` / `"None"` / `"NaN"`, map strictly as blank (empty cell).
   - Never fall back to header memo.

5. **Spend Category Master Lookup**:
   - Match `Ledger Account` against `spend category 1.xlsx` (155 accounts). If matched, populate Column Y. If not matched in master table, fallback to source `Spend Category` column if present.

6. **Debit / Credit Balance Invariant**:
   - Total Debits must equal Total Credits with `$0.00` variance.
