import os
import sys
import csv
import math
import re
import argparse
import datetime
import shutil
import openpyxl

MAX_BULK_FILES = 12

class _IlocIndexer:
    def __init__(self, rows):
        self._rows = rows
    def __getitem__(self, idx):
        return self._rows[idx]

class _LightweightDataFrame:
    def __init__(self, rows, columns=None):
        self._rows = rows
        self.iloc = _IlocIndexer(rows)
        if columns is not None:
            self.columns = list(columns)
        elif rows:
            self.columns = list(rows[0].keys())
        else:
            self.columns = []
    def __len__(self):
        return len(self._rows)
    def iterrows(self):
        for idx, r in enumerate(self._rows):
            yield idx, r

def _parse_csv_val(v):
    if v is None:
        return None
    s = v.strip()
    if s == "" or s.upper() in ("NAN", "NULL", "NONE"):
        return None
    try:
        if "." in s:
            return float(s)
        return int(s)
    except ValueError:
        return s

class _PandasLite:
    Timestamp = datetime.datetime

    @staticmethod
    def isna(val):
        if val is None:
            return True
        if isinstance(val, float) and math.isnan(val):
            return True
        return False

    @staticmethod
    def notna(val):
        return not _PandasLite.isna(val)

    @staticmethod
    def _extract_excel_sheet(ws):
        rows_iter = ws.iter_rows(values_only=True)
        header_row = None
        # Look in rows 1..10 for the real header row (Row 1 works immediately for standard files)
        for _ in range(10):
            cand = next(rows_iter, None)
            if cand is None:
                break
            non_empty = [v for v in cand if v is not None and str(v).strip() != ""]
            if len(non_empty) >= 2:
                header_row = cand
                break
        if not header_row:
            return [], []
        headers = [str(h).strip() if h is not None else f"Unnamed: {i}" for i, h in enumerate(header_row)]
        records = []
        for r in rows_iter:
            if r is None or all(v is None or (isinstance(v, str) and v.strip() == "") for v in r):
                continue
            row_dict = {}
            for i, h in enumerate(headers):
                val = r[i] if i < len(r) else None
                if isinstance(val, str) and val.strip() == "":
                    val = None
                row_dict[h] = val
            records.append(row_dict)
        return records, headers

    @staticmethod
    def read_excel(path):
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        ws = wb.active
        records, headers = _PandasLite._extract_excel_sheet(ws)
        if not records and len(wb.sheetnames) > 1:
            for sname in wb.sheetnames:
                cand_records, cand_headers = _PandasLite._extract_excel_sheet(wb[sname])
                if len(cand_headers) >= 2 and (cand_records or not headers):
                    records, headers = cand_records, cand_headers
                    if cand_records:
                        break
        wb.close()
        return _LightweightDataFrame(records, columns=headers)

    @staticmethod
    def read_csv(path):
        records = []
        headers = []
        with open(path, mode="r", encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            if reader.fieldnames:
                headers = [h.strip() for h in reader.fieldnames if h is not None]
            for r in reader:
                if not r or all(v is None or str(v).strip() == "" for v in r.values()):
                    continue
                row_dict = {}
                for k, v in r.items():
                    if k is not None:
                        row_dict[k.strip()] = _parse_csv_val(v)
                records.append(row_dict)
        return _LightweightDataFrame(records, columns=headers)

pd = _PandasLite()

def find_source_file(filename):
    if not filename:
        return filename
    if os.path.exists(str(filename)):
        return str(filename)
    bname = os.path.basename(str(filename))
    base_dirs = ["."]
    if getattr(sys, 'frozen', False):
        base_dirs.append(os.path.dirname(sys.executable))
    if hasattr(sys, '_MEIPASS'):
        base_dirs.append(sys._MEIPASS)
    user_home = os.path.expanduser("~")
    for extra in [
        os.path.join(user_home, "Downloads"),
        os.path.join(user_home, "Downloads", "EIB files"),
        os.path.join(user_home, "Desktop"),
        os.path.join(user_home, "Documents")
    ]:
        if os.path.exists(extra):
            base_dirs.append(extra)
    for b_dir in base_dirs:
        for folder in [".", "source", "target"]:
            candidate = os.path.join(b_dir, folder, bname)
            if os.path.exists(candidate):
                return candidate
    # Fallback matching for 2025 / 2026 raw source variations (e.g., Accounting_Journal_Raw_LE01_2025.xlsx -> 202501.xlsx)
    low_name = bname.lower()
    if "2025" in low_name:
        for b_dir in base_dirs:
            for folder in ["source", "."]:
                cand = os.path.join(b_dir, folder, "Accounting_Journal_Raw_LE01_202501.xlsx")
                if os.path.exists(cand):
                    return cand
    elif "2026" in low_name:
        for b_dir in base_dirs:
            for folder in ["source", "."]:
                cand = os.path.join(b_dir, folder, "Accounting_Journal_Raw_LE01_202601.xlsx")
                if os.path.exists(cand):
                    return cand
    return str(filename)

def find_target_file(filename):
    if not filename:
        return filename
    base_dirs = ["."]
    if hasattr(sys, '_MEIPASS'):
        base_dirs.append(sys._MEIPASS)
    for b_dir in base_dirs:
        for folder in ["target", ".", "source"]:
            candidate = os.path.join(b_dir, folder, os.path.basename(str(filename)))
            if os.path.exists(candidate):
                return candidate
    if os.path.exists(str(filename)):
        return str(filename)
    return str(filename)

DEFAULT_SOURCE_FILES = [
    find_source_file("Accounting_Journal_Raw_LE01_202601.xlsx"),
    find_source_file("Accounting_Journal_Raw_LE01_202501.xlsx"),
    find_source_file("Accounting_Journal_Raw_02_202601.xlsx"),
    find_source_file("Accounting_Journal_RawAgg_01_202501.xlsx"),
    find_source_file("FPEAST_HSTBAL_155_2023P05_20230711071408.csv")
]

TARGET_TEMPLATES = [
    {
        "id": "v45",
        "version": "v45.1",
        "title": "Workday Import Accounting Journal v45.1",
        "filename": "Import_Accounting_Journal_v45.1.template.xlsx",
        "badge": "Two-Sheet",
        "type": "two-sheet",
        "sheets": ["Import Accounting Journal", "Journal Entry Line Replacement"],
        "description": "Two-sheet layout: Sheet 1 (Headers) & Sheet 2 (Replacement Lines)."
    },
    {
        "id": "v46",
        "version": "v46.0",
        "title": "Workday Submit Accounting Journal v46.0",
        "filename": "Submit_Accounting_Journal_v46.0.xlsx",
        "badge": "Single-Sheet",
        "type": "single-sheet",
        "sheets": ["Submit Accounting Journal"],
        "description": "Enterprise single-sheet template with balancing worktags support."
    },
    {
        "id": "v47",
        "version": "v47.0",
        "title": "Workday Submit Accounting Journal v47.0",
        "filename": "Submit_Accounting_Journal_v47.0.xlsx",
        "badge": "Single-Sheet",
        "type": "single-sheet",
        "sheets": ["Submit Accounting Journal"],
        "description": "Single-sheet layout preserving 96 metadata columns, Row ID = 1, and Worktags+ dimensions."
    },
    {
        "id": "v48",
        "version": "v48.0",
        "title": "Workday Submit Accounting Journal v48.0",
        "filename": "Submit_Accounting_Journal_v48.0.xlsx",
        "badge": "Dedicated Worktags",
        "type": "single-sheet",
        "sheets": ["Submit Accounting Journal"],
        "description": "Latest Workday v48.0 architecture with 1:1 dedicated worktag columns (Cost Center, Spend Category, BU, etc.)."
    }
]

DEFAULT_TARGET_FILES = [
    find_target_file("Import_Accounting_Journal_v45.1.template.xlsx"),
    find_target_file("Submit_Accounting_Journal_v48.0.xlsx"),
    find_target_file("Submit_Accounting_Journal_v47.0.xlsx"),
    find_target_file("Submit_Accounting_Journal_v46.0.xlsx"),
    find_target_file("Import_Accounting_Journal_v45.1.xlsx")
]

def resolve_target_file(requested=None):
    """Resolves target template file by ID, version string, or filename."""
    if requested:
        for tpl in TARGET_TEMPLATES:
            if requested.lower() in (tpl["id"].lower(), tpl["version"].lower(), tpl["filename"].lower()):
                cand = find_target_file(tpl["filename"])
                if os.path.exists(cand):
                    return cand
        cand = find_target_file(requested)
        if os.path.exists(cand):
            return cand
    return find_target_file("Import_Accounting_Journal_v45.1.template.xlsx")

SPEND_MAPPING_FILE = "Spend Category Mapped Ledger Account.xlsx"

# Sheet names in target v45.1 workbook
SHEET_HEADER = "Import Accounting Journal"
SHEET_LINES = "Journal Entry Line Replacement"

def load_spend_category_map(mapping_file=None):
    """Loads ledger account -> spend category lookup dictionary from Excel."""
    if mapping_file is None:
        mapping_file = find_source_file(SPEND_MAPPING_FILE)
    spend_map = {}
    if not os.path.exists(mapping_file):
        print(f"Warning: Spend mapping file '{mapping_file}' not found.")
        return spend_map
    try:
        df = pd.read_excel(mapping_file)
        for _, row in df.iterrows():
            acct = str(row.get("Ledger Account", "")).strip()
            if acct.endswith(".0"):
                acct = acct[:-2]
            sc = str(row.get("Spend Category", "")).strip()
            if acct and sc and sc.lower() != "nan":
                spend_map[acct] = sc
        print(f"Loaded {len(spend_map)} spend category mappings from '{mapping_file}'.")
    except Exception as e:
        print(f"Error loading spend category map: {e}")
    return spend_map

def _detect_template_version_tag(template_file=None):
    if not template_file:
        return "v45.1"
    low = str(template_file).lower()
    if "v48" in low:
        return "v48.0"
    if "v47" in low:
        return "v47.0"
    if "v46" in low:
        return "v46.0"
    return "v45.1"

def get_output_filename(source_file, template_file=None):
    """Returns official Workday EIB filename for standard LE01 files, or unique filename for custom uploads."""
    bname = os.path.basename(str(source_file or ""))
    low = bname.lower()
    tpl_ver = _detect_template_version_tag(template_file)
    if tpl_ver == "v45.1":
        if "accounting_journal_raw" in low and "2025" in low:
            return "Import_Accounting_Journal_v45.0_LE01_2025_P1.xlsx"
        if "accounting_journal_raw" in low and "2026" in low:
            return "Import_Accounting_Journal_v45.0_LE01_2026_P2.xlsx"
    stem = os.path.splitext(bname)[0] or "Accounting_Journal"
    prefix = "Import_Accounting_Journal" if tpl_ver == "v45.1" else "Submit_Accounting_Journal"
    return f"{stem}_{prefix}_{tpl_ver}.xlsx"

def get_datewise_output_path(source_file, target_dir=None, template_file=None):
    """Generates the output path. If target_dir is given by user, writes directly there.
    Otherwise defaults to date-wise output/YYYY-MM-DD/ and synchronized root copy."""
    filename = get_output_filename(source_file, template_file=template_file)
    if target_dir and os.path.exists(target_dir) and os.path.isdir(target_dir):
        return os.path.join(target_dir, filename)
    date_str = datetime.date.today().strftime("%Y-%m-%d")
    output_dir = os.path.join("output", date_str)
    os.makedirs(output_dir, exist_ok=True)
    return os.path.join(output_dir, filename)

def to_yn(val):
    if pd.isna(val) or val is None:
        return None
    val_str = str(val).strip()
    if val_str in ("1", "1.0", "Y", "y", "True", "true", "YES", "yes"):
        return "Y"
    elif val_str in ("0", "0.0", "N", "n", "False", "false", "NO", "no"):
        return "N"
    return None

def clean_val(val, as_int=False):
    if pd.isna(val) or val is None or str(val).strip() == "" or str(val).strip().upper() in ("NULL", "NONE", "NAN"):
        return None
    s = str(val).strip()
    if as_int:
        if s.endswith(".0"):
            s = s[:-2]
        if s.isdigit():
            return str(int(s))
    return s

def clean_memo(val):
    """Cleans LineMemo for Workday EIB mapping.
    If the value is None, NaN, empty/whitespace, or 'NULL'/'NONE'/'NAN' (case-insensitive),
    returns None so the target cell under 'Memo' column remains strictly blank.
    """
    if pd.isna(val) or val is None:
        return None
    s = str(val).strip()
    if s == "" or s.upper() in ("NULL", "NONE", "NAN", "N/A"):
        return None
    return s

def clean_amount(val):
    """Populates amounts including zero (0.0), matching the source file."""
    if pd.isna(val) or val is None or str(val).strip() == "" or str(val).strip().upper() in ("NULL", "NONE", "NAN"):
        return None
    try:
        if isinstance(val, str):
            val = val.replace(",", "").replace("$", "").strip()
        return round(float(val), 2)
    except (ValueError, TypeError):
        return None

def clean_ledger_account(val):
    if pd.isna(val) or val is None or str(val).strip() == "" or str(val).strip().upper() in ("NULL", "NONE", "NAN"):
        return None
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    # Keep alphanumeric codes with letters or underscores as strings
    if s.isdigit():
        try:
            return int(s)
        except ValueError:
            return s
    return s

def clean_cost_center(val):
    """Formats Cost Center as a 3-digit number/string (e.g., '000', '090', '070', '200')."""
    if pd.isna(val) or val is None or str(val).strip() == "" or str(val).strip().upper() in ("NULL", "NONE", "NAN"):
        return None
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    try:
        num = int(s)
        return f"{num:03d}"
    except ValueError:
        return s

def parse_date(val):
    if pd.isna(val) or val is None or str(val).strip() == "":
        return None
    if isinstance(val, (datetime.datetime, datetime.date)):
        return val.date() if isinstance(val, datetime.datetime) else val
    val_str = str(val).strip().split()[0]
    for sep in ("-", "/"):
        if sep in val_str:
            parts = val_str.split(sep)
            if len(parts) == 3 and all(p.isdigit() for p in parts):
                try:
                    if len(parts[0]) == 4:
                        return datetime.date(int(parts[0]), int(parts[1]), int(parts[2]))
                    if len(parts[2]) == 4:
                        return datetime.date(int(parts[2]), int(parts[0]), int(parts[1]))
                except Exception:
                    pass
    return val_str

# ==============================================================================
# DYNAMIC SOURCE & TARGET SCHEMA ANALYZER + AUTO-MAPPER
# ==============================================================================
def _norm_token(name):
    if name is None:
        return ""
    s = str(name).strip().rstrip("*+").strip().lower()
    return re.sub(r"[^a-z0-9]+", "", s)

def _strip_workday_token(name):
    if name is None:
        return ""
    s = str(name).strip().rstrip("*+").strip().lower()
    for pfx in ("worktag_", "optionalbalancingworktags_", "optional_balancing_worktags_", "line_", "journal_", "header_"):
        if s.startswith(pfx):
            s = s[len(pfx):]
    for sfx in ("_reference_id", "_ref_id", "_id", "_code", "_ref"):
        if s.endswith(sfx):
            s = s[:-len(sfx)]
    return re.sub(r"[^a-z0-9]+", "", s)

def _tokenize_words(name):
    if name is None:
        return set()
    s = str(name).strip().rstrip("*+")
    s = re.sub(r"([a-z0-9])([A-Z])", r"\1 \2", s)
    parts = [p.lower() for p in re.split(r"[^a-zA-Z0-9]+", s) if p]
    drop = {"worktag", "optionalbalancingworktags", "optional", "balancing", "reference", "id", "ref"}
    filtered = {p for p in parts if p not in drop}
    return filtered or set(parts)

_WORKDAY_FIELD_ALIASES = {
    "intercompanyaffiliate": [
        "Worktag_Intercompany_Affiliate_ID", "IntercompanyAffiliate", "Intercompany Affiliate",
        "Intercompany_Affiliate_ID", "Intercompany_Affiliate", "Worktag_Affiliate_ID",
        "Affiliate", "Affiliate_ID", "AffiliateID", "Intercompany_Company_ID",
        "IntercompanyCompany", "Intercompany", "IC_Affiliate", "ICAffiliate",
        "TradingPartner", "Trading_Partner", "Trading Partner", "TradingPartnerID"
    ],
    "balancingworktagaffiliate": [
        "Worktag_Business_Unit_ID", "OptionalBalancingWorktags_Business_Unit_ID",
        "BalancingWorktagAffiliate", "Balancing Worktag Affiliate", "Balancing_Worktag_Affiliate",
        "BusinessUnit", "Business Unit", "Business_Unit_ID"
    ],
    "balancingworktag": [
        "BalancingWorktag", "Balancing Worktag", "Header_Business_Unit_ID",
        "OptionalBalancingWorktags_Business_Unit_ID", "Worktag_Business_Unit_ID"
    ],
    "costcenter": [
        "Worktag_Cost_Center_Reference_ID", "CostCenter", "Cost Center",
        "Cost_Center_Reference_ID", "Cost_Center_ID", "CostCenterID"
    ],
    "spendcategory": [
        "Worktag_Spend_Category_ID", "SpendCategory", "Spend Category", "Spend_Category_ID"
    ],
    "revenuecategory": [
        "Worktag_Revenue_Category_ID", "RevenueCategory", "Revenue Category", "Revenue_Category_ID"
    ],
    "customorganization": [
        "Worktag_Custom_Organization_Reference_ID", "CustomOrganization",
        "Custom Organization", "Custom_Organization_Reference_ID"
    ],
    "region": [
        "Worktag_Region_Reference_ID", "Region", "Region_Reference_ID", "RegionID"
    ],
    "bankaccount": [
        "Worktag_Bank_Account_ID", "BankAccount", "Bank Account", "Bank_Account_ID"
    ],
    "loan": [
        "Worktag_Loan_ID", "Loan", "Loan_ID"
    ],
    "customworktag1": ["Worktag_Custom_Worktag_1_ID", "CustomWorktag1", "Custom Worktag 1"],
    "customworktag2": ["Worktag_Custom_Worktag_2_ID", "CustomWorktag2", "Custom Worktag 2"],
    "customworktag3": ["Worktag_Custom_Worktag_3_ID", "CustomWorktag3", "Custom Worktag 3"],
    "customworktag4": ["Worktag_Custom_Worktag_4_ID", "CustomWorktag4", "Custom Worktag 4"],
    "customworktag5": ["Worktag_Custom_Worktag_5_ID", "CustomWorktag5", "Custom Worktag 5"],
    "supplier": ["Worktag_Supplier_ID", "Supplier", "Supplier_ID", "SupplierReferenceID", "Vendor", "VendorID"],
    "customer": ["Worktag_Customer_ID", "Customer", "Customer_ID", "CustomerReferenceID"],
    "fund": ["Worktag_Fund_ID", "Fund", "Fund_ID"],
    "grant": ["Worktag_Grant_ID", "Grant", "Grant_ID"],
    "program": ["Worktag_Program_ID", "Program", "Program_ID"],
    "project": ["Worktag_Project_ID", "Project", "Project_ID"],
    "location": ["Worktag_Location_ID", "Location", "Location_ID"],
    "taxcode": ["TaxCode", "Tax Code", "Tax_Code_ID"],
    "taxrate": ["TaxRate", "Tax Rate", "Tax_Rate_ID"],
    "taxtype": ["TaxType", "Tax Type", "Tax_Type_ID"],
    "taxapplicability": ["TaxApplicability", "Tax Applicability", "Tax_Applicability_ID"],
    "documentlink": ["DocumentLink", "Document Link"],
    "bookcode": ["BookCode", "Book Code", "Book_Code_ID"],
    "journalnumber": ["JournalNumber", "Journal Number"],
    "adjustmentjournal": ["AdjustmentJournal", "Adjustment Journal"],
    "includetaxlines": ["IncludeTaxLines", "Include Tax Lines"],
    "createreversal": ["CreateReversal", "Create Reversal"],
    "reversaldate": ["ReversalDate", "Reversal Date"],
    "controltotalamount": ["ControlTotalAmount", "Control Total Amount"],
    "currencyratetype": ["CurrencyRateType", "Currency Rate Type", "Currency_Rate_Type_ID"],
    "currencyrate": ["CurrencyRate", "Currency Rate"],
    "quantity": ["Quantity", "Qty"],
    "unitofmeasure": ["UnitOfMeasure", "Unit of Measure", "UOM"],
    "excludefromspendreport": ["ExcludeFromSpendReport", "Exclude from Spend Report"],
    "externalbusinessentityname": ["ExternalBusinessEntityName", "External Business Entity Name"],
    "externalbusinessentityaddress": ["ExternalBusinessEntityAddress", "External Business Entity Address"],
    "externalbusinessentitytaxid": ["ExternalBusinessEntityTaxID", "External Business Entity Tax ID"],
    "externalsupplierinvoicesource": ["ExternalSupplierInvoiceSource", "External Supplier Invoice Source", "External_Supplier_Invoice_Source_ID"],
}

_CORE_SOURCE_ALIASES = {
    "headerkey": ["HeaderKey", "Header Key", "Header_Key", "SpreadsheetKey", "Spreadsheet Key"],
    "journalkey": ["AccountingJournalID", "Accounting Journal ID", "JournalKey", "Journal Key", "JournalID", "Journal ID"],
    "addonly": ["AddOnlyJournal", "Add Only Journal", "Add Only", "AddOnly"],
    "submit": ["Submit", "SubmitJournal"],
    "company": ["CompanyReferenceID", "Company Reference ID", "Company_Reference_ID", "Company", "CompanyCode", "Company Code", "LegalEntity"],
    "currency": ["Currency", "CurrencyCode", "Currency Code", "Currency_ID"],
    "ledgertype": ["LedgerType", "Ledger Type", "Ledger_Type_ID"],
    "bookcode": ["BookCode", "Book Code", "Book_Code_ID"],
    "accountingdate": ["AccountingDate", "Accounting Date", "PostingDate", "Posting Date", "TransactionDate", "JournalDate"],
    "journalsource": ["JournalSource", "Journal Source"],
    "journalentrymemo": ["JournalEntryMemo", "Journal Entry Memo", "HeaderMemo", "Header Memo"],
    "journalextref": ["JournalExternalReferenceID", "Journal External Reference ID", "HeaderExternalReferenceID"],
    "journalnumber": ["JournalNumber", "Journal Number"],
    "row_id": ["ROW_ID", "RowID", "Row ID", "LineKey", "Line Key"],
    "lineorder": ["LineOrder", "Line Order", "ROW_ID", "RowID"],
    "linecompany": ["LineCompanyReferenceID", "Line Company Reference ID", "LineCompany", "Line Company", "CompanyReferenceID", "Company"],
    "ledgeraccount": ["LedgerAccountReferenceID", "Ledger Account Reference ID", "LedgerAccount", "Ledger Account", "GLAccount", "GL Account", "Account", "AccountNumber"],
    "accountset": ["LedgerAccountReferenceID_ParentID", "AccountSet", "Account Set", "Account_Set_ID"],
    "debitamount": ["DebitAmount", "Debit Amount", "Debit", "DrAmount", "Dr"],
    "creditamount": ["CreditAmount", "Credit Amount", "Credit", "CrAmount", "Cr"],
    "linecurrency": ["LineCurrency", "Line Currency", "Currency", "CurrencyCode"],
    "linememo": ["LineMemo", "Line Memo", "Memo", "LineDescription", "Description"],
    "lineextref": ["JournalLineExternalReferenceID", "Journal Line External Reference ID", "LineExternalReferenceID"],
    "budgetdate": ["BudgetDate", "Budget Date"],
    "revenuecategory": _WORKDAY_FIELD_ALIASES["revenuecategory"],
    "spendcategory": _WORKDAY_FIELD_ALIASES["spendcategory"],
    "costcenter": _WORKDAY_FIELD_ALIASES["costcenter"],
    "intercompanyaffiliate": _WORKDAY_FIELD_ALIASES["intercompanyaffiliate"],
    "balancingworktagaffiliate": _WORKDAY_FIELD_ALIASES["balancingworktagaffiliate"],
    "customorganization": _WORKDAY_FIELD_ALIASES["customorganization"],
    "customworktag1": _WORKDAY_FIELD_ALIASES["customworktag1"],
}

def _build_col_lookup(source_columns):
    """Maps raw, normalized, and stripped source column tokens to exact source column header strings."""
    exact = {str(c): c for c in source_columns}
    norm = {}
    stripped = {}
    for c in source_columns:
        nt = _norm_token(c)
        if nt and nt not in norm:
            norm[nt] = c
        st = _strip_workday_token(c)
        if st and st not in stripped:
            stripped[st] = c
    return {"exact": exact, "norm": norm, "stripped": stripped}

def _resolve_source_column(col_lookup, candidates):
    for cand in candidates:
        if cand in col_lookup["exact"]:
            return col_lookup["exact"][cand]
    for cand in candidates:
        nt = _norm_token(cand)
        if nt in col_lookup["norm"]:
            return col_lookup["norm"][nt]
    return None

def _pick_source_val(row, col_lookup, core_key):
    col_name = _resolve_source_column(col_lookup, _CORE_SOURCE_ALIASES.get(core_key, []))
    if col_name is not None:
        return row.get(col_name)
    return None

def _inspect_target_sheet(ws):
    """Reads Row 3 (restrictions), Row 4 (Workday format ID), Row 5 (column header) for every column."""
    cols = []
    for c in range(1, ws.max_column + 1):
        r3 = ws.cell(row=3, column=c).value
        r4 = ws.cell(row=4, column=c).value
        r5 = ws.cell(row=5, column=c).value
        if r5 is None and r4 is None:
            continue
        cols.append({
            "col_idx": c,
            "col_letter": openpyxl.utils.get_column_letter(c),
            "restriction": str(r3).strip() if r3 is not None else "",
            "format": str(r4).strip() if r4 is not None else "",
            "header": str(r5).strip() if r5 is not None else "",
            "norm": _norm_token(r5),
            "stripped_r4": _strip_workday_token(r4),
            "words": _tokenize_words(r5),
        })
    return cols

_IGNORED_TARGET_HEADERS = {"fields", "donotuse", "companieswithoutaccess", "filecontent"}
_AMBIGUOUS_R4_FORMATS = {"text", "yn", "yyyymmdd", "company", "externalsupplierinvoicesource", "accountset", "ledgeraccount", "currency", "currencyratetype"}
_SKIP_SOURCE_SUFFIXES = ("referenceidtype", "parentidtype", "idtype")

def discover_sheet_mappings(source_columns, ws, handled_target_cols, handled_source_cols, is_header_sheet=False):
    """
    Analyzes both source columns and target template columns (Row 5 header + Row 4 format)
    and automatically discovers mappings for any new or previously unmapped source/target columns.
    """
    target_cols = _inspect_target_sheet(ws)
    col_lookup = _build_col_lookup(source_columns)

    # Track Row 4 format frequencies so only unambiguous Row 4 IDs participate in format matching
    r4_counts = {}
    for tc in target_cols:
        sr4 = tc["stripped_r4"]
        if sr4:
            r4_counts[sr4] = r4_counts.get(sr4, 0) + 1

    # On Header sheet, avoid accidentally matching line-level columns into header slots
    header_disallowed_sources = {
        "row_id", "rowid", "lineorder", "debitamount", "creditamount", "linememo",
        "linecurrency", "ledgeraccountreferenceid", "ledgeraccountreferenceidparentid",
        "budgetdate", "linecompanyreferenceid"
    }

    discovered = []
    claimed_targets = set(handled_target_cols)
    claimed_sources = set(handled_source_cols)

    # Skip alternate ledger block unless source explicitly mentions 'alternate'
    has_alternate_source = any("alternate" in str(s).lower() for s in source_columns)

    for tc in target_cols:
        c_idx = tc["col_idx"]
        t_norm = tc["norm"]
        if c_idx in claimed_targets or not t_norm or t_norm in _IGNORED_TARGET_HEADERS:
            continue
        if not has_alternate_source and t_norm.startswith("alternateledger"):
            continue
        # Skip trailing duplicate Workday Tax/Attachment Row ID & Line Order columns after primary block
        if t_norm in ("rowid", "lineorder") and c_idx > 40:
            continue

        matched_src = None
        match_method = None

        # 1. Alias dictionary lookup (exact Workday field variants)
        alias_cands = _WORKDAY_FIELD_ALIASES.get(t_norm, [])
        if alias_cands:
            cand_src = _resolve_source_column(col_lookup, alias_cands)
            if cand_src and cand_src not in claimed_sources:
                if not (is_header_sheet and _norm_token(cand_src) in header_disallowed_sources):
                    matched_src = cand_src
                    match_method = "alias"

        # 2. Direct normalized / stripped Worktag match against Row 5 header
        if not matched_src:
            for src in source_columns:
                if src in claimed_sources:
                    continue
                s_low = str(src).strip().lower()
                if any(s_low.endswith(sfx) for sfx in _SKIP_SOURCE_SUFFIXES):
                    continue
                s_norm = _norm_token(src)
                if is_header_sheet and (s_norm in header_disallowed_sources or s_low.startswith("worktag_")):
                    continue
                s_strip = _strip_workday_token(src)
                if s_norm == t_norm or (s_strip and s_strip == t_norm):
                    matched_src = src
                    match_method = "header_norm"
                    break

        # 3. Unique Row 4 Format ID match (e.g. Region_Reference_ID <-> Worktag_Region_Reference_ID)
        sr4 = tc["stripped_r4"]
        if not matched_src and sr4 and sr4 not in _AMBIGUOUS_R4_FORMATS and r4_counts.get(sr4, 0) == 1:
            for src in source_columns:
                if src in claimed_sources:
                    continue
                s_low = str(src).strip().lower()
                if any(s_low.endswith(sfx) for sfx in _SKIP_SOURCE_SUFFIXES):
                    continue
                if is_header_sheet and s_low.startswith("worktag_"):
                    continue
                if _strip_workday_token(src) == sr4:
                    matched_src = src
                    match_method = "format_r4"
                    break

        # 4. Unambiguous multi-word token overlap for custom user columns
        if not matched_src and len(tc["words"]) >= 2:
            for src in source_columns:
                if src in claimed_sources:
                    continue
                s_low = str(src).strip().lower()
                if any(s_low.endswith(sfx) for sfx in _SKIP_SOURCE_SUFFIXES):
                    continue
                if is_header_sheet and (_norm_token(src) in header_disallowed_sources or s_low.startswith("worktag_")):
                    continue
                s_words = _tokenize_words(src)
                if len(s_words) >= 2 and (s_words == tc["words"]):
                    matched_src = src
                    match_method = "token_set"
                    break

        if matched_src:
            claimed_targets.add(c_idx)
            claimed_sources.add(matched_src)
            discovered.append({
                "source_col": matched_src,
                "target_sheet": ws.title,
                "target_col_idx": c_idx,
                "target_col_letter": tc["col_letter"],
                "target_header": tc["header"],
                "target_format": tc["format"],
                "match_method": match_method,
            })

    return discovered

def _coerce_auto_val(val, target_format):
    fmt = (target_format or "").lower()
    if fmt == "y/n":
        return to_yn(val) or clean_val(val)
    if "yyyy-mm-dd" in fmt:
        return parse_date(val)
    if fmt.startswith("number"):
        return clean_amount(val)
    return clean_val(val)

def _resolve_handled_src(col_lookup, keys):
    return {
        c for k in keys
        for c in [_resolve_source_column(col_lookup, _CORE_SOURCE_ALIASES[k])] if c
    }

def _plan_two_sheet_mappings(source_cols, ws_hdr, ws_line, col_lookup):
    """Builds shared dynamic mapping plan for two-sheet v45.1 workbooks."""
    line_hdr_cols = {
        str(ws_line.cell(row=5, column=c).value).strip().rstrip("*+"): c
        for c in range(1, ws_line.max_column + 1)
        if ws_line.cell(row=5, column=c).value is not None
    }
    col_intercompany = line_hdr_cols.get("Intercompany Affiliate")
    col_balancing_affiliate = line_hdr_cols.get("Balancing Worktag Affiliate", 31)

    handled_hdr_tgt = {2, 3, 5, 6, 12, 13, 14, 16, 17, 21, 22, 23, 33}
    handled_hdr_src = _resolve_handled_src(col_lookup, (
        "headerkey", "addonly", "journalkey", "company", "currency", "ledgertype",
        "accountingdate", "journalsource", "journalentrymemo", "journalextref"
    ))
    hdr_auto = discover_sheet_mappings(source_cols, ws_hdr, handled_hdr_tgt, handled_hdr_src, is_header_sheet=True)

    handled_line_tgt = {2, 3, 4, 5, 6, 7, 10, 11, 12, 20, 21, 22, 23, 25, 27}
    if col_intercompany:
        handled_line_tgt.add(col_intercompany)
    if col_balancing_affiliate:
        handled_line_tgt.add(col_balancing_affiliate)
    handled_line_src = _resolve_handled_src(col_lookup, (
        "headerkey", "row_id", "lineorder", "linecompany", "ledgeraccount", "accountset",
        "debitamount", "creditamount", "linecurrency", "linememo", "lineextref",
        "budgetdate", "revenuecategory", "spendcategory", "costcenter",
        "intercompanyaffiliate", "balancingworktagaffiliate"
    ))
    line_auto = discover_sheet_mappings(source_cols, ws_line, handled_line_tgt, handled_line_src, is_header_sheet=False)
    return col_intercompany, col_balancing_affiliate, hdr_auto, line_auto

def _plan_single_sheet_mappings(source_cols, ws, col_lookup):
    """Builds shared dynamic mapping plan for single-sheet v46/v47/v48 workbooks."""
    handled_tgt = {2, 3, 4, 5, 8, 9, 13, 15, 16, 17, 18, 19, 20, 24, 25, 26, 35, 36, 37, 38, 39, 42, 43, 44, 52, 53, 54, 55, 57, 59, 61, 63, 65}
    handled_src = _resolve_handled_src(col_lookup, (
        "headerkey", "addonly", "journalkey", "submit", "journalnumber", "company", "currency",
        "ledgertype", "bookcode", "accountingdate", "journalsource", "journalentrymemo",
        "journalextref", "row_id", "lineorder", "linecompany", "ledgeraccount", "accountset",
        "debitamount", "creditamount", "linecurrency", "linememo", "lineextref", "budgetdate",
        "costcenter", "spendcategory", "balancingworktagaffiliate", "customorganization",
        "customworktag1", "revenuecategory"
    ))
    return discover_sheet_mappings(source_cols, ws, handled_tgt, handled_src, is_header_sheet=False)

def analyze_schema(source_file, excel_template=None):
    """Inspects both source columns and target template columns and returns mapping discovery summary."""
    resolved_src = find_source_file(source_file)
    if not resolved_src or not os.path.exists(resolved_src):
        resolved_src = str(source_file)
    tpl_path = resolve_target_file(excel_template) if excel_template else resolve_target_file("v45")

    df = load_source_dataframe(resolved_src)
    source_cols = list(df.columns) if hasattr(df, "columns") and df.columns else (list(df.iloc[0].keys()) if len(df) > 0 else [])

    populated_cols = [
        c for c in source_cols
        if any(pd.notna(r.get(c)) and str(r.get(c)).strip() != "" for _, r in df.iterrows())
    ]

    wb = openpyxl.load_workbook(tpl_path, data_only=True, read_only=True)
    is_two_sheet = (SHEET_HEADER in wb.sheetnames and SHEET_LINES in wb.sheetnames)
    col_lookup = _build_col_lookup(source_cols)

    auto_mappings = []
    target_col_summary = {}

    if is_two_sheet:
        ws_hdr = wb[SHEET_HEADER]
        ws_line = wb[SHEET_LINES]
        target_col_summary[SHEET_HEADER] = ws_hdr.max_column
        target_col_summary[SHEET_LINES] = ws_line.max_column
        _, _, hdr_auto, line_auto = _plan_two_sheet_mappings(source_cols, ws_hdr, ws_line, col_lookup)
        auto_mappings.extend(hdr_auto)
        auto_mappings.extend(line_auto)
    else:
        sheet_name = "Submit Accounting Journal" if "Submit Accounting Journal" in wb.sheetnames else wb.sheetnames[0]
        ws = wb[sheet_name]
        target_col_summary[sheet_name] = ws.max_column
        auto_mappings.extend(_plan_single_sheet_mappings(source_cols, ws, col_lookup))

    wb.close()

    mapped_src_set = {
        c for k in _CORE_SOURCE_ALIASES
        for c in [_resolve_source_column(col_lookup, _CORE_SOURCE_ALIASES[k])] if c
    }
    mapped_src_set.update(m["source_col"] for m in auto_mappings)

    unmapped_populated = [
        c for c in populated_cols
        if c not in mapped_src_set and not any(str(c).lower().endswith(sfx) for sfx in _SKIP_SOURCE_SUFFIXES) and _norm_token(c) not in ("recordquantity", "accountingjournalreferenceid")
    ]

    interco_src = _resolve_source_column(col_lookup, _CORE_SOURCE_ALIASES["intercompanyaffiliate"])
    return {
        "source_file": resolved_src,
        "template_file": tpl_path,
        "template_version": _detect_template_version_tag(tpl_path),
        "total_rows": len(df),
        "source_columns": source_cols,
        "populated_source_columns": populated_cols,
        "target_sheets": target_col_summary,
        "intercompany_source_col": interco_src,
        "auto_discovered_mappings": auto_mappings,
        "unmapped_populated_columns": unmapped_populated,
    }

def select_file_interactive(prompt_title, options, default_index=0):
    """Displays a numbered menu in terminal and returns the selected option."""
    print(f"\n{prompt_title}")
    for idx, opt in enumerate(options, 1):
        extra = ""
        if os.path.exists(opt):
            sz_kb = round(os.path.getsize(opt) / 1024, 1)
            extra = f" ({sz_kb} KB)"
        print(f"  [{idx}] {opt}{extra}")

    default_val = options[default_index]

    # Non-interactive check (pipes, automated tasks, server subagent)
    if not sys.stdin.isatty():
        print(f">> Non-interactive terminal: auto-selected [{default_index + 1}] {default_val}")
        return default_val

    try:
        user_input = input(f">> Enter number [1-{len(options)}] (default {default_index + 1}): ").strip()
        if not user_input:
            print(f">> Selected default: {default_val}")
            return default_val
        if user_input.isdigit() and 1 <= int(user_input) <= len(options):
            chosen = options[int(user_input) - 1]
            print(f">> Selected: {chosen}")
            return chosen
        for opt in options:
            if user_input.lower() in opt.lower():
                print(f">> Selected: {opt}")
                return opt
        print(f">> Unrecognized input '{user_input}'. Using default: {default_val}")
        return default_val
    except (EOFError, KeyboardInterrupt):
        print(f"\n>> Using default: {default_val}")
        return default_val

def load_source_dataframe(source_path):
    """Loads a single source Excel or CSV file into a pandas DataFrame."""
    if not os.path.exists(source_path):
        raise FileNotFoundError(f"Cannot find source file: {source_path}")

    print(f"Reading source file: {source_path}...")
    if source_path.lower().endswith(".csv"):
        df = pd.read_csv(source_path)
    else:
        df = pd.read_excel(source_path)
    print(f"Loaded {len(df)} records from {source_path}.")
    return df

# ==============================================================================
# CODE FOR 2026 MAPPING
# ==============================================================================
def _safe_int_or_val(val, fallback):
    if val is None or pd.isna(val):
        return fallback
    s = str(val).strip()
    if s.endswith(".0") and s[:-2].isdigit():
        return int(s[:-2])
    if s.isdigit():
        return int(s)
    return s or fallback

def process_journal_2026(source_file, excel_template, output_file=None, output_dir=None):
    # Load Spend Category lookup map
    spend_map = load_spend_category_map()

    # Determine date-wise or user-selected output path
    custom_output = output_file is not None
    datewise_target = get_datewise_output_path(source_file, target_dir=output_dir, template_file=excel_template)
    if output_file is None:
        output_file = datewise_target

    if not os.path.exists(excel_template):
        raise FileNotFoundError(f"Cannot find target template: {excel_template}")

    # Check lock on destination
    try:
        if os.path.exists(output_file):
            with open(output_file, "r+b"):
                pass
    except (PermissionError, FileNotFoundError):
        print(f"\n[ERROR] '{output_file}' is currently open in Microsoft Excel.")
        print("Please close Microsoft Excel and re-run.\n")
        raise PermissionError(f"'{output_file}' is open in Excel.")

    df = load_source_dataframe(source_file)
    source_cols = list(df.columns) if hasattr(df, "columns") and df.columns else (list(df.iloc[0].keys()) if len(df) > 0 else [])
    col_lookup = _build_col_lookup(source_cols)

    print(f"Loading base Excel template: {excel_template}...")
    wb = openpyxl.load_workbook(excel_template)

    is_two_sheet = (SHEET_HEADER in wb.sheetnames and SHEET_LINES in wb.sheetnames)

    if is_two_sheet:
        print(f"Detected Workday EIB Template: v45.1 Two-Sheet ('{SHEET_HEADER}' & '{SHEET_LINES}')")
        ws_hdr = wb[SHEET_HEADER]
        ws_line = wb[SHEET_LINES]

        if ws_hdr.max_row >= 6:
            print(f"Clearing Sheet '{SHEET_HEADER}' data rows 6 to {ws_hdr.max_row}...")
            ws_hdr.delete_rows(6, ws_hdr.max_row - 5)

        if ws_line.max_row >= 6:
            print(f"Clearing Sheet '{SHEET_LINES}' data rows 6 to {ws_line.max_row}...")
            ws_line.delete_rows(6, ws_line.max_row - 5)

        # Resolve Header & Line column maps dynamically from Row 5 headers & auto-discover extra columns
        col_intercompany, col_balancing_affiliate, hdr_auto_mappings, line_auto_mappings = _plan_two_sheet_mappings(
            source_cols, ws_hdr, ws_line, col_lookup
        )
        all_auto_mappings = hdr_auto_mappings + line_auto_mappings

        if all_auto_mappings:
            print(f"Auto-Discovered {len(all_auto_mappings)} Additional Column Mapping(s):")
            for am in all_auto_mappings:
                print(f"  + [{am['target_sheet']}] '{am['source_col']}' -> Col {am['target_col_letter']} ({am['target_header']}) via {am['match_method']}")

        # Identify unique journals while preserving appearance order
        unique_headers = []
        seen_keys = []
        for _, row in df.iterrows():
            hk_raw = _pick_source_val(row, col_lookup, "headerkey")
            jk_raw = _pick_source_val(row, col_lookup, "journalkey")
            key = hk_raw if pd.notna(hk_raw) else (jk_raw or "1")
            if key not in seen_keys:
                seen_keys.append(key)
                unique_headers.append(row)

        hdr_row = 6
        header_key_map = {}
        for h_row in unique_headers:
            hk_raw = _pick_source_val(h_row, col_lookup, "headerkey")
            jk_raw = _pick_source_val(h_row, col_lookup, "journalkey")
            h_key = hk_raw if pd.notna(hk_raw) else (jk_raw or "1")
            hk_val = _safe_int_or_val(hk_raw, hdr_row - 5) if pd.notna(hk_raw) else (hdr_row - 5)
            header_key_map[h_key] = hk_val

            # Col B (2): Header Key*
            ws_hdr.cell(row=hdr_row, column=2, value=hk_val)

            # Col C (3): Add Only
            ws_hdr.cell(row=hdr_row, column=3, value=to_yn(_pick_source_val(h_row, col_lookup, "addonly")) or "Y")

            # Col F (6): Accounting Journal ID
            j_id = clean_val(jk_raw)
            ws_hdr.cell(row=hdr_row, column=6, value=j_id)

            # Col E (5): Accounting Journal -> Strictly left blank
            ws_hdr.cell(row=hdr_row, column=5, value=None)

            # Col L (12): Company*
            ws_hdr.cell(row=hdr_row, column=12, value=clean_val(_pick_source_val(h_row, col_lookup, "company")) or "LE01")

            # Col M (13): Currency*
            ws_hdr.cell(row=hdr_row, column=13, value=clean_val(_pick_source_val(h_row, col_lookup, "currency")) or "USD")

            # Col N (14): Ledger Type*
            lt = clean_val(_pick_source_val(h_row, col_lookup, "ledgertype")) or "Actuals"
            ws_hdr.cell(row=hdr_row, column=14, value="Actuals" if lt.upper() == "ACTUALS" else lt)

            # Col P (16): Accounting Date*
            ws_hdr.cell(row=hdr_row, column=16, value=parse_date(_pick_source_val(h_row, col_lookup, "accountingdate")))

            # Col Q (17): Journal Source*
            ws_hdr.cell(row=hdr_row, column=17, value=clean_val(_pick_source_val(h_row, col_lookup, "journalsource")) or "Manual_Journal")

            # Col U (21): Record Quantity -> Strictly left blank
            ws_hdr.cell(row=hdr_row, column=21, value=None)

            # Col V (22): Journal Entry Memo
            ws_hdr.cell(row=hdr_row, column=22, value=clean_val(_pick_source_val(h_row, col_lookup, "journalentrymemo")))

            # Col W (23): External Reference ID
            ws_hdr.cell(row=hdr_row, column=23, value=clean_val(_pick_source_val(h_row, col_lookup, "journalextref")))

            # Col AG (33): Auto Complete
            ws_hdr.cell(row=hdr_row, column=33, value="Y")

            # Auto-discovered header columns
            for am in hdr_auto_mappings:
                val = _coerce_auto_val(h_row.get(am["source_col"]), am["target_format"])
                if val is not None:
                    ws_hdr.cell(row=hdr_row, column=am["target_col_idx"], value=val)

            hdr_row += 1

        line_row = 6
        total_debits = 0.0
        total_credits = 0.0
        spend_matched_count = 0
        interco_mapped_count = 0

        for idx, row in df.iterrows():
            hk_raw = _pick_source_val(row, col_lookup, "headerkey")
            jk_raw = _pick_source_val(row, col_lookup, "journalkey")
            r_key = hk_raw if pd.notna(hk_raw) else (jk_raw or "1")
            if pd.notna(hk_raw):
                hk_val = _safe_int_or_val(hk_raw, header_key_map.get(r_key, 1))
            else:
                hk_val = header_key_map.get(r_key, 1)

            ws_line.cell(row=line_row, column=2, value=hk_val)

            raw_rid = _pick_source_val(row, col_lookup, "row_id")
            line_key_val = _safe_int_or_val(raw_rid, idx + 1) if pd.notna(raw_rid) else (idx + 1)
            ws_line.cell(row=line_row, column=3, value=line_key_val)

            raw_lo = _pick_source_val(row, col_lookup, "lineorder")
            line_order_val = _safe_int_or_val(raw_lo, line_key_val) if pd.notna(raw_lo) else line_key_val
            ws_line.cell(row=line_row, column=4, value=line_order_val)

            ws_line.cell(row=line_row, column=5, value=clean_val(_pick_source_val(row, col_lookup, "linecompany")) or "LE01")

            acct_clean = clean_ledger_account(_pick_source_val(row, col_lookup, "ledgeraccount"))
            ws_line.cell(row=line_row, column=6, value=acct_clean)
            ws_line.cell(row=line_row, column=7, value=clean_val(_pick_source_val(row, col_lookup, "accountset")) or "ACCOUNT_SET")

            deb_val = clean_amount(_pick_source_val(row, col_lookup, "debitamount"))
            crd_val = clean_amount(_pick_source_val(row, col_lookup, "creditamount"))

            if deb_val is not None:
                ws_line.cell(row=line_row, column=10, value=deb_val)
                total_debits += deb_val

            if crd_val is not None:
                ws_line.cell(row=line_row, column=11, value=crd_val)
                total_credits += crd_val

            ws_line.cell(row=line_row, column=12, value=clean_val(_pick_source_val(row, col_lookup, "linecurrency")) or "USD")

            # Col T (20): Memo - mapped directly from LineMemo. If empty or "NULL", left blank with no value (no fallback)
            memo = clean_memo(_pick_source_val(row, col_lookup, "linememo"))
            ws_line.cell(row=line_row, column=20, value=memo)

            ws_line.cell(row=line_row, column=21, value=clean_val(_pick_source_val(row, col_lookup, "lineextref")))
            ws_line.cell(row=line_row, column=22, value=parse_date(_pick_source_val(row, col_lookup, "budgetdate")))
            ws_line.cell(row=line_row, column=23, value=clean_val(_pick_source_val(row, col_lookup, "revenuecategory")))

            # Map Spend Category using Spend Category Mapped Ledger Account.xlsx
            acct_key = str(acct_clean).strip() if acct_clean is not None else ""
            mapped_sc = spend_map.get(acct_key) or clean_val(_pick_source_val(row, col_lookup, "spendcategory"))
            ws_line.cell(row=line_row, column=25, value=mapped_sc)
            if mapped_sc:
                spend_matched_count += 1

            cc_3digit = clean_cost_center(_pick_source_val(row, col_lookup, "costcenter"))
            cell_cc = ws_line.cell(row=line_row, column=27, value=cc_3digit)
            if cc_3digit is not None:
                cell_cc.number_format = "@"

            if col_intercompany:
                interco_val = clean_val(_pick_source_val(row, col_lookup, "intercompanyaffiliate"))
                ws_line.cell(row=line_row, column=col_intercompany, value=interco_val)
                if interco_val is not None:
                    interco_mapped_count += 1

            ws_line.cell(row=line_row, column=col_balancing_affiliate, value=clean_val(_pick_source_val(row, col_lookup, "balancingworktagaffiliate")))

            # Auto-discovered line columns (any extra source or target columns)
            for am in line_auto_mappings:
                val = _coerce_auto_val(row.get(am["source_col"]), am["target_format"])
                if val is not None:
                    ws_line.cell(row=line_row, column=am["target_col_idx"], value=val)

            line_row += 1

        total_journals = len(unique_headers)
        total_lines_written = len(df)
        variance = round(total_debits - total_credits, 2)

        print(f"\n[2026 MAPPING SUMMARY] Successfully mapped '{source_file}':")
        print(f"  - Sheet 1 ('{SHEET_HEADER}'): {total_journals} journal header rows written (Rows 6 to {hdr_row - 1})")
        print(f"  - Sheet 2 ('{SHEET_LINES}'): {total_lines_written} journal line rows written (Rows 6 to {line_row - 1})")
        print(f"  - Spend Categories Mapped: {spend_matched_count} / {total_lines_written} lines from '{SPEND_MAPPING_FILE}'")
        if col_intercompany:
            print(f"  - Intercompany Affiliate Mapped: {interco_mapped_count} / {total_lines_written} lines (Col {openpyxl.utils.get_column_letter(col_intercompany)})")
        print(f"  - Auto-Discovered Extra Mappings: {len(all_auto_mappings)}")
        print(f"  - Total Debits:  ${total_debits:,.2f}")
        print(f"  - Total Credits: ${total_credits:,.2f}")
        print(f"  - Net Variance:  ${variance:,.2f} ({'BALANCED' if variance == 0 else 'UNBALANCED'})")

    else:
        raise ValueError(f"Target '{excel_template}' must be a two-sheet template ('{SHEET_HEADER}' and '{SHEET_LINES}').")

    # Strictly enforce Column E (5) and Column U (21) blank in Import Accounting Journal sheet
    for r_idx in range(6, ws_hdr.max_row + 1):
        ws_hdr.cell(row=r_idx, column=5, value=None)
        ws_hdr.cell(row=r_idx, column=21, value=None)

    # Ensure target parent directory exists
    os.makedirs(os.path.dirname(output_file) if os.path.dirname(output_file) else ".", exist_ok=True)
    print(f"Saving workbook to output: {output_file}...")
    wb.save(output_file)

    # Only save synchronized root/output copies if no custom output_file was specified
    if not custom_output:
        for sync_path in [os.path.basename(output_file), os.path.join("output", os.path.basename(output_file))]:
            if os.path.abspath(output_file) != os.path.abspath(sync_path):
                try:
                    os.makedirs(os.path.dirname(sync_path) if os.path.dirname(sync_path) else ".", exist_ok=True)
                    wb.save(sync_path)
                    print(f"Also saved synchronized copy: {sync_path}")
                except Exception:
                    pass

    wb.close()
    print(f"SUCCESS: 2026 file mapped and saved.\n")

    return {
        "success": True,
        "template_version": "v45.1",
        "source_file": source_file,
        "target_file": output_file,
        "total_journals": total_journals,
        "total_lines_written": total_lines_written,
        "total_source_rows": len(df),
        "spend_categories_mapped": spend_matched_count,
        "intercompany_mapped_count": interco_mapped_count,
        "auto_mapped_columns": all_auto_mappings,
        "total_debit": round(total_debits, 2),
        "total_credit": round(total_credits, 2),
        "variance": variance,
        "is_balanced": variance == 0,
        "output_file": output_file
    }

# ==============================================================================
# CODE FOR 2025 MAPPING (UNIFIED WITH 2026 TWO-SHEET V45.1 MAPPING LOGIC)
# ==============================================================================
def process_journal_2025(source_file, excel_template, output_file=None, output_dir=None):
    """
    Uses the exact same two-sheet Import Accounting Journal (v45.1) mapping logic as 2026.
    """
    return process_journal_2026(source_file, excel_template, output_file=output_file, output_dir=output_dir)

# ==============================================================================
# CODE FOR SINGLE-SHEET TEMPLATES (v46.0, v47.0, v48.0)
# ==============================================================================
def process_journal_single_sheet(source_file, excel_template, output_file=None, output_dir=None):
    """Maps source journal entries into a Workday Submit Accounting Journal single-sheet template (v46, v47, v48)."""
    spend_map = load_spend_category_map()

    custom_output = output_file is not None
    datewise_target = get_datewise_output_path(source_file, target_dir=output_dir, template_file=excel_template)
    if output_file is None:
        output_file = datewise_target

    if not os.path.exists(excel_template):
        raise FileNotFoundError(f"Cannot find target template: {excel_template}")

    try:
        if os.path.exists(output_file):
            with open(output_file, "r+b"):
                pass
    except (PermissionError, FileNotFoundError):
        print(f"\n[ERROR] '{output_file}' is currently open in Microsoft Excel.")
        print("Please close Microsoft Excel and re-run.\n")
        raise PermissionError(f"'{output_file}' is open in Excel.")

    df = load_source_dataframe(source_file)
    source_cols = list(df.columns) if hasattr(df, "columns") and df.columns else (list(df.iloc[0].keys()) if len(df) > 0 else [])
    col_lookup = _build_col_lookup(source_cols)

    print(f"Loading single-sheet Excel template: {excel_template}...")
    wb = openpyxl.load_workbook(excel_template)

    sheet_name = "Submit Accounting Journal"
    if sheet_name not in wb.sheetnames:
        sheet_name = wb.sheetnames[0]
    ws = wb[sheet_name]

    if ws.max_row >= 6:
        print(f"Clearing Sheet '{sheet_name}' data rows 6 to {ws.max_row}...")
        ws.delete_rows(6, ws.max_row - 5)

    # Auto-discover any extra/new source or target columns beyond the core single-sheet positions
    auto_mappings = _plan_single_sheet_mappings(source_cols, ws, col_lookup)
    if auto_mappings:
        print(f"Auto-Discovered {len(auto_mappings)} Additional Column Mapping(s):")
        for am in auto_mappings:
            print(f"  + [{am['target_sheet']}] '{am['source_col']}' -> Col {am['target_col_letter']} ({am['target_header']}) via {am['match_method']}")

    total_debits = 0.0
    total_credits = 0.0
    spend_matched_count = 0

    curr_row = 6
    for idx, row in df.iterrows():
        # Col B (2): Spreadsheet Key*
        hk_raw = _pick_source_val(row, col_lookup, "headerkey")
        jk_raw = _pick_source_val(row, col_lookup, "journalkey")
        if pd.notna(hk_raw):
            key_val = _safe_int_or_val(hk_raw, 1)
        elif pd.notna(jk_raw):
            key_val = clean_val(jk_raw)
        else:
            key_val = 1
        ws.cell(row=curr_row, column=2, value=key_val)

        # Col C (3): Add Only
        ws.cell(row=curr_row, column=3, value=to_yn(_pick_source_val(row, col_lookup, "addonly")) or "Y")

        # Col D (4): Accounting Journal
        j_id = clean_val(jk_raw)
        ws.cell(row=curr_row, column=4, value=j_id)

        # Col E (5): Auto Complete
        ws.cell(row=curr_row, column=5, value="Y")

        # Col H (8): Accounting Journal ID
        ws.cell(row=curr_row, column=8, value=j_id)

        # Col I (9): Submit
        ws.cell(row=curr_row, column=9, value=to_yn(_pick_source_val(row, col_lookup, "submit")) or "Y")

        # Col M (13): Journal Number
        ws.cell(row=curr_row, column=13, value=clean_val(_pick_source_val(row, col_lookup, "journalnumber")))

        # Col O (15): Company*
        ws.cell(row=curr_row, column=15, value=clean_val(_pick_source_val(row, col_lookup, "company")) or "LE01")

        # Col P (16): Currency*
        ws.cell(row=curr_row, column=16, value=clean_val(_pick_source_val(row, col_lookup, "currency")) or "USD")

        # Col Q (17): Ledger Type*
        lt = clean_val(_pick_source_val(row, col_lookup, "ledgertype")) or "Actuals"
        ledger_type = "Actuals" if lt.upper() == "ACTUALS" else lt
        ws.cell(row=curr_row, column=17, value=ledger_type)

        # Col R (18): Book Code
        ws.cell(row=curr_row, column=18, value=clean_val(_pick_source_val(row, col_lookup, "bookcode")))

        # Col S (19): Accounting Date*
        acct_date = clean_val(_pick_source_val(row, col_lookup, "accountingdate"))
        if acct_date:
            if isinstance(acct_date, (pd.Timestamp, datetime.datetime, datetime.date)):
                acct_date_str = acct_date.strftime("%Y-%m-%d")
            else:
                acct_date_str = str(acct_date).split(" ")[0].strip()
            cell_ad = ws.cell(row=curr_row, column=19, value=acct_date_str)
            cell_ad.number_format = "@"

        # Col T (20): Journal Source*
        ws.cell(row=curr_row, column=20, value=clean_val(_pick_source_val(row, col_lookup, "journalsource")) or "Manual_Journal")

        # Col X (24): Record Quantity
        ws.cell(row=curr_row, column=24, value=clean_val(row.get("Record_Quantity")) or "N")

        # Col Y (25): Journal Entry Memo
        ws.cell(row=curr_row, column=25, value=clean_val(_pick_source_val(row, col_lookup, "journalentrymemo")))

        # Col Z (26): External Reference ID
        ws.cell(row=curr_row, column=26, value=clean_val(_pick_source_val(row, col_lookup, "journalextref")))

        # Col AI (35): Row ID* (set to 1 for single-sheet Workday format)
        ws.cell(row=curr_row, column=35, value=1)

        # Col AJ (36): Line Order
        raw_lo = _pick_source_val(row, col_lookup, "lineorder")
        raw_rid = _pick_source_val(row, col_lookup, "row_id")
        if pd.notna(raw_lo):
            l_order = _safe_int_or_val(raw_lo, idx + 1)
        elif pd.notna(raw_rid):
            l_order = _safe_int_or_val(raw_rid, idx + 1)
        else:
            l_order = idx + 1
        ws.cell(row=curr_row, column=36, value=l_order)

        # Col AK (37): Line Company
        ws.cell(row=curr_row, column=37, value=clean_val(_pick_source_val(row, col_lookup, "linecompany")) or clean_val(_pick_source_val(row, col_lookup, "company")) or "LE01")

        # Col AL (38): Ledger Account*
        acct_clean = clean_ledger_account(_pick_source_val(row, col_lookup, "ledgeraccount"))
        cell_acct = ws.cell(row=curr_row, column=38, value=acct_clean)
        if isinstance(acct_clean, int):
            cell_acct.number_format = "#,##0"
        elif acct_clean is not None:
            cell_acct.number_format = "@"

        # Col AM (39): Account Set
        ws.cell(row=curr_row, column=39, value=clean_val(_pick_source_val(row, col_lookup, "accountset")) or "ACCOUNT_SET")

        # Col AP (42): Debit Amount
        deb_val = clean_amount(_pick_source_val(row, col_lookup, "debitamount"))
        cell_deb = ws.cell(row=curr_row, column=42, value=deb_val)
        if deb_val is not None:
            cell_deb.number_format = "#,##0.00"
            total_debits += deb_val

        # Col AQ (43): Credit Amount
        crd_val = clean_amount(_pick_source_val(row, col_lookup, "creditamount"))
        cell_crd = ws.cell(row=curr_row, column=43, value=crd_val)
        if crd_val is not None:
            cell_crd.number_format = "#,##0.00"
            total_credits += crd_val

        # Col AR (44): Currency
        ws.cell(row=curr_row, column=44, value=clean_val(_pick_source_val(row, col_lookup, "linecurrency")) or "USD")

        # Col AZ (52): Memo - mapped directly from LineMemo. If empty or "NULL", strictly left blank with no value (zero fallback to JournalEntryMemo)
        memo = clean_memo(_pick_source_val(row, col_lookup, "linememo"))
        ws.cell(row=curr_row, column=52, value=memo)

        # Col BA (53): External Reference ID
        ws.cell(row=curr_row, column=53, value=clean_val(_pick_source_val(row, col_lookup, "lineextref")))

        # Col BB (54): Budget Date
        ws.cell(row=curr_row, column=54, value=parse_date(_pick_source_val(row, col_lookup, "budgetdate")))

        # Worktag Columns:
        # Col BC (55): Cost Center formatted as 3-digit number (e.g. '000', '090', '200')
        cc_3digit = clean_cost_center(_pick_source_val(row, col_lookup, "costcenter"))
        cell_cc = ws.cell(row=curr_row, column=55, value=cc_3digit)
        if cc_3digit is not None:
            cell_cc.number_format = "@"

        # Col BE (57): Spend Category mapped from Spend Category Mapped Ledger Account.xlsx
        acct_key = str(acct_clean).strip() if acct_clean is not None else ""
        mapped_sc = spend_map.get(acct_key) or clean_val(_pick_source_val(row, col_lookup, "spendcategory"))
        ws.cell(row=curr_row, column=57, value=mapped_sc)
        if mapped_sc:
            spend_matched_count += 1

        # Col BG (59): Business Unit
        ws.cell(row=curr_row, column=59, value=clean_val(_pick_source_val(row, col_lookup, "balancingworktagaffiliate")))

        # Col BI (61): Custom Organization
        ws.cell(row=curr_row, column=61, value=clean_val(_pick_source_val(row, col_lookup, "customorganization")))

        # Col BK (63): Custom Worktag 1
        ws.cell(row=curr_row, column=63, value=clean_val(_pick_source_val(row, col_lookup, "customworktag1")))

        # Col BM (65): Revenue Category
        ws.cell(row=curr_row, column=65, value=clean_val(_pick_source_val(row, col_lookup, "revenuecategory")))

        # Auto-discovered extra source/target columns
        for am in auto_mappings:
            val = _coerce_auto_val(row.get(am["source_col"]), am["target_format"])
            if val is not None:
                ws.cell(row=curr_row, column=am["target_col_idx"], value=val)

        curr_row += 1

    variance = round(total_debits - total_credits, 2)
    total_lines = len(df)
    version_id = _detect_template_version_tag(excel_template)

    print(f"\n[SINGLE-SHEET MAPPING SUMMARY ({version_id})] Successfully mapped '{source_file}':")
    print(f"  - Sheet ('{sheet_name}'): {total_lines} rows written (Rows 6 to {curr_row - 1})")
    print(f"  - Spend Categories Mapped: {spend_matched_count} / {total_lines} lines from '{SPEND_MAPPING_FILE}'")
    print(f"  - Auto-Discovered Extra Mappings: {len(auto_mappings)}")
    print(f"  - Total Debits:  ${total_debits:,.2f}")
    print(f"  - Total Credits: ${total_credits:,.2f}")
    print(f"  - Net Variance:  ${variance:,.2f} ({'BALANCED' if variance == 0 else 'UNBALANCED'})")

    os.makedirs(os.path.dirname(output_file) if os.path.dirname(output_file) else ".", exist_ok=True)
    print(f"Saving workbook to output: {output_file}...")
    wb.save(output_file)

    if not custom_output:
        root_copy = os.path.basename(output_file)
        if os.path.abspath(output_file) != os.path.abspath(root_copy):
            try:
                wb.save(root_copy)
                print(f"Also saved root copy: {root_copy}")
            except Exception:
                pass

    wb.close()
    print(f"SUCCESS: Single-sheet ({version_id}) file mapped and saved.\n")

    return {
        "success": True,
        "template_version": version_id,
        "source_file": source_file,
        "target_file": output_file,
        "total_journals": 1,
        "total_lines_written": total_lines,
        "total_source_rows": total_lines,
        "spend_categories_mapped": spend_matched_count,
        "auto_mapped_columns": auto_mappings,
        "total_debit": round(total_debits, 2),
        "total_credit": round(total_credits, 2),
        "variance": variance,
        "is_balanced": variance == 0,
        "output_file": output_file
    }

# ==============================================================================
# MAIN DISPATCHER & BULK PROCESSOR (MAX 12 FILES)
# ==============================================================================
def resolve_template_path(template_name):
    """Resolves template ID (v45, v46, v47, v48) or filename to full path."""
    if not template_name:
        return find_target_file("Import_Accounting_Journal_v45.1.template.xlsx")
    req = str(template_name).strip()
    for tpl in TARGET_TEMPLATES:
        if req.lower() in (tpl["id"].lower(), tpl["version"].lower(), tpl["filename"].lower()):
            return find_target_file(tpl["filename"])
    return find_target_file(req)

def process_journal(source_file=None, excel_template=None, output_file=None, output_dir=None):
    if excel_template is None:
        excel_template = find_target_file("Import_Accounting_Journal_v45.1.template.xlsx")
    else:
        excel_template = resolve_template_path(excel_template)

    if source_file is None or str(source_file).lower() in ("all", "both"):
        print("\n=== Processing BOTH 2025 and 2026 Accounting Journal Files ===")
        src_2026 = find_source_file("Accounting_Journal_Raw_LE01_202601.xlsx")
        if not os.path.exists(src_2026):
            src_2026 = find_source_file("Accounting_Journal_Raw_02_202601.xlsx")

        src_2025 = find_source_file("Accounting_Journal_Raw_LE01_202501.xlsx")
        if not os.path.exists(src_2025):
            src_2025 = find_source_file("Accounting_Journal_RawAgg_01_202501.xlsx")

        res_2026 = process_journal(source_file=src_2026, excel_template=excel_template, output_dir=output_dir)
        res_2025 = process_journal(source_file=src_2025, excel_template=excel_template, output_dir=output_dir)
        return {"2026": res_2026, "2025": res_2025}

    resolved_source = find_source_file(source_file)
    if not os.path.exists(resolved_source):
        resolved_source = source_file

    if not os.path.exists(excel_template):
        raise FileNotFoundError(f"Cannot find target template: {excel_template}")

    # Inspect template to decide two-sheet vs single-sheet
    wb = openpyxl.load_workbook(excel_template, read_only=True)
    sheets = wb.sheetnames
    wb.close()

    is_two_sheet = (SHEET_HEADER in sheets and SHEET_LINES in sheets)

    if is_two_sheet:
        if "2025" in str(resolved_source):
            return process_journal_2025(resolved_source, excel_template, output_file, output_dir=output_dir)
        else:
            return process_journal_2026(resolved_source, excel_template, output_file, output_dir=output_dir)
    else:
        return process_journal_single_sheet(resolved_source, excel_template, output_file, output_dir=output_dir)

def process_bulk(source_files, excel_template=None, output_dir=None, output_prefix=None):
    """Processes up to 12 source files sequentially into the selected Workday EIB template."""
    if not source_files:
        raise ValueError("No source files provided for bulk processing.")
    if len(source_files) > MAX_BULK_FILES:
        raise ValueError(f"Bulk processing supports a maximum of {MAX_BULK_FILES} files per batch (received {len(source_files)}).")

    tpl_path = resolve_template_path(excel_template) if excel_template else resolve_template_path("v45")
    tpl_ver = _detect_template_version_tag(tpl_path)
    out_dir = output_dir if (output_dir and os.path.isdir(output_dir)) else os.path.join("output", datetime.date.today().strftime("%Y-%m-%d"))
    os.makedirs(out_dir, exist_ok=True)

    results = []
    used_names = set()
    total_lines = 0
    total_debits = 0.0
    total_credits = 0.0
    succeeded = 0
    failed = 0

    for idx, src in enumerate(source_files, 1):
        bname = os.path.basename(str(src))
        stem = os.path.splitext(bname)[0] or f"File_{idx}"
        if output_prefix and str(output_prefix).strip():
            clean_pfx = os.path.splitext(os.path.basename(str(output_prefix).strip()))[0]
            candidate_name = f"{clean_pfx}_{stem}.xlsx"
        else:
            candidate_name = f"{stem}_{tpl_ver}.xlsx"

        # Prevent collisions if duplicate basenames exist in the batch
        if candidate_name.lower() in used_names:
            candidate_name = f"{os.path.splitext(candidate_name)[0]}_{idx}.xlsx"
        used_names.add(candidate_name.lower())
        out_path = os.path.join(out_dir, candidate_name)

        try:
            res = process_journal(source_file=src, excel_template=tpl_path, output_file=out_path, output_dir=out_dir)
            results.append(res)
            total_lines += res.get("total_lines_written", 0)
            total_debits += float(res.get("total_debit", 0.0))
            total_credits += float(res.get("total_credit", 0.0))
            succeeded += 1
        except Exception as e:
            failed += 1
            results.append({
                "success": False,
                "source_file": src,
                "error": str(e),
                "output_file": out_path
            })

    return {
        "success": failed == 0,
        "max_bulk_files": MAX_BULK_FILES,
        "total_files": len(source_files),
        "succeeded": succeeded,
        "failed": failed,
        "template_version": tpl_ver,
        "output_dir": out_dir,
        "total_lines_written": total_lines,
        "total_debit": round(total_debits, 2),
        "total_credit": round(total_credits, 2),
        "variance": round(total_debits - total_credits, 2),
        "results": results
    }

def main():
    parser = argparse.ArgumentParser(description="Workday Accounting Journal EIB Ingest Agent")
    parser.add_argument("--source", "-s", help="Source file (Excel or CSV), or 'all' to process both", default=None)
    parser.add_argument("--target", "-t", help="Target Excel base template", default=None)
    parser.add_argument("--all", "-a", action="store_true", help="Process both 2025 and 2026 files")
    parser.add_argument("--bulk", "-b", nargs="+", help=f"Bulk process up to {MAX_BULK_FILES} source files", default=None)
    parser.add_argument("--output-dir", "-o", help="Destination folder for output files", default=None)
    parser.add_argument("--analyze", action="store_true", help="Analyze source and target schemas without generating files")
    args = parser.parse_args()

    if args.analyze and args.source:
        import json
        report = analyze_schema(args.source, args.target)
        print(json.dumps(report, indent=2, default=str))
        return

    if args.bulk:
        summary = process_bulk(args.bulk, excel_template=args.target, output_dir=args.output_dir)
        print(f"\n[BULK SUMMARY] {summary['succeeded']}/{summary['total_files']} files processed -> {summary['output_dir']}")
        return

    if args.all:
        process_journal(source_file="all", excel_template=args.target, output_dir=args.output_dir)
        return

    # Discover available files in source and target directories as well as root
    existing_sources = [f for f in DEFAULT_SOURCE_FILES if os.path.exists(f)]
    for folder in [".", "source"]:
        if os.path.exists(folder):
            for fname in os.listdir(folder):
                full_p = os.path.join(folder, fname) if folder != "." else fname
                if (fname.endswith(".xlsx") or fname.endswith(".csv")) and "raw" in fname.lower() and full_p not in existing_sources:
                    existing_sources.append(full_p)
    existing_sources.append("Process BOTH 2025 and 2026 files")

    existing_targets = [f for f in DEFAULT_TARGET_FILES if os.path.exists(f)]
    for folder in [".", "target"]:
        if os.path.exists(folder):
            for fname in os.listdir(folder):
                full_p = os.path.join(folder, fname) if folder != "." else fname
                if fname.endswith(".xlsx") and ("import" in fname.lower() or "submit" in fname.lower()) and full_p not in existing_targets:
                    existing_targets.append(full_p)

    # 1. Ask user which file to select (show source file names)
    if args.source:
        chosen_source = args.source
    else:
        chosen_source = select_file_interactive(
            "Select source accounting file to process:",
            existing_sources,
            default_index=0
        )

    if "both" in chosen_source.lower() or "all" in chosen_source.lower():
        process_journal(source_file="all", excel_template=args.target, output_dir=args.output_dir)
        return

    # 2. Ask user where to map the data (show base template name)
    if args.target:
        chosen_target = args.target
    else:
        chosen_target = select_file_interactive(
            "Select base Workday EIB template to map data into:",
            existing_targets,
            default_index=0
        )

    # 3. Execute mapping individually via dispatcher
    process_journal(source_file=chosen_source, excel_template=chosen_target, output_dir=args.output_dir)

if __name__ == "__main__":
    main()