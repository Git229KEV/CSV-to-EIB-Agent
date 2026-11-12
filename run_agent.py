import os
import sys
import csv
import math
import argparse
import datetime
import shutil
import openpyxl

class _IlocIndexer:
    def __init__(self, rows):
        self._rows = rows
    def __getitem__(self, idx):
        return self._rows[idx]

class _LightweightDataFrame:
    def __init__(self, rows):
        self._rows = rows
        self.iloc = _IlocIndexer(rows)
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
    def read_excel(path):
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        ws = wb.active
        rows_iter = ws.iter_rows(values_only=True)
        header_row = next(rows_iter, None)
        if not header_row:
            wb.close()
            return _LightweightDataFrame([])
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
        wb.close()
        return _LightweightDataFrame(records)

    @staticmethod
    def read_csv(path):
        records = []
        with open(path, mode="r", encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            for r in reader:
                if not r or all(v is None or str(v).strip() == "" for v in r.values()):
                    continue
                row_dict = {}
                for k, v in r.items():
                    if k is not None:
                        row_dict[k.strip()] = _parse_csv_val(v)
                records.append(row_dict)
        return _LightweightDataFrame(records)

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

def get_output_filename(source_file):
    """Returns official Workday EIB filename matching strict standard:
    - 2025: Import_Accounting_Journal_v45.0_LE01_2025_P1.xlsx
    - 2026: Import_Accounting_Journal_v45.0_LE01_2026_P2.xlsx
    """
    if "2025" in str(source_file):
        return "Import_Accounting_Journal_v45.0_LE01_2025_P1.xlsx"
    else:
        return "Import_Accounting_Journal_v45.0_LE01_2026_P2.xlsx"

def get_datewise_output_path(source_file, target_dir=None, template_file=None):
    """Generates the output path. If target_dir is given by user, writes directly there.
    Otherwise defaults to date-wise output/YYYY-MM-DD/ and synchronized root copy."""
    filename = get_output_filename(source_file)
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
    if val_str in ("1", "1.0", "Y", "True", "true"):
        return "Y"
    elif val_str in ("0", "0.0", "N", "False", "false"):
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
    val_str = str(val).strip().split()[0]
    try:
        parts = [int(p) for p in val_str.split("-")]
        if len(parts) == 3:
            return datetime.date(parts[0], parts[1], parts[2])
    except Exception:
        pass
    return val_str

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
def process_journal_2026(source_file, excel_template, output_file=None, output_dir=None):
    # Load Spend Category lookup map
    spend_map = load_spend_category_map()

    # Determine date-wise or user-selected output path
    custom_output = output_file is not None
    datewise_target = get_datewise_output_path(source_file, target_dir=output_dir)
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

        # Identify unique journals while preserving appearance order
        unique_headers = []
        seen_keys = []
        for _, row in df.iterrows():
            key = row.get("HeaderKey") if "HeaderKey" in row and pd.notna(row.get("HeaderKey")) else (row.get("JournalKey") or row.get("AccountingJournalID"))
            if key not in seen_keys:
                seen_keys.append(key)
                unique_headers.append(row)

        hdr_row = 6
        header_key_map = {}
        for h_row in unique_headers:
            h_key = h_row.get("HeaderKey") if "HeaderKey" in h_row and pd.notna(h_row.get("HeaderKey")) else (h_row.get("JournalKey") or h_row.get("AccountingJournalID"))
            if "HeaderKey" in h_row and pd.notna(h_row.get("HeaderKey")):
                hk_val = int(h_row.get("HeaderKey"))
            else:
                hk_val = hdr_row - 5
            header_key_map[h_key] = hk_val

            # Col B (2): Header Key*
            ws_hdr.cell(row=hdr_row, column=2, value=hk_val)

            # Col C (3): Add Only
            ws_hdr.cell(row=hdr_row, column=3, value=to_yn(h_row.get("AddOnlyJournal")) or "Y")

            # Col F (6): Accounting Journal ID
            j_id = clean_val(h_row.get("AccountingJournalID")) or clean_val(h_row.get("JournalKey"))
            ws_hdr.cell(row=hdr_row, column=6, value=j_id)

            # Col E (5): Accounting Journal -> Strictly left blank
            ws_hdr.cell(row=hdr_row, column=5, value=None)

            # Col L (12): Company*
            ws_hdr.cell(row=hdr_row, column=12, value=clean_val(h_row.get("CompanyReferenceID")) or "LE01")

            # Col M (13): Currency*
            ws_hdr.cell(row=hdr_row, column=13, value=clean_val(h_row.get("Currency")) or "USD")

            # Col N (14): Ledger Type*
            lt = clean_val(h_row.get("LedgerType")) or "Actuals"
            ws_hdr.cell(row=hdr_row, column=14, value="Actuals" if lt.upper() == "ACTUALS" else lt)

            # Col P (16): Accounting Date*
            ws_hdr.cell(row=hdr_row, column=16, value=parse_date(h_row.get("AccountingDate")))

            # Col Q (17): Journal Source*
            ws_hdr.cell(row=hdr_row, column=17, value=clean_val(h_row.get("JournalSource")) or "Manual_Journal")

            # Col U (21): Record Quantity -> Strictly left blank
            ws_hdr.cell(row=hdr_row, column=21, value=None)

            # Col V (22): Journal Entry Memo
            ws_hdr.cell(row=hdr_row, column=22, value=clean_val(h_row.get("JournalEntryMemo")))

            # Col W (23): External Reference ID
            ws_hdr.cell(row=hdr_row, column=23, value=clean_val(h_row.get("JournalExternalReferenceID")))

            # Col AG (33): Auto Complete
            ws_hdr.cell(row=hdr_row, column=33, value="Y")

            hdr_row += 1

        line_row = 6
        total_debits = 0.0
        total_credits = 0.0
        spend_matched_count = 0

        # Resolve Sheet 2 column positions dynamically from Row 5 headers (supports updated v45.1 schema where
        # Intercompany Affiliate is Col 29/AC and Balancing Worktag Affiliate shifted to Col 31/AE)
        line_hdr_cols = {
            str(ws_line.cell(row=5, column=c).value).strip(): c
            for c in range(1, ws_line.max_column + 1)
            if ws_line.cell(row=5, column=c).value is not None
        }
        col_intercompany = line_hdr_cols.get("Intercompany Affiliate")
        col_balancing_affiliate = line_hdr_cols.get("Balancing Worktag Affiliate", 31)

        for idx, row in df.iterrows():
            r_key = row.get("HeaderKey") if "HeaderKey" in row and pd.notna(row.get("HeaderKey")) else (row.get("JournalKey") or row.get("AccountingJournalID"))
            if "HeaderKey" in row and pd.notna(row.get("HeaderKey")):
                hk_val = int(row.get("HeaderKey"))
            else:
                hk_val = header_key_map.get(r_key, 1)

            ws_line.cell(row=line_row, column=2, value=hk_val)

            if "ROW_ID" in row and pd.notna(row.get("ROW_ID")):
                line_key_val = int(row.get("ROW_ID"))
            else:
                line_key_val = idx + 1
            ws_line.cell(row=line_row, column=3, value=line_key_val)

            if "LineOrder" in row and pd.notna(row.get("LineOrder")):
                line_order_val = int(row.get("LineOrder"))
            else:
                line_order_val = line_key_val
            ws_line.cell(row=line_row, column=4, value=line_order_val)

            ws_line.cell(row=line_row, column=5, value=clean_val(row.get("LineCompanyReferenceID")) or "LE01")
            
            acct_clean = clean_ledger_account(row.get("LedgerAccountReferenceID"))
            ws_line.cell(row=line_row, column=6, value=acct_clean)
            ws_line.cell(row=line_row, column=7, value=clean_val(row.get("LedgerAccountReferenceID_ParentID")) or "ACCOUNT_SET")

            deb_val = clean_amount(row.get("DebitAmount"))
            crd_val = clean_amount(row.get("CreditAmount"))

            if deb_val is not None:
                ws_line.cell(row=line_row, column=10, value=deb_val)
                total_debits += deb_val

            if crd_val is not None:
                ws_line.cell(row=line_row, column=11, value=crd_val)
                total_credits += crd_val

            ws_line.cell(row=line_row, column=12, value=clean_val(row.get("LineCurrency")) or "USD")

            # Col T (20): Memo - mapped directly from LineMemo. If empty or "NULL", left blank with no value (no fallback)
            memo = clean_memo(row.get("LineMemo"))
            ws_line.cell(row=line_row, column=20, value=memo)

            ws_line.cell(row=line_row, column=21, value=clean_val(row.get("JournalLineExternalReferenceID")))
            ws_line.cell(row=line_row, column=22, value=parse_date(row.get("BudgetDate")))
            ws_line.cell(row=line_row, column=23, value=clean_val(row.get("Worktag_Revenue_Category_ID")))
            
            # Map Spend Category using Spend Category Mapped Ledger Account.xlsx
            acct_key = str(acct_clean).strip() if acct_clean is not None else ""
            mapped_sc = spend_map.get(acct_key) or clean_val(row.get("Worktag_Spend_Category_ID"))
            ws_line.cell(row=line_row, column=25, value=mapped_sc)
            if mapped_sc:
                spend_matched_count += 1

            cc_3digit = clean_cost_center(row.get("Worktag_Cost_Center_Reference_ID"))
            cell_cc = ws_line.cell(row=line_row, column=27, value=cc_3digit)
            if cc_3digit is not None:
                cell_cc.number_format = "@"

            if col_intercompany:
                interco_val = (
                    clean_val(row.get("Worktag_Intercompany_Affiliate_ID"))
                    or clean_val(row.get("IntercompanyAffiliate"))
                    or clean_val(row.get("Intercompany_Affiliate_ID"))
                )
                ws_line.cell(row=line_row, column=col_intercompany, value=interco_val)

            ws_line.cell(row=line_row, column=col_balancing_affiliate, value=clean_val(row.get("Worktag_Business_Unit_ID")) or clean_val(row.get("OptionalBalancingWorktags_Business_Unit_ID")))

            line_row += 1

        total_journals = len(unique_headers)
        total_lines_written = len(df)
        variance = round(total_debits - total_credits, 2)

        print(f"\n[2026 MAPPING SUMMARY] Successfully mapped '{source_file}':")
        print(f"  - Sheet 1 ('{SHEET_HEADER}'): {total_journals} journal header rows written (Rows 6 to {hdr_row - 1})")
        print(f"  - Sheet 2 ('{SHEET_LINES}'): {total_lines_written} journal line rows written (Rows 6 to {line_row - 1})")
        print(f"  - Spend Categories Mapped: {spend_matched_count} / {total_lines_written} lines from '{SPEND_MAPPING_FILE}'")
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
    datewise_target = get_datewise_output_path(source_file, target_dir=output_dir)
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
    print(f"Loading single-sheet Excel template: {excel_template}...")
    wb = openpyxl.load_workbook(excel_template)
    
    sheet_name = "Submit Accounting Journal"
    if sheet_name not in wb.sheetnames:
        sheet_name = wb.sheetnames[0]
    ws = wb[sheet_name]

    if ws.max_row >= 6:
        print(f"Clearing Sheet '{sheet_name}' data rows 6 to {ws.max_row}...")
        ws.delete_rows(6, ws.max_row - 5)

    total_debits = 0.0
    total_credits = 0.0
    spend_matched_count = 0

    curr_row = 6
    for idx, row in df.iterrows():
        # Col B (2): Spreadsheet Key*
        if "HeaderKey" in row and pd.notna(row.get("HeaderKey")):
            key_val = int(row.get("HeaderKey"))
        elif "JournalKey" in row and pd.notna(row.get("JournalKey")):
            key_val = clean_val(row.get("JournalKey"))
        else:
            key_val = 1
        ws.cell(row=curr_row, column=2, value=key_val)

        # Col C (3): Add Only
        ws.cell(row=curr_row, column=3, value=to_yn(row.get("AddOnlyJournal")) or "Y")

        # Col D (4): Accounting Journal
        j_id = clean_val(row.get("AccountingJournalID")) or clean_val(row.get("JournalKey"))
        ws.cell(row=curr_row, column=4, value=j_id)

        # Col E (5): Auto Complete
        ws.cell(row=curr_row, column=5, value="Y")

        # Col H (8): Accounting Journal ID
        ws.cell(row=curr_row, column=8, value=j_id)

        # Col I (9): Submit
        ws.cell(row=curr_row, column=9, value=to_yn(row.get("Submit")) or "Y")

        # Col M (13): Journal Number
        ws.cell(row=curr_row, column=13, value=clean_val(row.get("JournalNumber")))

        # Col O (15): Company*
        ws.cell(row=curr_row, column=15, value=clean_val(row.get("CompanyReferenceID")) or "LE01")

        # Col P (16): Currency*
        ws.cell(row=curr_row, column=16, value=clean_val(row.get("Currency")) or "USD")

        # Col Q (17): Ledger Type*
        lt = clean_val(row.get("LedgerType")) or "Actuals"
        ledger_type = "Actuals" if lt.upper() == "ACTUALS" else lt
        ws.cell(row=curr_row, column=17, value=ledger_type)

        # Col R (18): Book Code
        ws.cell(row=curr_row, column=18, value=clean_val(row.get("BookCode")))

        # Col S (19): Accounting Date*
        acct_date = clean_val(row.get("AccountingDate"))
        if acct_date:
            if isinstance(acct_date, (pd.Timestamp, datetime.datetime, datetime.date)):
                acct_date_str = acct_date.strftime("%Y-%m-%d")
            else:
                acct_date_str = str(acct_date).split(" ")[0].strip()
            cell_ad = ws.cell(row=curr_row, column=19, value=acct_date_str)
            cell_ad.number_format = "@"

        # Col T (20): Journal Source*
        ws.cell(row=curr_row, column=20, value=clean_val(row.get("JournalSource")) or "Manual_Journal")

        # Col X (24): Record Quantity
        ws.cell(row=curr_row, column=24, value=clean_val(row.get("Record_Quantity")) or "N")

        # Col Y (25): Journal Entry Memo
        ws.cell(row=curr_row, column=25, value=clean_val(row.get("JournalEntryMemo")))

        # Col Z (26): External Reference ID
        ws.cell(row=curr_row, column=26, value=clean_val(row.get("JournalExternalReferenceID")))

        # Col AI (35): Row ID* (set to 1 for single-sheet Workday format)
        ws.cell(row=curr_row, column=35, value=1)

        # Col AJ (36): Line Order
        if "LineOrder" in row and pd.notna(row.get("LineOrder")):
            l_order = int(row.get("LineOrder"))
        elif "ROW_ID" in row and pd.notna(row.get("ROW_ID")):
            l_order = int(row.get("ROW_ID"))
        else:
            l_order = idx + 1
        ws.cell(row=curr_row, column=36, value=l_order)

        # Col AK (37): Line Company
        ws.cell(row=curr_row, column=37, value=clean_val(row.get("LineCompanyReferenceID")) or clean_val(row.get("CompanyReferenceID")) or "LE01")

        # Col AL (38): Ledger Account*
        acct_clean = clean_ledger_account(row.get("LedgerAccountReferenceID"))
        cell_acct = ws.cell(row=curr_row, column=38, value=acct_clean)
        if isinstance(acct_clean, int):
            cell_acct.number_format = "#,##0"
        elif acct_clean is not None:
            cell_acct.number_format = "@"

        # Col AM (39): Account Set
        ws.cell(row=curr_row, column=39, value=clean_val(row.get("LedgerAccountReferenceID_ParentID")) or "ACCOUNT_SET")

        # Col AP (42): Debit Amount
        deb_val = clean_amount(row.get("DebitAmount"))
        cell_deb = ws.cell(row=curr_row, column=42, value=deb_val)
        if deb_val is not None:
            cell_deb.number_format = "#,##0.00"
            total_debits += deb_val

        # Col AQ (43): Credit Amount
        crd_val = clean_amount(row.get("CreditAmount"))
        cell_crd = ws.cell(row=curr_row, column=43, value=crd_val)
        if crd_val is not None:
            cell_crd.number_format = "#,##0.00"
            total_credits += crd_val

        # Col AR (44): Currency
        ws.cell(row=curr_row, column=44, value=clean_val(row.get("LineCurrency")) or "USD")

        # Col AZ (52): Memo - mapped directly from LineMemo. If empty or "NULL", strictly left blank with no value (zero fallback to JournalEntryMemo)
        memo = clean_memo(row.get("LineMemo"))
        ws.cell(row=curr_row, column=52, value=memo)

        # Col BA (53): External Reference ID
        ws.cell(row=curr_row, column=53, value=clean_val(row.get("JournalLineExternalReferenceID")))

        # Col BB (54): Budget Date
        ws.cell(row=curr_row, column=54, value=clean_val(row.get("BudgetDate")))

        # Worktag Columns:
        # Col BC (55): Cost Center formatted as 3-digit number (e.g. '000', '090', '200')
        cc_3digit = clean_cost_center(row.get("Worktag_Cost_Center_Reference_ID"))
        cell_cc = ws.cell(row=curr_row, column=55, value=cc_3digit)
        if cc_3digit is not None:
            cell_cc.number_format = "@"

        # Col BE (57): Spend Category mapped from Spend Category Mapped Ledger Account.xlsx
        acct_key = str(acct_clean).strip() if acct_clean is not None else ""
        mapped_sc = spend_map.get(acct_key) or clean_val(row.get("Worktag_Spend_Category_ID"))
        ws.cell(row=curr_row, column=57, value=mapped_sc)
        if mapped_sc:
            spend_matched_count += 1

        # Col BG (59): Business Unit
        ws.cell(row=curr_row, column=59, value=clean_val(row.get("Worktag_Business_Unit_ID")) or clean_val(row.get("OptionalBalancingWorktags_Business_Unit_ID")))

        # Col BI (61): Custom Organization
        ws.cell(row=curr_row, column=61, value=clean_val(row.get("Worktag_Custom_Organization_Reference_ID")))

        # Col BK (63): Custom Worktag 1
        ws.cell(row=curr_row, column=63, value=clean_val(row.get("Worktag_Custom_Worktag_1_ID")))

        # Col BM (65): Revenue Category
        ws.cell(row=curr_row, column=65, value=clean_val(row.get("Worktag_Revenue_Category_ID")))

        curr_row += 1

    variance = round(total_debits - total_credits, 2)
    total_lines = len(df)
    version_id = "v48.0" if "v48" in excel_template.lower() else ("v47.0" if "v47" in excel_template.lower() else "v46.0")

    print(f"\n[SINGLE-SHEET MAPPING SUMMARY ({version_id})] Successfully mapped '{source_file}':")
    print(f"  - Sheet ('{sheet_name}'): {total_lines} rows written (Rows 6 to {curr_row - 1})")
    print(f"  - Spend Categories Mapped: {spend_matched_count} / {total_lines} lines from '{SPEND_MAPPING_FILE}'")
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
        "total_debit": round(total_debits, 2),
        "total_credit": round(total_credits, 2),
        "variance": variance,
        "is_balanced": variance == 0,
        "output_file": output_file
    }

# ==============================================================================
# MAIN DISPATCHER
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

def main():
    parser = argparse.ArgumentParser(description="Workday Accounting Journal EIB Ingest Agent")
    parser.add_argument("--source", "-s", help="Source file (Excel or CSV), or 'all' to process both", default=None)
    parser.add_argument("--target", "-t", help="Target Excel base template", default=None)
    parser.add_argument("--all", "-a", action="store_true", help="Process both 2025 and 2026 files")
    args = parser.parse_args()

    if args.all:
        process_journal(source_file="all", excel_template=args.target)
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
        process_journal(source_file="all", excel_template=args.target)
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
    process_journal(source_file=chosen_source, excel_template=chosen_target)

if __name__ == "__main__":
    main()