import os
import sys
import base64
import tempfile
import subprocess
import webview

# Support PyInstaller runtime extraction path
if hasattr(sys, '_MEIPASS'):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Ensure current working directory is the executable's folder (not _MEIPASS temp folder)
if getattr(sys, 'frozen', False):
    os.chdir(os.path.dirname(os.path.abspath(sys.executable)))
else:
    os.chdir(os.path.dirname(os.path.abspath(__file__)))

from run_agent import (
    process_journal as run_agent_process,
    process_bulk as run_agent_process_bulk,
    analyze_schema as run_agent_analyze_schema,
    find_source_file,
    find_target_file,
    resolve_target_file,
    get_output_filename,
    load_source_dataframe,
    pd,
    TARGET_TEMPLATES,
    MAX_BULK_FILES
)

class DesktopApi:
    """Native Python API bridge callable directly from frontend JavaScript in desktop environment."""
    def __init__(self):
        self._window = None

    def get_version(self):
        return "Workday Accounting Journal EIB Ingest Studio — Enterprise Desktop"

    def is_desktop(self):
        return True

    def inspect_source_file(self, filepath, target_template="v45"):
        """Automatically inspects an uploaded/selected source file, counts records, and analyzes schema vs template."""
        try:
            resolved = find_source_file(filepath)
            if not resolved or not os.path.exists(resolved):
                if os.path.exists(str(filepath)):
                    resolved = str(filepath)
                else:
                    return {"success": False, "error": f"File not found: {filepath}"}
            df = load_source_dataframe(resolved)
            seen_keys = []
            for _, row in df.iterrows():
                key = row.get("HeaderKey") if "HeaderKey" in row and pd.notna(row.get("HeaderKey")) else (row.get("JournalKey") or row.get("AccountingJournalID"))
                if key not in seen_keys:
                    seen_keys.append(key)
            ext = os.path.splitext(resolved)[1].upper().lstrip(".") or "XLSX"

            schema = {}
            try:
                tpl_path = resolve_target_file(target_template)
                schema = run_agent_analyze_schema(resolved, tpl_path)
            except Exception:
                schema = {}

            return {
                "success": True,
                "filepath": resolved,
                "filename": os.path.basename(resolved),
                "size": os.path.getsize(resolved) if os.path.exists(resolved) else 0,
                "total_rows": len(df),
                "total_journals": max(len(seen_keys), 1) if len(df) > 0 else 0,
                "format": ext,
                "source_columns": schema.get("source_columns", []),
                "intercompany_source_col": schema.get("intercompany_source_col"),
                "auto_discovered_mappings": schema.get("auto_discovered_mappings", []),
                "unmapped_populated_columns": schema.get("unmapped_populated_columns", []),
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def analyze_schema(self, filepath, target_template="v45"):
        """Analyzes source columns vs selected target template columns and returns auto-discovered mappings."""
        try:
            tpl_path = resolve_target_file(target_template)
            report = run_agent_analyze_schema(filepath, tpl_path)
            return {"success": True, **report}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def get_default_paths(self):
        """Returns default workspace and output paths without pre-filling source file."""
        default_out = os.path.abspath("output")
        if not os.path.exists(default_out):
            try:
                os.makedirs(default_out, exist_ok=True)
            except Exception:
                pass

        return {
            "workspace_dir": os.path.abspath("."),
            "default_output_dir": default_out,
            "max_bulk_files": MAX_BULK_FILES
        }

    def select_file(self, allow_multiple=True, target_template="v45"):
        """Opens native OS file chooser dialog (supports single or bulk up to 12 files) and analyzes schema."""
        if not self._window:
            return {"success": False, "filepath": None, "files": []}
        try:
            try:
                result = self._window.create_file_dialog(
                    webview.OPEN_DIALOG,
                    allow_multiple=bool(allow_multiple),
                    file_types=('Excel and CSV Files (*.xlsx;*.xls;*.csv)', 'All files (*.*)')
                )
            except Exception:
                result = self._window.create_file_dialog(
                    webview.OPEN_DIALOG,
                    allow_multiple=bool(allow_multiple)
                )
            if result and len(result) > 0:
                selected_list = list(result)
                exceeded = len(selected_list) > MAX_BULK_FILES
                selected_count = len(selected_list)
                capped = selected_list[:MAX_BULK_FILES]

                inspected_files = []
                for p in capped:
                    info = self.inspect_source_file(p, target_template=target_template)
                    if info.get("success"):
                        inspected_files.append(info)
                    else:
                        inspected_files.append({
                            "success": False,
                            "filepath": p,
                            "filename": os.path.basename(p),
                            "size": os.path.getsize(p) if os.path.exists(p) else 0,
                            "error": info.get("error")
                        })

                first = inspected_files[0]
                return {
                    "success": True,
                    "filepath": first["filepath"],
                    "path": first["filepath"],
                    "filename": first["filename"],
                    "name": first["filename"],
                    "size": first.get("size", 0),
                    "total_rows": first.get("total_rows"),
                    "total_journals": first.get("total_journals"),
                    "format": first.get("format"),
                    "auto_discovered_mappings": first.get("auto_discovered_mappings", []),
                    "intercompany_source_col": first.get("intercompany_source_col"),
                    "unmapped_populated_columns": first.get("unmapped_populated_columns", []),
                    "files": inspected_files,
                    "exceeded_limit": exceeded,
                    "selected_count": selected_count,
                    "max_bulk_files": MAX_BULK_FILES,
                }
            return {"success": False, "cancelled": True, "filepath": None, "files": []}
        except Exception as e:
            print(f"[DesktopApp] File dialog error: {e}", flush=True)
        return {"success": False, "filepath": None, "files": []}

    def stage_uploaded_file(self, filename, base64_data, target_template="v45"):
        """Stages a file selected via HTML <input type='file'> into a local temp path and auto-detects structure."""
        try:
            if "," in base64_data:
                base64_data = base64_data.split(",", 1)[1]
            raw_bytes = base64.b64decode(base64_data)
            stage_dir = os.path.join(tempfile.gettempdir(), "WorkdayEIBStudio_Staged")
            os.makedirs(stage_dir, exist_ok=True)
            safe_name = os.path.basename(filename)
            staged_path = os.path.join(stage_dir, safe_name)
            with open(staged_path, "wb") as f:
                f.write(raw_bytes)
            info = self.inspect_source_file(staged_path, target_template=target_template)
            return {
                "success": True,
                "filepath": staged_path,
                "filename": safe_name,
                "size": len(raw_bytes),
                "total_rows": info.get("total_rows") if info.get("success") else None,
                "total_journals": info.get("total_journals") if info.get("success") else None,
                "format": info.get("format") if info.get("success") else None,
                "auto_discovered_mappings": info.get("auto_discovered_mappings", []),
                "intercompany_source_col": info.get("intercompany_source_col"),
                "unmapped_populated_columns": info.get("unmapped_populated_columns", []),
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def select_folder(self):
        """Opens native OS folder chooser dialog to select local save destination."""
        if not self._window:
            return {"success": False, "folder": None}
        try:
            result = self._window.create_file_dialog(
                webview.FOLDER_DIALOG
            )
            if result and len(result) > 0:
                selected_folder = result[0]
                return {
                    "success": True,
                    "path": selected_folder,
                    "folder": selected_folder
                }
        except Exception as e:
            print(f"[DesktopApp] Folder dialog error: {e}", flush=True)
        return {"success": False, "folder": None}

    def process_journal(self, filename, target_template="v45", output_dir=None, output_filename=None):
        """Directly executes Workday EIB ingestion and returns mapped results."""
        try:
            source_to_process = find_source_file(filename)
            if not source_to_process or not os.path.exists(source_to_process):
                if os.path.exists(filename):
                    source_to_process = filename
                else:
                    return {"success": False, "error": f"Source file '{filename}' could not be found."}

            if output_dir and not os.path.exists(output_dir):
                try:
                    os.makedirs(output_dir, exist_ok=True)
                except Exception:
                    output_dir = None

            custom_output_file = None
            if output_filename and str(output_filename).strip():
                clean_name = os.path.basename(str(output_filename).strip())
                if not clean_name.lower().endswith(".xlsx"):
                    clean_name += ".xlsx"
                base_out_dir = output_dir if output_dir else os.path.abspath("output")
                os.makedirs(base_out_dir, exist_ok=True)
                custom_output_file = os.path.join(base_out_dir, clean_name)

            target_path = resolve_target_file(target_template)
            res = run_agent_process(
                source_file=source_to_process,
                excel_template=target_path,
                output_file=custom_output_file,
                output_dir=output_dir
            )
            target_out = res.get("output_file") or res.get("target_file") or custom_output_file or target_path

            return {
                "success": True,
                "filename": filename,
                "target_template": target_template,
                "target_file": os.path.basename(target_out),
                "target_path": target_out,
                "summary": res
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def process_both(self, target_template="v45", output_dir=None):
        """Batch processes both 2026 and 2025 raw accounting files."""
        try:
            if output_dir and not os.path.exists(output_dir):
                try:
                    os.makedirs(output_dir, exist_ok=True)
                except Exception:
                    output_dir = None

            target_path = resolve_target_file(target_template)
            res = run_agent_process(source_file="all", excel_template=target_path, output_dir=output_dir)
            return {
                "success": True,
                "target_template": target_template,
                "summary": res
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def process_bulk(self, filepaths, target_template="v45", output_dir=None, output_prefix=None):
        """Bulk processes up to 12 selected source files into the chosen Workday EIB template."""
        try:
            if not filepaths or not isinstance(filepaths, (list, tuple)):
                return {"success": False, "error": "Please select at least 1 file for bulk processing."}
            if len(filepaths) > MAX_BULK_FILES:
                return {"success": False, "error": f"Bulk processing supports a maximum of {MAX_BULK_FILES} files per batch (selected {len(filepaths)})."}

            if output_dir and not os.path.exists(output_dir):
                try:
                    os.makedirs(output_dir, exist_ok=True)
                except Exception:
                    output_dir = None

            target_path = resolve_target_file(target_template)
            res = run_agent_process_bulk(
                source_files=list(filepaths),
                excel_template=target_path,
                output_dir=output_dir,
                output_prefix=output_prefix
            )
            return res
        except Exception as e:
            return {"success": False, "error": str(e)}

    def open_in_explorer(self, target_path=None):
        """Opens Windows Explorer at the folder or file location."""
        try:
            p = target_path if target_path else os.path.abspath("output")
            abs_p = os.path.abspath(p)
            if os.path.isfile(abs_p):
                subprocess.Popen(['explorer', '/select,', abs_p])
            elif os.path.isdir(abs_p):
                os.startfile(abs_p)
            return {"success": True}
        except Exception as e:
            print(f"[DesktopApp] Explorer error: {e}", flush=True)
            return {"success": False, "error": str(e)}

    def open_in_excel(self, target_path=None):
        """Opens the generated spreadsheet directly in Microsoft Excel."""
        try:
            target = target_path
            if not target or not os.path.exists(target):
                for cand in [
                    os.path.join("output", "Import_Accounting_Journal_v45.0_LE01_2026_P2.xlsx"),
                    os.path.join("output", "Import_Accounting_Journal_v45.0_LE01_2025_P1.xlsx")
                ]:
                    if os.path.exists(cand):
                        target = cand
                        break
            if target and os.path.exists(target):
                os.startfile(target)
                return {"success": True}
            return {"success": False, "error": f"File not found: {target_path}"}
        except Exception as e:
            print(f"[DesktopApp] Excel error: {e}", flush=True)
            return {"success": False, "error": str(e)}

    def get_sample_csv(self):
        """Loads sample source file info."""
        src = find_source_file("sample_journal_lines.csv")
        if not src or not os.path.exists(src):
            src = find_source_file("Accounting_Journal_Raw_LE01_202601.xlsx")
        if src and os.path.exists(src):
            return {
                "success": True,
                "filename": os.path.basename(src),
                "path": src
            }
        return {"success": False, "error": "Sample file not found"}

def start_desktop_app():
    html_path = os.path.join(BASE_DIR, "frontend", "index.html")
    if not os.path.exists(html_path):
        html_path = os.path.abspath(os.path.join("frontend", "index.html"))

    api = DesktopApi()
    window = webview.create_window(
        title="Workday Accounting Journal EIB Ingest Studio — Enterprise Desktop",
        url=html_path,
        width=1340,
        height=880,
        min_size=(1020, 680),
        text_select=True,
        confirm_close=False,
        js_api=api
    )
    api._window = window

    webview.start(debug=False)

if __name__ == "__main__":
    start_desktop_app()
