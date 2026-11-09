import os
import sys
import time
from playwright.sync_api import sync_playwright
from desktop_app import DesktopApi
import openpyxl

def run_e2e_tests():
    api = DesktopApi()
    html_path = os.path.abspath("frontend/index.html")
    file_url = f"file:///{html_path.replace(os.sep, '/')}"
    print(f"[*] Testing Desktop Studio UI directly via native file URL: {file_url} (No Server / No Port)...")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(viewport={"width": 1440, "height": 950})
        page = context.new_page()

        # Bridge Native DesktopApi methods to the browser window (simulating pywebview runtime)
        page.expose_function("api_get_default_paths", lambda: api.get_default_paths())
        page.expose_function("api_process_journal", lambda filename, tpl, out_dir: api.process_journal(filename, tpl, out_dir))
        page.expose_function("api_process_both", lambda tpl, out_dir: api.process_both(tpl, out_dir))
        page.expose_function("api_select_file", lambda: api.select_file())
        page.expose_function("api_select_folder", lambda: api.select_folder())
        page.expose_function("api_open_in_excel", lambda p=None: api.open_in_excel(p))
        page.expose_function("api_open_in_explorer", lambda p=None: api.open_in_explorer(p))
        page.expose_function("api_get_sample_csv", lambda: api.get_sample_csv())

        page.add_init_script("""
        window.pywebview = {
            api: {
                get_default_paths: () => window.api_get_default_paths(),
                process_journal: (f, t, o) => window.api_process_journal(f, t, o),
                process_both: (t, o) => window.api_process_both(t, o),
                select_file: () => window.api_select_file(),
                select_folder: () => window.api_select_folder(),
                open_in_excel: (p) => window.api_open_in_excel(p),
                open_in_explorer: (p) => window.api_open_in_explorer(p),
                get_sample_csv: () => window.api_get_sample_csv()
            }
        };
        window.dispatchEvent(new Event('pywebviewready'));
        """)

        print("[*] Navigating to Desktop Studio UI...")
        page.goto(file_url, wait_until="networkidle")

        # 1. Verify Desktop App Titlebar & Branding
        print("[*] Step 1: Verifying Desktop Application Header & Titlebar...")
        title = page.title()
        print(f"  [+] Page title: '{title}'")
        assert "Workday" in title and "Enterprise Desktop" in title, f"Unexpected title: {title}"

        header_text = page.locator("header").text_content()
        assert "Workday Accounting Journal EIB Ingest Studio" in header_text
        assert "DESKTOP EDITION (.EXE)" in header_text
        assert "v45.1" in header_text
        assert "Local Engine Active" in header_text
        print("  [+] Titlebar branding & desktop badges verified successfully.")

        # 2. Verify Panel 1: Source Accounting File selection
        print("[*] Step 2: Testing Panel 1 (Source File Input & Presets)...")
        source_input = page.locator("#sourcePathInput")
        page.wait_for_selector("#sourcePathInput")
        initial_source = source_input.input_value()
        print(f"  [+] Default source path input: '{initial_source}'")
        assert "Accounting_Journal_Raw_LE01_202601.xlsx" in initial_source or "2026" in initial_source

        # Test selecting 2025 source preset
        print("  [*] Clicking 2025 Preset Button (8 Lines)...")
        page.locator("#presetBtn-2025").click()
        page.wait_for_timeout(200)
        source_label_2025 = page.locator("#sourceFileNameLabel").text_content()
        assert "202501" in source_label_2025
        print(f"  [+] Active source switched to: '{source_label_2025}'")

        # Test selecting 2026 source preset
        print("  [*] Clicking 2026 Preset Button (222 Lines)...")
        page.locator("#presetBtn-2026").click()
        page.wait_for_timeout(200)
        source_label_2026 = page.locator("#sourceFileNameLabel").text_content()
        assert "202601" in source_label_2026
        print(f"  [+] Active source switched back to: '{source_label_2026}'")

        # 3. Verify Panel 2: Local Disk Save Location & Workday Naming Standards
        print("[*] Step 3: Testing Panel 2 (Save Location & Naming Standard Display)...")
        save_input = page.locator("#saveLocationInput")
        save_val = save_input.input_value()
        print(f"  [+] Configured Save Destination: '{save_val}'")
        assert len(save_val) > 0, "Save location should not be empty"

        browse_folder_btn = page.locator("#btnBrowseFolder")
        assert browse_folder_btn.is_visible()
        print(f"  [+] 'Browse Save Location...' button is present and visible.")

        naming_card_2026 = page.locator("#namingCard2026").text_content()
        naming_card_2025 = page.locator("#namingCard2025").text_content()
        assert "Import_Accounting_Journal_v45.0_LE01_2026_P2.xlsx" in naming_card_2026
        assert "Import_Accounting_Journal_v45.0_LE01_2025_P1.xlsx" in naming_card_2025
        print(f"  [+] Verified 2026 Output Naming Standard: Import_Accounting_Journal_v45.0_LE01_2026_P2.xlsx")
        print(f"  [+] Verified 2025 Output Naming Standard: Import_Accounting_Journal_v45.0_LE01_2025_P1.xlsx")

        # 4. Verify Panel 3: Target Workday EIB Template Selector (v45 to v48)
        print("[*] Step 4: Testing Panel 3 (Target Template Selector v45 - v48)...")
        for t_id in ["v45", "v46", "v47", "v48"]:
            card = page.locator(f"#tplCard-{t_id}")
            assert card.is_visible(), f"Template card #{t_id} must be visible"
        
        # Test switching to v48.0
        page.locator("#tplCard-v48").click()
        page.wait_for_timeout(200)
        badge_text_v48 = page.locator("#selectedTemplateBadge").text_content()
        top_badge_v48 = page.locator("#topVersionBadge").text_content()
        assert "v48.0" in badge_text_v48
        assert "v48.0" in top_badge_v48
        print("  [+] Successfully selected v48.0 template.")

        # Test switching to v46.0 and v47.0
        page.locator("#tplCard-v46").click()
        page.wait_for_timeout(100)
        assert "v46.0" in page.locator("#selectedTemplateBadge").text_content()

        page.locator("#tplCard-v47").click()
        page.wait_for_timeout(100)
        assert "v47.0" in page.locator("#selectedTemplateBadge").text_content()

        # Switch back to standard two-sheet v45.1
        page.locator("#tplCard-v45").click()
        page.wait_for_timeout(100)
        assert "v45.1" in page.locator("#selectedTemplateBadge").text_content()
        print("  [+] Template selector grid responsive and verified.")

        # 5. Execute 2026 Ingestion via Primary Action Button
        print("[*] Step 5: Executing 2026 Source Ingestion via Native Desktop IPC...")
        page.locator("#presetBtn-2026").click()
        page.wait_for_timeout(200)
        page.locator("#btnProcess").click()

        # Wait for execution status box to complete for 2026
        page.wait_for_selector("#execMessage:has-text('Import_Accounting_Journal_v45.0_LE01_2026_P2.xlsx')", timeout=20000)
        page.wait_for_selector("#btnProcess:not([disabled])", timeout=20000)
        exec_badge = page.locator("#execBadge").text_content()
        exec_msg = page.locator("#execMessage").text_content()
        print(f"  [+] Execution Badge: '{exec_badge}'")
        print(f"  [+] Execution Message: '{exec_msg}'")
        assert "100% BALANCED" in exec_badge
        assert "Import_Accounting_Journal_v45.0_LE01_2026_P2.xlsx" in exec_msg

        # 6. Execute 2025 Ingestion via Primary Action Button
        print("[*] Step 6: Executing 2025 Source Ingestion via Native Desktop IPC...")
        page.locator("#presetBtn-2025").click()
        page.wait_for_timeout(200)
        page.locator("#btnProcess").click()

        # Wait for execution status box to complete for 2025
        page.wait_for_selector("#execMessage:has-text('Import_Accounting_Journal_v45.0_LE01_2025_P1.xlsx')", timeout=20000)
        page.wait_for_selector("#btnProcess:not([disabled])", timeout=20000)
        exec_badge_2025 = page.locator("#execBadge").text_content()
        exec_msg_2025 = page.locator("#execMessage").text_content()
        print(f"  [+] 2025 Execution Badge: '{exec_badge_2025}'")
        print(f"  [+] 2025 Execution Message: '{exec_msg_2025}'")
        assert "100% BALANCED" in exec_badge_2025
        assert "Import_Accounting_Journal_v45.0_LE01_2025_P1.xlsx" in exec_msg_2025

        # 7. Check Terminal Console Log Content
        print("[*] Step 7: Verifying Execution Terminal Console Logs...")
        terminal_text = page.locator("#terminalLog").text_content()
        assert "Net Variance: $0.00 (100% BALANCED)" in terminal_text
        assert "Strict Null/Blank handling" in terminal_text
        assert "SUCCESS" in terminal_text
        print("  [+] Terminal log verified: Balance verified, strict null memo confirmed, status SUCCESS.")

        # 8. Take Desktop Screenshot for Verification
        screenshot_path = "playwright_desktop_app_verification.png"
        page.screenshot(path=screenshot_path, full_page=True)
        print(f"  [+] Captured high-resolution screenshot to: {screenshot_path}")

        browser.close()

    print("\n[SUCCESS] Playwright Desktop E2E Test Suite completely PASSED without server.py!")

    # 9. Verify Generated Excel Files On Local Disk
    print("\n[*] Step 8: Direct openpyxl file inspection on generated output files...")
    out_dir = r"c:\CSV to EIB\output"
    file_2026 = os.path.join(out_dir, "Import_Accounting_Journal_v45.0_LE01_2026_P2.xlsx")
    file_2025 = os.path.join(out_dir, "Import_Accounting_Journal_v45.0_LE01_2025_P1.xlsx")

    assert os.path.exists(file_2026), f"Expected 2026 output file does not exist: {file_2026}"
    assert os.path.exists(file_2025), f"Expected 2025 output file does not exist: {file_2025}"

    # Inspect 2026 file
    wb_2026 = openpyxl.load_workbook(file_2026, data_only=True)
    ws_2026 = wb_2026["Journal Entry Line Replacement"] if "Journal Entry Line Replacement" in wb_2026.sheetnames else wb_2026.active
    memos_2026 = [ws_2026.cell(row=r, column=12).value for r in range(6, ws_2026.max_row + 1)]
    null_memos_2026 = [m for m in memos_2026 if m is not None and str(m).strip().upper() == "NULL"]
    print(f"  [+] 2026 File ({ws_2026.max_row - 5} lines): null_memos count = {len(null_memos_2026)} (Must be 0)")
    assert len(null_memos_2026) == 0, f"Found literal 'NULL' in 2026 Memo: {null_memos_2026}"

    # Inspect 2025 file
    wb_2025 = openpyxl.load_workbook(file_2025, data_only=True)
    ws_2025 = wb_2025["Journal Entry Line Replacement"] if "Journal Entry Line Replacement" in wb_2025.sheetnames else wb_2025.active
    memos_2025 = [ws_2025.cell(row=r, column=12).value for r in range(6, ws_2025.max_row + 1)]
    null_memos_2025 = [m for m in memos_2025 if m is not None and str(m).strip().upper() == "NULL"]
    print(f"  [+] 2025 File ({ws_2025.max_row - 5} lines): null_memos count = {len(null_memos_2025)} (Must be 0)")
    assert len(null_memos_2025) == 0, f"Found literal 'NULL' in 2025 Memo: {null_memos_2025}"

    print("[SUCCESS] All files and LineMemo blank assertions verified 100% cleanly!")

if __name__ == "__main__":
    run_e2e_tests()
