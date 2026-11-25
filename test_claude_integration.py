import os
import sys
import json
import unittest

from claude_agent import ClaudeAgent, get_claude_agent, load_skill_text
import run_agent


class TestClaudeAIIntegration(unittest.TestCase):

    def test_skill_loading(self):
        """Verify the Workday EIB Mapping Skill loads correctly from SKILL.md."""
        skill_text, skill_path = load_skill_text()
        self.assertTrue(len(skill_text) > 200, "Skill text should be non-empty")
        self.assertIn("workday-eib-mapper", skill_text)
        self.assertIn("spend category 1.xlsx", skill_text)
        self.assertIn("LE04", skill_text)
        self.assertIn("3-digit", skill_text)
        print(f"[PASS] Skill loaded successfully from: {skill_path} ({len(skill_text)} chars)")

    def test_claude_agent_status_and_config(self):
        """Verify ClaudeAgent initialization and configuration persistence."""
        agent = get_claude_agent()
        status = agent.get_status()
        self.assertIn("available", status)
        self.assertIn("model", status)
        self.assertTrue(status["skill_loaded"])
        print(f"[PASS] ClaudeAgent status verified: model={status['model']}, skill_loaded={status['skill_loaded']}")

    def test_schema_analysis_fallback_resilience(self):
        """Verify schema analysis runs cleanly and provides fallback if Claude API key is absent or offline."""
        src_2026 = run_agent.find_source_file("Accounting_Journal_Raw_LE01_202601.xlsx")
        report = run_agent.analyze_schema(src_2026, "v45", use_claude=True)
        self.assertTrue(len(report["source_columns"]) > 20)
        self.assertEqual(report["total_rows"], 222)
        print(f"[PASS] Schema analysis completed with use_claude=True (rows={report['total_rows']}, auto_mappings={len(report['auto_discovered_mappings'])})")

    def test_mock_claude_ai_mapping_plan(self):
        """Verify that when Claude AI agent returns structured mappings, they correctly merge into the plan."""
        mock_claude_plan = {
            "success": True,
            "used_ai": True,
            "model": "claude-sonnet-5-5",
            "reasoning_summary": "Identified Trading_Partner as Intercompany Affiliate and Exclude_Spend as Exclude from Spend Report",
            "core_mappings": {
                "intercompany_affiliate": "Affliate Company",
                "cost_center": "Worktag_Cost_Center_Reference_ID",
            },
            "auto_discovered_mappings": [
                {
                    "source_col": "Exclude_Spend",
                    "target_sheet": "Journal Entry Line Replacement",
                    "target_header": "Exclude from Spend Report",
                    "confidence": 0.99,
                    "reasoning": "Direct semantic match to Exclude from Spend Report flag"
                }
            ]
        }

        src_cols = ["ROW_ID", "LineOrder", "LedgerAccountReferenceID", "DebitAmount", "CreditAmount", "Affliate Company", "Exclude_Spend"]
        col_lookup = run_agent._build_col_lookup(src_cols)
        
        import openpyxl
        wb = openpyxl.load_workbook(run_agent.find_target_file("Import_Accounting_Journal_v45.1.template.xlsx"))
        col_inter, col_bal, hdr_auto, line_auto = run_agent._plan_two_sheet_mappings(
            src_cols, wb[run_agent.SHEET_HEADER], wb[run_agent.SHEET_LINES], col_lookup, claude_plan=mock_claude_plan
        )
        wb.close()

        self.assertIsNotNone(col_inter)
        mapped_srcs = [m["source_col"] for m in line_auto]
        self.assertIn("Exclude_Spend", mapped_srcs)
        exclude_map = next(m for m in line_auto if m["source_col"] == "Exclude_Spend")
        self.assertIn("claude_ai", exclude_map["match_method"])
        self.assertEqual(exclude_map["confidence"], 0.99)
        print(f"[PASS] Claude AI mock plan merged accurately into line_auto mappings: {exclude_map}")

    def test_full_pipeline_verification(self):
        """Verify process_journal runs end-to-end with spend category 1.xlsx (155 rules), LE04, and balanced math."""
        res_2026 = run_agent.process_journal("Accounting_Journal_Raw_LE01_202601.xlsx", "v45", use_claude=False)
        self.assertTrue(res_2026["success"])
        self.assertEqual(res_2026["variance"], 0.0)
        self.assertEqual(res_2026["spend_categories_mapped"], 87)
        self.assertTrue(res_2026["is_balanced"])

        res_2025 = run_agent.process_journal("Accounting_Journal_Raw_LE01_202501.xlsx", "v45", use_claude=False)
        self.assertTrue(res_2025["success"])
        self.assertEqual(res_2025["variance"], 0.0)
        self.assertEqual(res_2025["spend_categories_mapped"], 2)
        print("[PASS] Full pipeline verified: 2026 mapped 87 spend categories, 2025 mapped 2 spend categories, both 100% balanced ($0.00 variance).")


if __name__ == "__main__":
    unittest.main()
