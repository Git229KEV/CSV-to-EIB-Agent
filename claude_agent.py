import os
import sys
import json
import urllib.request
import urllib.error
import time

# Support PyInstaller runtime extraction path
if hasattr(sys, '_MEIPASS'):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE = os.path.join(os.path.abspath("."), "claude_config.json")
DEFAULT_MODEL = "claude-sonnet-5-5"
FALLBACK_MODEL = "claude-haiku-5-5"
SUPPORTED_MODELS = [
    "claude-sonnet-5-5",
    "claude-haiku-5-5",
    "claude-opus-5-5",
    "claude-fable-5-2",
]

SKILL_RELATIVE_PATH = os.path.join("skills", "workday-eib-mapper", "SKILL.md")


def load_skill_text():
    """Loads the Workday EIB Mapping Skill instructions from disk or bundled resource."""
    candidates = [
        os.path.join(BASE_DIR, SKILL_RELATIVE_PATH),
        os.path.join(".", SKILL_RELATIVE_PATH),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), SKILL_RELATIVE_PATH),
    ]
    for cand in candidates:
        if os.path.exists(cand):
            try:
                with open(cand, "r", encoding="utf-8") as f:
                    return f.read(), cand
            except Exception:
                pass

    # Built-in fallback skill instructions if file is unreadable
    fallback_skill = """
# Workday Accounting Journal EIB Mapping Skill
- Two-sheet template v45.1: Sheet 1 (Import Accounting Journal) headers, Sheet 2 (Journal Entry Line Replacement) lines.
- Single-sheet templates v46.0, v47.0, v48.0: Submit Accounting Journal.
- Strip underscores directly from Intercompany Affiliate (e.g. 'LE_04' -> 'LE04').
- Format Cost Center as 3 digits with leading zeros (e.g. '0' -> '000', '90' -> '090').
- Map Ledger Account to Column F, Account Set to Column G.
- Debit and Credit amounts: zeros in source must be populated as 0.0.
- Empty or 'NULL' LineMemo must remain strictly blank.
"""
    return fallback_skill, "embedded_builtin"


class ClaudeAgent:
    """Claude AI Agent dedicated to Workday Accounting Journal EIB Schema Mapping."""

    def __init__(self, api_key=None, model=None):
        self._api_key = api_key or self._discover_api_key()
        self._model = model or self._discover_model()
        self.skill_text, self.skill_source = load_skill_text()

    def _discover_api_key(self):
        """Looks for Claude / Anthropic API key in environment or local config."""
        # 1. Environment variables
        for var in ("ANTHROPIC_API_KEY", "CLAUDE_API_KEY", "ANTHROPIC_KEY"):
            val = os.environ.get(var)
            if val and val.strip():
                return val.strip()

        # 2. Local claude_config.json
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    cfg = json.load(f)
                    val = cfg.get("api_key")
                    if val and str(val).strip():
                        return str(val).strip()
            except Exception:
                pass

        # 3. User home .claude config / key if present
        home_key_file = os.path.join(os.path.expanduser("~"), ".anthropic_api_key")
        if os.path.exists(home_key_file):
            try:
                with open(home_key_file, "r", encoding="utf-8") as f:
                    k = f.read().strip()
                    if k:
                        return k
            except Exception:
                pass

        return None

    def _discover_model(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    cfg = json.load(f)
                    val = cfg.get("model")
                    if val and str(val).strip():
                        return str(val).strip()
            except Exception:
                pass
        return os.environ.get("ANTHROPIC_MODEL", DEFAULT_MODEL)

    def set_config(self, api_key=None, model=None):
        """Updates and persists Claude AI Agent credentials locally."""
        if api_key is not None:
            self._api_key = str(api_key).strip() if str(api_key).strip() else None
        if model is not None and str(model).strip():
            self._model = str(model).strip()

        try:
            cfg = {}
            if os.path.exists(CONFIG_FILE):
                try:
                    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                        cfg = json.load(f)
                except Exception:
                    cfg = {}
            if self._api_key:
                cfg["api_key"] = self._api_key
            elif "api_key" in cfg:
                del cfg["api_key"]
            cfg["model"] = self._model
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(cfg, f, indent=2)
            return True
        except Exception as e:
            print(f"[ClaudeAgent] Failed to persist config: {e}")
            return False

    def is_available(self):
        return bool(self._api_key and len(self._api_key) > 5)

    def get_status(self):
        masked_key = ""
        if self._api_key:
            if len(self._api_key) > 12:
                masked_key = f"{self._api_key[:6]}...{self._api_key[-4:]}"
            else:
                masked_key = "***configured***"

        return {
            "available": self.is_available(),
            "has_key": bool(self._api_key),
            "masked_key": masked_key,
            "model": self._model,
            "skill_loaded": bool(self.skill_text),
            "skill_source": self.skill_source,
            "config_file": CONFIG_FILE,
        }

    def _call_anthropic_api(self, messages, system_prompt="", max_tokens=2500, temperature=None):
        """Makes direct HTTP POST to Anthropic Messages API using standard library (zero external dependency)."""
        if not self.is_available():
            raise ValueError("Anthropic API key is not configured. Please provide an API key.")

        url = "https://api.anthropic.com/v1/messages"
        headers = {
            "x-api-key": self._api_key,
            "anthropic-version": "2023-06-01",
            "content-type": "application/json",
            "user-agent": "WorkdayEIBStudio-ClaudeAgent/1.0",
        }

        payload = {
            "model": self._model,
            "max_tokens": max_tokens,
            "messages": messages,
        }
        if system_prompt and system_prompt.strip():
            payload["system"] = system_prompt.strip()

        # Omit temperature for 5.x / fable / reasoning models where it is deprecated by Anthropic
        m_lower = (self._model or "").lower()
        deprecates_temperature = any(token in m_lower for token in ("fable", "5-2", "5-5", "-5-", "opus-5", "sonnet-5", "haiku-5"))
        if temperature is not None and not deprecates_temperature:
            payload["temperature"] = temperature

        def _execute(data_payload):
            req = urllib.request.Request(
                url,
                data=json.dumps(data_payload).encode("utf-8"),
                headers=headers,
                method="POST",
            )
            with urllib.request.urlopen(req, timeout=60) as resp:
                resp_data = json.loads(resp.read().decode("utf-8"))
                content_blocks = resp_data.get("content", [])
                text_parts = [b.get("text", "") for b in content_blocks if b.get("type") == "text"]
                return "".join(text_parts), resp_data

        try:
            return _execute(payload)
        except urllib.error.HTTPError as e:
            err_body = e.read().decode("utf-8", errors="replace")
            # Auto-recovery: If temperature was rejected, retry immediately without it
            if "temperature" in err_body and "temperature" in payload:
                del payload["temperature"]
                try:
                    return _execute(payload)
                except Exception:
                    pass
            raise RuntimeError(f"Claude API error HTTP {e.code}: {err_body}")
        except Exception as e:
            raise RuntimeError(f"Network error calling Claude API: {str(e)}")

    def test_connection(self):
        """Tests Claude API key validity with a lightweight test ping."""
        if not self.is_available():
            return {"success": False, "error": "No API key provided."}
        try:
            msg = [{"role": "user", "content": "Respond with the word READY"}]
            text, _ = self._call_anthropic_api(msg, system_prompt="You are a ping tester.", max_tokens=10)
            return {
                "success": True,
                "message": text.strip(),
                "model": self._model,
            }
        except Exception as e:
            return {"success": False, "error": str(e)}

    def analyze_and_map(self, source_columns, sample_rows, template_summary):
        """
        Invokes Claude AI Agent with the Workday EIB Mapping Skill to analyze source columns,
        semantic relationships, and target template columns, returning structured mappings.
        """
        if not self.is_available():
            return {
                "success": False,
                "error": "Claude AI Agent is not configured with an API key. Using rule-based fallback.",
                "used_ai": False,
            }

        system_prompt = f"""You are the Workday Accounting Journal EIB Ingest AI Agent.
Your job is to analyze raw source accounting columns and map them intelligently to Workday EIB spreadsheet columns.

You MUST adhere to the following domain skill instructions:
{self.skill_text}

OUTPUT INSTRUCTION:
Return ONLY a valid JSON object (no markdown code blocks, no preamble, no commentary outside JSON).
Keep auto_discovered_mappings concise and only for novel or non-standard source columns not already covered in core_mappings.
The JSON must follow this exact schema:
{{
  "reasoning_summary": "High-level explanation of source format and mapping strategy",
  "core_mappings": {{
    "header_key": "source_column_name or null",
    "line_key": "source_column_name or null",
    "line_order": "source_column_name or null",
    "company": "source_column_name or null",
    "ledger_account": "source_column_name or null",
    "account_set": "source_column_name or null",
    "debit": "source_column_name or null",
    "credit": "source_column_name or null",
    "currency": "source_column_name or null",
    "memo": "source_column_name or null",
    "budget_date": "source_column_name or null",
    "cost_center": "source_column_name or null",
    "intercompany_affiliate": "source_column_name or null",
    "spend_category": "source_column_name or null",
    "revenue_category": "source_column_name or null",
    "balancing_worktag": "source_column_name or null"
  }},
  "auto_discovered_mappings": [
    {{
      "source_col": "exact_source_column_name",
      "target_sheet": "Import Accounting Journal or Journal Entry Line Replacement or Submit Accounting Journal",
      "target_header": "target_column_header",
      "target_col_idx": 12,
      "confidence": 0.95,
      "reasoning": "Explanation"
    }}
  ],
  "special_rules": {{
    "strip_intercompany_underscores": true,
    "pad_cost_center_3_digits": true,
    "zero_amounts_as_float": true
  }}
}}
"""

        user_content = {
            "task": "Map source accounting columns to Workday EIB spreadsheet template.",
            "source_columns": source_columns,
            "sample_rows_preview": sample_rows[:3] if sample_rows else [],
            "target_template_info": template_summary,
        }

        messages = [
            {"role": "user", "content": json.dumps(user_content, indent=2, default=str)}
        ]

        try:
            start_t = time.time()
            raw_text, raw_resp = self._call_anthropic_api(messages, system_prompt=system_prompt, max_tokens=8192)
            elapsed = round(time.time() - start_t, 2)

            # Clean JSON from potential markdown wrapping ```json ... ```
            clean_text = raw_text.strip()
            if clean_text.startswith("```"):
                lines = clean_text.splitlines()
                if lines[0].startswith("```"):
                    lines = lines[1:]
                if lines and lines[-1].startswith("```"):
                    lines = lines[:-1]
                clean_text = "\n".join(lines).strip()

            parsed = None
            try:
                parsed = json.loads(clean_text)
            except Exception:
                # Attempt regex recovery if trailing brackets were truncated
                idx_open = clean_text.find("{")
                idx_close = clean_text.rfind("}")
                if idx_open != -1 and idx_close != -1 and idx_close > idx_open:
                    try:
                        parsed = json.loads(clean_text[idx_open:idx_close+1])
                    except Exception:
                        pass
                if not parsed:
                    raise

            return {
                "success": True,
                "used_ai": True,
                "model": self._model,
                "elapsed_seconds": elapsed,
                "reasoning_summary": parsed.get("reasoning_summary", ""),
                "core_mappings": parsed.get("core_mappings", {}),
                "auto_discovered_mappings": parsed.get("auto_discovered_mappings", []),
                "special_rules": parsed.get("special_rules", {}),
                "raw_response": parsed,
            }
        except Exception as e:
            return {
                "success": False,
                "used_ai": False,
                "error": f"Claude AI Agent mapping failed: {str(e)}",
                "fallback_to_rules": True,
            }


# Singleton accessor
_agent_instance = None

def get_claude_agent():
    global _agent_instance
    if _agent_instance is None:
        _agent_instance = ClaudeAgent()
    return _agent_instance
