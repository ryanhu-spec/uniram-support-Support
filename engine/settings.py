"""
Settings for Uniram AI Support — Azure Function version.
All secrets are read from Azure Function Application Settings (environment variables).
DO NOT hardcode credentials here. Set them in Azure Portal → Function App → Configuration.
"""
import os

# ── Microsoft Graph API ──────────────────────────────────────────────────────
GRAPH_TENANT_ID     = os.environ.get("GRAPH_TENANT_ID",     "")
GRAPH_CLIENT_ID     = os.environ.get("GRAPH_CLIENT_ID",     "")
GRAPH_CLIENT_SECRET = os.environ.get("GRAPH_CLIENT_SECRET", "")

# ── OpenAI ───────────────────────────────────────────────────────────────────
OPENAI_API_KEY      = os.environ.get("OPENAI_API_KEY",      "")

# ── Mailboxes ────────────────────────────────────────────────────────────────
SUPPORT_MAILBOX     = os.environ.get("SUPPORT_MAILBOX",     "support@uniram.com")

# ── Azure Storage ────────────────────────────────────────────────────────────
AZURE_STORAGE_CONNECTION_STRING = os.environ.get("AZURE_STORAGE_CONNECTION_STRING", "")

# ── Escalation ───────────────────────────────────────────────────────────────
ESCALATION_EMAIL    = os.environ.get("ESCALATION_EMAIL",    "ken@uniram.com")
ESCALATION_CC       = os.environ.get("ESCALATION_CC",       "finn.sun@uniram.com")

# ── AI Models ────────────────────────────────────────────────────────────────
CLASSIFY_MODEL  = "gpt-4o-mini"
REPLY_MODEL     = "gpt-4o"
EMBED_MODEL     = "text-embedding-3-small"

# ── Tuning ───────────────────────────────────────────────────────────────────
CONFIDENCE_THRESHOLD         = float(os.environ.get("CONFIDENCE_THRESHOLD", "0.65"))
ESCALATION_CONFIDENCE_THRESHOLD = CONFIDENCE_THRESHOLD
MAX_EMAILS_PER_RUN           = int(os.environ.get("MAX_EMAILS_PER_RUN", "20"))
SCAN_BATCH_SIZE              = 50
SCAN_MAX_EMAILS              = None

# ── Knowledge base path (overridden at runtime by function_app.py) ───────────
KB_PATH = os.environ.get("KB_PATH", "/tmp/uniram_kb")
