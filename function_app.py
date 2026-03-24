"""
Uniram AI Support — Azure Function Entry Point
Timer Trigger: runs every 30 minutes to process support@uniram.com

v15: KB persistence fix — always download latest KB from Blob on cold start,
     and upload updated KB back to Blob after every run.
"""
# ── sqlite3 patch: must be done BEFORE importing chromadb ──
import sys as _sys
import os as _os

# Patch sys.path to include bundled packages first
_func_dir = _os.path.dirname(_os.path.abspath(__file__))
_pkg_dir = _os.path.join(_func_dir, ".python_packages", "lib", "site-packages")
if _pkg_dir not in _sys.path:
    _sys.path.insert(0, _pkg_dir)

# Replace sqlite3 with pysqlite3-binary (sqlite >= 3.35.0 required by ChromaDB)
try:
    import pysqlite3 as _pysqlite3
    _sys.modules["sqlite3"] = _pysqlite3
except ImportError:
    pass  # Fall back to system sqlite3

import azure.functions as func
import logging
import os, sys, zipfile, shutil

app = func.FunctionApp()

KB_PATH      = "/tmp/uniram_kb"
KB_ZIP_PATH  = "/tmp/knowledge_base.zip"
KB_BLOB_NAME = "knowledge_base.zip"
KB_CONTAINER = "uniram-support-kb"


@app.timer_trigger(
    schedule="0 */30 * * * *",   # every 30 minutes
    arg_name="myTimer",
    run_on_startup=False,
    use_monitor=False
)
def support_ai_reply(myTimer: func.TimerRequest) -> None:
    if myTimer.past_due:
        logging.warning("Timer is past due — running now anyway.")
    logging.info("Uniram AI Support: starting email processing run...")

    # ── Patch sys.path so our modules are importable ──
    func_dir = os.path.dirname(__file__)
    if func_dir not in sys.path:
        sys.path.insert(0, func_dir)

    # ── Always download latest KB from Blob Storage (ensures cold-start safety) ──
    _download_knowledge_base()

    # ── Run the AI reply engine ──
    try:
        from engine.ai_reply_engine import process_emails
        process_emails(dry_run=False, kb_path=KB_PATH)
        logging.info("Uniram AI Support: run complete.")
    except Exception as e:
        logging.error(f"Uniram AI Support ERROR: {e}", exc_info=True)
        raise
    finally:
        # ── Always upload updated KB back to Blob Storage ──
        try:
            _upload_knowledge_base()
        except Exception as e:
            logging.warning(f"KB upload failed (non-fatal): {e}")


def _download_knowledge_base():
    """Always download the latest knowledge_base.zip from Blob Storage."""
    conn_str = os.environ.get("AzureWebJobsStorage", "")
    sas_url  = os.environ.get("KB_SAS_URL", "")

    # Clean up any stale /tmp KB to ensure fresh load
    if os.path.exists(KB_PATH):
        shutil.rmtree(KB_PATH, ignore_errors=True)
    if os.path.exists(KB_ZIP_PATH):
        os.remove(KB_ZIP_PATH)

    if conn_str:
        # Preferred: use AzureWebJobsStorage connection string (no expiry)
        logging.info("Downloading KB via AzureWebJobsStorage connection string...")
        from azure.storage.blob import BlobServiceClient
        client = BlobServiceClient.from_connection_string(conn_str)
        blob = client.get_blob_client(container=KB_CONTAINER, blob=KB_BLOB_NAME)
        with open(KB_ZIP_PATH, "wb") as f:
            data = blob.download_blob()
            data.readinto(f)
    elif sas_url:
        # Fallback: use pre-signed SAS URL
        logging.info("Downloading KB via SAS URL...")
        import requests as req
        with req.get(sas_url, stream=True, timeout=120) as resp:
            resp.raise_for_status()
            with open(KB_ZIP_PATH, "wb") as f:
                for chunk in resp.iter_content(chunk_size=1024 * 1024):
                    f.write(chunk)
    else:
        raise ValueError("Neither AzureWebJobsStorage nor KB_SAS_URL is set.")

    os.makedirs(KB_PATH, exist_ok=True)
    with zipfile.ZipFile(KB_ZIP_PATH, "r") as zf:
        zf.extractall(KB_PATH)
    logging.info(f"Knowledge base extracted to {KB_PATH}")


def _upload_knowledge_base():
    """Zip the updated KB and upload back to Blob Storage for persistence."""
    conn_str = os.environ.get("AzureWebJobsStorage", "")
    if not conn_str:
        logging.warning("AzureWebJobsStorage not set — skipping KB upload.")
        return

    if not os.path.exists(KB_PATH):
        logging.warning("KB path does not exist — nothing to upload.")
        return

    # Re-zip the KB directory
    updated_zip = "/tmp/knowledge_base_updated.zip"
    with zipfile.ZipFile(updated_zip, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(KB_PATH):
            for file in files:
                full_path = os.path.join(root, file)
                arcname = os.path.relpath(full_path, KB_PATH)
                zf.write(full_path, arcname)

    # Upload to Blob Storage (overwrite)
    from azure.storage.blob import BlobServiceClient
    client = BlobServiceClient.from_connection_string(conn_str)
    blob = client.get_blob_client(container=KB_CONTAINER, blob=KB_BLOB_NAME)
    with open(updated_zip, "rb") as f:
        blob.upload_blob(f, overwrite=True)

    zip_size = os.path.getsize(updated_zip) / 1024
    logging.info(f"KB uploaded back to Blob Storage ({zip_size:.1f} KB)")

    # Cleanup temp zip
    os.remove(updated_zip)
