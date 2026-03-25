"""
Uniram AI Support — Azure Function Entry Point
Timer Trigger: runs every 30 minutes to process support@uniram.com

v20: Ken approval flow — Approve/Reject HTTP triggers for draft replies
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


@app.route(route="approve_support_reply", methods=["GET", "POST"])
def approve_support_reply(req: func.HttpRequest) -> func.HttpResponse:
    """
    HTTP endpoint for Ken to Approve or Reject Jennifer's draft reply.

    Query params:
        token:  Pending approval token (UUID)
        action: 'approve' or 'reject'

    Approve flow:
        - Sends Jennifer's draft to the customer (CC Ken)
        - Marks token as approved
        - Jennifer learns from the approval next cycle

    Reject flow:
        - Returns an HTML page where Ken can type the correct answer
        - On submit (POST with 'correct_answer'), sends Ken's answer to customer (CC Ken)
        - Saves Q&A to KB immediately
    """
    func_dir = os.path.dirname(__file__)
    if func_dir not in sys.path:
        sys.path.insert(0, func_dir)

    logging.info("approve_support_reply endpoint called")

    token  = req.params.get("token", "")
    action = req.params.get("action", "")

    # Handle POST from reject form (Ken submitting correct answer)
    correct_answer = ""
    if req.method == "POST":
        try:
            body = req.get_json()
            token          = body.get("token", token)
            action         = body.get("action", action)
            correct_answer = body.get("correct_answer", "")
        except Exception:
            try:
                form = req.form
                token          = form.get("token", token)
                action         = form.get("action", action)
                correct_answer = form.get("correct_answer", "")
            except Exception:
                pass

    if not token:
        return func.HttpResponse(
            body=_error_page("Missing token parameter."),
            mimetype="text/html", status_code=400
        )

    try:
        from engine.pending_approvals import get_pending, mark_done
        pending = get_pending(token)
    except Exception as e:
        logging.error(f"approve_support_reply: failed to load pending store: {e}")
        return func.HttpResponse(
            body=_error_page(f"Storage error: {e}"),
            mimetype="text/html", status_code=500
        )

    if not pending:
        return func.HttpResponse(
            body=_error_page("This approval link has expired or was already processed."),
            mimetype="text/html", status_code=404
        )

    if pending.get("status") != "pending":
        return func.HttpResponse(
            body=_already_done_page(pending),
            mimetype="text/html", status_code=200
        )

    # ── APPROVE ──────────────────────────────────────────────────────────────
    if action == "approve":
        try:
            _send_approved_reply(pending)
            mark_done(token, "approved")
            logging.info(f"[Approve] token={token} | to={pending['sender_email']}")
            return func.HttpResponse(
                body=_approve_success_page(pending),
                mimetype="text/html", status_code=200
            )
        except Exception as e:
            logging.error(f"approve_support_reply approve error: {e}", exc_info=True)
            return func.HttpResponse(
                body=_error_page(str(e)),
                mimetype="text/html", status_code=500
            )

    # ── REJECT — show form (GET) or process answer (POST) ────────────────────
    elif action == "reject":
        if req.method == "POST" and correct_answer.strip():
            try:
                _send_ken_answer(pending, correct_answer)
                _learn_from_rejection(pending, correct_answer)
                mark_done(token, "rejected_and_taught")
                logging.info(f"[Reject+Teach] token={token} | to={pending['sender_email']}")
                return func.HttpResponse(
                    body=_teach_success_page(pending),
                    mimetype="text/html", status_code=200
                )
            except Exception as e:
                logging.error(f"approve_support_reply reject error: {e}", exc_info=True)
                return func.HttpResponse(
                    body=_error_page(str(e)),
                    mimetype="text/html", status_code=500
                )
        else:
            # Show the reject/teach form
            return func.HttpResponse(
                body=_reject_form_page(pending, token),
                mimetype="text/html", status_code=200
            )

    else:
        return func.HttpResponse(
            body=_error_page(f"Unknown action: '{action}'. Use 'approve' or 'reject'."),
            mimetype="text/html", status_code=400
        )


# ── Helper: send Jennifer's approved draft to customer ───────────────────────

def _send_approved_reply(pending: dict):
    """Send Jennifer's draft reply to the customer, CC Ken."""
    func_dir = os.path.dirname(__file__)
    if func_dir not in sys.path:
        sys.path.insert(0, func_dir)

    from engine.ai_reply_engine import (
        get_graph_token, send_email, SUPPORT_MAILBOX,
        ESCALATION_EMAIL, ESCALATION_CC
    )

    customer_email = pending["sender_email"]
    subject        = pending["subject"]
    draft          = pending["draft_reply"]
    sender_name    = pending.get("sender_name", "")

    # Build a clean reply HTML
    draft_html = draft.replace("\n", "<br>")
    body_html = f"""<div style="font-family:Arial,sans-serif;font-size:14px;color:#333;">
{draft_html}
<br><br>
<table style="font-family:Arial,sans-serif;font-size:12px;color:#333;">
<tr><td><strong>Jennifer</strong> | Technical Support</td></tr>
<tr><td>Uni-ram Corporation</td></tr>
<tr><td>381 Bentley Street, Markham, Ontario L3R 9T2, Canada</td></tr>
<tr><td>Tel: 905-477-5911 | <a href="http://www.uniram.com">www.uniram.com</a></td></tr>
</table>
</div>"""

    reply_subject = subject if subject.startswith("Re:") else f"Re: {subject}"

    send_email(
        from_mailbox=SUPPORT_MAILBOX,
        to_addresses=[customer_email],
        cc_addresses=[ESCALATION_EMAIL],
        subject=reply_subject,
        body_html=body_html
    )


def _send_ken_answer(pending: dict, correct_answer: str):
    """Send Ken's corrected answer to the customer, CC Ken."""
    func_dir = os.path.dirname(__file__)
    if func_dir not in sys.path:
        sys.path.insert(0, func_dir)

    from engine.ai_reply_engine import send_email, SUPPORT_MAILBOX, ESCALATION_EMAIL

    customer_email = pending["sender_email"]
    subject        = pending["subject"]

    answer_html = correct_answer.replace("\n", "<br>")
    body_html = f"""<div style="font-family:Arial,sans-serif;font-size:14px;color:#333;">
{answer_html}
<br><br>
<table style="font-family:Arial,sans-serif;font-size:12px;color:#333;">
<tr><td><strong>Jennifer</strong> | Technical Support</td></tr>
<tr><td>Uni-ram Corporation</td></tr>
<tr><td>381 Bentley Street, Markham, Ontario L3R 9T2, Canada</td></tr>
<tr><td>Tel: 905-477-5911 | <a href="http://www.uniram.com">www.uniram.com</a></td></tr>
</table>
</div>"""

    reply_subject = subject if subject.startswith("Re:") else f"Re: {subject}"

    send_email(
        from_mailbox=SUPPORT_MAILBOX,
        to_addresses=[customer_email],
        cc_addresses=[ESCALATION_EMAIL],
        subject=reply_subject,
        body_html=body_html
    )


def _learn_from_rejection(pending: dict, correct_answer: str):
    """Save Ken's correct answer to ChromaDB KB immediately."""
    try:
        _download_knowledge_base()

        func_dir = os.path.dirname(__file__)
        if func_dir not in sys.path:
            sys.path.insert(0, func_dir)

        from engine.ai_reply_engine import auto_learn_from_reply

        question = pending.get("subject", "") + "\n" + pending.get("email_body", "")[:400]
        product  = pending.get("product", "")
        category = pending.get("category", "")

        auto_learn_from_reply(
            original_question=question,
            reply_text=correct_answer,
            product=product,
            category=category,
            source="ken_correction",
            kb_path=KB_PATH
        )

        _upload_knowledge_base()
        logging.info(f"[Learn] Saved Ken's correction to KB for: {pending.get('subject', '')[:60]}")
    except Exception as e:
        logging.warning(f"[Learn] Failed to save to KB (non-fatal): {e}")


# ── HTML page helpers ─────────────────────────────────────────────────────────

def _approve_success_page(pending: dict) -> str:
    customer = pending.get("sender_name") or pending.get("sender_email", "")
    subject  = pending.get("subject", "")
    return f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8"><title>Reply Sent</title>
<style>
  body {{font-family:Arial,sans-serif;display:flex;justify-content:center;
         align-items:center;min-height:100vh;margin:0;background:#f0f4f8;}}
  .card {{background:white;border-radius:8px;padding:40px;max-width:520px;
          text-align:center;box-shadow:0 4px 20px rgba(0,0,0,0.1);}}
  h2 {{color:#28a745;margin:0 0 8px;}}
  .subject {{background:#f8f9fc;border:1px solid #e0e0e0;padding:8px 16px;
             border-radius:4px;font-size:14px;color:#333;margin:12px 0;}}
</style></head>
<body><div class="card">
  <div style="font-size:48px;margin-bottom:16px;">✅</div>
  <h2>Reply Sent</h2>
  <p>Jennifer's reply has been sent to <strong>{customer}</strong>.</p>
  <div class="subject">{subject}</div>
  <p style="font-size:12px;color:#888;margin-top:16px;">You have been CC'd on the reply.</p>
  <p style="font-size:11px;color:#aaa;">— Uniram AI Support System</p>
</div></body></html>"""


def _teach_success_page(pending: dict) -> str:
    customer = pending.get("sender_name") or pending.get("sender_email", "")
    subject  = pending.get("subject", "")
    return f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8"><title>Sent & Learned</title>
<style>
  body {{font-family:Arial,sans-serif;display:flex;justify-content:center;
         align-items:center;min-height:100vh;margin:0;background:#f0f4f8;}}
  .card {{background:white;border-radius:8px;padding:40px;max-width:520px;
          text-align:center;box-shadow:0 4px 20px rgba(0,0,0,0.1);}}
  h2 {{color:#1a1a2e;margin:0 0 8px;}}
  .badge {{display:inline-block;background:#28a745;color:white;padding:4px 12px;
           border-radius:12px;font-size:12px;margin:4px;}}
</style></head>
<body><div class="card">
  <div style="font-size:48px;margin-bottom:16px;">🎓</div>
  <h2>Sent &amp; Learned</h2>
  <p>Your answer has been sent to <strong>{customer}</strong>.</p>
  <div class="subject" style="background:#f8f9fc;border:1px solid #e0e0e0;padding:8px 16px;
       border-radius:4px;font-size:14px;color:#333;margin:12px 0;">{subject}</div>
  <span class="badge">✓ Reply sent</span>
  <span class="badge">✓ Knowledge base updated</span>
  <p style="font-size:12px;color:#888;margin-top:16px;">
    Jennifer has learned from your correction and will apply it to similar questions in the future.
  </p>
  <p style="font-size:11px;color:#aaa;">— Uniram AI Support System</p>
</div></body></html>"""


def _reject_form_page(pending: dict, token: str) -> str:
    customer      = pending.get("sender_name") or pending.get("sender_email", "")
    subject       = pending.get("subject", "")
    draft         = pending.get("draft_reply", "").replace("<", "&lt;").replace(">", "&gt;")
    customer_msg  = pending.get("email_body", "")[:600].replace("<", "&lt;").replace(">", "&gt;")
    key_param     = f"&code={os.environ.get('FUNCTION_KEY', '')}" if os.environ.get('FUNCTION_KEY') else ""
    action_url    = f"/api/approve_support_reply?token={token}&action=reject{key_param}"

    return f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8"><title>Teach Jennifer</title>
<style>
  body {{font-family:Arial,sans-serif;margin:0;background:#f0f4f8;padding:24px;}}
  .card {{background:white;border-radius:8px;padding:32px;max-width:700px;
          margin:0 auto;box-shadow:0 4px 20px rgba(0,0,0,0.1);}}
  h2 {{color:#1a1a2e;margin:0 0 4px;}}
  .label {{font-weight:bold;font-size:13px;color:#555;margin:16px 0 4px;}}
  .box {{background:#f8f9fc;border:1px solid #e0e0e0;border-radius:4px;
         padding:12px;font-size:13px;color:#444;white-space:pre-wrap;}}
  .draft-box {{background:#fffdf0;border:1px solid #FFD700;border-radius:4px;
               padding:12px;font-size:13px;color:#444;white-space:pre-wrap;}}
  textarea {{width:100%;min-height:180px;padding:12px;font-size:14px;
             border:2px solid #1a1a2e;border-radius:6px;box-sizing:border-box;
             font-family:Arial,sans-serif;resize:vertical;}}
  .btn {{display:inline-block;background:#1a1a2e;color:#fff;padding:12px 32px;
         border-radius:6px;border:none;font-size:15px;font-weight:bold;
         cursor:pointer;margin-top:16px;width:100%;}}
  .btn:hover {{background:#2d2d4e;}}
</style></head>
<body><div class="card">
  <h2>✏️ Teach Jennifer</h2>
  <p style="color:#555;margin:4px 0 16px;">
    Correct Jennifer's draft and send your answer to <strong>{customer}</strong>.
    Jennifer will learn from your correction.
  </p>

  <div class="label">Subject</div>
  <div class="box">{subject}</div>

  <div class="label">Customer's message</div>
  <div class="box">{customer_msg}</div>

  <div class="label">Jennifer's draft (for reference)</div>
  <div class="draft-box">{draft}</div>

  <form method="POST" action="{action_url}">
    <input type="hidden" name="token" value="{token}">
    <input type="hidden" name="action" value="reject">
    <div class="label">Your correct answer (will be sent to customer &amp; saved to KB)</div>
    <textarea name="correct_answer" placeholder="Type the correct answer here..." required></textarea>
    <button type="submit" class="btn">📤 Send &amp; Teach Jennifer</button>
  </form>

  <p style="font-size:11px;color:#aaa;margin-top:20px;text-align:center;">
    — Uniram AI Support System
  </p>
</div></body></html>"""


def _already_done_page(pending: dict) -> str:
    status = pending.get("status", "")
    resolved = pending.get("resolved_at", "")[:16].replace("T", " ")
    label = "Approved ✅" if status == "approved" else "Rejected & Taught 🎓"
    return f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8"><title>Already Processed</title>
<style>
  body {{font-family:Arial,sans-serif;display:flex;justify-content:center;
         align-items:center;min-height:100vh;margin:0;background:#f0f4f8;}}
  .card {{background:white;border-radius:8px;padding:40px;max-width:480px;
          text-align:center;box-shadow:0 4px 20px rgba(0,0,0,0.1);}}
</style></head>
<body><div class="card">
  <div style="font-size:48px;margin-bottom:16px;">ℹ️</div>
  <h2 style="color:#555;">Already Processed</h2>
  <p>This approval was already handled: <strong>{label}</strong></p>
  <p style="font-size:13px;color:#888;">Resolved at: {resolved} UTC</p>
  <p style="font-size:11px;color:#aaa;margin-top:16px;">— Uniram AI Support System</p>
</div></body></html>"""


def _error_page(msg: str) -> str:
    return f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8"><title>Error</title>
<style>
  body {{font-family:Arial,sans-serif;display:flex;justify-content:center;
         align-items:center;min-height:100vh;margin:0;background:#f0f4f8;}}
  .card {{background:white;border-radius:8px;padding:40px;max-width:480px;
          text-align:center;box-shadow:0 4px 20px rgba(0,0,0,0.1);}}
  .err {{background:#fff5f5;border:1px solid #f5c6cb;padding:12px;border-radius:4px;
         font-size:12px;color:#721c24;margin-top:12px;text-align:left;word-break:break-all;}}
</style></head>
<body><div class="card">
  <div style="font-size:48px;margin-bottom:16px;">❌</div>
  <h2 style="color:#dc3545;">Error</h2>
  <p>Something went wrong.</p>
  <div class="err">{msg}</div>
  <p style="font-size:12px;color:#aaa;margin-top:16px;">
    Contact ryan.hu@uniram.com if this persists.
  </p>
</div></body></html>"""


# ── KB helpers ────────────────────────────────────────────────────────────────

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
        logging.info("Downloading KB via AzureWebJobsStorage connection string...")
        from azure.storage.blob import BlobServiceClient
        client = BlobServiceClient.from_connection_string(conn_str)
        blob = client.get_blob_client(container=KB_CONTAINER, blob=KB_BLOB_NAME)
        with open(KB_ZIP_PATH, "wb") as f:
            data = blob.download_blob()
            data.readinto(f)
    elif sas_url:
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

    updated_zip = "/tmp/knowledge_base_updated.zip"
    with zipfile.ZipFile(updated_zip, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(KB_PATH):
            for file in files:
                full_path = os.path.join(root, file)
                arcname = os.path.relpath(full_path, KB_PATH)
                zf.write(full_path, arcname)

    from azure.storage.blob import BlobServiceClient
    client = BlobServiceClient.from_connection_string(conn_str)
    blob = client.get_blob_client(container=KB_CONTAINER, blob=KB_BLOB_NAME)
    with open(updated_zip, "rb") as f:
        blob.upload_blob(f, overwrite=True)

    zip_size = os.path.getsize(updated_zip) / 1024
    logging.info(f"KB uploaded back to Blob Storage ({zip_size:.1f} KB)")
    os.remove(updated_zip)
