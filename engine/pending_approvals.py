"""
Pending Approvals Store
Stores draft replies awaiting Ken's approval in Azure Blob Storage.
Each pending item contains:
  - original email data
  - Jennifer's draft reply
  - classification metadata
  - timestamp
"""

import json
import uuid
import os
from datetime import datetime, timezone

BLOB_CONTAINER  = "uniram-support-kb"
BLOB_PENDING    = "pending_approvals.json"


def _get_blob_client():
    from azure.storage.blob import BlobServiceClient
    conn_str = os.environ.get("AzureWebJobsStorage", "")
    if not conn_str:
        raise ValueError("AzureWebJobsStorage not set")
    client = BlobServiceClient.from_connection_string(conn_str)
    return client.get_blob_client(container=BLOB_CONTAINER, blob=BLOB_PENDING)


def _load() -> dict:
    try:
        blob = _get_blob_client()
        data = blob.download_blob().readall()
        return json.loads(data)
    except Exception:
        return {}


def _save(data: dict):
    blob = _get_blob_client()
    blob.upload_blob(json.dumps(data, ensure_ascii=False, indent=2), overwrite=True)


def save_pending(email: dict, draft_reply: str, classification: dict,
                 confidence: float) -> str:
    """Save a pending approval and return the approval token (UUID)."""
    token = str(uuid.uuid4()).replace("-", "")[:16]
    data = _load()
    data[token] = {
        "token": token,
        "created_at": datetime.now(timezone.utc).isoformat(),
        "email_id": email.get("id", ""),
        "subject": email.get("subject", ""),
        "sender_email": email.get("from", {}).get("emailAddress", {}).get("address", ""),
        "sender_name": email.get("from", {}).get("emailAddress", {}).get("name", ""),
        "email_body": email.get("body", {}).get("content", email.get("bodyPreview", "")),
        "draft_reply": draft_reply,
        "category": classification.get("category", ""),
        "product": classification.get("product_model", ""),
        "language": classification.get("language", "en"),
        "confidence": confidence,
        "status": "pending",  # pending | approved | rejected
    }
    _save(data)
    return token


def get_pending(token: str) -> dict:
    """Retrieve a pending approval by token. Returns None if not found."""
    data = _load()
    return data.get(token)


def mark_done(token: str, status: str):
    """Mark a pending approval as approved or rejected."""
    data = _load()
    if token in data:
        data[token]["status"] = status
        data[token]["resolved_at"] = datetime.now(timezone.utc).isoformat()
        _save(data)


def cleanup_old(days: int = 7):
    """Remove entries older than N days."""
    from datetime import timedelta
    data = _load()
    cutoff = datetime.now(timezone.utc) - timedelta(days=days)
    cleaned = {
        k: v for k, v in data.items()
        if datetime.fromisoformat(v["created_at"]) > cutoff
    }
    if len(cleaned) < len(data):
        _save(cleaned)
