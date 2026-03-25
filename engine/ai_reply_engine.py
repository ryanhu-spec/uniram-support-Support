"""
Step 3: AI Auto-Reply Engine
- Monitors support@uniram.com Inbox for all emails (read and unread)
- Moves processed emails to Processed subfolder under Inbox
- Classifies each email (technical vs non-technical)
- Retrieves relevant context from ChromaDB knowledge base
- Generates a professional reply using GPT-4o
- Sends reply via Microsoft Graph API (reply-to-thread with original email quoted)
- Escalates to Ken if confidence is low or question is too complex
- Auto-learns: saves new Ken/Finn replies back to knowledge base
- v19: Confidence threshold 0.65→0.45; smarter skip for internally-handled threads

Changelog:
- v2: Signature updated to "Technical Support" (single name Jennifer)
- v2: Reply now quotes original email in thread
- v2: HTML signature with Uni-ram logo (inline base64)
- v2: Fixed process_emails() signature to accept kb_path param
"""

import sys, os, json, re, time, argparse, requests, base64
from datetime import datetime

from engine.settings import (
    GRAPH_TENANT_ID, GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET,
    OPENAI_API_KEY, SUPPORT_MAILBOX
)
from openai import OpenAI
import chromadb
from chromadb import Documents, EmbeddingFunction, Embeddings

openai_client = OpenAI(api_key=OPENAI_API_KEY, base_url="https://api.openai.com/v1")

# ─────────────────────────────────────────────
# Config
# ─────────────────────────────────────────────
ESCALATION_EMAIL     = "ken@uniram.com"
ESCALATION_CC        = "finn.sun@uniram.com"
CONFIDENCE_THRESHOLD = 0.45  # Lowered from 0.65 — Jennifer should attempt to answer more
DEFAULT_KB_PATH      = "/tmp/uniram_kb"
LOG_PATH             = "/tmp/reply_log.jsonl"

# ─────────────────────────────────────────────
# Uni-ram logo (inline base64 — loaded once at startup)
# ─────────────────────────────────────────────
def _load_logo_b64() -> str:
    """Load Uni-ram logo as base64 from assets folder next to this file."""
    # Try several possible paths
    candidates = [
        os.path.join(os.path.dirname(__file__), "..", "assets", "image002.png"),
        "/home/ubuntu/jennifer_project/assets/image002.png",
    ]
    for path in candidates:
        if os.path.exists(path):
            with open(path, "rb") as f:
                return base64.b64encode(f.read()).decode()
    return ""  # fallback: no logo

LOGO_B64 = _load_logo_b64()

# ─────────────────────────────────────────────
# HTML Signature (Technical Support)
# ─────────────────────────────────────────────
def build_signature_html() -> str:
    logo_tag = (
        f'<img src="data:image/png;base64,{LOGO_B64}" width="120" alt="Uni-ram" style="display:block;">'
        if LOGO_B64 else ""
    )
    return f"""
<table cellpadding="0" cellspacing="0" border="0"
       style="font-family:Arial,sans-serif;font-size:13px;color:#333;margin-top:20px;">
  <tr>
    <td style="padding-right:15px;vertical-align:top;">{logo_tag}</td>
    <td style="border-left:3px solid #FFD700;padding-left:15px;vertical-align:top;line-height:1.7;">
      <strong>Jennifer</strong><br>
      Technical Support<br>
      Uni-ram Corporation<br>
      381 Bentley Street, Markham,<br>
      Ontario L3R 9T2, Canada<br>
      Tel: 905-477-5911 &nbsp;|&nbsp; Toll-Free: 1-800-417-9133<br>
      <a href="http://www.uniram.com" style="color:#0563C1;">www.uniram.com</a>
      &nbsp;&nbsp;
      <a href="mailto:support@uniram.com" style="color:#0563C1;">support@uniram.com</a>
      <br><br>
      <em>Discover why millions of users rely on Uni-ram products</em><br>
      Your opinion matters to us! Please share your experience with us by clicking
      <a href="https://www.uniram.com/feedback" style="color:#0563C1;"><strong>HERE</strong></a>.
      We value your feedback!
    </td>
  </tr>
</table>"""

# ─────────────────────────────────────────────
# Embedding function (direct OpenAI)
# ─────────────────────────────────────────────
class UniramEmbeddingFunction(EmbeddingFunction):
    def __init__(self):
        self.client = OpenAI(api_key=OPENAI_API_KEY, base_url="https://api.openai.com/v1")
    def __call__(self, input: Documents) -> Embeddings:
        response = self.client.embeddings.create(model="text-embedding-3-small", input=input)
        return [item.embedding for item in response.data]

def get_collection(kb_path: str = DEFAULT_KB_PATH):
    client = chromadb.PersistentClient(path=kb_path)
    return client.get_or_create_collection(
        name="uniram_support",
        embedding_function=UniramEmbeddingFunction(),
        metadata={"hnsw:space": "cosine"}
    )

# ─────────────────────────────────────────────
# Graph API helpers
# ─────────────────────────────────────────────
_token_cache = {"token": None, "expires": 0}

def get_graph_token():
    if _token_cache["token"] and time.time() < _token_cache["expires"] - 60:
        return _token_cache["token"]
    r = requests.post(
        f"https://login.microsoftonline.com/{GRAPH_TENANT_ID}/oauth2/v2.0/token",
        data={"grant_type": "client_credentials", "client_id": GRAPH_CLIENT_ID,
              "client_secret": GRAPH_CLIENT_SECRET, "scope": "https://graph.microsoft.com/.default"}
    )
    r.raise_for_status()
    data = r.json()
    _token_cache["token"] = data["access_token"]
    _token_cache["expires"] = time.time() + data.get("expires_in", 3600)
    return _token_cache["token"]

def get_inbox_emails(mailbox: str, max_count: int = 50):
    """Get all emails from Inbox (not Processed subfolder) regardless of read status."""
    token = get_graph_token()
    # Fetch from Inbox directly — Processed is a subfolder so won't appear here
    r = requests.get(
        f"https://graph.microsoft.com/v1.0/users/{mailbox}/mailFolders/Inbox/messages",
        headers={"Authorization": f"Bearer {token}"},
        params={
            "$top": max_count,
            "$select": "id,subject,from,toRecipients,ccRecipients,receivedDateTime,body,bodyPreview,conversationId",
            "$orderby": "receivedDateTime asc"
        }
    )
    return r.json().get("value", []) if r.status_code == 200 else []

def get_unread_emails(mailbox: str, max_count: int = 50):
    """Alias for backward compatibility — now returns all inbox emails."""
    return get_inbox_emails(mailbox, max_count)

def mark_as_read(mailbox: str, message_id: str):
    token = get_graph_token()
    requests.patch(
        f"https://graph.microsoft.com/v1.0/users/{mailbox}/messages/{message_id}",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={"isRead": True}
    )

def move_to_processed(mailbox: str, message_id: str):
    """Move a processed email to the 'Processed' subfolder under Inbox."""
    token = get_graph_token()

    # Get or create the Processed folder under Inbox
    # First, find the Inbox folder ID
    r = requests.get(
        f"https://graph.microsoft.com/v1.0/users/{mailbox}/mailFolders/Inbox/childFolders",
        headers={"Authorization": f"Bearer {token}"},
        params={"$filter": "displayName eq 'Processed'", "$top": 1}
    )
    folders = r.json().get("value", []) if r.status_code == 200 else []

    if folders:
        processed_folder_id = folders[0]["id"]
    else:
        # Create the Processed folder
        cr = requests.post(
            f"https://graph.microsoft.com/v1.0/users/{mailbox}/mailFolders/Inbox/childFolders",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={"displayName": "Processed"}
        )
        if cr.status_code in (200, 201):
            processed_folder_id = cr.json()["id"]
        else:
            print(f"    [move] Could not create Processed folder: {cr.status_code}")
            return

    # Move the message
    mv = requests.post(
        f"https://graph.microsoft.com/v1.0/users/{mailbox}/messages/{message_id}/move",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={"destinationId": processed_folder_id}
    )
    if mv.status_code == 201:
        print(f"    Moved to Processed folder")
    else:
        print(f"    [move] Failed to move: {mv.status_code}")


def learn_from_folder_history(mailbox: str, folder_names: list,
                              kb_path: str = DEFAULT_KB_PATH,
                              max_per_folder: int = 200):
    """
    One-time historical learning: scan specified subfolders under Inbox,
    extract Q&A pairs from Ken/Finn replies and add to knowledge base.
    Folders: ['2023', '2024', '2025', 'Ken', 'Finn', 'Coop', 'Luis', 'Norbu', 'Shipping']
    """
    token = get_graph_token()
    total_learned = 0

    # Get all child folders of Inbox
    r = requests.get(
        f"https://graph.microsoft.com/v1.0/users/{mailbox}/mailFolders/Inbox/childFolders",
        headers={"Authorization": f"Bearer {token}"},
        params={"$top": 50}
    )
    if r.status_code != 200:
        print(f"  [history] Could not list folders: {r.status_code}")
        return

    all_folders = {f["displayName"]: f["id"] for f in r.json().get("value", [])}
    print(f"  [history] Found folders: {list(all_folders.keys())}")

    for folder_name in folder_names:
        folder_id = all_folders.get(folder_name)
        if not folder_id:
            print(f"  [history] Folder '{folder_name}' not found — skipping")
            continue

        print(f"  [history] Scanning folder: {folder_name}")
        learned_in_folder = 0
        next_link = (
            f"https://graph.microsoft.com/v1.0/users/{mailbox}/mailFolders/{folder_id}/messages"
            f"?$top=50&$select=id,subject,from,body,bodyPreview,receivedDateTime,conversationId"
            f"&$orderby=receivedDateTime desc"
        )

        fetched = 0
        while next_link and fetched < max_per_folder:
            mr = requests.get(next_link, headers={"Authorization": f"Bearer {token}"}, timeout=30)
            if mr.status_code != 200:
                break
            data = mr.json()
            messages = data.get("value", [])
            fetched += len(messages)

            for msg in messages:
                subject = msg.get("subject", "")
                sender_addr = msg.get("from", {}).get("emailAddress", {}).get("address", "").lower()
                body_text = extract_text(msg)

                # Only learn from replies by Ken, Finn, or internal staff (not customers)
                is_internal = any(domain in sender_addr for domain in ["@uniram.com"])
                if not is_internal:
                    continue

                # Skip very short replies (auto-acks, etc.)
                if len(body_text.strip()) < 50:
                    continue

                # Extract the technical content (before quoted original)
                reply_body = re.split(
                    r'(From:|-----Original|________________________________|On .* wrote:)',
                    body_text
                )[0].strip()

                if len(reply_body) < 30:
                    continue

                # Use GPT to extract Q&A pair
                try:
                    from openai import OpenAI
                    client = OpenAI(api_key=OPENAI_API_KEY, base_url="https://api.openai.com/v1")
                    extract_resp = client.chat.completions.create(
                        model="gpt-4o-mini",
                        messages=[{
                            "role": "system",
                            "content": (
                                "You are extracting technical Q&A pairs from support email threads. "
                                "Given a reply from a Uniram engineer, extract: "
                                "1) The implied customer question/problem "
                                "2) The technical answer/solution provided. "
                                "Return JSON: {\"question\": \"...\", \"answer\": \"...\", \"product\": \"URS500|URS900|URS600|unknown\", \"is_technical\": true/false}. "
                                "If not technical, return {\"is_technical\": false}."
                            )
                        }, {
                            "role": "user",
                            "content": f"Subject: {subject}\n\nReply:\n{reply_body[:800]}"
                        }],
                        response_format={"type": "json_object"},
                        max_tokens=300,
                        temperature=0
                    )
                    result = json.loads(extract_resp.choices[0].message.content)
                    if not result.get("is_technical", False):
                        continue
                    question = result.get("question", subject)
                    answer = result.get("answer", reply_body)
                    product = result.get("product", "")

                    auto_learn_from_reply(
                        original_question=question,
                        reply_text=answer,
                        product=product,
                        category="equipment_fault",
                        source=f"history_{folder_name}",
                        kb_path=kb_path
                    )
                    learned_in_folder += 1
                    total_learned += 1

                except Exception as e:
                    continue  # Skip on any GPT error

            next_link = data.get("@odata.nextLink")

        print(f"  [history] {folder_name}: learned {learned_in_folder} items")

    print(f"  [history] Total learned from history: {total_learned} items")
    return total_learned


def build_reply_html(reply_text: str, original_email: dict) -> str:
    """Build full HTML reply: body + signature + quoted original email."""
    # Convert plain text reply to HTML paragraphs
    paragraphs = reply_text.strip().split("\n")
    body_html = "".join(
        f"<p>{p}</p>" if p.strip() else "<br>"
        for p in paragraphs
    )

    # Quoted original email
    orig_from = original_email.get("from", {}).get("emailAddress", {}).get("address", "")
    orig_date = original_email.get("receivedDateTime", "")[:10]
    orig_subject = original_email.get("subject", "")
    orig_body = original_email.get("body", {}).get("content", "")

    quoted = f"""
<br>
<hr style="border:none;border-top:1px solid #ccc;margin:12px 0;">
<div style="color:#666;font-size:12px;">
  <strong>From:</strong> {orig_from}<br>
  <strong>Sent:</strong> {orig_date}<br>
  <strong>To:</strong> {SUPPORT_MAILBOX}<br>
  <strong>Subject:</strong> {orig_subject}<br>
  <br>
  {orig_body}
</div>"""

    return f"<html><body>{body_html}{build_signature_html()}{quoted}</body></html>"

def send_reply(mailbox: str, message_id: str, reply_html: str) -> bool:
    """Reply to an existing email thread using Graph API /reply endpoint."""
    token = get_graph_token()
    r = requests.post(
        f"https://graph.microsoft.com/v1.0/users/{mailbox}/messages/{message_id}/reply",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={
            "message": {
                "body": {"contentType": "HTML", "content": reply_html}
            },
            "comment": ""
        }
    )
    return r.status_code == 202

def forward_email(mailbox: str, message_id: str, to_addresses: list,
                  comment: str = "") -> bool:
    """Forward an existing email using Graph API /forward — preserves all attachments and inline images."""
    token = get_graph_token()
    r = requests.post(
        f"https://graph.microsoft.com/v1.0/users/{mailbox}/messages/{message_id}/forward",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={
            "comment": comment,
            "toRecipients": [{"emailAddress": {"address": a}} for a in to_addresses]
        }
    )
    return r.status_code == 202

def send_email(from_mailbox: str, to_addresses: list, subject: str,
               body_html: str, cc_addresses: list = None) -> bool:
    """Send a new email (used for escalation to Ken)."""
    token = get_graph_token()
    message = {
        "subject": subject,
        "body": {"contentType": "HTML", "content": body_html},
        "toRecipients": [{"emailAddress": {"address": a}} for a in to_addresses],
    }
    if cc_addresses:
        message["ccRecipients"] = [{"emailAddress": {"address": a}} for a in cc_addresses]
    r = requests.post(
        f"https://graph.microsoft.com/v1.0/users/{from_mailbox}/sendMail",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={"message": message, "saveToSentItems": True}
    )
    return r.status_code in (200, 201, 202)

# ─────────────────────────────────────────────
# Email parsing
# ─────────────────────────────────────────────
def extract_text(email: dict) -> str:
    body_html = email.get("body", {}).get("content", email.get("bodyPreview", ""))
    text = re.sub(r'<[^>]+>', ' ', body_html)
    return re.sub(r'\s+', ' ', text).strip()

def get_sender_address(email: dict) -> str:
    return email.get("from", {}).get("emailAddress", {}).get("address", "")

def get_sender_name(email: dict) -> str:
    return email.get("from", {}).get("emailAddress", {}).get("name", "")

# ─────────────────────────────────────────────
# Classification
# ─────────────────────────────────────────────
CLASSIFY_PROMPT = """You are analyzing incoming emails to Uniram Corporation's support mailbox.
Uniram manufactures automotive service equipment: spray gun washers, solvent recyclers, tire changers, wheel balancers, vehicle lifts.

Classify this email and return JSON:
- "is_technical": true/false — is this a genuine technical support or product question?
- "category": "equipment_fault" | "parts_inquiry" | "installation" | "operation" | "maintenance" | "warranty" | "solvent_safety" | "pricing_inquiry" | "general_inquiry" | "not_technical"
  NOTE: Use "solvent_safety" if the question involves solvent type selection, compatible solvents, temperature settings for specific solvents, flash point, boiling point, or chemical compatibility.
  NOTE: Use "pricing_inquiry" ONLY if the customer explicitly asks about price, quote, cost, discount, bulk pricing, or requests a formal quotation for a product. Examples: "how much does it cost", "can you send me a quote", "what is the price for X".
  NOTE: Do NOT use "pricing_inquiry" for emails asking how to get support, how to fix a problem, or reporting that equipment is not working — even if they mention a product model. Those should be "equipment_fault" or "general_inquiry".
  NOTE: Set "is_technical" to true for "pricing_inquiry" so it gets routed properly.
- "core_question": the specific question being asked (null if not_technical)
- "product_model": product model mentioned (e.g. "URS500", "UGW-110") or null
- "urgency": "high" | "normal" | "low"
- "language": detected language code (e.g. "en", "fr", "es", "ja")
- "is_vague": true/false — set to true if the email is too vague to diagnose (e.g. just says "it doesn't work" or "not working" without describing any specific symptom, error, behavior, or what they already tried). Set to false if the customer has provided enough detail to attempt a diagnosis."""

def classify_email(email: dict) -> dict:
    subject = email.get("subject", "")
    body_text = extract_text(email)
    sender = get_sender_address(email)
    prompt = f"Subject: {subject}\nFrom: {sender}\nBody:\n{body_text[:2000]}"
    resp = openai_client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[{"role": "system", "content": CLASSIFY_PROMPT},
                  {"role": "user", "content": prompt}],
        response_format={"type": "json_object"}, temperature=0
    )
    return json.loads(resp.choices[0].message.content)

# ─────────────────────────────────────────────
# Safety-critical keyword check
# ─────────────────────────────────────────────
SAFETY_KEYWORDS = [
    "flash point", "flashpoint", "boiling point", "what temperature", "which solvent",
    "compatible solvent", "solvent type", "what solvent", "which chemical",
    "can i use", "is it safe to use", "acetone", "lacquer thinner", "methanol",
    "ethanol", "isopropyl", "toluene", "xylene", "naphtha", "mineral spirits",
    "temperature setting", "set temperature", "boil", "ignite", "flammable",
    "explosion", "fire hazard", "chemical", "vapor"
]

def is_safety_critical(text: str, category: str) -> bool:
    if category == "solvent_safety":
        return True
    text_lower = text.lower()
    matched = [kw for kw in SAFETY_KEYWORDS if kw in text_lower]
    return len(matched) >= 2

# ─────────────────────────────────────────────
# Knowledge base retrieval
# ─────────────────────────────────────────────
def retrieve_context(question: str, product: str = None,
                     n_results: int = 5, kb_path: str = DEFAULT_KB_PATH) -> list:
    collection = get_collection(kb_path)
    query = f"{product}: {question}" if product else question
    results = collection.query(query_texts=[query], n_results=n_results)
    contexts = []
    if results and results.get("documents"):
        for doc, meta in zip(results["documents"][0], results["metadatas"][0]):
            contexts.append({
                "text": doc,
                "source": meta.get("source", ""),
                "type": meta.get("type", ""),
            })
    return contexts

# ─────────────────────────────────────────────
# Reply generation
# ─────────────────────────────────────────────
REPLY_SYSTEM_PROMPT = """You are Jennifer, Technical Support at Uni-ram Corporation, a manufacturer of professional automotive service equipment (spray gun washers, solvent recyclers, tire changers, wheel balancers, vehicle lifts).

Your job is to write professional, helpful email replies to customer technical support questions. You write as a real person, not a robot.

Guidelines:
- Be direct and to the point. No unnecessary pleasantries like "I hope this email finds you well."
- Don't start with "Dear". Use the customer's first name if available, otherwise just start with the content.
- Give clear, actionable answers based on the provided context.
- If referencing the manual, mention the specific section or page number.
- If you cannot fully answer, say so honestly and let them know one of our engineers will follow up shortly.
- Keep the tone professional but warm and approachable — like a knowledgeable colleague, not a robot.
- Reply in the same language as the customer's email.
- Do NOT include a signature in your reply — it will be added automatically.
- End your reply naturally without any signature block."""

def generate_reply(customer_name: str, question: str, contexts: list,
                   category: str, product: str, language: str = "en") -> tuple:
    """Returns (reply_text, confidence_score, needs_escalation, escalation_reason)"""
    context_text = "\n\n".join([
        f"[Source: {c['source']}]\n{c['text']}" for c in contexts
    ]) if contexts else "No relevant context found in knowledge base."

    user_prompt = f"""Customer name: {customer_name or 'Customer'}
Product: {product or 'Not specified'}
Category: {category}
Customer question: {question}

Relevant knowledge base context:
{context_text}

Write a helpful reply email body (no signature). Also provide a confidence score (0.0-1.0).
Return as JSON: {{"reply": "...", "confidence": 0.0-1.0, "needs_escalation": true/false, "escalation_reason": "..." }}"""

    resp = openai_client.chat.completions.create(
        model="gpt-4o",
        messages=[{"role": "system", "content": REPLY_SYSTEM_PROMPT},
                  {"role": "user", "content": user_prompt}],
        response_format={"type": "json_object"}, temperature=0.3
    )
    result = json.loads(resp.choices[0].message.content)
    return (result.get("reply", ""), result.get("confidence", 0.5),
            result.get("needs_escalation", False), result.get("escalation_reason", ""))

# ─────────────────────────────────────────────
# Escalation to Ken
# ─────────────────────────────────────────────
def escalate_to_ken(email: dict, classification: dict, ai_draft: str,
                    escalation_reason: str, dry_run: bool = False) -> bool:
    subject = email.get("subject", "")
    sender_name = get_sender_name(email)
    sender_addr = get_sender_address(email)
    body_text = extract_text(email)[:800]
    received = email.get("receivedDateTime", "")[:10]

    escalation_body = f"""<p>Hi Ken,</p>
<p>An incoming support email needs your attention. The AI assistant couldn't resolve this one with sufficient confidence.</p>
<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:14px;">
  <tr><td><b>From</b></td><td>{sender_name} &lt;{sender_addr}&gt;</td></tr>
  <tr><td><b>Subject</b></td><td>{subject}</td></tr>
  <tr><td><b>Received</b></td><td>{received}</td></tr>
  <tr><td><b>Category</b></td><td>{classification.get('category', '')}</td></tr>
  <tr><td><b>Product</b></td><td>{classification.get('product_model') or 'Not specified'}</td></tr>
  <tr><td><b>Reason for escalation</b></td><td>{escalation_reason}</td></tr>
</table>
<br>
<p><b>Customer's message:</b></p>
<blockquote style="border-left:3px solid #ccc;padding-left:12px;color:#555;">{body_text}</blockquote>
<br>
<p><b>AI draft reply (for reference):</b></p>
<blockquote style="border-left:3px solid #f90;padding-left:12px;color:#555;">{ai_draft.replace(chr(10), '<br>')}</blockquote>
<br>
<p>Please reply directly to the customer at <a href="mailto:{sender_addr}">{sender_addr}</a>.</p>
<p style="color:#888;font-size:12px;">— Uniram AI Support System</p>"""

    escalation_subject = f"[Support Escalation] {subject}"

    if dry_run:
        print(f"\n  [DRY RUN] Would escalate to {ESCALATION_EMAIL}")
        print(f"  Subject: {escalation_subject}")
        print(f"  Reason: {escalation_reason}")
        return True

    return send_email(
        from_mailbox=SUPPORT_MAILBOX,
        to_addresses=[ESCALATION_EMAIL],
        cc_addresses=[ESCALATION_CC],
        subject=escalation_subject,
        body_html=escalation_body
    )

# ─────────────────────────────────────────────
# Auto-learn: save Ken/Finn replies to KB
# ─────────────────────────────────────────────
def auto_learn_from_reply(original_question: str, reply_text: str,
                          product: str, category: str, source: str,
                          kb_path: str = DEFAULT_KB_PATH):
    collection = get_collection(kb_path)
    doc_text = f"Question: {original_question}\nAnswer: {reply_text}"
    doc_id = f"autolearn_{datetime.now().strftime('%Y%m%d%H%M%S%f')}"
    collection.upsert(
        ids=[doc_id],
        documents=[doc_text],
        metadatas=[{"source": f"autolearn:{source}", "type": "email_qa",
                    "category": category, "product": product or ""}]
    )

# ─────────────────────────────────────────────
# Learn from Ken's feedback on escalation emails
# ─────────────────────────────────────────────
KEN_MAILBOX = "ken@uniram.com"
KEN_FEEDBACK_TAG = "Jennifer-Feedback-Processed"

KEN_INTENT_PROMPT = """You are analyzing a reply from Ken Wu (Uniram's lead engineer) to an escalated support email.
Ken's reply is short and informal. Extract his intent and any technical answer he may have provided.

Return JSON with:
- "intent": one of "junk" | "handled" | "sales" | "unknown"
  - "junk": Ken says it's spam, marketing, not relevant, wrong inbox, cold call, solicitation
  - "handled": Ken says he replied, took care of it, resolved it, or provides a technical answer
  - "sales": Ken says it's a sales lead, new customer inquiry, should go to Kate/Luna/sales
  - "unknown": cannot determine intent
- "technical_answer": if Ken provided a technical answer or diagnosis, extract it here (string or null)
- "confidence": 0.0-1.0 how confident you are in the intent classification"""

def parse_ken_intent_gpt(ken_reply: str, original_subject: str = "") -> dict:
    """Use GPT to understand Ken's intent from his reply. Returns dict with intent, technical_answer, confidence."""
    try:
        prompt = f"Original escalation subject: {original_subject}\n\nKen's reply:\n{ken_reply[:800]}"
        resp = openai_client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": KEN_INTENT_PROMPT},
                {"role": "user", "content": prompt}
            ],
            response_format={"type": "json_object"},
            temperature=0
        )
        return json.loads(resp.choices[0].message.content)
    except Exception as e:
        print(f"  [learn] GPT intent parse failed: {e}")
        return {"intent": "unknown", "technical_answer": None, "confidence": 0.0}

def learn_from_ken_feedback(kb_path: str = DEFAULT_KB_PATH, dry_run: bool = False):
    """Scan Ken's inbox for replies to [Support Escalation] emails and learn from them."""
    token = get_graph_token()

    # Fetch recent unread emails in Ken's inbox that are replies to escalation threads
    r = requests.get(
        f"https://graph.microsoft.com/v1.0/users/{KEN_MAILBOX}/messages",
        headers={"Authorization": f"Bearer {token}"},
        params={
            "$filter": "isRead eq false",
            "$top": 20,
            "$select": "id,subject,from,body,bodyPreview,conversationId,receivedDateTime,categories",
            "$orderby": "receivedDateTime asc"
        }
    )
    if r.status_code != 200:
        print(f"  [learn] Could not fetch Ken's inbox: {r.status_code}")
        return

    ken_emails = r.json().get("value", [])
    learned = 0

    for email in ken_emails:
        subject = email.get("subject", "")
        sender_addr = email.get("from", {}).get("emailAddress", {}).get("address", "").lower()
        categories = email.get("categories", [])

        # Only process Ken's own replies to escalation threads
        # Ken's reply will have subject like "Re: [Support Escalation] ..."
        if "support escalation" not in subject.lower():
            continue
        if sender_addr != "ken@uniram.com":
            continue
        if KEN_FEEDBACK_TAG in categories:
            continue  # already processed

        body_text = extract_text(email)
        # Extract just Ken's top reply (before the quoted original)
        # Ken's reply is typically the first few lines before "-----Original Message-----" or "From:"
        ken_reply = re.split(r'(From:|-----Original|________________________________)', body_text)[0].strip()

        # Use GPT to understand Ken's intent
        gpt_result = parse_ken_intent_gpt(ken_reply, original_subject=subject)
        intent = gpt_result.get("intent", "unknown")
        tech_answer = gpt_result.get("technical_answer") or ken_reply
        confidence = gpt_result.get("confidence", 0.0)

        print(f"  [learn] Ken replied to: {subject[:60]}")
        print(f"  [learn] GPT intent: {intent} (confidence: {confidence:.2f}) | Reply: {ken_reply[:80]}")

        if not dry_run:
            if intent == "junk":
                auto_learn_from_reply(
                    original_question=subject,
                    reply_text="[Ken marked as junk/spam — skip similar emails from this sender]",
                    product="", category="not_technical", source="ken_feedback_junk",
                    kb_path=kb_path
                )
                print(f"  [learn] Saved junk signal to KB")

            elif intent == "handled":
                # Use GPT-extracted technical answer if available, otherwise use Ken's raw reply
                original_q_match = re.search(
                    r'Customer.s message[:\s]+([\s\S]{20,500}?)(?:AI draft|Please reply|\Z)',
                    body_text, re.IGNORECASE
                )
                original_q = original_q_match.group(1).strip() if original_q_match else subject
                auto_learn_from_reply(
                    original_question=original_q,
                    reply_text=tech_answer,
                    product="", category="equipment_fault", source="ken_feedback_handled",
                    kb_path=kb_path
                )
                print(f"  [learn] Saved Ken's answer to KB: {tech_answer[:80]}")

            elif intent == "sales":
                auto_learn_from_reply(
                    original_question=subject,
                    reply_text="[Ken flagged as sales lead — route to sales@uniram.com]",
                    product="", category="pricing_inquiry", source="ken_feedback_sales",
                    kb_path=kb_path
                )
                print(f"  [learn] Saved sales routing signal to KB")

            # Tag the email so we don't process it again
            requests.patch(
                f"https://graph.microsoft.com/v1.0/users/{KEN_MAILBOX}/messages/{email['id']}",
                headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
                json={"isRead": True, "categories": categories + [KEN_FEEDBACK_TAG]}
            )
            learned += 1
        else:
            print(f"  [DRY RUN] Would process Ken feedback: intent={intent} (GPT confidence: {confidence:.2f})")

    print(f"  [learn] Ken feedback scan complete — {learned} items learned")

# ─────────────────────────────────────────────
# Logging
# ─────────────────────────────────────────────
def log_action(email: dict, classification: dict, action: str,
               confidence: float, reply_preview: str = ""):
    entry = {
        "timestamp": datetime.now().isoformat(),
        "email_id": email.get("id", ""),
        "subject": email.get("subject", ""),
        "sender": get_sender_address(email),
        "category": classification.get("category", ""),
        "product": classification.get("product_model", ""),
        "action": action,
        "confidence": confidence,
        "reply_preview": reply_preview[:200]
    }
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(json.dumps(entry, ensure_ascii=False) + "\n")

# ─────────────────────────────────────────────
# Main processing loop
# ─────────────────────────────────────────────
# Only skip genuine auto-replies and calendar noise — everything else gets routed
NON_TECHNICAL_SKIP = [
    "out of office", "automatic reply", "auto-reply", "autoreply",
    "no-reply", "noreply",
    "calendar invite", "accepted:", "declined:", "tentative:"
]

INTERNAL_DOMAINS = ["@uniram.com"]

def is_internal_address(addr: str) -> bool:
    return any(domain in addr.lower() for domain in INTERNAL_DOMAINS)

def should_skip(email: dict) -> bool:
    subject = email.get("subject", "")
    body_preview = email.get("bodyPreview", "")
    text = (subject + " " + body_preview).lower()

    # Rule 1: Auto-reply / calendar noise
    if any(kw in text for kw in NON_TECHNICAL_SKIP):
        return True

    # Rule 2: Already handled internally
    # If email is forwarded by internal staff AND another internal staff is already CC'd,
    # someone is already handling it — skip to avoid double-escalation to Ken
    sender = email.get("from", {}).get("emailAddress", {}).get("address", "").lower()
    cc_list = [r.get("emailAddress", {}).get("address", "").lower()
               for r in email.get("ccRecipients", [])]
    to_list = [r.get("emailAddress", {}).get("address", "").lower()
               for r in email.get("toRecipients", [])]
    all_recipients = cc_list + to_list

    if is_internal_address(sender):
        # Internal staff already CC'd (excluding support@ itself) means someone is on it
        internal_cc = [a for a in all_recipients
                       if is_internal_address(a) and "support@uniram.com" not in a]
        if internal_cc:
            return True

    # Rule 3: [Support Escalation] replies bouncing back into inbox — already processed
    if "[support escalation]" in subject.lower() and is_internal_address(sender):
        return True

    return False

def process_emails(dry_run: bool = False, kb_path: str = DEFAULT_KB_PATH):
    print(f"\n{'='*60}")
    print(f"  Uniram AI Support — Auto-Reply Engine v19")
    print(f"  Mode: {'DRY RUN (no emails sent)' if dry_run else 'LIVE'}")
    print(f"  Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"{'='*60}\n")

    # Phase 0: Learn from Ken's feedback before processing new emails
    print("Phase 0: Scanning Ken's inbox for feedback on escalations...")
    try:
        learn_from_ken_feedback(kb_path=kb_path, dry_run=dry_run)
    except Exception as e:
        print(f"  [learn] Error during Ken feedback scan: {e}")
    print()

    # Phase 0.5: One-time historical learning (runs only if LEARN_HISTORY env var is set)
    # Process one folder at a time to avoid Azure Function 10-min timeout
    # Set LEARN_HISTORY=true and LEARN_HISTORY_FOLDER=<folder_name> to process a specific folder
    # When all folders are done, set LEARN_HISTORY=false
    if os.environ.get("LEARN_HISTORY", "").lower() == "true":
        ALL_HISTORY_FOLDERS = ["2023", "2024", "2025", "Ken", "Finn", "Coop", "Luis", "Norbu", "Shipping"]
        # Which folder to process this run (default: first one)
        current_folder = os.environ.get("LEARN_HISTORY_FOLDER", ALL_HISTORY_FOLDERS[0])
        print(f"Phase 0.5: Learning from historical folder: {current_folder}")
        try:
            learned = learn_from_folder_history(
                mailbox=SUPPORT_MAILBOX,
                folder_names=[current_folder],
                kb_path=kb_path,
                max_per_folder=50  # 50 per run to stay within timeout
            )
            # Advance to next folder automatically
            try:
                idx = ALL_HISTORY_FOLDERS.index(current_folder)
                if idx + 1 < len(ALL_HISTORY_FOLDERS):
                    next_folder = ALL_HISTORY_FOLDERS[idx + 1]
                    import subprocess as _sp
                    _sp.run([
                        "az", "functionapp", "config", "appsettings", "set",
                        "--name", "uniram-support",
                        "--resource-group", "uniram-reports-rg",
                        "--settings", f"LEARN_HISTORY_FOLDER={next_folder}",
                        "-o", "none"
                    ], capture_output=True, timeout=30)
                    print(f"  [history] Next run will process: {next_folder}")
                else:
                    # All folders done — disable LEARN_HISTORY
                    import subprocess as _sp
                    _sp.run([
                        "az", "functionapp", "config", "appsettings", "set",
                        "--name", "uniram-support",
                        "--resource-group", "uniram-reports-rg",
                        "--settings", "LEARN_HISTORY=false",
                        "-o", "none"
                    ], capture_output=True, timeout=30)
                    print("  [history] All folders complete — LEARN_HISTORY disabled")
            except Exception as e2:
                print(f"  [history] Could not advance folder: {e2}")
        except Exception as e:
            print(f"  [history] Error: {e}")
        print()

    emails = get_inbox_emails(SUPPORT_MAILBOX, max_count=50)
    print(f"Found {len(emails)} emails in Inbox (all, not just unread) for {SUPPORT_MAILBOX}\n")

    if not emails:
        print("  Nothing to process.")
        return

    stats = {"processed": 0, "replied": 0, "escalated": 0, "skipped": 0}

    for email in emails:
        subject = email.get("subject", "")
        sender  = get_sender_address(email)
        msg_id  = email["id"]
        print(f"─── Processing: {subject[:60]}")
        print(f"    From: {sender}")

        # Quick skip for obvious non-technical
        if should_skip(email):
            print(f"    Skipped (rule-based filter)\n")
            mark_as_read(SUPPORT_MAILBOX, msg_id)
            move_to_processed(SUPPORT_MAILBOX, msg_id)
            stats["skipped"] += 1
            log_action(email, {}, "skipped", 0)
            continue

        # Step 1: Classify
        classification = classify_email(email)
        category = classification.get("category", "")
        product  = classification.get("product_model")
        language = classification.get("language", "en")
        print(f"    Category: {category} | Product: {product or 'N/A'} | Lang: {language}")

        if not classification.get("is_technical"):
            print(f"    Not clearly technical — escalating to Ken for review\n")
            escalate_to_ken(
                email=email,
                classification=classification,
                ai_draft="[This inquiry appears to be non-technical or outside Uniram's product scope. Please review and respond if appropriate.]",
                escalation_reason="Email classified as non-technical or outside Uniram product scope",
                dry_run=dry_run
            )
            mark_as_read(SUPPORT_MAILBOX, msg_id)
            move_to_processed(SUPPORT_MAILBOX, msg_id)
            stats["escalated"] += 1
            log_action(email, classification, "escalated_non_technical", 0)
            continue

        # Pricing inquiry — forward to sales (native forward preserves images/attachments)
        if category == "pricing_inquiry":
            print(f"    Pricing inquiry — forwarding to sales@uniram.com\n")
            sender_addr = get_sender_address(email)
            product = classification.get('product_model') or 'Not specified'
            forward_comment = (
                f"Hi Sales Team,\n\n"
                f"A customer has sent a pricing inquiry to the support mailbox. Please follow up.\n"
                f"Product: {product}\n"
                f"Please reply directly to the customer at {sender_addr}.\n\n"
                f"— Uniram AI Support System"
            )
            if not dry_run:
                success = forward_email(
                    mailbox=SUPPORT_MAILBOX,
                    message_id=msg_id,
                    to_addresses=["sales@uniram.com"],
                    comment=forward_comment
                )
                if success:
                    print(f"    Forwarded (with attachments) to sales@uniram.com")
                else:
                    print(f"    Forward failed — falling back to summary email")
                    # Fallback: send summary without attachments
                    body_text_preview = extract_text(email)[:800]
                    received = email.get("receivedDateTime", "")[:10]
                    sender_name = get_sender_name(email)
                    fallback_body = f"""<p>Hi Sales Team,</p>
<p>A customer has sent a pricing inquiry to the support mailbox. Please follow up.</p>
<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:14px;">
  <tr><td><b>From</b></td><td>{sender_name} &lt;{sender_addr}&gt;</td></tr>
  <tr><td><b>Subject</b></td><td>{subject}</td></tr>
  <tr><td><b>Received</b></td><td>{received}</td></tr>
  <tr><td><b>Product</b></td><td>{product}</td></tr>
</table>
<br>
<p><b>Customer's message:</b></p>
<blockquote style="border-left:3px solid #ccc;padding-left:12px;color:#555;">{body_text_preview}</blockquote>
<br>
<p>Please reply directly to the customer at <a href="mailto:{sender_addr}">{sender_addr}</a>.</p>
<p style="color:#888;font-size:12px;">— Uniram AI Support System</p>"""
                    send_email(
                        from_mailbox=SUPPORT_MAILBOX,
                        to_addresses=["sales@uniram.com"],
                        subject=f"[Pricing Inquiry] {subject}",
                        body_html=fallback_body
                    )
            else:
                print(f"  [DRY RUN] Would forward pricing inquiry (with attachments) to sales@uniram.com")
            mark_as_read(SUPPORT_MAILBOX, msg_id)
            move_to_processed(SUPPORT_MAILBOX, msg_id)
            stats["escalated"] += 1
            log_action(email, classification, "forwarded_to_sales", 0)
            continue

        # Vague email check — ask customer for more details before escalating
        if classification.get("is_vague"):
            customer_name = get_sender_name(email)
            product_label = product or "your unit"
            first_name = customer_name.split()[0] if customer_name and customer_name != "Customer" else None
            greeting = first_name if first_name else "Hi"
            language = classification.get("language", "en")

            if language == "fr":
                clarify_text = (
                    f"{greeting},\n\n"
                    f"Merci de nous avoir contactés au sujet de {product_label}.\n\n"
                    f"Pour vous aider efficacement, pourriez-vous nous fournir quelques détails supplémentaires :\n"
                    f"- Quel est le symptôme exact ? (ex. : ne démarre pas, fuite, bruit inhabituel, message d'erreur)\n"
                    f"- Depuis combien de temps le problème existe-t-il ?\n"
                    f"- Y a-t-il eu un événement déclencheur (chute, surchauffe, changement de solvant) ?\n"
                    f"- Des photos ou vidéos seraient très utiles si possible.\n\n"
                    f"Dès que nous aurons ces informations, nous pourrons vous orienter rapidement."
                )
            else:
                clarify_text = (
                    f"{greeting},\n\n"
                    f"Thanks for reaching out about your {product_label}.\n\n"
                    f"To help you as quickly as possible, could you share a few more details:\n"
                    f"- What exactly is the symptom? (e.g. won't start, leaking, unusual noise, error code/light)\n"
                    f"- How long has this been happening?\n"
                    f"- Was there anything that triggered it — a drop, overheating, solvent change, power issue?\n"
                    f"- Photos or a short video would be really helpful if you can share them.\n\n"
                    f"Once we have those details we can point you in the right direction quickly."
                )

            clarify_html = build_reply_html(clarify_text, email)
            print(f"    Vague email — sending clarification request to {sender}")
            if not dry_run:
                send_reply(SUPPORT_MAILBOX, msg_id, clarify_html)
            else:
                print(f"  [DRY RUN] Would send clarification request")
            mark_as_read(SUPPORT_MAILBOX, msg_id)
            move_to_processed(SUPPORT_MAILBOX, msg_id)
            stats["replied"] += 1
            log_action(email, classification, "clarification_requested", 0)
            continue

        core_question = classification.get("core_question") or extract_text(email)[:500]

        # Step 2: Retrieve context
        contexts = retrieve_context(core_question, product=product,
                                    n_results=5, kb_path=kb_path)
        print(f"    Retrieved {len(contexts)} context chunks from knowledge base")

        # Step 3: Generate reply
        customer_name = get_sender_name(email)
        reply_text, confidence, needs_escalation, escalation_reason = generate_reply(
            customer_name=customer_name,
            question=core_question,
            contexts=contexts,
            category=category,
            product=product,
            language=language
        )
        print(f"    AI confidence: {confidence:.2f} | Needs escalation: {needs_escalation}")

        stats["processed"] += 1

        # Step 4: Safety-critical override — ALWAYS escalate
        body_text_full = extract_text(email)
        if is_safety_critical(body_text_full, category):
            print(f"    SAFETY-CRITICAL: forced escalation to Ken")
            escalate_to_ken(
                email=email,
                classification=classification,
                ai_draft="[SAFETY-CRITICAL: solvent/chemical/temperature question — requires engineer review]",
                escalation_reason="Safety-critical: solvent/chemical/temperature question",
                dry_run=dry_run
            )
            mark_as_read(SUPPORT_MAILBOX, msg_id)
            move_to_processed(SUPPORT_MAILBOX, msg_id)
            stats["escalated"] += 1
            log_action(email, classification, "escalated_safety_critical", 0)
            continue

        # Step 5: Reply or escalate
        if confidence >= CONFIDENCE_THRESHOLD and not needs_escalation:
            reply_html = build_reply_html(reply_text, email)

            if dry_run:
                print(f"\n  [DRY RUN] Would reply to: {sender}")
                print(f"  Reply preview:\n{reply_text[:400]}\n")
            else:
                success = send_reply(SUPPORT_MAILBOX, msg_id, reply_html)
                if success:
                    print(f"    Reply sent to {sender}")
                    stats["replied"] += 1
                else:
                    print(f"    Failed to send reply")

            mark_as_read(SUPPORT_MAILBOX, msg_id)
            move_to_processed(SUPPORT_MAILBOX, msg_id)
            log_action(email, classification, "replied", confidence, reply_text)

        else:
            reason = escalation_reason or f"Low confidence ({confidence:.2f})"
            escalate_to_ken(email, classification, reply_text, reason, dry_run=dry_run)
            if not dry_run:
                mark_as_read(SUPPORT_MAILBOX, msg_id)
                move_to_processed(SUPPORT_MAILBOX, msg_id)
            stats["escalated"] += 1
            print(f"    Escalated to Ken — {reason}")
            log_action(email, classification, "escalated", confidence)

        print()

    print(f"{'='*60}")
    print(f"  DONE: {stats['processed']} processed | {stats['replied']} replied | "
          f"{stats['escalated']} escalated | {stats['skipped']} skipped")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()
    process_emails(dry_run=args.dry_run)
