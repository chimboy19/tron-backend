# services/graph_email_service.py
import os
import json
import re
import time
import csv
import base64
import logging
import tempfile
import requests
from datetime import datetime, timezone
from urllib.parse import quote
from msal import ConfidentialClientApplication
from fpdf import FPDF, XPos, YPos

from config import Config
from models import db, IncomingQuotationRequest, Customer, Quotation,User,Notification
from utils.ocr_utils import extract_items_from_text, extract_items_from_attachment
from utils.xlsx_loader import generate_new_ucod

log = logging.getLogger(__name__)

# ==========================
# 🔧 HELPER FUNCTIONS (SAME AS OUTLOOK VERSION)
# ==========================
def normalize_text(text: str) -> str:
    if not text:
        return ""
    text = re.sub(r'[０-９]', lambda m: str(ord(m.group()) - 0xFEE0), text)
    return re.sub(r'\s+', ' ', text.strip())




def is_quotation_request(subject: str, body: str) -> bool:
    full_text = normalize_text(f"{subject} {body}")
    exclusions = [
        r"microsoft.*account.*security",
        r"login.*alert",
        r"two[-\s]?factor.*code",
    ]
    for pattern in exclusions:
        if re.search(pattern, full_text, re.IGNORECASE):
            log.info(f"⏭️ Skipping email (matched exclusion pattern: {pattern})")
            return False
    match = re.search(r"(quotation|quote|見積|お見積|H\d{6})", full_text, re.IGNORECASE)
    if match:
        log.info(f"✅ Quotation request detected: '{match.group()}'")
        return True
    log.info("❌ Not a quotation request")
    return False




def extract_customer_info_from_email(body: str) -> tuple[str, str]:
    sender_name = ""
    name_match = re.search(r'(?:from|送信者|名前|氏名|From|Name)[:：]?\s*([^\n\r,。@]+)', body, re.IGNORECASE)
    if name_match:
        sender_name = name_match.group(1).strip()
    tel = ""
    tel_match = re.search(r'(\d{2,4}[-\s]?\d{2,4}[-\s]?\d{3,4})', body)
    if tel_match:
        tel = tel_match.group(1).replace(" ", "-")
    if sender_name:
        sender_name = re.sub(r'[<>\[\]\(\)\d]', '', sender_name).strip()
        sender_name = sender_name.split("@")[0].strip()
    return sender_name or "新規顧客", tel




def load_processed_emails():
    try:
        if os.path.exists("data/processed_graph_emails.json"):
            with open("data/processed_graph_emails.json", "r", encoding='utf-8') as f:
                return set(json.load(f))
    except Exception as e:
        log.warning(f"Could not load processed emails: {e}")
    return set()

def save_processed_emails(processed_set):
    try:
        os.makedirs("data", exist_ok=True)
        with open("data/processed_graph_emails.json", "w", encoding='utf-8') as f:
            json.dump(list(processed_set), f, ensure_ascii=False)
    except Exception as e:
        log.warning(f"Could not save processed emails: {e}")



# ==========================
# 🔐 GRAPH AUTH
# ==========================
_TOKEN_CACHE = {}

def get_graph_token():
    global _TOKEN_CACHE
    now = time.time()
    cached = _TOKEN_CACHE.get('token')
    if cached and cached['expires_at'] > now:
        return cached['access_token']

    app = ConfidentialClientApplication(
        client_id=Config.GRAPH_CLIENT_ID,
        client_credential=Config.GRAPH_CLIENT_SECRET,
        authority=f"https://login.microsoftonline.com/{Config.GRAPH_TENANT_ID}"  
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"]) 

    if "access_token" not in result:
        error = result.get("error")
        desc = result.get("error_description")
        log.error(f"❌ Graph token acquisition failed: {error} - {desc}")
        return None

    _TOKEN_CACHE['token'] = {
        'access_token': result['access_token'],
        'expires_at': now + result['expires_in'] - 300
    }
    log.info("✅ Graph token acquired successfully")
    return result['access_token']

def graph_request(method, url, **kwargs):
    token = get_graph_token()
    if not token:
        raise Exception("No Graph token")
    headers = {"Authorization": f"Bearer {token}"}
    if "json" in kwargs:
        headers["Content-Type"] = "application/json"
    resp = requests.request(method, f"https://graph.microsoft.com/v1.0{url}", headers=headers, **kwargs)  
    resp.raise_for_status()
    return resp.json()



# ==========================
# ✉️ DRAFT REPLY CREATION (DOES NOT SEND)
# ==========================
def create_quotation_draft_via_graph(mailbox: str, message_id: str, to_email: str, customer, quotation_id, db_session):
    """Creates a reply DRAFT with PDF + CSV attachments using Microsoft Graph."""
    try:
        quotation = db_session.query(Quotation).filter_by(id=quotation_id).first()
        if not quotation:
            log.error(f"❌ Quotation {quotation_id} not found")
            return

        # -------------------------
        # 1. Build results payload
        # -------------------------
        results = []
        total_amount = 0

        for item in quotation.items:
            product = item.product
            if product:
                results.append({
                    "input_code": product.hcod,
                    "HNM": product.hnm,
                    "qty": item.quantity,
                    "price": float(item.unit_price or 0),
                    "lead_time_days": item.lead_time_days or 0,
                    "delivery_date": item.estimated_delivery_date.strftime("%Y/%m/%d") if item.estimated_delivery_date else "TBD",
                    "stock_status": item.stock_status or "Unknown",
                    "supplier_name": getattr(item, 'supplier_name', 'Internal')
                })
                total_amount += float(item.unit_price or 0) * item.quantity
            else:
                results.append({
                    "input_code": getattr(item, 'notes', 'Unknown').split(':')[-1].strip()
                    if 'Digi-Key' in getattr(item, 'notes', '') else getattr(item, 'input_code', 'Unknown'),
                    "HNM": getattr(item, 'notes', ''),
                    "qty": item.quantity,
                    "price": float(item.unit_price or 0),
                    "lead_time_days": item.lead_time_days or 0,
                    "delivery_date": item.estimated_delivery_date.strftime("%Y/%m/%d") if item.estimated_delivery_date else "TBD",
                    "stock_status": item.stock_status or "Unknown",
                    "supplier_name": getattr(item, 'supplier_name', 'External')
                })
                total_amount += float(item.unit_price or 0) * item.quantity

        timestamp = int(time.time())
        temp_dir = tempfile.gettempdir()

        # -------------------------
        # 2. Generate CSV
        # -------------------------
        csv_path = os.path.join(temp_dir, f"order_{timestamp}.csv")
        with open(csv_path, "w", encoding="utf-8", newline="") as f:
            writer = csv.writer(f)
            writer.writerow([
                "Manufacturer Part Number", "Quantity", "Lead Time (Days)",
                "Estimated Delivery Date", "Stock Status", "Supplier"
            ])
            for r in results:
                mpn = r.get("HNM") or r.get("input_code") or ''
                writer.writerow([
                    mpn,
                    r.get("qty", 1),
                    r.get("lead_time_days", 0),
                    r.get("delivery_date", "TBD"),
                    r.get("stock_status", "Unknown"),
                    r.get("supplier_name", "Unknown")
                ])

        # -------------------------
        # 3. Generate PDF
        # -------------------------
        pdf_path = os.path.join(temp_dir, f"quotation_{timestamp}.pdf")
        pdf = FPDF()
        pdf.add_page()

        font_path = os.path.join(os.path.dirname(__file__), "..", "fonts", "NotoSansJP-Regular.ttf")
        if os.path.exists(font_path):
            pdf.add_font("NotoSansJP", "", font_path, uni=True)
            pdf.set_font("NotoSansJP", size=12)
        else:
            pdf.set_font("Arial", size=12)

        pdf.cell(0, 10, f"{customer.name} 様", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.cell(0, 8, "お見積書", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.ln(4)

        headers = ["品番", "品名/説明", "数量", "単価", "金額", "納期", "納入予定日", "在庫状況", "仕入先"]
        pdf.cell(0, 8, " | ".join(headers), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.cell(0, 4, "-" * 180, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

        total = 0
        for r in results:
            qty = int(r.get("qty", 1))
            price = float(r.get("price", 0))
            line_total = qty * price
            total += line_total

            desc = (r.get("HNM") or "")[:35]
            pdf.cell(
                0, 8,
                f"{r['input_code']} | {desc} | {qty} | {int(price):,} | {int(line_total):,} | "
                f"{r.get('lead_time_days', 0)}日 | {r.get('delivery_date', 'TBD')} | {r.get('stock_status', 'Unknown')} | {r.get('supplier_name', 'Internal')[:10]}",
                new_x=XPos.LMARGIN, new_y=YPos.NEXT
            )

        pdf.output(pdf_path)

        # -------------------------
        # 4. Build HTML body
        # -------------------------
        pdf_filename = os.path.basename(pdf_path)
        csv_filename = os.path.basename(csv_path)
        html_body = build_html_email(customer.name, results, pdf_filename, csv_filename)

        # Read attachments
        with open(pdf_path, "rb") as f:
            pdf_bytes = f.read()
        with open(csv_path, "rb") as f:
            csv_bytes = f.read()

        # -------------------------
        # 5. Create reply draft
        # -------------------------
        log.info(f"📧 Creating draft reply to message {message_id}...")
        draft_resp = graph_request("POST", f"/users/{mailbox}/messages/{message_id}/createReply", json={})
        draft_id = draft_resp["id"]

        # -------------------------
        # 6. Update draft content
        # -------------------------
        graph_request(
            "PATCH",
            f"/users/{mailbox}/messages/{draft_id}",
            json={
                "subject": f"【お見積り】{customer.name} 様",
                "toRecipients": [{"emailAddress": {"address": to_email}}],
                "body": {"contentType": "html", "content": html_body}
            }
        )

        # -------------------------
        # 7. Upload attachments
        # -------------------------
        graph_request(
            "POST",
            f"/users/{mailbox}/messages/{draft_id}/attachments",
            json={
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": pdf_filename,
                "contentType": "application/pdf",
                "contentBytes": base64.b64encode(pdf_bytes).decode("utf-8")
            }
        )
        graph_request(
            "POST",
            f"/users/{mailbox}/messages/{draft_id}/attachments",
            json={
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": csv_filename,
                "contentType": "text/csv",
                "contentBytes": base64.b64encode(csv_bytes).decode("utf-8")
            }
        )

        log.info(f"✅ Draft reply created successfully (Draft ID: {draft_id})")

    except Exception as e:
        log.error(f"❌ Failed to create reply draft: {e}", exc_info=True)
        raise

    finally:
        # Safe cleanup
        try:
            if os.path.exists(pdf_path):
                os.remove(pdf_path)
            if os.path.exists(csv_path):
                os.remove(csv_path)
        except Exception:
            pass




def build_html_email(customer_name, results, pdf_filename, csv_filename):
    table_style = "border-collapse:collapse;width:100%;max-width:800px;"
    th_style = "border:1px solid #d0d0d0;padding:8px;text-align:left;font-weight:bold;background:#f7f7f7;"
    td_style = "border:1px solid #d0d0d0;padding:8px;text-align:left;"
    right_style = td_style + "text-align:right;"
    center_style = td_style + "text-align:center;"
    rows_html = ""
    total_price = 0
    for r in results:
        qty = int(r.get('qty', 1))
        price = int(float(r.get('price', 0) or 0))
        line_total = price * qty
        total_price += line_total
        desc = (r.get('HNM', '') or '')[:50]
        lead_time = r.get('lead_time_days', 0)
        delivery_date = r.get('delivery_date', 'TBD')
        stock_status = r.get('stock_status', 'Unknown')
        supplier = r.get('supplier_name', 'Internal')
        stock_style = center_style
        if "在庫あり" in stock_status:
            stock_style += "background-color:#d4edda;"
        elif "在庫なし" in stock_status:
            stock_style += "background-color:#f8d7da;"
        elif "要確認" in stock_status or "確認中" in stock_status:
            stock_style += "background-color:#fff3cd;"
        supplier_style = center_style
        if "Digi-Key" in supplier:
            supplier_style += "background-color:#e3f2fd;"
        elif "Internal" in supplier:
            supplier_style += "background-color:#f3e5f5;"
        rows_html += (
            "<tr>"
            f"<td style='{td_style}'>{r.get('input_code','')}</td>"
            f"<td style='{td_style}'>{desc}</td>"
            f"<td style='{right_style}'>{qty}</td>"
            f"<td style='{right_style}'>¥{price:,}</td>"
            f"<td style='{right_style}'>¥{line_total:,}</td>"
            f"<td style='{center_style}'>{lead_time}</td>"
            f"<td style='{center_style}'>{delivery_date}</td>"
            f"<td style='{stock_style}'>{stock_status}</td>"
            f"<td style='{supplier_style}'>{supplier}</td>"
            "</tr>"
        )
    total_row = (
        "<tr>"
        f"<td style='{td_style}' colspan='4'><strong>合計金額（概算）</strong></td>"
        f"<td style='{right_style}'><strong>¥{int(total_price):,}</strong></td>"
        f"<td style='{center_style}' colspan='4'></td>"
        "</tr>"
    )
    return f"""
    <html>
    <body style="font-family: -apple-system, 'Segoe UI', Roboto, 'Noto Sans JP', Arial, sans-serif; color:#222;">
      <div style="max-width:900px;margin:0 auto;">
        <p>{customer_name} 様</p>
        <p>いつもお世話になっております。下記の通りお見積書を添付いたします。</p>
        <table role="presentation" style="{table_style}">
          <thead>
            <tr>
              <th style="{th_style}">品番</th>
              <th style="{th_style}">品名 / 説明</th>
              <th style="{th_style}">数量</th>
              <th style="{th_style}">単価</th>
              <th style="{th_style}">金額</th>
              <th style="{th_style}">納期 (日数)</th>
              <th style="{th_style}">納入予定日</th>
              <th style="{th_style}">在庫状況</th>
              <th style="{th_style}">仕入先</th>
            </tr>
          </thead>
          <tbody>
            {rows_html}
            {total_row}
          </tbody>
        </table>
        <p style="margin-top:16px;">添付ファイル：</p>
        <ul>
          <li>見積書 (PDF): {pdf_filename}</li>
          <li>発注用CSV: {csv_filename}</li>
        </ul>
        <p style="margin-top:12px;">※納期・在庫は別途確認の上、正式なご案内を差し上げます。</p>
        <p>何卒よろしくお願いいたします。<br>営業部</p>
      </div>
    </body>
    </html>
    """

# ==========================
# 📧 EMAIL PROCESSING
# ==========================
def process_graph_email(email_data, db_session, quotation_service, inventory_service, mailbox: str):
    subject = email_data.get("subject", "")
    body_preview = email_data.get("bodyPreview", "")
    sender_email = email_data["from"]["emailAddress"]["address"]
    received_time = email_data.get("receivedDateTime")
    message_id = email_data["id"]

    log.info(f"📧 Processing Graph email: '{subject}' from {sender_email}")

    # Extract from body
    items = extract_items_from_text(body_preview)

    # Extract from attachments
    if email_data.get("hasAttachments"):
        attach_resp = graph_request("GET", f"/users/{mailbox}/messages/{message_id}/attachments")
        for att in attach_resp.get("value", []):
            if att.get("contentType", "").startswith(("image/", "application/pdf")):
                content_bytes = base64.b64decode(att["contentBytes"])
                ext = os.path.splitext(att["name"])[1] or ".pdf"
                with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
                    tmp.write(content_bytes)
                    tmp_path = tmp.name
                try:
                    items.extend(extract_items_from_attachment(tmp_path))
                finally:
                    os.unlink(tmp_path)

    if not items:
        log.info("❌ No valid item codes found.")
        return False
    
    valid_items = []
    for code, qty in items:
        if code and code.strip() and code.strip() != "```":
            try:
                qty_val = int(qty) if qty else 1
            except:
                qty_val = 1
            valid_items.append((code.strip(), qty_val))
    if not valid_items:
        return False

    agg = {}
    for code, qty in valid_items:
        if isinstance(code, str):
            if re.match(r'^H\d{6}$', code.strip(), re.IGNORECASE):
                final_code = code.strip().upper()
            else:
                final_code = code.strip()
        else:
            final_code = str(code).strip()
        if final_code:
            agg[final_code] = agg.get(final_code, 0) + qty
    item_list = [(k, v) for k, v in agg.items()]
    log.info(f"📦 Aggregated items: {item_list}")
    customer_name, customer_tel = extract_customer_info_from_email(body_preview)

    incoming = IncomingQuotationRequest(
        subject=subject,
        body=body_preview,
        sender=sender_email,
        received_date=datetime.fromisoformat(received_time.replace("Z", "+00:00")).replace(tzinfo=None),
        status='pending',
        items_data=[{"hcod": hcod, "qty": qty} for hcod, qty in item_list],
        customer_name=customer_name,
        customer_tel=customer_tel
    )
    db_session.add(incoming)
    db_session.commit()

    # Process quotation
    customer = db_session.query(Customer).filter_by(email=sender_email).first()
    if not customer:
        existing_ucods = {c.ucod for c in db_session.query(Customer).all()}
        new_ucod = generate_new_ucod(existing_ucods)
        customer = Customer(ucod=new_ucod, name=customer_name, email=sender_email, phone=customer_tel)
        db_session.add(customer)
        db_session.commit()
        log.info(f"👤 Created new customer: {customer.name} ({customer.ucod})")

    result = quotation_service.create_quotation(customer.id, [{"hcod": hcod, "qty": qty} for hcod, qty in item_list])
    if result.get("success"):
        incoming.status = 'processed'
        db_session.commit()
        log.info(f"✅ Quotation {result['quotation_id']} created for incoming request {incoming.id}")

        try:
            create_quotation_draft_via_graph(
                mailbox=mailbox,
                message_id=message_id,
                to_email=sender_email,
                customer=customer,
                quotation_id=result['quotation_id'],
                db_session=db_session
            )
            admin_users = User.query.filter_by(role='admin').all()
            for admin in admin_users:
                notif = Notification(
                    user_id=admin.id,
                    type='draft_created',
                    message=f"Draft reply created for quotation email from {sender_email}. Check Drafts folder."
                )
                db.session.add(notif)
            db.session.commit()
        except Exception as e:
            log.error(f" Draft creation failed: {e}")

      
        graph_request("PATCH", f"/users/{mailbox}/messages/{message_id}", json={"isRead": True})
        return True
    else:
        incoming.status = 'error'
        incoming.notes = result.get("error", "Unknown")
        db_session.commit()
        return False

# ==========================
# 🔄 POLLING FUNCTION
# ==========================
def poll_and_process_emails_graph(db_session, quotation_service, inventory_service):
    processed_ids = load_processed_emails()
    new_processed = set()
    processed_count = 0
    mailbox = quote(Config.GRAPH_MAIL_BOX) if hasattr(Config, 'GRAPH_MAIL_BOX') else quote(Config.GRAPH_MAILBOX)

    try:
        resp = graph_request("GET", f"/users/{mailbox}/mailFolders/inbox/messages?$filter=isRead eq false&$top=50")
        emails = resp.get("value", [])

        for email in emails:
            email_id = email["id"]
            if email_id in processed_ids:
                continue
            if is_quotation_request(email.get("subject", ""), email.get("bodyPreview", "")):
                if process_graph_email(email, db_session, quotation_service, inventory_service, mailbox):
                    processed_count += 1
                    new_processed.add(email_id)

        if new_processed:
            processed_ids.update(new_processed)
            save_processed_emails(processed_ids)

        return processed_count
    except Exception as e:
        log.error(f"Graph polling error: {e}", exc_info=True)
        return 0