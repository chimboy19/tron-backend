


# import sys
# import os
# import json
# import re
# import tempfile
# import base64
# import logging
# import pythoncom
# import win32com.client
# import time
# import csv
# from fpdf import FPDF, XPos, YPos
# from models import db, IncomingQuotationRequest, Customer, Product, Quotation, QuotationItem
# from utils.ocr_utils import ocr_pdf_with_openai, extract_items_from_text, extract_items_from_attachment, correct_ocr_code
# from config import Config
# from datetime import datetime

# backend_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
# if backend_dir not in sys.path:
#     sys.path.insert(0, backend_dir)

# log = logging.getLogger(__name__)

# # ==========================
# # 🔧 HELPER FUNCTIONS (EXACT SAME AS STANDALONE SCRIPT)
# # ==========================
# def normalize_text(text: str) -> str:
#     if not text:
#         return ""
#     text = re.sub(r'[０-９]', lambda m: str(ord(m.group()) - 0xFEE0), text)
#     return re.sub(r'\s+', ' ', text.strip())

# # def is_quotation_request(subject: str, body: str) -> bool:
# #     full_text = normalize_text(f"{subject} {body}")
    
# 
# #     exclude_patterns = [
# #         r'microsoft.*account.*connected',
# #         r'account.*security.*notification',
# #         r'new app.*connected',
# #         r'security.*alert',
# #         r'login.*notification'
# #     ]
    
# #     for pattern in exclude_patterns:
# #         if re.search(pattern, full_text, re.IGNORECASE):
# #             log.info(f"⏭️ Skipping non-quotation email (matches exclude pattern: {pattern})")
# #             return False
    
# #    
# #     japanese_keywords = [
# #         '見積', '見積もり', 'お見積', '見積書', '見積依頼',
# #         '価格', '見積価格', 'お見積もり', '見積り',
# #         'quotation', 'quote', 'pricing', 'estimate'
# #     ]
    
# #     # Also look for H-codes in the email
# #     h_code_pattern = r'H\d{6}'
    
# #     has_quotation_keyword = any(keyword in full_text for keyword in japanese_keywords)
# #     has_h_codes = bool(re.search(h_code_pattern, full_text, re.IGNORECASE))
    
# #     if has_quotation_keyword or has_h_codes:
# #         log.info(f"✅ Genuine quotation request detected - Japanese keywords or H-codes found")
# #         return True
                
# #     return False

# def is_quotation_request(subject: str, body: str) -> bool:
#     # Combine and normalize text
#     full_text = normalize_text(f"{subject} {body}")

#     # --- Optional exclusions (very safe ones only) ---
#     exclusions = [
#         r"microsoft.*account.*security",
#         r"login.*alert",
#         r"two[-\s]?factor.*code",
#     ]
#     for pattern in exclusions:
#         if re.search(pattern, full_text, re.IGNORECASE):
#             log.info(f"⏭️ Skipping email (matched exclusion pattern: {pattern})")
#             return False

#     # --- Main quotation detection (simple + reliable) ---
#     match = re.search(
#         r"(quotation|quote|見積|お見積|H\d{6})",
#         full_text,
#         re.IGNORECASE
#     )

#     if match:
#         log.info(f"✅ Quotation request detected: '{match.group()}'")
#         return True

#     log.info("❌ Not a quotation request")
#     return False


# def extract_customer_info_from_email(msg) -> tuple[str, str]:
#     body = (msg.Body or "").strip()
#     sender_name = ""
#     try:
#         if hasattr(msg, "SenderName") and msg.SenderName:
#             sender_name = msg.SenderName.strip()
#         elif hasattr(msg, "Sender") and msg.Sender:
#             sender_name = msg.Sender.Name.strip()
#     except Exception:
#         pass

#     if not sender_name:
#         name_match = re.search(r'(?:from|送信者|名前|氏名|From|Name)[:：]?\s*([^\n\r,。@]+)', body, re.IGNORECASE)
#         if name_match:
#             sender_name = name_match.group(1).strip()

#     tel = ""
#     tel_match = re.search(r'(\d{2,4}[-\s]?\d{2,4}[-\s]?\d{3,4})', body)
#     if tel_match:
#         tel = tel_match.group(1).replace(" ", "-")

#     if sender_name:
#         sender_name = re.sub(r'[<>\[\]\(\)\d]', '', sender_name).strip()
#         sender_name = sender_name.split("@")[0].strip()

#     return sender_name or "新規顧客", tel

# def load_processed_emails():
#     """Load processed emails from file"""
#     try:
#         if os.path.exists("data/processed_emails.json"):
#             with open("data/processed_emails.json", "r", encoding='utf-8') as f:
#                 return set(json.load(f))
#     except Exception as e:
#         log.warning(f"Could not load processed emails: {e}")
#     return set()

# def save_processed_emails(processed_set):
#     """Save processed emails to file"""
#     try:
#         os.makedirs("data", exist_ok=True)
#         with open("data/processed_emails.json", "w", encoding='utf-8') as f:
#             json.dump(list(processed_set), f, ensure_ascii=False)
#     except Exception as e:
#         log.warning(f"Could not save processed emails: {e}")

# # ==========================
# # 📧 EMAIL PROCESSING (EXACT SAME LOGIC AS STANDALONE)
# # ==========================
# def process_outlook_email(msg, db_session, quotation_service, inventory_service):
#     """
#     Core logic to process a single Outlook email message for quotation requests.
#     Uses the EXACT SAME logic as the standalone script.
#     """
#     subject, body = msg.Subject or "", msg.Body or ""
#     sender = getattr(msg, "SenderEmailAddress", None) or "unknown@example.com"

#     log.info(f"📧 Processing email: '{subject}' from {sender}")

#     # Extract items from email body (SAME AS STANDALONE)
#     items = extract_items_from_text(body)
#     log.info(f"📝 Found {len(items)} items in email body")

#     # Extract items from attachments (SAME AS STANDALONE)
#     try:
#         if getattr(msg, "Attachments", None) and msg.Attachments.Count > 0:
#             log.info(f"📎 Found {msg.Attachments.Count} attachments")
#             for i in range(1, msg.Attachments.Count + 1):
#                 try:
#                     attachment = msg.Attachments.Item(i)
#                     filename = getattr(attachment, "FileName", "attachment")
#                     log.info(f"📎 Processing attachment: {filename}")
                    
#                     attachment_items = extract_items_from_attachment(attachment)
#                     items.extend(attachment_items)
#                     log.info(f"📦 Extracted {len(attachment_items)} items from attachment {filename}")
                    
#                 except Exception as e:
#                     log.warning(f"Attachment parse error: {e}", exc_info=True)
#     except Exception as e:
#         log.warning(f"Failed to read attachments container: {e}", exc_info=True)

#     log.info(f"📦 Total raw item candidates: {items}")

#     if not items:
#         log.info("❌ No valid item codes found in email body or attachments.")
#         return False

#     # Filter and aggregate items (SAME AS STANDALONE)
#     valid_items = []
#     for code, qty in items:
#         if code and code.strip() and code.strip() != "```":
#             try:
#                 qty_val = int(qty) if qty else 1
#             except (ValueError, TypeError):
#                 qty_val = 1
#             valid_items.append((code.strip(), qty_val))
    
#     if not valid_items:
#         log.info("No valid items after filtering.")
#         return False

#     # Aggregate items by code (SAME AS STANDALONE)
#     agg = {}
#     for code, qty in valid_items:
#         if isinstance(code, str):
#             if re.match(r'^H\d{6}$', code.strip(), re.IGNORECASE):
#                 final_code = code.strip().upper()
#             else:
#                 final_code = code.strip()
#         else:
#             final_code = str(code).strip()
#         if final_code:
#             agg[final_code] = agg.get(final_code, 0) + qty

#     item_list = [(k, v) for k, v in agg.items()]
#     log.info(f"📦 Aggregated unique items to process: {item_list}")

#     # *** STORE IN DATABASE (FLASK-SPECIFIC) ***
#     customer_name, customer_tel = extract_customer_info_from_email(msg)
    
#     incoming_request = IncomingQuotationRequest(
#         subject=subject,
#         body=body,
#         sender=sender,
#         received_date=msg.ReceivedTime,
#         status='pending',
#         items_data=[{"hcod": hcod, "qty": qty} for hcod, qty in item_list],
#         customer_name=customer_name,
#         customer_tel=customer_tel
#     )
#     db_session.add(incoming_request)
#     db_session.commit()
#     log.info(f"💾 Stored incoming email request (ID: {incoming_request.id})")

#     # *** PROCESS WITH QUOTATION SERVICE (FLASK-SPECIFIC) ***
#     try:
#         # Find or create customer
#         customer = db_session.query(Customer).filter_by(email=sender).first()
#         if not customer:
#             from utils.xlsx_loader import generate_new_ucod
#             existing_ucods = {c.ucod for c in db_session.query(Customer).all()}
#             new_ucod = generate_new_ucod(existing_ucods)
#             customer = Customer(
#                 ucod=new_ucod, 
#                 name=customer_name or "Email Customer", 
#                 email=sender, 
#                 phone=customer_tel
#             )
#             db_session.add(customer)
#             db_session.commit()
#             log.info(f"👤 Created new customer: {customer.name} ({customer.ucod})")

#         # Process quotation using the service
#         result = quotation_service.create_quotation(
#             customer.id, 
#             [{"hcod": hcod, "qty": qty} for hcod, qty in item_list]
#         )

#         if result.get("success"):
#             incoming_request.status = 'processed'
#             incoming_request.notes = f"Successfully created quotation ID {result.get('quotation_id')}"
#             db_session.commit()
#             log.info(f"✅ Incoming request {incoming_request.id} processed successfully. Quotation ID: {result.get('quotation_id')}")
            
#             # *** NEW: CREATE OUTLOOK DRAFT AND MOVE EMAIL (LIKE AUTOMATION SCRIPT) ***
#             try:
#                 create_quotation_draft_and_move_email(msg, customer, result, incoming_request, db_session)
#             except Exception as draft_error:
#                 log.error(f"❌ Failed to create draft/move email: {draft_error}")
#                 # Don't fail the whole process if draft creation fails
                
#             return True
#         else:
#             incoming_request.status = 'error'
#             incoming_request.notes = f"Processing failed: {result.get('error', 'Unknown error')}"
#             db_session.commit()
#             log.error(f"❌ Failed to process incoming request {incoming_request.id}: {result.get('error', 'Unknown error')}")
#             return False
#     except Exception as e:
#         log.error(f"❌ Error processing incoming request {incoming_request.id}: {e}", exc_info=True)
#         incoming_request.status = 'error'
#         incoming_request.notes = f"Processing error: {str(e)}"
#         db_session.commit()
#         return False

# def create_quotation_draft_and_move_email(original_msg, customer, quotation_result, incoming_request, db_session):
#     """Create Outlook draft with quotation and move original email to processed folder."""
#     try:
#         import win32com.client
#         import tempfile
#         import os
        
#         quotation_id = quotation_result.get('quotation_id')
#         log.info(f"📧 Creating Outlook draft for quotation {quotation_id}")
        
#         # Get quotation details from database
#         quotation = db_session.query(Quotation).filter_by(id=quotation_id).first()
#         if not quotation:
#             log.error(f"❌ Quotation {quotation_id} not found in database")
#             return
        
#         # Generate quotation results for PDF/CSV
#         results = []
#         total_amount = 0
#         for item in quotation.items:
#             product = item.product
#             if product:
#                 results.append({
#                     "input_code": product.hcod,
#                     "HNM": product.hnm,
#                     "qty": item.quantity,
#                     "price": float(item.unit_price or 0),
#                     "lead_time_days": item.lead_time_days or 0,
#                     "delivery_date": item.estimated_delivery_date.strftime("%Y/%m/%d") if item.estimated_delivery_date else "TBD",
#                     "stock_status": item.stock_status or "Unknown",
#                     "supplier_name": getattr(item, 'supplier_name', 'Internal')
#                 })
#                 total_amount += float(item.unit_price or 0) * item.quantity
#             else:
#                 # Handle Digi-Key items or items without product association
#                 results.append({
#                     "input_code": getattr(item, 'notes', 'Unknown').split(':')[-1].strip() if 'Digi-Key' in getattr(item, 'notes', '') else getattr(item, 'input_code', 'Unknown'),
#                     "HNM": getattr(item, 'notes', ''),
#                     "qty": item.quantity,
#                     "price": float(item.unit_price or 0),
#                     "lead_time_days": item.lead_time_days or 0,
#                     "delivery_date": item.estimated_delivery_date.strftime("%Y/%m/%d") if item.estimated_delivery_date else "TBD",
#                     "stock_status": item.stock_status or "Unknown",
#                     "supplier_name": getattr(item, 'supplier_name', 'External')
#                 })
#                 total_amount += float(item.unit_price or 0) * item.quantity
        
#         # Add total row
#         results.append({
#             "input_code": "合計",
#             "HNM": "",
#             "qty": "",
#             "price": "",
#             "lead_time_days": "",
#             "delivery_date": "",
#             "stock_status": "",
#             "supplier_name": "",
#             "line_total": total_amount
#         })
        
#         # Generate PDF and CSV files (like automation script)
#         timestamp = int(time.time())
#         temp_dir = tempfile.gettempdir()
        
#         # Generate CSV
#         csv_path = os.path.join(temp_dir, f"order_{timestamp}.csv")
#         with open(csv_path, 'w', encoding='utf-8', newline='') as f:
#             writer = csv.writer(f)
#             writer.writerow(['Manufacturer Part Number', 'Quantity', 'Lead Time (Days)', 'Estimated Delivery Date', 'Stock Status', 'Supplier'])
#             for r in results:
#                 if r.get("input_code") == "合計":
#                     continue
#                 mpn = r.get('HNM') or r.get('input_code') or ''
#                 qty = r.get('qty', 1)
#                 lead_time = r.get('lead_time_days', 0)
#                 delivery_date = r.get('delivery_date', 'TBD')
#                 stock_status = r.get('stock_status', 'Unknown')
#                 supplier = r.get('supplier_name', 'Unknown')
#                 writer.writerow([mpn, qty, lead_time, delivery_date, stock_status, supplier])
        
#         # Generate PDF
#         pdf_path = os.path.join(temp_dir, f"quotation_{timestamp}.pdf")
#         pdf = FPDF()
#         pdf.add_page()
        
#         # Try to use Japanese font
#         font_path = os.path.join(os.path.dirname(__file__), "..", "fonts", "NotoSansJP-Regular.ttf")
#         if os.path.exists(font_path):
#             try:
#                 pdf.add_font("NotoSansJP", "", font_path, uni=True)
#                 pdf.set_font("NotoSansJP", size=12)
#             except Exception as e:
#                 log.warning(f"Failed to load NotoSansJP font: {e}. Falling back.")
#                 pdf.set_font("Arial", size=12)
#         else:
#             pdf.set_font("Arial", size=12)
            
#         pdf.cell(0, 10, f"{customer.name} 様", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
#         pdf.cell(0, 8, "お見積書", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
#         pdf.ln(2)
#         pdf.cell(0, 6, "以下の通りお見積り申し上げます。", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
#         pdf.ln(4)
        
#         headers = ["品番", "品名/説明", "数量", "単価 (JPY)", "金額 (JPY)", "納期 (日数)", "納入予定日", "在庫状況", "仕入先"]
#         header_line = " | ".join(headers)
#         pdf.cell(0, 8, header_line, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
#         pdf.cell(0, 4, "-" * 180, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        
#         total = 0
#         for r in results:
#             if r.get("input_code") == "合計":
#                 continue
                
#             qty = int(r.get('qty', 1))
#             price = float(r.get('price') or 0)
#             line_total = price * qty
#             total += line_total
#             desc = (r.get('HNM') or '')[:35]  # Shorter description to fit
#             lead_time = r.get('lead_time_days', 0)
#             delivery_date = r.get('delivery_date', 'TBD')
#             stock_status = r.get('stock_status', 'Unknown')
#             supplier = r.get('supplier_name', 'Internal')[:10]  # Shorten supplier name
            
#             pdf.cell(0, 8, f"{r.get('input_code', '')} | {desc} | {qty} | {int(price):,} | {int(line_total):,} | {lead_time}日 | {delivery_date} | {stock_status} | {supplier}", 
#                     new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        
#         pdf.cell(0, 6, "-" * 180, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
#         pdf.cell(0, 10, f"合計金額: ¥{int(total):,}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
#         pdf.ln(6)
#         pdf.cell(0, 6, "備考:", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
#         pdf.cell(0, 6, "納期・在庫は別途ご確認ください。税抜価格で記載。", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
#         pdf.output(pdf_path)
        
#         # Create HTML email body (like automation script)
#         pdf_filename = os.path.basename(pdf_path)
#         csv_filename = os.path.basename(csv_path)
#         html_body = build_html_email(customer.name, results, pdf_filename, csv_filename)
        
#         # Create Outlook draft
#         outlook = win32com.client.Dispatch("Outlook.Application")
#         mail = outlook.CreateItem(0)
#         mail.To = original_msg.SenderEmailAddress
#         mail.Subject = f"【お見積り】{customer.name} 様"
#         mail.HTMLBody = html_body
#         mail.Attachments.Add(os.path.abspath(pdf_path))
#         mail.Attachments.Add(os.path.abspath(csv_path))
#         mail.Save()
#         log.info("✅ HTML draft created with PDF + CSV attachments.")
        
#         # Move original message to processed folder
#         namespace = outlook.GetNamespace("MAPI")
#         inbox = namespace.GetDefaultFolder(6)
#         try:
#             processed_folder = inbox.Folders("Quotations Processed")
#         except Exception:
#             processed_folder = inbox.Folders.Add("Quotations Processed")
#         original_msg.Move(processed_folder)
#         log.info("✅ Moved original message to processed folder.")
        
#         # Clean up temporary files
#         try:
#             os.remove(pdf_path)
#             os.remove(csv_path)
#         except Exception as e:
#             log.warning(f"Could not clean up temporary files: {e}")
            
#     except Exception as e:
#         log.error(f"❌ Error creating draft/moving email: {e}", exc_info=True)
#         raise

# def build_html_email(customer_name, results, pdf_filename, csv_filename):
#     """Build HTML email body like the automation script."""
#     table_style = "border-collapse:collapse;width:100%;max-width:800px;"
#     th_style = "border:1px solid #d0d0d0;padding:8px;text-align:left;font-weight:bold;background:#f7f7f7;"
#     td_style = "border:1px solid #d0d0d0;padding:8px;text-align:left;"
#     right_style = td_style + "text-align:right;"
#     center_style = td_style + "text-align:center;"
    
#     rows_html = ""
#     total_price = 0
#     for r in results:
#         if r.get("input_code") == "合計":
#             continue
            
#         try:
#             qty = int(r.get('qty', 1))
#         except:
#             qty = 1
#         price = int(float(r.get('price', 0) or 0))
#         line_total = price * qty
#         total_price += line_total
#         desc = (r.get('HNM', '') or '')[:50]  # Shorter description
#         lead_time = r.get('lead_time_days', 0)
#         delivery_date = r.get('delivery_date', 'TBD')
#         stock_status = r.get('stock_status', 'Unknown')
#         supplier = r.get('supplier_name', 'Internal')
        
#         # Add color coding for stock status
#         stock_style = center_style
#         if "在庫あり" in stock_status:
#             stock_style += "background-color:#d4edda;"
#         elif "在庫なし" in stock_status:
#             stock_style += "background-color:#f8d7da;"
#         elif "要確認" in stock_status or "確認中" in stock_status:
#             stock_style += "background-color:#fff3cd;"
            
#         # Add supplier color coding
#         supplier_style = center_style
#         if "Digi-Key" in supplier:
#             supplier_style += "background-color:#e3f2fd;"
#         elif "Internal" in supplier:
#             supplier_style += "background-color:#f3e5f5;"
            
#         rows_html += (
#             "<tr>"
#             f"<td style='{td_style}'>{r.get('input_code','')}</td>"
#             f"<td style='{td_style}'>{desc}</td>"
#             f"<td style='{right_style}'>{qty}</td>"
#             f"<td style='{right_style}'>¥{price:,}</td>"
#             f"<td style='{right_style}'>¥{line_total:,}</td>"
#             f"<td style='{center_style}'>{lead_time}</td>"
#             f"<td style='{center_style}'>{delivery_date}</td>"
#             f"<td style='{stock_style}'>{stock_status}</td>"
#             f"<td style='{supplier_style}'>{supplier}</td>"
#             "</tr>"
#         )
    
#     total_row = (
#         "<tr>"
#         f"<td style='{td_style}' colspan='4'><strong>合計金額（概算）</strong></td>"
#         f"<td style='{right_style}'><strong>¥{int(total_price):,}</strong></td>"
#         f"<td style='{center_style}' colspan='4'></td>"
#         "</tr>"
#     )
    
#     html = f"""
#     <html>
#     <body style="font-family: -apple-system, 'Segoe UI', Roboto, 'Noto Sans JP', Arial, sans-serif; color:#222;">
#       <div style="max-width:900px;margin:0 auto;">
#         <p>{customer_name} 様</p>
#         <p>いつもお世話になっております。下記の通りお見積書を添付いたします。</p>
#         <table role="presentation" style="{table_style}">
#           <thead>
#             <tr>
#               <th style="{th_style}">品番</th>
#               <th style="{th_style}">品名 / 説明</th>
#               <th style="{th_style}">数量</th>
#               <th style="{th_style}">単価</th>
#               <th style="{th_style}">金額</th>
#               <th style="{th_style}">納期 (日数)</th>
#               <th style="{th_style}">納入予定日</th>
#               <th style="{th_style}">在庫状況</th>
#               <th style="{th_style}">仕入先</th>
#             </tr>
#           </thead>
#           <tbody>
#             {rows_html}
#             {total_row}
#           </tbody>
#         </table>
#         <p style="margin-top:16px;">添付ファイル：</p>
#         <ul>
#           <li>見積書 (PDF): {pdf_filename}</li>
#           <li>発注用CSV: {csv_filename}</li>
#         </ul>
#         <p style="margin-top:12px;">※納期・在庫は別途確認の上、正式なご案内を差し上げます。</p>
#         <p>何卒よろしくお願いいたします。<br>営業部</p>
#       </div>
#     </body>
#     </html>
#     """
#     return html

# # ==========================
# # 📧 POLLING FUNCTION (SAME LOGIC AS STANDALONE)
# # ==========================
# def poll_and_process_emails(outlook, db_session, quotation_service, inventory_service):
#     """Poll Outlook for new emails and process quotation requests."""
#     try:
#         # Load previously processed emails
#         processed_email_ids = load_processed_emails()
#         new_processed = set()
        
#         namespace = outlook.GetNamespace("MAPI")
#         inbox = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
#         messages = inbox.Items
        
#         # Sort by received time (newest first)
#         messages.Sort("[ReceivedTime]", True)
        
#         processed_count = 0
        
#         for i, msg in enumerate(messages):
#             if i >= 50:  # Limit to 50 most recent emails
#                 break
                
#             try:
#                 subject = msg.Subject or ""
#                 body = msg.Body or ""
                
#                 # Create unique identifier
#                 email_id = f"{msg.EntryID}_{msg.ReceivedTime}"
                
#                 # Skip if already processed
#                 if email_id in processed_email_ids:
#                     log.debug(f"⏭️ Skipping already processed email: {subject[:30]}...")
#                     continue
                
#                 # ✅ USE THE SIMPLE DETECTION (SAME AS STANDALONE)
#                 if is_quotation_request(subject, body):
#                     log.info(f"🔍 Found quotation request: {subject[:50]}...")
                    
#                     if process_outlook_email(msg, db_session, quotation_service, inventory_service):
#                         processed_count += 1
#                         new_processed.add(email_id)
#                         log.info(f"✅ Marked email as processed: {email_id}")
                        
#             except Exception as e:
#                 log.error(f"Error processing email {i}: {e}")
#                 continue
        
#         # Save newly processed emails
#         if new_processed:
#             processed_email_ids.update(new_processed)
#             save_processed_emails(processed_email_ids)
#             log.info(f"💾 Saved {len(new_processed)} new processed emails")
                
#         log.info(f"📧 Scanning complete. Processed {processed_count} new quotation requests.")
#         return processed_count
        
#     except Exception as e:
#         log.error(f"Error polling emails: {e}")
#         return 0

# def run_outlook_monitor(outlook_app_instance, db_session, quotation_service, inventory_service):
#     """Background thread function to continuously poll Outlook."""
#     pythoncom.CoInitialize() # Initialize COM for this thread
#     try:
#         while True:
#             # Call the polling function from this service
#             processed_count = poll_and_process_emails(outlook_app_instance, db_session, quotation_service, inventory_service)
#             # Sleep for configured interval (import CHECK_INTERVAL from config if needed)
#             time.sleep(Config.CHECK_INTERVAL) # Use Config.CHECK_INTERVAL
#     except Exception as e:
#         log.error(f"Outlook monitor thread error: {e}", exc_info=True)
#     finally:
#         pythoncom.CoUninitialize() # Uninitialize COM