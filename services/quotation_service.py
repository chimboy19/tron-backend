

# backend/services/quotation_service.py
from models import Quotation, QuotationItem, Product, Customer, Stock, SupplierLeadTime, Supplier
from sqlalchemy.exc import IntegrityError
from services.inventory_service import InventoryService
from utils.xlsx_loader import load_initial_data, load_xlsx_to_db
from utils.lead_time_calculator import calculate_lead_time_and_status
from utils.calendar_utils import get_delivery_date, load_calendar
from utils.ocr_utils import correct_ocr_code
import datetime
import re
import logging
from decimal import Decimal

log = logging.getLogger(__name__)


class QuotationService:
    def __init__(self, db_session, inventory_service):
        self.db_session = db_session
        self.inventory_service = inventory_service

    def load_mhn1(self):
        """Load MHN1 data from DB for OCR correction."""
        try:
            products = self.db_session.query(Product).all()
            mhn1_data = {}
            for i, product in enumerate(products):
                mhn1_data[i] = {
                    'HCOD': product.hcod,
                    'HNM': product.hnm,
                    'HNMM': getattr(product, 'description', ''),
                    'HSRS': getattr(product, 'category', ''),
                    'TANKAU': float(product.unit_price or 0),
                    'MCOD': getattr(product, 'supplier_code', '')
                }
            return mhn1_data
        except Exception as e:
            log.warning(f"Could not load MHN1 data: {e}", exc_info=True)
            return {}
        



    def find_product_in_database(self, product_code):
        """Lookup product using HCOD first, then HNM if no match."""
        if not product_code:
            return None

        clean_code = str(product_code).replace("-", "").replace("_", "").replace(" ", "").strip().upper()

        product = self.db_session.query(Product).filter_by(hcod=clean_code).first()
        if product:
            log.info(f"✅ Exact HCOD match: '{clean_code}' -> '{product.hcod}'")
            return product
        product = self.db_session.query(Product).filter(Product.hcod.ilike(clean_code)).first()
        if product:
            log.info(f"✅ Case-insensitive HCOD match: '{clean_code}' -> '{product.hcod}'")
            return product
        if not clean_code.startswith("H"):
            h_prefixed = "H" + clean_code
            product = self.db_session.query(Product).filter_by(hcod=h_prefixed).first()
            if product:
                log.info(f"✅ H-prefixed HCOD match: '{clean_code}' -> '{product.hcod}'")
                return product

        product = self.db_session.query(Product).filter(Product.hcod.ilike(f"%{clean_code}%")).first()
        if product:
            log.info(f"✅ Partial HCOD match: '{clean_code}' -> '{product.hcod}'")
            return product
        product = self.db_session.query(Product).filter(Product.hnm.ilike(f"%{clean_code}%")).first()
        if product:
            log.info(f"✅ HNM match: '{clean_code}' found in '{product.hnm}'")
            return product

        log.info(f"❌ No database match found for: '{clean_code}'")
        return None
    
    

    def parse_delivery_date(self, date_input):
      
        if not date_input:
            return None
        try:
           
            if isinstance(date_input, datetime.date) and not isinstance(date_input, datetime.datetime):
                return date_input
            if isinstance(date_input, datetime.datetime):
                return date_input.date()
        except Exception:
          
            pass
        if isinstance(date_input, (bytes, bytearray)):
            try:
                date_input = date_input.decode('utf-8', errors='ignore')
            except Exception:
                date_input = str(date_input)
        try:
            s = str(date_input).strip()
            if not s:
                return None
            if '-' in s:
                s_clean = re.sub(r'[^0-9\-]', '', s)
                parts = [p for p in s_clean.split('-') if p != '']
                if len(parts) >= 3 and all(p.isdigit() for p in parts[:3]):
                    try:
                        y = int(parts[0])
                        m = int(parts[1])
                        d = int(parts[2])
                        return datetime.date(y, m, d)
                    except Exception:
                        try:
                            return datetime.datetime.strptime('-'.join(parts[:3]), "%Y-%m-%d").date()
                        except Exception:
                            return None
                else:
                    try:
                        return datetime.datetime.strptime(s, "%Y-%m-%d").date()
                    except Exception:
                        pass
            digits = re.sub(r'\D', '', s)
            if len(digits) == 8 and digits.isdigit():
                try:
                    return datetime.datetime.strptime(digits, "%Y%m%d").date()
                except Exception:
                    return None
            log.warning(f"Unknown date format in parse_delivery_date: '{s}'")
            return None
        except Exception as e:
            log.warning(f"Exception parsing delivery date '{date_input}': {e}", exc_info=True)
            return None
        


    def safe_float_convert(self, value):
        if value is None:
            return 0.0
        try:
            if isinstance(value, Decimal):
                return float(value)
            return float(value)
        except (TypeError, ValueError):
            return 0.0
        


    def normalize_text(self, text: str) -> str:
        if not text:
            return ""
        text = re.sub(r'[０-９]', lambda m: str(ord(m.group()) - 0xFEE0), text)
        return re.sub(r'\s+', ' ', text.strip())




    def extract_items_from_text(self, text: str):
        """Extract items (H-codes and quantities) from OCR/text input."""
        if not text:
            return []

        text = self.normalize_text(text)
        items = []

        lines = re.split(r'[\n\r]+', text)
        for line in lines:
            line = line.strip()
            if not line or line in ["```", "~~~"]:
                continue
            hcod_match = re.search(
                r'H\s*(\d{6})\s*(?:[:：\-\s]*[x×*]?\s*(?:qty|数量|個|pcs|pc|ea|本|枚|セット|units?)?\s*[:：]?\s*)?(\d+)?',
                line,
                re.IGNORECASE
            )
            if hcod_match:
                code = f"H{hcod_match.group(1)}"
                qty = int(hcod_match.group(2)) if hcod_match.group(2) else 1
                items.append((code, qty))
                continue
            part_qty = re.search(
                r'([\w\s\-–―/／\(\)\[\]\u3000-\u9FFFΩμa-zA-Z0-9\.\+\=\&]{2,80}?)\s*[x×＊*]\s*(\d{1,5})',
                line,
                re.IGNORECASE
            )
            if part_qty:
                name = part_qty.group(1).strip()
                qty = int(part_qty.group(2))
                if len(name) >= 2:
                    items.append((name, qty))
                continue
            fallback = re.search(
                r'([\w\s\-–―/／\(\)\[\]\u3000-\u9FFFΩμa-zA-Z0-9\.\+\=\&]{2,80}?)\s+(\d{1,4})\s*$',
                line,
                re.IGNORECASE
            )
            if fallback:
                name = fallback.group(1).strip()
                qty = int(fallback.group(2))
                if len(name) >= 2:
                    items.append((name, qty))

        return items



    

    def create_quotation(self, customer_id, items_data):
        customer = self.db_session.query(Customer).filter_by(id=customer_id).first()
        if not customer:
            return {"error": "Customer not found"}

        quotation = Quotation(customer_id=customer_id, status='draft')
        self.db_session.add(quotation)
        self.db_session.flush()  
        total_amount = 0.0
        errors = []
        try:
            mhn1_data = self.load_mhn1()
            known_parts = [p.get('HNM', '').strip().upper() for p in mhn1_data.values() if p.get('HNM')]
        except Exception as e:
            log.debug(f"Could not preload MHN1 parts: {e}")
            known_parts = []
        calendar_dict = {}
        try:
            calendar_dict = load_calendar()
        except Exception:
            log.debug("Could not load calendar; get_delivery_date will assume defaults")

        for item_data in items_data:
            hcod_input = item_data.get('hcod') or item_data.get('part_number') or item_data.get('part') or item_data.get('pn')
            qty = item_data.get('qty') or item_data.get('quantity') or item_data.get('q') or 1
            try:
                qty = int(qty)
            except Exception:
                qty = 1
            required_date = item_data.get('required_date')
            if not hcod_input or str(hcod_input).strip() in ["```", "~~~", ""]:
                log.info(f"Skipping invalid OCR artifact: {hcod_input}")
                continue

            corrected_hcod = correct_ocr_code(hcod_input, known_parts)
            lookup_code = corrected_hcod or hcod_input

            product = self.find_product_in_database(lookup_code)
            if not product:
                log.info(f"🔍 {lookup_code} not in DB — attempting Digi-Key lookup...")
                try:
                    from app import search_digikey
                    price, stock, dk_pn = search_digikey(lookup_code)
                except Exception as e:
                    log.warning(f"Digi-Key lookup failed: {e}", exc_info=True)
                    price, stock, dk_pn = None, 0, ""

                if price is not None:
                    try:
                        stock = int(stock or 0)
                    except Exception:
                        stock = 0
                    lead_time_days = 3 if stock > 0 else 14
                    if not isinstance(lead_time_days, int) or lead_time_days < 0:
                        log.warning(f"Invalid lead_time_days from Digi-Key handling: {lead_time_days} -> forcing 0")
                        lead_time_days = 0

                    delivery_date = None
                    try:
                        delivery_date = get_delivery_date(lead_time_days, calendar_dict)
                    except Exception as e:
                        log.warning(f"get_delivery_date failed for Digi-Key item {lookup_code}: {e}", exc_info=True)
                        delivery_date = None
                    if not isinstance(delivery_date, datetime.date):
                        log.debug(f"Digi-Key delivery_date not a date (will set None): {delivery_date}")
                        delivery_date = None
                    try:
                        ident = (dk_pn or lookup_code).upper()
                        ident_clean = re.sub(r'[^A-Z0-9\-\_\.]', '_', ident)

                        dk_hcod = f"DK-{ident_clean}"
                        placeholder = self.db_session.query(Product).filter_by(hcod=dk_hcod).first()
                        if not placeholder:
                            placeholder = Product(
                                hcod=dk_hcod,
                                hnm=str(lookup_code)[:200],
                                description=f"External supplier item (Digi-Key) - source PN: {dk_pn or lookup_code}",
                                category="EXTERNAL",
                                supplier_code="DIGIKEY",
                                unit_cost=0.0,
                                unit_price=float(price) if price is not None else 0.0
                            )
                            self.db_session.add(placeholder)
                            self.db_session.flush()
                        product_id_to_use = placeholder.id
                    except Exception as e:
                        log.exception(f"Failed to create/find placeholder Product for Digi-Key item {lookup_code}: {e}")
                        product_id_to_use = None

                    
                    if not product_id_to_use:
                        try:
                            fallback_prod = Product(
                                hcod=f"UNKNOWN-{int(datetime.datetime.now().timestamp())}",
                                hnm=f"Unknown item ({lookup_code})",
                                description="Fallback placeholder for external item",
                                category="UNKNOWN",
                                supplier_code="UNKNOWN",
                                unit_cost=0.0,
                                unit_price=float(price) if price is not None else 0.0
                            )
                            self.db_session.add(fallback_prod)
                            self.db_session.flush()
                            product_id_to_use = fallback_prod.id
                        except Exception as e:
                            log.exception(f"Fallback placeholder creation also failed for {lookup_code}: {e}")
                            product_id_to_use = None

                    
                    q_item = QuotationItem(
                        quotation_id=quotation.id,
                        product_id=product_id_to_use,
                        quantity=qty,
                        unit_price=price,
                        lead_time_days=lead_time_days,
                        estimated_delivery_date=delivery_date,
                        stock_status=(f"在庫あり ({stock}個) - Digi-Key" if stock > 0 else "在庫なし - Digi-Key"),
                        notes=f"Digi-Key PN: {dk_pn or lookup_code}"
                    )
                    self.db_session.add(q_item)
                    total_amount += self.safe_float_convert(price) * qty
                else:
                   
                    try:
                        placeholder = Product(
                            hcod=f"UNKNOWN-{int(datetime.datetime.now().timestamp())}",
                            hnm=f"Unknown item ({lookup_code})",
                            description="Placeholder product for review",
                            category="UNKNOWN",
                            supplier_code="UNKNOWN",
                            unit_cost=0.0,
                            unit_price=0.0
                        )
                        self.db_session.add(placeholder)
                        self.db_session.flush()
                        product_id = placeholder.id
                    except IntegrityError:
                        self.db_session.rollback()
                        product_id = None
                        log.warning(f"Failed to create placeholder for {lookup_code}", exc_info=True)

                    q_item = QuotationItem(
                        quotation_id=quotation.id,
                        product_id=product_id,
                        quantity=qty,
                        unit_price=0.0,
                        lead_time_days=999,
                        estimated_delivery_date=None,
                        stock_status="要確認 (内部/外部共に見つかりません)"
                    )
                    self.db_session.add(q_item)
                    errors.append(f"Item {lookup_code} not found internally or on Digi-Key.")
                continue

          
            calc_result = {}
            try:
                calc_result = self.inventory_service.calculate_lead_time_and_status(
                    product.hcod,
                    product.supplier_code,
                    qty,
                    customer_delivery_date=required_date
                ) or {}
            except Exception as e:
                log.warning(f"calculate_lead_time_and_status failed for {product.hcod}: {e}", exc_info=True)
                calc_result = {}

          
            lt_raw = calc_result.get("lead_time_days", 0)
            try:
                lead_time_days = int(lt_raw if lt_raw is not None else 0)
            except Exception:
                log.warning(f"Invalid lead_time_days '{lt_raw}' for {product.hcod} -> forcing 0")
                lead_time_days = 0
            if lead_time_days < 0:
                log.warning(f"Negative lead_time_days {lead_time_days} for {product.hcod} -> forcing 0")
                lead_time_days = 0

            
            estimated_delivery_date = None
           
            delivery_date_raw = calc_result.get("delivery_date")
            estimated_delivery_date = self.parse_delivery_date(delivery_date_raw)
            if estimated_delivery_date is None:
               
                try:
                    estimated_delivery_date = get_delivery_date(lead_time_days, calendar_dict)
                except Exception as e:
                    log.warning(f"Fallback get_delivery_date failed for {product.hcod}: {e}", exc_info=True)
                    estimated_delivery_date = None

           
            if not isinstance(estimated_delivery_date, datetime.date):
                log.debug(f"Final estimated_delivery_date invalid for {product.hcod}: {estimated_delivery_date} -> setting None")
                estimated_delivery_date = None

            unit_price = self.safe_float_convert(product.unit_price)

            q_item = QuotationItem(
                quotation_id=quotation.id,
                product_id=product.id,
                quantity=qty,
                unit_price=unit_price,
                lead_time_days=lead_time_days,
                estimated_delivery_date=estimated_delivery_date,
                stock_status=calc_result.get("stock_status", "Unknown")
            )
            self.db_session.add(q_item)
            total_amount += unit_price * qty

        try:
            self.db_session.commit()
            return {
                "success": True,
                "quotation_id": quotation.id,
                "total_amount": total_amount,
                "errors": errors if errors else None
            }
        except IntegrityError:
            self.db_session.rollback()
            log.exception("Database integrity error while creating quotation.")
            return {"error": "Database integrity error occurred while creating quotation."}
        



    def is_quotation_request(self, subject: str, body: str) -> bool:
        full_text = self.normalize_text(f"{subject} {body}")
        return bool(re.search(r'(H\d{6}|見積|お見積|quote|quotation)', full_text, re.IGNORECASE))
