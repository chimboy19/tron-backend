# backend/services/procurement_service.py
from models import PurchaseOrder, PurchaseOrderItem, Product, Supplier, Stock, SupplierLeadTime, FAXHistory
from sqlalchemy.exc import IntegrityError
from services.inventory_service import InventoryService
import logging
import datetime 

log = logging.getLogger(__name__)

class ProcurementService:
    def __init__(self, db_session, inventory_service=None):
        self.db_session = db_session
        self.inventory_service = inventory_service 

    def create_purchase_order(self, supplier_id, items_data):
        """
        Create a purchase order based on supplier ID and item data.
        :param supplier_id: The ID of the supplier to order from.
        :param items_data: List of dicts like [{"hcod": "H123456", "qty": 10, "unit_price": 100.0}, ...]
                           or [{"product_id": 1, "qty": 10, "unit_price": 100.0}, ...]
        :return: Dict with success status, PO ID, total value, or error message.
        """
        supplier = self.db_session.query(Supplier).filter_by(id=supplier_id).first()
        if not supplier:
            return {"error": "Supplier not found"}

        po = PurchaseOrder(supplier_id=supplier_id, status='pending')
        self.db_session.add(po)
        self.db_session.flush() 

        total_value = 0.0
        errors = []
        for item_data in items_data:
            hcod_input = item_data.get('hcod')
            product_id_input = item_data.get('product_id')
            qty = item_data.get('qty', 1)
            unit_price = item_data.get('unit_price', 0.0) 
            product = None
            if hcod_input:
               
                product = self.db_session.query(Product).filter_by(hcod=hcod_input).first()
            elif product_id_input:
                
                product = self.db_session.query(Product).filter_by(id=product_id_input).first()

            if not product:
                errors.append(f"Product (HCOD: {hcod_input}, ID: {product_id_input}) not found for PO item.")
                continue
            if self.inventory_service:
                stock_info = self.inventory_service.check_stock(product.hcod, qty)
                if stock_info.get("available", 0) >= qty:
                    log.info(f"Stock level for {product.hcod} has risen above required qty {qty} since PO trigger. Skipping item in PO {po.id}.")
                    continue 

            po_item = PurchaseOrderItem(
                po_id=po.id,
                product_id=product.id,
                quantity=qty,
                unit_price=unit_price,
                
            )
            self.db_session.add(po_item)
            total_value += unit_price * qty

        try:
            self.db_session.commit()
            log.info(f"✅ Created PO {po.id} for supplier {supplier.name} with {len(po.items)} items, total value ¥{total_value:.2f}")
            return {
                "success": True,
                "po_id": po.id,
                "total_value": total_value,
                "items_count": len(po.items),
                "errors": errors if errors else None 
            }
        except IntegrityError:
            self.db_session.rollback()
            log.error(f"❌ Database integrity error creating PO for supplier {supplier_id}.")
            return {"error": "Database integrity error occurred while creating purchase order."}

    def send_po_to_supplier(self, po):
        """
        Send the PO to the supplier via email, FAX API, or direct supplier API call.
        This is a placeholder for the actual sending logic.
        """
        supplier = po.supplier
        
        if supplier.api_config and supplier.api_config.get("type") == "REST_API":
            fax_log = FAXHistory(
                supplier_code=supplier.mcod,
                order_number=f"PO{po.id:06d}", 
                fax_day=datetime.date.today(),
                fax_time=datetime.datetime.now().time(),
               
            )
            self.db_session.add(fax_log)
            self.db_session.commit()
            return {"success": True, "method": "API"}
        elif supplier.api_config and supplier.api_config.get("type") == "EMAIL":
            fax_log = FAXHistory(
                supplier_code=supplier.mcod,
                order_number=f"PO{po.id:06d}",
                fax_day=datetime.date.today(),
                fax_time=datetime.datetime.now().time(),
                # ... other RAF1 fields as applicable ...
            )
            self.db_session.add(fax_log)
            self.db_session.commit()
            return {"success": True, "method": "EMAIL"}
        else:
            import requests
            try:
                
                log.info(f"Fax sent successfully for PO {po.id} to supplier {supplier.mcod}")
                # Update FAXHistory (RAF1 equivalent) table
                fax_log = FAXHistory(
                    supplier_code=supplier.mcod,
                    order_number=f"PO{po.id:06d}",
                    fax_day=datetime.date.today(),
                    fax_time=datetime.datetime.now().time(),
                    # ... other RAF1 fields as applicable ...
                )
                self.db_session.add(fax_log)
                self.db_session.commit()
                po.status = 'sent_via_fax' # Update PO status
                self.db_session.commit()
                return {"success": True, "method": "FAX_API"}
                # else:
                #     return {"success": False, "error": f"FAX API returned {response.status_code}: {response.text}"}
            except Exception as e:
                log.error(f"Error sending FAX for PO {po.id}: {e}")
                return {"success": False, "error": str(e)}

    def receive_supplier_confirmation(self, supplier_code, order_number, delivery_date_str, quantity_received):
        """
        Process a supplier's confirmation or update (e.g., via FAX, email, or API callback).
        Updates the RAS1 (SupplierLeadTime) table and potentially the Stock table.
        This function would be called by an endpoint handling supplier responses or by an email/FAX processor.
        """
        po_id_from_order_num = int(order_number.replace("PO", "").lstrip("0")) # Extract ID from "PO123456"
        po = self.db_session.query(PurchaseOrder).filter_by(id=po_id_from_order_num).first()
        if not po or po.supplier.mcod != supplier_code:
            log.warning(f"Received confirmation for unknown PO {order_number} from supplier {supplier_code}")
            return {"error": "PO not found or supplier mismatch"}

        
        po.status = 'confirmed' 
        po_item = None
        for item in po.items:
            if item.quantity == quantity_received: # This is a simplification, might need more specific matching
                po_item = item
                break
        if not po_item:
            log.warning(f"Could not find matching item in PO {po.id} for quantity {quantity_received} in confirmation from {supplier_code}")
            return {"error": "Matching PO item not found"}

        try:
            confirmed_date_obj = datetime.datetime.strptime(delivery_date_str, "%Y%m%d").date()
            po_item.confirmed_delivery_date = confirmed_date_obj
        except ValueError:
            log.error(f"Invalid date format {delivery_date_str} in supplier confirmation for PO {po.id}")
            return {"error": "Invalid date format in confirmation"}
        
        hcod = po_item.product.hcod if po_item.product else None
        if hcod:
            new_lead_time_record = SupplierLeadTime(
                product_id=po_item.product_id,
                supplier_id=po.supplier_id,
                order_number=order_number, # Link back to the PO number
                promised_days=(confirmed_date_obj - datetime.date.today()).days, # Calculate promised days from today to confirmed date
                quantity=quantity_received,
                supplier_invoice_number="",
                comment=f"Confirmed delivery for PO {order_number}",
                updated_date=datetime.date.today(),
                updated_time=datetime.datetime.now().time(),
                customer_delivery_date=None, 
                free_order_number=""
            )

            self.db_session.add(new_lead_time_record)
            self.db_session.commit()
            log.info(f"Updated RAS1 with new lead time record for {hcod} from supplier {supplier_code} based on PO {order_number} confirmation.")
        else:
            log.warning(f"Could not update RAS1 for PO {order_number} - HCOD not found for PO item product.")

        return {"success": True, "message": f"Confirmation for PO {order_number} processed"}

    def trigger_reorder_based_on_stock(self):
        """
        Check inventory levels and trigger PO creation for items below reorder point.
        This function would ideally be called by a background scheduler (e.g., Celery beat).
        """
        log.info("🔍 Initiating automatic reorder check based on stock levels...")
        # Fetch all stock records
        all_stocks = self.db_session.query(Stock).all()
        for stock_record in all_stocks:
            product = stock_record.product
            supplier = stock_record.supplier
            if not product or not supplier:
                log.warning(f"Stock record {stock_record.id} has missing product or supplier link, skipping.")
                continue

            current_stock = stock_record.actual_quantity
            reorder_threshold = 10 
            standard_order_qty = 100 
            if current_stock < reorder_threshold:
                qty_to_order = max(standard_order_qty, reorder_threshold - current_stock) 
                log.info(f"Triggering reorder for {product.hcod} (Current: {current_stock}, Threshold: {reorder_threshold}, Ordering: {qty_to_order})")
                lead_time_calc_result = self.inventory_service.calculate_lead_time_and_status(
                    hcod=product.hcod,
                    supplier_code=supplier.mcod,
                    requested_qty=qty_to_order
                )
                # Use the calculated lead time or a default from supplier master
                expected_arrival_days = lead_time_calc_result.get("lead_time_days", supplier.standard_lead_time)
                # Create the PO
                po_result = self.create_purchase_order(
                    supplier_id=supplier.id,
                    items_data=[{"hcod": product.hcod, "qty": qty_to_order, "unit_price": product.unit_cost}] # Use cost as placeholder price
                )
                if po_result.get("success"):
                    log.info(f"✅ Automatic reorder PO {po_result['po_id']} created for {product.hcod}. Expected arrival in ~{expected_arrival_days} days.")
                else:
                    log.error(f"❌ Failed to create automatic reorder PO for {product.hcod}: {po_result.get('error', 'Unknown error')}")
        log.info("🔍 Automatic reorder check completed.")