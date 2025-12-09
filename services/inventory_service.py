# backend/services/inventory_service.py
from models import Stock, Product, Supplier, SupplierLeadTime, Calendar
from utils.lead_time_calculator import calculate_lead_time_and_status 
from utils.calendar_utils import get_delivery_date
import logging

log = logging.getLogger(__name__)

class InventoryService:
    def __init__(self, db_session):
        self.db_session = db_session

    def check_stock(self, hcod, requested_qty):
        stock_records = self.db_session.query(Stock).join(Product).filter(Product.hcod == hcod).all()
        total_available = sum([s.actual_quantity for s in stock_records])

        if total_available >= requested_qty:
            status = f"在庫あり ({total_available}個)"
        elif total_available > 0:
            status = f"一部在庫 ({total_available}個)・発注必要"
        else:
            status = "在庫なし (発注必要)"

        return {"available": total_available, "status": status}

    def calculate_lead_time_and_status(self, hcod, supplier_code, requested_qty, customer_delivery_date=None):
        """Wrapper around the utility function to use the service's session."""
        return calculate_lead_time_and_status(
            hcod, supplier_code, requested_qty,
            customer_delivery_date=customer_delivery_date,
            db_session=self.db_session 
        )