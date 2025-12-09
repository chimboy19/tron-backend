# backend/utils/lead_time_calculator.py
from models import Stock, Product, Supplier, SupplierLeadTime, Calendar
from utils.calendar_utils import get_delivery_date,load_calendar
from datetime import datetime

def calculate_lead_time_and_status(hcod, supplier_code, requested_qty, customer_delivery_date=None, db_session=None):
    
    if not db_session:
        from app import db 
        db_session = db.session
    stock_records = db_session.query(Stock).join(Product).filter(Product.hcod == hcod).all()
    total_available_stock = sum([s.actual_quantity for s in stock_records])

    if requested_qty <= total_available_stock:
        lead_time_days = 4
        delivery_date = get_delivery_date(lead_time_days, load_calendar()) 
        stock_status = f"在庫あり ({total_available_stock}個)"
        print(f"Item {hcod} (Qty {requested_qty}) found in stock (Qty {total_available_stock}). Lead time: {lead_time_days} days, Delivery: {delivery_date}")
        return {"lead_time_days": lead_time_days, "delivery_date": delivery_date.isoformat(), "stock_status": stock_status}

   
    supplier = db_session.query(Supplier).filter_by(mcod=supplier_code).first()
    if not supplier:
        
        return {"lead_time_days": 999, "delivery_date": "要問合せ", "stock_status": "要確認"}

    standard_lead_time = supplier.standard_lead_time 
    product_id = db_session.query(Product.id).filter_by(hcod=hcod).scalar()
    supplier_id = supplier.id
    if not product_id or not supplier_id:
        return {"lead_time_days": 999, "delivery_date": "要問合せ", "stock_status": "要確認"}

    historical_record = db_session.query(SupplierLeadTime).filter_by(
        product_id=product_id, supplier_id=supplier_id
    ).order_by(SupplierLeadTime.updated_date.desc()).first() 

   
    if historical_record and historical_record.promised_days and historical_record.promised_days <= 90: 
        lead_time_days = historical_record.promised_days
        print(f"Using historical lead time {lead_time_days} days for {hcod} from supplier {supplier.name} (Record ID: {historical_record.id})")
    else:
        lead_time_days = standard_lead_time
        print(f"Using standard lead time {standard_lead_time} days for supplier {supplier.name}")

    delivery_date = get_delivery_date(lead_time_days, load_calendar())
    print(f"Item {hcod} (Qty {requested_qty}) needs ordering. Supplier {supplier_code}. Lead time: {lead_time_days} days, Delivery: {delivery_date}")

   
    if customer_delivery_date:
        req_date_obj = datetime.strptime(customer_delivery_date, "%Y-%m-%d").date()
        if delivery_date > req_date_obj:
            print(f"Warning: Calculated delivery date {delivery_date} is later than required date {req_date_obj}")
    if total_available_stock > 0:
        stock_status = f"一部在庫 ({total_available_stock}個)・発注必要"
    else:
        stock_status = f"発注必要 (在庫: {total_available_stock}個)"

    print(f"Item {hcod} (Qty {requested_qty}) needs ordering. Supplier {supplier_code}. Lead time: {lead_time_days} days, Delivery: {delivery_date}, Status: {stock_status}")
    return {"lead_time_days": lead_time_days, "delivery_date": delivery_date.isoformat(), "stock_status": stock_status}



