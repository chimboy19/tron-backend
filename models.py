# backend/models.py
from flask_sqlalchemy import SQLAlchemy
from flask_jwt_extended import create_access_token
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime, timedelta
import enum

db = SQLAlchemy()

class User(db.Model):
    __tablename__ = 'users'
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    role = db.Column(db.Enum('admin', 'customer', name='user_roles'), default='customer')
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    last_login = db.Column(db.DateTime)
    notifications = db.relationship('Notification', back_populates='user', lazy='dynamic')
    customer_id = db.Column(db.Integer, db.ForeignKey('customers.id'))
    customer = db.relationship('Customer', back_populates='user')

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)

    def get_auth_token(self):
        
        return create_access_token(identity=self.id, additional_claims={"role": self.role}, expires_delta=timedelta(hours=24))




# --- Core Master Data Models (Based on TREE layouts) ---
class Product(db.Model):
    __tablename__ = 'products'
    id = db.Column(db.Integer, primary_key=True)
    hcod = db.Column(db.String(50), unique=True, nullable=False) # H123456 (from MHN1 layout)
    hnm = db.Column(db.String(255), nullable=False) # 品名/説明 (from MHN1 layout)
    description = db.Column(db.Text) # HNMT (from MHN1 layout)
    category = db.Column(db.String(100)) # HSRS or derived from MHN1
    supplier_code = db.Column(db.String(50), db.ForeignKey('suppliers.mcod'), nullable=False) # MCODD (from MHN1 layout)
    unit_cost = db.Column(db.Numeric(10, 2)) # IRINE (from MHN1 layout)
    unit_price = db.Column(db.Numeric(10, 2)) # @A (from MHN1 layout)
    supplier = db.relationship('Supplier', back_populates='products')
    stocks = db.relationship('Stock', back_populates='product')
    quotation_items = db.relationship('QuotationItem', back_populates='product')
    supplier_lead_times = db.relationship('SupplierLeadTime', back_populates='product')


class Order(db.Model):
    __tablename__ = 'orders'

    id = db.Column(db.Integer, primary_key=True)
    order_number =db.Column(db.String(50), unique=True, nullable=False, index=True)
    customer_id = db.Column(db.Integer, db.ForeignKey('customers.id'), nullable=False)
    status = db.Column(db.String(20), default='confirmed')  # e.g., confirmed, shipped, delivered
    date_created = db.Column(db.DateTime, default=datetime.utcnow)
    total_amount = db.Column(db.Float, default=0.0)

    # Relationships
    customer = db.relationship("Customer", back_populates="orders")
    items = db.relationship("OrderItem", back_populates="order", cascade="all, delete-orphan")

    def __repr__(self):
        return f"<Order {self.order_number}>"
    

class OrderItem(db.Model):
    __tablename__ = 'order_items'

    id = db.Column(db.Integer, primary_key=True)
    order_id = db.Column(db.Integer, db.ForeignKey('orders.id'), nullable=False)
    product_id = db.Column(db.Integer, db.ForeignKey('products.id'), nullable=False)
    quantity = db.Column(db.Integer, nullable=False)
    unit_price = db.Column(db.Float, nullable=False)
    lead_time_days = db.Column(db.Integer, nullable=True)
    estimated_delivery_date = db.Column(db.Date, nullable=True)
    stock_status = db.Column(db.String(50), nullable=True)  
    # Relationships
    order = db.relationship("Order", back_populates="items")
    product = db.relationship("Product")

    def __repr__(self):
        return f"<OrderItem order={self.order_id} product={self.product_id}>"


class Supplier(db.Model):
    __tablename__ = 'suppliers'
    id = db.Column(db.Integer, primary_key=True)
    mcod = db.Column(db.String(10), unique=True, nullable=False) 
    name = db.Column(db.String(255), nullable=False) # MNM (from MMK1 layout)
    address = db.Column(db.Text) # ADRNS, ADR1S, ADR2S (from MMK1 layout)
    contact_person = db.Column(db.String(255)) # MNMT (from MMK1 layout)
    phone = db.Column(db.String(20)) # TELS (from MMK1 layout)
    standard_lead_time = db.Column(db.Integer, default=7) # NEBIP (納入日数) (from MMK1 layout)
    api_config = db.Column(db.JSON) # Store API credentials/endpoints

    products = db.relationship('Product', back_populates='supplier')
    stocks = db.relationship('Stock', back_populates='supplier')
    lead_times = db.relationship('SupplierLeadTime', back_populates='supplier')

class Customer(db.Model):
    __tablename__ = 'customers'
    id = db.Column(db.Integer, primary_key=True)
    ucod = db.Column(db.String(10), unique=True, nullable=False) # 1209 (from MUS1 layout)
    name = db.Column(db.String(255), nullable=False) # UNM (from MUS1 layout)
    address = db.Column(db.Text) # ADRNS, ADR1S, ADR2S (from MUS1 layout)
    phone = db.Column(db.String(20)) # TELS (from MUS1 layout)
    email = db.Column(db.String(255)) # EMAIL (from MUS1 layout)
    terms = db.Column(db.String(50)) # SHCOD (from MUS1 layout)
    user = db.relationship('User', back_populates='customer', uselist=False, cascade="all, delete-orphan")
    quotations = db.relationship('Quotation', back_populates='customer', cascade="all, delete-orphan")
    orders = db.relationship("Order", back_populates="customer", cascade="all, delete-orphan")



# --- Data Models based on TREE layouts ---
class Stock(db.Model):
    __tablename__ = 'stocks'
    id = db.Column(db.Integer, primary_key=True)
    product_id = db.Column(db.Integer, db.ForeignKey('products.id'), nullable=False)
    supplier_id= db.Column(db.Integer, db.ForeignKey('suppliers.id'), nullable=False)
    warehouse_code = db.Column(db.String(10), default='01') # SOKO (from TNZ1 layout)
    actual_quantity = db.Column(db.Integer, default=0) # JZSU (実在庫数) (from TNZ1 layout)
    shelf_quantity = db.Column(db.Integer, default=0) # ZSSU (棚在庫数) (from TNZ1 layout)
    location_floor = db.Column(db.String(10)) # TANAFL (from TNZ1 layout)
    manufacturer = db.Column(db.String(10)) 
    hnm= db.Column(db.String(10)) 
    location_block = db.Column(db.String(10)) # TANABL (from TNZ1 layout)
    location_number = db.Column(db.String(10)) # TANANO (from TNZ1 layout)
    location_stage = db.Column(db.String(10)) # TANAST (from TNZ1 layout)
    product = db.relationship('Product', back_populates='stocks')
    supplier = db.relationship('Supplier', back_populates='stocks')



class Calendar(db.Model):
    __tablename__ = 'calendar'
    id = db.Column(db.Integer, primary_key=True)
    date = db.Column(db.Date, unique=True, nullable=False) # DATE (from DSET layout)
    date_number = db.Column(db.Integer) # DATENO (from DSET layout)
    day_of_week = db.Column(db.String(10)) # YOBI (曜日) (from DSET layout)
    is_holiday = db.Column(db.Boolean, default=False) # YASUMI (休み) (from DSET layout)
    is_shipping_stop = db.Column(db.Boolean, default=False) # NYSTOP (入荷止) (from DSET layout)
    delivery_course = db.Column(db.String(10)) # KOSU (配送コース) (from DSET layout)
    course_date = db.Column(db.Date) # DATE1 (from DSET layout)
    reverse_course_date = db.Column(db.Date) # DATE2 (from DSET layout)
    course_next_day = db.Column(db.Date) # DATE3 (from DSET layout)
    logistics_holiday = db.Column(db.Boolean, default=False) # DSETXA (物流休み) (from DSET layout)
    

class SupplierLeadTime(db.Model):
    __tablename__ = 'supplier_lead_times'
    id = db.Column(db.Integer, primary_key=True)
    product_id = db.Column(db.Integer, db.ForeignKey('products.id'), nullable=False) # Implicitly links HCOD via foreign key
    supplier_id = db.Column(db.Integer, db.ForeignKey('suppliers.id'), nullable=False) # MCOD (from RAS1 layout)
    order_number = db.Column(db.String(20)) # TNO (旧VER注番) (from RAS1 layout)
    promised_days = db.Column(db.Integer) # ASDAYS (仕入先回答納期) (from RAS1 layout)
    quantity = db.Column(db.Integer) # ASSU (回答数量) (from RAS1 layout)
    supplier_invoice_number = db.Column(db.String(30)) # SIDEN (仕入先・伝票番号) (from RAS1 layout)
    comment = db.Column(db.Text) # ASCMT (コメント) (from RAS1 layout)
    updated_date = db.Column(db.Date) # UPDDAT (更新日) (from RAS1 layout)
    updated_time = db.Column(db.Time) # UPDTIM (更新時刻) (from RAS1 layout)
    control_a = db.Column(db.String(1)) # RAS1XA (制御項目Ａ) (from RAS1 layout)
    control_b = db.Column(db.String(1)) # RAS1XB (制御項目Ｂ) (from RAS1 layout)
    control_c = db.Column(db.String(1)) # RAS1XC (制御項目Ｃ) (from RAS1 layout)
    control_d = db.Column(db.String(1)) # RAS1XD (制御項目Ｄ) (from RAS1 layout)
    control_e = db.Column(db.Integer) # RAS1XE (制御項目Ｅ) (from RAS1 layout)
    control_f = db.Column(db.String(10)) # RAS1XF (制御項目Ｆ) (from RAS1 layout)
    control_g = db.Column(db.Integer) # RAS1XG (制御項目Ｇ) (from RAS1 layout)
    control_h = db.Column(db.Integer) # RAS1XH (制御項目Ｈ) (from RAS1 layout)
    control_i = db.Column(db.Integer) # RAS1XI (制御項目Ｉ) (from RAS1 layout)
    control_j = db.Column(db.Integer) # RAS1XJ (制御項目Ｊ) (from RAS1 layout)
    control_k = db.Column(db.Integer) # RAS1XK (制御項目Ｋ) (from RAS1 layout)
    customer_delivery_date = db.Column(db.Date) # NODAYU (得意先納期) (from RAS1 layout)
    free_order_number = db.Column(db.String(15)) # TNOF (フリー注番) (from RAS1 layout)

    product = db.relationship('Product', back_populates='supplier_lead_times')
    supplier = db.relationship('Supplier', back_populates='lead_times')

class FAXHistory(db.Model): # Represents RAF1 layout conceptually
    __tablename__ = 'fax_history'
    id = db.Column(db.Integer, primary_key=True)
    delete_flag = db.Column(db.String(1)) # RAF1D (削除) (from RAF1 layout)
    fax_day = db.Column(db.Date) # FAXDAY (送信日) (from RAF1 layout)
    fax_time = db.Column(db.Time) # FAXTIM (送信時間) (from RAF1 layout)
    job_number = db.Column(db.Integer) # JOBNBR (ＪＯＢ№) (from RAF1 layout)
    spool_number_print = db.Column(db.Integer) # SPLNO (スプール№印刷用) (from RAF1 layout)
    spool_number_log = db.Column(db.Integer) # LOGSPL (スプール№ ログ用) (from RAF1 layout)
    supplier_code = db.Column(db.String(10), db.ForeignKey('suppliers.mcod')) # MCOD (仕入先コード) (from RAF1 layout)
    order_number = db.Column(db.String(20)) # ODRNO (注文書送り先№) (from RAF1 layout)
    resend_count = db.Column(db.Integer, default=0) # FAXSU (再送信回数) (from RAF1 layout)
    printer_output_flag = db.Column(db.String(1)) # FLAG (ﾌﾟﾘﾝﾀｰ出力 ﾌﾗｸﾞ) (from RAF1 layout)
    resend_flag = db.Column(db.String(1)) # FLAGS (再送信 ﾌﾗｸﾞ) (from RAF1 layout)
    send_number = db.Column(db.String(20)) # SNDNO (送信№) (from RAF1 layout)
    printer_flag_p = db.Column(db.String(1)) # FLAGP (ﾌﾟﾘﾝﾀｰ出力 ﾌﾗｸﾞ) (from RAF1 layout)
    fax_type_flag = db.Column(db.String(1)) # FLAG1 (ＦＡＸの種類) (from RAF1 layout)
    control_a = db.Column(db.String(1)) # RAF1XA (制御項目Ａ) (from RAF1 layout)
    control_b = db.Column(db.String(1)) # RAF1XB (制御項目Ｂ) (from RAF1 layout)
    control_c = db.Column(db.String(1)) # RAF1XC (制御項目Ｃ) (from RAF1 layout)
    control_d = db.Column(db.String(1)) # RAF1XD (制御項目Ｄ) (from RAF1 layout)
    control_e = db.Column(db.String(1)) # RAF1XE (制御項目Ｅ) (from RAF1 layout)
    job_number_inquiry = db.Column(db.String(10)) # JOBNBA (ＪＯＢ№（照会）) (from RAF1 layout)
    spool_number_sent = db.Column(db.String(10)) # SPLNOA (スプール№（送信）) (from RAF1 layout)
    job = db.Column(db.String(10)) # JOB (ジョブ) (from RAF1 layout)
    user = db.Column(db.String(10)) # USER (ユーザー) (from RAF1 layout)
    fax_management_number = db.Column(db.String(10)) # UFSNO (ＦＡＸ管理№) (from RAF1 layout)
    customer_code = db.Column(db.String(10), db.ForeignKey('customers.ucod')) # UCOD (得意先コード) (from RAF1 layout)
    destination_code = db.Column(db.String(10)) # UFNO (送り先コード) (from RAF1 layout)
    sending_department = db.Column(db.String(10)) # GPCOD (送信部署) (from RAF1 layout)
    sender_code = db.Column(db.String(10)) # WCODF (送信者コード) (from RAF1 layout)
    destination_type = db.Column(db.String(1)) # UFMUC (送信先識別：仕／得) (from RAF1 layout)

    supplier = db.relationship('Supplier')
    customer = db.relationship('Customer')

class TNZ2Record(db.Model): 
    __tablename__ = 'tnz2_records'
    id = db.Column(db.Integer, primary_key=True)
    delete_flag = db.Column(db.String(1)) # TNZ2D (削除コード) (from TNZ2 layout)
    document_number = db.Column(db.String(10)) # DENNO (伝票番号) (from TNZ2 layout)
    line_number = db.Column(db.String(10)) # GYONO (行番号) (from TNZ2 layout)
    customer_code = db.Column(db.String(10), db.ForeignKey('customers.ucod')) # UCOD (得意先コード) (from TNZ2 layout)
    supplier_code = db.Column(db.String(10), db.ForeignKey('suppliers.mcod')) # MCOD (仕入先コード) (from TNZ2 layout)
    item_description = db.Column(db.Text) # HNAME (品名／注残) (from TNZ2 layout)
    quantity = db.Column(db.Integer) # SURYO (数量) (from TNZ2 layout)
    unit_price = db.Column(db.Numeric(10, 2)) # TANKA (単価) (from TNZ2 layout)
    standard_cost = db.Column(db.Numeric(10, 2)) # IRINE (標準入値) (from TNZ2 layout)
    customer_delivery_date = db.Column(db.Date) # NODAYU (得意先納期) (from TNZ2 layout)
    shipment_date = db.Column(db.Date) # SYDAY (出庫日) (from TNZ2 layout)
    customer_name = db.Column(db.String(32)) # UNMX (諸口得意先名) (from TNZ2 layout)
    model_2 = db.Column(db.String(32)) # HNM2 (形式２) (from TNZ2 layout)
    estimate_number = db.Column(db.String(9)) # MNO (見積番号) (from TNZ2 layout)
    job_number = db.Column(db.String(6)) # KNO (工番) (from TNZ2 layout)
    estimate_number_detail = db.Column(db.String(12)) # MNO2 (見積番号（明細）) (from TNZ2 layout)
    job_number_detail = db.Column(db.String(9)) # KNO2 (工番（明細）) (from TNZ2 layout)
    sales_category = db.Column(db.String(1)) # BURIK (売上区分) (from TNZ2 layout)
    group_identifier = db.Column(db.String(1)) # BGPSY (ｸﾞﾙｰﾌﾟ識別営／設) (from TNZ2 layout)
    product_code = db.Column(db.String(10), db.ForeignKey('products.hcod')) # HCOD (品番) (from TNZ2 layout)
    warehouse_code = db.Column(db.String(10)) # SOKO (倉庫区分) (from TNZ2 layout)
    inventory_level = db.Column(db.String(1)) # ZLVL (在庫レベル) (from TNZ2 layout)
    location_floor = db.Column(db.String(1)) # TANAFL (棚フロア) (from TNZ2 layout)
    location_block = db.Column(db.String(1)) # TANABL (棚ブロック) (from TNZ2 layout)
    location_number = db.Column(db.String(2)) # TANANO (棚番) (from TNZ2 layout)
    location_stage = db.Column(db.String(2)) # TANAST (棚段) (from TNZ2 layout)
    delivery_code = db.Column(db.String(5)) # HAISO (配送コード) (from TNZ2 layout)
    storage_shelf = db.Column(db.String(6)) # KTANA (格納棚番) (from TNZ2 layout)

    customer = db.relationship('Customer')
    supplier = db.relationship('Supplier')
    product = db.relationship('Product')

# --- Application Business Models ---
class Quotation(db.Model):
    __tablename__ = 'quotations'
    id = db.Column(db.Integer, primary_key=True)
    customer_id = db.Column(db.Integer, db.ForeignKey('customers.id'), nullable=False)
    date_created = db.Column(db.DateTime, default=datetime.utcnow)
    status = db.Column(db.String(20), default='draft') # draft, sent, accepted, etc.
    notes = db.Column(db.Text) # For internal notes

    customer = db.relationship('Customer', back_populates='quotations')
    items = db.relationship('QuotationItem', back_populates='quotation', cascade='all, delete-orphan')

class QuotationItem(db.Model):
    __tablename__ = 'quotation_items'
    id = db.Column(db.Integer, primary_key=True)
    quotation_id = db.Column(db.Integer, db.ForeignKey('quotations.id'), nullable=False)
    product_id = db.Column(db.Integer, db.ForeignKey('products.id'), nullable=False)
    quantity = db.Column(db.Integer, nullable=False)
    unit_price = db.Column(db.Numeric(10, 2)) # Calculated price
    lead_time_days = db.Column(db.Integer) # Calculated lead time
    estimated_delivery_date = db.Column(db.Date) # Calculated delivery date
    stock_status = db.Column(db.String(50)) # In Stock, Out of Stock, Partial, etc.
    notes = db.Column(db.Text) # For item-specific notes

    quotation = db.relationship('Quotation', back_populates='items')
    product = db.relationship('Product', back_populates='quotation_items')

class PurchaseOrder(db.Model):
    __tablename__ = 'purchase_orders'
    id = db.Column(db.Integer, primary_key=True)
    supplier_id = db.Column(db.Integer, db.ForeignKey('suppliers.id'), nullable=False)
    date_created = db.Column(db.DateTime, default=datetime.utcnow)
    status = db.Column(db.String(20), default='pending') # pending, sent, confirmed, shipped, received
    notes = db.Column(db.Text)

    supplier = db.relationship('Supplier')
    items = db.relationship('PurchaseOrderItem', back_populates='po', cascade='all, delete-orphan')

class PurchaseOrderItem(db.Model):
    __tablename__ = 'purchase_order_items'
    id = db.Column(db.Integer, primary_key=True)
    po_id = db.Column(db.Integer, db.ForeignKey('purchase_orders.id'), nullable=False)
    product_id = db.Column(db.Integer, db.ForeignKey('products.id'), nullable=False)
    quantity = db.Column(db.Integer, nullable=False)
    unit_price = db.Column(db.Numeric(10, 2))
    confirmed_delivery_date = db.Column(db.Date) 
    po = db.relationship('PurchaseOrder', back_populates='items')
    product = db.relationship('Product')

class IncomingQuotationRequest(db.Model): 
    __tablename__ = 'incoming_quotation_requests'
    id = db.Column(db.Integer, primary_key=True)
    subject = db.Column(db.String(500), nullable=False) 
    body = db.Column(db.Text)
    sender = db.Column(db.String(255), nullable=False) 
    received_date = db.Column(db.DateTime) 
    status = db.Column(db.String(20), default='pending') 
    items_data = db.Column(db.JSON) 
    customer_name = db.Column(db.String(255)) 
    customer_tel = db.Column(db.String(20)) 
    customer_ucod = db.Column(db.String(10), db.ForeignKey('customers.ucod')) 
    customer_id = db.Column(db.Integer, db.ForeignKey('customers.id'))

    notes = db.Column(db.Text)

    customer = db.relationship('Customer', foreign_keys=[customer_id])

    
    customer_by_ucod = db.relationship(
        'Customer',
        foreign_keys=[customer_ucod],
        viewonly=True
    )




class Notification(db.Model):
    __tablename__ = 'notification'
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False, index=True)
    type = db.Column(db.String(50), nullable=False)  
    related_id = db.Column(db.Integer, nullable=True) 
    message = db.Column(db.String(200), nullable=False)
    is_read = db.Column(db.Boolean, default=False, index=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, index=True) 
    user = db.relationship('User', back_populates='notifications')