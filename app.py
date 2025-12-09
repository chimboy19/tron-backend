
import os
import re
import time
# import csv

import logging
import tempfile

import requests
import datetime
from datetime import timedelta
from fpdf import FPDF, XPos, YPos
from dotenv import load_dotenv
from openai import OpenAI
from difflib import get_close_matches
from models import db, Product, Supplier, Customer, Stock,Notification, Calendar, SupplierLeadTime, Quotation, QuotationItem, PurchaseOrder, PurchaseOrderItem, FAXHistory, TNZ2Record, User, IncomingQuotationRequest,Order,OrderItem
from utils.xlsx_loader import load_initial_data
from utils.lead_time_calculator import calculate_lead_time_and_status
from utils.calendar_utils import load_calendar, is_business_day, add_business_days, get_delivery_date
from config import Config
from flask import Flask, request, jsonify, send_file
from flask_sqlalchemy import SQLAlchemy
from flask_jwt_extended import JWTManager, create_access_token, jwt_required, get_jwt_identity ,get_jwt
from sqlalchemy import func
import httpx
from pdf2image import convert_from_bytes
from werkzeug.utils import secure_filename
import threading
from flask_migrate import Migrate
from io import BytesIO
from fpdf import FPDF
from flask_cors import CORS
from sqlalchemy.exc import NoResultFound





load_dotenv()

logging.basicConfig(level=logging.INFO)
log = logging.getLogger(__name__)

# ==========================
# 🔧 FLASK APPLICATION SETUP
# ==========================
app = Flask(__name__)
app.config.from_object(Config)
CORS(app, origins=["http://localhost:5173"])
db.init_app(app)
jwt = JWTManager(app)
migrate = Migrate(app, db)

# Global service instances
inventory_service = None
quotation_service = None
procurement_service = None

# ==========================
# 🔌 DIGI-KEY INTEGRATION
# ==========================
_DIGIKEY_TOKEN = None

def _get_digikey_token():
    global _DIGIKEY_TOKEN
    if _DIGIKEY_TOKEN:
        return _DIGIKEY_TOKEN
    token_url = "https://api.digikey.com/v1/oauth2/token"
    payload = {
        "client_id": Config.DIGIKEY_CLIENT_ID,
        "client_secret": Config.DIGIKEY_CLIENT_SECRET,
        "grant_type": "client_credentials"
    }
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    try:
        log.info("🔑 Requesting Digi-Key token via httpx...")
        with httpx.Client(timeout=20) as client:
            res = client.post(token_url, data=payload, headers=headers)
            if res.status_code != 200:
                log.warning(f" Digi-Key token fetch failed [{res.status_code}]: {res.text[:300]}")
                return None
            data = res.json()
            _DIGIKEY_TOKEN = data.get("access_token")
            if _DIGIKEY_TOKEN:
                log.info("✅ Digi-Key token fetched successfully (httpx).")
            else:
                log.warning("Digi-Key token response had no access_token field.")
            return _DIGIKEY_TOKEN
    except Exception as e:
        log.warning(f" Digi-Key token fetch failed via httpx: {e}", exc_info=True)
        return None


def search_digikey(part_number: str):
    try:
        if not part_number:
            return None, 0, ""
        part_number = part_number.replace(" ", "").strip().upper()
        token = _get_digikey_token()
        if not token:
            log.debug("No Digi-Key token available - skipping Digi-Key search.")
            return None, 0, ""
        from urllib.parse import quote
        encoded = quote(part_number.strip())
        url = f"https://api.digikey.com/products/v4/search/{encoded}/pricing?limit=5&inStock=false&excludeMarketplace=true"
        headers = {
            "Authorization": f"Bearer {token}",
            "X-DIGIKEY-Client-Id": Config.DIGIKEY_CLIENT_ID,
            "X-DIGIKEY-Locale-Site": "JP",
            "X-DIGIKEY-Locale-Language": "EN",
            "X-DIGIKEY-Locale-Currency": "JPY",
            "Accept": "application/json"
        }
        res = requests.get(url, headers=headers, timeout=10)
        if res.status_code == 401:
            log.info("Digi-Key token unauthorized; refreshing token and retrying.")
            global _DIGIKEY_TOKEN
            _DIGIKEY_TOKEN = None
            token = _get_digikey_token()
            if not token:
                return None, 0, ""
            headers["Authorization"] = f"Bearer {token}"
            res = requests.get(url, headers=headers, timeout=10)
        if res.status_code != 200:
            log.warning(f"Digi-Key search returned status {res.status_code}: {res.text}")
            return None, 0, ""
        data = res.json()
        product_pricings = data.get("ProductPricings") or data.get("Products") or []
        if not product_pricings:
            return None, 0, ""
        p0 = product_pricings[0] if isinstance(product_pricings, list) else product_pricings
        price = None
        stock = 0
        dk_pn = ""
        if isinstance(p0, dict):
            dk_pn = p0.get("ManufacturerProductNumber") or ""
            stock = p0.get("QuantityAvailable") or 0
            try:
                variations = p0.get("ProductVariations") or []
                if variations:
                    std = variations[0].get("StandardPricing") or []
                    if std:
                        price = float(std[0].get("UnitPrice") or 0)
                if price is None:
                    std_top = p0.get("StandardPricing") or []
                    if std_top:
                        price = float(std_top[0].get("UnitPrice") or 0)
            except Exception:
                price = None
            if price is None:
                nested = p0.get("Product") or {}
                if nested:
                    variations = nested.get("ProductVariations") or []
                    if variations:
                        std = variations[0].get("StandardPricing") or []
                        if std:
                            price = float(std[0].get("UnitPrice") or 0)
        try:
            stock = int(stock or 0)
        except Exception:
            stock = 0
        if price is not None:
            try:
                price = float(price)
            except Exception:
                price = None
        return (price if price is not None else None, stock, dk_pn or part_number)
    except Exception as e:
        log.warning(f"Digi-Key search error for '{part_number}': {e}", exc_info=True)
        return None, 0, ""



# to create initail admin  , one time
@app.route('/api/auth/create_initial_admin', methods=['POST'])
def create_initial_admin():
    
    if User.query.filter_by(role='admin').first():
        log.warning(" Initial admin creation attempted but admin already exists")
        return jsonify({
            "success": False, 
            "message": "Initial admin already exists. Use protected endpoint for additional admins."
        }), 400
    
    data = request.get_json()
    username = data.get('username')
    password = data.get('password')
    email = data.get('email')
    
    if not username or not password:
        return jsonify({"success": False, "message": "Username and password are required"}), 400
    
   
    if User.query.filter_by(username=username).first():
        return jsonify({"success": False, "message": "Username already exists"}), 400
    
    try:
        new_user = User(
            username=username,
            email=email,
            role='admin'
        )
        new_user.set_password(password)
        
        db.session.add(new_user)
        db.session.commit()
        
        log.info(f"🎉 INITIAL ADMIN CREATED: '{username}' - SECURE THIS ENDPOINT!")
        
        return jsonify({
            "success": True, 
            "message": "Initial admin user created successfully. DISABLE THIS ENDPOINT IN PRODUCTION.",
            "user_id": new_user.id,
            "security_warning": "DISABLE /api/auth/create_initial_admin IN PRODUCTION"
        }), 201
        
    except Exception as e:
        db.session.rollback()
        log.error(f"Error creating initial admin user: {e}")
        return jsonify({"success": False, "message": "Error creating initial admin"}), 500




# ==========================
# 🔐 AUTHENTICATION ROUTES
# ==========================


@app.route('/api/auth/login', methods=['POST'])
def login():
    username = request.json.get('username', None)
    password = request.json.get('password', None)

    user = User.query.filter_by(username=username).first()

    if user and user.check_password(password) and user.role == 'admin':
        # Identity MUST be a string
        access_token = create_access_token(
            identity=str(user.id),
            additional_claims={"role": user.role}
        )

        return jsonify({
            "success": True,
            "access_token": access_token,
            "user_role": user.role
        }), 200

    return jsonify({
        "success": False,
        "message": "Bad username or password"
    }), 401


#  for admin to register new admin users
@app.route('/api/auth/register_admin', methods=['POST'])
@jwt_required()
def register_admin_protected():
    """Register a new admin user (protected - requires existing admin)"""
    current_user_id = get_jwt_identity()
    current_user = User.query.get(current_user_id)
    
    if not current_user or current_user.role != 'admin':
        return jsonify({"success": False, "message": "Admin access required"}), 403
    
    data = request.get_json()
    username = data.get('username')
    password = data.get('password')
    email = data.get('email')
    
    if not username or not password:
        return jsonify({"success": False, "message": "Username and password are required"}), 400
    if User.query.filter_by(username=username).first():
        return jsonify({"success": False, "message": "Username already exists"}), 400
    
    if User.query.filter_by(email=email).first():
        return jsonify({"success": False, "message": "Email already exists"}), 400
    
    try:
        new_user = User(
            username=username,
            email=email,
            role='admin'
        )
        new_user.set_password(password)
        
        db.session.add(new_user)
        db.session.commit()
        log.info(f"🆕 Admin user '{username}' created by admin '{current_user.username}' (ID: {current_user.id})")
        
        return jsonify({
            "success": True, 
            "message": "Admin user created successfully",
            "user_id": new_user.id
        }), 201
        
    except Exception as e:
        db.session.rollback()
        log.error(f"Error creating admin user: {e}")
        return jsonify({"success": False, "message": "Error creating user"}), 500


# register customer
@app.route('/api/auth/register_customer', methods=['POST'])
def register_customer():
    """Register a new customer user"""
    data = request.get_json()
    username = data.get('username')
    password = data.get('password')
    email = data.get('email')
    customer_name = data.get('customer_name')
    phone = data.get('phone', '')
    address = data.get('address', '')
    
    if not all([username, password, email, customer_name]):
        return jsonify({
            "success": False, 
            "message": "Username, password, email, and customer_name are required"
        }), 400
    
    if User.query.filter_by(username=username).first():
        return jsonify({"success": False, "message": "Username already exists"}), 400
    
    if User.query.filter_by(email=email).first():
        return jsonify({"success": False, "message": "Email already exists"}), 400
    
    try:
        from utils.xlsx_loader import generate_new_ucod
        existing_ucods = {c.ucod for c in Customer.query.all()}
        new_ucod = generate_new_ucod(existing_ucods)
        
        new_customer = Customer(
            ucod=new_ucod,
            name=customer_name,
            email=email,
            phone=phone,
            address=address
        )
        
        new_user = User(
            username=username,
            email=email,
            role='customer'
        )
        new_user.set_password(password)
        new_user.customer = new_customer
        
        db.session.add(new_customer)
        db.session.add(new_user)
        db.session.commit()
        
        return jsonify({
            "success": True, 
            "message": "Customer account created successfully",
            "customer_id": new_customer.id,
            "user_id": new_user.id,
            "ucod": new_ucod
        }), 201
        
    except Exception as e:
        db.session.rollback()
        log.error(f"Error creating customer user: {e}")
        return jsonify({"success": False, "message": "Error creating customer account"}), 500
    


@app.route('/api/auth/check_username/<username>', methods=['GET'])
def check_username_availability(username):
    user = User.query.filter_by(username=username).first()
    return jsonify({"available": user is None})



@app.route('/api/auth/check_email/<email>', methods=['GET'])
def check_email_availability(email):
    user = User.query.filter_by(email=email).first()
    return jsonify({"available": user is None})



#customer login
@app.route('/api/customer_portal/login', methods=['POST'])
def customer_portal_login():
    email = request.json.get('email', None)
    password = request.json.get('password', None)

    user = User.query.join(Customer).filter(Customer.email == email).first()

    if user and user.role == 'customer' and user.check_password(password):
        access_token = create_access_token(
            identity=str(user.id),
            additional_claims={
                "role": user.role,
                "customer_id": user.customer.id
            }
        )

        return jsonify({
            "success": True,
            "access_token": access_token,
            "customer_name": user.customer.name,
            "customer_id": user.customer.id
        })

    else:
        return jsonify({"success": False, "message": "Invalid email or password"}), 401


# ==========================
# 📊 DASHBOARD ROUTES
# ==========================
@app.route('/api/dashboard/stats', methods=['GET'])
@jwt_required()
def get_dashboard_stats():
    current_user_id = get_jwt_identity()
    user = User.query.get(current_user_id)
    if not user or user.role != 'admin':
        return jsonify({"msg": "Admin access required"}), 403
    total_products = Product.query.count()
    total_customers = Customer.query.count()
    pending_quotations = Quotation.query.filter_by(status='draft').count()
    pending_email_requests = IncomingQuotationRequest.query.filter_by(status='pending').count()
    total_stocks = db.session.query(func.sum(Stock.actual_quantity)).scalar() or 0
    return jsonify({
        "total_products": total_products,
        "total_customers": total_customers,
        "pending_quotations": pending_quotations,
        "pending_email_requests": pending_email_requests,
        "total_stock_quantity": total_stocks
    })



@app.route('/api/notifications', methods=['GET'])
@jwt_required()
def get_notifications():
    current_user_id = int(get_jwt_identity())
    user = User.query.get(current_user_id)
    if not user:
        return jsonify({"error": "Unauthorized"}), 403
    notifications = Notification.query.filter_by(user_id=user.id)\
                                      .order_by(Notification.created_at.desc())\
                                      .limit(20).all()
    return jsonify([{
        "id": n.id,
        "type": n.type,
        "message": n.message,
        "is_read": n.is_read,
        "created_at": n.created_at.isoformat()
    } for n in notifications])



@app.route('/api/notifications/<int:notification_id>/read', methods=['PATCH'])
@jwt_required()
def mark_notification_read(notification_id):
    current_user_id = int(get_jwt_identity())
    notif = Notification.query.filter_by(id=notification_id, user_id=current_user_id).first_or_404()
    notif.is_read = True
    db.session.commit()
    return jsonify({"success": True})


@app.route('/api/notifications/mark_all_read', methods=['POST'])
@jwt_required()
def mark_all_notifications_read():
    current_user_id = int(get_jwt_identity())
    Notification.query.filter_by(user_id=current_user_id, is_read=False)\
                      .update({"is_read": True})
    db.session.commit()
    return jsonify({"success": True})


@app.route('/api/notifications/clear_all', methods=['DELETE'])
@jwt_required()
def clear_all_notifications():
    current_user_id = get_jwt_identity()
    try:
        if isinstance(current_user_id, str):
            current_user_id = int(current_user_id)
        elif not isinstance(current_user_id, int):
            raise ValueError("Invalid identity type")
    except (ValueError, TypeError):
        return jsonify({"error": "Invalid user identity"}), 400

    Notification.query.filter_by(user_id=current_user_id).delete()
    db.session.commit()
    return jsonify({"message": "All notifications cleared"}), 200


# ==========================
# 📦 MASTER DATA ROUTES
# ==========================
# 
# 
@app.route('/api/products', methods=['GET'])
@jwt_required()
def get_products():
    current_user_id = int(get_jwt_identity())
    user = User.query.get(current_user_id)
    if not user or user.role != 'admin':
        return jsonify({"msg": "Admin access required"}), 403

    products = Product.query.all()
    result = []
    for p in products:
        total_stock = db.session.query(func.sum(Stock.actual_quantity))\
                                .filter(Stock.product_id == p.id)\
                                .scalar() or 0
        stock_status = f"在庫あり ({total_stock}個)" if total_stock > 0 else "在庫なし"

        result.append({
            "id": p.id,
            "hcod": p.hcod,
            "hnm": p.hnm,
            "supplier_code": p.supplier_code,
            "unit_price": float(p.unit_price or 0),
            "stock_status": stock_status  
        })
    return jsonify(result)





@app.route('/api/products/search', methods=['GET'])
@jwt_required()
def search_products_admin():
    current_user_id = get_jwt_identity()
    user = User.query.get(int(current_user_id))
    if not user or user.role != 'admin':
        return jsonify({"error": "Admin access required"}), 403

    query = request.args.get('q', '').strip()
    if not query:
        return jsonify([])

    products = Product.query.filter(
        db.or_(
            Product.hcod.ilike(f'%{query}%'),
            Product.hnm.ilike(f'%{query}%')
        )
    ).limit(50).all()

    result = []
    for p in products:
        
        total_stock = db.session.query(func.sum(Stock.actual_quantity))\
                                .filter(Stock.product_id == p.id)\
                                .scalar() or 0
        stock_status = f"在庫あり ({total_stock}個)" if total_stock > 0 else "在庫なし"

        result.append({
            "id": p.id,
            "hcod": p.hcod,
            "hnm": p.hnm,
            "supplier_code": p.supplier_code,
            "unit_price": float(p.unit_price or 0),
            "stock_status": stock_status,
            "description": p.description
        })
    return jsonify(result)



@app.route('/api/products/<int:product_id>', methods=['PUT'])
@jwt_required()
def update_product(product_id):
    current_user_id = get_jwt_identity()
    user = User.query.get(int(current_user_id))
    if not user or user.role != 'admin':
        return jsonify({"msg": "Admin access required"}), 403

    product = Product.query.get_or_404(product_id)
    data = request.get_json()

    product.hcod = data.get('hcod', product.hcod)
    product.hnm = data.get('hnm', product.hnm)
    product.supplier_code = data.get('supplier_code', product.supplier_code)
    product.unit_price = float(data.get('unit_price', product.unit_price))
    product.description = data.get('description', product.description)

    db.session.commit()
    return jsonify({"success": True, "message": "Product updated"}), 200


@app.route('/api/suppliers', methods=['GET'])
@jwt_required()
def get_suppliers():
    current_user_id = get_jwt_identity()
    user = User.query.get(current_user_id)
    if not user or user.role != 'admin':
        return jsonify({"msg": "Admin access required"}), 403
    suppliers = Supplier.query.all()
    return jsonify([{"id": s.id, "mcod": s.mcod, "name": s.name, "standard_lead_time": s.standard_lead_time} for s in suppliers])



@app.route('/api/suppliers/<int:supplier_id>', methods=['PUT'])
@jwt_required()
def update_supplier(supplier_id):
    current_user_id = get_jwt_identity()
    user = User.query.get(int(current_user_id))
    if not user or user.role != 'admin':
        return jsonify({"msg": "Admin access required"}), 403

    supplier = Supplier.query.get_or_404(supplier_id)
    data = request.get_json()

    supplier.mcod = data.get('mcod', supplier.mcod)
    supplier.name = data.get('name', supplier.name)
    supplier.standard_lead_time = int(data.get('standard_lead_time', supplier.standard_lead_time))

    db.session.commit()
    return jsonify({"success": True, "message": "Supplier updated"}), 200



@app.route('/api/orders', methods=['GET'])
@jwt_required()
def get_all_orders():
    current_user_id = get_jwt_identity()
    user = User.query.get(int(current_user_id))
    if not user or user.role != 'admin':
        return jsonify({"msg": "Admin access required"}), 403

    orders = Order.query.order_by(Order.date_created.desc()).all()
    result = []
    for o in orders:
        customer = Customer.query.get(o.customer_id)
        result.append({
            "order_number": o.order_number,
            "customer_name": customer.name if customer else "Unknown",
            "customer_id": o.customer_id,
            "status": o.status,
            "date_created": o.date_created.isoformat(),
            "items": [
                {
                    "hcod": item.product.hcod,
                    "quantity": item.quantity,
                    "unit_price": float(item.unit_price),
                    "estimated_delivery_date": item.estimated_delivery_date.isoformat() if item.estimated_delivery_date else None,
                    "stock_status": item.stock_status
                }
                for item in o.items
            ]
        })
    return jsonify(result), 200


@app.route('/api/quotations/<int:quotation_id>/delivery', methods=['PATCH'])
@jwt_required()
def update_quotation_delivery_date(quotation_id):
    current_user_id = get_jwt_identity()
    user = User.query.get(int(current_user_id))

    if not user or user.role != 'admin':
        return jsonify({"msg": "Admin access required"}), 403

    quotation = Quotation.query.get_or_404(quotation_id)

    data = request.get_json()
    new_date_str = data.get('estimated_delivery_date')

    if not new_date_str:
        return jsonify({"error": "estimated_delivery_date is required"}), 400

    try:
        new_date = datetime.date.fromisoformat(new_date_str)
    except ValueError:
        return jsonify({"error": "Invalid date format, use YYYY-MM-DD"}), 400

    for item in quotation.items:
        item.estimated_delivery_date = new_date

    customer_user = User.query.filter(
        User.customer_id == quotation.customer_id,
        User.role == "customer"
    ).first()

    if customer_user:
        notif = Notification(
            user_id=customer_user.id,
            type='delivery_updated',
            related_id=quotation.id,
            message=f"Delivery date updated for quotation #{quotation.id}"
        )
        db.session.add(notif)

    db.session.commit()

    return jsonify({
        "success": True,
        "message": "Delivery date updated"
    }), 200



@app.route('/api/orders/<order_number>/delivery', methods=['PATCH'])
@jwt_required()
def update_order_delivery_date(order_number):
    current_user_id = get_jwt_identity()
    user = User.query.get(int(current_user_id))

    # Admin-only access
    if not user or user.role != 'admin':
        return jsonify({"msg": "Admin access required"}), 403

   
    order = Order.query.filter_by(order_number=order_number).first_or_404()

    data = request.get_json()
    new_date_str = data.get('estimated_delivery_date')

    if not new_date_str:
        return jsonify({"error": "estimated_delivery_date is required"}), 400

    try:
        new_date = datetime.date.fromisoformat(new_date_str)
    except ValueError:
        return jsonify({"error": "Invalid date format, use YYYY-MM-DD"}), 400

    
    for item in order.items:
        item.estimated_delivery_date = new_date
    customer_user = User.query.filter(
        User.customer_id == order.customer_id,
        User.role == "customer"
    ).first()

    if customer_user:
        notif = Notification(
            user_id=customer_user.id,
            type='order_delivery_updated',
            related_id=order.id,
            message=f"Delivery date updated for order {order.order_number}"
        )
        db.session.add(notif)

    db.session.commit()

    return jsonify({
        "success": True,
        "message": "Order delivery date updated"
    }), 200


# # Edit quotation item (admin only)
# @app.route('/api/quotations/<int:quotation_id>/items/<int:item_id>', methods=['PATCH'])
# @jwt_required()
# def update_quotation_item(quotation_id, item_id):
#     current_user_id = get_jwt_identity()
#     user = User.query.get(int(current_user_id))
#     if not user or user.role != 'admin':
#         return jsonify({"error": "Admin access required"}), 403

#     quotation = Quotation.query.get_or_404(quotation_id)
#     item = QuotationItem.query.filter_by(id=item_id, quotation_id=quotation_id).first_or_404()

#     data = request.get_json()
#     changes_made = False

#     # Update unit_price
#     if 'unit_price' in data:
#         item.unit_price = float(data['unit_price'])
#         changes_made = True

#     # Update estimated_delivery_date
#     if 'estimated_delivery_date' in data:
#         try:
#             item.estimated_delivery_date = datetime.date.fromisoformat(data['estimated_delivery_date'])
#             changes_made = True
#         except ValueError:
#             return jsonify({"error": "Invalid date format. Use YYYY-MM-DD"}), 400

#     if changes_made:
#         # Notify customer
#         from models import Notification
#         notif = Notification(
#             user_id=quotation.customer.user.id,
#             type='quotation_updated',
#             related_id=quotation.id,
#             message=f"Quotation #{quotation.id} has been updated by admin"
#         )
#         db.session.add(notif)
#         db.session.commit()
#         return jsonify({
#             "success": True,
#             "message": "Item updated and customer notified"
#         }), 200
#     else:
#         return jsonify({"success": False, "message": "No valid fields to update"}), 400




# Edit quotation item (admin only)
@app.route('/api/quotations/<int:quotation_id>/items/<int:item_id>', methods=['PATCH'])
@jwt_required()
def update_quotation_item(quotation_id, item_id):
    current_user_id = get_jwt_identity()
    user = User.query.get(int(current_user_id))
    if not user or user.role != 'admin':
        return jsonify({"error": "Admin access required"}), 403

    quotation = Quotation.query.get_or_404(quotation_id)
    item = QuotationItem.query.filter_by(id=item_id, quotation_id=quotation_id).first_or_404()

    data = request.get_json()
    changes_made = False

    # Update unit_price
    if 'unit_price' in data:
        try:
            item.unit_price = float(data['unit_price'])
            changes_made = True
        except (TypeError, ValueError):
            return jsonify({"error": "Invalid unit_price: must be a number"}), 400

    # Update estimated_delivery_date
    if 'estimated_delivery_date' in data:
        date_str = data['estimated_delivery_date']
        if date_str:
            try:
                item.estimated_delivery_date = datetime.date.fromisoformat(date_str)
                changes_made = True
            except ValueError:
                return jsonify({"error": "Invalid date format. Use YYYY-MM-DD"}), 400
        else:
            # Allow clearing the date
            item.estimated_delivery_date = None
            changes_made = True

    if not changes_made:
        return jsonify({"success": False, "message": "No valid fields to update"}), 400

    # Commit the item update first
    db.session.commit()

    # 🔒 Safely notify customer only if user exists
    customer_user_id = None
    if quotation.customer and quotation.customer.user:
        customer_user_id = quotation.customer.user.id
    elif quotation.customer_id:
        # Fallback: try to load customer explicitly (in case relationship wasn't loaded)
        from models import Customer
        customer = Customer.query.get(quotation.customer_id)
        if customer and customer.user:
            customer_user_id = customer.user.id

    if customer_user_id:
        from models import Notification
        notif = Notification(
            user_id=customer_user_id,
            type='quotation_updated',
            related_id=quotation.id,
            message=f"Quotation #{quotation.id} has been updated by admin"
        )
        db.session.add(notif)
        db.session.commit()
        message = "Item updated and customer notified"
    else:
        # Log for debugging, but don't fail
        app.logger.warning(f"Quotation {quotation_id} has no linked customer user — notification skipped.")
        message = "Item updated (no customer to notify)"

    return jsonify({
        "success": True,
        "message": message
    }), 200


@app.route('/api/customers', methods=['GET'])
@jwt_required()
def get_customers():
    current_user_id = get_jwt_identity()
    user = User.query.get(current_user_id)
    if not user or user.role != 'admin':
        return jsonify({"msg": "Admin access required"}), 403
    customers = Customer.query.all()
    return jsonify([{"id": c.id, "ucod": c.ucod, "name": c.name, "email": c.email} for c in customers])



@app.route('/api/products', methods=['POST'])
@jwt_required()
def add_product():
    current_user_id = get_jwt_identity()
    user = User.query.get(int(current_user_id))
    if not user or user.role != 'admin':
        return jsonify({"msg": "Admin access required"}), 403

    data = request.get_json()
    hcod = data.get('hcod')
    hnm = data.get('hnm')
    supplier_code = data.get('supplier_code')
    unit_price = data.get('unit_price')
    description = data.get('description', '')

    if not hcod or not hnm or supplier_code is None:
        return jsonify({"error": "hcod, hnm, and supplier_code are required"}), 400

    if Product.query.filter_by(hcod=hcod).first():
        return jsonify({"error": "Product with this hcod already exists"}), 400

    product = Product(
        hcod=hcod,
        hnm=hnm,
        supplier_code=supplier_code,
        unit_price=float(unit_price) if unit_price else 0.0,
        description=description
    )
    db.session.add(product)
    db.session.commit()

    return jsonify({"success": True, "product_id": product.id}), 201



@app.route('/api/products/<int:product_id>', methods=['DELETE'])
@jwt_required()
def delete_product(product_id):
    current_user_id = get_jwt_identity()
    user = User.query.get(int(current_user_id))
    if not user or user.role != 'admin':
        return jsonify({"msg": "Admin access required"}), 403

    product = Product.query.get_or_404(product_id)
    Stock.query.filter_by(product_id=product.id).delete()
    QuotationItem.query.filter_by(product_id=product.id).delete()
    OrderItem.query.filter_by(product_id=product.id).delete()
    db.session.delete(product)
    db.session.commit()

    return jsonify({"success": True, "message": "Product and related data deleted"}), 200




@app.route('/api/inventory/stock/<int:product_id>', methods=['PUT'])
@jwt_required()
def update_stock(product_id):
    current_user_id = get_jwt_identity()
    user = User.query.get(int(current_user_id))
    if not user or user.role != 'admin':
        return jsonify({"msg": "Admin access required"}), 403

    product = Product.query.get_or_404(product_id)
    supplier = Supplier.query.filter_by(mcod=product.supplier_code).first()
    if not supplier:
        return jsonify({"error": f"No supplier found for product supplier_code: {product.supplier_code}"}), 400

    data = request.get_json()
    new_qty = data.get('actual_quantity')
    if new_qty is None:
        return jsonify({"error": "actual_quantity is required"}), 400

    stock = Stock.query.filter_by(product_id=product.id, supplier_id=supplier.id).first()
    if not stock:
        stock = Stock(
            product_id=product.id,
            supplier_id=supplier.id,  
            actual_quantity=int(new_qty)
        )
        db.session.add(stock)
    else:
        stock.actual_quantity = int(new_qty)
    
    db.session.commit()
    return jsonify({"success": True, "stock_level": stock.actual_quantity}), 200



@app.route('/api/suppliers', methods=['POST'])
@jwt_required()
def add_supplier():
    current_user_id = get_jwt_identity()
    user = User.query.get(int(current_user_id))
    if not user or user.role != 'admin':
        return jsonify({"msg": "Admin access required"}), 403

    data = request.get_json()
    mcod = data.get('mcod')
    name = data.get('name')
    lead_time = data.get('standard_lead_time', 0)

    if not mcod or not name:
        return jsonify({"error": "mcod and name are required"}), 400

    if Supplier.query.filter_by(mcod=mcod).first():
        return jsonify({"error": "Supplier with this mcod already exists"}), 400

    supplier = Supplier(mcod=mcod, name=name, standard_lead_time=int(lead_time))
    db.session.add(supplier)
    db.session.commit()
    return jsonify({"success": True, "supplier_id": supplier.id}), 201




@app.route('/api/suppliers/<int:supplier_id>', methods=['DELETE'])
@jwt_required()
def delete_supplier(supplier_id):
    current_user_id = get_jwt_identity()
    user = User.query.get(int(current_user_id))
    if not user or user.role != 'admin':
        return jsonify({"msg": "Admin access required"}), 403

    supplier = Supplier.query.get_or_404(supplier_id)

    
    Product.query.filter_by(supplier_code=supplier.mcod).delete()
    Stock.query.filter_by(supplier_id=supplier.id).delete()

    db.session.delete(supplier)
    db.session.commit()
    return jsonify({"success": True, "message": "Supplier and related products deleted"}), 200



# Add supplier search (admin-only)
@app.route('/api/suppliers/search', methods=['GET'])
@jwt_required()
def search_suppliers():
    current_user_id = get_jwt_identity()
    user = User.query.get(int(current_user_id))
    if not user or user.role != 'admin':
        return jsonify({"error": "Admin access required"}), 403

    query = request.args.get('q', '').strip()
    if not query:
        return jsonify([])

    suppliers = Supplier.query.filter(
        db.or_(
            Supplier.mcod.ilike(f'%{query}%'),
            Supplier.name.ilike(f'%{query}%')
        )
    ).limit(50).all()

    return jsonify([
        {
            "id": s.id,
            "mcod": s.mcod,
            "name": s.name,
            "standard_lead_time": s.standard_lead_time,
            "address": s.address,
            "phone": s.phone
        }
        for s in suppliers
    ])

# ==========================
# 📦 INVENTORY ROUTES
# ==========================
@app.route('/api/inventory/check', methods=['POST'])
@jwt_required()
def check_inventory():
    hcod = request.json.get('hcod')
    qty = request.json.get('qty', 1)
    if not hcod:
        return jsonify({"error": "H-CODE is required"}), 400
    result = inventory_service.check_stock(hcod, qty)
    return jsonify(result)




@app.route('/api/inventory/lead_time', methods=['POST'])
@jwt_required()
def calculate_lead_time_api():
    hcod = request.json.get('hcod')
    supplier_code = request.json.get('supplier_code')
    qty = request.json.get('qty', 1)
    required_date = request.json.get('required_date')
    if not hcod or not supplier_code:
        return jsonify({"error": "H-CODE and Supplier Code are required"}), 400
    result = inventory_service.calculate_lead_time_and_status(hcod, supplier_code, qty, customer_delivery_date=required_date)
    return jsonify(result)

# ==========================
# 📝 QUOTATION ROUTES
# ==========================


@app.route('/api/quotations', methods=['GET'])
@jwt_required()
def get_quotations():
    current_user_id = get_jwt_identity()
    user = User.query.get(current_user_id)
    if not user:
        return jsonify({"msg": "Unauthorized"}), 401
    if user.role == 'admin':
        quotations = Quotation.query.options(db.joinedload(Quotation.customer)).all()
    elif user.role == 'customer' and user.customer:
        quotations = Quotation.query.filter_by(customer_id=user.customer.id).options(db.joinedload(Quotation.customer)).all()
    else:
        return jsonify({"msg": "Unauthorized"}), 403

    result = []
    for q in quotations:
        items = []
        for qi in q.items:
            items.append({
                "id": qi.id,
                "product_hcod": qi.product.hcod if qi.product else "N/A",
                "quantity": qi.quantity,
                "unit_price": float(qi.unit_price or 0),
                "lead_time_days": qi.lead_time_days,
                "estimated_delivery_date": qi.estimated_delivery_date.isoformat() if qi.estimated_delivery_date else None,
                "stock_status": qi.stock_status
            })
        result.append({
            "id": q.id,
            "customer_id": q.customer_id,
            "customer_ucod": q.customer.ucod if q.customer else "—",  # ✅ ADD THIS
            "status": q.status,
            "date_created": q.date_created.isoformat(),
            "items": items,
            "total_amount": sum(float(item["unit_price"]) * item["quantity"] for item in items)
        })
    return jsonify(result)


@app.route('/api/quotations', methods=['POST'])
@jwt_required()
def create_quotation():
    current_user_id = get_jwt_identity()
    user = User.query.get(current_user_id)
    if not user:
        return jsonify({"msg": "Unauthorized"}), 401

    data = request.get_json(force=True)
    if not data:
        return jsonify({"error": "Invalid or missing JSON body"}), 400

    customer_id_from_request = data.get('customer_id')

    if user.role == 'admin':
        if not customer_id_from_request:
            return jsonify({"error": "Customer ID is required for admin users"}), 400
        customer_id = customer_id_from_request

    elif user.role == 'customer' and user.customer:
        customer_id = user.customer.id
        if customer_id_from_request and customer_id_from_request != customer_id:
            return jsonify({"error": "Customer cannot create quotation for another customer"}), 403

    else:
        return jsonify({"msg": "Unauthorized"}), 403

    items_data = data.get('items', [])
    if not items_data:
        return jsonify({"error": "Items are required"}), 400

    result = quotation_service.create_quotation(customer_id, items_data)

    if result.get("success"):
        return jsonify(result), 201
    else:
        return jsonify(result), 400


# @app.route('/api/quotations/<int:quotation_id>', methods=['GET'])
# @jwt_required()
# def get_quotation(quotation_id):
#     quotation = Quotation.query.get_or_404(quotation_id)
#     current_user_id = get_jwt_identity()
#     user = User.query.get(current_user_id)
#     if not user or (user.role != 'admin' and (user.role != 'customer' or user.customer.id != quotation.customer_id)):
#         return jsonify({"msg": "Unauthorized"}), 403
#     items = [{
#         "product_hcod": qi.product.hcod if qi.product else "N/A",
#         "quantity": qi.quantity,
#         "unit_price": float(qi.unit_price or 0),
#         "lead_time_days": qi.lead_time_days,
#         "estimated_delivery_date": qi.estimated_delivery_date.isoformat() if qi.estimated_delivery_date else "TBD",
#         "stock_status": qi.stock_status
#     } for qi in quotation.items]
#     return jsonify({
#         "id": quotation.id,
#         "customer_id": quotation.customer_id,
#         "status": quotation.status,
#         "date_created": quotation.date_created.isoformat(),
#         "items": items,
#         "total_amount": sum([float(i["unit_price"]) * i["quantity"] for i in items])
#     })

@app.route('/api/quotations/<int:quotation_id>', methods=['GET'])
@jwt_required()
def get_quotation(quotation_id):
    current_user_id = get_jwt_identity()
    user = User.query.get(current_user_id)

    if not user:
        return jsonify({"msg": "Unauthorized"}), 401

    quotation = Quotation.query.options(
        db.joinedload(Quotation.items).joinedload(QuotationItem.product),
        db.joinedload(Quotation.customer)
    ).get_or_404(quotation_id)

    # Customer access check
    if user.role == "customer" and quotation.customer_id != user.customer.id:
        return jsonify({"msg": "Unauthorized"}), 403

    # Serialize clean items safely
    items = []
    for qi in quotation.items:
        if qi is None:
            continue  # 🔥 skip null entries (prevents frontend crash)

        items.append({
            "id": qi.id,
            "product_hcod": qi.product.hcod if qi.product else "N/A",
            "description": qi.product.description if qi.product else "",
            "quantity": qi.quantity or 0,
            "unit_price": float(qi.unit_price or 0),
            "lead_time_days": qi.lead_time_days,
            "estimated_delivery_date": qi.estimated_delivery_date.isoformat()
            if qi.estimated_delivery_date else None,
            "stock_status": qi.stock_status or "—",
        })

    return jsonify({
        "id": quotation.id,
        "customer_id": quotation.customer_id,
        "customer_ucod": quotation.customer.ucod if quotation.customer else "—",
        "status": quotation.status,
        "date_created": quotation.date_created.isoformat(),
        "items": items,  # 🔥 ALWAYS a clean valid list
        "total_amount": sum(float(i["unit_price"]) * i["quantity"] for i in items),
    })


# ==========================
# 📧 INCOMING EMAIL QUOTATION REQUESTS
# ==========================
@app.route('/api/incoming_requests', methods=['GET'])
@jwt_required()
def get_incoming_requests():
    current_user_id = get_jwt_identity()
    user = User.query.get(current_user_id)
    if not user or user.role != 'admin':
        return jsonify({"msg": "Admin access required"}), 403
    requests = IncomingQuotationRequest.query.order_by(IncomingQuotationRequest.received_date.desc()).all()
    return jsonify([{
        "id": r.id,
        "subject": r.subject,
        "sender": r.sender,
        "received_date": r.received_date.isoformat() if r.received_date else None,
        "status": r.status,
        "items": r.items_data,
        "customer_name": r.customer_name,
        "customer_id": r.customer_id,
        "notes": r.notes
    } for r in requests])


@app.route('/api/incoming_requests/<int:request_id>/process', methods=['POST'])
@jwt_required()
def process_incoming_request(request_id):
    current_user_id = get_jwt_identity()
    user = User.query.get(current_user_id)
    if not user or user.role != 'admin':
        return jsonify({"msg": "Admin access required"}), 403
    req = IncomingQuotationRequest.query.get_or_404(request_id)
    if req.status != 'pending':
        return jsonify({"error": "Request is not in pending state"}), 400
    try:
        customer_id = req.customer_id
        if not customer_id:
            customer = Customer.query.filter_by(email=req.sender).first()
            if not customer:
                from utils.xlsx_loader import generate_new_ucod
                existing_ucods = {c.ucod for c in Customer.query.all()}
                new_ucod = generate_new_ucod(existing_ucods)
                customer = Customer(ucod=new_ucod, name=req.customer_name or "Email Customer", email=req.sender)
                new_user = User(username=req.sender.split('@')[0], email=req.sender, role='customer')
                new_user.set_password('default_temp_password')
                new_user.customer = customer
                db.session.add(new_user)
                db.session.add(customer)
                db.session.commit()
                customer_id = customer.id
                req.customer_id = customer_id
        items_data = req.items_data
        result = quotation_service.create_quotation(customer_id, items_data)
        if result.get("success"):
            req.status = 'processed'
            req.notes = f"Successfully created quotation ID {result.get('quotation_id')}"
            db.session.commit()
            return jsonify({"message": "Request processed successfully", "quotation_id": result.get("quotation_id")})
        else:
            req.status = 'error'
            req.notes = f"Processing failed: {result.get('error', 'Unknown error')}"
            db.session.commit()
            return jsonify({"error": result.get("error", "Unknown error during processing")}), 500
    except Exception as e:
        log.error(f"Error processing incoming request {request_id}: {e}", exc_info=True)
        req.status = 'error'
        req.notes = f"Processing error: {str(e)}"
        db.session.commit()
        return jsonify({"error": "Internal server error during processing"}), 500

# ==========================
# 🛒 PROCUREMENT ROUTES
# ==========================
@app.route('/api/purchase_orders', methods=['GET'])
@jwt_required()
def get_purchase_orders():
    current_user_id = get_jwt_identity()
    user = User.query.get(current_user_id)
    if not user or user.role != 'admin':
        return jsonify({"msg": "Admin access required"}), 403
    pos = PurchaseOrder.query.all()
    return jsonify([{
        "id": po.id,
        "supplier_id": po.supplier_id,
        "status": po.status,
        "date_created": po.date_created.isoformat(),
        "total_items": len(po.items)
    } for po in pos])

@app.route('/api/purchase_orders', methods=['POST'])
@jwt_required()
def create_purchase_order():
    current_user_id = get_jwt_identity()
    user = User.query.get(current_user_id)
    if not user or user.role != 'admin':
        return jsonify({"msg": "Admin access required"}), 403
    data = request.get_json()
    supplier_id = data.get('supplier_id')
    items_data = data.get('items', [])
    if not supplier_id or not items_data:
        return jsonify({"error": "Supplier ID and items are required"}), 400
    from services.procurement_service import ProcurementService
    proc_service = ProcurementService(db.session, inventory_service)
    result = proc_service.create_purchase_order(supplier_id, items_data)
    if result.get("success"):
        return jsonify(result), 201
    else:
        return jsonify(result), 400



# Get order details (customer or admin)
@app.route('/api/orders/<order_number>', methods=['GET'])
@jwt_required()
def get_order_details(order_number):
    claims = get_jwt()
    role = claims.get("role")
    
    order = Order.query.filter_by(order_number=order_number).first_or_404()
    
    # Authorization
    if role == 'customer':
        customer_id = claims.get("customer_id")
        if not customer_id or order.customer_id != customer_id:
            return jsonify({"error": "Unauthorized"}), 403
    elif role != 'admin':
        return jsonify({"error": "Unauthorized"}), 403

    # Fetch customer name
    customer = Customer.query.get(order.customer_id)
    
    items = []
    for item in order.items:
        items.append({
            "hcod": item.product.hcod if item.product else "N/A",
            "quantity": item.quantity,
            "unit_price": float(item.unit_price or 0),
            "estimated_delivery_date": item.estimated_delivery_date.isoformat() if item.estimated_delivery_date else None,
            "stock_status": item.stock_status,
            "lead_time_days": item.lead_time_days
        })

    return jsonify({
        "order_number": order.order_number,
        "customer_ucod": customer.ucod if customer else "—",
        "customer_name": customer.name if customer else "—",
        "status": order.status,
        "date_created": order.date_created.isoformat(),
        "total_amount": float(order.total_amount or sum(item["unit_price"] * item["quantity"] for item in items)),
        "items": items
    })

# ==========================
# 📅 CALENDAR UTILITIES
# ==========================
@app.route('/api/calendar/is_business_day/<date_str>', methods=['GET'])
@jwt_required()
def check_business_day(date_str):
    try:
        date_obj = datetime.datetime.strptime(date_str, "%Y%m%d").date()
        calendar_data = load_calendar()
        is_bd = is_business_day(date_obj, calendar_data)
        return jsonify({"date": date_str, "is_business_day": is_bd})
    except ValueError:
        return jsonify({"error": "Invalid date format, expected YYYYMMDD"}), 400



# ==========================
# 📄 REPORTING / FILE GENERATION
# ==========================

@app.route('/api/reports/quotations/<int:quotation_id>/pdf', methods=['GET'])
@jwt_required()
def get_quotation_pdf(quotation_id):
    # Fetch quotation
    quotation = Quotation.query.get_or_404(quotation_id)
    current_user_id = get_jwt_identity()
    user = User.query.get(current_user_id)

    # Authorization
    if not user or (user.role != 'admin' and (user.role != 'customer' or user.customer.id != quotation.customer_id)):
        return jsonify({"msg": "Unauthorized"}), 403

    # Prepare PDF
    pdf_buffer = BytesIO()
    pdf = FPDF()
    pdf.add_page()
    font_path = os.path.join(os.path.dirname(__file__), "fonts", "NotoSansJP-Regular.ttf")
    if os.path.exists(font_path):
        pdf.add_font("NotoSansJP", "", font_path, uni=True)
        pdf.set_font("NotoSansJP", size=12)
    else:
        raise FileNotFoundError(" NotoSansJP-Regular.ttf not found. Japanese characters cannot be rendered.")

    pdf.cell(0, 10, f"{quotation.customer.name} 様", ln=True)
    pdf.cell(0, 8, "お見積書", ln=True)
    pdf.ln(2)
    pdf.cell(0, 6, "以下の通りお見積り申し上げます。", ln=True)
    pdf.ln(4)

    # Table headers
    headers = ["品番", "品名/説明", "数量", "単価 (JPY)", "金額 (JPY)", "納期 (日数)", "納入予定日", "在庫状況"]
    header_line = " | ".join(headers)
    pdf.cell(0, 8, header_line, ln=True)
    pdf.cell(0, 4, "-" * 160, ln=True)

    total = 0
   
    for qi in quotation.items:
        qty = int(qi.quantity) if qi.quantity else 1
        price = float(qi.unit_price or 0)
        line_total = price * qty
        total += line_total
        desc = (qi.product.hnm if qi.product else qi.input_code or '')[:40]
        lead_time = qi.lead_time_days or 0

        
        if qi.estimated_delivery_date:
            formatted_date = qi.estimated_delivery_date.strftime("%Y/%m/%d")
        else:
            formatted_date = "TBD"
        stock_status = qi.stock_status or "N/A"
        product_code = qi.product.hcod if qi.product else qi.input_code or 'N/A'
        pdf.cell(
            0, 8,
            f"{product_code} | {desc} | {qty} | {int(price):,} | {int(line_total):,} | {lead_time}日 | {formatted_date} | {stock_status}",
            new_x=XPos.LMARGIN,
            new_y=YPos.NEXT
        )

    pdf.cell(0, 6, "-" * 160, ln=True)
    pdf.cell(0, 10, f"合計金額: ¥{int(total):,}", ln=True)
    pdf.ln(6)
    pdf.cell(0, 6, "備考:", ln=True)
    pdf.cell(0, 6, "納期・在庫は別途ご確認ください。税抜価格で記載。", ln=True)
    pdf_buffer.write(pdf.output(dest='S'))
    pdf_buffer.seek(0)

    return send_file(
        pdf_buffer,
        as_attachment=True,
        download_name=f"quotation_{quotation_id}.pdf",
        mimetype='application/pdf'
    )


# ==========================
# 🔍 CUSTOMER PORTAL ROUTES
# ==========================
@app.route('/api/customer_portal/products/search', methods=['GET'])
@jwt_required()
def search_products_for_customer():
    current_user_id = get_jwt_identity()
    claims = get_jwt()
    role = claims.get("role")

    if role != "customer":
        return jsonify({"error": "Unauthorized"}), 403
    query = request.args.get('q', '')
    if not query:
        return jsonify({"error": "Query string 'q' is required"}), 400
    products = Product.query.filter(
        db.or_(
            Product.hcod.ilike(f'%{query}%'),
            Product.hnm.ilike(f'%{query}%')
        )
    ).limit(20).all()

    results = []

    for p in products:
        stock_records = Stock.query.filter_by(product_id=p.id).all()
        total_stock = sum(s.actual_quantity for s in stock_records)

        stock_status = (
            f"在庫あり ({total_stock}個)" if total_stock > 0 else "在庫なし"
        )

        results.append({
            "hcod": p.hcod,
            "hnm": p.hnm,
            "description": p.description,
            "unit_price": float(p.unit_price or 0),
            "stock_status": stock_status
        })

    return jsonify(results)



@app.route('/api/customer_portal/info', methods=['GET'])
@jwt_required()
def get_customer_info():
    current_user_id = get_jwt_identity()
    claims = get_jwt()

    role = claims.get("role")
    customer_id = claims.get("customer_id")

    if role != "customer":
        return jsonify({"error": "Unauthorized"}), 403

    if not customer_id:
        return jsonify({"error": "Customer ID not found"}), 404

    customer = Customer.query.get_or_404(customer_id)

    return jsonify({
        "ucod": customer.ucod,
        "name": customer.name,
        "email": customer.email,
        "phone": customer.phone,
        "address": customer.address
    })


@app.route('/api/customer_portal/dashboard_stats', methods=['GET'])
@jwt_required()
def get_customer_dashboard_stats():
    claims = get_jwt()
    if claims.get("role") != "customer":
        return jsonify({"error": "Unauthorized"}), 403
    customer_id = claims.get("customer_id")
    if not customer_id:
        return jsonify({"error": "Customer ID missing"}), 404

    total_quotations = Quotation.query.filter_by(customer_id=customer_id).count()
    total_orders = Order.query.filter_by(customer_id=customer_id).count()
    total_products = db.session.query(func.count(Product.id)).scalar() or 0

    return jsonify({
        "total_quotations": total_quotations,
        "total_orders": total_orders,
        "total_products_available": total_products
    })



@app.route('/api/customer_portal/request_quote', methods=['POST'])
@jwt_required()
def create_customer_quotation_request():

    current_user_id = get_jwt_identity()
    claims = get_jwt()
    role = claims.get("role")
    customer_id = claims.get("customer_id")

    # Authorization check
    if role != "customer":
        return jsonify({"error": "Unauthorized"}), 403

    if not customer_id:
        return jsonify({"error": "Customer ID not found"}), 404

    data = request.get_json()
    items_data = data.get('items', [])
    notes = data.get('notes', '')

    if not items_data:
        return jsonify({"error": "Items are required"}), 400

    customer = Customer.query.get_or_404(customer_id)
    temp_request = IncomingQuotationRequest(
        subject=f"Portal Quote Request from {customer.name}",
        body=notes,
        sender=customer.email,
        received_date=datetime.datetime.utcnow(),
        status='pending',
        items_data=items_data,
        customer_name=customer.name,
        customer_id=customer_id
    )

    db.session.add(temp_request)
    db.session.commit()

    try:
      
        result = quotation_service.create_quotation(customer.id, items_data)
        if result.get("success"):
            quotation_id = result.get("quotation_id")

            temp_request.status = 'processed'
            temp_request.notes = f"Successfully created quotation ID {quotation_id}"
            db.session.commit()
            admin_users = User.query.filter_by(role='admin').all()
            for admin in admin_users:
                notif = Notification(
                    user_id=admin.id,
                    type='new_quotation',
                    related_id=quotation_id,
                    message=f"New quotation #{quotation_id} from {customer.name}"
                )
                db.session.add(notif)
            customer_user = User.query.filter_by(
                customer_id=customer.id,
                role='customer'
            ).first()
            if customer_user:
                notif2 = Notification(
                    user_id=customer_user.id,
                    type='quotation_created',
                    related_id=quotation_id,
                    message=f"Your quotation request has been processed — quotation #{quotation_id} created."
                )
                db.session.add(notif2)
            db.session.commit()

            return jsonify({
                "message": "Quote request submitted successfully",
                "quotation_id": quotation_id
            }), 201
        else:
            error_message = result.get("error", "Unknown error")

            temp_request.status = 'error'
            temp_request.notes = f"Processing failed: {error_message}"
            db.session.commit()

            return jsonify({"error": error_message}), 500
    except Exception as e:
        log.error(f"[Customer Portal Quote Request Error] {e}", exc_info=True)

        temp_request.status = 'error'
        temp_request.notes = f"Processing exception: {str(e)}"
        db.session.commit()
        return jsonify({"error": "Internal server error during processing"}), 500




@app.route('/api/quotations/<int:quotation_id>', methods=['DELETE'])
@jwt_required()
def delete_quotation(quotation_id):
    current_user_id = get_jwt_identity()
    user = User.query.get(current_user_id)

    if not user:
        return jsonify({"success": False, "message": "Unauthorized"}), 401
    quotation = Quotation.query.get_or_404(quotation_id)   
    if user.role == 'customer':
        if not user.customer or user.customer.id != quotation.customer_id:
            return jsonify({"success": False, "message": "You do not have permission to delete this quotation"}), 403    
        if quotation.status != 'draft':
            return jsonify({
                "success": False,
                "message": "Customers can only delete quotations with status 'draft'"
            }), 403
    elif user.role != 'admin':
        return jsonify({"success": False, "message": "Unauthorized"}), 403
    try:
       
        for item in quotation.items:
            db.session.delete(item)

        db.session.delete(quotation)
        db.session.commit()

        return jsonify({
            "success": True,
            "message": "Quotation deleted successfully"
        }), 200

    except Exception as e:
        db.session.rollback()
        return jsonify({
            "success": False,
            "message": f"Error deleting quotation: {str(e)}"
        }), 500


    


@app.route('/api/customer_portal/upload_request', methods=['POST'])
@jwt_required()
def upload_customer_request_file():

    
    current_user_id = get_jwt_identity()
    claims = get_jwt()
    role = claims.get("role")
    customer_id = claims.get("customer_id")

    if role != "customer":
        return jsonify({"error": "Unauthorized"}), 403

    if not customer_id:
        return jsonify({"error": "Customer ID not found"}), 404


    # ---------------- FILE VALIDATION ----------------
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400
    
    file = request.files['file']

    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400

    ALLOWED_EXTENSIONS = {'.pdf', '.csv', '.txt', '.xlsx', '.xls'}
    ext = '.' + (file.filename.rsplit('.', 1)[1].lower() if '.' in file.filename else '')

    if ext not in ALLOWED_EXTENSIONS:
        return jsonify({"error": "File type not allowed"}), 400


    # ---------------- SAVE FILE ----------------
    filename = secure_filename(file.filename)
    temp_dir = os.path.join(tempfile.gettempdir(), "customer_uploads")
    os.makedirs(temp_dir, exist_ok=True)

    filepath = os.path.join(temp_dir, f"{customer_id}_{int(time.time())}_{filename}")
    file.save(filepath)

    try:
        # OCR / file processing
        from utils.ocr_utils import process_uploaded_file_for_items
        items = process_uploaded_file_for_items(filepath)

        customer = Customer.query.get_or_404(customer_id)

        temp_request = IncomingQuotationRequest(
            subject=f"Portal File Upload Quote Request from {customer.name}",
            body=f"Uploaded file: {filename}",
            sender=customer.email,
            received_date=datetime.datetime.utcnow(),
            status='pending',
            items_data=items,
            customer_name=customer.name,
            customer_id=customer_id
        )

        db.session.add(temp_request)
        db.session.commit()


        # ---------------- CREATE QUOTATION ----------------
        result = quotation_service.create_quotation(customer.id, items)

        if result.get("success"):
            quotation_id = result.get("quotation_id")

            temp_request.status = 'processed'
            temp_request.notes = f"Successfully created quotation ID {quotation_id}"
            db.session.commit()

           
            if os.path.exists(filepath):
                os.remove(filepath)

            admin_users = User.query.filter_by(role='admin').all()
            for admin in admin_users:
                notif = Notification(
                    user_id=admin.id,
                    type='new_quotation',
                    related_id=quotation_id,
                    message=f"New quotation #{quotation_id} from {customer.name}"
                )
                db.session.add(notif)

            # Notify customer portal user
            customer_user = User.query.filter_by(
                customer_id=customer.id,
                role='customer'
            ).first()

            if customer_user:
                notif2 = Notification(
                    user_id=customer_user.id,
                    type='quotation_created',
                    related_id=quotation_id,
                    message=f"Your uploaded file has been processed — quotation #{quotation_id} created."
                )
                db.session.add(notif2)

            db.session.commit()

            return jsonify({
                "message": "File uploaded and quotation created",
                "quotation_id": quotation_id
            }), 201

        else:
            error_msg = result.get("error", "Unknown error")

            temp_request.status = 'error'
            temp_request.notes = f"Processing failed: {error_msg}"
            db.session.commit()

            if os.path.exists(filepath):
                os.remove(filepath)

            return jsonify({"error": error_msg}), 500

    except Exception as e:
        log.error(f"Error processing uploaded file {filepath}: {e}", exc_info=True)

        # ensure temporary file is removed
        if os.path.exists(filepath):
            os.remove(filepath)

        # update request record if it exists
        if 'temp_request' in locals():
            temp_request.status = 'error'
            temp_request.notes = f"Processing exception: {str(e)}"
            db.session.commit()

        return jsonify({"error": "Error processing uploaded file"}), 500





@app.route('/api/customer_portal/order', methods=['POST'])
@jwt_required()
def place_customer_order():
    claims = get_jwt()

    # Ensure role === customer
    if claims.get("role") != "customer":
        return jsonify({"error": "Unauthorized"}), 403

    customer_id = claims.get("customer_id")
    if not customer_id:
        return jsonify({"error": "Customer ID missing"}), 400

    data = request.get_json() or {}
    items_data = data.get('items', [])

    if not items_data:
        return jsonify({"error": "Items are required"}), 400

    # Validate customer
    customer = Customer.query.get(customer_id)
    if not customer:
        return jsonify({"error": "Customer not found"}), 404

    # ----------------------------------------
    # Generate new unique order number for today
    # ----------------------------------------
    today = datetime.datetime.now().strftime("%Y%m%d")
    last_order = (
        Order.query
        .filter(Order.order_number.like(f"ORD-{today}-%"))
        .order_by(Order.id.desc())
        .first()
    )

    if last_order:
        last_num = int(last_order.order_number.split("-")[-1])
        new_num = last_num + 1
    else:
        new_num = 1

    order_number = f"ORD-{today}-{new_num:03d}"

    # ----------------------------------------
    # Create order
    # ----------------------------------------
    order = Order(
        order_number=order_number,
        customer_id=customer_id,
        status='confirmed'
    )
    db.session.add(order)
    db.session.flush()  # get order.id

    total_amount = 0
    final_delivery_date = None

    for item in items_data:
        hcod = item.get('hcod')
        qty = item.get('quantity', 1)

        product = Product.query.filter_by(hcod=hcod).first()
        if not product:
            db.session.rollback()
            return jsonify({"error": f"Product not found: {hcod}"}), 404

        # Lead time logic
        lead_time_result = calculate_lead_time_and_status(
            hcod,
            product.supplier_code,
            qty,
            customer_delivery_date=None
        )

        delivery_date = lead_time_result.get("estimated_delivery_date")
        if delivery_date:
            delivery_date_obj = datetime.date.fromisoformat(delivery_date)
            final_delivery_date = delivery_date_obj
        else:
            delivery_date_obj = None

        unit_price = float(product.unit_price or 0)

        order_item = OrderItem(
            order_id=order.id,
            product_id=product.id,
            quantity=qty,
            unit_price=unit_price,
            lead_time_days=lead_time_result.get("lead_time_days"),
            estimated_delivery_date=delivery_date_obj,
            stock_status=lead_time_result.get("stock_status")
        )
        db.session.add(order_item)

        total_amount += unit_price * qty

    # ----------------------------------------
    # Commit order + items
    # ----------------------------------------
    db.session.commit()

    # ----------------------------------------
    # Create notifications for admins
    # ----------------------------------------
    admin_users = User.query.filter_by(role='admin').all()
    for admin in admin_users:
        notif = Notification(
            user_id=admin.id,
            type='new_order',
            related_id=order.id,
            message=f"New order {order_number} from {customer.name}"
        )
        db.session.add(notif)

    db.session.commit()

    return jsonify({
        "success": True,
        "order_number": order_number,
        "delivery_date": final_delivery_date.isoformat() if final_delivery_date else None,
        "total_amount": total_amount
    }), 201



def allowed_file(filename):
    ALLOWED_EXTENSIONS = {'.pdf', '.csv', '.txt', '.xlsx', '.xls'}
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS



@app.route('/api/customer_portal/orders', methods=['GET'])
@jwt_required()
def get_customer_orders():
    claims = get_jwt()
    if claims.get("role") != "customer":
        return jsonify({"error": "Unauthorized"}), 403
    customer_id = claims.get("customer_id")
    if not customer_id:
        return jsonify({"error": "Customer ID missing"}), 404

    orders = Order.query.filter_by(customer_id=customer_id)\
                        .options(db.joinedload(Order.items)).order_by(Order.date_created.desc()).all()
    result = []
    for o in orders:
        items = []
        for item in o.items:
            items.append({
                "hcod": item.product.hcod if item.product else "N/A",
                "quantity": item.quantity,
                "unit_price": float(item.unit_price or 0),
                "estimated_delivery_date": item.estimated_delivery_date.isoformat() if item.estimated_delivery_date else None,
                "stock_status": item.stock_status
            })
        result.append({
            "order_number": o.order_number,
            "status": o.status,
            "date_created": o.date_created.isoformat(),
            "items": items
        })
    return jsonify(result)



@app.route('/api/customer_portal/notifications', methods=['GET'])
@jwt_required()
def get_customer_notifications():
    claims = get_jwt()
    if claims.get("role") != "customer":
        return jsonify({"error": "Unauthorized"}), 403
    customer_id = claims.get("customer_id")

    # New quotations
    new_quotations = Quotation.query.filter_by(
        customer_id=customer_id,
        status='draft'
    ).count()

    # New orders
    new_orders = Order.query.filter_by(
        customer_id=customer_id
    ).count()

    return jsonify({
        "total": new_quotations + new_orders,
        "items": [
            {"type": "quotation", "count": new_quotations, "message": f"{new_quotations} new quotations"},
            {"type": "order", "count": new_orders, "message": f"{new_orders} orders placed"}
        ]
    })





def run_graph_email_monitor():
    """Poll Microsoft Graph for new quotation emails."""
    log.info("📧 Microsoft Graph email monitor started...")
    from services.graph_email_service import poll_and_process_emails_graph
    while True:
        try:
            with app.app_context():
                processed_count = poll_and_process_emails_graph(
                    db.session, quotation_service, inventory_service
                )
                if processed_count > 0:
                    log.info(f"✅ Processed {processed_count} new quotation emails via Microsoft Graph")
        except Exception as e:
            log.error(f"❌ Graph email monitor error: {e}", exc_info=True)
        time.sleep(Config.CHECK_INTERVAL)




def start_background_services():
    """Start background services including Outlook monitor"""
    try:
       
        monitor_thread = threading.Thread(target=run_graph_email_monitor, daemon=True)
        monitor_thread.start()
        log.info("✅ Started Outlook monitor background thread.")
    except Exception as e:
        log.error(f"❌ Failed to start Outlook monitor: {e}")
        log.info("⚠️ Outlook monitor failed to start, but main application will continue running...")

# ==========================
# 🚀 MAIN ENTRY POINT - FIXED VERSION``
# ==========================
if __name__ == '__main__':
    with app.app_context():
        db.create_all()
        # load_initial_data()

    from services.inventory_service import InventoryService
    from services.quotation_service import QuotationService
    from services.procurement_service import ProcurementService

    inventory_service = InventoryService(db.session)
    quotation_service = QuotationService(db.session, inventory_service)
    procurement_service = ProcurementService(db.session, inventory_service)

    
    start_background_services()

    # Run Flask app without reloader to avoid threading issues
    # app.run(debug=True, host='0.0.0.0', port=5000, use_reloader=False)
    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_DEBUG", "False").lower() == "true"
    app.run(debug=debug, host='0.0.0.0', port=port, use_reloader=False)