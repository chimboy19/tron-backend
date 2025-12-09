
# # backend/utils/xlsx_loader.py
# import sys
# import os
# from datetime import date

# backend_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
# if backend_dir not in sys.path:
#     sys.path.insert(0, backend_dir)

# import os
# from models import db, Product, Supplier, Customer, Stock, Calendar, SupplierLeadTime, FAXHistory, TNZ2Record, IncomingQuotationRequest, User
# import pandas as pd


# def load_initial_data():
#     """Load initial data from XLSX files ONLY if tables are empty."""
#     from models import Product
#     if Product.query.first() is not None:
#         print("✅ Data already loaded. Skipping XLSX import.")
#         return

#     print(" Starting initial data load from XLSX files into the database...")

    
#     load_xlsx_to_db(
#         "data/MUH1.xlsx",
#         Product,
#         {
#             "HCOD": "hcod",
#             "HNMT": "hnm",
#             "HNMM": "description",
#             "HSRS": "category",
#             'TANKAU': 'unit_price',
#             "MCOD": "supplier_code"
#         }
#     )
#     load_xlsx_to_db(
#         "data/MMK1.xlsx",
#         Supplier,
#         {
#             "MCOD": "mcod",
#             "MNM": "name",
#             "ADRNS": "address",
#             "TELS": "phone",
#             "NEBIP": "standard_lead_time"
#         }
#     )
#     load_xlsx_to_db(
#         "data/MUS1.xlsx",
#         Customer,
#         {
#             "UCOD": "ucod",
#             "UNM": "name",
#             "ADRNS": "address",
#             "TELS": "phone"
#         }
#     )


#     load_xlsx_to_db(
#         "data/MZS1.xlsx",
#         Stock,
#         {
#             "HCOD": "product_hcod",
#             "MCODD": "supplier_mcod",
#             "MKRCD": "manufacturer",
#             "SOKO": "warehouse_code",
#             "JZSU": "actual_quantity",
#             "ZSSU": "shelf_quantity",
#             "TANAFL": "location_floor",
#             "TANABL": "location_block",
#             "TANANO": "location_number",
#             "TANAST": "location_stage"
#         },
#         resolve_foreign_keys=True
#     )
#     # load_xlsx_to_db(
#     #     "data/TNZ1.xlsx",
#     #     Stock,
#     #     {
#     #         "HCOD": "product_hcod",
#     #         "MCODD": "supplier_mcod",
#     #         "SOKO": "warehouse_code",
#     #         "JZSU": "actual_quantity",
#     #         "ZSSU": "shelf_quantity",
#     #         "TANAFL": "location_floor",
#     #         "TANABL": "location_block",
#     #         "TANANO": "location_number",
#     #         "TANAST": "location_stage"
#     #     },
#     #     resolve_foreign_keys=True
#     # )
#     load_xlsx_to_db(
#         "data/DSET.xlsx",
#         Calendar,
#         {
#             "DATE": "date",
#             "DATENO": "date_number",
#             "YOBI": "day_of_week",
#             "YASUMI": "is_holiday",
#             "NYSTOP": "is_shipping_stop",
#             "KOSU": "delivery_course",
#             "DATE1": "course_date",
#             "DATE2": "reverse_course_date",
#             "DATE3": "course_next_day",
#             "DSETXA": "logistics_holiday"
#         }
#     )
#     #  RAS1.xlsx lacks HCOD → disabled
#     # load_xlsx_to_db(...)
#     load_xlsx_to_db(
#         "data/RAF1.xlsx",
#         FAXHistory,
#         {
#             "RAF1D": "delete_flag",
#             "FAXDAY": "fax_day",
#             "MCOD": "supplier_code",
#             "UCOD": "customer_code"
#         }
#     )
#     load_xlsx_to_db(
#         "data/TNZ2.xlsx",
#         TNZ2Record,
#         {
#             "TNZ2D": "delete_flag",
#             "DENNO": "document_number",
#             "UCOD": "customer_code",
#             "MCOD": "supplier_code",
#             "HNAME": "item_description",
#             "SURYO": "quantity",
#             "TANKA": "unit_price",
#             "IRINE": "standard_cost",
#             "NODAYU": "customer_delivery_date",
#             "SYDAY": "shipment_date",
#             "HCOD": "product_code"
#         }
#     )

#     print("✅ Initial data load completed.")


# def load_xlsx_to_db(xlsx_path, model_class, column_mapping, resolve_foreign_keys=False):
#     if not os.path.exists(xlsx_path):
#         print(f"File {xlsx_path} does not exist, skipping.")
#         return

#     print(f"Loading data from {xlsx_path} into {model_class.__tablename__}...")
#     try:
#         df = pd.read_excel(xlsx_path)
#         df = df.fillna('').astype(str).replace('nan', '')

       
#         df.rename(columns=column_mapping, inplace=True)

        
#         if model_class == Calendar:
#             for col in ['date', 'course_date', 'reverse_course_date', 'course_next_day']:
#                 if col in df.columns:
#                     df[col] = df[col].apply(yymmdd_to_date)
#             df = df.dropna(subset=['date'])
#             df = df.drop_duplicates(subset=['date'], keep='first')

#         elif model_class == FAXHistory:
#             if 'fax_day' in df.columns:
#                 df['fax_day'] = df['fax_day'].apply(yyyymmdd_to_date)

#         elif model_class == TNZ2Record:
#             for col in ['customer_delivery_date', 'shipment_date']:
#                 if col in df.columns:
#                     df[col] = df[col].apply(yyyymmdd_to_date)

       
#         if model_class == Calendar:
#             for col in ['is_holiday', 'is_shipping_stop', 'logistics_holiday']:
#                 if col in df.columns:
#                     df[col] = df[col].apply(lambda x: x == '1' if x != '' else False)

      
#         required_fields = [
#             col.name for col in model_class.__table__.columns
#             if not col.nullable and col.name in df.columns
#         ]
#         if required_fields:
#             mask = True
#             for field in required_fields:
#                 mask = mask & (df[field] != '')
#             df = df[mask].copy()

#         records = df.to_dict('records')
#         if not records:
#             print(f"✅ Successfully loaded 0 records into {model_class.__tablename__} from {xlsx_path}")
#             return

       
#         if model_class == Product:
#             seen = set()
#             unique_records = []
#             for r in records:
#                 hcod = r.get('hcod')
#                 if hcod and hcod not in seen:
#                     seen.add(hcod)
#                     unique_records.append(r)
#             records = unique_records

       
#         if resolve_foreign_keys:
#             if model_class in (Stock,):
               
#                 product_map = {}
#                 for p in Product.query.all():
                   
#                     key1 = str(p.hcod).strip().upper()
#                     key2 = key1.lstrip('H')  
#                     key3 = f"H{key2}" if not key2.startswith('H') else key2  
#                     product_map[key1] = p.id
#                     product_map[key2] = p.id
#                     product_map[key3] = p.id

#                 supplier_map = {}
#                 for s in Supplier.query.all():
                    
#                     key1 = str(s.mcod).strip().upper()
#                     key2 = key1.rstrip('.0')  
#                     key3 = key1.lstrip('0')   
                    
#                     supplier_map[key1] = s.id
#                     supplier_map[key2] = s.id
#                     supplier_map[key3] = s.id

#                 new_records = []
#                 skipped_records = []
                
#                 for r in records:
#                     hcod_raw = r.pop('product_hcod', None)
#                     mcod_raw = r.pop('supplier_mcod', None)

                    
#                     hcod_norm = str(hcod_raw).strip().upper() if hcod_raw else ''
                   
#                     if hcod_norm.endswith('.0') and len(hcod_norm) > 2:
#                         hcod_norm = hcod_norm[:-2]

#                     hcod_variations = [
#                         hcod_norm,
#                         hcod_norm.lstrip('H'), 
#                         f"H{hcod_norm}" if not hcod_norm.startswith('H') else hcod_norm,  # With H
#                         hcod_norm.zfill(6), 
#                     ]
                    
                   
#                     mcod_norm = str(mcod_raw).rstrip('.0').strip().upper() if mcod_raw else ''
#                     mcod_variations = [
#                         mcod_norm,
#                         mcod_norm.rstrip('.0'),
#                         mcod_norm.lstrip('0'),
#                         mcod_norm.zfill(4),
#                     ]

                   
#                     product_id = None
#                     supplier_id = None
                    
#                     for hcod_var in hcod_variations:
#                         if hcod_var in product_map:
#                             product_id = product_map[hcod_var]
#                             break
                    
#                     for mcod_var in mcod_variations:
#                         if mcod_var in supplier_map:
#                             supplier_id = supplier_map[mcod_var]
#                             break

#                     if product_id and supplier_id:
#                         r['product_id'] = product_id
#                         r['supplier_id'] = supplier_id
#                         new_records.append(r)
#                     else:
#                         skipped_records.append({
#                             'hcod_raw': hcod_raw,
#                             'hcod_norm': hcod_norm,
#                             'mcod_raw': mcod_raw,
#                             'mcod_norm': mcod_norm
#                         })

#                 if skipped_records:
#                     print(f"⚠️ Skipped {len(skipped_records)} records due to FK mismatches")
                   
#                     for i, skipped in enumerate(skipped_records[:10]):
#                         print(f"  {i+1}. HCOD: '{skipped['hcod_raw']}'->'{skipped['hcod_norm']}', MCOD: '{skipped['mcod_raw']}'->'{skipped['mcod_norm']}'")
#                     if len(skipped_records) > 10:
#                         print(f"  ... and {len(skipped_records) - 10} more")
                
#                 records = new_records

#             elif model_class in (FAXHistory, TNZ2Record):
                
#                 customer_map = {}
#                 for c in Customer.query.all():
#                     key = str(c.ucod).strip().upper()
#                     customer_map[key] = c.id
                
#                 supplier_map = {}
#                 for s in Supplier.query.all():
#                     key = str(s.mcod).strip().upper()
#                     supplier_map[key] = s.id

#                 new_records = []
#                 for r in records:
#                     ucod = r.pop('customer_code', None)
#                     mcod = r.pop('supplier_code', None)
                    
#                     if ucod and mcod and ucod in customer_map and mcod in supplier_map:
#                         r['customer_id'] = customer_map[ucod]
#                         r['supplier_id'] = supplier_map[mcod]
#                         if model_class == TNZ2Record:
#                             hcod = r.pop('product_code', None)
#                             if hcod:
                               
#                                 product = Product.query.filter(
#                                     db.or_(
#                                         Product.hcod == hcod,
#                                         Product.hcod == hcod.lstrip('H'),
#                                         Product.hcod == f"H{hcod}" if not hcod.startswith('H') else hcod
#                                     )
#                                 ).first()
#                                 if product:
#                                     r['product_id'] = product.id
#                         new_records.append(r)
#                 records = new_records

#         if not records:
#             print(f"✅ Successfully loaded 0 records into {model_class.__tablename__} from {xlsx_path} (FK resolution failed)")
#             return

#         db.session.bulk_insert_mappings(model_class, records)
#         db.session.commit()
#         print(f"✅ Successfully loaded {len(records)} records into {model_class.__tablename__} from {xlsx_path}")

#     except Exception as e:
#         print(f" Error loading {xlsx_path}: {e}")
#         import traceback
#         traceback.print_exc()
#         db.session.rollback()


# def yymmdd_to_date(val):
#     """Convert YYMMDD string to date (assumes 2000s for years 00-49)."""
#     if not val or val in ('', '0', 'nan'):
#         return None
#     try:
#         s = str(val).zfill(6)
#         year = int(s[:2])
#         year = 2000 + year if year < 50 else 1900 + year
#         return date(year, int(s[2:4]), int(s[4:6]))
#     except:
#         return None


# def yyyymmdd_to_date(val):
#     """Convert YYYYMMDD string to date."""
#     if not val or val in ('', '0', 'nan'):
#         return None
#     try:
#         s = str(val)
#         return date(int(s[:4]), int(s[4:6]), int(s[6:8]))
#     except:
#         return None


# def generate_new_ucod(existing_ucods: set) -> str:
#     """Generate a new unique customer code."""
#     existing_ints = {int(u) for u in existing_ucods if str(u).isdigit()}
#     new_id = 1001
#     while new_id in existing_ints:
#         new_id += 1
#     return str(new_id)


# if __name__ == "__main__":
#     from app import app, db
#     with app.app_context():
#         db.create_all()
#         load_initial_data()








# backend/utils/xlsx_loader.py
import sys
import os
from datetime import date

backend_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if backend_dir not in sys.path:
    sys.path.insert(0, backend_dir)

import pandas as pd
from models import db, Product, Supplier, Customer, Stock, Calendar, FAXHistory, TNZ2Record


def load_initial_data():
    """Load initial data from XLSX files ONLY if tables are empty."""
    if Product.query.first() is not None:
        print("✅ Data already loaded. Skipping XLSX import.")
        return

    print("Starting initial data load from XLSX files into the database...")

    # 1. Load reference tables FIRST
    load_xlsx_to_db(
        "data/MMK1.xlsx",
        Supplier,
        {
            "MCOD": "mcod",
            "MNM": "name",
            "ADRNS": "address",
            "TELS": "phone",
            "NEBIP": "standard_lead_time"
        }
    )
    # Debug: print loaded supplier codes
    supplier_codes = {s.mcod for s in Supplier.query.all()}
    print(f"✅ Loaded {len(supplier_codes)} supplier codes (sample): {sorted(list(supplier_codes))[:5]}")

    load_xlsx_to_db(
        "data/MUS1.xlsx",
        Customer,
        {
            "UCOD": "ucod",
            "UNM": "name",
            "ADRNS": "address",
            "TELS": "phone"
        }
    )
    # Debug: print loaded customer codes
    customer_codes = {c.ucod for c in Customer.query.all()}
    print(f"✅ Loaded {len(customer_codes)} customer codes (sample): {sorted(list(customer_codes))[:5]}")

    # 2. Load dependent tables
    load_xlsx_to_db(
        "data/MUH1.xlsx",
        Product,
        {
            "HCOD": "hcod",
            "HNMT": "hnm",
            "HNMM": "description",
            "HSRS": "category",
            'TANKAU': 'unit_price',
            "MCOD": "supplier_code"
        }
    )

    load_xlsx_to_db(
        "data/MZS1.xlsx",
        Stock,
        {
            "HCOD": "product_hcod",
            "MCODD": "supplier_mcod",
            "MKRCD": "manufacturer",
            "SOKO": "warehouse_code",
            "JZSU": "actual_quantity",
            "ZSSU": "shelf_quantity",
            "TANAFL": "location_floor",
            "TANABL": "location_block",
            "TANANO": "location_number",
            "TANAST": "location_stage"
        },
        resolve_foreign_keys=True
    )

    load_xlsx_to_db(
        "data/DSET.xlsx",
        Calendar,
        {
            "DATE": "date",
            "DATENO": "date_number",
            "YOBI": "day_of_week",
            "YASUMI": "is_holiday",
            "NYSTOP": "is_shipping_stop",
            "KOSU": "delivery_course",
            "DATE1": "course_date",
            "DATE2": "reverse_course_date",
            "DATE3": "course_next_day",
            "DSETXA": "logistics_holiday"
        }
    )

    load_xlsx_to_db(
        "data/RAF1.xlsx",
        FAXHistory,
        {
            "RAF1D": "delete_flag",
            "FAXDAY": "fax_day",
            "MCOD": "supplier_code",
            "UCOD": "customer_code"
        },
        resolve_foreign_keys=True  # Enable FK resolution
    )

    load_xlsx_to_db(
        "data/TNZ2.xlsx",
        TNZ2Record,
        {
            "TNZ2D": "delete_flag",
            "DENNO": "document_number",
            "UCOD": "customer_code",
            "MCOD": "supplier_code",
            "HNAME": "item_description",
            "SURYO": "quantity",
            "TANKA": "unit_price",
            "IRINE": "standard_cost",
            "NODAYU": "customer_delivery_date",
            "SYDAY": "shipment_date",
            "HCOD": "product_code"
        },
        resolve_foreign_keys=True  # Enable FK resolution
    )

    print("✅ Initial data load completed.")


def load_xlsx_to_db(xlsx_path, model_class, column_mapping, resolve_foreign_keys=False):
    if not os.path.exists(xlsx_path):
        print(f"File {xlsx_path} does not exist, skipping.")
        return

    print(f"Loading data from {xlsx_path} into {model_class.__tablename__}...")
    try:
        df = pd.read_excel(xlsx_path)
        df = df.fillna('').astype(str).replace('nan', '')
        df.rename(columns=column_mapping, inplace=True)

        # Date conversions
        if model_class == Calendar:
            for col in ['date', 'course_date', 'reverse_course_date', 'course_next_day']:
                if col in df.columns:
                    df[col] = df[col].apply(yymmdd_to_date)
            df = df.dropna(subset=['date']).drop_duplicates(subset=['date'], keep='first')
            for col in ['is_holiday', 'is_shipping_stop', 'logistics_holiday']:
                if col in df.columns:
                    df[col] = df[col].apply(lambda x: x == '1' if x != '' else False)

        elif model_class == FAXHistory and 'fax_day' in df.columns:
            df['fax_day'] = df['fax_day'].apply(yyyymmdd_to_date)

        elif model_class == TNZ2Record:
            for col in ['customer_delivery_date', 'shipment_date']:
                if col in df.columns:
                    df[col] = df[col].apply(yyyymmdd_to_date)

        # Required field filtering
        required_fields = [
            col.name for col in model_class.__table__.columns
            if not col.nullable and col.name in df.columns
        ]
        if required_fields:
            mask = df[required_fields].ne('').all(axis=1)
            df = df[mask].copy()

        records = df.to_dict('records')
        if not records:
            print(f"✅ Successfully loaded 0 records into {model_class.__tablename__} from {xlsx_path}")
            return

        # Deduplication for Product
        if model_class == Product:
            seen = set()
            unique_records = []
            for r in records:
                hcod = r.get('hcod')
                if hcod and hcod not in seen:
                    seen.add(hcod)
                    unique_records.append(r)
            records = unique_records

        # Resolve foreign keys if requested
        if resolve_foreign_keys:
            if model_class == Stock:
                # Reuse your existing logic (unchanged)
                product_map = {}
                for p in Product.query.all():
                    key1 = str(p.hcod).strip().upper()
                    key2 = key1.lstrip('H')
                    key3 = f"H{key2}" if not key2.startswith('H') else key2
                    product_map[key1] = p.id
                    product_map[key2] = p.id
                    product_map[key3] = p.id

                supplier_map = {}
                for s in Supplier.query.all():
                    key1 = str(s.mcod).strip().upper()
                    key2 = key1.rstrip('.0')
                    key3 = key1.lstrip('0')
                    supplier_map[key1] = s.id
                    supplier_map[key2] = s.id
                    supplier_map[key3] = s.id

                new_records = []
                skipped_records = []
                for r in records:
                    hcod_raw = r.pop('product_hcod', None)
                    mcod_raw = r.pop('supplier_mcod', None)

                    hcod_norm = str(hcod_raw).strip().upper() if hcod_raw else ''
                    if hcod_norm.endswith('.0'):
                        hcod_norm = hcod_norm[:-2]
                    hcod_variations = [hcod_norm, hcod_norm.lstrip('H'), f"H{hcod_norm}" if not hcod_norm.startswith('H') else hcod_norm, hcod_norm.zfill(6)]

                    mcod_norm = str(mcod_raw).rstrip('.0').strip().upper() if mcod_raw else ''
                    mcod_variations = [mcod_norm, mcod_norm.rstrip('.0'), mcod_norm.lstrip('0'), mcod_norm.zfill(4)]

                    product_id = next((product_map[h] for h in hcod_variations if h in product_map), None)
                    supplier_id = next((supplier_map[m] for m in mcod_variations if m in supplier_map), None)

                    if product_id and supplier_id:
                        r['product_id'] = product_id
                        r['supplier_id'] = supplier_id
                        new_records.append(r)
                    else:
                        skipped_records.append({'hcod_raw': hcod_raw, 'mcod_raw': mcod_raw})

                if skipped_records:
                    print(f"⚠️ Skipped {len(skipped_records)} records due to FK mismatches")
                    for i, skipped in enumerate(skipped_records[:10]):
                        print(f"  {i+1}. HCOD: '{skipped['hcod_raw']}', MCOD: '{skipped['mcod_raw']}'")
                    if len(skipped_records) > 10:
                        print(f"  ... and {len(skipped_records) - 10} more")
                records = new_records

            elif model_class in (FAXHistory, TNZ2Record):
                # Build FK maps
                customer_map = {str(c.ucod).strip().upper(): c.id for c in Customer.query.all()}
                supplier_map = {str(s.mcod).strip().upper(): s.id for s in Supplier.query.all()}

                new_records = []
                missing_customers = set()
                missing_suppliers = set()

                for r in records:
                    ucod = str(r.pop('customer_code', '')).strip().upper()
                    mcod = str(r.pop('supplier_code', '')).strip().upper()

                    if ucod not in customer_map:
                        missing_customers.add(ucod)
                        continue
                    if mcod not in supplier_map:
                        missing_suppliers.add(mcod)
                        continue

                    r['customer_id'] = customer_map[ucod]
                    r['supplier_id'] = supplier_map[mcod]

                    # Resolve product for TNZ2Record
                    if model_class == TNZ2Record:
                        hcod = r.pop('product_code', None)
                        if hcod:
                            product = Product.query.filter(
                                db.or_(
                                    Product.hcod == hcod,
                                    Product.hcod == hcod.lstrip('H'),
                                    Product.hcod == f"H{hcod}" if not hcod.startswith('H') else hcod
                                )
                            ).first()
                            if product:
                                r['product_id'] = product.id
                    new_records.append(r)

                if missing_customers:
                    print(f"⚠️ Unknown customer codes in {xlsx_path}: {sorted(missing_customers)[:5]}")
                if missing_suppliers:
                    print(f"⚠️ Unknown supplier codes in {xlsx_path}: {sorted(missing_suppliers)[:5]}")

                records = new_records

        if not records:
            print(f"✅ Successfully loaded 0 records into {model_class.__tablename__} from {xlsx_path} (FK resolution failed)")
            return

        db.session.bulk_insert_mappings(model_class, records)
        db.session.commit()
        print(f"✅ Successfully loaded {len(records)} records into {model_class.__tablename__} from {xlsx_path}")

    except Exception as e:
        print(f" Error loading {xlsx_path}: {e}")
        import traceback
        traceback.print_exc()
        db.session.rollback()


def yymmdd_to_date(val):
    if not val or val in ('', '0', 'nan'):
        return None
    try:
        s = str(val).zfill(6)
        year = int(s[:2])
        year = 2000 + year if year < 50 else 1900 + year
        return date(year, int(s[2:4]), int(s[4:6]))
    except:
        return None


def yyyymmdd_to_date(val):
    if not val or val in ('', '0', 'nan'):
        return None
    try:
        s = str(val)
        return date(int(s[:4]), int(s[4:6]), int(s[6:8]))
    except:
        return None

def generate_new_ucod(existing_ucods: set) -> str:
    """Generate a new unique customer code."""
    existing_ints = {int(u) for u in existing_ucods if str(u).isdigit()}
    new_id = 1001
    while new_id in existing_ints:
        new_id += 1
    return str(new_id)




if __name__ == "__main__":
    from app import app
    with app.app_context():
        db.create_all()
        load_initial_data()