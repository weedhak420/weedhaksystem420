import os
import datetime
import uuid
import shutil
import random
import pandas as pd
from decimal import Decimal
from threading import Lock
from flask import (Blueprint, render_template, request, redirect, url_for, flash, 
                   session, jsonify, send_file, current_app)
from werkzeug.security import generate_password_hash, check_password_hash
from sqlalchemy import func, desc, and_, or_, text
from app import db
from app.models import (User, Customer, Product, Order, OrderItem, InventoryEntry, 
                        ActivityLog, Notification, Settings, ShopSettings, Banner, 
                        FeaturedProduct, Store)
from app.utils import (save_file, delete_file, log_activity, create_notification, 
                       generate_qr_code, generate_qr_code_base64, add_product_to_google_sheets,
                       update_product_in_google_sheets, sync_all_products_to_google_sheets,
                       validate_api_key, add_product_to_google_sheets_realtime,
                       create_n8n_webhook_data, send_to_n8n_webhook)

main = Blueprint('main', __name__)
stock_update_lock = Lock()

# API Authentication decorator
def require_api_key(f):
    """Decorator to require API key for API endpoints"""
    def decorated_function(*args, **kwargs):
        api_key = request.headers.get('X-API-Key') or request.args.get('api_key')
        if not api_key or not validate_api_key(api_key):
            return jsonify({'error': 'Invalid or missing API key'}), 401
        return f(*args, **kwargs)
    decorated_function.__name__ = f.__name__
    return decorated_function

@main.context_processor
def inject_shop_settings():
    shop_settings = ShopSettings.query.first()
    if not shop_settings:
        store = Store.query.first()
        if not store:
            store = Store(name='Default Store')
            db.session.add(store)
            db.session.commit()
        shop_settings = ShopSettings(store_id=store.id)
        db.session.add(shop_settings)
        db.session.commit()
    return dict(shop_settings=shop_settings)

@main.context_processor
def inject_notifications():
    if 'user_id' in session:
        unread_count = Notification.query.filter_by(user_id=session['user_id'], is_read=False).count()
        recent_notifications = Notification.query.filter_by(user_id=session['user_id']).order_by(Notification.created_at.desc()).limit(5).all()
        return dict(unread_notifications_count=unread_count, recent_notifications=recent_notifications)
    return dict(unread_notifications_count=0, recent_notifications=[])

# API Routes for Product Management
@main.route('/api/products', methods=['GET'])
@require_api_key
def api_get_products():
    """API endpoint to get all products"""
    try:
        products = Product.query.all()
        products_data = []
        
        for product in products:
            products_data.append({
                'id': product.id,
                'name': product.name,
                'flavor': product.flavor,
                'description': product.description,
                'price': float(product.price),
                'cost': float(product.cost),
                'wholesale_price': float(product.wholesale_price) if product.wholesale_price else None,
                'stock': product.stock,
                'barcode': product.barcode,
                'image_path': product.image_path,
                'created_at': product.created_at.isoformat() if product.created_at else None,
                'profit_margin': product.profit_margin
            })
        
        return jsonify({
            'success': True,
            'products': products_data,
            'total': len(products_data)
        })
        
    except Exception as e:
        current_app.logger.error(f"API get products error: {str(e)}")
        return jsonify({'success': False, 'error': str(e)}), 500

@main.route('/api/products', methods=['POST'])
@require_api_key
def api_add_product():
    """API endpoint to add a new product"""
    try:
        data = request.get_json()
        
        # Validate required fields
        required_fields = ['name', 'flavor', 'price', 'cost']
        for field in required_fields:
            if field not in data:
                return jsonify({'success': False, 'error': f'Missing required field: {field}'}), 400
        
        # Create new product
        product = Product(
            name=data['name'],
            flavor=data['flavor'],
            description=data.get('description', ''),
            price=float(data['price']),
            cost=float(data['cost']),
            wholesale_price=float(data.get('wholesale_price')) if data.get('wholesale_price') else None,
            stock=int(data.get('stock', 0)),
            barcode=data.get('barcode', '')
        )
        
        db.session.add(product)
        db.session.commit()
        
        # Prepare data for Google Sheets
        product_data = {
            'id': product.id,
            'name': product.name,
            'flavor': product.flavor,
            'description': product.description,
            'price': float(product.price),
            'cost': float(product.cost),
            'stock': product.stock,
            'barcode': product.barcode or ''
        }
        
        # Sync to Google Sheets
        sheets_success = add_product_to_google_sheets(product_data)
        
        # Log activity
        log_activity('api_add_product', 'product', product.id, f'Added product via API: {product.name} {product.flavor}')
        
        # Create notification
        create_notification(f'สินค้าใหม่ถูกเพิ่มผ่าน API: {product.name} {product.flavor}', 'new_product', product.id)
        
        return jsonify({
            'success': True,
            'message': 'Product added successfully',
            'product': {
                'id': product.id,
                'name': product.name,
                'flavor': product.flavor,
                'description': product.description,
                'price': float(product.price),
                'cost': float(product.cost),
                'wholesale_price': float(product.wholesale_price) if product.wholesale_price else None,
                'stock': product.stock,
                'barcode': product.barcode,
                'created_at': product.created_at.isoformat(),
                'profit_margin': product.profit_margin
            },
            'google_sheets_synced': sheets_success
        }), 201
        
    except ValueError as e:
        return jsonify({'success': False, 'error': f'Invalid data type: {str(e)}'}), 400
    except Exception as e:
        db.session.rollback()
        current_app.logger.error(f"API add product error: {str(e)}")
        return jsonify({'success': False, 'error': str(e)}), 500

@main.route('/api/products/<int:product_id>', methods=['PUT'])
@require_api_key
def api_update_product(product_id):
    """API endpoint to update a product"""
    try:
        product = Product.query.get_or_404(product_id)
        data = request.get_json()
        
        # Update product fields
        if 'name' in data:
            product.name = data['name']
        if 'flavor' in data:
            product.flavor = data['flavor']
        if 'description' in data:
            product.description = data['description']
        if 'price' in data:
            product.price = float(data['price'])
        if 'cost' in data:
            product.cost = float(data['cost'])
        if 'wholesale_price' in data:
            product.wholesale_price = float(data['wholesale_price']) if data['wholesale_price'] else None
        if 'stock' in data:
            product.stock = int(data['stock'])
        if 'barcode' in data:
            product.barcode = data['barcode']
        
        db.session.commit()
        
        # Prepare data for Google Sheets
        product_data = {
            'id': product.id,
            'name': product.name,
            'flavor': product.flavor,
            'description': product.description,
            'price': float(product.price),
            'cost': float(product.cost),
            'stock': product.stock,
            'barcode': product.barcode or ''
        }
        
        # Sync to Google Sheets
        sheets_success = update_product_in_google_sheets(product.id, product_data)
        
        # Log activity
        log_activity('api_update_product', 'product', product.id, f'Updated product via API: {product.name} {product.flavor}')
        
        return jsonify({
            'success': True,
            'message': 'Product updated successfully',
            'product': {
                'id': product.id,
                'name': product.name,
                'flavor': product.flavor,
                'description': product.description,
                'price': float(product.price),
                'cost': float(product.cost),
                'wholesale_price': float(data['wholesale_price']) if data.get('wholesale_price') else None,
                'stock': product.stock,
                'barcode': product.barcode,
                'created_at': product.created_at.isoformat() if product.created_at else None,
                'profit_margin': product.profit_margin
            },
            'google_sheets_synced': sheets_success
        })
        
    except ValueError as e:
        return jsonify({'success': False, 'error': f'Invalid data type: {str(e)}'}), 400
    except Exception as e:
        db.session.rollback()
        current_app.logger.error(f"API update product error: {str(e)}")
        return jsonify({'success': False, 'error': str(e)}), 500

@main.route('/api/products/<int:product_id>', methods=['DELETE'])
@require_api_key
def api_delete_product(product_id):
    """API endpoint to delete a product"""
    try:
        product = Product.query.get_or_404(product_id)
        product_name = f"{product.name} {product.flavor}"
        
        # Delete associated image if exists
        if product.image_path:
            delete_file(product.image_path)
        
        # Log activity before deletion
        log_activity('api_delete_product', 'product', product.id, f'Deleted product via API: {product_name}')
        
        db.session.delete(product)
        db.session.commit()
        
        # Note: We don't delete from Google Sheets automatically to preserve data
        # You might want to add a "deleted" column instead
        
        return jsonify({
            'success': True,
            'message': f'Product "{product_name}" deleted successfully'
        })
        
    except Exception as e:
        db.session.rollback()
        current_app.logger.error(f"API delete product error: {str(e)}")
        return jsonify({'success': False, 'error': str(e)}), 500

@main.route('/api/products/sync-to-sheets', methods=['POST'])
@require_api_key
def api_sync_products_to_sheets():
    """API endpoint to sync all products to Google Sheets"""
    try:
        success = sync_all_products_to_google_sheets()
        
        if success:
            log_activity('api_sync_products', details='Synced all products to Google Sheets via API')
            return jsonify({
                'success': True,
                'message': 'All products synced to Google Sheets successfully'
            })
        else:
            return jsonify({
                'success': False,
                'error': 'Failed to sync products to Google Sheets'
            }), 500
            
    except Exception as e:
        current_app.logger.error(f"API sync products error: {str(e)}")
        return jsonify({'success': False, 'error': str(e)}), 500

@main.route('/api/products/search', methods=['GET'])
@require_api_key
def api_search_products_authenticated():
    """API endpoint to search products"""
    try:
        query = request.args.get('q', '').strip()
        limit = min(int(request.args.get('limit', 50)), 100)  # Max 100 results
        
        if not query:
            return jsonify({'success': False, 'error': 'Search query is required'}), 400
        
        # Search products by name, flavor, or barcode
        products = Product.query.filter(
            or_(
                Product.name.ilike(f'%{query}%'),
                Product.flavor.ilike(f'%{query}%'),
                Product.barcode.ilike(f'%{query}%') if query else False
            )
        ).limit(limit).all()
        
        products_data = []
        for product in products:
            products_data.append({
                'id': product.id,
                'name': product.name,
                'flavor': product.flavor,
                'description': product.description,
                'price': float(product.price),
                'cost': float(product.cost),
                'wholesale_price': float(product.wholesale_price) if product.wholesale_price else None,
                'stock': product.stock,
                'barcode': product.barcode,
                'image_path': product.image_path,
                'profit_margin': product.profit_margin
            })
        
        return jsonify({
            'success': True,
            'products': products_data,
            'total': len(products_data),
            'query': query
        })
        
    except Exception as e:
        current_app.logger.error(f"API search products error: {str(e)}")
        return jsonify({'success': False, 'error': str(e)}), 500

@main.route('/api/products/low-stock', methods=['GET'])
@require_api_key
def api_get_low_stock_products():
    """API endpoint to get low stock products"""
    try:
        settings = Settings.query.first()
        threshold = int(request.args.get('threshold', settings.low_stock_threshold if settings else 10))
        
        products = Product.query.filter(Product.stock <= threshold).all()
        
        products_data = []
        for product in products:
            products_data.append({
                'id': product.id,
                'name': product.name,
                'flavor': product.flavor,
                'stock': product.stock,
                'price': float(product.price),
                'cost': float(product.cost),
                'barcode': product.barcode
            })
        
        return jsonify({
            'success': True,
            'products': products_data,
            'total': len(products_data),
            'threshold': threshold
        })
        
    except Exception as e:
        current_app.logger.error(f"API low stock products error: {str(e)}")
        return jsonify({'success': False, 'error': str(e)}), 500

# Existing routes continue here...
@main.route('/')
def index():
    shop_settings = ShopSettings.query.first()
    # Banners for stock page
    shop_stock_banners_top = Banner.query.filter_by(page_location='shop_stock', position='top', is_active=True).all()
    shop_stock_banners_middle = Banner.query.filter_by(page_location='shop_stock', position='middle', is_active=True).all()
    shop_stock_banners_bottom = Banner.query.filter_by(page_location='shop_stock', position='bottom', is_active=True).all()
    
    return render_template('shop_stock.html', 
                           shop_settings=shop_settings,
                           shop_stock_banners_top=shop_stock_banners_top,
                           shop_stock_banners_middle=shop_stock_banners_middle,
                           shop_stock_banners_bottom=shop_stock_banners_bottom)

@main.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        
        user = User.query.filter_by(username=username).first()
        
        if user and check_password_hash(user.password, password):
            session['user_id'] = user.id
            session['username'] = user.username
            session['role'] = user.role
            
            log_activity('login')
            flash('เข้าสู่ระบบสำเร็จ', 'success')
            return redirect(url_for('main.dashboard'))
        
        flash('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง', 'danger')
    
    return render_template('login.html')

@main.route('/logout')
def logout():
    if 'user_id' in session:
        log_activity('logout')
    
    session.clear()
    flash('ออกจากระบบสำเร็จ', 'success')
    return redirect(url_for('main.login'))

@main.route('/dashboard')
def dashboard():
    if 'user_id' not in session:
        return redirect(url_for('main.login'))
    
    total_sales = db.session.query(func.sum(Order.total_amount)).scalar() or 0
    pending_orders = Order.query.filter_by(payment_status='pending').count()
    
    settings = Settings.query.first()
    low_stock_threshold = settings.low_stock_threshold if settings else 10
    low_stock_products = Product.query.filter(Product.stock < low_stock_threshold).count()
    low_stock_products_list = Product.query.filter(Product.stock < low_stock_threshold).all()
    
    recent_orders = Order.query.order_by(Order.order_date.desc()).limit(5).all()
    
    return render_template('dashboard.html', 
                          total_sales=total_sales,
                          pending_orders=pending_orders,
                          low_stock_products=low_stock_products,
                          low_stock_products_list=low_stock_products_list,
                          recent_orders=recent_orders)

# API Routes for notifications and sales trend
@main.route('/api/notifications/unread')
def api_notifications_unread():
    if 'user_id' not in session:
        return jsonify({'error': 'Unauthorized'}), 401
    
    notifications = Notification.query.filter_by(
        user_id=session['user_id'], 
        is_read=False
    ).order_by(Notification.created_at.desc()).limit(10).all()
    
    count = Notification.query.filter_by(user_id=session['user_id'], is_read=False).count()
    
    notifications_data = []
    for notification in notifications:
        notifications_data.append({
            'id': notification.id,
            'message': notification.message,
            'type': notification.type,
            'related_id': notification.related_id,
            'created_at': notification.created_at.strftime('%Y-%m-%d %H:%M'),
            'is_read': notification.is_read
        })
    
    return jsonify({
        'count': count,
        'notifications': notifications_data
    })

@main.route('/api/sales_trend')
def api_sales_trend():
    if 'user_id' not in session:
        return jsonify({'error': 'Unauthorized'}), 401
    
    try:
        today = datetime.datetime.now().date()
        
        # Daily sales for last 30 days with real data
        daily_dates = []
        daily_sales = []
        for i in range(30):
            date = today - datetime.timedelta(days=29-i)
            daily_dates.append(date.strftime('%m/%d'))
            
            # Get actual sales for this date
            sales = db.session.query(func.sum(Order.total_amount)).filter(
                func.date(Order.order_date) == date,
                Order.payment_status == 'paid'
            ).scalar() or 0
            daily_sales.append(float(sales))
        
        # Weekday sales analysis (0=Monday, 6=Sunday) - real data
        weekday_sales = []
        for weekday in range(7):
            # Get average sales for each weekday over the last 12 weeks
            twelve_weeks_ago = today - datetime.timedelta(weeks=12)
            
            # Calculate which weekday number this corresponds to in MySQL (1=Sunday, 2=Monday, etc.)
            mysql_weekday = (weekday + 2) % 7 + 1
            
            avg_sales = db.session.query(func.avg(Order.total_amount)).filter(
                func.dayofweek(Order.order_date) == mysql_weekday,
                func.date(Order.order_date) >= twelve_weeks_ago,
                Order.payment_status == 'paid'
            ).scalar() or 0
            weekday_sales.append(float(avg_sales))
        
        # Hourly sales pattern - real data from last 30 days
        hourly_sales = []
        thirty_days_ago = today - datetime.timedelta(days=30)
        
        for hour in range(24):
            avg_sales = db.session.query(func.avg(Order.total_amount)).filter(
                func.hour(Order.order_date) == hour,
                func.date(Order.order_date) >= thirty_days_ago,
                Order.payment_status == 'paid'
            ).scalar() or 0
            hourly_sales.append(float(avg_sales))
        
        # Monthly sales data for forecast - last 12 months
        forecast_dates = []
        actual_sales = []
        
        for i in range(12):
            # Calculate the date for i months ago
            if today.month - i <= 0:
                month = today.month - i + 12
                year = today.year - 1
            else:
                month = today.month - i
                year = today.year
            
            month_date = datetime.date(year, month, 1)
            forecast_dates.insert(0, month_date.strftime('%Y-%m'))
            
            # Get actual sales for this month
            if month == 12:
                next_month_date = datetime.date(year + 1, 1, 1)
            else:
                next_month_date = datetime.date(year, month + 1, 1)
            
            month_sales = db.session.query(func.sum(Order.total_amount)).filter(
                func.date(Order.order_date) >= month_date,
                func.date(Order.order_date) < next_month_date,
                Order.payment_status == 'paid'
            ).scalar() or 0
            actual_sales.insert(0, float(month_sales))
        
        # Generate forecast based on trend analysis
        forecast_sales = [None] * 12  # Initialize with None for actual months
        
        # Calculate trend using linear regression on last 6 months
        recent_sales = [x for x in actual_sales[-6:] if x > 0]
        if len(recent_sales) >= 3:
            # Simple linear trend calculation
            x_values = list(range(len(recent_sales)))
            n = len(recent_sales)
            sum_x = sum(x_values)
            sum_y = sum(recent_sales)
            sum_xy = sum(x * y for x, y in zip(x_values, recent_sales))
            sum_x2 = sum(x * x for x in x_values)
            
            # Calculate slope and intercept
            slope = (n * sum_xy - sum_x * sum_y) / (n * sum_x2 - sum_x * sum_x) if (n * sum_x2 - sum_x * sum_x) != 0 else 0
            intercept = (sum_y - slope * sum_x) / n
            
            # Generate forecast for next 3 months
            for i in range(3):
                forecast_value = intercept + slope * (len(recent_sales) + i)
                # Add some seasonality based on historical patterns
                month_index = (len(actual_sales) + i) % 12
                seasonal_factor = 1.0
                
                # Calculate seasonal factor based on historical data
                if len(actual_sales) >= 12:
                    same_month_sales = [actual_sales[j] for j in range(month_index, len(actual_sales), 12) if actual_sales[j] > 0]
                    if same_month_sales:
                        avg_same_month = sum(same_month_sales) / len(same_month_sales)
                        overall_avg = sum(x for x in actual_sales if x > 0) / len([x for x in actual_sales if x > 0])
                        seasonal_factor = avg_same_month / overall_avg if overall_avg > 0 else 1.0
                
                forecast_value *= seasonal_factor
                forecast_sales.append(max(0, forecast_value))  # Ensure non-negative
                
                # Add corresponding date
                future_month = today.month + i + 1
                future_year = today.year
                if future_month > 12:
                    future_month -= 12
                    future_year += 1
                
                future_date = datetime.date(future_year, future_month, 1)
                forecast_dates.append(future_date.strftime('%Y-%m'))
                actual_sales.append(None)
        else:
            # Fallback: simple average-based forecast
            avg_recent = sum(recent_sales) / len(recent_sales) if recent_sales else 0
            for i in range(3):
                forecast_sales.append(avg_recent * (1 + random.uniform(-0.1, 0.2)))
                
                future_month = today.month + i + 1
                future_year = today.year
                if future_month > 12:
                    future_month -= 12
                    future_year += 1
                
                future_date = datetime.date(future_year, future_month, 1)
                forecast_dates.append(future_date.strftime('%Y-%m'))
                actual_sales.append(None)
        
        # Forecast metrics
        forecast_next_month = forecast_sales[-3] if len(forecast_sales) > 2 else 0
        current_month_sales = actual_sales[-1] if actual_sales and actual_sales[-1] is not None else 0
        
        # If current month sales is 0, use last month with data
        if current_month_sales == 0:
            for i in range(len(actual_sales) - 1, -1, -1):
                if actual_sales[i] is not None and actual_sales[i] > 0:
                    current_month_sales = actual_sales[i]
                    break
        
        forecast_change_percentage = 0
        if current_month_sales > 0:
            forecast_change_percentage = ((forecast_next_month - current_month_sales) / current_month_sales * 100)
        
        # Top products forecast based on actual sales velocity
        thirty_days_ago = today - datetime.timedelta(days=30)
        sixty_days_ago = today - datetime.timedelta(days=60)
        
        # Get products with sales in last 30 days
        recent_product_sales = db.session.query(
            Product.id,
            Product.name,
            Product.flavor,
            func.sum(OrderItem.quantity).label('recent_quantity'),
            func.sum(OrderItem.price * OrderItem.quantity).label('recent_revenue')
        ).join(OrderItem).join(Order).filter(
            func.date(Order.order_date) >= thirty_days_ago,
            Order.payment_status == 'paid'
        ).group_by(Product.id).all()
        
        # Get products with sales in previous 30 days (30-60 days ago)
        previous_product_sales = db.session.query(
            Product.id,
            func.sum(OrderItem.quantity).label('previous_quantity')
        ).join(OrderItem).join(Order).filter(
            func.date(Order.order_date) >= sixty_days_ago,
            func.date(Order.order_date) < thirty_days_ago,
            Order.payment_status == 'paid'
        ).group_by(Product.id).all()
        
        # Create lookup for previous sales
        previous_sales_lookup = {item.id: float(item.previous_quantity) for item in previous_product_sales}
        
        # Calculate growth rate and forecast
        forecast_top_products = []
        for product in recent_product_sales:
            recent_qty = float(product.recent_quantity)
            previous_qty = previous_sales_lookup.get(product.id, 0)
            
            # Calculate growth rate
            growth_rate = 0
            if previous_qty > 0:
                growth_rate = (recent_qty - previous_qty) / previous_qty
            elif recent_qty > 0:
                growth_rate = 1.0  # New product with sales
            
            # Score based on recent sales volume and growth
            score = recent_qty * (1 + max(0, growth_rate))
            
            forecast_top_products.append({
                'name': product.name,
                'flavor': product.flavor,
                'recent_quantity': recent_qty,
                'growth_rate': growth_rate,
                'score': score
            })
        
        # Sort by score and take top 5
        forecast_top_products = sorted(forecast_top_products, key=lambda x: x['score'], reverse=True)[:5]
        
        return jsonify({
            'daily_dates': daily_dates,
            'daily_sales': daily_sales,
            'weekday_sales': weekday_sales,
            'hourly_sales': hourly_sales,
            'forecast_dates': forecast_dates,
            'actual_sales': actual_sales,
            'forecast_sales': forecast_sales,
            'forecast_next_month': forecast_next_month,
            'forecast_change_percentage': forecast_change_percentage,
            'forecast_top_products': forecast_top_products
        })
    
    except Exception as e:
        current_app.logger.error(f"Sales trend API error: {str(e)}")
        return jsonify({'error': 'Internal server error'}), 500

@main.route('/api/customer_analysis')
def api_customer_analysis():
    if 'user_id' not in session:
        return jsonify({'error': 'Unauthorized'}), 401
    
    try:
        # Top customers by total spent
        top_customers_query = db.session.query(
            Customer.id,
            Customer.name,
            func.count(Order.id).label('order_count'),
            func.sum(Order.total_amount).label('total_spent'),
            func.max(Order.order_date).label('last_order_date')
        ).join(Order).group_by(Customer.id).order_by(func.sum(Order.total_amount).desc()).limit(10)
        
        top_customers = []
        for c in top_customers_query.all():
            top_customers.append({
                'id': c.id,
                'name': c.name,
                'order_count': c.order_count,
                'total_spent': float(c.total_spent) if c.total_spent else 0,
                'last_order_date': c.last_order_date.strftime('%Y-%m-%d') if c.last_order_date else None
            })
        
        # Get inactive customers (haven't ordered in 30 days)
        thirty_days_ago = datetime.datetime.now() - datetime.timedelta(days=30)
        
        inactive_customers_query = db.session.query(
            Customer.id,
            Customer.name,
            func.max(Order.order_date).label('last_order_date'),
            func.sum(Order.total_amount).label('total_spent')
        ).join(Order).group_by(Customer.id).having(
            func.max(Order.order_date) < thirty_days_ago
        ).order_by(func.max(Order.order_date)).limit(10)
        
        inactive_customers = []
        for c in inactive_customers_query.all():
            inactive_customers.append({
                'id': c.id,
                'name': c.name,
                'last_order_date': c.last_order_date.strftime('%Y-%m-%d') if c.last_order_date else None,
                'total_spent': float(c.total_spent) if c.total_spent else 0
            })
        
        # Calculate customer distribution
        # Regular: ordered more than once in last 30 days
        # Casual: ordered once in last 30 days
        # New: first order in last 30 days
        # Inactive: no orders in last 30 days
        
        # Get all customers with their last order date
        customer_stats = db.session.query(
            Customer.id,
            func.count(Order.id).label('order_count'),
            func.max(Order.order_date).label('last_order_date'),
            func.min(Order.order_date).label('first_order_date')
        ).outerjoin(Order).group_by(Customer.id).all()
        
        regular_count = 0
        casual_count = 0
        new_count = 0
        inactive_count = 0
        
        for stat in customer_stats:
            if not stat.last_order_date:
                # No orders
                continue
                
            if stat.last_order_date >= thirty_days_ago:
                # Active in last 30 days
                if stat.first_order_date >= thirty_days_ago:
                    # First order in last 30 days
                    new_count += 1
                elif stat.order_count > 1:
                    # Multiple orders
                    regular_count += 1
                else:
                    # Single order
                    casual_count += 1
            else:
                # No orders in last 30 days
                inactive_count += 1
        
        customer_distribution = [regular_count, casual_count, new_count, inactive_count]
        
        # Calculate purchase frequency
        purchase_frequency = [0, 0, 0, 0, 0]  # [1, 2-3, 4-6, 7-12, >12]
        
        for stat in customer_stats:
            if not stat.order_count:
                continue
                
            if stat.order_count == 1:
                purchase_frequency[0] += 1
            elif 2 <= stat.order_count <= 3:
                purchase_frequency[1] += 1
            elif 4 <= stat.order_count <= 6:
                purchase_frequency[2] += 1
            elif 7 <= stat.order_count <= 12:
                purchase_frequency[3] += 1
            else:
                purchase_frequency[4] += 1
        
        return jsonify({
            'top_customers': top_customers,
            'inactive_customers': inactive_customers,
            'customer_distribution': customer_distribution,
            'purchase_frequency': purchase_frequency
        })
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@main.route('/api/activity_logs')
def api_activity_logs():
    if 'user_id' not in session or session.get('role') != 'admin':
        return jsonify({'error': 'Unauthorized'}), 401
    
    try:
        page = request.args.get('page', 1, type=int)
        per_page = 20
        
        logs = ActivityLog.query.order_by(ActivityLog.timestamp.desc()).paginate(
            page=page, per_page=per_page, error_out=False
        )
        
        logs_data = []
        for log in logs.items:
            logs_data.append({
                'id': log.id,
                'user': log.user.username if log.user else 'System',
                'action': log.action,
                'entity_type': log.entity_type,
                'entity_id': log.entity_id,
                'details': log.details,
                'timestamp': log.timestamp.strftime('%Y-%m-%d %H:%M:%S'),
                'ip_address': log.ip_address
            })
        
        return jsonify({
            'logs': logs_data,
            'has_next': logs.has_next,
            'has_prev': logs.has_prev,
            'page': logs.page,
            'pages': logs.pages,
            'total': logs.total
        })
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@main.route('/realtime_data')
def realtime_data():
    if 'user_id' not in session:
        return jsonify({'error': 'Unauthorized'}), 401
    
    period = request.args.get('period', 'day')
    
    today = datetime.datetime.now().date()
    if period == 'day':
        start_date = today
        end_date = today
    elif period == 'week':
        start_date = today - datetime.timedelta(days=today.weekday())
        end_date = today
    elif period == 'month':
        start_date = today.replace(day=1)
        end_date = today
    elif period == 'year':
        start_date = today.replace(month=1, day=1)
        end_date = today
    else:
        return jsonify({'error': 'Invalid period'}), 400
    
    page = request.args.get('page', 1, type=int)
    per_page = 100
    
    orders_query = Order.query.filter(
        func.date(Order.order_date) >= start_date,
        func.date(Order.order_date) <= end_date
    ).order_by(Order.order_date.desc())
    
    orders_paginated = orders_query.paginate(page=page, per_page=per_page, error_out=False)
    
    total_sales = db.session.query(func.sum(Order.total_amount)).filter(
        func.date(Order.order_date) >= start_date,
        func.date(Order.order_date) <= end_date
    ).scalar() or 0
    total_sales = float(total_sales)
    
    if period == 'day':
        prev_start_date = today - datetime.timedelta(days=1)
        prev_end_date = prev_start_date
    elif period == 'week':
        prev_start_date = start_date - datetime.timedelta(weeks=1)
        prev_end_date = start_date - datetime.timedelta(days=1)
    elif period == 'month':
        prev_month = start_date.month - 1
        prev_year = start_date.year
        if prev_month == 0:
            prev_month = 12
            prev_year -= 1
        prev_start_date = datetime.date(prev_year, prev_month, 1)
        if start_date.month == 1:
            prev_end_date = datetime.date(start_date.year - 1, 12, 31)
        else:
            prev_end_date = start_date - datetime.timedelta(days=1)
    elif period == 'year':
        prev_start_date = datetime.date(start_date.year - 1, 1, 1)
        prev_end_date = datetime.date(start_date.year - 1, 12, 31)
    
    prev_sales = db.session.query(func.sum(Order.total_amount)).filter(
        func.date(Order.order_date) >= prev_start_date,
        func.date(Order.order_date) <= prev_end_date
    ).scalar() or 0
    prev_sales = float(prev_sales)
    
    sales_change = 0
    if prev_sales > 0:
        sales_change = ((total_sales - prev_sales) / prev_sales) * 100
    
    today_sales = db.session.query(func.sum(Order.total_amount)).filter(
        func.date(Order.order_date) == today
    ).scalar() or 0
    today_sales = float(today_sales)
    
    today_order_count = Order.query.filter(
        func.date(Order.order_date) == today
    ).count()
    
    order_items = db.session.query(
        OrderItem.product_id,
        func.sum(OrderItem.quantity).label('quantity'),
        func.sum(OrderItem.price * OrderItem.quantity).label('revenue')
    ).join(Order).filter(
        func.date(Order.order_date) >= start_date,
        func.date(Order.order_date) <= end_date
    ).group_by(OrderItem.product_id).all()
    
    total_profit = 0
    total_cost = 0
    total_units_sold = 0
    
    for item in order_items:
        product = Product.query.get(item.product_id)
        if product:
            quantity = float(item.quantity) if hasattr(item.quantity, 'as_integer_ratio') else item.quantity
            revenue = float(item.revenue) if hasattr(item.revenue, 'as_integer_ratio') else item.revenue
            cost = float(product.cost)
            item_cost = cost * quantity
            item_profit = revenue - item_cost
            total_profit += item_profit
            total_cost += item_cost
            total_units_sold += quantity
    
    prev_order_items = db.session.query(
        OrderItem.product_id,
        func.sum(OrderItem.quantity).label('quantity'),
        func.sum(OrderItem.price * OrderItem.quantity).label('revenue')
    ).join(Order).filter(
        func.date(Order.order_date) >= prev_start_date,
        func.date(Order.order_date) <= prev_end_date
    ).group_by(OrderItem.product_id).all()
    
    prev_total_profit = 0
    prev_total_units = 0
    
    for item in prev_order_items:
        product = Product.query.get(item.product_id)
        if product:
            quantity = float(item.quantity) if hasattr(item.quantity, 'as_integer_ratio') else item.quantity
            revenue = float(item.revenue) if hasattr(item.revenue, 'as_integer_ratio') else item.revenue
            cost = float(product.cost)
            item_cost = cost * quantity
            item_profit = revenue - item_cost
            prev_total_profit += item_profit
            prev_total_units += quantity
    
    profit_change = 0
    if prev_total_profit > 0:
        profit_change = ((total_profit - prev_total_profit) / prev_total_profit) * 100
    
    units_change = 0
    if prev_total_units > 0:
        units_change = ((total_units_sold - prev_total_units) / prev_total_units) * 100
    
    avg_profit_margin = (total_profit / total_sales) * 100 if total_sales > 0 else 0
    cost_percentage = (total_cost / total_sales) * 100 if total_sales > 0 else 0
    
    profit_analysis = {
        'total_sales': total_sales,
        'total_cost': total_cost,
        'total_profit': total_profit
    }
    
    hourly_sales = []
    hourly_labels = []
    
    for hour in range(24):
        hour_start = datetime.datetime.combine(today, datetime.time(hour, 0))
        hour_end = datetime.datetime.combine(today, datetime.time(hour, 59, 59))
        
        hour_sales = db.session.query(func.sum(Order.total_amount)).filter(
            Order.order_date >= hour_start,
            Order.order_date <= hour_end
        ).scalar() or 0
        hour_sales = float(hour_sales)
        
        hourly_sales.append({'labels': f"{hour:02d}:00", 'values': hour_sales})
        hourly_labels.append(f"{hour:02d}:00")
    
    monthly_sales = {'labels': [], 'current_year': [], 'last_year': [], 'target': []}
    current_year = today.year
    for month in range(1, 13):
        month_name = datetime.date(current_year, month, 1).strftime('%b')
        monthly_sales['labels'].append(month_name)
        
        month_start = datetime.date(current_year, month, 1)
        month_end = (datetime.date(current_year, month + 1, 1) - datetime.timedelta(days=1)) if month < 12 else datetime.date(current_year, 12, 31)
        
        month_sales_val = db.session.query(func.sum(Order.total_amount)).filter(
            func.date(Order.order_date) >= month_start,
            func.date(Order.order_date) <= month_end
        ).scalar() or 0
        monthly_sales['current_year'].append(float(month_sales_val))
        
        last_year_month_start = datetime.date(current_year - 1, month, 1)
        last_year_month_end = (datetime.date(current_year - 1, month + 1, 1) - datetime.timedelta(days=1)) if month < 12 else datetime.date(current_year - 1, 12, 31)
        
        last_year_month_sales_val = db.session.query(func.sum(Order.total_amount)).filter(
            func.date(Order.order_date) >= last_year_month_start,
            func.date(Order.order_date) <= last_year_month_end
        ).scalar() or 0
        last_year_month_sales = float(last_year_month_sales_val)
        monthly_sales['last_year'].append(last_year_month_sales)
        monthly_sales['target'].append(last_year_month_sales * 1.1)
    
    top_products_by_units = db.session.query(
        Product.id, Product.name, Product.flavor, func.sum(OrderItem.quantity).label('units_sold')
    ).join(OrderItem, Product.id == OrderItem.product_id).join(Order, OrderItem.order_id == Order.id).filter(
        func.date(Order.order_date) >= start_date, func.date(Order.order_date) <= end_date
    ).group_by(Product.id).order_by(func.sum(OrderItem.quantity).desc()).limit(5).all()
    
    top_products_by_revenue = db.session.query(
        Product.id, Product.name, Product.flavor, func.sum(OrderItem.price * OrderItem.quantity).label('revenue')
    ).join(OrderItem, Product.id == OrderItem.product_id).join(Order, OrderItem.order_id == Order.id).filter(
        func.date(Order.order_date) >= start_date, func.date(Order.order_date) <= end_date
    ).group_by(Product.id).order_by(func.sum(OrderItem.price * OrderItem.quantity).desc()).limit(5).all()
    
    top_products_by_profit = []
    for product in top_products_by_revenue:
        product_obj = Product.query.get(product.id)
        if product_obj:
            quantity_sold = db.session.query(func.sum(OrderItem.quantity)).join(Order, OrderItem.order_id == Order.id).filter(
                OrderItem.product_id == product.id, func.date(Order.order_date) >= start_date, func.date(Order.order_date) <= end_date
            ).scalar() or 0
            
            quantity_sold = float(quantity_sold) if hasattr(quantity_sold, 'as_integer_ratio') else quantity_sold
            revenue = float(product.revenue) if hasattr(product.revenue, 'as_integer_ratio') else product.revenue
            cost = float(product_obj.cost)
            profit = revenue - (cost * quantity_sold)
            
            top_products_by_profit.append({'id': product.id, 'name': product.name, 'flavor': product.flavor, 'profit': profit})
    
    top_products_by_profit = sorted(top_products_by_profit, key=lambda x: x['profit'], reverse=True)[:5]
    
    top_products = {
        'by_units': [{'name': f"{p.name} {p.flavor}", 'value': float(p.units_sold)} for p in top_products_by_units],
        'by_revenue': [{'name': f"{p.name} {p.flavor}", 'value': float(p.revenue)} for p in top_products_by_revenue],
        'by_profit': [{'name': f"{p['name']} {p['flavor']}", 'value': float(p['profit'])} for p in top_products_by_profit]
    }
    
    return jsonify({
        'sales_change': float(sales_change), 'today_sales': float(today_sales), 'today_order_count': today_order_count,
        'total_profit': float(total_profit), 'profit_change': float(profit_change), 'avg_profit_margin': float(avg_profit_margin),
        'total_units_sold': total_units_sold, 'units_change': float(units_change), 'total_cost': float(total_cost),
        'cost_percentage': float(cost_percentage), 'profit_analysis': profit_analysis,
        'hourly_sales': {'labels': hourly_labels, 'values': [float(h['values']) for h in hourly_sales]},
        'monthly_sales': monthly_sales, 'top_products': top_products, 'has_more': orders_paginated.has_next,
        'next_page': page + 1 if orders_paginated.has_next else None
    })

@main.route('/products')
def products():
    if 'user_id' not in session:
        return redirect(url_for('main.login'))
    
    page = request.args.get('page', 1, type=int)
    per_page = 20
    settings = Settings.query.first()
    low_stock_threshold = settings.low_stock_threshold if settings else 10
    
    products_paginated = Product.query.paginate(page=page, per_page=per_page, error_out=False)
    
    low_stock = [p for p in products_paginated.items if p.stock < low_stock_threshold]
    
    return render_template('products.html', products=products_paginated.items, pagination=products_paginated, low_stock=low_stock)

@main.route('/add_product', methods=['GET', 'POST'])
def add_product():
    if 'user_id' not in session:
        return redirect(url_for('main.login'))
    
    if request.method == 'POST':
        wholesale_price = request.form.get('wholesale_price')
        product = Product(
            name=request.form['name'],
            flavor=request.form['flavor'],
            description=request.form['description'],
            price=float(request.form['price']),
            cost=float(request.form['cost']),
            wholesale_price=float(wholesale_price) if wholesale_price else None,
            stock=int(request.form['stock']),
            barcode=request.form.get('barcode', '')
        )
        
        if 'image' in request.files and request.files['image'].filename:
            file = request.files['image']
            image_path = save_file(file, 'products')
            if image_path:
                product.image_path = image_path
            else:
                return render_template('add_product.html')
        
        db.session.add(product)
        db.session.commit()
        
        # Prepare data for Google Sheets and n8n
        user_info = {
            'user_id': session.get('user_id'),
            'username': session.get('username')
        }
        
        product_data = {
            'id': product.id,
            'name': product.name,
            'flavor': product.flavor,
            'description': product.description,
            'price': float(product.price),
            'cost': float(product.cost),
            'wholesale_price': float(product.wholesale_price) if product.wholesale_price else None,
            'stock': product.stock,
            'barcode': product.barcode or '',
            'created_at': product.created_at.isoformat() if product.created_at else datetime.datetime.now().isoformat()
        }
        
        # Sync to Google Sheets with real-time data
        sheets_success = add_product_to_google_sheets_realtime(product_data, 'ADD', user_info)
        
        # Create n8n webhook data
        webhook_data = create_n8n_webhook_data(product_data, 'ADD', user_info)
        
        # Send to n8n webhook
        webhook_success = send_to_n8n_webhook(webhook_data)
        
        log_activity('add_product', 'product', product.id, f'Added product: {product.name} {product.flavor}')
        flash('เพิ่มสินค้าสำเร็จ', 'success')
        return redirect(url_for('main.products'))
    
    return render_template('add_product.html')

@main.route('/edit_product/<int:id>', methods=['GET', 'POST'])
def edit_product(id):
    if 'user_id' not in session:
        return redirect(url_for('main.login'))
    
    product = Product.query.get_or_404(id)
    
    if request.method == 'POST':
        product.name = request.form['name']
        product.flavor = request.form['flavor']
        product.description = request.form['description']
        product.price = float(request.form['price'])
        product.cost = float(request.form['cost'])
        wholesale_price = request.form.get('wholesale_price', '')
        product.wholesale_price = float(wholesale_price) if wholesale_price else None
        product.stock = int(request.form['stock'])
        product.barcode = request.form.get('barcode', '')
        
        if 'image' in request.files and request.files['image'].filename:
            file = request.files['image']
            new_image_path = save_file(file, 'products')
            if new_image_path:
                if product.image_path:
                    delete_file(product.image_path)
                product.image_path = new_image_path
            else:
                return render_template('edit_product.html', product=product)
        
        db.session.commit()
        
        # Prepare data for Google Sheets and n8n
        user_info = {
            'user_id': session.get('user_id'),
            'username': session.get('username')
        }
        
        product_data = {
            'id': product.id,
            'name': product.name,
            'flavor': product.flavor,
            'description': product.description,
            'price': float(product.price),
            'cost': float(product.cost),
            'wholesale_price': float(product.wholesale_price) if product.wholesale_price else None,
            'stock': product.stock,
            'barcode': product.barcode or '',
            'created_at': product.created_at.isoformat() if product.created_at else datetime.datetime.now().isoformat()
        }
        
        # Sync to Google Sheets with real-time data
        sheets_success = add_product_to_google_sheets_realtime(product_data, 'UPDATE', user_info)
        
        # Create n8n webhook data
        webhook_data = create_n8n_webhook_data(product_data, 'UPDATE', user_info)
        
        # Send to n8n webhook
        webhook_success = send_to_n8n_webhook(webhook_data)
        
        log_activity('edit_product', 'product', product.id, f'Edited product: {product.name} {product.flavor}')
        flash('แก้ไขสินค้าสำเร็จ', 'success')
        return redirect(url_for('main.products'))
    
    return render_template('edit_product.html', product=product)

@main.route('/delete_product/<int:id>', methods=['POST'])
def delete_product(id):
    if 'user_id' not in session:
        return redirect(url_for('main.login'))
    if session.get('role') != 'admin':
        flash('คุณไม่มีสิทธิ์ในการลบสินค้า', 'danger')
        return redirect(url_for('main.products'))
    
    product = Product.query.get_or_404(id)
    if product.image_path:
        delete_file(product.image_path)
    
    log_activity('delete_product', 'product', product.id, f'Deleted product: {product.name} {product.flavor}')
    db.session.delete(product)
    db.session.commit()
    
    flash('ลบสินค้าสำเร็จ', 'success')
    return redirect(url_for('main.products'))

@main.route('/customers')
def customers():
    if 'user_id' not in session:
        return redirect(url_for('main.login'))
    page = request.args.get('page', 1, type=int)
    per_page = 20
    customers_paginated = Customer.query.paginate(page=page, per_page=per_page, error_out=False)
    return render_template('customers.html', customers=customers_paginated.items, pagination=customers_paginated)

@main.route('/add_customer', methods=['GET', 'POST'])
def add_customer():
    if 'user_id' not in session:
        return redirect(url_for('main.login'))
    if request.method == 'POST':
        customer = Customer(
            name=request.form['name'],
            phone=request.form.get('phone', ''),
            address=request.form.get('address', ''),
            line_id=request.form.get('line_id', '')
        )
        db.session.add(customer)
        db.session.commit()
        log_activity('add_customer', 'customer', customer.id, f'Added customer: {customer.name}')
        flash('เพิ่มลูกค้าสำเร็จ', 'success')
        return redirect(url_for('main.customers'))
    return render_template('add_customer.html')

@main.route('/edit_customer/<int:id>', methods=['GET', 'POST'])
def edit_customer(id):
    if 'user_id' not in session:
        return redirect(url_for('main.login'))
    customer = Customer.query.get_or_404(id)
    if request.method == 'POST':
        customer.name = request.form['name']
        customer.phone = request.form.get('phone', '')
        customer.address = request.form.get('address', '')
        customer.line_id = request.form.get('line_id', '')
        db.session.commit()
        log_activity('edit_customer', 'customer', customer.id, f'Edited customer: {customer.name}')
        flash('แก้ไขข้อมูลลูกค้าสำเร็จ', 'success')
        return redirect(url_for('main.customers'))
    return render_template('edit_customer.html', customer=customer)

@main.route('/view_customer/<int:id>')
def view_customer(id):
    if 'user_id' not in session:
        return redirect(url_for('main.login'))
    customer = Customer.query.get_or_404(id)
    orders = Order.query.filter_by(customer_id=id).order_by(Order.order_date.desc()).all()
    return render_template('view_customer.html', customer=customer, orders=orders)

@main.route('/orders')
def orders():
    if 'user_id' not in session:
        return redirect(url_for('main.login'))
    page = request.args.get('page', 1, type=int)
    per_page = 20
    orders_paginated = Order.query.options(db.joinedload(Order.customer)).order_by(Order.order_date.desc()).paginate(page=page, per_page=per_page, error_out=False)
    return render_template('orders.html', orders=orders_paginated.items, pagination=orders_paginated)

@main.route('/add_order', methods=['GET', 'POST'])
def add_order():
    if 'user_id' not in session:
        return redirect(url_for('main.login'))
    
    if request.method == 'POST':
        try:
            customer_id = request.form.get('customer_id')
            if not customer_id:
                flash('กรุณาเลือกลูกค้า', 'danger')
                customers = Customer.query.all()
                products = Product.query.filter(Product.stock > 0).all()
                return render_template('add_order.html', customers=customers, products=products)

            current_time = datetime.datetime.now(datetime.timezone.utc)
            order = Order(
                customer_id=customer_id,
                shipping_address=request.form.get('shipping_address', ''),
                notes=request.form.get('notes', ''),
                order_date=current_time
            )
            db.session.add(order)
            db.session.flush()
            
            total_amount = 0
            order_items_created = []
            
            product_ids = request.form.getlist('product_id[]')
            quantities = request.form.getlist('quantity[]')
            
            if not product_ids or not quantities:
                flash('กรุณาเลือกสินค้าอย่างน้อย 1 รายการ', 'danger')
                db.session.rollback()
                customers = Customer.query.all()
                products = Product.query.filter(Product.stock > 0).all()
                return render_template('add_order.html', customers=customers, products=products)
            
            for i in range(len(product_ids)):
                if not product_ids[i] or not quantities[i]:
                    continue
                    
                try:
                    product_id = int(product_ids[i])
                    quantity = int(quantities[i])
                    if quantity <= 0:
                        continue
                except (ValueError, TypeError):
                    continue
                
                with stock_update_lock:
                    product = Product.query.get(product_id)
                    if not product:
                        flash(f'ไม่พบสินค้าที่เลือก', 'danger')
                        db.session.rollback()
                        customers = Customer.query.all()
                        products = Product.query.filter(Product.stock > 0).all()
                        return render_template('add_order.html', customers=customers, products=products)
                    
                    if product.stock < quantity:
                        flash(f'สินค้า {product.name} {product.flavor} มีไม่เพียงพอ (คงเหลือ: {product.stock})', 'danger')
                        db.session.rollback()
                        customers = Customer.query.all()
                        products = Product.query.filter(Product.stock > 0).all()
                        return render_template('add_order.html', customers=customers, products=products)
                    
                    # Create order item
                    order_item = OrderItem(
                        order_id=order.id, 
                        product_id=product_id, 
                        quantity=quantity, 
                        price=product.price
                    )
                    db.session.add(order_item)
                    order_items_created.append(order_item)
                    
                    # Create inventory entry
                    inventory_entry = InventoryEntry(
                        product_id=product_id, 
                        quantity=-quantity, 
                        date=current_time, 
                        notes=f'Order #{order.id}', 
                        user_id=session['user_id']
                    )
                    db.session.add(inventory_entry)
                    
                    # Update product stock
                    product.stock -= quantity
                    total_amount += product.price * quantity
                    
                    # Check for low stock notification
                    settings = Settings.query.first()
                    low_stock_threshold = settings.low_stock_threshold if settings else 10
                    if product.stock < low_stock_threshold:
                        create_notification(
                            f'สินค้า {product.name} {product.flavor} เหลือ {product.stock} ชิ้น', 
                            'low_stock', 
                            product.id
                        )
            
            if total_amount == 0:
                flash('กรุณาเลือกสินค้าและระบุจำนวนที่ถูกต้อง', 'danger')
                db.session.rollback()
                customers = Customer.query.all()
                products = Product.query.filter(Product.stock > 0).all()
                return render_template('add_order.html', customers=customers, products=products)
            
            order.total_amount = total_amount
            order.qr_code_path = generate_qr_code(order.id)
            
            # Handle payment slip upload
            if 'payment_slip' in request.files and request.files['payment_slip'].filename:
                file = request.files['payment_slip']
                payment_slip_path = save_file(file, 'slips')
                if payment_slip_path:
                    order.payment_slip = payment_slip_path
                    order.payment_date = current_time
                    order.payment_status = 'paid'
            
            db.session.commit()
            
            # Prepare user info for Google Sheets and webhook
            user_info = {
                'user_id': session.get('user_id'),
                'username': session.get('username')
            }
            
            # Update stock in Google Sheets
            try:
                from app.utils import update_stock_in_google_sheets_for_order, create_order_webhook_data, send_to_n8n_webhook
                sheets_success = update_stock_in_google_sheets_for_order(order_items_created, 'SALE', user_info)
                
                # Create and send webhook data for order
                webhook_data = create_order_webhook_data(order, order_items_created, 'CREATE', user_info)
                webhook_success = send_to_n8n_webhook(webhook_data)
                
                if sheets_success:
                    flash('เพิ่มคำสั่งซื้อสำเร็จ และอัพเดทข้อมูลใน Google Sheets แล้ว', 'success')
                else:
                    flash('เพิ่มคำสั่งซื้อสำเร็จ แต่ไม่สามารถอัพเดท Google Sheets ได้', 'warning')
            except Exception as e:
                current_app.logger.error(f"Error in Google Sheets/webhook integration: {str(e)}")
                flash('เพิ่มคำสั่งซื้อสำเร็จ แต่เกิดข้อผิดพลาดในการอัพเดทข้อมูลภายนอก', 'warning')
            
            create_notification(f'มีคำสั่งซื้อใหม่ #{order.id}', 'new_order', order.id)
            log_activity('add_order', 'order', order.id, f'Added order for customer: {order.customer.name}')
            
            return redirect(url_for('main.view_order', id=order.id))
            
        except Exception as e:
            db.session.rollback()
            current_app.logger.error(f"Error in add_order: {str(e)}")
            flash(f'เกิดข้อผิดพลาด: {str(e)}', 'danger')
    
    customers = Customer.query.all()
    products = Product.query.filter(Product.stock > 0).all()
    return render_template('add_order.html', customers=customers, products=products)

@main.route('/view_order/<int:id>')
def view_order(id):
    if 'user_id' not in session:
        return redirect(url_for('main.login'))
    order = Order.query.get_or_404(id)
    return render_template('view_order.html', order=order)

@main.route('/delete_order/<int:id>', methods=['POST'])
def delete_order(id):
    order = Order.query.get_or_404(id)
    try:
        current_time = datetime.datetime.now(datetime.timezone.utc)
        for item in order.order_items:
            product = Product.query.get(item.product_id)
            if product:
                product.stock += item.quantity
                db.session.add(InventoryEntry(product_id=item.product_id, quantity=item.quantity, date=current_time, notes=f'Order #{order.id} deleted', user_id=session.get('user_id')))
        
        # Prepare user info for Google Sheets and webhook
        user_info = {
            'user_id': session.get('user_id'),
            'username': session.get('username')
        }
        
        # Update stock in Google Sheets (return stock)
        from app.utils import update_stock_in_google_sheets_for_order, create_order_webhook_data, send_to_n8n_webhook
        sheets_success = update_stock_in_google_sheets_for_order(order.order_items, 'RETURN', user_info)
        
        # Create and send webhook data for order deletion
        webhook_data = create_order_webhook_data(order, order.order_items, 'DELETE', user_info)
        webhook_success = send_to_n8n_webhook(webhook_data)
        
        db.session.delete(order)
        db.session.commit()
        
        if sheets_success:
            flash('คำสั่งซื้อถูกลบเรียบร้อยแล้ว และอัพเดทข้อมูลใน Google Sheets แล้ว', 'success')
        else:
            flash('คำสั่งซื้อถูกลบเรียบร้อยแล้ว แต่ไม่สามารถอัพเดท Google Sheets ได้', 'warning')
    except Exception as e:
        db.session.rollback()
        flash(f'เกิดข้อผิดพลาด: {str(e)}', 'danger')
    return redirect(url_for('main.orders'))

@main.route('/update_payment/<int:id>', methods=['POST'])
def update_payment(id):
    if 'user_id' not in session:
        return redirect(url_for('main.login'))
    order = Order.query.get_or_404(id)
    if 'payment_slip' in request.files and request.files['payment_slip'].filename:
        file = request.files['payment_slip']
        if order.payment_slip:
            delete_file(order.payment_slip)
        payment_slip = save_file(file, 'slips')
        if payment_slip:
            order.payment_slip = payment_slip
            order.payment_date = datetime.datetime.now()
            order.payment_status = 'paid'
            db.session.commit()
            create_notification(f'มีการชำระเงินสำหรับคำสั่งซื้อ #{order.id}', 'payment', order.id)
            log_activity('update_payment', 'order', order.id)
            flash('อัปเดตการชำระเงินสำเร็จ', 'success')
    return redirect(url_for('main.view_order', id=id))

@main.route('/generate_payment_qr/<int:order_id>')
def generate_payment_qr(order_id):
    if 'user_id' not in session and not request.args.get('public'):
        return redirect(url_for('main.login'))
    order = Order.query.get_or_404(order_id)
    if not order.qr_code_path:
        order.qr_code_path = generate_qr_code(order.id)
        db.session.commit()
    qr_code_url = url_for('static', filename=order.qr_code_path, _external=True)
    return render_template('payment_qr.html', order=order, qr_code=qr_code_url)

@main.route('/receipt/<int:order_id>')
def receipt(order_id):
    if 'user_id' not in session and not request.args.get('public'):
        return redirect(url_for('main.login'))
    order = Order.query.get_or_404(order_id)
    return render_template('receipt.html', order=order)

@main.route('/inventory')
def inventory():
    if 'user_id' not in session:
        return redirect(url_for('main.login'))
    page = request.args.get('page', 1, type=int)
    per_page = 20
    products_paginated = Product.query.paginate(page=page, per_page=per_page, error_out=False)
    return render_template('inventory.html', products=products_paginated.items, pagination=products_paginated)

@main.route('/add_inventory', methods=['GET', 'POST'])
def add_inventory():
    if 'user_id' not in session:
        return redirect(url_for('main.login'))
    if request.method == 'POST':
        product_id = request.form.get('product_id')
        if not product_id:
            product = Product.query.filter_by(name=request.form.get('name'), flavor=request.form.get('flavor')).first()
            if product:
                product_id = product.id
            else:
                flash('กรุณาเลือกสินค้า', 'danger')
                return redirect(url_for('main.add_inventory'))
        
        try:
            quantity = int(request.form['quantity'])
            if quantity <= 0:
                flash('จำนวนต้องมากกว่า 0', 'danger')
                return redirect(url_for('main.add_inventory'))
        except ValueError:
            flash('กรุณากรอกจำนวนเป็นตัวเลข', 'danger')
            return redirect(url_for('main.add_inventory'))
        
        with stock_update_lock:
            product = Product.query.get_or_404(product_id)
            db.session.add(InventoryEntry(product_id=product_id, quantity=quantity, date=datetime.datetime.now(datetime.timezone.utc), notes=request.form.get('notes', ''), user_id=session.get('user_id')))
            product.stock += quantity
            db.session.commit()
        
        log_activity('add_inventory', 'product', product_id, f'Added {quantity} to {product.name}')
        flash('เพิ่มสต็อกสำเร็จ', 'success')
        return redirect(url_for('main.inventory'))
    
    product = Product.query.get(request.args.get('product_id')) if request.args.get('product_id') else None
    return render_template('add_inventory.html', selected_product=product)

@main.route('/inventory_history')
def inventory_history():
    if 'user_id' not in session:
        return redirect(url_for('main.login'))
    page = request.args.get('page', 1, type=int)
    per_page = 20
    history_paginated = InventoryEntry.query.join(Product).join(User).order_by(InventoryEntry.date.desc()).paginate(page=page, per_page=per_page, error_out=False)
    return render_template('inventory_history.html', history=history_paginated.items, pagination=history_paginated)

@main.route('/realtime_stock')
def realtime_stock():
    if 'user_id' not in session:
        return redirect(url_for('main.login'))
    return render_template('realtime_stock.html')

@main.route('/api/realtime_stock_data')
def api_realtime_stock_data():
    if 'user_id' not in session:
        return jsonify({'error': 'Unauthorized'}), 401
    products = Product.query.all()
    settings = Settings.query.first()
    low_stock_threshold = settings.low_stock_threshold if settings else 10
    products_data = [{'id': p.id, 'name': p.name, 'flavor': p.flavor, 'price': float(p.price), 'cost': float(p.cost), 'wholesale_price': float(p.wholesale_price) if p.wholesale_price else None, 'stock': p.stock} for p in products]
    return jsonify({'products': products_data, 'low_stock_threshold': low_stock_threshold})

@main.route('/api/public_stock_data')
def api_public_stock_data():
    """Public API endpoint for stock data - no authentication required"""
    try:
        products = Product.query.filter(Product.stock > 0).all()
        products_data = []
        for product in products:
            products_data.append({
                'id': product.id,
                'name': product.name,
                'flavor': product.flavor,
                'price': float(product.price),
                'stock': product.stock,
                'image_path': product.image_path
            })
        return jsonify({'products': products_data})
    except Exception as e:
        return jsonify({'error': 'Failed to fetch stock data'}), 500

@main.route('/add_inventory_barcode', methods=['GET', 'POST'])
def add_inventory_barcode():
    if 'user_id' not in session:
        return redirect(url_for('main.login'))
    return render_template('add_inventory_barcode.html')

@main.route('/api/scan_barcode', methods=['POST'])
def scan_barcode():
    if 'user_id' not in session:
        return jsonify({'error': 'Unauthorized'}), 401
    product = Product.query.filter_by(barcode=request.json.get('barcode')).first()
    if not product:
        return jsonify({'found': False})
    return jsonify({'found': True, 'product': {'id': product.id, 'name': product.name, 'flavor': product.flavor, 'price': product.price, 'stock': product.stock}})

@main.route('/api/search_products')
def api_search_products_internal():
    if 'user_id' not in session:
        return jsonify({'error': 'Unauthorized'}), 401
    query = request.args.get('query', '').strip()
    products = Product.query.filter(or_(Product.name.ilike(f'%{query}%'), Product.flavor.ilike(f'%{query}%'))).limit(10).all() if query else []
    return jsonify({'products': [{'id': p.id, 'name': p.name, 'flavor': p.flavor, 'stock': p.stock} for p in products]})

@main.route('/api/search_inventory_products')
def api_search_inventory_products_internal():
    if 'user_id' not in session:
        return jsonify({'error': 'Unauthorized'}), 401
    
    query = request.args.get('query', '').strip()
    if not query:
        return jsonify({'products': []})
    
    # Search products by name, flavor, or barcode
    products = Product.query.filter(
        or_(
            Product.name.ilike(f'%{query}%'),
            Product.flavor.ilike(f'%{query}%'),
            Product.barcode.ilike(f'%{query}%') if query else False
        )
    ).limit(20).all()
    
    products_data = []
    for product in products:
        products_data.append({
            'id': product.id,
            'name': product.name,
            'flavor': product.flavor,
            'price': float(product.price),
            'cost': float(product.cost),
            'stock': product.stock,
            'barcode': product.barcode or '',
            'profit': float(product.price - product.cost),
            'stock_value': float(product.price * product.stock)
        })
    
    return jsonify({'products': products_data})

@main.route('/reports')
def reports():
    if 'user_id' not in session:
        return redirect(url_for('main.login'))
    return render_template('reports.html')

@main.route('/reports/customer_analysis')
def customer_analysis():
    if 'user_id' not in session:
        return redirect(url_for('main.login'))
    return render_template('reports/customer_analysis.html')

@main.route('/reports/sales_trend')
def sales_trend():
    if 'user_id' not in session:
        return redirect(url_for('main.login'))

    return render_template('reports/sales_trend.html')

@main.route('/admin/activity_logs')
def activity_logs():
    if 'user_id' not in session or session.get('role') != 'admin':
        flash('You do not have permission to view this page.', 'danger')
        return redirect(url_for('main.dashboard'))

    page = request.args.get('page', 1, type=int)
    per_page = 20
    logs = ActivityLog.query.order_by(ActivityLog.timestamp.desc()).paginate(page=page, per_page=per_page, error_out=False)
    
    return render_template('admin/activity_logs.html', logs=logs)

@main.route('/reports/sales', methods=['GET', 'POST'])
def sales_report():
    if 'user_id' not in session:
        return redirect(url_for('main.login'))
    if request.method == 'POST':
        start_date = datetime.datetime.strptime(request.form['start_date'], '%Y-%m-%d').date()
        end_date = datetime.datetime.strptime(request.form['end_date'], '%Y-%m-%d').date()
        page = request.args.get('page', 1, type=int)
        per_page = 50
        
        orders_query = Order.query.filter(func.date(Order.order_date) >= start_date, func.date(Order.order_date) <= end_date).order_by(Order.order_date.desc())
        orders_paginated = orders_query.paginate(page=page, per_page=per_page, error_out=False)
        
        total_sales = db.session.query(func.sum(Order.total_amount)).filter(func.date(Order.order_date) >= start_date, func.date(Order.order_date) <= end_date).scalar() or 0
        
        order_items = db.session.query(OrderItem.product_id, func.sum(OrderItem.quantity).label('quantity'), func.sum(OrderItem.price * OrderItem.quantity).label('amount')).join(Order).filter(func.date(Order.order_date) >= start_date, func.date(Order.order_date) <= end_date).group_by(OrderItem.product_id).all()
        products = {p.id: p for p in Product.query.filter(Product.id.in_([item.product_id for item in order_items])).all()}
        
        product_sales = {}
        for item in order_items:
            product = products.get(item.product_id)
            if product:
                cost = float(product.cost) * float(item.quantity)
                product_sales[f"{product.name} ({product.flavor})"] = {'quantity': item.quantity, 'amount': item.amount, 'cost': cost, 'profit': float(item.amount) - cost}
        
        total_profit = sum(item['profit'] for item in product_sales.values())
        
        return render_template('sales_report.html', orders=orders_paginated.items, total_sales=total_sales, total_orders=orders_query.count(), total_profit=total_profit, product_sales=product_sales, start_date=start_date, end_date=end_date, pagination=orders_paginated)
    
    return render_template('sales_report.html')

@main.route('/reports/product', methods=['GET', 'POST'])
def product_report():
    if 'user_id' not in session:
        return redirect(url_for('main.login'))
    products = Product.query.all()
    if request.method == 'POST':
        product_id = request.form['product_id']
        product = Product.query.get_or_404(product_id)
        page = request.args.get('page', 1, type=int)
        per_page = 20
        orders_query = Order.query.join(OrderItem).filter(OrderItem.product_id == product_id).distinct()
        orders_paginated = orders_query.paginate(page=page, per_page=per_page, error_out=False)
        
        order_items = OrderItem.query.filter_by(product_id=product_id).all()
        total_sold = sum(item.quantity for item in order_items)
        total_revenue = sum(float(item.price) * item.quantity for item in order_items)
        total_profit = total_revenue - (float(product.cost) * total_sold)
        
        return render_template('product_report.html', products=products, product=product, orders=orders_paginated.items, total_sold=total_sold, total_revenue=total_revenue, total_profit=total_profit, pagination=orders_paginated)
    
    return render_template('product_report.html', products=products)

@main.route('/reports/profit')
def profit_analysis():
    if 'user_id' not in session:
        return redirect(url_for('main.login'))
    
    products = Product.query.all()
    total_value = sum(float(p.price) * p.stock for p in products)
    total_cost = sum(float(p.cost) * p.stock for p in products)
    
    # Sample product for demonstration (first product or create dummy data)
    sample_product = None
    if products:
        first_product = products[0]
        sample_product = {
            'name': first_product.name,
            'flavor': first_product.flavor,
            'price': float(first_product.price),
            'cost': float(first_product.cost),
            'profit': float(first_product.price) - float(first_product.cost),
            'profit_percentage': ((float(first_product.price) - float(first_product.cost)) / float(first_product.price) * 100) if first_product.price > 0 else 0
        }
    
    # Top products by margin - create list of dicts with profit_margin
    top_products_by_margin = []
    for product in products:
        profit_margin = ((float(product.price) - float(product.cost)) / float(product.price) * 100) if product.price > 0 else 0
        top_products_by_margin.append({
            'name': product.name,
            'flavor': product.flavor,
            'price': float(product.price),
            'cost': float(product.cost),
            'profit_margin': profit_margin
        })
    top_products_by_margin = sorted(top_products_by_margin, key=lambda x: x['profit_margin'], reverse=True)[:5]
    
    # Top products by total profit
    top_products_by_total_profit = []
    for product in products:
        total_profit = (float(product.price) - float(product.cost)) * product.stock
        top_products_by_total_profit.append({
            'name': product.name,
            'flavor': product.flavor,
            'price': float(product.price),
            'cost': float(product.cost),
            'stock': product.stock,
            'total_profit': total_profit
        })
    top_products_by_total_profit = sorted(top_products_by_total_profit, key=lambda x: x['total_profit'], reverse=True)[:5]
    
    # Generate sample historical data for charts
    months = []
    revenue_data = []
    cost_data = []
    profit_data = []
    
    today = datetime.datetime.now().date()
    for i in range(6):
        month_date = today.replace(day=1) - datetime.timedelta(days=30*i)
        months.insert(0, month_date.strftime('%b %Y'))
        
        # Get actual sales data for this month
        month_start = month_date.replace(day=1)
        if month_date.month == 12:
            month_end = datetime.date(month_date.year + 1, 1, 1) - datetime.timedelta(days=1)
        else:
            month_end = month_date.replace(month=month_date.month + 1, day=1) - datetime.timedelta(days=1)
        
        month_revenue = db.session.query(func.sum(Order.total_amount)).filter(
            func.date(Order.order_date) >= month_start,
            func.date(Order.order_date) <= month_end
        ).scalar() or 0
        
        # Calculate cost based on sold items
        month_items = db.session.query(
            OrderItem.product_id,
            func.sum(OrderItem.quantity).label('quantity')
        ).join(Order).filter(
            func.date(Order.order_date) >= month_start,
            func.date(Order.order_date) <= month_end
        ).group_by(OrderItem.product_id).all()
        
        month_cost = 0
        for item in month_items:
            product = Product.query.get(item.product_id)
            if product:
                month_cost += float(product.cost) * float(item.quantity)
        
        revenue_data.insert(0, float(month_revenue))
        cost_data.insert(0, float(month_cost))
        profit_data.insert(0, float(month_revenue) - float(month_cost))
    
    return render_template('profit_analysis.html',
                          products=products,
                          total_value=total_value,
                          total_cost=total_cost,
                          total_profit_potential=total_value - total_cost,
                          sample_product=sample_product,
                          top_products_by_margin=top_products_by_margin,
                          top_products_by_total_profit=top_products_by_total_profit,
                          months=months,
                          revenue_data=revenue_data,
                          cost_data=cost_data,
                          profit_data=profit_data)

@main.route('/profile', methods=['GET', 'POST'])
def profile():
    if 'user_id' not in session:
        return redirect(url_for('main.login'))
    user = User.query.get(session['user_id'])
    if request.method == 'POST':
        if not check_password_hash(user.password, request.form['current_password']):
            flash('รหัสผ่านปัจจุบันไม่ถูกต้อง', 'danger')
        elif request.form['new_password'] != request.form['confirm_password']:
            flash('รหัสผ่านใหม่ไม่ตรงกัน', 'danger')
        else:
            user.password = generate_password_hash(request.form['new_password'])
            db.session.commit()
            log_activity('change_password')
            flash('เปลี่ยนรหัสผ่านสำเร็จ', 'success')
        return redirect(url_for('main.profile'))
    return render_template('profile.html', user=user)

@main.route('/manage_users')
def manage_users():
    if session.get('role') != 'admin':
        return redirect(url_for('main.dashboard'))
    users = User.query.all()
    return render_template('admin/users.html', users=users)

@main.route('/add_user', methods=['GET', 'POST'])
def add_user():
    if session.get('role') != 'admin':
        return redirect(url_for('main.dashboard'))
    if request.method == 'POST':
        if User.query.filter_by(username=request.form['username']).first():
            flash('ชื่อผู้ใช้นี้มีอยู่แล้ว', 'danger')
            return redirect(url_for('main.add_user'))
        user = User(username=request.form['username'], password=generate_password_hash(request.form['password']), name=request.form.get('name', ''), role=request.form['role'])
        db.session.add(user)
        db.session.commit()
        log_activity('add_user', 'user', user.id, f'Added user: {user.username}')
        flash('เพิ่มผู้ใช้สำเร็จ', 'success')
        return redirect(url_for('main.manage_users'))
    return render_template('admin/add_user.html')

@main.route('/edit_user/<int:id>', methods=['GET', 'POST'])
def edit_user(id):
    if session.get('role') != 'admin':
        return redirect(url_for('main.dashboard'))
    user = User.query.get_or_404(id)
    if request.method == 'POST':
        if request.form.get('password'):
            user.password = generate_password_hash(request.form['password'])
        user.name = request.form.get('name', '')
        user.role = request.form['role']
        db.session.commit()
        log_activity('edit_user', 'user', user.id, f'Edited user: {user.username}')
        flash('แก้ไขผู้ใช้สำเร็จ', 'success')
        return redirect(url_for('main.manage_users'))
    return render_template('admin/edit_user.html', user=user)

@main.route('/notifications')
def notifications():
    if 'user_id' not in session:
        return redirect(url_for('main.login'))
    
    user_notifications = Notification.query.filter_by(user_id=session['user_id']).order_by(Notification.created_at.desc()).all()
    
    try:
        Notification.query.filter_by(user_id=session['user_id'], is_read=False).update({'is_read': True})
        db.session.commit()
    except Exception as e:
        db.session.rollback()
        flash('Could not mark notifications as read.', 'danger')

    return render_template('notifications.html', notifications=user_notifications)

@main.route('/delete_user/<int:id>', methods=['POST'])
def delete_user(id):
    if session.get('role') != 'admin':
        return redirect(url_for('main.dashboard'))
    if id == session['user_id']:
        flash('คุณไม่สามารถลบบัญชีของตัวเองได้', 'danger')
        return redirect(url_for('main.manage_users'))
    user = User.query.get_or_404(id)
    log_activity('delete_user', 'user', user.id, f'Deleted user: {user.username}')
    db.session.delete(user)
    db.session.commit()
    flash('ลบผู้ใช้สำเร็จ', 'success')
    return redirect(url_for('main.manage_users'))

@main.route('/storefront_sale', methods=['GET', 'POST'])
def storefront_sale():
    if request.method == 'POST':
        customer = Customer.query.filter_by(name=request.form.get('customer_name', 'Walk-in Customer')).first()
        if not customer:
            customer = Customer(name=request.form.get('customer_name', 'Walk-in Customer'), phone=request.form.get('customer_phone', ''))
            db.session.add(customer)
            db.session.flush()
        
        current_time = datetime.datetime.now(datetime.timezone.utc)
        order = Order(customer_id=customer.id, order_date=current_time, status='completed', payment_status='paid', notes=f"Payment Method: {request.form.get('payment_method', 'cash')}")
        db.session.add(order)
        db.session.flush()
        
        total_amount = 0
        for i in range(len(request.form.getlist('product_id[]'))):
            product_id = int(request.form.getlist('product_id[]')[i])
            quantity = int(request.form.getlist('quantity[]')[i])
            if quantity <= 0: continue
            
            product = Product.query.get(product_id)
            if not product or product.stock < quantity:
                flash(f'สินค้า {product.name if product else ""} มีไม่เพียงพอ', 'danger')
                continue
            
            db.session.add(OrderItem(order_id=order.id, product_id=product_id, quantity=quantity, price=product.price))
            product.stock -= quantity
            # Use the current logged-in user or a system user ID (1 for admin)
            system_user_id = session.get('user_id', 1)  # Default to user ID 1 if no session
            db.session.add(InventoryEntry(product_id=product_id, quantity=-quantity, date=current_time, notes=f'Storefront Sale - Order #{order.id}', user_id=system_user_id))
            total_amount += product.price * quantity

        order.total_amount = total_amount
        try:
            db.session.commit()
            
            # Prepare user info for Google Sheets and webhook
            user_info = {
                'user_id': session.get('user_id', 1),  # Default to admin if no session
                'username': session.get('username', 'System')
            }
            
            # Update stock in Google Sheets
            from app.utils import update_stock_in_google_sheets_for_order, create_order_webhook_data, send_to_n8n_webhook
            sheets_success = update_stock_in_google_sheets_for_order(order.order_items, 'SALE', user_info)
            
            # Create and send webhook data for storefront sale
            webhook_data = create_order_webhook_data(order, order.order_items, 'STOREFRONT_SALE', user_info)
            webhook_success = send_to_n8n_webhook(webhook_data)
            
            if sheets_success:
                flash('บันทึกการขายเรียบร้อยแล้ว และอัพเดทข้อมูลใน Google Sheets แล้ว', 'success')
            else:
                flash('บันทึกการขายเรียบร้อยแล้ว แต่ไม่สามารถอัพเดท Google Sheets ได้', 'warning')
            return redirect(url_for('main.view_order', id=order.id))
        except Exception as e:
            db.session.rollback()
            flash(f'เกิดข้อผิดพลาด: {str(e)}', 'danger')

    products = Product.query.filter(Product.stock > 0).all()
    return render_template('storefront_sale.html', products=products)

@main.route('/backup_data')
def backup_data():
    if session.get('role') != 'admin':
        return redirect(url_for('main.dashboard'))
    
    # Get the absolute path to the project root
    project_root = os.path.abspath(os.path.dirname(current_app.root_path))
    backup_dir = os.path.join(project_root, current_app.config['UPLOAD_FOLDER'], 'backups')
    os.makedirs(backup_dir, exist_ok=True)
    timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    
    # Excel backup
    excel_filename = f'data_export_{timestamp}.xlsx'
    excel_path = os.path.join(backup_dir, excel_filename)
    
    # Database backup
    db_filename = f'database_backup_{timestamp}.sql'
    db_backup_path = os.path.join(backup_dir, db_filename)
    
    excel_created = False
    db_backup_created = False
    
    try:
        # Create Excel backup
        with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
            # Export all tables to Excel
            tables_to_export = [
                'user', 'customer', 'product', 'order', 'order_item', 
                'inventory_entry', 'activity_log', 'notification', 
                'settings', 'shop_settings', 'banner', 'featured_product', 'store'
            ]
            
            for table_name in tables_to_export:
                try:
                    # Use raw SQL to ensure compatibility
                    query = f"SELECT * FROM {table_name}"
                    df = pd.read_sql(query, db.engine)
                    df.to_excel(writer, sheet_name=table_name, index=False)
                except Exception as e:
                    current_app.logger.warning(f"Could not export table {table_name}: {str(e)}")
                    continue
        
        excel_created = True
        current_app.logger.info(f"Excel backup created: {excel_filename}")
        
    except Exception as e:
        current_app.logger.error(f"Error creating Excel backup: {str(e)}")
    
    # Create database backup
    try:
        db_uri = current_app.config.get('SQLALCHEMY_DATABASE_URI', '')
        
        if db_uri.startswith('sqlite:///'):
            # SQLite database - copy the file
            db_source_path = db_uri.replace('sqlite:///', '')
            if not os.path.isabs(db_source_path):
                db_source_path = os.path.join(project_root, db_source_path)
            
            if os.path.exists(db_source_path):
                # For SQLite, copy the .db file instead of .sql
                db_filename = f'database_backup_{timestamp}.db'
                db_backup_path = os.path.join(backup_dir, db_filename)
                shutil.copy2(db_source_path, db_backup_path)
                db_backup_created = True
                current_app.logger.info(f"SQLite backup created: {db_filename}")
            else:
                current_app.logger.error(f"SQLite database file not found: {db_source_path}")
                
        elif 'mysql' in db_uri:
            # MySQL database - create SQL dump
            try:
                # Create SQL dump using modern SQLAlchemy syntax
                tables_to_backup = [
                    'user', 'customer', 'product', 'order', 'order_item', 
                    'inventory_entry', 'activity_log', 'notification', 
                    'settings', 'shop_settings', 'banner', 'featured_product', 'store'
                ]
                
                with open(db_backup_path, 'w', encoding='utf-8') as f:
                    f.write(f"-- Database backup created on {datetime.datetime.now()}\n")
                    f.write(f"-- Generated by Marbo9K System\n\n")
                    
                    for table_name in tables_to_backup:
                        try:
                            # Get table structure
                            f.write(f"-- Table: {table_name}\n")
                            
                            # Use modern SQLAlchemy syntax
                            with db.engine.connect() as connection:
                                result = connection.execute(text(f"SELECT * FROM {table_name}"))
                                rows = result.fetchall()
                                
                                if rows:
                                    # Get column names
                                    columns = result.keys()
                                    
                                    # Create INSERT statements
                                    for row in rows:
                                        values = []
                                        for value in row:
                                            if value is None:
                                                values.append('NULL')
                                            elif isinstance(value, str):
                                                # Escape single quotes
                                                escaped_value = value.replace("'", "''")
                                                values.append(f"'{escaped_value}'")
                                            elif isinstance(value, datetime.datetime):
                                                values.append(f"'{value.strftime('%Y-%m-%d %H:%M:%S')}'")
                                            elif isinstance(value, datetime.date):
                                                values.append(f"'{value.strftime('%Y-%m-%d')}'")
                                            else:
                                                values.append(str(value))
                                        
                                        columns_str = ', '.join(columns)
                                        values_str = ', '.join(values)
                                        f.write(f"INSERT INTO {table_name} ({columns_str}) VALUES ({values_str});\n")
                            
                            f.write(f"\n")
                            
                        except Exception as e:
                            current_app.logger.warning(f"Could not backup table {table_name}: {str(e)}")
                            f.write(f"-- Error backing up table {table_name}: {str(e)}\n\n")
                            continue
                
                db_backup_created = True
                current_app.logger.info(f"MySQL backup created: {db_filename}")
                
            except Exception as e:
                current_app.logger.error(f"Error creating MySQL backup: {str(e)}")
        
        else:
            current_app.logger.warning(f"Unsupported database type for backup: {db_uri}")
            
    except Exception as e:
        current_app.logger.error(f"Error in database backup process: {str(e)}")
    
    # Log the backup activity
    backup_status = []
    if excel_created:
        backup_status.append(f"Excel={excel_filename}")
    if db_backup_created:
        backup_status.append(f"DB={db_filename}")
    
    log_activity('backup_data', details=f'Created backups: {", ".join(backup_status) if backup_status else "Failed"}')
    
    # Determine flash message and template variables
    if excel_created and db_backup_created:
        flash('สำรองข้อมูลสำเร็จ (ทั้งไฟล์ Excel และฐานข้อมูล)', 'success')
        return render_template('admin/backup.html', 
                               excel_export=f'backups/{excel_filename}',
                               db_backup=f'backups/{db_filename}')
    elif excel_created:
        flash('สำรองข้อมูล Excel สำเร็จ แต่ไม่สามารถสำรองานข้อมูลได้', 'warning')
        return render_template('admin/backup.html', 
                               excel_export=f'backups/{excel_filename}',
                               db_backup=None)
    elif db_backup_created:
        flash('สำรองฐานข้อมูลสำเร็จ แต่ไม่สามารถสร้างไฟล์ Excel ได้', 'warning')
        return render_template('admin/backup.html', 
                               excel_export=None,
                               db_backup=f'backups/{db_filename}')
    else:
        flash('เกิดข้อผิดพลาดในการสำรองข้อมูล กรุณาลองใหม่อีกครั้ง', 'danger')
        return redirect(url_for('main.dashboard'))

@main.route('/download_backup/<path:filename>')
def download_backup(filename):
    if session.get('role') != 'admin':
        return redirect(url_for('main.dashboard'))
    
    # Get the absolute path to the project root
    project_root = os.path.abspath(os.path.dirname(current_app.root_path))
    backup_file_path = os.path.join(project_root, current_app.config['UPLOAD_FOLDER'], filename)
    
    # Check if file exists
    if not os.path.exists(backup_file_path):
        flash('ไม่พบไฟล์สำรองข้อมูลที่ต้องการ', 'danger')
        return redirect(url_for('main.backup_data'))
    
    return send_file(backup_file_path, as_attachment=True)

@main.route('/settings', methods=['GET', 'POST'])
def settings():
    if session.get('role') != 'admin':
        return redirect(url_for('main.dashboard'))
    
    settings = Settings.query.first()
    if not settings:
        settings = Settings()
        db.session.add(settings)
        db.session.commit()

    if request.method == 'POST':
        settings.company_name = request.form['company_name']
        settings.company_address = request.form['company_address']
        settings.company_phone = request.form['company_phone']
        settings.company_email = request.form['company_email']
        settings.low_stock_threshold = int(request.form['low_stock_threshold'])
        
        if 'company_logo' in request.files and request.files['company_logo'].filename:
            file = request.files['company_logo']
            if settings.company_logo:
                delete_file(settings.company_logo)
            logo_path = save_file(file, 'logos')
            if logo_path:
                settings.company_logo = logo_path
        
        db.session.commit()
        flash('บันทึกการตั้งค่าสำเร็จ', 'success')
        return redirect(url_for('main.settings'))
        
    return render_template('admin/settings.html', settings=settings)

# Dummy shop routes for completion
@main.route('/shop')
def shop():
    return redirect(url_for('main.index'))

@main.route('/shop_editor', methods=['GET', 'POST'])
def shop_editor():
    if session.get('role') != 'admin':
        return redirect(url_for('main.dashboard'))
    
    shop_settings = ShopSettings.query.first()
    if not shop_settings:
        store = Store.query.first()
        if not store:
            store = Store(name='Default Store')
            db.session.add(store)
            db.session.commit()
        shop_settings = ShopSettings(store_id=store.id)
        db.session.add(shop_settings)
        db.session.commit()

    products = Product.query.order_by(Product.name).all()
    featured_product_ids = [fp.product_id for fp in FeaturedProduct.query.all()]
    
    # Banners for main shop
    shop_banners_top = Banner.query.filter_by(page_location='shop', position='top').all()
    shop_banners_middle = Banner.query.filter_by(page_location='shop', position='middle').all()
    shop_banners_bottom = Banner.query.filter_by(page_location='shop', position='bottom').all()

    # Banners for stock page
    shop_stock_banners_top = Banner.query.filter_by(page_location='shop_stock', position='top').all()
    shop_stock_banners_middle = Banner.query.filter_by(page_location='shop_stock', position='middle').all()
    shop_stock_banners_bottom = Banner.query.filter_by(page_location='shop_stock', position='bottom').all()

    return render_template('shop_editor.html', 
                           shop_settings=shop_settings,
                           products=products,
                           featured_product_ids=featured_product_ids,
                           shop_banners_top=shop_banners_top,
                           shop_banners_middle=shop_banners_middle,
                           shop_banners_bottom=shop_banners_bottom,
                           shop_stock_banners_top=shop_stock_banners_top,
                           shop_stock_banners_middle=shop_stock_banners_middle,
                           shop_stock_banners_bottom=shop_stock_banners_bottom)

@main.route('/update_shop_hero', methods=['POST'])
def update_shop_hero():
    if session.get('role') != 'admin':
        return redirect(url_for('main.dashboard'))
    settings = ShopSettings.query.first()
    if not settings:
        flash('Shop settings have not been initialized.', 'danger')
        return redirect(url_for('main.shop_editor'))
    settings.hero_title = request.form['hero_title']
    settings.hero_subtitle = request.form['hero_subtitle']
    if 'hero_background' in request.files and request.files['hero_background'].filename:
        file = request.files['hero_background']
        if settings.hero_background:
            delete_file(settings.hero_background)
        hero_background_path = save_file(file, 'banners')
        if hero_background_path:
            settings.hero_background = hero_background_path
    settings.hero_text_color = request.form['hero_text_color']
    settings.hero_button_text = request.form['hero_button_text']
    db.session.commit()
    flash('Shop hero section updated!', 'success')
    return redirect(url_for('main.shop_editor'))

@main.route('/update_featured_products', methods=['POST'])
def update_featured_products():
    if session.get('role') != 'admin':
        return redirect(url_for('main.dashboard'))
    settings = ShopSettings.query.first()
    if not settings:
        flash('Shop settings have not been initialized.', 'danger')
        return redirect(url_for('main.shop_editor'))
    FeaturedProduct.query.delete()
    featured_products = request.form.getlist('featured_products[]')
    for product_id in featured_products:
        fp = FeaturedProduct(product_id=product_id)
        db.session.add(fp)
    db.session.commit()
    flash('Featured products updated!', 'success')
    return redirect(url_for('main.shop_editor'))

@main.route('/update_shop_theme', methods=['POST'])
def update_shop_theme():
    if session.get('role') != 'admin':
        return redirect(url_for('main.dashboard'))
    settings = ShopSettings.query.first()
    if not settings:
        flash('Shop settings have not been initialized.', 'danger')
        return redirect(url_for('main.shop_editor'))
    settings.primary_color = request.form['primary_color']
    settings.secondary_color = request.form['secondary_color']
    settings.accent_color = request.form['accent_color']
    settings.text_color = request.form['text_color']
    settings.font_family = request.form['font_family']
    settings.border_radius = request.form['border_radius']
    db.session.commit()
    flash('Shop theme updated!', 'success')
    return redirect(url_for('main.shop_editor'))

@main.route('/update_shop_footer', methods=['POST'])
def update_shop_footer():
    if session.get('role') != 'admin':
        return redirect(url_for('main.dashboard'))
    settings = ShopSettings.query.first()
    if not settings:
        flash('Shop settings have not been initialized.', 'danger')
        return redirect(url_for('main.shop_editor'))
    settings.footer_text = request.form['footer_text']
    settings.contact_phone = request.form['contact_phone']
    settings.contact_email = request.form['contact_email']
    settings.social_facebook = request.form['social_facebook']
    settings.social_instagram = request.form['social_instagram']
    settings.social_line = request.form['social_line']
    db.session.commit()
    flash('Shop footer updated!', 'success')
    return redirect(url_for('main.shop_editor'))

@main.route('/update_shop_icons', methods=['POST'])
def update_shop_icons():
    if session.get('role') != 'admin':
        return redirect(url_for('main.dashboard'))
    settings = ShopSettings.query.first()
    if not settings:
        flash('Shop settings have not been initialized.', 'danger')
        return redirect(url_for('main.shop_editor'))
    if 'favicon' in request.files and request.files['favicon'].filename:
        file = request.files['favicon']
        if settings.favicon_path:
            delete_file(settings.favicon_path)
        favicon_path = save_file(file, 'logos')
        if favicon_path:
            settings.favicon_path = favicon_path
    if 'navbar_logo' in request.files and request.files['navbar_logo'].filename:
        file = request.files['navbar_logo']
        if settings.navbar_logo_path:
            delete_file(settings.navbar_logo_path)
        navbar_logo_path = save_file(file, 'logos')
        if navbar_logo_path:
            settings.navbar_logo_path = navbar_logo_path
    db.session.commit()
    flash('Shop icons updated!', 'success')
    return redirect(url_for('main.shop_editor'))

@main.route('/add_banner', methods=['POST'])
def add_banner():
    if session.get('role') != 'admin':
        return redirect(url_for('main.dashboard'))
    title = request.form['banner_title']
    link = request.form['banner_link']
    position = request.form['banner_position']
    page_location = request.form['page_location']
    
    banner = Banner(title=title, link=link, position=position, page_location=page_location)
    
    if 'banner_image' in request.files and request.files['banner_image'].filename:
        file = request.files['banner_image']
        image_path = save_file(file, 'banners')
        if image_path:
            banner.image_path = image_path
    
    db.session.add(banner)
    db.session.commit()
    flash('Banner added successfully!', 'success')
    return redirect(url_for('main.shop_editor'))

@main.route('/delete_banner/<int:banner_id>', methods=['POST'])
def delete_banner(banner_id):
    if session.get('role') != 'admin':
        return redirect(url_for('main.dashboard'))
    banner = Banner.query.get_or_404(banner_id)
    if banner.image_path:
        delete_file(banner.image_path)
    db.session.delete(banner)
    db.session.commit()
    flash('Banner deleted successfully!', 'success')
    return redirect(url_for('main.shop_editor'))

@main.route('/toggle_banner/<int:banner_id>', methods=['POST'])
def toggle_banner(banner_id):
    if session.get('role') != 'admin':
        return redirect(url_for('main.dashboard'))
    banner = Banner.query.get_or_404(banner_id)
    banner.is_active = not banner.is_active
    db.session.commit()
    status = 'activated' if banner.is_active else 'deactivated'
    flash(f'Banner {status} successfully!', 'success')
    return redirect(url_for('main.shop_editor'))

@main.route('/shop_success')
def shop_success():
    return render_template('shop_success.html')

@main.route('/upload_slip_and_location', methods=['GET', 'POST'])
def upload_slip_and_location():
    if request.method == 'POST':
        order_id = request.form.get('order_id')
        if not order_id:
            flash('กรุณาระบุหมายเลขคำสั่งซื้อ', 'danger')
            return render_template('upload_slip_and_location.html')
        
        order = Order.query.get(order_id)
        if not order:
            flash('ไม่พบคำสั่งซื้อที่ระบุ', 'danger')
            return render_template('upload_slip_and_location.html')
        
        # Handle payment slip upload
        if 'payment_slip' in request.files and request.files['payment_slip'].filename:
            file = request.files['payment_slip']
            if order.payment_slip:
                delete_file(order.payment_slip)
            payment_slip = save_file(file, 'slips')
            if payment_slip:
                order.payment_slip = payment_slip
                order.payment_date = datetime.datetime.now()
                order.payment_status = 'paid'
        
        # Handle location update
        location = request.form.get('location')
        if location:
            order.shipping_address = location
        
        db.session.commit()
        flash('อัปเดตข้อมูลสำเร็จ', 'success')
        return redirect(url_for('main.shop_success'))
    
    return render_template('upload_slip_and_location.html')

@main.route('/debug_index')
def debug_index():
    return render_template('debug_index.html')
