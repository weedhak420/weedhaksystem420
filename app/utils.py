import os
import time
import base64
import uuid
from io import BytesIO
import qrcode
from flask import flash, current_app, session, request
from werkzeug.utils import secure_filename
from app import db
from app.models import ActivityLog, Notification, User
import datetime
import requests

# Google Sheets imports
try:
    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
    GOOGLE_SHEETS_AVAILABLE = True
except ImportError:
    GOOGLE_SHEETS_AVAILABLE = False
    print("Google Sheets API not available. Install google-api-python-client to enable Google Sheets integration.")

# Global flag to track Google Sheets status
GOOGLE_SHEETS_STATUS = {
    'enabled': True,
    'last_error': None,
    'error_count': 0,
    'last_success': None
}

def allowed_file(filename):
    ALLOWED_EXTENSIONS = {'jpg', 'jpeg', 'png', 'gif', 'webp'}
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def save_file(file, folder):
    """Save uploaded file and return the relative path"""
    if file and file.filename:
        if not allowed_file(file.filename):
            flash('รูปแบบไฟล์ไม่ถูกต้อง กรุณาอัปโหลดไฟล์รูปภาพเท่านั้น (JPG, PNG, GIF, WEBP)', 'danger')
            return None
        
        if file.content_length and file.content_length > current_app.config['MAX_CONTENT_LENGTH']:
            flash(f"ขนาดไฟล์ใหญ่เกินไป กรุณาอัปโหลดไฟล์ขนาดไม่เกิน {current_app.config['MAX_CONTENT_LENGTH'] / (1024 * 1024):.0f}MB", 'danger')
            return None
            
        filename = secure_filename(file.filename)
        if not filename:
            return None
        
        # Add timestamp to filename to avoid conflicts
        name, ext = os.path.splitext(filename)
        timestamp = str(int(datetime.datetime.now().timestamp()))
        filename = f"{name}_{timestamp}{ext}"
        
        # Create folder if it doesn't exist
        folder_path = os.path.join(current_app.config['UPLOAD_FOLDER'], folder)
        os.makedirs(folder_path, exist_ok=True)
        
        # Save file
        file_path = os.path.join(folder_path, filename)
        try:
            file.save(file_path)
            return f"{folder}/{filename}"
        except Exception as e:
            current_app.logger.error(f"Error saving file: {str(e)}")
            flash('เกิดข้อผิดพลาดในการอัปโหลดรูปภาพ กรุณาลองใหม่อีกครั้ง', 'danger')
            return None
    return None

def delete_file(file_path):
    """Delete file from filesystem"""
    if file_path:
        full_path = os.path.join(current_app.config['UPLOAD_FOLDER'], file_path)
        full_path = os.path.normpath(full_path)
        try:
            if os.path.exists(full_path):
                os.remove(full_path)
        except Exception as e:
            current_app.logger.error(f"Error deleting file {file_path}: {str(e)}")

def log_activity(action, entity_type=None, entity_id=None, details=None):
    """Log user activity"""
    try:
        user_id = session.get('user_id') if session else None
        ip_address = request.remote_addr if request else None
        
        activity = ActivityLog(
            user_id=user_id,
            action=action,
            entity_type=entity_type,
            entity_id=entity_id,
            details=details,
            ip_address=ip_address
        )
        
        db.session.add(activity)
        db.session.commit()
    except Exception as e:
        if current_app:
            current_app.logger.error(f"Error logging activity: {str(e)}")
        else:
            print(f"Error logging activity: {str(e)}")

def create_notification(message, notification_type=None, related_id=None, user_id=None):
    """Create notification for user(s)"""
    try:
        if user_id:
            users = [User.query.get(user_id)]
        else:
            # Send to all admin users
            users = User.query.filter_by(role='admin').all()
        
        for user in users:
            if user:
                notification = Notification(
                    user_id=user.id,
                    message=message,
                    type=notification_type,
                    related_id=related_id
                )
                db.session.add(notification)
        
        db.session.commit()
    except Exception as e:
        if current_app:
            current_app.logger.error(f"Error creating notification: {str(e)}")
        else:
            print(f"Error creating notification: {str(e)}")

def generate_qr_code(order_id):
    """Generate QR code for order and return the file path"""
    try:
        # Create QR code data
        qr_data = f"Order #{order_id} - Marbo9K Shop"
        
        # Generate QR code
        qr = qrcode.QRCode(version=1, box_size=10, border=5)
        qr.add_data(qr_data)
        qr.make(fit=True)
        
        # Create QR code image
        qr_image = qr.make_image(fill_color="black", back_color="white")
        
        # Save QR code
        filename = f"qrcode_{order_id}.png"
        folder_path = os.path.join(current_app.config['UPLOAD_FOLDER'], 'qrcodes')
        os.makedirs(folder_path, exist_ok=True)
        
        file_path = os.path.join(folder_path, filename)
        qr_image.save(file_path)
        
        return f"qrcodes/{filename}"
    except Exception as e:
        if current_app:
            current_app.logger.error(f"Error generating QR code: {str(e)}")
        else:
            print(f"Error generating QR code: {str(e)}")
        return None

def generate_qr_code_base64(data):
    """Generate QR code and return as base64 string"""
    try:
        qr = qrcode.QRCode(version=1, box_size=10, border=5)
        qr.add_data(data)
        qr.make(fit=True)
        
        qr_image = qr.make_image(fill_color="black", back_color="white")
        
        # Convert to base64
        buffer = BytesIO()
        qr_image.save(buffer, format='PNG')
        buffer.seek(0)
        
        qr_base64 = base64.b64encode(buffer.getvalue()).decode()
        return f"data:image/png;base64,{qr_base64}"
    except Exception as e:
        if current_app:
            current_app.logger.error(f"Error generating QR code base64: {str(e)}")
        else:
            print(f"Error generating QR code base64: {str(e)}")
        return None

# Google Sheets Integration Functions

def check_google_sheets_config():
    """Check if Google Sheets is properly configured"""
    if not GOOGLE_SHEETS_AVAILABLE:
        return False, "Google Sheets API library not installed"
    
    # Handle case when running outside Flask context
    try:
        credentials_file = current_app.config.get('GOOGLE_SHEETS_CREDENTIALS_FILE')
        spreadsheet_id = current_app.config.get('GOOGLE_SHEETS_SPREADSHEET_ID')
    except RuntimeError:
        # Running outside Flask context, try to get from environment or config file
        try:
            import config
            credentials_file = getattr(config, 'GOOGLE_SHEETS_CREDENTIALS_FILE', None)
            spreadsheet_id = getattr(config, 'GOOGLE_SHEETS_SPREADSHEET_ID', None)
        except ImportError:
            return False, "Cannot access configuration (no Flask context and no config.py)"
    
    if not credentials_file:
        return False, "GOOGLE_SHEETS_CREDENTIALS_FILE not configured"
    
    if not os.path.exists(credentials_file):
        return False, f"Credentials file not found: {credentials_file}"
    
    if not spreadsheet_id:
        return False, "GOOGLE_SHEETS_SPREADSHEET_ID not configured"
    
    return True, "Configuration OK"

def update_google_sheets_status(success=True, error_message=None):
    """Update Google Sheets status tracking"""
    global GOOGLE_SHEETS_STATUS
    
    if success:
        GOOGLE_SHEETS_STATUS['last_success'] = datetime.datetime.now()
        GOOGLE_SHEETS_STATUS['error_count'] = 0
        GOOGLE_SHEETS_STATUS['last_error'] = None
        # Re-enable if it was disabled due to temporary errors
        if not GOOGLE_SHEETS_STATUS['enabled'] and GOOGLE_SHEETS_STATUS['error_count'] == 0:
            GOOGLE_SHEETS_STATUS['enabled'] = True
            if current_app:
                current_app.logger.info("Google Sheets integration re-enabled after successful operation")
            else:
                print("Google Sheets integration re-enabled after successful operation")
    else:
        GOOGLE_SHEETS_STATUS['error_count'] += 1
        GOOGLE_SHEETS_STATUS['last_error'] = error_message
        
        # Disable after multiple consecutive errors
        if GOOGLE_SHEETS_STATUS['error_count'] >= 3:
            GOOGLE_SHEETS_STATUS['enabled'] = False
            if current_app:
                current_app.logger.warning(f"Google Sheets integration disabled after {GOOGLE_SHEETS_STATUS['error_count']} consecutive errors")
            else:
                print(f"Google Sheets integration disabled after {GOOGLE_SHEETS_STATUS['error_count']} consecutive errors")

def _execute_sheets_api_call(api_call, operation_name="Unknown"):
    """Wrapper to execute Google Sheets API calls with centralized error handling."""
    global GOOGLE_SHEETS_STATUS
    
    if not GOOGLE_SHEETS_STATUS['enabled']:
        message = f"Google Sheets integration is disabled, skipping {operation_name}"
        if current_app:
            current_app.logger.info(message)
        else:
            print(message)
        return None

    # Check configuration before making API call
    config_ok, config_message = check_google_sheets_config()
    if not config_ok:
        error_message = f"Google Sheets configuration error: {config_message}"
        if current_app:
            current_app.logger.error(error_message)
        else:
            print(f"ERROR: {error_message}")
        update_google_sheets_status(False, config_message)
        return None

    try:
        result = api_call.execute()
        update_google_sheets_status(True)
        return result
        
    except HttpError as error:
        status_code = error.resp.status if hasattr(error.resp, 'status') else None
        error_message = f"Google Sheets API error in {operation_name}: {error} (Status: {status_code})"
        
        if current_app:
            current_app.logger.error(error_message)
        else:
            print(f"ERROR: {error_message}")
        
        if status_code == 403:
            permission_error = (
                "Google Sheets permission denied (403 Forbidden). "
                "Please ensure the service account has 'Editor' permissions for the Google Sheet. "
                "Steps to fix: 1) Open your Google Sheet, 2) Click 'Share', "
                "3) Add the service account email with 'Editor' permissions"
            )
            if current_app:
                current_app.logger.error(permission_error)
            else:
                print(f"ERROR: {permission_error}")
            update_google_sheets_status(False, permission_error)
            
            # Create notification for admin users if in Flask context
            try:
                create_notification(
                    "Google Sheets integration failed: Permission denied. Please check service account permissions.",
                    notification_type="error",
                    related_id=None
                )
            except:
                pass  # Ignore if not in Flask context
            
        elif status_code == 404:
            not_found_error = (
                "Google Sheets API error: Spreadsheet or range not found (404 Not Found). "
                "Please check your SPREADSHEET_ID and RANGE configuration."
            )
            if current_app:
                current_app.logger.error(not_found_error)
            else:
                print(f"ERROR: {not_found_error}")
            update_google_sheets_status(False, not_found_error)
            
        elif status_code == 400:
            bad_request_error = (
                f"Google Sheets API error: Bad request (400). "
                f"This usually means the sheet name or range format is incorrect. "
                f"Error details: {error}"
            )
            if current_app:
                current_app.logger.error(bad_request_error)
            else:
                print(f"ERROR: {bad_request_error}")
            update_google_sheets_status(False, bad_request_error)
            
            # Create notification for admin users if in Flask context
            try:
                create_notification(
                    "Google Sheets integration failed: Invalid sheet name or range format. Please check your Google Sheets configuration.",
                    notification_type="error",
                    related_id=None
                )
            except:
                pass  # Ignore if not in Flask context
            
        else:
            update_google_sheets_status(False, error_message)
            
        return None
        
    except Exception as e:
        error_message = f"Unexpected error in Google Sheets API ({operation_name}): {str(e)}"
        if current_app:
            current_app.logger.error(error_message)
        else:
            print(f"ERROR: {error_message}")
        update_google_sheets_status(False, error_message)
        return None

def get_google_sheets_service():
    """Initialize and return Google Sheets service"""
    config_ok, config_message = check_google_sheets_config()
    if not config_ok:
        error_message = f"Google Sheets configuration error: {config_message}"
        if current_app:
            current_app.logger.error(error_message)
        else:
            print(f"ERROR: {error_message}")
        return None
    
    try:
        # Handle case when running outside Flask context
        try:
            credentials_file = current_app.config.get('GOOGLE_SHEETS_CREDENTIALS_FILE')
        except RuntimeError:
            # Running outside Flask context, try to get from config file
            import config
            credentials_file = getattr(config, 'GOOGLE_SHEETS_CREDENTIALS_FILE', None)
        
        # Define the scope
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
        
        # Load credentials
        credentials = Credentials.from_service_account_file(credentials_file, scopes=SCOPES)
        
        # Build the service
        service = build('sheets', 'v4', credentials=credentials)
        return service
        
    except Exception as e:
        error_message = f"Error initializing Google Sheets service: {str(e)}"
        if current_app:
            current_app.logger.error(error_message)
        else:
            print(f"ERROR: {error_message}")
        update_google_sheets_status(False, error_message)
        return None

def get_google_sheets_status():
    """Get current Google Sheets integration status"""
    return GOOGLE_SHEETS_STATUS.copy()

def reset_google_sheets_integration():
    """Reset Google Sheets integration status (for admin use)"""
    global GOOGLE_SHEETS_STATUS
    GOOGLE_SHEETS_STATUS = {
        'enabled': True,
        'last_error': None,
        'error_count': 0,
        'last_success': None
    }
    message = "Google Sheets integration status reset"
    if current_app:
        current_app.logger.info(message)
    else:
        print(message)

def get_or_create_sheet(service, spreadsheet_id, sheet_name):
    """Get existing sheet or create new one if it doesn't exist"""
    try:
        # Get spreadsheet metadata
        sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = sheet_metadata.get('sheets', [])
        sheet_names = [sheet['properties']['title'] for sheet in sheets]
        
        if sheet_name not in sheet_names:
            message = f"Sheet '{sheet_name}' not found, creating new sheet"
            if current_app:
                current_app.logger.info(message)
            else:
                print(message)
            
            # Create new sheet
            requests = [{
                'addSheet': {
                    'properties': {
                        'title': sheet_name,
                        'gridProperties': {
                            'rowCount': 1000,
                            'columnCount': 26
                        }
                    }
                }
            }]
            
            batch_update_call = service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={'requests': requests}
            )
            
            result = _execute_sheets_api_call(batch_update_call, f"create_sheet_{sheet_name}")
            if result:
                message = f"Sheet '{sheet_name}' created successfully"
                if current_app:
                    current_app.logger.info(message)
                else:
                    print(message)
                return True
            else:
                error_message = f"Failed to create sheet '{sheet_name}'"
                if current_app:
                    current_app.logger.error(error_message)
                else:
                    print(f"ERROR: {error_message}")
                return False
        else:
            message = f"Sheet '{sheet_name}' already exists"
            if current_app:
                current_app.logger.info(message)
            else:
                print(message)
            return True
            
    except Exception as e:
        error_message = f"Error checking/creating sheet '{sheet_name}': {str(e)}"
        if current_app:
            current_app.logger.error(error_message)
        else:
            print(f"ERROR: {error_message}")
        return False

def setup_google_sheets_structure():
    """Setup Google Sheets with proper headers and structure"""
    if not GOOGLE_SHEETS_STATUS['enabled']:
        message = "Google Sheets integration is disabled, skipping structure setup"
        if current_app:
            current_app.logger.info(message)
        else:
            print(message)
        return False

    service = get_google_sheets_service()
    if not service:
        return False

    # Handle case when running outside Flask context
    try:
        spreadsheet_id = current_app.config.get('GOOGLE_SHEETS_SPREADSHEET_ID')
    except RuntimeError:
        import config
        spreadsheet_id = getattr(config, 'GOOGLE_SHEETS_SPREADSHEET_ID', None)

    # Ensure Products sheet exists
    if not get_or_create_sheet(service, spreadsheet_id, 'Products'):
        return False

    # Define headers for the products sheet
    headers = [
        'Timestamp', 'Action', 'Product_ID', 'Product_Name', 'Flavor',
        'Description', 'Price', 'Cost', 'Wholesale_Price', 'Stock',
        'Barcode', 'Profit_Margin', 'Stock_Value', 'User_ID', 'User_Name',
        'Created_At', 'Updated_At'
    ]

    # Check if headers already exist
    range_name = 'Products!A1:Q1'
    api_call_get = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=range_name
    )
    result = _execute_sheets_api_call(api_call_get, "check_headers")

    existing_values = result.get('values', []) if result else []

    # If no headers exist or result is None (error occurred), create them
    if not result or not existing_values or not existing_values[0]:
        message = "Creating headers in Products sheet"
        if current_app:
            current_app.logger.info(message)
        else:
            print(message)
        
        body = {'values': [headers]}
        api_call_update = service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=range_name,
            valueInputOption='RAW',
            body=body
        )
        update_result = _execute_sheets_api_call(api_call_update, "create_headers")
        if update_result:
            message = "Google Sheets headers created successfully"
            if current_app:
                current_app.logger.info(message)
            else:
                print(message)
            return True
        else:
            error_message = "Failed to create Google Sheets headers"
            if current_app:
                current_app.logger.error(error_message)
            else:
                print(f"ERROR: {error_message}")
            return False

    message = "Google Sheets headers already exist"
    if current_app:
        current_app.logger.info(message)
    else:
        print(message)
    return True

def add_product_to_google_sheets(product_data):
    """Add product to Google Sheets"""
    if not GOOGLE_SHEETS_STATUS['enabled']:
        message = "Google Sheets integration is disabled, skipping product addition"
        if current_app:
            current_app.logger.info(message)
        else:
            print(message)
        return False

    service = get_google_sheets_service()
    if not service:
        return False

    # Handle case when running outside Flask context
    try:
        spreadsheet_id = current_app.config.get('GOOGLE_SHEETS_SPREADSHEET_ID')
    except RuntimeError:
        import config
        spreadsheet_id = getattr(config, 'GOOGLE_SHEETS_SPREADSHEET_ID', None)

    # Ensure sheet structure is set up
    if not setup_google_sheets_structure():
        return False

    # Prepare the data row
    values = [
        [
            product_data.get('id', ''),
            product_data.get('name', ''),
            product_data.get('flavor', ''),
            product_data.get('description', ''),
            product_data.get('price', 0),
            product_data.get('cost', 0),
            product_data.get('stock', 0),
            product_data.get('barcode', '')
        ]
    ]

    body = {'values': values}

    # Use a simpler range that should work
    range_name = 'Products!A:H'
    api_call = service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id,
        range=range_name,
        valueInputOption='RAW',
        insertDataOption='INSERT_ROWS',
        body=body
    )
    result = _execute_sheets_api_call(api_call, "add_product")

    if result:
        message = f"Product added to Google Sheets: {result.get('updates', {}).get('updatedCells', 0)} cells updated"
        if current_app:
            current_app.logger.info(message)
        else:
            print(message)
        return True
    
    return False

def update_product_in_google_sheets(product_id, product_data):
    """Update product in Google Sheets"""
    if not GOOGLE_SHEETS_STATUS['enabled']:
        message = "Google Sheets integration is disabled, skipping product update"
        if current_app:
            current_app.logger.info(message)
        else:
            print(message)
        return False

    service = get_google_sheets_service()
    if not service:
        return False

    # Handle case when running outside Flask context
    try:
        spreadsheet_id = current_app.config.get('GOOGLE_SHEETS_SPREADSHEET_ID')
    except RuntimeError:
        import config
        spreadsheet_id = getattr(config, 'GOOGLE_SHEETS_SPREADSHEET_ID', None)

    # Ensure sheet structure is set up
    if not setup_google_sheets_structure():
        return False

    # First, find the row with the product ID
    range_name = 'Products!A:H'
    api_call_get = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=range_name
    )
    result = _execute_sheets_api_call(api_call_get, "get_products_for_update")

    if not result:
        return False

    values = result.get('values', [])
    row_index = None

    # Find the row with matching product ID (skip header row)
    for i, row in enumerate(values[1:], start=2):
        if len(row) > 0 and str(row[0]) == str(product_id):
            row_index = i
            break

    if row_index is None:
        # Product not found, add it instead
        message = f"Product {product_id} not found in Google Sheets, adding as new product"
        if current_app:
            current_app.logger.info(message)
        else:
            print(message)
        return add_product_to_google_sheets(product_data)

    # Update the row
    update_values = [
        product_data.get('id', ''),
        product_data.get('name', ''),
        product_data.get('flavor', ''),
        product_data.get('description', ''),
        product_data.get('price', 0),
        product_data.get('cost', 0),
        product_data.get('stock', 0),
        product_data.get('barcode', '')
    ]

    update_range = f"Products!A{row_index}:H{row_index}"
    body = {'values': [update_values]}

    api_call_update = service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=update_range,
        valueInputOption='RAW',
        body=body
    )
    result_update = _execute_sheets_api_call(api_call_update, "update_product")

    if result_update:
        message = f"Product updated in Google Sheets: {result_update.get('updatedCells', 0)} cells updated"
        if current_app:
            current_app.logger.info(message)
        else:
            print(message)
        return True
    
    return False

def sync_all_products_to_google_sheets():
    """Sync all products from database to Google Sheets"""
    if not GOOGLE_SHEETS_STATUS['enabled']:
        message = "Google Sheets integration is disabled, skipping full sync"
        if current_app:
            current_app.logger.info(message)
        else:
            print(message)
        return False

    from app.models import Product
    
    service = get_google_sheets_service()
    if not service:
        return False

    # Handle case when running outside Flask context
    try:
        spreadsheet_id = current_app.config.get('GOOGLE_SHEETS_SPREADSHEET_ID')
    except RuntimeError:
        import config
        spreadsheet_id = getattr(config, 'GOOGLE_SHEETS_SPREADSHEET_ID', None)

    # Ensure sheet structure is set up
    if not setup_google_sheets_structure():
        return False

    # Get all products from database
    products = Product.query.all()

    # Prepare header row
    header = ['ID', 'Name', 'Flavor', 'Description', 'Price', 'Cost', 'Stock', 'Barcode']
    values = [header]

    # Add product data
    for product in products:
        values.append([
            product.id,
            product.name,
            product.flavor,
            product.description or '',
            float(product.price),
            float(product.cost),
            product.stock,
            product.barcode or ''
        ])

    # Clear existing data and add new data
    range_name = 'Products!A:H'

    # Clear the range first
    api_call_clear = service.spreadsheets().values().clear(
        spreadsheetId=spreadsheet_id,
        range=range_name
    )
    if not _execute_sheets_api_call(api_call_clear, "clear_products"):
        return False

    # Add new data
    body = {'values': values}

    api_call_update = service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=range_name,
        valueInputOption='RAW',
        body=body
    )
    result = _execute_sheets_api_call(api_call_update, "sync_all_products")

    if result:
        message = f"All products synced to Google Sheets: {result.get('updatedCells', 0)} cells updated"
        if current_app:
            current_app.logger.info(message)
        else:
            print(message)
        return True
    
    return False

def validate_api_key(api_key):
    """Validate API key"""
    try:
        expected_key = current_app.config.get('API_KEY')
    except RuntimeError:
        import config
        expected_key = getattr(config, 'API_KEY', None)
    
    return api_key == expected_key

def add_product_to_google_sheets_realtime(product_data, action='ADD', user_info=None):
    """Add product to Google Sheets with timestamp and real-time data for n8n"""
    if not GOOGLE_SHEETS_STATUS['enabled']:
        message = "Google Sheets integration is disabled, skipping realtime update"
        if current_app:
            current_app.logger.info(message)
        else:
            print(message)
        return False

    service = get_google_sheets_service()
    if not service:
        return False

    # Handle case when running outside Flask context
    try:
        spreadsheet_id = current_app.config.get('GOOGLE_SHEETS_SPREADSHEET_ID')
    except RuntimeError:
        import config
        spreadsheet_id = getattr(config, 'GOOGLE_SHEETS_SPREADSHEET_ID', None)

    # Setup structure if needed
    if not setup_google_sheets_structure():
        return False

    # Get current timestamp
    now = datetime.datetime.now()
    timestamp = now.strftime('%Y-%m-%d %H:%M:%S')

    # Calculate additional fields
    price = float(product_data.get('price', 0))
    cost = float(product_data.get('cost', 0))
    stock = int(product_data.get('stock', 0))
    profit_margin = ((price - cost) / price * 100) if price > 0 else 0
    stock_value = price * stock

    # Prepare the data row
    values = [[
        timestamp,
        action,
        product_data.get('id', ''),
        product_data.get('name', ''),
        product_data.get('flavor', ''),
        product_data.get('description', ''),
        price,
        cost,
        float(product_data.get('wholesale_price', 0)) if product_data.get('wholesale_price') else '',
        stock,
        product_data.get('barcode', ''),
        round(profit_margin, 2),
        round(stock_value, 2),
        user_info.get('user_id', '') if user_info else '',
        user_info.get('username', '') if user_info else '',
        product_data.get('created_at', timestamp),
        timestamp
    ]]

    body = {'values': values}

    # Append the data to the sheet
    range_name = 'Products!A:Q'
    api_call = service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id,
        range=range_name,
        valueInputOption='RAW',
        insertDataOption='INSERT_ROWS',
        body=body
    )
    result = _execute_sheets_api_call(api_call, "add_realtime_product")

    if result:
        message = f"Product {action} synced to Google Sheets: {result.get('updates', {}).get('updatedCells', 0)} cells updated"
        if current_app:
            current_app.logger.info(message)
        else:
            print(message)
        # Also update the main products sheet for current inventory
        update_main_products_sheet(product_data)
        return True
    
    return False

def update_main_products_sheet(product_data):
    """Update main products sheet for current inventory status"""
    if not GOOGLE_SHEETS_STATUS['enabled']:
        return False

    service = get_google_sheets_service()
    if not service:
        return False

    # Handle case when running outside Flask context
    try:
        spreadsheet_id = current_app.config.get('GOOGLE_SHEETS_SPREADSHEET_ID')
    except RuntimeError:
        import config
        spreadsheet_id = getattr(config, 'GOOGLE_SHEETS_SPREADSHEET_ID', None)

    # Ensure Current_Inventory sheet exists
    if not get_or_create_sheet(service, spreadsheet_id, 'Current_Inventory'):
        return False

    # Check if headers exist in Current_Inventory sheet
    headers = [
        'Product_ID', 'Product_Name', 'Flavor', 'Description', 
        'Price', 'Cost', 'Wholesale_Price', 'Stock', 'Barcode',
        'Profit_Margin', 'Stock_Value', 'Last_Updated'
    ]

    # Check if headers already exist
    header_range = 'Current_Inventory!A1:L1'
    api_call_get_headers = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=header_range
    )
    header_result = _execute_sheets_api_call(api_call_get_headers, "check_inventory_headers")

    existing_headers = header_result.get('values', []) if header_result else []

    # If no headers exist, create them
    if not header_result or not existing_headers or not existing_headers[0]:
        body = {'values': [headers]}
        header_update_call = service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=header_range,
            valueInputOption='RAW',
            body=body
        )
        if not _execute_sheets_api_call(header_update_call, "create_inventory_headers"):
            return False

    # Find and update or insert product in Current_Inventory sheet
    product_id = str(product_data.get('id', ''))
    if not product_id:
        return False

    # Get current data
    get_call = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range='Current_Inventory!A:L'
    )
    result = _execute_sheets_api_call(get_call, "get_inventory_data")
    values = result.get('values', []) if result else []
    row_index = None

    # Find the row with matching product ID (skip header row)
    for i, row in enumerate(values[1:], start=2):
        if len(row) > 0 and str(row[0]) == product_id:
            row_index = i
            break

    # Prepare update data
    price = float(product_data.get('price', 0))
    cost = float(product_data.get('cost', 0))
    stock = int(product_data.get('stock', 0))
    profit_margin = ((price - cost) / price * 100) if price > 0 else 0
    stock_value = price * stock
    timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    update_values = [
        product_data.get('id', ''),
        product_data.get('name', ''),
        product_data.get('flavor', ''),
        product_data.get('description', ''),
        price,
        cost,
        float(product_data.get('wholesale_price', 0)) if product_data.get('wholesale_price') else '',
        stock,
        product_data.get('barcode', ''),
        round(profit_margin, 2),
        round(stock_value, 2),
        timestamp
    ]

    if row_index:
        # Update existing row
        update_range = f"Current_Inventory!A{row_index}:L{row_index}"
        body = {'values': [update_values]}
        update_call = service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id, range=update_range,
            valueInputOption='RAW', body=body
        )
        _execute_sheets_api_call(update_call, "update_inventory_row")
    else:
        # Append new row
        body = {'values': [update_values]}
        append_call = service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id, range='Current_Inventory!A:L',
            valueInputOption='RAW', insertDataOption='INSERT_ROWS', body=body
        )
        _execute_sheets_api_call(append_call, "append_inventory_row")

    return True

def create_n8n_webhook_data(product_data, action='ADD', user_info=None):
    """Create structured data for n8n webhook consumption"""
    timestamp = datetime.datetime.now()
    
    # Handle case when running outside Flask context
    try:
        environment = current_app.config.get('ENV', 'production')
    except RuntimeError:
        environment = 'development'
    
    webhook_data = {
        'event': 'product_change',
        'action': action,
        'timestamp': timestamp.isoformat(),
        'timestamp_unix': int(timestamp.timestamp()),
        'product': {
            'id': product_data.get('id'),
            'name': product_data.get('name'),
            'flavor': product_data.get('flavor'),
            'description': product_data.get('description'),
            'price': float(product_data.get('price', 0)),
            'cost': float(product_data.get('cost', 0)),
            'wholesale_price': float(product_data.get('wholesale_price', 0)) if product_data.get('wholesale_price') else None,
            'stock': int(product_data.get('stock', 0)),
            'barcode': product_data.get('barcode'),
            'profit_margin': ((float(product_data.get('price', 0)) - float(product_data.get('cost', 0))) / float(product_data.get('price', 1)) * 100),
            'stock_value': float(product_data.get('price', 0)) * int(product_data.get('stock', 0))
        },
        'user': user_info if user_info else None,
        'system': {
            'source': 'marbo9k_system',
            'version': '1.0',
            'environment': environment
        },
        'google_sheets_status': get_google_sheets_status()
    }
    
    return webhook_data

def send_to_n8n_webhook(webhook_data):
    """Send data to n8n webhook if configured"""
    try:
        # Handle case when running outside Flask context
        try:
            webhook_url = current_app.config.get('N8N_WEBHOOK_URL')
        except RuntimeError:
            import config
            webhook_url = getattr(config, 'N8N_WEBHOOK_URL', None)
        
        if not webhook_url:
            message = "N8N webhook URL not configured, skipping webhook"
            if current_app:
                current_app.logger.info(message)
            else:
                print(message)
            return False
        
        response = requests.post(
            webhook_url,
            json=webhook_data,
            headers={'Content-Type': 'application/json'},
            timeout=10
        )
        
        if response.status_code == 200:
            message = "Data sent to n8n webhook successfully"
            if current_app:
                current_app.logger.info(message)
            else:
                print(message)
            return True
        else:
            error_message = f"n8n webhook returned status {response.status_code}: {response.text}"
            if current_app:
                current_app.logger.error(error_message)
            else:
                print(f"ERROR: {error_message}")
            return False
            
    except Exception as e:
        error_message = f"Error sending to n8n webhook: {str(e)}"
        if current_app:
            current_app.logger.error(error_message)
        else:
            print(f"ERROR: {error_message}")
        return False

def update_stock_in_google_sheets_for_order(order_items, action='SALE', user_info=None):
    """Update stock in Google Sheets when order is created/deleted"""
    if not GOOGLE_SHEETS_STATUS['enabled']:
        message = "Google Sheets integration is disabled, skipping stock update"
        if current_app:
            current_app.logger.info(message)
        else:
            print(message)
        return False

    service = get_google_sheets_service()
    if not service:
        return False

    # Handle case when running outside Flask context
    try:
        spreadsheet_id = current_app.config.get('GOOGLE_SHEETS_SPREADSHEET_ID')
    except RuntimeError:
        import config
        spreadsheet_id = getattr(config, 'GOOGLE_SHEETS_SPREADSHEET_ID', None)

    # Ensure Stock_Transactions sheet exists
    if not get_or_create_sheet(service, spreadsheet_id, 'Stock_Transactions'):
        return False

    # Setup headers for Stock_Transactions sheet
    headers = [
        'Timestamp', 'Action', 'Order_ID', 'Product_ID', 'Product_Name', 
        'Flavor', 'Quantity_Changed', 'New_Stock_Level', 'Unit_Price',
        'Total_Value', 'User_ID', 'User_Name'
    ]

    # Check if headers exist
    header_range = 'Stock_Transactions!A1:L1'
    api_call_get_headers = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=header_range
    )
    header_result = _execute_sheets_api_call(api_call_get_headers, "check_stock_headers")

    existing_headers = header_result.get('values', []) if header_result else []

    # Create headers if they don't exist
    if not header_result or not existing_headers or not existing_headers[0]:
        body = {'values': [headers]}
        header_update_call = service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=header_range,
            valueInputOption='RAW',
            body=body
        )
        if not _execute_sheets_api_call(header_update_call, "create_stock_headers"):
            return False

    # Process each order item
    from app.models import Product
    timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    all_transactions = []

    for item in order_items:
        product = Product.query.get(item.product_id)
        if not product:
            continue

        # Calculate quantity change (negative for sales, positive for returns)
        quantity_change = -item.quantity if action == 'SALE' else item.quantity
        
        transaction_row = [
            timestamp,
            action,
            item.order_id,
            product.id,
            product.name,
            product.flavor,
            quantity_change,
            product.stock,  # Current stock level after the transaction
            float(item.price),
            float(item.price * item.quantity),
            user_info.get('user_id', '') if user_info else '',
            user_info.get('username', '') if user_info else ''
        ]
        all_transactions.append(transaction_row)

        # Also update the Current_Inventory sheet
        product_data = {
            'id': product.id,
            'name': product.name,
            'flavor': product.flavor,
            'description': product.description,
            'price': float(product.price),
            'cost': float(product.cost),
            'wholesale_price': float(product.wholesale_price) if product.wholesale_price else None,
            'stock': product.stock,
            'barcode': product.barcode or ''
        }
        update_main_products_sheet(product_data)

    # Add all transactions to Stock_Transactions sheet
    if all_transactions:
        body = {'values': all_transactions}
        append_call = service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id,
            range='Stock_Transactions!A:L',
            valueInputOption='RAW',
            insertDataOption='INSERT_ROWS',
            body=body
        )
        result = _execute_sheets_api_call(append_call, "add_stock_transactions")

        if result:
            message = f"Stock transactions added to Google Sheets: {len(all_transactions)} transactions"
            if current_app:
                current_app.logger.info(message)
            else:
                print(message)
            return True

    return False

def create_order_webhook_data(order, order_items, action='CREATE', user_info=None):
    """Create structured data for order webhook"""
    timestamp = datetime.datetime.now()
    
    # Handle case when running outside Flask context
    try:
        environment = current_app.config.get('ENV', 'production')
    except RuntimeError:
        environment = 'development'
    
    # Prepare order items data
    items_data = []
    for item in order_items:
        from app.models import Product
        product = Product.query.get(item.product_id)
        if product:
            items_data.append({
                'product_id': item.product_id,
                'product_name': product.name,
                'flavor': product.flavor,
                'quantity': item.quantity,
                'unit_price': float(item.price),
                'total_price': float(item.price * item.quantity),
                'new_stock_level': product.stock
            })
    
    webhook_data = {
        'event': 'order_change',
        'action': action,
        'timestamp': timestamp.isoformat(),
        'timestamp_unix': int(timestamp.timestamp()),
        'order': {
            'id': order.id,
            'customer_id': order.customer_id,
            'customer_name': order.customer.name if order.customer else None,
            'total_amount': float(order.total_amount),
            'payment_status': order.payment_status,
            'status': order.status,
            'order_date': order.order_date.isoformat() if order.order_date else None,
            'items': items_data,
            'item_count': len(items_data)
        },
        'user': user_info if user_info else None,
        'system': {
            'source': 'marbo9k_system',
            'version': '1.0',
            'environment': environment
        },
        'google_sheets_status': get_google_sheets_status()
    }
    
    return webhook_data

def test_google_sheets_connection():
    """Test Google Sheets connection and return detailed status"""
    test_result = {
        'success': False,
        'message': '',
        'details': {},
        'suggestions': []
    }
    
    # Check if Google Sheets API is available
    if not GOOGLE_SHEETS_AVAILABLE:
        test_result['message'] = 'Google Sheets API library not installed'
        test_result['suggestions'].append('Install google-api-python-client: pip install google-api-python-client')
        return test_result
    
    # Check configuration
    config_ok, config_message = check_google_sheets_config()
    if not config_ok:
        test_result['message'] = config_message
        test_result['suggestions'].append('Check your configuration settings')
        return test_result
    
    # Try to get service
    service = get_google_sheets_service()
    if not service:
        test_result['message'] = 'Failed to initialize Google Sheets service'
        test_result['suggestions'].append('Check your credentials file')
        return test_result
    
    # Try to access spreadsheet
    try:
        # Handle case when running outside Flask context
        try:
            spreadsheet_id = current_app.config.get('GOOGLE_SHEETS_SPREADSHEET_ID')
        except RuntimeError:
            import config
            spreadsheet_id = getattr(config, 'GOOGLE_SHEETS_SPREADSHEET_ID', None)
        
        # Get spreadsheet metadata
        metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        
        test_result['success'] = True
        test_result['message'] = 'Google Sheets connection successful'
        test_result['details'] = {
            'spreadsheet_title': metadata.get('properties', {}).get('title', 'Unknown'),
            'sheet_count': len(metadata.get('sheets', [])),
            'sheets': [sheet['properties']['title'] for sheet in metadata.get('sheets', [])]
        }
        
        # Reset error status on successful connection
        update_google_sheets_status(True)
        
    except HttpError as error:
        status_code = error.resp.status if hasattr(error.resp, 'status') else None
        test_result['message'] = f'Google Sheets API error: {error} (Status: {status_code})'
        
        if status_code == 403:
            test_result['suggestions'].extend([
                'Ensure the service account has Editor permissions for the Google Sheet',
                'Share the Google Sheet with the service account email',
                'Check if the Google Sheets API is enabled in Google Cloud Console'
            ])
        elif status_code == 404:
            test_result['suggestions'].extend([
                'Check if the SPREADSHEET_ID is correct',
                'Ensure the spreadsheet exists and is accessible'
            ])
            
    except Exception as e:
        test_result['message'] = f'Unexpected error: {str(e)}'
        test_result['suggestions'].append('Check your network connection and try again')
    
    return test_result
