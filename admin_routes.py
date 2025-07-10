"""
Admin routes for Google Sheets management
Add these routes to your main Flask application
"""

from flask import Blueprint, render_template, request, redirect, url_for, flash, jsonify, session
from app.utils import (
    test_google_sheets_connection,
    get_google_sheets_status,
    reset_google_sheets_integration,
    sync_all_products_to_google_sheets,
    setup_google_sheets_structure
)

admin_bp = Blueprint('admin', __name__, url_prefix='/admin')

def require_admin():
    """Decorator to require admin access"""
    if 'user_id' not in session or session.get('role') != 'admin':
        flash('Access denied. Admin privileges required.', 'danger')
        return redirect(url_for('main.login'))
    return None

@admin_bp.route('/google_sheets')
def google_sheets_management():
    """Google Sheets management dashboard"""
    auth_check = require_admin()
    if auth_check:
        return auth_check
    
    # Get current status
    status = get_google_sheets_status()
    
    # Test connection
    test_result = test_google_sheets_connection()
    
    return render_template('admin/google_sheets.html', 
                         status=status, 
                         test_result=test_result)

@admin_bp.route('/google_sheets/test', methods=['POST'])
def test_google_sheets():
    """Test Google Sheets connection"""
    auth_check = require_admin()
    if auth_check:
        return auth_check
    
    test_result = test_google_sheets_connection()
    
    if test_result['success']:
        flash('Google Sheets connection test successful!', 'success')
    else:
        flash(f'Google Sheets connection test failed: {test_result["message"]}', 'danger')
    
    return redirect(url_for('admin.google_sheets_management'))

@admin_bp.route('/google_sheets/reset', methods=['POST'])
def reset_google_sheets():
    """Reset Google Sheets integration status"""
    auth_check = require_admin()
    if auth_check:
        return auth_check
    
    reset_google_sheets_integration()
    flash('Google Sheets integration status has been reset.', 'info')
    
    return redirect(url_for('admin.google_sheets_management'))

@admin_bp.route('/google_sheets/sync', methods=['POST'])
def sync_products():
    """Manually sync all products to Google Sheets"""
    auth_check = require_admin()
    if auth_check:
        return auth_check
    
    success = sync_all_products_to_google_sheets()
    
    if success:
        flash('All products have been synced to Google Sheets successfully!', 'success')
    else:
        flash('Failed to sync products to Google Sheets. Check the logs for details.', 'danger')
    
    return redirect(url_for('admin.google_sheets_management'))

@admin_bp.route('/google_sheets/setup', methods=['POST'])
def setup_sheets_structure():
    """Setup Google Sheets structure"""
    auth_check = require_admin()
    if auth_check:
        return auth_check
    
    success = setup_google_sheets_structure()
    
    if success:
        flash('Google Sheets structure has been set up successfully!', 'success')
    else:
        flash('Failed to set up Google Sheets structure. Check the logs for details.', 'danger')
    
    return redirect(url_for('admin.google_sheets_management'))

@admin_bp.route('/google_sheets/status')
def google_sheets_status_api():
    """API endpoint for Google Sheets status"""
    auth_check = require_admin()
    if auth_check:
        return jsonify({'error': 'Unauthorized'}), 401
    
    status = get_google_sheets_status()
    test_result = test_google_sheets_connection()
    
    return jsonify({
        'status': status,
        'test_result': test_result
    })
