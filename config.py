import os
from dotenv import load_dotenv
from datetime import timedelta

load_dotenv()

class Config:
    SECRET_KEY = os.environ.get('SECRET_KEY', 'marbo9k-secure-api-key-2024')
    SQLALCHEMY_DATABASE_URI = os.environ.get('DATABASE_URL', 'mysql+pymysql://root:@localhost/marbo9k')
    SQLALCHEMY_TRACK_MODIFICATIONS = False
    UPLOAD_FOLDER = 'static'
    MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 16MB max upload size
 
    # Session settings
    PERMANENT_SESSION_LIFETIME = timedelta(hours=24)
    
    # API settings
    API_KEY = os.environ.get('API_KEY') or 'marbo9k-secure-api-key-2024'
    
    # Google Sheets Integration
    GOOGLE_SHEETS_CREDENTIALS_FILE = os.environ.get('credentials.json') or 'credentials.json'
    GOOGLE_SHEETS_SPREADSHEET_ID = os.environ.get('1jJZMBg77rbb29-Z3vdtzV-Qkr_4Mt1qlsr3iCmQ2wcs') or '1jJZMBg77rbb29-Z3vdtzV-Qkr_4Mt1qlsr3iCmQ2wcs'
    GOOGLE_SHEETS_RANGE = os.environ.get('Products!A:Q') or 'Products!A:Q'
    
    # N8N Webhook configuration
    N8N_WEBHOOK_URL = os.environ.get('N8N_WEBHOOK_URL') or None

    
    # Application settings
    ENV = os.environ.get('FLASK_ENV') or 'production'
    DEBUG = os.environ.get('FLASK_DEBUG') == '1'

# Development configuration
class DevelopmentConfig(Config):
    DEBUG = True
    ENV = 'development'

# Production configuration
class ProductionConfig(Config):
    DEBUG = False
    ENV = 'production'

# Testing configuration
class TestingConfig(Config):
    TESTING = True
    SQLALCHEMY_DATABASE_URI = 'sqlite:///:memory:'
    WTF_CSRF_ENABLED = False

# Configuration dictionary
config = {
    'development': DevelopmentConfig,
    'production': ProductionConfig,
    'testing': TestingConfig,
    'default': DevelopmentConfig
}

