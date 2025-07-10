from app import db
import datetime

class User(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password = db.Column(db.String(200), nullable=False)
    name = db.Column(db.String(100))
    role = db.Column(db.String(20), default='staff')
    created_at = db.Column(db.DateTime, default=datetime.datetime.now)

class Customer(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    phone = db.Column(db.String(20))
    address = db.Column(db.Text)
    line_id = db.Column(db.String(50))
    created_at = db.Column(db.DateTime, default=datetime.datetime.now)
    orders = db.relationship('Order', backref='customer', lazy=True)

class Product(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    flavor = db.Column(db.String(100), nullable=False)
    description = db.Column(db.Text)
    price = db.Column(db.Float, nullable=False)
    cost = db.Column(db.Float, nullable=False)
    wholesale_price = db.Column(db.Float)
    stock = db.Column(db.Integer, default=0)
    barcode = db.Column(db.String(50))
    image_path = db.Column(db.String(200))
    created_at = db.Column(db.DateTime, default=datetime.datetime.now)
    
    @property
    def profit_margin(self):
        if self.price > 0:
            return ((self.price - self.cost) / self.price) * 100
        return 0

class Order(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    customer_id = db.Column(db.Integer, db.ForeignKey('customer.id'), nullable=False)
    order_date = db.Column(db.DateTime, default=datetime.datetime.now)
    total_amount = db.Column(db.Float, default=0)
    shipping_address = db.Column(db.Text)
    payment_slip = db.Column(db.String(200))
    payment_date = db.Column(db.DateTime)
    payment_status = db.Column(db.String(20), default='pending')
    status = db.Column(db.String(20), default='pending')
    notes = db.Column(db.Text)
    qr_code_path = db.Column(db.String(200))
    pickup_location = db.Column(db.String(500))
    order_items = db.relationship('OrderItem', backref='order', lazy=True, cascade="all, delete-orphan")
    
    @property
    def item_count(self):
        return sum(item.quantity for item in self.order_items)

class OrderItem(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    order_id = db.Column(db.Integer, db.ForeignKey('order.id'), nullable=False)
    product_id = db.Column(db.Integer, db.ForeignKey('product.id'), nullable=False)
    quantity = db.Column(db.Integer, nullable=False)
    price = db.Column(db.Float, nullable=False)
    product = db.relationship('Product')

class InventoryEntry(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    product_id = db.Column(db.Integer, db.ForeignKey('product.id'), nullable=False)
    quantity = db.Column(db.Integer, nullable=False)
    date = db.Column(db.DateTime, default=lambda: datetime.datetime.now(datetime.timezone.utc))
    notes = db.Column(db.Text)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=True)
    product = db.relationship('Product')
    user = db.relationship('User')

class ActivityLog(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'))
    action = db.Column(db.String(100), nullable=False)
    entity_type = db.Column(db.String(50))
    entity_id = db.Column(db.Integer)
    details = db.Column(db.Text)
    ip_address = db.Column(db.String(50))
    timestamp = db.Column(db.DateTime, default=datetime.datetime.now)
    user = db.relationship('User')

class Notification(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'))
    message = db.Column(db.String(255), nullable=False)
    type = db.Column(db.String(50))
    related_id = db.Column(db.Integer)
    is_read = db.Column(db.Boolean, default=False)
    created_at = db.Column(db.DateTime, default=datetime.datetime.now)
    user = db.relationship('User')

class Settings(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    company_name = db.Column(db.String(100))
    company_logo = db.Column(db.String(200))
    company_address = db.Column(db.Text)
    company_phone = db.Column(db.String(20))
    company_email = db.Column(db.String(100))
    low_stock_threshold = db.Column(db.Integer, default=10)

class Store(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False, unique=True)
    shop_settings = db.relationship('ShopSettings', backref='store', uselist=False)
    
class ShopSettings(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    store_id = db.Column(db.Integer, db.ForeignKey('store.id'), nullable=False)
    hero_title = db.Column(db.String(200), default="ยินดีต้อนรับสู่ Marbo 9k Shop")
    hero_subtitle = db.Column(db.String(200), default="เลือกสินค้าคุณภาพที่คุณชื่นชอบ ส่งตรงถึงบ้านคุณ")
    hero_background = db.Column(db.String(200))
    hero_text_color = db.Column(db.String(20), default="#ffffff")
    hero_button_text = db.Column(db.String(50))
    featured_title = db.Column(db.String(100), default="สินค้าแนะนำ")
    featured_subtitle = db.Column(db.String(200))
    primary_color = db.Column(db.String(20), default="#0066cc")
    secondary_color = db.Column(db.String(20), default="#001f3f")
    accent_color = db.Column(db.String(20), default="#ffc107")
    text_color = db.Column(db.String(20), default="#333333")
    font_family = db.Column(db.String(50), default="Kanit")
    border_radius = db.Column(db.Integer, default=8)
    footer_text = db.Column(db.String(200), default="&copy; 2025 Marbo 9k Shop. สงวนลิขสิท���ิ์ทุกประการ.")
    contact_phone = db.Column(db.String(20))
    contact_email = db.Column(db.String(100))
    social_facebook = db.Column(db.String(200))
    social_instagram = db.Column(db.String(200))
    social_line = db.Column(db.String(50))
    favicon_path = db.Column(db.String(200))
    navbar_logo_path = db.Column(db.String(200))
    updated_at = db.Column(db.DateTime, default=datetime.datetime.now, onupdate=datetime.datetime.now)

class Banner(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(100), nullable=False)
    image_path = db.Column(db.String(200), nullable=False)
    link = db.Column(db.String(200))
    position = db.Column(db.String(20), default="top")
    page_location = db.Column(db.String(50), default="shop")
    is_active = db.Column(db.Boolean, default=True)
    created_at = db.Column(db.DateTime, default=datetime.datetime.now)
    updated_at = db.Column(db.DateTime, default=datetime.datetime.now, onupdate=datetime.datetime.now)

class FeaturedProduct(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    product_id = db.Column(db.Integer, db.ForeignKey('product.id'), nullable=False)
    position = db.Column(db.Integer, default=0)
    created_at = db.Column(db.DateTime, default=datetime.datetime.now)
    product = db.relationship('Product')