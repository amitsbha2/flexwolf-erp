import sqlite3

# CONNECTION
conn = sqlite3.connect(
    "garment_factory.db",
    check_same_thread=False
)

cursor = conn.cursor()

# =========================
# MACHINE MASTER
# =========================
cursor.execute("""
CREATE TABLE IF NOT EXISTS machines (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    machine_name TEXT UNIQUE
)
""")

# =========================
# KARIGAR MASTER
# =========================
cursor.execute("""
CREATE TABLE IF NOT EXISTS karigars (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    karigar_name TEXT,
    machine_name TEXT
)
""")

# =========================
# ARTICLE MASTER
# =========================
cursor.execute("""
CREATE TABLE IF NOT EXISTS articles (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    article_name TEXT UNIQUE
)
""")

# =========================
# ARTICLE RATE MASTER
# =========================
cursor.execute("""
CREATE TABLE IF NOT EXISTS article_rates (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    article_name TEXT,
    machine_name TEXT,
    rate REAL
)
""")

# =========================
# PRODUCTION
# =========================
cursor.execute("""
CREATE TABLE IF NOT EXISTS production (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entry_date TEXT,
    karigar_name TEXT,
    machine_name TEXT,
    article_name TEXT,
    qty INTEGER,
    rate REAL,
    amount REAL,
    payment_status TEXT
)
""")

# =========================
# FABRIC STOCK
# =========================
cursor.execute("""
CREATE TABLE IF NOT EXISTS fabric_stock (
    id INTEGER PRIMARY KEY AUTOINCREMENT,

    entry_date TEXT,

    fabric_name TEXT,

    color TEXT,

    total_rolls INTEGER,

    total_weight REAL
)
""")

# =========================
# CUTTING
# =========================
cursor.execute("""
CREATE TABLE IF NOT EXISTS cutting (

    id INTEGER PRIMARY KEY AUTOINCREMENT,

    entry_date TEXT,

    fabric_name TEXT,

    color TEXT,

    lot_no TEXT,

    used_rolls INTEGER,

    s_qty INTEGER,

    m_qty INTEGER,

    l_qty INTEGER,

    xl_qty INTEGER,

    xxl_qty INTEGER,

    xxxl_qty INTEGER,

    xxxxl_qty INTEGER,

    total_pieces INTEGER
)
""")

# =========================
# BOX PACKING
# =========================
cursor.execute("""
CREATE TABLE IF NOT EXISTS box_packing (

    id INTEGER PRIMARY KEY AUTOINCREMENT,

    packing_date TEXT,

    article_name TEXT,

    color TEXT,

    box_number TEXT,

    s_qty INTEGER,

    m_qty INTEGER,

    l_qty INTEGER,

    xl_qty INTEGER,

    xxl_qty INTEGER,

    xxxl_qty INTEGER,

    xxxxl_qty INTEGER,

    total_pieces INTEGER
)
""")

# =========================
# ADVANCE PAYMENTS
# =========================
cursor.execute("""
CREATE TABLE IF NOT EXISTS advances (

    id INTEGER PRIMARY KEY AUTOINCREMENT,

    entry_date TEXT,

    karigar_name TEXT,

    amount REAL,

    remarks TEXT
)
""")

# ==========================================
# STAFF SALARY TABLE
# ==========================================

cursor.execute("""
CREATE TABLE IF NOT EXISTS staff_salary (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    entry_date TEXT,
    staff_name TEXT,
    monthly_salary REAL,
    advance REAL,
    final_salary REAL,
    payment_status TEXT
)
""")

conn.commit()

# SAVE
conn.commit()

print("DATABASE READY SUCCESSFULLY")
# =========================
# USERS TABLE
# =========================

cursor.execute("""
CREATE TABLE IF NOT EXISTS users (

    id INTEGER PRIMARY KEY AUTOINCREMENT,

    username TEXT UNIQUE,

    password TEXT,

    role TEXT
)
""")

# DEFAULT ADMIN USER

cursor.execute("""
INSERT OR IGNORE INTO users
(username, password, role)

VALUES
('admin', '1234', 'Admin')
""")

conn.commit()
# ==========================================
# STAFF ADVANCE TABLE
# ==========================================

cursor.execute("""
CREATE TABLE IF NOT EXISTS staff_advances (

    id INTEGER PRIMARY KEY AUTOINCREMENT,

    entry_date TEXT,

    staff_name TEXT,

    amount REAL,

    remarks TEXT
    status TEXT DEFAULT 'Pending'
)
""")

conn.commit()
# =========================
# STAFF MASTER TABLE
# =========================

cursor.execute("""
CREATE TABLE IF NOT EXISTS staff_master (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    staff_name TEXT UNIQUE
)
""")

conn.commit()
# ==========================================
# VENDORS TABLE
# ==========================================

cursor.execute("""
CREATE TABLE IF NOT EXISTS vendors (

    id INTEGER PRIMARY KEY AUTOINCREMENT,

    vendor_name TEXT UNIQUE,

    vendor_type TEXT
)
""")

# ==========================================
# VENDOR BILLS
# ==========================================

cursor.execute("""
CREATE TABLE IF NOT EXISTS vendor_bills (

    id INTEGER PRIMARY KEY AUTOINCREMENT,

    entry_date TEXT,

    vendor_name TEXT,

    bill_no TEXT,

    bill_amount REAL
)
""")

# ==========================================
# VENDOR PAYMENTS
# ==========================================

cursor.execute("""
CREATE TABLE IF NOT EXISTS vendor_payments (

    id INTEGER PRIMARY KEY AUTOINCREMENT,

    entry_date TEXT,

    vendor_name TEXT,

    payment_amount REAL,

    payment_mode TEXT,

    remarks TEXT
)
""")

conn.commit()
