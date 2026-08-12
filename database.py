import sqlite3
import hashlib
import os
from datetime import datetime

from constants import (
    ASSIGNMENT_FULL_SERVICE, ASSIGNMENT_PARTIAL, ASSIGNMENT_UNPLANNED,
    MATERIAL_STATUS_DONE, MATERIAL_STATUS_PENDING, MATERIAL_STATUS_NA,
    SIZES, RAPORLU_LASTIK,
)

DB_PATH = os.path.join(os.path.dirname(__file__), 'app_data.db')


def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


def _add_column_if_missing(conn, table, column, coltype_default):
    cols = [r['name'] for r in conn.execute(f"PRAGMA table_info({table})").fetchall()]
    if column not in cols:
        conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {coltype_default}")


def _rename_column_if_exists(conn, table, old, new):
    cols = [r['name'] for r in conn.execute(f"PRAGMA table_info({table})").fetchall()]
    if old in cols and new not in cols:
        try:
            conn.execute(f"ALTER TABLE {table} RENAME COLUMN {old} TO {new}")
        except sqlite3.OperationalError:
            pass


def init_db():
    conn = get_conn()
    c = conn.cursor()

    # Users
    c.execute('''CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        full_name TEXT NOT NULL,
        role TEXT NOT NULL DEFAULT 'user',
        created_at TEXT NOT NULL
    )''')

    # Manufacturers
    c.execute('''CREATE TABLE IF NOT EXISTS manufacturers (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE NOT NULL,
        contact_person TEXT DEFAULT '',
        phone TEXT DEFAULT '',
        email TEXT DEFAULT '',
        created_at TEXT NOT NULL
    )''')

    # Ürün grupları (Kadın: string, cheeky... / Erkek: boxer, slip...)
    c.execute('''CREATE TABLE IF NOT EXISTS product_groups (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        gender TEXT NOT NULL,
        name TEXT NOT NULL,
        created_at TEXT NOT NULL,
        UNIQUE(gender, name)
    )''')

    # Ürün adları (belirli model isimleri; cinsiyet + ürün grubunu otomatik taşır)
    c.execute('''CREATE TABLE IF NOT EXISTS product_names (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE NOT NULL,
        gender TEXT NOT NULL,
        urun_grubu TEXT NOT NULL,
        created_at TEXT NOT NULL
    )''')

    # Kumaş ana verisi
    c.execute('''CREATE TABLE IF NOT EXISTS fabrics (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        icerik TEXT DEFAULT '',
        kumas_turu TEXT DEFAULT 'Örme',
        en REAL DEFAULT 0,
        gr_m2 REAL DEFAULT 0,
        urun_adi_id INTEGER,
        created_at TEXT NOT NULL,
        FOREIGN KEY (urun_adi_id) REFERENCES product_names(id)
    )''')

    # Lastik ana verisi (Baskılı Lastik / Jakarlı Lastik / Raporlu Lastik)
    c.execute('''CREATE TABLE IF NOT EXISTS elastics (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tur TEXT NOT NULL,
        ad TEXT NOT NULL,
        boyut TEXT DEFAULT '',
        urun_adi_id INTEGER,
        created_at TEXT NOT NULL,
        FOREIGN KEY (urun_adi_id) REFERENCES product_names(id)
    )''')

    # Uygulama ayarları
    c.execute('''CREATE TABLE IF NOT EXISTS app_settings (
        id INTEGER PRIMARY KEY CHECK (id = 1),
        firma_adi TEXT DEFAULT 'Tekstil Ltd'
    )''')

    # İrsaliyeler (gelen mal kayıtları) - item_id, kategoriye göre order_fabrics / order_elastics /
    # order_accessories / (Kutu için NULL, order_id yeterli) tablosundaki satırı işaret eder.
    c.execute('''CREATE TABLE IF NOT EXISTS irsaliyeler (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        order_id INTEGER NOT NULL,
        item_id INTEGER,
        kategori TEXT NOT NULL,
        irsaliye_no TEXT DEFAULT '',
        tedarikci TEXT DEFAULT '',
        miktar REAL DEFAULT 0,
        birim TEXT DEFAULT '',
        gelis_tarihi TEXT DEFAULT '',
        aciklama TEXT DEFAULT '',
        created_at TEXT NOT NULL,
        FOREIGN KEY (order_id) REFERENCES orders(id)
    )''')

    # Reçete kütüphanesi: cinsiyet + ürün grubu + beden -> kumaş(gr)/lastik(adet, cm) reçetesi
    c.execute('''CREATE TABLE IF NOT EXISTS recipes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        gender TEXT NOT NULL,
        urun_grubu TEXT NOT NULL,
        size TEXT NOT NULL,
        kumas_gr REAL DEFAULT 0,
        lastik_adet REAL DEFAULT 0,
        lastik_cm REAL DEFAULT 0,
        created_at TEXT NOT NULL,
        UNIQUE(gender, urun_grubu, size)
    )''')

    # Orders (ana sipariş kaydı)
    c.execute('''CREATE TABLE IF NOT EXISTS orders (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        model_name TEXT NOT NULL,
        gender TEXT NOT NULL DEFAULT 'Kadın',
        urun_grubu TEXT DEFAULT '',
        manufacturer_id INTEGER,
        assignment_type TEXT NOT NULL DEFAULT 'planlama',
        package_size INTEGER NOT NULL DEFAULT 1,
        kumas_fire_orani REAL NOT NULL DEFAULT 3,
        lastik_fire_orani REAL NOT NULL DEFAULT 3,
        deadline TEXT,
        siparis_tarihi TEXT DEFAULT '',
        total_quantity INTEGER DEFAULT 0,
        total_boxes INTEGER DEFAULT 0,
        status TEXT DEFAULT 'Planlama',
        urun_foto TEXT DEFAULT '',

        kutu_manufacturer_id INTEGER,
        kutu_siparis_adet REAL DEFAULT 0,
        kutu_gelen_adet REAL DEFAULT 0,
        kutu_siparis_tarihi TEXT DEFAULT '',
        kutu_termin_tarihi TEXT DEFAULT '',

        created_at TEXT NOT NULL,
        created_by INTEGER,
        FOREIGN KEY (manufacturer_id) REFERENCES manufacturers(id),
        FOREIGN KEY (kutu_manufacturer_id) REFERENCES manufacturers(id),
        FOREIGN KEY (created_by) REFERENCES users(id)
    )''')

    # Beden bazlı kutu adedi + reçete (siparişe kaydedilen anlık kopya)
    c.execute('''CREATE TABLE IF NOT EXISTS order_sizes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        order_id INTEGER NOT NULL,
        size TEXT NOT NULL,
        box_qty INTEGER NOT NULL DEFAULT 0,
        kumas_gr REAL NOT NULL DEFAULT 0,
        lastik_adet REAL NOT NULL DEFAULT 0,
        lastik_cm REAL NOT NULL DEFAULT 0,
        FOREIGN KEY (order_id) REFERENCES orders(id)
    )''')

    # --- Kumaş / Lastik / Aksesuar artık BİRBİRİNDEN BAĞIMSIZ satır listeleri ---
    c.execute('''CREATE TABLE IF NOT EXISTS order_fabrics (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        order_id INTEGER NOT NULL,
        seq INTEGER NOT NULL DEFAULT 1,
        kumas_fabric_id INTEGER,
        kumas_renk TEXT DEFAULT '',
        kumas_foto TEXT DEFAULT '',
        kumas_manufacturer_id INTEGER,
        kumas_siparis_kg REAL DEFAULT 0,
        kumas_gelen_kg REAL DEFAULT 0,
        kumas_siparis_tarihi TEXT DEFAULT '',
        kumas_termin_tarihi TEXT DEFAULT '',
        FOREIGN KEY (order_id) REFERENCES orders(id),
        FOREIGN KEY (kumas_fabric_id) REFERENCES fabrics(id),
        FOREIGN KEY (kumas_manufacturer_id) REFERENCES manufacturers(id)
    )''')

    c.execute('''CREATE TABLE IF NOT EXISTS order_elastics (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        order_id INTEGER NOT NULL,
        seq INTEGER NOT NULL DEFAULT 1,
        lastik_elastic_id INTEGER,
        lastik_renk TEXT DEFAULT '',
        lastik_foto TEXT DEFAULT '',
        lastik_manufacturer_id INTEGER,
        lastik_siparis_cm REAL DEFAULT 0,
        lastik_gelen_cm REAL DEFAULT 0,
        lastik_siparis_adet REAL DEFAULT 0,
        lastik_gelen_adet REAL DEFAULT 0,
        lastik_siparis_tarihi TEXT DEFAULT '',
        lastik_termin_tarihi TEXT DEFAULT '',
        FOREIGN KEY (order_id) REFERENCES orders(id),
        FOREIGN KEY (lastik_elastic_id) REFERENCES elastics(id),
        FOREIGN KEY (lastik_manufacturer_id) REFERENCES manufacturers(id)
    )''')

    c.execute('''CREATE TABLE IF NOT EXISTS order_accessories (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        order_id INTEGER NOT NULL,
        seq INTEGER NOT NULL DEFAULT 1,
        aksesuar_adi TEXT DEFAULT '',
        aksesuar_renk TEXT DEFAULT '',
        aksesuar_foto TEXT DEFAULT '',
        aksesuar_manufacturer_id INTEGER,
        aksesuar_siparis_adet REAL DEFAULT 0,
        aksesuar_gelen_adet REAL DEFAULT 0,
        aksesuar_siparis_tarihi TEXT DEFAULT '',
        aksesuar_termin_tarihi TEXT DEFAULT '',
        FOREIGN KEY (order_id) REFERENCES orders(id),
        FOREIGN KEY (aksesuar_manufacturer_id) REFERENCES manufacturers(id)
    )''')

    # Logs
    c.execute('''CREATE TABLE IF NOT EXISTS logs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER,
        username TEXT,
        action TEXT NOT NULL,
        details TEXT DEFAULT '',
        created_at TEXT NOT NULL,
        FOREIGN KEY (user_id) REFERENCES users(id)
    )''')

    conn.commit()

    # --- Eski şemadan (varsa) rename/migrasyon ---
    _rename_column_if_exists(conn, 'recipes', 'lastik_mt', 'lastik_cm')
    _rename_column_if_exists(conn, 'order_sizes', 'lastik_mt', 'lastik_cm')
    conn.commit()

    migrations = [
        ('manufacturers', 'contact_person', "TEXT DEFAULT ''"),
        ('manufacturers', 'phone', "TEXT DEFAULT ''"),
        ('manufacturers', 'email', "TEXT DEFAULT ''"),
        ('orders', 'urun_grubu', "TEXT DEFAULT ''"),
        ('orders', 'siparis_tarihi', "TEXT DEFAULT ''"),
        ('orders', 'kutu_manufacturer_id', "INTEGER"),
        ('orders', 'kutu_siparis_adet', "REAL DEFAULT 0"),
        ('orders', 'kutu_gelen_adet', "REAL DEFAULT 0"),
        ('orders', 'kutu_siparis_tarihi', "TEXT DEFAULT ''"),
        ('orders', 'kutu_termin_tarihi', "TEXT DEFAULT ''"),
        ('fabrics', 'kumas_turu', "TEXT DEFAULT 'Örme'"),
        ('fabrics', 'urun_adi_id', "INTEGER"),
        ('elastics', 'urun_adi_id', "INTEGER"),
    ]
    for table, col, definition in migrations:
        try:
            _add_column_if_missing(conn, table, col, definition)
        except sqlite3.OperationalError:
            pass
    conn.commit()

    # Varsayılan ayarlar satırı
    c.execute("SELECT COUNT(*) FROM app_settings WHERE id = 1")
    if c.fetchone()[0] == 0:
        c.execute("INSERT INTO app_settings (id, firma_adi) VALUES (1, 'Paul Kenzie')")
        conn.commit()

    # Varsayılan admin kullanıcı
    c.execute("SELECT COUNT(*) FROM users")
    if c.fetchone()[0] == 0:
        pw_hash = hash_password("admin123")
        c.execute('''INSERT INTO users (username, password_hash, full_name, role, created_at)
                     VALUES (?, ?, ?, ?, ?)''',
                  ('admin', pw_hash, 'Yönetici', 'admin', datetime.now().isoformat()))
        conn.commit()

    conn.close()


# --- Auth ---
def hash_password(password):
    salt = "paulkenzie_erp"
    return hashlib.sha256((salt + password).encode()).hexdigest()


def verify_password(password, password_hash):
    return hash_password(password) == password_hash


def authenticate(username, password):
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT * FROM users WHERE username = ?", (username,))
    user = c.fetchone()
    conn.close()
    if user and verify_password(password, user['password_hash']):
        return dict(user)
    return None


def add_log(user_id, username, action, details=''):
    conn = get_conn()
    conn.execute('''INSERT INTO logs (user_id, username, action, details, created_at)
                    VALUES (?, ?, ?, ?, ?)''',
                 (user_id, username, action, details, datetime.now().isoformat()))
    conn.commit()
    conn.close()


def get_logs(limit=200):
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT * FROM logs ORDER BY created_at DESC LIMIT ?", (limit,))
    rows = [dict(r) for r in c.fetchall()]
    conn.close()
    return rows


# --- Ayarlar (firma adı) ---
def get_settings():
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT * FROM app_settings WHERE id = 1")
    row = c.fetchone()
    conn.close()
    return dict(row) if row else {'firma_adi': 'Tekstil Ltd'}


def update_settings(firma_adi):
    conn = get_conn()
    conn.execute("UPDATE app_settings SET firma_adi=? WHERE id=1", (firma_adi,))
    conn.commit()
    conn.close()


# --- Manufacturers ---
def get_manufacturers():
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT * FROM manufacturers ORDER BY name")
    rows = [dict(r) for r in c.fetchall()]
    conn.close()
    return rows


def add_manufacturer(name, contact_person='', phone='', email=''):
    conn = get_conn()
    try:
        conn.execute('''INSERT INTO manufacturers (name, contact_person, phone, email, created_at)
                        VALUES (?, ?, ?, ?, ?)''',
                     (name, contact_person, phone, email, datetime.now().isoformat()))
        conn.commit()
        return True
    except sqlite3.IntegrityError:
        return False
    finally:
        conn.close()


def update_manufacturer(mf_id, contact_person, phone, email):
    conn = get_conn()
    conn.execute('''UPDATE manufacturers SET contact_person = ?, phone = ?, email = ? WHERE id = ?''',
                 (contact_person, phone, email, mf_id))
    conn.commit()
    conn.close()


def delete_manufacturer(mf_id):
    conn = get_conn()
    conn.execute("DELETE FROM manufacturers WHERE id = ?", (mf_id,))
    conn.commit()
    conn.close()


# --- Product groups (ürün grupları) ---
def get_product_groups(gender=None):
    conn = get_conn()
    c = conn.cursor()
    if gender:
        c.execute("SELECT * FROM product_groups WHERE gender = ? ORDER BY name", (gender,))
    else:
        c.execute("SELECT * FROM product_groups ORDER BY gender, name")
    rows = [dict(r) for r in c.fetchall()]
    conn.close()
    return rows


def add_product_group(gender, name):
    conn = get_conn()
    try:
        conn.execute("INSERT INTO product_groups (gender, name, created_at) VALUES (?, ?, ?)",
                     (gender, name, datetime.now().isoformat()))
        conn.commit()
        return True
    except sqlite3.IntegrityError:
        return False
    finally:
        conn.close()


def delete_product_group(pg_id):
    conn = get_conn()
    conn.execute("DELETE FROM product_groups WHERE id = ?", (pg_id,))
    conn.commit()
    conn.close()


# --- Product names (ürün adları: model ismi + cinsiyet + ürün grubu) ---
def get_product_names(gender=None):
    conn = get_conn()
    c = conn.cursor()
    if gender:
        c.execute("SELECT * FROM product_names WHERE gender = ? ORDER BY name", (gender,))
    else:
        c.execute("SELECT * FROM product_names ORDER BY name")
    rows = [dict(r) for r in c.fetchall()]
    conn.close()
    return rows


def get_product_name_by_name(name):
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT * FROM product_names WHERE name = ?", (name,))
    row = c.fetchone()
    conn.close()
    return dict(row) if row else None


def add_product_name(name, gender, urun_grubu):
    conn = get_conn()
    try:
        conn.execute("INSERT INTO product_names (name, gender, urun_grubu, created_at) VALUES (?, ?, ?, ?)",
                     (name, gender, urun_grubu, datetime.now().isoformat()))
        conn.commit()
        return True
    except sqlite3.IntegrityError:
        return False
    finally:
        conn.close()


def update_product_name(pn_id, name, gender, urun_grubu):
    conn = get_conn()
    try:
        conn.execute("UPDATE product_names SET name=?, gender=?, urun_grubu=? WHERE id=?",
                     (name, gender, urun_grubu, pn_id))
        conn.commit()
        return True
    except sqlite3.IntegrityError:
        return False
    finally:
        conn.close()


def delete_product_name(pn_id):
    conn = get_conn()
    conn.execute("DELETE FROM product_names WHERE id = ?", (pn_id,))
    conn.commit()
    conn.close()


# --- Fabrics (kumaş ana verisi) ---
def get_fabrics():
    conn = get_conn()
    c = conn.cursor()
    c.execute('''SELECT f.*, p.name as urun_adi FROM fabrics f
                 LEFT JOIN product_names p ON f.urun_adi_id = p.id
                 ORDER BY f.name''')
    rows = [dict(r) for r in c.fetchall()]
    conn.close()
    return rows


def add_fabric(name, icerik, kumas_turu, en, gr_m2, urun_adi_id=None):
    conn = get_conn()
    conn.execute('''INSERT INTO fabrics (name, icerik, kumas_turu, en, gr_m2, urun_adi_id, created_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?)''',
                 (name, icerik, kumas_turu, en, gr_m2, urun_adi_id, datetime.now().isoformat()))
    conn.commit()
    conn.close()


def update_fabric(fabric_id, name, icerik, kumas_turu, en, gr_m2, urun_adi_id=None):
    conn = get_conn()
    conn.execute('''UPDATE fabrics SET name=?, icerik=?, kumas_turu=?, en=?, gr_m2=?, urun_adi_id=?
                    WHERE id=?''', (name, icerik, kumas_turu, en, gr_m2, urun_adi_id, fabric_id))
    conn.commit()
    conn.close()


def delete_fabric(fabric_id):
    conn = get_conn()
    conn.execute("DELETE FROM fabrics WHERE id = ?", (fabric_id,))
    conn.commit()
    conn.close()


# --- Elastics (lastik ana verisi) ---
def get_elastics():
    conn = get_conn()
    c = conn.cursor()
    c.execute('''SELECT e.*, p.name as urun_adi FROM elastics e
                 LEFT JOIN product_names p ON e.urun_adi_id = p.id
                 ORDER BY e.tur, e.boyut''')
    rows = [dict(r) for r in c.fetchall()]
    conn.close()
    return rows


def add_elastic(tur, boyut, urun_adi_id=None):
    conn = get_conn()
    conn.execute('''INSERT INTO elastics (tur, ad, boyut, urun_adi_id, created_at)
                    VALUES (?, ?, ?, ?, ?)''',
                 (tur, tur, boyut, urun_adi_id, datetime.now().isoformat()))
    conn.commit()
    conn.close()


def update_elastic(elastic_id, tur, boyut, urun_adi_id=None):
    conn = get_conn()
    conn.execute("UPDATE elastics SET tur=?, ad=?, boyut=?, urun_adi_id=? WHERE id=?",
                (tur, tur, boyut, urun_adi_id, elastic_id))
    conn.commit()
    conn.close()


def delete_elastic(elastic_id):
    conn = get_conn()
    conn.execute("DELETE FROM elastics WHERE id = ?", (elastic_id,))
    conn.commit()
    conn.close()


# --- Recipes (reçete kütüphanesi) ---
def get_recipe_groups():
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT DISTINCT gender, urun_grubu FROM recipes ORDER BY gender, urun_grubu")
    rows = [dict(r) for r in c.fetchall()]
    conn.close()
    return rows


def get_recipe(gender, urun_grubu):
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT * FROM recipes WHERE gender = ? AND urun_grubu = ?", (gender, urun_grubu))
    rows = [dict(r) for r in c.fetchall()]
    conn.close()
    result = {s: {'kumas_gr': 0, 'lastik_adet': 0, 'lastik_cm': 0} for s in SIZES}
    for r in rows:
        result[r['size']] = {'kumas_gr': r['kumas_gr'], 'lastik_adet': r['lastik_adet'], 'lastik_cm': r['lastik_cm']}
    return result


def upsert_recipe(gender, urun_grubu, size, kumas_gr, lastik_adet, lastik_cm):
    conn = get_conn()
    conn.execute('''INSERT INTO recipes (gender, urun_grubu, size, kumas_gr, lastik_adet, lastik_cm, created_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(gender, urun_grubu, size)
                    DO UPDATE SET kumas_gr = excluded.kumas_gr, lastik_adet = excluded.lastik_adet,
                                  lastik_cm = excluded.lastik_cm''',
                 (gender, urun_grubu, size, kumas_gr, lastik_adet, lastik_cm, datetime.now().isoformat()))
    conn.commit()
    conn.close()


def delete_recipe_group(gender, urun_grubu):
    conn = get_conn()
    conn.execute("DELETE FROM recipes WHERE gender = ? AND urun_grubu = ?", (gender, urun_grubu))
    conn.commit()
    conn.close()


# --- Users ---
def get_users():
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT id, username, full_name, role, created_at FROM users ORDER BY username")
    rows = [dict(r) for r in c.fetchall()]
    conn.close()
    return rows


def add_user(username, password, full_name, role='user'):
    conn = get_conn()
    try:
        pw_hash = hash_password(password)
        conn.execute('''INSERT INTO users (username, password_hash, full_name, role, created_at)
                        VALUES (?, ?, ?, ?, ?)''',
                     (username, pw_hash, full_name, role, datetime.now().isoformat()))
        conn.commit()
        return True
    except sqlite3.IntegrityError:
        return False
    finally:
        conn.close()


def update_user_password(user_id, new_password):
    conn = get_conn()
    pw_hash = hash_password(new_password)
    conn.execute("UPDATE users SET password_hash = ? WHERE id = ?", (pw_hash, user_id))
    conn.commit()
    conn.close()


def update_user_role(user_id, role):
    conn = get_conn()
    conn.execute("UPDATE users SET role = ? WHERE id = ?", (role, user_id))
    conn.commit()
    conn.close()


def delete_user(user_id):
    conn = get_conn()
    conn.execute("DELETE FROM users WHERE id = ?", (user_id,))
    conn.commit()
    conn.close()


# --- Orders ---
def add_order(model_name, gender, urun_grubu, package_size,
              kumas_fire_orani, lastik_fire_orani, deadline, siparis_tarihi, sizes,
              fabrics, elastics, accessories, user_id, urun_foto=''):
    """
    Yeni sipariş her zaman 'Planlama Bekleyen' (manufacturer_id=None) olarak oluşturulur.

    sizes: [{'size': 'M', 'box_qty': 10, 'kumas_gr': 150, 'lastik_adet': 2, 'lastik_cm': 50}, ...]
    fabrics: [{'kumas_fabric_id':.., 'kumas_renk':.., 'kumas_foto':..}, ...]  (en az 1 satır)
    elastics: [{'lastik_elastic_id':.., 'lastik_renk':.., 'lastik_foto':..}, ...]  (en az 1 satır)
    accessories: [{'aksesuar_adi':.., 'aksesuar_renk':.., 'aksesuar_foto':..}, ...]  (0 veya daha fazla)
    """
    conn = get_conn()
    c = conn.cursor()

    total_boxes = sum(s['box_qty'] for s in sizes)
    total_quantity = total_boxes * package_size

    c.execute('''INSERT INTO orders
                 (model_name, gender, urun_grubu, manufacturer_id, assignment_type, package_size,
                  kumas_fire_orani, lastik_fire_orani, deadline, siparis_tarihi,
                  total_quantity, total_boxes, status, urun_foto, created_at, created_by)
                 VALUES (?, ?, ?, NULL, ?, ?, ?, ?, ?, ?, ?, ?, 'Planlama', ?, ?, ?)''',
              (model_name, gender, urun_grubu, ASSIGNMENT_UNPLANNED, package_size,
               kumas_fire_orani, lastik_fire_orani, deadline, siparis_tarihi,
               total_quantity, total_boxes, urun_foto, datetime.now().isoformat(), user_id))
    order_id = c.lastrowid

    for s in sizes:
        if s['box_qty'] > 0:
            c.execute('''INSERT INTO order_sizes (order_id, size, box_qty, kumas_gr, lastik_adet, lastik_cm)
                         VALUES (?, ?, ?, ?, ?, ?)''',
                      (order_id, s['size'], s['box_qty'], s.get('kumas_gr', 0),
                       s.get('lastik_adet', 0), s.get('lastik_cm', 0)))

    fabric_ids, elastic_ids, accessory_ids = [], [], []

    for i, f in enumerate(fabrics or [{}], start=1):
        c.execute('''INSERT INTO order_fabrics (order_id, seq, kumas_fabric_id, kumas_renk, kumas_foto)
                     VALUES (?, ?, ?, ?, ?)''',
                  (order_id, i, f.get('kumas_fabric_id'), f.get('kumas_renk', ''), f.get('kumas_foto', '')))
        fabric_ids.append(c.lastrowid)

    for i, el in enumerate(elastics or [{}], start=1):
        c.execute('''INSERT INTO order_elastics (order_id, seq, lastik_elastic_id, lastik_renk, lastik_foto)
                     VALUES (?, ?, ?, ?, ?)''',
                  (order_id, i, el.get('lastik_elastic_id'), el.get('lastik_renk', ''), el.get('lastik_foto', '')))
        elastic_ids.append(c.lastrowid)

    for i, a in enumerate(accessories or [], start=1):
        if not (a.get('aksesuar_adi') or '').strip():
            continue
        c.execute('''INSERT INTO order_accessories (order_id, seq, aksesuar_adi, aksesuar_renk, aksesuar_foto)
                     VALUES (?, ?, ?, ?, ?)''',
                  (order_id, i, a.get('aksesuar_adi', ''), a.get('aksesuar_renk', ''), a.get('aksesuar_foto', '')))
        accessory_ids.append(c.lastrowid)

    conn.commit()
    conn.close()
    return order_id, fabric_ids, elastic_ids, accessory_ids


def set_order_photo(order_id, path):
    conn = get_conn()
    conn.execute("UPDATE orders SET urun_foto = ? WHERE id = ?", (path, order_id))
    conn.commit()
    conn.close()


def set_fabric_photo(fabric_row_id, path):
    conn = get_conn()
    conn.execute("UPDATE order_fabrics SET kumas_foto = ? WHERE id = ?", (path, fabric_row_id))
    conn.commit()
    conn.close()


def set_elastic_photo(elastic_row_id, path):
    conn = get_conn()
    conn.execute("UPDATE order_elastics SET lastik_foto = ? WHERE id = ?", (path, elastic_row_id))
    conn.commit()
    conn.close()


def set_accessory_photo(accessory_row_id, path):
    conn = get_conn()
    conn.execute("UPDATE order_accessories SET aksesuar_foto = ? WHERE id = ?", (path, accessory_row_id))
    conn.commit()
    conn.close()


def _material_status(material_manufacturer_id, order_manufacturer_id):
    if order_manufacturer_id and material_manufacturer_id == order_manufacturer_id:
        return MATERIAL_STATUS_DONE
    return MATERIAL_STATUS_PENDING


def _compute_material_totals(order, sizes):
    """Sipariş geneli toplam kumaş(kg)/lastik(adet,cm) ihtiyacı - satır sayısından bağımsız."""
    kumas_fire_mult = 1 + (order['kumas_fire_orani'] or 0) / 100.0
    lastik_fire_mult = 1 + (order['lastik_fire_orani'] or 0) / 100.0
    kumas_gr_total = sum(s['box_qty'] * s['kumas_gr'] for s in sizes) * kumas_fire_mult
    lastik_adet_total = sum(s['box_qty'] * s['lastik_adet'] for s in sizes) * lastik_fire_mult
    lastik_cm_total = sum(s['box_qty'] * s['lastik_cm'] for s in sizes) * lastik_fire_mult
    return {
        'kumas_kg_total': round(kumas_gr_total / 1000.0, 3),
        'lastik_adet_total': round(lastik_adet_total, 1),
        'lastik_cm_total': round(lastik_cm_total, 1),
    }


def get_order_detail(order_id):
    conn = get_conn()
    c = conn.cursor()
    c.execute('''SELECT o.*, m.name as manufacturer_name, u.full_name as created_by_name,
                        km.name as kutu_manufacturer_name
                 FROM orders o
                 LEFT JOIN manufacturers m ON o.manufacturer_id = m.id
                 LEFT JOIN manufacturers km ON o.kutu_manufacturer_id = km.id
                 LEFT JOIN users u ON o.created_by = u.id
                 WHERE o.id = ?''', (order_id,))
    order = c.fetchone()
    if not order:
        conn.close()
        return None
    order = dict(order)

    c.execute("SELECT * FROM order_sizes WHERE order_id = ? ORDER BY id", (order_id,))
    sizes = [dict(r) for r in c.fetchall()]
    order['sizes'] = sizes

    totals = _compute_material_totals(order, sizes)
    order.update(totals)

    c.execute('''SELECT ofa.*, f.name as kumas_adi, f.icerik as kumas_icerik, f.kumas_turu as kumas_turu,
                        m.name as kumas_manufacturer_name
                 FROM order_fabrics ofa
                 LEFT JOIN fabrics f ON ofa.kumas_fabric_id = f.id
                 LEFT JOIN manufacturers m ON ofa.kumas_manufacturer_id = m.id
                 WHERE ofa.order_id = ? ORDER BY ofa.seq''', (order_id,))
    fabric_rows = [dict(r) for r in c.fetchall()]
    n_fab = len(fabric_rows) or 1
    for f in fabric_rows:
        f['kumas_kg_required'] = round(totals['kumas_kg_total'] / n_fab, 3)
        f['kumas_kalan_kg'] = round(f['kumas_kg_required'] - (f.get('kumas_gelen_kg') or 0), 3)
        f['kumas_status'] = _material_status(f.get('kumas_manufacturer_id'), order['manufacturer_id'])
    order['fabrics'] = fabric_rows

    c.execute('''SELECT oel.*, e.tur as lastik_tur, e.boyut as lastik_boyut,
                        m.name as lastik_manufacturer_name
                 FROM order_elastics oel
                 LEFT JOIN elastics e ON oel.lastik_elastic_id = e.id
                 LEFT JOIN manufacturers m ON oel.lastik_manufacturer_id = m.id
                 WHERE oel.order_id = ? ORDER BY oel.seq''', (order_id,))
    elastic_rows = [dict(r) for r in c.fetchall()]
    n_el = len(elastic_rows) or 1
    for e in elastic_rows:
        is_raporlu = (e.get('lastik_tur') == RAPORLU_LASTIK)
        e['is_raporlu'] = is_raporlu
        e['lastik_cm_required'] = round(totals['lastik_cm_total'] / n_el, 1)
        e['lastik_adet_required'] = round(totals['lastik_adet_total'] / n_el, 1)
        e['lastik_cm_kalan'] = round(e['lastik_cm_required'] - (e.get('lastik_gelen_cm') or 0), 1)
        e['lastik_adet_kalan'] = round(e['lastik_adet_required'] - (e.get('lastik_gelen_adet') or 0), 1)
        e['lastik_status'] = _material_status(e.get('lastik_manufacturer_id'), order['manufacturer_id'])
    order['elastics'] = elastic_rows

    c.execute('''SELECT oac.*, m.name as aksesuar_manufacturer_name
                 FROM order_accessories oac
                 LEFT JOIN manufacturers m ON oac.aksesuar_manufacturer_id = m.id
                 WHERE oac.order_id = ? ORDER BY oac.seq''', (order_id,))
    accessory_rows = [dict(r) for r in c.fetchall()]
    n_ak = len(accessory_rows) or 1
    for a in accessory_rows:
        a['aksesuar_adet_required'] = round((order['total_quantity'] or 0) / n_ak, 1)
        a['aksesuar_kalan_adet'] = round(a['aksesuar_adet_required'] - (a.get('aksesuar_gelen_adet') or 0), 1)
        a['aksesuar_status'] = _material_status(a.get('aksesuar_manufacturer_id'), order['manufacturer_id'])
    order['accessories'] = accessory_rows

    conn.close()
    return order


def get_orders_overview(assigned=None, include_completed=False, year=None):
    """assigned=True -> üretici atanmış, False -> planlama bekleyen, None -> hepsi"""
    conn = get_conn()
    c = conn.cursor()
    query = '''SELECT o.*, m.name as manufacturer_name
               FROM orders o
               LEFT JOIN manufacturers m ON o.manufacturer_id = m.id
               WHERE 1=1'''
    params = []
    if not include_completed:
        query += " AND o.status NOT IN ('Tamamlandı','İptal')"
    if assigned is True:
        query += " AND o.manufacturer_id IS NOT NULL"
    elif assigned is False:
        query += " AND o.manufacturer_id IS NULL"
    if year:
        query += " AND strftime('%Y', o.created_at) = ?"
        params.append(str(year))
    query += " ORDER BY (o.deadline IS NULL), o.deadline ASC"
    c.execute(query, params)
    orders = [dict(r) for r in c.fetchall()]
    conn.close()

    return [get_order_detail(o['id']) for o in orders]


def get_dashboard_counts():
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT COUNT(*) FROM orders")
    total = c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM orders WHERE status NOT IN ('Tamamlandı','İptal')")
    active = c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM orders WHERE status NOT IN ('Tamamlandı','İptal') AND manufacturer_id IS NOT NULL")
    assigned = c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM orders WHERE status NOT IN ('Tamamlandı','İptal') AND manufacturer_id IS NULL")
    unplanned = c.fetchone()[0]
    today = datetime.now().date().isoformat()
    c.execute("SELECT COUNT(*) FROM orders WHERE status NOT IN ('Tamamlandı','İptal') AND deadline IS NOT NULL AND deadline <= ?",
              (today,))
    urgent = c.fetchone()[0]
    conn.close()
    return {'total': total, 'active': active, 'assigned': assigned, 'unplanned': unplanned, 'urgent': urgent}


def get_completed_count_by_year(year):
    conn = get_conn()
    c = conn.cursor()
    c.execute('''SELECT COUNT(*) FROM orders
                 WHERE status = 'Tamamlandı' AND strftime('%Y', created_at) = ?''', (str(year),))
    count = c.fetchone()[0]
    conn.close()
    return count


def update_order_kutu(order_id, kutu_manufacturer_id, kutu_siparis_adet, kutu_gelen_adet,
                      kutu_siparis_tarihi, kutu_termin_tarihi):
    conn = get_conn()
    conn.execute('''UPDATE orders SET kutu_manufacturer_id=?, kutu_siparis_adet=?,
                    kutu_gelen_adet=?, kutu_siparis_tarihi=?, kutu_termin_tarihi=? WHERE id=?''',
                (kutu_manufacturer_id, kutu_siparis_adet, kutu_gelen_adet,
                 kutu_siparis_tarihi, kutu_termin_tarihi, order_id))
    conn.commit()
    conn.close()


def update_fabric_row(fabric_row_id, kumas_fabric_id, kumas_renk, manufacturer_id,
                      siparis_kg, gelen_kg, siparis_tarihi, termin_tarihi):
    conn = get_conn()
    conn.execute('''UPDATE order_fabrics SET kumas_fabric_id=?, kumas_renk=?, kumas_manufacturer_id=?,
                    kumas_siparis_kg=?, kumas_gelen_kg=?, kumas_siparis_tarihi=?, kumas_termin_tarihi=?
                    WHERE id=?''',
                (kumas_fabric_id, kumas_renk, manufacturer_id, siparis_kg, gelen_kg,
                 siparis_tarihi, termin_tarihi, fabric_row_id))
    conn.commit()
    conn.close()


def update_elastic_row(elastic_row_id, lastik_elastic_id, lastik_renk, manufacturer_id,
                       siparis_cm, gelen_cm, siparis_adet, gelen_adet, siparis_tarihi, termin_tarihi):
    conn = get_conn()
    conn.execute('''UPDATE order_elastics SET lastik_elastic_id=?, lastik_renk=?, lastik_manufacturer_id=?,
                    lastik_siparis_cm=?, lastik_gelen_cm=?, lastik_siparis_adet=?, lastik_gelen_adet=?,
                    lastik_siparis_tarihi=?, lastik_termin_tarihi=? WHERE id=?''',
                (lastik_elastic_id, lastik_renk, manufacturer_id, siparis_cm, gelen_cm,
                 siparis_adet, gelen_adet, siparis_tarihi, termin_tarihi, elastic_row_id))
    conn.commit()
    conn.close()


def update_accessory_row(accessory_row_id, aksesuar_adi, aksesuar_renk, manufacturer_id,
                         siparis_adet, gelen_adet, siparis_tarihi, termin_tarihi):
    conn = get_conn()
    conn.execute('''UPDATE order_accessories SET aksesuar_adi=?, aksesuar_renk=?, aksesuar_manufacturer_id=?,
                    aksesuar_siparis_adet=?, aksesuar_gelen_adet=?, aksesuar_siparis_tarihi=?,
                    aksesuar_termin_tarihi=? WHERE id=?''',
                (aksesuar_adi, aksesuar_renk, manufacturer_id, siparis_adet, gelen_adet,
                 siparis_tarihi, termin_tarihi, accessory_row_id))
    conn.commit()
    conn.close()


def add_fabric_row(order_id, kumas_fabric_id='', kumas_renk=''):
    conn = get_conn()
    c = conn.cursor()
    seq = c.execute("SELECT COALESCE(MAX(seq),0)+1 FROM order_fabrics WHERE order_id=?", (order_id,)).fetchone()[0]
    c.execute("INSERT INTO order_fabrics (order_id, seq, kumas_fabric_id, kumas_renk) VALUES (?, ?, ?, ?)",
             (order_id, seq, kumas_fabric_id or None, kumas_renk))
    conn.commit()
    conn.close()


def add_elastic_row(order_id, lastik_elastic_id='', lastik_renk=''):
    conn = get_conn()
    c = conn.cursor()
    seq = c.execute("SELECT COALESCE(MAX(seq),0)+1 FROM order_elastics WHERE order_id=?", (order_id,)).fetchone()[0]
    c.execute("INSERT INTO order_elastics (order_id, seq, lastik_elastic_id, lastik_renk) VALUES (?, ?, ?, ?)",
             (order_id, seq, lastik_elastic_id or None, lastik_renk))
    conn.commit()
    conn.close()


def add_accessory_row(order_id, aksesuar_adi='', aksesuar_renk=''):
    conn = get_conn()
    c = conn.cursor()
    seq = c.execute("SELECT COALESCE(MAX(seq),0)+1 FROM order_accessories WHERE order_id=?", (order_id,)).fetchone()[0]
    c.execute("INSERT INTO order_accessories (order_id, seq, aksesuar_adi, aksesuar_renk) VALUES (?, ?, ?, ?)",
             (order_id, seq, aksesuar_adi, aksesuar_renk))
    conn.commit()
    conn.close()


def delete_fabric_row(row_id):
    conn = get_conn()
    conn.execute("DELETE FROM order_fabrics WHERE id = ?", (row_id,))
    conn.commit()
    conn.close()


def delete_elastic_row(row_id):
    conn = get_conn()
    conn.execute("DELETE FROM order_elastics WHERE id = ?", (row_id,))
    conn.commit()
    conn.close()


def delete_accessory_row(row_id):
    conn = get_conn()
    conn.execute("DELETE FROM order_accessories WHERE id = ?", (row_id,))
    conn.commit()
    conn.close()


# --- İrsaliyeler (gelen mal kaydı) - kaydedildiğinde ilgili "gelen" alanını otomatik günceller ---
def add_irsaliye(order_id, item_id, kategori, irsaliye_no, tedarikci, miktar, birim, gelis_tarihi, aciklama=''):
    conn = get_conn()
    c = conn.cursor()
    c.execute('''INSERT INTO irsaliyeler (order_id, item_id, kategori, irsaliye_no, tedarikci, miktar,
                birim, gelis_tarihi, aciklama, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
             (order_id, item_id, kategori, irsaliye_no, tedarikci, miktar, birim, gelis_tarihi,
              aciklama, datetime.now().isoformat()))
    irs_id = c.lastrowid

    # Kalan/gelen miktarları otomatik düş
    if kategori == 'Kumaş' and item_id:
        c.execute("UPDATE order_fabrics SET kumas_gelen_kg = COALESCE(kumas_gelen_kg,0) + ? WHERE id = ?",
                 (miktar, item_id))
    elif kategori == 'Lastik' and item_id:
        if birim == 'adet':
            c.execute("UPDATE order_elastics SET lastik_gelen_adet = COALESCE(lastik_gelen_adet,0) + ? WHERE id = ?",
                     (miktar, item_id))
        else:
            c.execute("UPDATE order_elastics SET lastik_gelen_cm = COALESCE(lastik_gelen_cm,0) + ? WHERE id = ?",
                     (miktar, item_id))
    elif kategori == 'Aksesuar' and item_id:
        c.execute("UPDATE order_accessories SET aksesuar_gelen_adet = COALESCE(aksesuar_gelen_adet,0) + ? WHERE id = ?",
                 (miktar, item_id))
    elif kategori == 'Kutu':
        c.execute("UPDATE orders SET kutu_gelen_adet = COALESCE(kutu_gelen_adet,0) + ? WHERE id = ?",
                 (miktar, order_id))

    conn.commit()
    conn.close()
    return irs_id


def get_irsaliyeler(order_id=None):
    conn = get_conn()
    c = conn.cursor()
    if order_id:
        c.execute("SELECT * FROM irsaliyeler WHERE order_id = ? ORDER BY created_at DESC", (order_id,))
    else:
        c.execute("SELECT * FROM irsaliyeler ORDER BY created_at DESC LIMIT 200")
    rows = [dict(r) for r in c.fetchall()]
    conn.close()
    return rows


# --- Yedekleme (backup) ---
def export_all_data():
    conn = get_conn()
    c = conn.cursor()
    tables = ['users', 'manufacturers', 'product_groups', 'product_names', 'fabrics', 'elastics', 'recipes',
              'orders', 'order_sizes', 'order_fabrics', 'order_elastics', 'order_accessories',
              'logs', 'app_settings', 'irsaliyeler']
    result = {}
    for t in tables:
        c.execute(f"SELECT * FROM {t}")
        result[t] = [dict(r) for r in c.fetchall()]
    conn.close()
    return result


def update_order_status(order_id, status):
    conn = get_conn()
    conn.execute("UPDATE orders SET status = ? WHERE id = ?", (status, order_id))
    conn.commit()
    conn.close()


def update_order_assignment(order_id, manufacturer_id, full_service,
                            fabric_assignments, elastic_assignments, accessory_assignments):
    """
    *_assignments: [{'id': row_id, 'manufacturer_id': ..}]
    """
    conn = get_conn()
    c = conn.cursor()

    if manufacturer_id is None:
        assignment_type = ASSIGNMENT_UNPLANNED
    elif full_service:
        assignment_type = ASSIGNMENT_FULL_SERVICE
    else:
        assignment_type = ASSIGNMENT_PARTIAL

    c.execute('''UPDATE orders SET manufacturer_id = ?, assignment_type = ? WHERE id = ?''',
              (manufacturer_id, assignment_type, order_id))

    if full_service and manufacturer_id:
        c.execute("UPDATE order_fabrics SET kumas_manufacturer_id = ? WHERE order_id = ?", (manufacturer_id, order_id))
        c.execute("UPDATE order_elastics SET lastik_manufacturer_id = ? WHERE order_id = ?", (manufacturer_id, order_id))
        c.execute("UPDATE order_accessories SET aksesuar_manufacturer_id = ? WHERE order_id = ?", (manufacturer_id, order_id))
    else:
        for fa in fabric_assignments:
            c.execute("UPDATE order_fabrics SET kumas_manufacturer_id = ? WHERE id = ? AND order_id = ?",
                     (fa.get('manufacturer_id'), fa['id'], order_id))
        for ea in elastic_assignments:
            c.execute("UPDATE order_elastics SET lastik_manufacturer_id = ? WHERE id = ? AND order_id = ?",
                     (ea.get('manufacturer_id'), ea['id'], order_id))
        for aa in accessory_assignments:
            c.execute("UPDATE order_accessories SET aksesuar_manufacturer_id = ? WHERE id = ? AND order_id = ?",
                     (aa.get('manufacturer_id'), aa['id'], order_id))

    conn.commit()
    conn.close()


def update_order_basic(order_id, model_name, gender, urun_grubu, deadline, siparis_tarihi,
                        kumas_fire_orani, lastik_fire_orani):
    conn = get_conn()
    conn.execute('''UPDATE orders SET model_name = ?, gender = ?, urun_grubu = ?, deadline = ?,
                    siparis_tarihi = ?, kumas_fire_orani = ?, lastik_fire_orani = ? WHERE id = ?''',
                (model_name, gender, urun_grubu, deadline, siparis_tarihi,
                 kumas_fire_orani, lastik_fire_orani, order_id))
    conn.commit()
    conn.close()


def delete_order(order_id):
    conn = get_conn()
    conn.execute("DELETE FROM order_sizes WHERE order_id = ?", (order_id,))
    conn.execute("DELETE FROM order_fabrics WHERE order_id = ?", (order_id,))
    conn.execute("DELETE FROM order_elastics WHERE order_id = ?", (order_id,))
    conn.execute("DELETE FROM order_accessories WHERE order_id = ?", (order_id,))
    conn.execute("DELETE FROM orders WHERE id = ?", (order_id,))
    conn.commit()
    conn.close()
