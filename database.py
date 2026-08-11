import sqlite3
import hashlib
import os
import math
from datetime import datetime

from constants import (
    ASSIGNMENT_FULL_SERVICE, ASSIGNMENT_PARTIAL, ASSIGNMENT_UNPLANNED,
    MATERIAL_STATUS_DONE, MATERIAL_STATUS_PENDING, MATERIAL_STATUS_NA,
    SIZES, CLOSED_STATUSES,
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

    # Kumaş ana verisi
    c.execute('''CREATE TABLE IF NOT EXISTS fabrics (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        icerik TEXT DEFAULT '',
        kumas_turu TEXT DEFAULT 'Örme',
        en REAL DEFAULT 0,
        gr_m2 REAL DEFAULT 0,
        fiyat REAL DEFAULT 0,
        created_at TEXT NOT NULL
    )''')

    # Lastik ana verisi
    c.execute('''CREATE TABLE IF NOT EXISTS elastics (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tur TEXT NOT NULL,
        ad TEXT NOT NULL,
        boyut TEXT DEFAULT '',
        urun_grubu TEXT DEFAULT '',
        fiyat REAL DEFAULT 0,
        created_at TEXT NOT NULL
    )''')

    # Uygulama ayarları (kur, firma, barkod prefix) - tek satır (id=1)
    c.execute('''CREATE TABLE IF NOT EXISTS app_settings (
        id INTEGER PRIMARY KEY CHECK (id = 1),
        usd_kur REAL DEFAULT 34.5,
        eur_kur REAL DEFAULT 37.2,
        varsayilan_para TEXT DEFAULT 'TL',
        firma_adi TEXT DEFAULT 'Tekstil Ltd',
        barkod_prefix TEXT DEFAULT 'TXT'
    )''')

    # İrsaliyeler (gelen mal kayıtları)
    c.execute('''CREATE TABLE IF NOT EXISTS irsaliyeler (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        order_id INTEGER NOT NULL,
        color_id INTEGER,
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

    # Reçete kütüphanesi: cinsiyet + ürün grubu + beden -> kumaş/lastik reçetesi
    c.execute('''CREATE TABLE IF NOT EXISTS recipes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        gender TEXT NOT NULL,
        urun_grubu TEXT NOT NULL,
        size TEXT NOT NULL,
        kumas_gr REAL DEFAULT 0,
        lastik_adet REAL DEFAULT 0,
        lastik_mt REAL DEFAULT 0,
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
        total_quantity INTEGER DEFAULT 0,
        total_boxes INTEGER DEFAULT 0,
        status TEXT DEFAULT 'Planlama',
        urun_foto TEXT DEFAULT '',
        model_kodu TEXT DEFAULT '',

        para_birimi TEXT DEFAULT 'TL',
        iscilik_birim REAL DEFAULT 0,
        genel_gider_yuzde REAL DEFAULT 10,
        kar_yuzde REAL DEFAULT 30,

        kutu_manufacturer_id INTEGER,
        kutu_fiyat REAL DEFAULT 0,
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
        lastik_mt REAL NOT NULL DEFAULT 0,
        FOREIGN KEY (order_id) REFERENCES orders(id)
    )''')

    # Paket içi renkler / malzeme detayları
    c.execute('''CREATE TABLE IF NOT EXISTS order_colors (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        order_id INTEGER NOT NULL,
        seq INTEGER NOT NULL DEFAULT 1,
        color_name TEXT NOT NULL DEFAULT '',

        kumas_fabric_id INTEGER,
        kumas_renk TEXT DEFAULT '',
        kumas_foto TEXT DEFAULT '',
        kumas_manufacturer_id INTEGER,
        kumas_siparis_kg REAL DEFAULT 0,
        kumas_gelen_kg REAL DEFAULT 0,
        kumas_siparis_tarihi TEXT DEFAULT '',
        kumas_termin_tarihi TEXT DEFAULT '',

        lastik_elastic_id INTEGER,
        lastik_renk TEXT DEFAULT '',
        lastik_foto TEXT DEFAULT '',
        lastik_manufacturer_id INTEGER,
        lastik_siparis_mt REAL DEFAULT 0,
        lastik_gelen_mt REAL DEFAULT 0,
        lastik_siparis_tarihi TEXT DEFAULT '',
        lastik_termin_tarihi TEXT DEFAULT '',

        aksesuar_adi TEXT DEFAULT '',
        aksesuar_renk TEXT DEFAULT '',
        aksesuar_foto TEXT DEFAULT '',
        aksesuar_manufacturer_id INTEGER,
        aksesuar_fiyat REAL DEFAULT 0,
        aksesuar_siparis_adet REAL DEFAULT 0,
        aksesuar_gelen_adet REAL DEFAULT 0,
        aksesuar_siparis_tarihi TEXT DEFAULT '',
        aksesuar_termin_tarihi TEXT DEFAULT '',

        FOREIGN KEY (order_id) REFERENCES orders(id),
        FOREIGN KEY (kumas_fabric_id) REFERENCES fabrics(id),
        FOREIGN KEY (lastik_elastic_id) REFERENCES elastics(id),
        FOREIGN KEY (kumas_manufacturer_id) REFERENCES manufacturers(id),
        FOREIGN KEY (lastik_manufacturer_id) REFERENCES manufacturers(id),
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

    # --- Migrasyonlar: eski şemalarda eksik kolonları ekle ---
    migrations = [
        ('manufacturers', 'contact_person', "TEXT DEFAULT ''"),
        ('manufacturers', 'phone', "TEXT DEFAULT ''"),
        ('manufacturers', 'email', "TEXT DEFAULT ''"),
        ('orders', 'urun_grubu', "TEXT DEFAULT ''"),
        ('orders', 'kumas_fire_orani', "REAL NOT NULL DEFAULT 3"),
        ('orders', 'lastik_fire_orani', "REAL NOT NULL DEFAULT 3"),
        ('orders', 'urun_foto', "TEXT DEFAULT ''"),
        ('orders', 'gender', "TEXT NOT NULL DEFAULT 'Kadın'"),
        ('orders', 'assignment_type', "TEXT NOT NULL DEFAULT 'planlama'"),
        ('orders', 'package_size', "INTEGER NOT NULL DEFAULT 1"),
        ('orders', 'total_boxes', "INTEGER DEFAULT 0"),
        ('order_colors', 'kumas_fabric_id', "INTEGER"),
        ('order_colors', 'kumas_renk', "TEXT DEFAULT ''"),
        ('order_colors', 'kumas_foto', "TEXT DEFAULT ''"),
        ('order_colors', 'kumas_siparis_kg', "REAL DEFAULT 0"),
        ('order_colors', 'kumas_gelen_kg', "REAL DEFAULT 0"),
        ('order_colors', 'kumas_siparis_tarihi', "TEXT DEFAULT ''"),
        ('order_colors', 'kumas_termin_tarihi', "TEXT DEFAULT ''"),
        ('order_colors', 'lastik_elastic_id', "INTEGER"),
        ('order_colors', 'lastik_renk', "TEXT DEFAULT ''"),
        ('order_colors', 'lastik_foto', "TEXT DEFAULT ''"),
        ('order_colors', 'aksesuar_foto', "TEXT DEFAULT ''"),
        ('fabrics', 'kumas_turu', "TEXT DEFAULT 'Örme'"),
        ('fabrics', 'fiyat', "REAL DEFAULT 0"),
        ('elastics', 'fiyat', "REAL DEFAULT 0"),
        ('orders', 'model_kodu', "TEXT DEFAULT ''"),
        ('orders', 'para_birimi', "TEXT DEFAULT 'TL'"),
        ('orders', 'iscilik_birim', "REAL DEFAULT 0"),
        ('orders', 'genel_gider_yuzde', "REAL DEFAULT 10"),
        ('orders', 'kar_yuzde', "REAL DEFAULT 30"),
        ('orders', 'kutu_manufacturer_id', "INTEGER"),
        ('orders', 'kutu_fiyat', "REAL DEFAULT 0"),
        ('orders', 'kutu_siparis_adet', "REAL DEFAULT 0"),
        ('orders', 'kutu_gelen_adet', "REAL DEFAULT 0"),
        ('orders', 'kutu_siparis_tarihi', "TEXT DEFAULT ''"),
        ('orders', 'kutu_termin_tarihi', "TEXT DEFAULT ''"),
        ('order_colors', 'lastik_siparis_mt', "REAL DEFAULT 0"),
        ('order_colors', 'lastik_gelen_mt', "REAL DEFAULT 0"),
        ('order_colors', 'lastik_siparis_tarihi', "TEXT DEFAULT ''"),
        ('order_colors', 'lastik_termin_tarihi', "TEXT DEFAULT ''"),
        ('order_colors', 'aksesuar_fiyat', "REAL DEFAULT 0"),
        ('order_colors', 'aksesuar_siparis_adet', "REAL DEFAULT 0"),
        ('order_colors', 'aksesuar_gelen_adet', "REAL DEFAULT 0"),
        ('order_colors', 'aksesuar_siparis_tarihi', "TEXT DEFAULT ''"),
        ('order_colors', 'aksesuar_termin_tarihi', "TEXT DEFAULT ''"),
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
        c.execute('''INSERT INTO app_settings (id, usd_kur, eur_kur, varsayilan_para, firma_adi, barkod_prefix)
                     VALUES (1, 34.5, 37.2, 'TL', 'Paul Kenzie', 'PK')''')
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


# --- Ayarlar (kur, firma, barkod) ---
def get_settings():
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT * FROM app_settings WHERE id = 1")
    row = c.fetchone()
    conn.close()
    return dict(row) if row else {
        'usd_kur': 34.5, 'eur_kur': 37.2, 'varsayilan_para': 'TL',
        'firma_adi': 'Tekstil Ltd', 'barkod_prefix': 'TXT',
    }


def update_settings(usd_kur, eur_kur, varsayilan_para, firma_adi, barkod_prefix):
    conn = get_conn()
    conn.execute('''UPDATE app_settings SET usd_kur=?, eur_kur=?, varsayilan_para=?,
                    firma_adi=?, barkod_prefix=? WHERE id=1''',
                (usd_kur, eur_kur, varsayilan_para, firma_adi, barkod_prefix))
    conn.commit()
    conn.close()


def generate_model_code(barkod_prefix, order_id, renk=''):
    renk_kod = ''.join([ch for ch in (renk or '').upper() if ch.isalnum()])[:2] or 'XX'
    return f"{barkod_prefix}-{order_id:03d}-{renk_kod}"


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


# --- Fabrics (kumaş ana verisi) ---
def get_fabrics():
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT * FROM fabrics ORDER BY name")
    rows = [dict(r) for r in c.fetchall()]
    conn.close()
    return rows


def add_fabric(name, icerik, kumas_turu, en, gr_m2):
    conn = get_conn()
    conn.execute('''INSERT INTO fabrics (name, icerik, kumas_turu, en, gr_m2, created_at)
                    VALUES (?, ?, ?, ?, ?, ?)''',
                 (name, icerik, kumas_turu, en, gr_m2, datetime.now().isoformat()))
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
    c.execute("SELECT * FROM elastics ORDER BY tur, ad")
    rows = [dict(r) for r in c.fetchall()]
    conn.close()
    return rows


def add_elastic(tur, ad, boyut, urun_grubu=''):
    conn = get_conn()
    conn.execute('''INSERT INTO elastics (tur, ad, boyut, urun_grubu, created_at)
                    VALUES (?, ?, ?, ?, ?)''',
                 (tur, ad, boyut, urun_grubu, datetime.now().isoformat()))
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


def get_urun_gruplari_with_recipe(gender=None):
    conn = get_conn()
    c = conn.cursor()
    if gender:
        c.execute("SELECT DISTINCT urun_grubu FROM recipes WHERE gender = ? ORDER BY urun_grubu", (gender,))
    else:
        c.execute("SELECT DISTINCT urun_grubu FROM recipes ORDER BY urun_grubu")
    rows = [r['urun_grubu'] for r in c.fetchall()]
    conn.close()
    return rows


def get_recipe(gender, urun_grubu):
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT * FROM recipes WHERE gender = ? AND urun_grubu = ?", (gender, urun_grubu))
    rows = [dict(r) for r in c.fetchall()]
    conn.close()
    result = {s: {'kumas_gr': 0, 'lastik_adet': 0, 'lastik_mt': 0} for s in SIZES}
    for r in rows:
        result[r['size']] = {'kumas_gr': r['kumas_gr'], 'lastik_adet': r['lastik_adet'], 'lastik_mt': r['lastik_mt']}
    return result


def upsert_recipe(gender, urun_grubu, size, kumas_gr, lastik_adet, lastik_mt):
    conn = get_conn()
    conn.execute('''INSERT INTO recipes (gender, urun_grubu, size, kumas_gr, lastik_adet, lastik_mt, created_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(gender, urun_grubu, size)
                    DO UPDATE SET kumas_gr = excluded.kumas_gr, lastik_adet = excluded.lastik_adet,
                                  lastik_mt = excluded.lastik_mt''',
                 (gender, urun_grubu, size, kumas_gr, lastik_adet, lastik_mt, datetime.now().isoformat()))
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
              kumas_fire_orani, lastik_fire_orani, deadline, sizes, colors, user_id, urun_foto='',
              para_birimi='TL', iscilik_birim=0, genel_gider_yuzde=10, kar_yuzde=30,
              kutu_manufacturer_id=None, kutu_fiyat=0):
    """
    Yeni sipariş her zaman 'Planlama Bekleyen' (manufacturer_id=None) olarak oluşturulur.
    Üretici ataması sadece sipariş düzenleme sayfasında yapılır.

    sizes:  [{'size': 'M', 'box_qty': 10, 'kumas_gr': 150, 'lastik_adet': 2, 'lastik_mt': 0.5}, ...]
    colors: [{'color_name':.., 'kumas_fabric_id':.., 'kumas_renk':.., 'kumas_foto':..,
              'lastik_elastic_id':.., 'lastik_renk':.., 'lastik_foto':..,
              'aksesuar_adi':.., 'aksesuar_renk':.., 'aksesuar_foto':..}, ...]
    """
    conn = get_conn()
    c = conn.cursor()

    total_boxes = sum(s['box_qty'] for s in sizes)
    total_quantity = total_boxes * package_size

    c.execute('''INSERT INTO orders
                 (model_name, gender, urun_grubu, manufacturer_id, assignment_type, package_size,
                  kumas_fire_orani, lastik_fire_orani, deadline, total_quantity, total_boxes, status,
                  urun_foto, para_birimi, iscilik_birim, genel_gider_yuzde, kar_yuzde,
                  kutu_manufacturer_id, kutu_fiyat, kutu_siparis_adet,
                  created_at, created_by)
                 VALUES (?, ?, ?, NULL, ?, ?, ?, ?, ?, ?, ?, 'Planlama', ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
              (model_name, gender, urun_grubu, ASSIGNMENT_UNPLANNED, package_size,
               kumas_fire_orani, lastik_fire_orani, deadline, total_quantity, total_boxes,
               urun_foto, para_birimi, iscilik_birim, genel_gider_yuzde, kar_yuzde,
               kutu_manufacturer_id, kutu_fiyat, math.ceil(total_quantity / package_size) if package_size else total_boxes,
               datetime.now().isoformat(), user_id))
    order_id = c.lastrowid

    # Ayarlar'daki barkod prefix ile model kodu oluştur
    settings = c.execute("SELECT barkod_prefix FROM app_settings WHERE id=1").fetchone()
    prefix = settings['barkod_prefix'] if settings else 'TXT'
    ilk_renk = colors[0].get('color_name', '') if colors else ''
    model_kodu = generate_model_code(prefix, order_id, ilk_renk)
    c.execute("UPDATE orders SET model_kodu = ? WHERE id = ?", (model_kodu, order_id))

    for s in sizes:
        if s['box_qty'] > 0:
            c.execute('''INSERT INTO order_sizes (order_id, size, box_qty, kumas_gr, lastik_adet, lastik_mt)
                         VALUES (?, ?, ?, ?, ?, ?)''',
                      (order_id, s['size'], s['box_qty'], s.get('kumas_gr', 0),
                       s.get('lastik_adet', 0), s.get('lastik_mt', 0)))

    color_ids = []
    for i, col in enumerate(colors, start=1):
        c.execute('''INSERT INTO order_colors
                     (order_id, seq, color_name, kumas_fabric_id, kumas_renk, kumas_foto,
                      lastik_elastic_id, lastik_renk, lastik_foto,
                      aksesuar_adi, aksesuar_renk, aksesuar_foto)
                     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                  (order_id, i, col.get('color_name', f'Renk {i}'),
                   col.get('kumas_fabric_id'), col.get('kumas_renk', ''), col.get('kumas_foto', ''),
                   col.get('lastik_elastic_id'), col.get('lastik_renk', ''), col.get('lastik_foto', ''),
                   col.get('aksesuar_adi', ''), col.get('aksesuar_renk', ''), col.get('aksesuar_foto', '')))
        color_ids.append(c.lastrowid)

    conn.commit()
    conn.close()
    return order_id, color_ids


def set_order_photo(order_id, path):
    conn = get_conn()
    conn.execute("UPDATE orders SET urun_foto = ? WHERE id = ?", (path, order_id))
    conn.commit()
    conn.close()


def set_color_photos(color_id, kumas_foto=None, lastik_foto=None, aksesuar_foto=None):
    conn = get_conn()
    if kumas_foto is not None:
        conn.execute("UPDATE order_colors SET kumas_foto = ? WHERE id = ?", (kumas_foto, color_id))
    if lastik_foto is not None:
        conn.execute("UPDATE order_colors SET lastik_foto = ? WHERE id = ?", (lastik_foto, color_id))
    if aksesuar_foto is not None:
        conn.execute("UPDATE order_colors SET aksesuar_foto = ? WHERE id = ?", (aksesuar_foto, color_id))
    conn.commit()
    conn.close()


def _material_status(material_manufacturer_id, order_manufacturer_id, has_material=True):
    if not has_material:
        return MATERIAL_STATUS_NA
    if order_manufacturer_id and material_manufacturer_id == order_manufacturer_id:
        return MATERIAL_STATUS_DONE
    return MATERIAL_STATUS_PENDING


def _compute_order_totals(order, sizes, colors):
    kumas_fire_mult = 1 + (order['kumas_fire_orani'] or 0) / 100.0
    lastik_fire_mult = 1 + (order['lastik_fire_orani'] or 0) / 100.0
    total_boxes = sum(s['box_qty'] for s in sizes)

    colors_out = []
    for col in colors:
        color_kumas_gr = sum(s['box_qty'] * s['kumas_gr'] for s in sizes) * kumas_fire_mult
        color_lastik_adet = sum(s['box_qty'] * s['lastik_adet'] for s in sizes) * lastik_fire_mult
        color_lastik_mt = sum(s['box_qty'] * s['lastik_mt'] for s in sizes) * lastik_fire_mult
        has_aksesuar = bool((col.get('aksesuar_adi') or '').strip())

        kumas_kg_required = round(color_kumas_gr / 1000.0, 3)
        kumas_gelen = col.get('kumas_gelen_kg') or 0
        kumas_kalan = round(kumas_kg_required - kumas_gelen, 3)

        colors_out.append({
            **col,
            'pieces': total_boxes,
            'kumas_kg': kumas_kg_required,
            'kumas_kalan_kg': kumas_kalan,
            'lastik_adet_toplam': round(color_lastik_adet, 1),
            'lastik_mt_toplam': round(color_lastik_mt, 1),
            'kumas_status': _material_status(col.get('kumas_manufacturer_id'), order['manufacturer_id']),
            'lastik_status': _material_status(col.get('lastik_manufacturer_id'), order['manufacturer_id']),
            'aksesuar_status': _material_status(col.get('aksesuar_manufacturer_id'), order['manufacturer_id'], has_aksesuar),
        })

    return colors_out


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

    c.execute('''SELECT oc.*, f.name as kumas_adi, f.icerik as kumas_icerik, f.kumas_turu as kumas_turu,
                        f.fiyat as kumas_fiyat_kg,
                        e.tur as lastik_tur, e.ad as lastik_adi, e.boyut as lastik_genislik,
                        e.fiyat as lastik_fiyat_mt,
                        km.name as kumas_manufacturer_name,
                        lm.name as lastik_manufacturer_name, am.name as aksesuar_manufacturer_name
                 FROM order_colors oc
                 LEFT JOIN fabrics f ON oc.kumas_fabric_id = f.id
                 LEFT JOIN elastics e ON oc.lastik_elastic_id = e.id
                 LEFT JOIN manufacturers km ON oc.kumas_manufacturer_id = km.id
                 LEFT JOIN manufacturers lm ON oc.lastik_manufacturer_id = lm.id
                 LEFT JOIN manufacturers am ON oc.aksesuar_manufacturer_id = am.id
                 WHERE oc.order_id = ? ORDER BY oc.seq''', (order_id,))
    colors = [dict(r) for r in c.fetchall()]
    conn.close()

    order['sizes'] = sizes
    order['colors'] = _compute_order_totals(order, sizes, colors)
    order['kumas_kg_total'] = round(sum(c_['kumas_kg'] for c_ in order['colors']), 3)
    order['lastik_mt_total'] = round(sum(c_['lastik_mt_toplam'] for c_ in order['colors']), 1)
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


def compute_order_costs(order):
    """Kur'u TL'ye çevirerek maliyet/kâr hesaplar. order = get_order_detail() çıktısı."""
    settings = get_settings()
    kur = {'TL': 1.0, 'USD': settings['usd_kur'], 'EUR': settings['eur_kur']}.get(order.get('para_birimi', 'TL'), 1.0)

    kumas_maliyet = sum((c_['kumas_kg'] or 0) * (c_.get('kumas_fiyat_kg') or 0) for c_ in order['colors'])
    lastik_maliyet = sum((c_['lastik_mt_toplam'] or 0) * (c_.get('lastik_fiyat_mt') or 0) for c_ in order['colors'])
    kutu_maliyet = (order.get('kutu_fiyat') or 0) * (order.get('total_boxes') or 0)
    aksesuar_maliyet = sum((c_['pieces'] or 0) * (c_.get('aksesuar_fiyat') or 0)
                           for c_ in order['colors'] if (c_.get('aksesuar_adi') or '').strip())
    iscilik_maliyet = (order.get('iscilik_birim') or 0) * (order.get('total_quantity') or 0)

    malzeme_toplam = kumas_maliyet + lastik_maliyet + kutu_maliyet + aksesuar_maliyet
    ara_toplam = malzeme_toplam + iscilik_maliyet
    genel_gider = ara_toplam * (order.get('genel_gider_yuzde') or 0) / 100.0
    toplam_maliyet_pb = ara_toplam + genel_gider
    kar_pb = toplam_maliyet_pb * (order.get('kar_yuzde') or 0) / 100.0
    satis_toplam_pb = toplam_maliyet_pb + kar_pb

    qty = order.get('total_quantity') or 1

    return {
        'para_birimi': order.get('para_birimi', 'TL'),
        'kur': kur,
        'kumas_maliyet': round(kumas_maliyet, 2),
        'lastik_maliyet': round(lastik_maliyet, 2),
        'kutu_maliyet': round(kutu_maliyet, 2),
        'aksesuar_maliyet': round(aksesuar_maliyet, 2),
        'iscilik_maliyet': round(iscilik_maliyet, 2),
        'toplam_maliyet_pb': round(toplam_maliyet_pb, 2),
        'satis_toplam_pb': round(satis_toplam_pb, 2),
        'kar_toplam_pb': round(kar_pb, 2),
        'birim_maliyet_pb': round(toplam_maliyet_pb / qty, 2),
        'satis_birim_pb': round(satis_toplam_pb / qty, 2),
        'toplam_maliyet_tl': round(toplam_maliyet_pb * kur, 2),
        'satis_toplam_tl': round(satis_toplam_pb * kur, 2),
        'kar_toplam_tl': round(kar_pb * kur, 2),
    }


def update_order_kutu(order_id, kutu_manufacturer_id, kutu_fiyat, kutu_siparis_adet, kutu_gelen_adet,
                      kutu_siparis_tarihi, kutu_termin_tarihi):
    conn = get_conn()
    conn.execute('''UPDATE orders SET kutu_manufacturer_id=?, kutu_fiyat=?, kutu_siparis_adet=?,
                    kutu_gelen_adet=?, kutu_siparis_tarihi=?, kutu_termin_tarihi=? WHERE id=?''',
                (kutu_manufacturer_id, kutu_fiyat, kutu_siparis_adet, kutu_gelen_adet,
                 kutu_siparis_tarihi, kutu_termin_tarihi, order_id))
    conn.commit()
    conn.close()


def update_color_lastik_tracking(color_id, fiyat, siparis_mt, gelen_mt, siparis_tarihi, termin_tarihi):
    conn = get_conn()
    conn.execute('''UPDATE order_colors SET lastik_siparis_mt=?, lastik_gelen_mt=?,
                    lastik_siparis_tarihi=?, lastik_termin_tarihi=? WHERE id=?''',
                (siparis_mt, gelen_mt, siparis_tarihi, termin_tarihi, color_id))
    conn.execute("UPDATE elastics SET fiyat=? WHERE id = (SELECT lastik_elastic_id FROM order_colors WHERE id=?)",
                (fiyat, color_id))
    conn.commit()
    conn.close()


def update_color_aksesuar_tracking(color_id, fiyat, siparis_adet, gelen_adet, siparis_tarihi, termin_tarihi):
    conn = get_conn()
    conn.execute('''UPDATE order_colors SET aksesuar_fiyat=?, aksesuar_siparis_adet=?, aksesuar_gelen_adet=?,
                    aksesuar_siparis_tarihi=?, aksesuar_termin_tarihi=? WHERE id=?''',
                (fiyat, siparis_adet, gelen_adet, siparis_tarihi, termin_tarihi, color_id))
    conn.commit()
    conn.close()


# --- İrsaliyeler ---
def add_irsaliye(order_id, color_id, kategori, irsaliye_no, tedarikci, miktar, birim, gelis_tarihi, aciklama=''):
    conn = get_conn()
    c = conn.cursor()
    c.execute('''INSERT INTO irsaliyeler (order_id, color_id, kategori, irsaliye_no, tedarikci, miktar,
                birim, gelis_tarihi, aciklama, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
             (order_id, color_id, kategori, irsaliye_no, tedarikci, miktar, birim, gelis_tarihi,
              aciklama, datetime.now().isoformat()))
    irs_id = c.lastrowid
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
    tables = ['users', 'manufacturers', 'product_groups', 'fabrics', 'elastics', 'recipes',
              'orders', 'order_sizes', 'order_colors', 'logs', 'app_settings', 'irsaliyeler']
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


def update_order_assignment(order_id, manufacturer_id, full_service, color_assignments):
    """
    color_assignments: [{'id': color_id, 'kumas_manufacturer_id':.., 'lastik_manufacturer_id':.., 'aksesuar_manufacturer_id':..}]
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
        c.execute('''UPDATE order_colors
                     SET kumas_manufacturer_id = ?, lastik_manufacturer_id = ?,
                         aksesuar_manufacturer_id = CASE WHEN aksesuar_adi != '' THEN ? ELSE aksesuar_manufacturer_id END
                     WHERE order_id = ?''',
                  (manufacturer_id, manufacturer_id, manufacturer_id, order_id))
    else:
        for ca in color_assignments:
            c.execute('''UPDATE order_colors
                         SET kumas_manufacturer_id = ?, lastik_manufacturer_id = ?, aksesuar_manufacturer_id = ?
                         WHERE id = ? AND order_id = ?''',
                      (ca.get('kumas_manufacturer_id'), ca.get('lastik_manufacturer_id'),
                       ca.get('aksesuar_manufacturer_id'), ca['id'], order_id))

    conn.commit()
    conn.close()


def update_order_basic(order_id, model_name, gender, urun_grubu, deadline,
                        kumas_fire_orani, lastik_fire_orani):
    conn = get_conn()
    conn.execute('''UPDATE orders SET model_name = ?, gender = ?, urun_grubu = ?, deadline = ?,
                    kumas_fire_orani = ?, lastik_fire_orani = ? WHERE id = ?''',
                (model_name, gender, urun_grubu, deadline, kumas_fire_orani, lastik_fire_orani, order_id))
    conn.commit()
    conn.close()


def update_fabric_tracking(color_id, siparis_kg, gelen_kg, siparis_tarihi, termin_tarihi):
    conn = get_conn()
    conn.execute('''UPDATE order_colors SET kumas_siparis_kg = ?, kumas_gelen_kg = ?,
                    kumas_siparis_tarihi = ?, kumas_termin_tarihi = ? WHERE id = ?''',
                (siparis_kg, gelen_kg, siparis_tarihi, termin_tarihi, color_id))
    conn.commit()
    conn.close()


def delete_order(order_id):
    conn = get_conn()
    conn.execute("DELETE FROM order_sizes WHERE order_id = ?", (order_id,))
    conn.execute("DELETE FROM order_colors WHERE order_id = ?", (order_id,))
    conn.execute("DELETE FROM orders WHERE id = ?", (order_id,))
    conn.commit()
    conn.close()
