import streamlit as st
import pandas as pd
import os
import io
import zipfile
import json
from datetime import datetime, date, timedelta
import database as db
from constants import (
    SIZES, GENDERS, PACKAGE_SIZES, ORDER_STATUSES, CLOSED_STATUSES, ELASTIC_TYPES, FABRIC_TYPES,
    DEFAULT_PRODUCT_GROUPS, IRSALIYE_KATEGORILERI, RAPORLU_LASTIK,
    MATERIAL_STATUS_DONE, MATERIAL_STATUS_PENDING, MATERIAL_STATUS_NA,
    MATERIAL_STATUS_COLORS,
)

try:
    import openpyxl
    HAS_EXCEL = True
except ImportError:
    HAS_EXCEL = False

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas
    from reportlab.lib.units import mm
    HAS_PDF = True
except ImportError:
    HAS_PDF = False

# --- Page config ---
st.set_page_config(page_title="Paul Kenzie ERP", layout="wide", page_icon="🧥",
                    initial_sidebar_state="expanded")

# --- Initialize DB ---
db.init_db()

UPLOAD_DIR = os.path.join(os.path.dirname(__file__), 'uploads')
os.makedirs(UPLOAD_DIR, exist_ok=True)

TR_DATE_FMT = "DD.MM.YYYY"

# ======================================================================
# GÖRSEL TEMA / CSS
# ======================================================================
CUSTOM_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

html, body, .stApp { font-family: 'Inter', -apple-system, sans-serif; }
h1, h2, h3, h4, h5, h6,
p, li, label,
.stMarkdown, .stCaption,
.stButton button, .stDownloadButton button,
.stTextInput input, .stNumberInput input, .stDateInput input, .stTextArea textarea,
.stSelectbox div[data-baseweb="select"] *,
.stRadio label, .stCheckbox label,
div[data-testid="stMetricValue"], div[data-testid="stMetricLabel"] {
    font-family: 'Inter', -apple-system, sans-serif !important;
}
[data-testid="stIconMaterial"],
[data-testid="stExpanderToggleIcon"],
span[class*="material-icons"],
span[class*="material-symbols"] {
    font-family: 'Material Symbols Rounded', 'Material Icons' !important;
}

#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header[data-testid="stHeader"] {background: transparent;}

.stApp { background: #f4f6fa; }

/* --- Kompakt giriş kutuları (Reçeteler hariç genel görünüm) --- */
.stTextInput input, .stNumberInput input, .stDateInput input {
    padding: 5px 10px !important;
    font-size: 0.84rem !important;
    min-height: 32px !important;
}
.stSelectbox div[data-baseweb="select"] > div {
    min-height: 32px !important;
    font-size: 0.84rem !important;
}

/* --- Sidebar --- */
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0f172a 0%, #1e2a44 100%);
    border-right: 1px solid #1e293b;
}
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3 { color: #e2e8f0 !important; }
section[data-testid="stSidebar"] div[role="radiogroup"] label {
    padding: 8px 12px; border-radius: 8px; margin-bottom: 2px;
}
section[data-testid="stSidebar"] hr { border-color: rgba(255,255,255,0.12); }
section[data-testid="stSidebar"] .stButton button {
    background: rgba(255,255,255,0.08); color: #e2e8f0 !important; border: 1px solid rgba(255,255,255,0.15);
    border-radius: 8px;
}
section[data-testid="stSidebar"] .stButton button:hover { background: rgba(239,68,68,0.85); border-color: transparent; }

.pk-brand { padding: 6px 4px 18px 4px; }
.pk-brand .pk-logo { font-size: 1.5rem; font-weight: 800; letter-spacing: 0.5px; color: #f8fafc !important; }
.pk-brand .pk-sub { font-size: 0.72rem; color: #94a3b8 !important; letter-spacing: 1.5px; text-transform: uppercase; }
.pk-user-card {
    background: rgba(255,255,255,0.06); border-radius: 10px; padding: 10px 12px; margin-bottom: 14px;
    border: 1px solid rgba(255,255,255,0.08);
}

h1, h2, h3 { color: #0f172a; font-weight: 700 !important; }
.pk-page-header {
    display:flex; align-items:center; justify-content:space-between;
    padding-bottom: 14px; margin-bottom: 18px; border-bottom: 2px solid #e2e8f0;
}
.pk-page-title { font-size: 1.5rem; font-weight: 800; color: #0f172a; }
.pk-page-sub { color: #64748b; font-size: 0.85rem; margin-top: 2px; }

div[data-testid="stMetric"] {
    background: #ffffff; border: 1px solid #e2e8f0; border-radius: 12px;
    padding: 14px 16px; box-shadow: 0 1px 3px rgba(15,23,42,0.04);
}
div[data-testid="stMetric"] label { color: #64748b !important; font-weight: 600; }

.pk-card {
    background:#fff; border:1px solid #e2e8f0; border-radius:12px; padding:18px 20px;
    box-shadow: 0 1px 3px rgba(15,23,42,0.04); margin-bottom: 16px;
}
.pk-section-title {
    font-size: 1.05rem; font-weight: 700; color:#0f172a; margin: 18px 0 10px 0;
}
.pk-badge {
    display:inline-block; padding: 3px 11px; border-radius: 999px; font-size: 0.74rem;
    font-weight: 700; white-space: nowrap;
}
div[data-testid="stDataFrame"] { border-radius: 10px; overflow: hidden; border: 1px solid #e2e8f0; }
.stButton button { border-radius: 8px; font-weight: 600; border: 1px solid #cbd5e1; }
.stTabs [data-baseweb="tab"] { font-weight: 600; }
div[data-testid="stForm"] { background:#fff; border:1px solid #e2e8f0; border-radius:12px; padding: 18px 20px; }

.pk-steps { display:flex; gap: 6px; margin-bottom: 18px; }
.pk-step {
    flex:1; text-align:center; padding: 10px 6px; border-radius: 8px; font-weight:700; font-size:0.85rem;
    background:#eef2f7; color:#94a3b8; border: 1px solid #e2e8f0;
}
.pk-step.done { background:#dcfce7; color:#15803d; border-color:#bbf7d0; }
.pk-step.active { background:#1e3a8a; color:#fff; border-color:#1e3a8a; }

.pk-track-table { width:100%; border-collapse: collapse; font-size: 0.86rem; }
.pk-track-table th { text-align:left; padding: 8px 10px; background:#f1f5f9; color:#475569; font-weight:700; border-bottom: 2px solid #e2e8f0; }
.pk-track-table td { padding: 8px 10px; border-bottom: 1px solid #e2e8f0; }

/* --- Sipariş kartı metrik kutucukları (gölgeli, durum renkli) --- */
.pk-chip-row { display:flex; gap:10px; margin: 6px 0 14px 0; flex-wrap: wrap; }
.pk-chip {
    flex: 1; min-width: 110px; border-radius: 10px; padding: 10px 14px;
    box-shadow: 0 2px 6px rgba(15,23,42,0.08);
    border: 1px solid rgba(0,0,0,0.04);
}
.pk-chip .pk-chip-label { font-size: 0.7rem; font-weight: 700; text-transform: uppercase; opacity: 0.75; }
.pk-chip .pk-chip-value { font-size: 1.15rem; font-weight: 800; margin-top: 2px; }
.pk-chip.green { background: #dcfce7; color: #15803d; }
.pk-chip.yellow { background: #fef9c3; color: #a16207; }
.pk-chip.red { background: #fee2e2; color: #b91c1c; }
.pk-chip.gray { background: #f1f5f9; color: #475569; }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# --- Session defaults ---
if 'logged_in' not in st.session_state:
    st.session_state.logged_in = False
if 'user' not in st.session_state:
    st.session_state.user = None
if 'edit_order_id' not in st.session_state:
    st.session_state.edit_order_id = None

# ======================================================================
# YARDIMCI FONKSİYONLAR
# ======================================================================

def tr_date(iso_str):
    """ISO (YYYY-MM-DD) tarihi Türkçe (GG.AA.YYYY) formatına çevirir."""
    if not iso_str:
        return '-'
    try:
        return datetime.strptime(iso_str[:10], '%Y-%m-%d').strftime('%d.%m.%Y')
    except Exception:
        return iso_str


def urgency_bucket(deadline_str):
    """15+ gün -> green, 1-14 gün -> yellow, <=0 gün -> red, tarih yoksa -> gray"""
    if not deadline_str:
        return 'gray'
    try:
        dl = datetime.strptime(deadline_str[:10], '%Y-%m-%d').date()
        diff = (dl - date.today()).days
        if diff <= 0:
            return 'red'
        elif diff <= 14:
            return 'yellow'
        else:
            return 'green'
    except Exception:
        return 'gray'


def get_urgency_label(deadline_str):
    if not deadline_str:
        return '⚪ Belirtilmedi'
    try:
        dl = datetime.strptime(deadline_str[:10], '%Y-%m-%d').date()
        diff = (dl - date.today()).days
        if diff < 0:
            return f'🔴 GECİKMİŞ ({abs(diff)} gün)'
        elif diff == 0:
            return '🔴 BUGÜN'
        elif diff <= 3:
            return f'🔴 {diff} gün kaldı'
        elif diff <= 7:
            return f'🟠 {diff} gün kaldı'
        elif diff <= 14:
            return f'🟡 {diff} gün kaldı'
        else:
            return f'🟢 {diff} gün kaldı'
    except Exception:
        return '⚪ Belirtilmedi'


def fabric_row_color(termin_str):
    if not termin_str:
        return '#ffffff'
    bucket = urgency_bucket(termin_str)
    return {'green': '#dcfce7', 'yellow': '#ffedd5', 'red': '#fee2e2', 'gray': '#ffffff'}[bucket]


def badge(text, color):
    return f'<span class="pk-badge" style="background:{color}1f;color:{color};">{text}</span>'


def material_badge(status):
    color = MATERIAL_STATUS_COLORS.get(status, '#94a3b8')
    icon = {'Tamamlandı': '✅', 'Bekliyor': '⏳', 'Yok': '—'}.get(status, '')
    return badge(f'{icon} {status}', color)


def page_header(title, subtitle=''):
    st.markdown(
        f'<div class="pk-page-header"><div><div class="pk-page-title">{title}</div>'
        f'<div class="pk-page-sub">{subtitle}</div></div></div>',
        unsafe_allow_html=True,
    )


def mf_options(manufacturers, allow_none_label="— Atama yok —"):
    names = [allow_none_label] + [m['name'] for m in manufacturers]
    ids = [None] + [m['id'] for m in manufacturers]
    return names, ids


def none_select(label, manufacturers, key, index=0, none_label="— Atama yok —"):
    names, ids = mf_options(manufacturers, none_label)
    sel = st.selectbox(label, names, index=index, key=key)
    return ids[names.index(sel)]


def save_uploaded_file(uploaded_file, prefix):
    if uploaded_file is None:
        return None
    ext = os.path.splitext(uploaded_file.name)[1] or '.jpg'
    filename = f"{prefix}{ext}"
    path = os.path.join(UPLOAD_DIR, filename)
    with open(path, 'wb') as f:
        f.write(uploaded_file.getbuffer())
    return path


def show_photo(path, caption=''):
    if path and os.path.exists(path):
        st.image(path, caption=caption, width=140)
    else:
        st.caption("📷 Fotoğraf yok")


def render_chip(label, value, bucket):
    st.markdown(f'''<div class="pk-chip {bucket}">
        <div class="pk-chip-label">{label}</div>
        <div class="pk-chip-value">{value}</div>
        </div>''', unsafe_allow_html=True)


def order_chip_row(o):
    """Toplam Adet / Toplam Kutu / Kumaş / Lastik kutucukları - tamamlanma & termin durumuna göre renkli."""
    order_bucket = urgency_bucket(o.get('deadline'))

    fab_done = all(f['kumas_status'] == MATERIAL_STATUS_DONE for f in o['fabrics']) if o['fabrics'] else False
    fab_terms = [f.get('kumas_termin_tarihi') for f in o['fabrics'] if f.get('kumas_termin_tarihi')]
    fab_bucket = 'green' if fab_done else (urgency_bucket(min(fab_terms)) if fab_terms else 'gray')

    el_done = all(e['lastik_status'] == MATERIAL_STATUS_DONE for e in o['elastics']) if o['elastics'] else False
    el_terms = [e.get('lastik_termin_tarihi') for e in o['elastics'] if e.get('lastik_termin_tarihi')]
    el_bucket = 'green' if el_done else (urgency_bucket(min(el_terms)) if el_terms else 'gray')

    html = '<div class="pk-chip-row">'
    for label, value, bucket in [
        ('Toplam Adet', f"{o['total_quantity']:,}", order_bucket),
        ('Toplam Kutu', f"{o['total_boxes']:,}", order_bucket),
        ('Kumaş (kg)', f"{o['kumas_kg_total']}", fab_bucket),
        ('Lastik', f"{o['lastik_cm_total']} cm", el_bucket),
    ]:
        html += f'''<div class="pk-chip {bucket}">
            <div class="pk-chip-label">{label}</div>
            <div class="pk-chip-value">{value}</div>
            </div>'''
    html += '</div>'
    st.markdown(html, unsafe_allow_html=True)


def render_fabric_tracking_table(fabrics):
    rows_html = ""
    for f in fabrics:
        bg = fabric_row_color(f.get('kumas_termin_tarihi'))
        rows_html += f"""<tr style="background:{bg};">
            <td>{f.get('kumas_adi') or '-'}</td>
            <td>{f.get('kumas_renk') or '-'}</td>
            <td>{f['kumas_kg_required']}</td>
            <td>{f.get('kumas_siparis_kg') or 0}</td>
            <td>{f.get('kumas_gelen_kg') or 0}</td>
            <td>{f['kumas_kalan_kg']}</td>
            <td>{tr_date(f.get('kumas_siparis_tarihi'))}</td>
            <td>{tr_date(f.get('kumas_termin_tarihi'))}</td>
        </tr>"""
    html = f"""<table class="pk-track-table">
        <tr><th>Kumaş</th><th>Renk</th><th>Gerekli (kg)</th><th>Sipariş (kg)</th>
        <th>Gelen (kg)</th><th>Kalan (kg)</th><th>Sipariş Tarihi</th><th>Termin</th></tr>
        {rows_html}
    </table>"""
    st.markdown(html, unsafe_allow_html=True)


# ======================================================================
# GİRİŞ SAYFASI
# ======================================================================
def page_login():
    st.markdown("<br><br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.markdown(
            '<div style="text-align:center; margin-bottom:8px;">'
            '<div style="font-size:2.2rem;">🧥</div>'
            '<div style="font-size:1.6rem; font-weight:800; color:#0f172a;">PAUL KENZIE</div>'
            '<div style="color:#64748b; letter-spacing:2px; font-size:0.75rem; text-transform:uppercase;">Üretim Yönetim Sistemi (ERP)</div>'
            '</div>', unsafe_allow_html=True)
        st.markdown('<div class="pk-card">', unsafe_allow_html=True)
        with st.form("login_form"):
            username = st.text_input("Kullanıcı Adı")
            password = st.text_input("Şifre", type="password")
            submitted = st.form_submit_button("🔐 Giriş Yap", width='stretch', type="primary")
            if submitted:
                if username and password:
                    user = db.authenticate(username, password)
                    if user:
                        st.session_state.logged_in = True
                        st.session_state.user = user
                        db.add_log(user['id'], user['username'], 'Giriş', f'{user["full_name"]} sisteme giriş yaptı')
                        st.rerun()
                    else:
                        st.error("Kullanıcı adı veya şifre hatalı!")
                else:
                    st.error("Lütfen tüm alanları doldurun.")
        st.markdown('</div>', unsafe_allow_html=True)


# ======================================================================
# DASHBOARD
# ======================================================================
def page_dashboard():
    page_header("📊 Genel Bakış", "Sipariş durumlarının özet görünümü")

    counts = db.get_dashboard_counts()
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Aktif Sipariş", counts['active'])
    c2.metric("🏭 Atanmış", counts['assigned'])
    c3.metric("🗂️ Planlama Bekleyen", counts['unplanned'])
    c4.metric("🔴 Acil / Gecikmiş", counts['urgent'])
    c5.metric("Toplam Sipariş", counts['total'])

    st.markdown('<div class="pk-section-title">✅ Tamamlanmış Siparişler</div>', unsafe_allow_html=True)
    this_year = date.today().year
    last_year = this_year - 1
    cc1, cc2 = st.columns(2)
    cc1.metric(f"{this_year}", db.get_completed_count_by_year(this_year))
    cc2.metric(f"{last_year}", db.get_completed_count_by_year(last_year))


# ======================================================================
# YENİ SİPARİŞ SİHİRBAZI
# ======================================================================
STEP_LABELS = ["Temel Bilgiler", "Sipariş Adedi", "Kumaş / Lastik / Aksesuar", "Onay"]
NEW_PN_SENTINEL = "➕ Yeni Ürün Adı..."


def render_steps(active):
    html = '<div class="pk-steps">'
    for i, label in enumerate(STEP_LABELS, 1):
        cls = 'done' if i < active else ('active' if i == active else '')
        icon = '✓ ' if i < active else ''
        html += f'<div class="pk-step {cls}">{icon}{i}. {label}</div>'
    html += '</div>'
    st.markdown(html, unsafe_allow_html=True)


def reset_wizard():
    keys_to_clear = [k for k in st.session_state.keys() if k.startswith(('size_box_', 'w_', 'fab_', 'el_', 'ak_', '_pn_applied'))]
    for k in keys_to_clear:
        del st.session_state[k]
    st.session_state.order_step = 1
    st.session_state.order_data = {}
    st.session_state.wiz_fabric_count = 1
    st.session_state.wiz_elastic_count = 1
    st.session_state.wiz_accessory_count = 0


def _on_pn_change():
    sel = st.session_state.get('w_urun_adi_sel')
    if sel and sel != NEW_PN_SENTINEL:
        rec = db.get_product_name_by_name(sel)
        if rec:
            st.session_state['w_gender'] = rec['gender']
            groups = [g['name'] for g in db.get_product_groups(rec['gender'])]
            if rec['urun_grubu'] in groups:
                st.session_state['w_urun_grubu'] = rec['urun_grubu']


def compute_preview_totals(sizes, kumas_fire, lastik_fire):
    kumas_fire_mult = 1 + (kumas_fire or 0) / 100.0
    lastik_fire_mult = 1 + (lastik_fire or 0) / 100.0
    kumas_gr_total = sum(s['box_qty'] * s['kumas_gr'] for s in sizes) * kumas_fire_mult
    lastik_adet_total = sum(s['box_qty'] * s['lastik_adet'] for s in sizes) * lastik_fire_mult
    lastik_cm_total = sum(s['box_qty'] * s['lastik_cm'] for s in sizes) * lastik_fire_mult
    return {
        'kumas_kg_total': round(kumas_gr_total / 1000.0, 3),
        'lastik_adet_total': round(lastik_adet_total, 1),
        'lastik_cm_total': round(lastik_cm_total, 1),
    }


def page_yeni_siparis():
    page_header("📝 Yeni Sipariş", "Ürün, sipariş adedi ve malzeme bilgilerini girin")

    if 'order_step' not in st.session_state:
        st.session_state.order_step = 1
    if 'order_data' not in st.session_state:
        st.session_state.order_data = {}
    if 'wiz_fabric_count' not in st.session_state:
        st.session_state.wiz_fabric_count = 1
    if 'wiz_elastic_count' not in st.session_state:
        st.session_state.wiz_elastic_count = 1
    if 'wiz_accessory_count' not in st.session_state:
        st.session_state.wiz_accessory_count = 0

    step = st.session_state.order_step
    render_steps(step)

    # ---------------- STEP 1: Temel Bilgiler ----------------
    if step == 1:
        product_names_all = db.get_product_names()
        pn_display = [NEW_PN_SENTINEL] + [p['name'] for p in product_names_all]

        c1, c2 = st.columns(2)
        with c1:
            pn_sel = st.selectbox("ÜRÜN ADI *", pn_display, key="w_urun_adi_sel", on_change=_on_pn_change)
            if pn_sel == NEW_PN_SENTINEL:
                new_model_name = st.text_input("Yeni Ürün Adı", key="w_new_pn_name")
                save_new_pn = st.checkbox("Bu ürün adını kaydet (tekrar kullanmak için)", value=True, key="w_save_pn")
                model_name = new_model_name
            else:
                model_name = pn_sel
                save_new_pn = False

            gender = st.radio("CİNSİYET *", GENDERS, horizontal=True, key="w_gender")
            groups = db.get_product_groups(gender)
            if not groups:
                st.warning(f"'{gender}' için Ayarlar > Ürün Grupları bölümünden ürün grubu tanımlamalısınız.")
                urun_grubu = None
            else:
                group_names = [g['name'] for g in groups]
                if st.session_state.get('w_urun_grubu') not in group_names:
                    st.session_state['w_urun_grubu'] = group_names[0]
                urun_grubu = st.selectbox("ÜRÜN GRUBU *", group_names, key="w_urun_grubu")
            package_size = st.selectbox("KUTU İÇİ ADEDİ *", PACKAGE_SIZES,
                                         format_func=lambda x: f"{x}'li paket", key="w_package_size")
        with c2:
            siparis_tarihi = st.date_input("SİPARİŞ TARİHİ *", value=date.today(), format=TR_DATE_FMT, key="w_siparis_tarihi")
            deadline = st.date_input("TERMİN TARİHİ *", min_value=date.today(), format=TR_DATE_FMT, key="w_deadline")
            kumas_fire = st.number_input("KUMAŞ FİRE ORANI (%)", min_value=0.0, max_value=50.0, value=3.0, step=0.5,
                                          key="w_kumas_fire")
            lastik_fire = st.number_input("LASTİK FİRE ORANI (%)", min_value=0.0, max_value=50.0, value=3.0, step=0.5,
                                           key="w_lastik_fire")
            urun_foto = st.file_uploader("ÜRÜN FOTOĞRAFI (opsiyonel)", type=['png', 'jpg', 'jpeg'], key="w_urun_foto")

        manufacturers = db.get_manufacturers()
        st.markdown("**📦 Kutu**")
        names, ids = mf_options(manufacturers, "— Belirtilmedi —")
        kutu_mf_sel = st.selectbox("Kutu Üreticisi (opsiyonel)", names, key="w_kutu_mf")
        kutu_manufacturer_id = ids[names.index(kutu_mf_sel)]

        if st.button("Devam Et →", width='stretch', type="primary"):
            if not model_name or not model_name.strip():
                st.error("Ürün adı boş olamaz!")
            elif not urun_grubu:
                st.error("Ürün grubu seçilmedi!")
            else:
                if pn_sel == NEW_PN_SENTINEL and save_new_pn:
                    db.add_product_name(model_name.strip(), gender, urun_grubu)
                st.session_state.order_data.update({
                    'model_name': model_name.strip(),
                    'gender': gender,
                    'urun_grubu': urun_grubu,
                    'package_size': package_size,
                    'siparis_tarihi': siparis_tarihi.isoformat(),
                    'deadline': deadline.isoformat(),
                    'kumas_fire_orani': kumas_fire,
                    'lastik_fire_orani': lastik_fire,
                    '_urun_foto_file': urun_foto,
                    'kutu_manufacturer_id': kutu_manufacturer_id,
                })
                st.session_state.order_step = 2
                st.rerun()

    # ---------------- STEP 2: Sipariş Adedi ----------------
    elif step == 2:
        data = st.session_state.order_data
        st.markdown(f"**Ürün:** {data['model_name']}  |  **Cinsiyet:** {data['gender']}  |  "
                    f"**Ürün Grubu:** {data['urun_grubu']}  |  **Paket:** {data['package_size']}'li")

        recipe = db.get_recipe(data['gender'], data['urun_grubu'])
        has_recipe = any(recipe[s]['kumas_gr'] > 0 or recipe[s]['lastik_cm'] > 0 for s in SIZES)
        if not has_recipe:
            st.warning("Bu ürün grubu için Ayarlar > Reçeteler bölümünde henüz kumaş/lastik reçetesi tanımlanmamış. "
                       "Kutu adetlerini girebilirsiniz ama kumaş/lastik ihtiyacı 0 hesaplanacaktır.")

        st.markdown("#### Sipariş Adedi (Beden Bazlı Kutu Sayısı)")
        cols = st.columns(7)
        for i, size in enumerate(SIZES):
            with cols[i]:
                st.number_input(size, min_value=0, value=0, step=1, key=f"size_box_{size}", label_visibility="visible")
                r = recipe[size]
                st.caption(f"K:{r['kumas_gr']}gr L:{r['lastik_cm']}cm")

        active_sizes = [s for s in SIZES if st.session_state.get(f"size_box_{s}", 0) > 0]

        st.markdown("---")
        cb1, cb2 = st.columns(2)
        with cb1:
            if st.button("← Geri", width='stretch'):
                st.session_state.order_step = 1
                st.rerun()
        with cb2:
            if st.button("Devam Et →", width='stretch', type="primary"):
                if not active_sizes:
                    st.error("En az bir beden için kutu adedi girin!")
                else:
                    sizes_payload = []
                    for size in active_sizes:
                        r = recipe[size]
                        sizes_payload.append({
                            'size': size,
                            'box_qty': int(st.session_state.get(f"size_box_{size}", 0)),
                            'kumas_gr': r['kumas_gr'],
                            'lastik_adet': r['lastik_adet'],
                            'lastik_cm': r['lastik_cm'],
                        })
                    st.session_state.order_data['sizes'] = sizes_payload
                    st.session_state.order_data['total_boxes'] = sum(s['box_qty'] for s in sizes_payload)
                    st.session_state.order_step = 3
                    st.rerun()

    # ---------------- STEP 3: Kumaş / Lastik / Aksesuar (bağımsız, dinamik satırlar) ----------------
    elif step == 3:
        data = st.session_state.order_data
        fabrics_master = db.get_fabrics()
        elastics_master = db.get_elastics()
        preview = compute_preview_totals(data['sizes'], data['kumas_fire_orani'], data['lastik_fire_orani'])

        fabric_names = ["— Tanımlı değil —"] + [f"{f['name']} ({f['kumas_turu']})" for f in fabrics_master]
        fabric_ids = [None] + [f['id'] for f in fabrics_master]
        elastic_labels = ["— Tanımlı değil —"] + [f"{e['tur']} ({e['boyut']})" for e in elastics_master]
        elastic_ids = [None] + [e['id'] for e in elastics_master]

        st.markdown("### 🧵 Kumaş")
        n_fab = st.session_state.wiz_fabric_count
        for i in range(1, n_fab + 1):
            c1, c2, c3 = st.columns(3)
            with c1:
                st.selectbox("Kumaş", fabric_names, key=f"fab_sel_{i}")
            with c2:
                st.text_input("Renk", key=f"fab_renk_{i}")
            with c3:
                st.file_uploader("Fotoğraf", type=['png', 'jpg', 'jpeg'], key=f"fab_foto_{i}")
        st.caption(f"Bu {n_fab} satıra toplam kumaş ihtiyacı ({preview['kumas_kg_total']} kg) eşit bölünecektir.")
        if st.button("➕ Kumaş Satırı Ekle", key="add_fab_row"):
            st.session_state.wiz_fabric_count += 1
            st.rerun()
        if n_fab > 1 and st.button("➖ Son Kumaş Satırını Kaldır", key="rm_fab_row"):
            for k in [f"fab_sel_{n_fab}", f"fab_renk_{n_fab}", f"fab_foto_{n_fab}"]:
                st.session_state.pop(k, None)
            st.session_state.wiz_fabric_count -= 1
            st.rerun()

        st.markdown("---")
        st.markdown("### ➰ Lastik")
        n_el = st.session_state.wiz_elastic_count
        for i in range(1, n_el + 1):
            c1, c2, c3 = st.columns(3)
            with c1:
                esel = st.selectbox("Lastik", elastic_labels, key=f"el_sel_{i}")
            with c2:
                st.text_input("Renk", key=f"el_renk_{i}")
            with c3:
                st.file_uploader("Fotoğraf", type=['png', 'jpg', 'jpeg'], key=f"el_foto_{i}")
            sel_idx = elastic_labels.index(esel) if esel in elastic_labels else 0
            sel_elastic_id = elastic_ids[sel_idx]
            sel_elastic = next((e for e in elastics_master if e['id'] == sel_elastic_id), None)
            if sel_elastic and sel_elastic['tur'] == RAPORLU_LASTIK:
                per_row_adet = round(preview['lastik_adet_total'] / n_el, 1)
                st.caption(f"📏 Raporlu lastik seçildi — bu satır ADET bazlı takip edilecek (≈{per_row_adet} adet).")
            else:
                per_row_cm = round(preview['lastik_cm_total'] / n_el, 1)
                st.caption(f"📏 Bu satır CM bazlı takip edilecek (≈{per_row_cm} cm).")
        if st.button("➕ Lastik Satırı Ekle", key="add_el_row"):
            st.session_state.wiz_elastic_count += 1
            st.rerun()
        if n_el > 1 and st.button("➖ Son Lastik Satırını Kaldır", key="rm_el_row"):
            for k in [f"el_sel_{n_el}", f"el_renk_{n_el}", f"el_foto_{n_el}"]:
                st.session_state.pop(k, None)
            st.session_state.wiz_elastic_count -= 1
            st.rerun()

        st.markdown("---")
        st.markdown("### 🔘 Aksesuar (opsiyonel)")
        n_ak = st.session_state.wiz_accessory_count
        for i in range(1, n_ak + 1):
            c1, c2, c3 = st.columns(3)
            with c1:
                st.text_input("Aksesuar Adı", key=f"ak_adi_{i}")
            with c2:
                st.text_input("Renk", key=f"ak_renk_{i}")
            with c3:
                st.file_uploader("Fotoğraf", type=['png', 'jpg', 'jpeg'], key=f"ak_foto_{i}")
        if st.button("➕ Aksesuar Satırı Ekle", key="add_ak_row"):
            st.session_state.wiz_accessory_count += 1
            st.rerun()
        if n_ak > 0 and st.button("➖ Son Aksesuar Satırını Kaldır", key="rm_ak_row"):
            for k in [f"ak_adi_{n_ak}", f"ak_renk_{n_ak}", f"ak_foto_{n_ak}"]:
                st.session_state.pop(k, None)
            st.session_state.wiz_accessory_count -= 1
            st.rerun()

        st.markdown("---")
        cb1, cb2 = st.columns(2)
        with cb1:
            if st.button("← Geri", width='stretch', key="step3_back"):
                st.session_state.order_step = 2
                st.rerun()
        with cb2:
            if st.button("Devam Et →", width='stretch', type="primary", key="step3_next"):
                fabrics_in = []
                for i in range(1, n_fab + 1):
                    sel = st.session_state.get(f"fab_sel_{i}")
                    idx = fabric_names.index(sel) if sel in fabric_names else 0
                    fabrics_in.append({
                        'kumas_fabric_id': fabric_ids[idx],
                        'kumas_renk': st.session_state.get(f"fab_renk_{i}", ''),
                        '_foto_file': st.session_state.get(f"fab_foto_{i}"),
                    })
                elastics_in = []
                for i in range(1, n_el + 1):
                    sel = st.session_state.get(f"el_sel_{i}")
                    idx = elastic_labels.index(sel) if sel in elastic_labels else 0
                    elastics_in.append({
                        'lastik_elastic_id': elastic_ids[idx],
                        'lastik_renk': st.session_state.get(f"el_renk_{i}", ''),
                        '_foto_file': st.session_state.get(f"el_foto_{i}"),
                    })
                accessories_in = []
                for i in range(1, n_ak + 1):
                    accessories_in.append({
                        'aksesuar_adi': st.session_state.get(f"ak_adi_{i}", ''),
                        'aksesuar_renk': st.session_state.get(f"ak_renk_{i}", ''),
                        '_foto_file': st.session_state.get(f"ak_foto_{i}"),
                    })
                st.session_state.order_data['fabrics'] = fabrics_in
                st.session_state.order_data['elastics'] = elastics_in
                st.session_state.order_data['accessories'] = accessories_in
                st.session_state.order_step = 4
                st.rerun()

    # ---------------- STEP 4: Onay ----------------
    elif step == 4:
        data = st.session_state.order_data
        st.markdown("#### Sipariş Özeti")

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Ürün", data['model_name'])
        c2.metric("Ürün Grubu", data['urun_grubu'])
        c3.metric("Paket", f"{data['package_size']}'li")
        c4.metric("Toplam Kutu", data.get('total_boxes', 0))

        total_qty = data.get('total_boxes', 0) * data['package_size']
        st.markdown(f"**Cinsiyet:** {data['gender']}  |  **Sipariş Tarihi:** {tr_date(data['siparis_tarihi'])}  |  "
                    f"**Termin:** {tr_date(data['deadline'])}  |  **Toplam Adet:** {total_qty}  |  "
                    f"**Kumaş Fire:** %{data['kumas_fire_orani']}  |  **Lastik Fire:** %{data['lastik_fire_orani']}")

        st.markdown("---")
        st.markdown("**Beden Dağılımı**")
        st.dataframe(pd.DataFrame([
            {'Beden': s['size'], 'Kutu Adedi': s['box_qty'], 'Kumaş (gr/parça)': s['kumas_gr'],
             'Lastik (adet/parça)': s['lastik_adet'], 'Lastik (cm/parça)': s['lastik_cm']}
            for s in data['sizes']
        ]), width='stretch', hide_index=True)

        preview = compute_preview_totals(data['sizes'], data['kumas_fire_orani'], data['lastik_fire_orani'])
        fabrics_master = {f['id']: f for f in db.get_fabrics()}
        elastics_master = {e['id']: e for e in db.get_elastics()}

        st.markdown("---")
        st.markdown(f"**🧵 Kumaş** — Toplam ihtiyaç: {preview['kumas_kg_total']} kg")
        fab_rows = []
        n_fab = len(data['fabrics'])
        for f in data['fabrics']:
            fab = fabrics_master.get(f['kumas_fabric_id'])
            fab_rows.append({'Kumaş': fab['name'] if fab else '—', 'Renk': f['kumas_renk'] or '—',
                             'Gerekli (kg)': round(preview['kumas_kg_total'] / n_fab, 3)})
        st.dataframe(pd.DataFrame(fab_rows), width='stretch', hide_index=True)

        st.markdown(f"**➰ Lastik** — Toplam ihtiyaç: {preview['lastik_cm_total']} cm / {preview['lastik_adet_total']} adet (Raporlu)")
        el_rows = []
        n_el = len(data['elastics'])
        for el in data['elastics']:
            ela = elastics_master.get(el['lastik_elastic_id'])
            is_raporlu = ela and ela['tur'] == RAPORLU_LASTIK
            gerekli = round(preview['lastik_adet_total'] / n_el, 1) if is_raporlu else round(preview['lastik_cm_total'] / n_el, 1)
            birim = 'adet' if is_raporlu else 'cm'
            el_rows.append({'Lastik': ela['tur'] if ela else '—', 'Renk': el['lastik_renk'] or '—',
                            'Gerekli': f"{gerekli} {birim}"})
        st.dataframe(pd.DataFrame(el_rows), width='stretch', hide_index=True)

        if data['accessories']:
            st.markdown(f"**🔘 Aksesuar**")
            ak_rows = [{'Aksesuar': a['aksesuar_adi'] or '—', 'Renk': a['aksesuar_renk'] or '—'} for a in data['accessories']]
            st.dataframe(pd.DataFrame(ak_rows), width='stretch', hide_index=True)

        st.markdown("---")
        cb1, cb2 = st.columns(2)
        with cb1:
            if st.button("← Geri", width='stretch'):
                st.session_state.order_step = 3
                st.rerun()
        with cb2:
            if st.button("✅ Siparişi Oluştur", width='stretch', type="primary"):
                order_id, fab_ids, el_ids, ak_ids = db.add_order(
                    data['model_name'], data['gender'], data['urun_grubu'], data['package_size'],
                    data['kumas_fire_orani'], data['lastik_fire_orani'],
                    data['deadline'], data['siparis_tarihi'], data['sizes'],
                    data['fabrics'], data['elastics'], data['accessories'], st.session_state.user['id']
                )

                if data.get('_urun_foto_file') is not None:
                    path = save_uploaded_file(data['_urun_foto_file'], f"order_{order_id}_urun")
                    db.set_order_photo(order_id, path)

                for i, f in enumerate(data['fabrics']):
                    if f.get('_foto_file'):
                        path = save_uploaded_file(f['_foto_file'], f"order_{order_id}_fabric_{fab_ids[i]}")
                        db.set_fabric_photo(fab_ids[i], path)
                for i, el in enumerate(data['elastics']):
                    if el.get('_foto_file'):
                        path = save_uploaded_file(el['_foto_file'], f"order_{order_id}_elastic_{el_ids[i]}")
                        db.set_elastic_photo(el_ids[i], path)
                for i, a in enumerate(data['accessories']):
                    if a.get('_foto_file'):
                        path = save_uploaded_file(a['_foto_file'], f"order_{order_id}_accessory_{ak_ids[i]}")
                        db.set_accessory_photo(ak_ids[i], path)

                if data.get('kutu_manufacturer_id'):
                    db.update_order_kutu(order_id, data['kutu_manufacturer_id'], 0, 0, '', '')

                db.add_log(st.session_state.user['id'], st.session_state.user['username'],
                          'Yeni Sipariş', f'{data["model_name"]} siparişi oluşturuldu (ID: {order_id})')
                st.success("Sipariş başarıyla oluşturuldu! 'Planlama Bekleyen Siparişler' listesinde görünecek.")
                reset_wizard()
                st.rerun()


# ======================================================================
# SİPARİŞ DETAY GÖRÜNÜMÜ (salt okunur)
# ======================================================================
def render_order_detail_readonly(o):
    order_chip_row(o)

    if o.get('urun_foto'):
        show_photo(o['urun_foto'], "Ürün Fotoğrafı")

    st.caption(f"Sipariş Tarihi: {tr_date(o.get('siparis_tarihi'))}  |  Termin: {tr_date(o.get('deadline'))}")

    st.markdown("**Beden Dağılımı**")
    st.dataframe(pd.DataFrame([
        {'Beden': s['size'], 'Kutu': s['box_qty'], 'Kumaş (gr/parça)': s['kumas_gr'],
         'Lastik (adet/parça)': s['lastik_adet'], 'Lastik (cm/parça)': s['lastik_cm']}
        for s in o['sizes']
    ]), width='stretch', hide_index=True)

    st.markdown("**🧵 Kumaş**")
    for f in o['fabrics']:
        cc1, cc2 = st.columns([3, 1])
        with cc1:
            st.markdown(
                f"{material_badge(f['kumas_status'])} **{f.get('kumas_adi') or '-'}** ({f.get('kumas_turu') or '-'}) · "
                f"{f.get('kumas_icerik') or '-'} · Renk: {f.get('kumas_renk') or '-'} · "
                f"Üretici: {f.get('kumas_manufacturer_name') or '—'} · Gerekli: {f['kumas_kg_required']} kg",
                unsafe_allow_html=True)
        with cc2:
            show_photo(f.get('kumas_foto'))
    render_fabric_tracking_table(o['fabrics'])

    st.markdown("**➰ Lastik**")
    for e in o['elastics']:
        birim = 'adet' if e['is_raporlu'] else 'cm'
        gerekli = e['lastik_adet_required'] if e['is_raporlu'] else e['lastik_cm_required']
        gelen = e.get('lastik_gelen_adet') if e['is_raporlu'] else e.get('lastik_gelen_cm')
        cc1, cc2 = st.columns([3, 1])
        with cc1:
            st.markdown(
                f"{material_badge(e['lastik_status'])} **{e.get('lastik_tur') or '-'}** ({e.get('lastik_boyut') or '-'}) · "
                f"Renk: {e.get('lastik_renk') or '-'} · Üretici: {e.get('lastik_manufacturer_name') or '—'} · "
                f"Gerekli: {gerekli} {birim} · Gelen: {gelen or 0} {birim}",
                unsafe_allow_html=True)
        with cc2:
            show_photo(e.get('lastik_foto'))

    if o['accessories']:
        st.markdown("**🔘 Aksesuar**")
        for a in o['accessories']:
            cc1, cc2 = st.columns([3, 1])
            with cc1:
                st.markdown(
                    f"{material_badge(a['aksesuar_status'])} **{a['aksesuar_adi']}** · Renk: {a.get('aksesuar_renk') or '-'} · "
                    f"Üretici: {a.get('aksesuar_manufacturer_name') or '—'} · "
                    f"Gerekli: {a['aksesuar_adet_required']} adet · Gelen: {a.get('aksesuar_gelen_adet') or 0} adet",
                    unsafe_allow_html=True)
            with cc2:
                show_photo(a.get('aksesuar_foto'))

    st.markdown("**📦 Kutu**")
    kutu_bg = fabric_row_color(o.get('kutu_termin_tarihi'))
    kutu_kalan = (o.get('kutu_siparis_adet') or 0) - (o.get('kutu_gelen_adet') or 0)
    st.markdown(f"""<div style="background:{kutu_bg}; padding:10px 14px; border-radius:8px; font-size:0.86rem;">
        Üretici: {o.get('kutu_manufacturer_name') or '—'} &nbsp;|&nbsp;
        Gerekli: {o.get('total_boxes', 0)} kutu &nbsp;|&nbsp;
        Sipariş: {o.get('kutu_siparis_adet') or 0} &nbsp;|&nbsp;
        Gelen: {o.get('kutu_gelen_adet') or 0} &nbsp;|&nbsp;
        Kalan: {kutu_kalan} &nbsp;|&nbsp;
        Termin: {tr_date(o.get('kutu_termin_tarihi'))}
        </div>""", unsafe_allow_html=True)


def order_list_page(title, subtitle, assigned_filter, is_admin, include_completed=False, year_filter=False):
    page_header(title, subtitle)

    if year_filter:
        this_year = date.today().year
        last_year = this_year - 1
        year = st.selectbox("Yıl", [this_year, last_year])
        orders = db.get_orders_overview(assigned=None, include_completed=True, year=year)
        orders = [o for o in orders if o['status'] == 'Tamamlandı']
    else:
        orders = db.get_orders_overview(assigned=assigned_filter, include_completed=include_completed)

    search = st.text_input("🔍 Ürün adına göre ara")
    if search:
        orders = [o for o in orders if search.lower() in o['model_name'].lower()]

    if not orders:
        st.info("Bu kritere uyan sipariş bulunamadı.")
        return

    for o in orders:
        title_line = (f"{o['model_name']} — {o['gender']} — {o.get('urun_grubu') or '-'} — "
                      f"{o['package_size']}'li — {o.get('manufacturer_name') or 'Atanmamış'} — "
                      f"{get_urgency_label(o.get('deadline',''))} — {o['status']}")
        with st.expander(title_line):
            render_order_detail_readonly(o)
            if is_admin:
                st.markdown("---")
                if st.button("✏️ Düzenle", key=f"editnav_{o['id']}"):
                    st.session_state.edit_order_id = o['id']
                    st.rerun()


def page_atanmis_siparisler():
    order_list_page("🏭 Atanmış Siparişler", "Üreticiye atanmış aktif siparişler",
                    assigned_filter=True, is_admin=(st.session_state.user['role'] == 'admin'))


def page_planlama_bekleyen():
    order_list_page("🗂️ Planlama Bekleyen Siparişler", "Henüz üretici ataması yapılmamış siparişler",
                    assigned_filter=False, is_admin=(st.session_state.user['role'] == 'admin'))


def page_tamamlanmis_siparisler():
    order_list_page("✅ Tamamlanmış Siparişler", "Yıla göre tamamlanmış sipariş arşivi",
                    assigned_filter=None, is_admin=(st.session_state.user['role'] == 'admin'),
                    include_completed=True, year_filter=True)


# ======================================================================
# İRSALİYE / PDF
# ======================================================================
def generate_irsaliye_pdf(irsaliye, order):
    if not HAS_PDF:
        return None
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4
    c.setFont("Helvetica-Bold", 16)
    c.drawString(20 * mm, height - 20 * mm, "İRSALİYE - Paul Kenzie ERP")
    c.setFont("Helvetica", 10)
    y = height - 35 * mm
    c.drawString(20 * mm, y, f"İrsaliye No: {irsaliye.get('irsaliye_no', '-')}")
    y -= 7 * mm
    c.drawString(20 * mm, y, f"Geliş Tarihi: {tr_date(irsaliye.get('gelis_tarihi'))}")
    y -= 7 * mm
    c.drawString(20 * mm, y, f"Kategori: {irsaliye.get('kategori', '-')} | Ürün: {order.get('model_name','-')}")
    y -= 7 * mm
    c.drawString(20 * mm, y, f"Tedarikçi: {irsaliye.get('tedarikci', '-')} | "
                             f"Miktar: {irsaliye.get('miktar', 0)} {irsaliye.get('birim', '')}")
    y -= 10 * mm
    c.setFont("Helvetica", 9)
    if irsaliye.get('aciklama'):
        c.drawString(20 * mm, y, f"Açıklama: {irsaliye['aciklama']}")
        y -= 10 * mm
    c.setFont("Helvetica", 8)
    c.drawString(20 * mm, y, f"Oluşturulma: {datetime.now().strftime('%d.%m.%Y %H:%M')} | Sistem: Paul Kenzie ERP")
    c.showPage()
    c.save()
    buffer.seek(0)
    return buffer.getvalue()


def render_irsaliye_form(order_id, item_id, kategori, tedarikci_default, birim):
    key_prefix = f"irs_{kategori}_{order_id}_{item_id}"
    with st.form(f"form_{key_prefix}"):
        c1, c2, c3 = st.columns(3)
        with c1:
            irsaliye_no = st.text_input("İrsaliye No", value=f"IRS-{datetime.now().strftime('%Y%m%d')}-{order_id}",
                                        key=f"no_{key_prefix}")
            tedarikci = st.text_input("Tedarikçi", value=tedarikci_default, key=f"ted_{key_prefix}")
        with c2:
            miktar = st.number_input(f"Miktar ({birim})", min_value=0.0, value=0.0, key=f"mik_{key_prefix}")
            gelis_tarihi = st.date_input("Geliş Tarihi", value=date.today(), format=TR_DATE_FMT, key=f"gt_{key_prefix}")
        with c3:
            aciklama = st.text_area("Açıklama", key=f"ac_{key_prefix}")
        if st.form_submit_button("💾 İrsaliyeyi Kaydet (kalan miktar otomatik düşer)", type="primary"):
            irs_id = db.add_irsaliye(order_id, item_id, kategori, irsaliye_no, tedarikci, miktar, birim,
                                     gelis_tarihi.isoformat(), aciklama)
            db.add_log(st.session_state.user['id'], st.session_state.user['username'],
                      'İrsaliye Oluşturuldu', f"{kategori} - {irsaliye_no}")
            st.session_state[f'last_irs_{key_prefix}'] = irs_id
            st.success("İrsaliye kaydedildi, gelen/kalan miktar güncellendi!")
            st.rerun()

    last_irs_id = st.session_state.get(f'last_irs_{key_prefix}')
    if last_irs_id:
        irsaliyeler = db.get_irsaliyeler(order_id)
        irs = next((i for i in irsaliyeler if i['id'] == last_irs_id), None)
        if irs and HAS_PDF:
            order = db.get_order_detail(order_id)
            pdf_bytes = generate_irsaliye_pdf(irs, order)
            if pdf_bytes:
                st.download_button(f"📄 PDF İndir - {irs['irsaliye_no']}", data=pdf_bytes,
                                  file_name=f"{irs['irsaliye_no']}.pdf", mime="application/pdf",
                                  key=f"dl_{key_prefix}_{last_irs_id}")
        elif irs and not HAS_PDF:
            st.warning("PDF kütüphanesi (reportlab) kurulu değil.")


# ======================================================================
# SİPARİŞ DÜZENLE (ayrı sayfa, sadece admin)
# ======================================================================
def page_order_edit(order_id):
    o = db.get_order_detail(order_id)
    if not o:
        st.error("Sipariş bulunamadı.")
        if st.button("← Listeye Dön"):
            st.session_state.edit_order_id = None
            st.rerun()
        return

    if st.button("← Listeye Dön"):
        st.session_state.edit_order_id = None
        st.rerun()

    page_header(f"✏️ Sipariş Düzenle — {o['model_name']}", f"ID: {o['id']}")

    manufacturers = db.get_manufacturers()
    fabrics_master = db.get_fabrics()
    elastics_master = db.get_elastics()
    fabric_labels = ["— Tanımlı değil —"] + [f"{f['name']} ({f['kumas_turu']})" for f in fabrics_master]
    fabric_ids_m = [None] + [f['id'] for f in fabrics_master]
    elastic_labels = ["— Tanımlı değil —"] + [f"{e['tur']} ({e['boyut']})" for e in elastics_master]
    elastic_ids_m = [None] + [e['id'] for e in elastics_master]

    # --- Temel bilgiler ---
    st.markdown('<div class="pk-section-title">Temel Bilgiler</div>', unsafe_allow_html=True)
    with st.form(f"edit_basic_{order_id}"):
        c1, c2 = st.columns(2)
        with c1:
            model_name = st.text_input("Ürün Adı", value=o['model_name'])
            gender = st.radio("Cinsiyet", GENDERS, index=GENDERS.index(o['gender']) if o['gender'] in GENDERS else 0,
                              horizontal=True)
            groups = db.get_product_groups(gender)
            group_names = [g['name'] for g in groups] or [o.get('urun_grubu') or '']
            ug_index = group_names.index(o['urun_grubu']) if o.get('urun_grubu') in group_names else 0
            urun_grubu = st.selectbox("Ürün Grubu", group_names, index=ug_index)
        with c2:
            try:
                st_val = datetime.strptime(o['siparis_tarihi'], '%Y-%m-%d').date() if o.get('siparis_tarihi') else date.today()
            except Exception:
                st_val = date.today()
            siparis_tarihi = st.date_input("Sipariş Tarihi", value=st_val, format=TR_DATE_FMT)
            try:
                dl_value = datetime.strptime(o['deadline'], '%Y-%m-%d').date() if o.get('deadline') else date.today()
            except Exception:
                dl_value = date.today()
            deadline = st.date_input("Termin Tarihi", value=dl_value, format=TR_DATE_FMT)
            kumas_fire = st.number_input("Kumaş Fire Oranı (%)", min_value=0.0, max_value=50.0,
                                          value=float(o['kumas_fire_orani']), step=0.5)
            lastik_fire = st.number_input("Lastik Fire Oranı (%)", min_value=0.0, max_value=50.0,
                                           value=float(o['lastik_fire_orani']), step=0.5)
        if st.form_submit_button("💾 Temel Bilgileri Kaydet", type="primary"):
            db.update_order_basic(order_id, model_name, gender, urun_grubu, deadline.isoformat(),
                                  siparis_tarihi.isoformat(), kumas_fire, lastik_fire)
            db.add_log(st.session_state.user['id'], st.session_state.user['username'],
                      'Sipariş Güncellendi', f"{model_name} temel bilgiler (ID: {order_id})")
            st.success("Kaydedildi!")
            st.rerun()

    # --- Üretici ataması ---
    st.markdown('<div class="pk-section-title">Üretici Ataması</div>', unsafe_allow_html=True)
    with st.form(f"edit_assignment_{order_id}"):
        names, ids = mf_options(manufacturers, "— Planlamada bırak (atama yapma) —")
        cur_idx = ids.index(o['manufacturer_id']) if o['manufacturer_id'] in ids else 0
        sel = st.selectbox("Üretici", names, index=cur_idx)
        new_mf_id = ids[names.index(sel)]
        full_service = st.checkbox("Bu üretici tüm süreçlerle (kumaş/lastik/aksesuar) ilgilenecek",
                                    value=(o['assignment_type'] == 'tam_hizmet'))
        new_status = st.selectbox("Üretim Durumu", ORDER_STATUSES, index=ORDER_STATUSES.index(o['status']))

        fabric_assignments, elastic_assignments, accessory_assignments = [], [], []
        if not full_service:
            if o['fabrics']:
                st.markdown("**Kumaş Üretici Ataması**")
                for f in o['fabrics']:
                    f_idx = ids.index(f['kumas_manufacturer_id']) if f['kumas_manufacturer_id'] in ids else 0
                    label = f"{f.get('kumas_adi') or 'Kumaş'} — {f.get('kumas_renk') or ''}"
                    f_sel = st.selectbox(label, names, index=f_idx, key=f"fmf_{order_id}_{f['id']}")
                    fabric_assignments.append({'id': f['id'], 'manufacturer_id': ids[names.index(f_sel)]})
            if o['elastics']:
                st.markdown("**Lastik Üretici Ataması**")
                for e in o['elastics']:
                    e_idx = ids.index(e['lastik_manufacturer_id']) if e['lastik_manufacturer_id'] in ids else 0
                    label = f"{e.get('lastik_tur') or 'Lastik'} — {e.get('lastik_renk') or ''}"
                    e_sel = st.selectbox(label, names, index=e_idx, key=f"emf_{order_id}_{e['id']}")
                    elastic_assignments.append({'id': e['id'], 'manufacturer_id': ids[names.index(e_sel)]})
            if o['accessories']:
                st.markdown("**Aksesuar Üretici Ataması**")
                for a in o['accessories']:
                    a_idx = ids.index(a['aksesuar_manufacturer_id']) if a['aksesuar_manufacturer_id'] in ids else 0
                    label = f"{a['aksesuar_adi']} — {a.get('aksesuar_renk') or ''}"
                    a_sel = st.selectbox(label, names, index=a_idx, key=f"amf_{order_id}_{a['id']}")
                    accessory_assignments.append({'id': a['id'], 'manufacturer_id': ids[names.index(a_sel)]})

        if st.form_submit_button("💾 Atamayı Kaydet", type="primary"):
            db.update_order_assignment(order_id, new_mf_id, full_service,
                                       fabric_assignments, elastic_assignments, accessory_assignments)
            db.update_order_status(order_id, new_status)
            db.add_log(st.session_state.user['id'], st.session_state.user['username'],
                      'Sipariş Ataması Güncellendi', f"{o['model_name']} (ID: {order_id})")
            st.success("Kaydedildi!")
            st.rerun()

    # --- Kutu takibi ---
    st.markdown('<div class="pk-section-title">📦 Kutu Takibi</div>', unsafe_allow_html=True)
    with st.form(f"edit_kutu_{order_id}"):
        names, ids = mf_options(manufacturers, "— Belirtilmedi —")
        k_idx = ids.index(o.get('kutu_manufacturer_id')) if o.get('kutu_manufacturer_id') in ids else 0
        kutu_mf_sel = st.selectbox("Kutu Üreticisi", names, index=k_idx)
        kutu_manufacturer_id = ids[names.index(kutu_mf_sel)]

        t1, t2, t3, t4 = st.columns(4)
        with t1:
            kutu_siparis_adet = st.number_input("Sipariş (adet)", min_value=0.0, value=float(o.get('kutu_siparis_adet') or 0))
        with t2:
            kutu_gelen_adet = st.number_input("Gelen (adet)", min_value=0.0, value=float(o.get('kutu_gelen_adet') or 0))
        with t3:
            try:
                ks_val = datetime.strptime(o['kutu_siparis_tarihi'], '%Y-%m-%d').date() if o.get('kutu_siparis_tarihi') else date.today()
            except Exception:
                ks_val = date.today()
            kutu_siparis_tarihi = st.date_input("Sipariş Tarihi", value=ks_val, format=TR_DATE_FMT, key="kutu_st")
        with t4:
            try:
                kt_val = datetime.strptime(o['kutu_termin_tarihi'], '%Y-%m-%d').date() if o.get('kutu_termin_tarihi') else date.today()
            except Exception:
                kt_val = date.today()
            kutu_termin_tarihi = st.date_input("Termin Tarihi", value=kt_val, format=TR_DATE_FMT, key="kutu_tt")

        if st.form_submit_button("💾 Kutu Bilgisini Kaydet", type="primary"):
            db.update_order_kutu(order_id, kutu_manufacturer_id, kutu_siparis_adet, kutu_gelen_adet,
                                 kutu_siparis_tarihi.isoformat(), kutu_termin_tarihi.isoformat())
            db.add_log(st.session_state.user['id'], st.session_state.user['username'],
                      'Kutu Güncellendi', f"{o['model_name']} (ID: {order_id})")
            st.success("Kaydedildi!")
            st.rerun()

    if st.button("🚚 Kutu İrsaliyesi Oluştur (PDF)", key=f"kutu_irs_{order_id}"):
        st.session_state[f'show_irs_kutu_{order_id}'] = True
    if st.session_state.get(f'show_irs_kutu_{order_id}'):
        render_irsaliye_form(order_id, None, "Kutu", o.get('kutu_manufacturer_name') or '', "adet")

    # --- Kumaş satırları ---
    st.markdown('<div class="pk-section-title">🧵 Kumaş</div>', unsafe_allow_html=True)
    render_fabric_tracking_table(o['fabrics'])
    for f in o['fabrics']:
        with st.expander(f"{f.get('kumas_adi') or 'Kumaş'} — {f.get('kumas_renk') or ''}"):
            with st.form(f"edit_fabric_{f['id']}"):
                r1, r2 = st.columns(2)
                with r1:
                    f_idx = fabric_ids_m.index(f.get('kumas_fabric_id')) if f.get('kumas_fabric_id') in fabric_ids_m else 0
                    fsel = st.selectbox("Kumaş", fabric_labels, index=f_idx, key=f"ef_{f['id']}")
                    new_fabric_id = fabric_ids_m[fabric_labels.index(fsel)]
                    kumas_renk = st.text_input("Renk", value=f.get('kumas_renk') or '', key=f"ekr_{f['id']}")
                with r2:
                    kumas_foto_file = st.file_uploader("Fotoğraf (değiştir)", type=['png', 'jpg', 'jpeg'], key=f"ekf_{f['id']}")
                    show_photo(f.get('kumas_foto'), "Mevcut")

                st.markdown("**Tedarik Takibi**")
                t1, t2, t3, t4 = st.columns(4)
                with t1:
                    siparis_kg = st.number_input("Sipariş (kg)", min_value=0.0, value=float(f.get('kumas_siparis_kg') or 0), key=f"esk_{f['id']}")
                with t2:
                    gelen_kg = st.number_input("Gelen (kg)", min_value=0.0, value=float(f.get('kumas_gelen_kg') or 0), key=f"egk_{f['id']}")
                with t3:
                    try:
                        st_val = datetime.strptime(f['kumas_siparis_tarihi'], '%Y-%m-%d').date() if f.get('kumas_siparis_tarihi') else date.today()
                    except Exception:
                        st_val = date.today()
                    siparis_tarihi = st.date_input("Sipariş T.", value=st_val, format=TR_DATE_FMT, key=f"est_{f['id']}")
                with t4:
                    try:
                        tt_val = datetime.strptime(f['kumas_termin_tarihi'], '%Y-%m-%d').date() if f.get('kumas_termin_tarihi') else date.today()
                    except Exception:
                        tt_val = date.today()
                    termin_tarihi = st.date_input("Termin", value=tt_val, format=TR_DATE_FMT, key=f"ett_{f['id']}")

                if not full_service:
                    names2, ids2 = mf_options(manufacturers, "— Atama yok —")
                    fm_idx = ids2.index(f.get('kumas_manufacturer_id')) if f.get('kumas_manufacturer_id') in ids2 else 0
                    fm_sel = st.selectbox("Üretici", names2, index=fm_idx, key=f"efmf_{f['id']}")
                    row_manufacturer_id = ids2[names2.index(fm_sel)]
                else:
                    row_manufacturer_id = f.get('kumas_manufacturer_id')
                    st.caption(f"Üretici (tam hizmet): {o.get('manufacturer_name') or '—'}")

                if st.form_submit_button("💾 Kaydet", type="primary"):
                    db.update_fabric_row(f['id'], new_fabric_id, kumas_renk, row_manufacturer_id,
                                         siparis_kg, gelen_kg, siparis_tarihi.isoformat(), termin_tarihi.isoformat())
                    if kumas_foto_file:
                        path = save_uploaded_file(kumas_foto_file, f"order_{order_id}_fabric_{f['id']}")
                        db.set_fabric_photo(f['id'], path)
                    db.add_log(st.session_state.user['id'], st.session_state.user['username'],
                              'Kumaş Güncellendi', f"{o['model_name']} (ID: {order_id})")
                    st.success("Kaydedildi!")
                    st.rerun()

            irs1, irs2 = st.columns(2)
            with irs1:
                if st.button("🚚 İrsaliye Oluştur", key=f"irs_fab_{f['id']}"):
                    st.session_state[f'show_irs_fab_{f["id"]}'] = True
            with irs2:
                if st.button("🗑️ Bu Kumaş Satırını Sil", key=f"del_fab_{f['id']}"):
                    db.delete_fabric_row(f['id'])
                    st.rerun()
            if st.session_state.get(f'show_irs_fab_{f["id"]}'):
                render_irsaliye_form(order_id, f['id'], "Kumaş", f.get('kumas_manufacturer_name') or '', "kg")

    if st.button("➕ Yeni Kumaş Satırı Ekle", key=f"add_fab_edit_{order_id}"):
        db.add_fabric_row(order_id)
        st.rerun()

    # --- Lastik satırları ---
    st.markdown('<div class="pk-section-title">➰ Lastik</div>', unsafe_allow_html=True)
    for e in o['elastics']:
        with st.expander(f"{e.get('lastik_tur') or 'Lastik'} — {e.get('lastik_renk') or ''}"):
            with st.form(f"edit_elastic_{e['id']}"):
                r1, r2 = st.columns(2)
                with r1:
                    e_idx = elastic_ids_m.index(e.get('lastik_elastic_id')) if e.get('lastik_elastic_id') in elastic_ids_m else 0
                    esel = st.selectbox("Lastik", elastic_labels, index=e_idx, key=f"ee_{e['id']}")
                    new_elastic_id = elastic_ids_m[elastic_labels.index(esel)]
                    lastik_renk = st.text_input("Renk", value=e.get('lastik_renk') or '', key=f"elr_{e['id']}")
                with r2:
                    lastik_foto_file = st.file_uploader("Fotoğraf (değiştir)", type=['png', 'jpg', 'jpeg'], key=f"elf_{e['id']}")
                    show_photo(e.get('lastik_foto'), "Mevcut")

                sel_elastic = next((el for el in elastics_master if el['id'] == new_elastic_id), None)
                is_raporlu = sel_elastic and sel_elastic['tur'] == RAPORLU_LASTIK

                st.markdown("**Tedarik Takibi**")
                if is_raporlu:
                    st.caption("📏 Raporlu lastik — adet bazlı takip")
                    t1, t2 = st.columns(2)
                    with t1:
                        siparis_adet = st.number_input("Sipariş (adet)", min_value=0.0, value=float(e.get('lastik_siparis_adet') or 0), key=f"elsa_{e['id']}")
                    with t2:
                        gelen_adet = st.number_input("Gelen (adet)", min_value=0.0, value=float(e.get('lastik_gelen_adet') or 0), key=f"elga_{e['id']}")
                    siparis_cm, gelen_cm = e.get('lastik_siparis_cm') or 0, e.get('lastik_gelen_cm') or 0
                else:
                    st.caption("📏 CM bazlı takip")
                    t1, t2 = st.columns(2)
                    with t1:
                        siparis_cm = st.number_input("Sipariş (cm)", min_value=0.0, value=float(e.get('lastik_siparis_cm') or 0), key=f"elsc_{e['id']}")
                    with t2:
                        gelen_cm = st.number_input("Gelen (cm)", min_value=0.0, value=float(e.get('lastik_gelen_cm') or 0), key=f"elgc_{e['id']}")
                    siparis_adet, gelen_adet = e.get('lastik_siparis_adet') or 0, e.get('lastik_gelen_adet') or 0

                t3, t4 = st.columns(2)
                with t3:
                    try:
                        ls_val = datetime.strptime(e['lastik_siparis_tarihi'], '%Y-%m-%d').date() if e.get('lastik_siparis_tarihi') else date.today()
                    except Exception:
                        ls_val = date.today()
                    lastik_siparis_tarihi = st.date_input("Sipariş T.", value=ls_val, format=TR_DATE_FMT, key=f"elst_{e['id']}")
                with t4:
                    try:
                        lt_val = datetime.strptime(e['lastik_termin_tarihi'], '%Y-%m-%d').date() if e.get('lastik_termin_tarihi') else date.today()
                    except Exception:
                        lt_val = date.today()
                    lastik_termin_tarihi = st.date_input("Termin", value=lt_val, format=TR_DATE_FMT, key=f"eltt_{e['id']}")

                if not full_service:
                    names2, ids2 = mf_options(manufacturers, "— Atama yok —")
                    em_idx = ids2.index(e.get('lastik_manufacturer_id')) if e.get('lastik_manufacturer_id') in ids2 else 0
                    em_sel = st.selectbox("Üretici", names2, index=em_idx, key=f"eemf_{e['id']}")
                    row_manufacturer_id = ids2[names2.index(em_sel)]
                else:
                    row_manufacturer_id = e.get('lastik_manufacturer_id')
                    st.caption(f"Üretici (tam hizmet): {o.get('manufacturer_name') or '—'}")

                if st.form_submit_button("💾 Kaydet", type="primary"):
                    db.update_elastic_row(e['id'], new_elastic_id, lastik_renk, row_manufacturer_id,
                                          siparis_cm, gelen_cm, siparis_adet, gelen_adet,
                                          lastik_siparis_tarihi.isoformat(), lastik_termin_tarihi.isoformat())
                    if lastik_foto_file:
                        path = save_uploaded_file(lastik_foto_file, f"order_{order_id}_elastic_{e['id']}")
                        db.set_elastic_photo(e['id'], path)
                    db.add_log(st.session_state.user['id'], st.session_state.user['username'],
                              'Lastik Güncellendi', f"{o['model_name']} (ID: {order_id})")
                    st.success("Kaydedildi!")
                    st.rerun()

            irs1, irs2 = st.columns(2)
            with irs1:
                if st.button("🚚 İrsaliye Oluştur", key=f"irs_el_{e['id']}"):
                    st.session_state[f'show_irs_el_{e["id"]}'] = True
            with irs2:
                if st.button("🗑️ Bu Lastik Satırını Sil", key=f"del_el_{e['id']}"):
                    db.delete_elastic_row(e['id'])
                    st.rerun()
            if st.session_state.get(f'show_irs_el_{e["id"]}'):
                birim = 'adet' if e['is_raporlu'] else 'cm'
                render_irsaliye_form(order_id, e['id'], "Lastik", e.get('lastik_manufacturer_name') or '', birim)

    if st.button("➕ Yeni Lastik Satırı Ekle", key=f"add_el_edit_{order_id}"):
        db.add_elastic_row(order_id)
        st.rerun()

    # --- Aksesuar satırları ---
    st.markdown('<div class="pk-section-title">🔘 Aksesuar</div>', unsafe_allow_html=True)
    for a in o['accessories']:
        with st.expander(f"{a['aksesuar_adi']} — {a.get('aksesuar_renk') or ''}"):
            with st.form(f"edit_accessory_{a['id']}"):
                r1, r2 = st.columns(2)
                with r1:
                    aksesuar_adi = st.text_input("Aksesuar Adı", value=a['aksesuar_adi'], key=f"eaa_{a['id']}")
                    aksesuar_renk = st.text_input("Renk", value=a.get('aksesuar_renk') or '', key=f"ear_{a['id']}")
                with r2:
                    aksesuar_foto_file = st.file_uploader("Fotoğraf (değiştir)", type=['png', 'jpg', 'jpeg'], key=f"eaf_{a['id']}")
                    show_photo(a.get('aksesuar_foto'), "Mevcut")

                st.markdown("**Tedarik Takibi**")
                t1, t2, t3, t4 = st.columns(4)
                with t1:
                    siparis_adet = st.number_input("Sipariş (adet)", min_value=0.0, value=float(a.get('aksesuar_siparis_adet') or 0), key=f"easa_{a['id']}")
                with t2:
                    gelen_adet = st.number_input("Gelen (adet)", min_value=0.0, value=float(a.get('aksesuar_gelen_adet') or 0), key=f"eaga_{a['id']}")
                with t3:
                    try:
                        as_val = datetime.strptime(a['aksesuar_siparis_tarihi'], '%Y-%m-%d').date() if a.get('aksesuar_siparis_tarihi') else date.today()
                    except Exception:
                        as_val = date.today()
                    aksesuar_siparis_tarihi = st.date_input("Sipariş T.", value=as_val, format=TR_DATE_FMT, key=f"east_{a['id']}")
                with t4:
                    try:
                        att_val = datetime.strptime(a['aksesuar_termin_tarihi'], '%Y-%m-%d').date() if a.get('aksesuar_termin_tarihi') else date.today()
                    except Exception:
                        att_val = date.today()
                    aksesuar_termin_tarihi = st.date_input("Termin", value=att_val, format=TR_DATE_FMT, key=f"eatt_{a['id']}")

                if not full_service:
                    names2, ids2 = mf_options(manufacturers, "— Atama yok —")
                    am_idx = ids2.index(a.get('aksesuar_manufacturer_id')) if a.get('aksesuar_manufacturer_id') in ids2 else 0
                    am_sel = st.selectbox("Üretici", names2, index=am_idx, key=f"eamf_{a['id']}")
                    row_manufacturer_id = ids2[names2.index(am_sel)]
                else:
                    row_manufacturer_id = a.get('aksesuar_manufacturer_id')
                    st.caption(f"Üretici (tam hizmet): {o.get('manufacturer_name') or '—'}")

                if st.form_submit_button("💾 Kaydet", type="primary"):
                    db.update_accessory_row(a['id'], aksesuar_adi, aksesuar_renk, row_manufacturer_id,
                                            siparis_adet, gelen_adet, aksesuar_siparis_tarihi.isoformat(),
                                            aksesuar_termin_tarihi.isoformat())
                    if aksesuar_foto_file:
                        path = save_uploaded_file(aksesuar_foto_file, f"order_{order_id}_accessory_{a['id']}")
                        db.set_accessory_photo(a['id'], path)
                    db.add_log(st.session_state.user['id'], st.session_state.user['username'],
                              'Aksesuar Güncellendi', f"{o['model_name']} (ID: {order_id})")
                    st.success("Kaydedildi!")
                    st.rerun()

            irs1, irs2 = st.columns(2)
            with irs1:
                if st.button("🚚 İrsaliye Oluştur", key=f"irs_ak_{a['id']}"):
                    st.session_state[f'show_irs_ak_{a["id"]}'] = True
            with irs2:
                if st.button("🗑️ Bu Aksesuar Satırını Sil", key=f"del_ak_{a['id']}"):
                    db.delete_accessory_row(a['id'])
                    st.rerun()
            if st.session_state.get(f'show_irs_ak_{a["id"]}'):
                render_irsaliye_form(order_id, a['id'], "Aksesuar", a.get('aksesuar_manufacturer_name') or '', "adet")

    if st.button("➕ Yeni Aksesuar Satırı Ekle", key=f"add_ak_edit_{order_id}"):
        db.add_accessory_row(order_id, "Yeni Aksesuar")
        st.rerun()

    # --- Silme ---
    st.markdown('<div class="pk-section-title">Tehlikeli Bölge</div>', unsafe_allow_html=True)
    del_key = f"confirm_del_{order_id}"
    if not st.session_state.get(del_key):
        if st.button("🗑️ Siparişi Sil"):
            st.session_state[del_key] = True
            st.rerun()
    else:
        st.warning("Bu işlem geri alınamaz. Silmeyi onaylıyor musunuz?")
        cc1, cc2 = st.columns(2)
        with cc1:
            if st.button("Evet, Sil", type="primary"):
                db.delete_order(order_id)
                db.add_log(st.session_state.user['id'], st.session_state.user['username'],
                          'Sipariş Silindi', f"{o['model_name']} (ID: {order_id})")
                st.session_state[del_key] = False
                st.session_state.edit_order_id = None
                st.rerun()
        with cc2:
            if st.button("Vazgeç"):
                st.session_state[del_key] = False
                st.rerun()


# ======================================================================
# AYARLAR (Admin only)
# ======================================================================
def page_ayarlar():
    page_header("⚙️ Ayarlar", "Ana verileri (üretici, ürün adı/grubu, kumaş, lastik, reçete) yönetin")

    tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9 = st.tabs(
        ["🏭 Üreticiler", "🏷️ Ürün Grupları", "🏷️ Ürün Adları", "🧵 Kumaşlar", "➰ Lastikler",
         "📐 Reçeteler", "🏢 Firma", "👤 Kullanıcılar", "🧾 Log"])

    # --- Üreticiler ---
    with tab1:
        with st.form("add_manufacturer"):
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                new_mf = st.text_input("Üretici Adı")
            with c2:
                contact_person = st.text_input("İletişim Kişisi")
            with c3:
                phone = st.text_input("Telefon")
            with c4:
                email = st.text_input("E-posta")
            submitted = st.form_submit_button("Ekle", width='stretch', type="primary")
            if submitted and new_mf.strip():
                if db.add_manufacturer(new_mf.strip(), contact_person.strip(), phone.strip(), email.strip()):
                    db.add_log(st.session_state.user['id'], st.session_state.user['username'],
                              'Üretici Eklendi', new_mf.strip())
                    st.success(f"'{new_mf.strip()}' eklendi!")
                    st.rerun()
                else:
                    st.error("Bu üretici zaten mevcut!")

        st.markdown("---")
        st.markdown("#### Mevcut Üreticiler")
        manufacturers = db.get_manufacturers()
        if manufacturers:
            for mf in manufacturers:
                with st.expander(f"{mf['name']}"):
                    with st.form(f"edit_mf_{mf['id']}"):
                        c1, c2, c3 = st.columns(3)
                        with c1:
                            cp = st.text_input("İletişim Kişisi", value=mf.get('contact_person') or '', key=f"mfcp_{mf['id']}")
                        with c2:
                            ph = st.text_input("Telefon", value=mf.get('phone') or '', key=f"mfph_{mf['id']}")
                        with c3:
                            em = st.text_input("E-posta", value=mf.get('email') or '', key=f"mfem_{mf['id']}")
                        if st.form_submit_button("💾 Kaydet"):
                            db.update_manufacturer(mf['id'], cp, ph, em)
                            st.success("Güncellendi!")
                            st.rerun()
                    if st.button("🗑️ Üreticiyi Sil", key=f"del_mf_{mf['id']}"):
                        db.delete_manufacturer(mf['id'])
                        db.add_log(st.session_state.user['id'], st.session_state.user['username'],
                                  'Üretici Silindi', mf['name'])
                        st.rerun()
        else:
            st.info("Henüz üretici tanımlanmamış.")

    # --- Ürün Grupları ---
    with tab2:
        st.caption("Örnek: Kadın için string, cheeky, slip, boyshort, short, pantolon, triangle bra; "
                  "Erkek için boxer, slip, short, pantolon.")
        if st.button("⚡ Varsayılan Ürün Gruplarını Ekle"):
            for gender, gnames in DEFAULT_PRODUCT_GROUPS.items():
                for gname in gnames:
                    db.add_product_group(gender, gname)
            st.success("Varsayılan ürün grupları eklendi!")
            st.rerun()

        with st.form("add_product_group"):
            c1, c2, c3 = st.columns([1, 2, 1])
            with c1:
                pg_gender = st.radio("Cinsiyet", GENDERS, horizontal=True)
            with c2:
                pg_name = st.text_input("Ürün Grubu Adı (ör. String, Boxer)")
            with c3:
                st.write("")
                submit_pg = st.form_submit_button("Ekle", width='stretch', type="primary")
            if submit_pg:
                if pg_name.strip():
                    if db.add_product_group(pg_gender, pg_name.strip()):
                        st.success(f"'{pg_name.strip()}' eklendi!")
                        st.rerun()
                    else:
                        st.error("Bu ürün grubu zaten mevcut!")

        st.markdown("---")
        for gender in GENDERS:
            st.markdown(f"#### {gender}")
            groups = db.get_product_groups(gender)
            if groups:
                for g in groups:
                    c1, c2 = st.columns([4, 1])
                    c1.write(g['name'])
                    if c2.button("🗑️", key=f"del_pg_{g['id']}"):
                        db.delete_product_group(g['id'])
                        st.rerun()
            else:
                st.caption("Henüz ürün grubu yok.")

    # --- Ürün Adları ---
    with tab3:
        st.caption("Belirli model isimleri tanımlayın (ör. 'Zebra Boxer'); Yeni Sipariş'te bu ad seçildiğinde "
                  "cinsiyet ve ürün grubu otomatik doldurulur.")
        with st.form("add_product_name"):
            c1, c2, c3, c4 = st.columns([2, 1, 2, 1])
            with c1:
                pn_name = st.text_input("Ürün Adı")
            with c2:
                pn_gender = st.radio("Cinsiyet", GENDERS, horizontal=True, key="pn_gender_add")
            with c3:
                pn_groups = db.get_product_groups(pn_gender)
                pn_urun_grubu = st.selectbox("Ürün Grubu", [g['name'] for g in pn_groups] if pn_groups else ["—"])
            with c4:
                st.write("")
                submit_pn = st.form_submit_button("Ekle", width='stretch', type="primary")
            if submit_pn:
                if pn_name.strip() and pn_groups:
                    if db.add_product_name(pn_name.strip(), pn_gender, pn_urun_grubu):
                        st.success(f"'{pn_name.strip()}' eklendi!")
                        st.rerun()
                    else:
                        st.error("Bu ürün adı zaten mevcut!")
                elif not pn_groups:
                    st.error("Önce Ürün Grupları sekmesinden bu cinsiyet için grup ekleyin.")

        st.markdown("---")
        st.markdown("#### Mevcut Ürün Adları")
        product_names = db.get_product_names()
        if product_names:
            for pn in product_names:
                with st.expander(f"{pn['name']} — {pn['gender']} / {pn['urun_grubu']}"):
                    with st.form(f"edit_pn_{pn['id']}"):
                        c1, c2, c3 = st.columns(3)
                        with c1:
                            new_name = st.text_input("Ürün Adı", value=pn['name'], key=f"pnn_{pn['id']}")
                        with c2:
                            new_gender = st.radio("Cinsiyet", GENDERS, horizontal=True,
                                                  index=GENDERS.index(pn['gender']) if pn['gender'] in GENDERS else 0,
                                                  key=f"png_{pn['id']}")
                        with c3:
                            gset = db.get_product_groups(new_gender)
                            gnames = [g['name'] for g in gset] or [pn['urun_grubu']]
                            gidx = gnames.index(pn['urun_grubu']) if pn['urun_grubu'] in gnames else 0
                            new_ug = st.selectbox("Ürün Grubu", gnames, index=gidx, key=f"pnu_{pn['id']}")
                        if st.form_submit_button("💾 Kaydet"):
                            db.update_product_name(pn['id'], new_name, new_gender, new_ug)
                            st.success("Güncellendi!")
                            st.rerun()
                    if st.button("🗑️ Sil", key=f"del_pn_{pn['id']}"):
                        db.delete_product_name(pn['id'])
                        st.rerun()
        else:
            st.info("Henüz ürün adı tanımlanmamış.")

    # --- Kumaşlar ---
    with tab4:
        product_names_all = db.get_product_names()
        pn_names_display = ["—"] + [p['name'] for p in product_names_all]
        with st.form("add_fabric"):
            c1, c2, c3, c4, c5 = st.columns(5)
            with c1:
                f_name = st.text_input("Kumaş Adı")
            with c2:
                f_turu = st.selectbox("Tür", FABRIC_TYPES)
            with c3:
                f_icerik = st.text_input("İçerik", placeholder="%95 Pamuk")
            with c4:
                f_en = st.number_input("En (cm)", min_value=0.0, value=0.0, step=1.0)
            with c5:
                f_grm2 = st.number_input("Gramaj (gr/m²)", min_value=0.0, value=0.0, step=1.0)
            f_urun_adi = st.selectbox("Ürün Adı (opsiyonel)", pn_names_display)
            if st.form_submit_button("Ekle", width='stretch', type="primary"):
                if f_name.strip():
                    urun_adi_id = None
                    if f_urun_adi != "—":
                        rec = next((p for p in product_names_all if p['name'] == f_urun_adi), None)
                        urun_adi_id = rec['id'] if rec else None
                    db.add_fabric(f_name.strip(), f_icerik.strip(), f_turu, f_en, f_grm2, urun_adi_id)
                    st.success(f"'{f_name.strip()}' eklendi!")
                    st.rerun()
                else:
                    st.error("Kumaş adı gerekli!")

        st.markdown("---")
        st.markdown("#### Mevcut Kumaşlar")
        fabrics = db.get_fabrics()
        if fabrics:
            for f in fabrics:
                with st.expander(f"{f['name']} ({f['kumas_turu']})"):
                    with st.form(f"edit_fabric_master_{f['id']}"):
                        c1, c2, c3, c4, c5 = st.columns(5)
                        with c1:
                            nname = st.text_input("Kumaş Adı", value=f['name'], key=f"fmn_{f['id']}")
                        with c2:
                            nturu = st.selectbox("Tür", FABRIC_TYPES, index=FABRIC_TYPES.index(f['kumas_turu']) if f['kumas_turu'] in FABRIC_TYPES else 0, key=f"fmt_{f['id']}")
                        with c3:
                            nicerik = st.text_input("İçerik", value=f['icerik'] or '', key=f"fmi_{f['id']}")
                        with c4:
                            nen = st.number_input("En (cm)", min_value=0.0, value=float(f['en'] or 0), key=f"fme_{f['id']}")
                        with c5:
                            ngrm2 = st.number_input("Gramaj", min_value=0.0, value=float(f['gr_m2'] or 0), key=f"fmg_{f['id']}")
                        pn_idx = pn_names_display.index(f['urun_adi']) if f.get('urun_adi') in pn_names_display else 0
                        n_urun_adi = st.selectbox("Ürün Adı", pn_names_display, index=pn_idx, key=f"fmu_{f['id']}")
                        if st.form_submit_button("💾 Kaydet"):
                            urun_adi_id = None
                            if n_urun_adi != "—":
                                rec = next((p for p in product_names_all if p['name'] == n_urun_adi), None)
                                urun_adi_id = rec['id'] if rec else None
                            db.update_fabric(f['id'], nname, nicerik, nturu, nen, ngrm2, urun_adi_id)
                            st.success("Güncellendi!")
                            st.rerun()
                    if st.button("🗑️ Sil", key=f"del_fab_master_{f['id']}"):
                        db.delete_fabric(f['id'])
                        st.rerun()
        else:
            st.info("Henüz kumaş tanımlanmamış.")

    # --- Lastikler ---
    with tab5:
        product_names_all = db.get_product_names()
        pn_names_display = ["—"] + [p['name'] for p in product_names_all]
        with st.form("add_elastic"):
            c1, c2, c3 = st.columns(3)
            with c1:
                e_ad = st.selectbox("Lastik Adı", ELASTIC_TYPES)
            with c2:
                e_boyut = st.text_input("Boyut", placeholder="2 cm")
            with c3:
                e_urun_adi = st.selectbox("Ürün Adı (opsiyonel)", pn_names_display)
            if st.form_submit_button("Ekle", width='stretch', type="primary"):
                urun_adi_id = None
                if e_urun_adi != "—":
                    rec = next((p for p in product_names_all if p['name'] == e_urun_adi), None)
                    urun_adi_id = rec['id'] if rec else None
                db.add_elastic(e_ad, e_boyut.strip(), urun_adi_id)
                st.success(f"'{e_ad}' eklendi!")
                st.rerun()

        st.markdown("---")
        st.markdown("#### Mevcut Lastikler")
        elastics = db.get_elastics()
        if elastics:
            for e in elastics:
                with st.expander(f"{e['tur']} — {e['boyut'] or '-'}"):
                    with st.form(f"edit_elastic_master_{e['id']}"):
                        c1, c2, c3 = st.columns(3)
                        with c1:
                            nad = st.selectbox("Lastik Adı", ELASTIC_TYPES,
                                               index=ELASTIC_TYPES.index(e['tur']) if e['tur'] in ELASTIC_TYPES else 0,
                                               key=f"emn_{e['id']}")
                        with c2:
                            nboyut = st.text_input("Boyut", value=e['boyut'] or '', key=f"emb_{e['id']}")
                        with c3:
                            pn_idx = pn_names_display.index(e['urun_adi']) if e.get('urun_adi') in pn_names_display else 0
                            n_urun_adi = st.selectbox("Ürün Adı", pn_names_display, index=pn_idx, key=f"emu_{e['id']}")
                        if st.form_submit_button("💾 Kaydet"):
                            urun_adi_id = None
                            if n_urun_adi != "—":
                                rec = next((p for p in product_names_all if p['name'] == n_urun_adi), None)
                                urun_adi_id = rec['id'] if rec else None
                            db.update_elastic(e['id'], nad, nboyut, urun_adi_id)
                            st.success("Güncellendi!")
                            st.rerun()
                    if st.button("🗑️ Sil", key=f"del_ela_master_{e['id']}"):
                        db.delete_elastic(e['id'])
                        st.rerun()
        else:
            st.info("Henüz lastik tanımlanmamış.")

    # --- Reçeteler ---
    with tab6:
        st.caption("Cinsiyet + Ürün Grubu + Beden bazında kumaş (gr) ve lastik (adet/cm) reçetesi tanımlayın. "
                  "Yeni sipariş oluşturulurken bu değerler otomatik uygulanır.")

        # Düzenle butonundan gelen bekleyen seçim varsa, radio/selectbox oluşturulmadan ÖNCE uygula
        pending = st.session_state.pop('_pending_recipe_edit', None)
        if pending:
            st.session_state['recipe_gender_edit'] = pending[0]
            st.session_state['recipe_ug_edit'] = pending[1]

        if 'recipe_gender_edit' not in st.session_state:
            st.session_state['recipe_gender_edit'] = GENDERS[0]

        r_gender = st.radio("Cinsiyet", GENDERS, horizontal=True, key="recipe_gender_edit")
        groups = db.get_product_groups(r_gender)
        if not groups:
            st.warning(f"Önce '{r_gender}' için Ürün Grupları sekmesinden bir ürün grubu ekleyin.")
        else:
            group_names = [g['name'] for g in groups]
            if st.session_state.get('recipe_ug_edit') not in group_names:
                st.session_state['recipe_ug_edit'] = group_names[0]
            r_urun_grubu = st.selectbox("Ürün Grubu", group_names, key="recipe_ug_edit")
            existing = db.get_recipe(r_gender, r_urun_grubu)

            with st.form(f"recipe_form"):
                st.markdown(f"**{r_gender} — {r_urun_grubu} reçetesi**")
                size_vals = {}
                cols = st.columns(7)
                for i, size in enumerate(SIZES):
                    with cols[i]:
                        st.markdown(f"**{size}**")
                        kg = st.number_input("Kumaş (gr)", min_value=0.0, value=float(existing[size]['kumas_gr']),
                                             step=1.0, key=f"rk_{size}")
                        la = st.number_input("Lastik (adet)", min_value=0.0, value=float(existing[size]['lastik_adet']),
                                             step=1.0, key=f"ra_{size}")
                        lc = st.number_input("Lastik (cm)", min_value=0.0, value=float(existing[size]['lastik_cm']),
                                             step=1.0, key=f"rc_{size}")
                        size_vals[size] = (kg, la, lc)
                if st.form_submit_button("💾 Reçeteyi Kaydet", type="primary"):
                    for size, (kg, la, lc) in size_vals.items():
                        db.upsert_recipe(r_gender, r_urun_grubu, size, kg, la, lc)
                    db.add_log(st.session_state.user['id'], st.session_state.user['username'],
                              'Reçete Güncellendi', f"{r_gender} - {r_urun_grubu}")
                    st.success("Reçete kaydedildi!")
                    st.rerun()

        st.markdown("---")
        st.markdown("#### Tanımlı Reçeteler")
        recipe_groups = db.get_recipe_groups()
        if recipe_groups:
            for rg in recipe_groups:
                c1, c2, c3 = st.columns([3, 1, 1])
                with c1:
                    st.markdown(f"**{rg['gender']} — {rg['urun_grubu']}**")
                with c2:
                    if st.button("✏️ Düzenle", key=f"editrec_{rg['gender']}_{rg['urun_grubu']}"):
                        st.session_state['_pending_recipe_edit'] = (rg['gender'], rg['urun_grubu'])
                        st.rerun()
                with c3:
                    if st.button("🗑️ Sil", key=f"delrec_{rg['gender']}_{rg['urun_grubu']}"):
                        db.delete_recipe_group(rg['gender'], rg['urun_grubu'])
                        st.rerun()
                r = db.get_recipe(rg['gender'], rg['urun_grubu'])
                df = pd.DataFrame([{'Beden': s, 'Kumaş (gr)': r[s]['kumas_gr'],
                                    'Lastik (adet)': r[s]['lastik_adet'], 'Lastik (cm)': r[s]['lastik_cm']}
                                   for s in SIZES])
                st.dataframe(df, width='stretch', hide_index=True)
                st.markdown("---")
        else:
            st.info("Henüz reçete tanımlanmamış.")

    # --- Firma ---
    with tab7:
        settings = db.get_settings()
        with st.form("settings_form"):
            firma_adi = st.text_input("Firma Adı", value=settings['firma_adi'])
            if st.form_submit_button("💾 Kaydet", type="primary"):
                db.update_settings(firma_adi.strip())
                db.add_log(st.session_state.user['id'], st.session_state.user['username'],
                          'Ayarlar Güncellendi', 'Firma adı')
                st.success("Kaydedildi!")
                st.rerun()

    # --- Kullanıcılar ---
    with tab8:
        with st.form("add_user"):
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                new_username = st.text_input("Kullanıcı Adı")
            with col2:
                new_fullname = st.text_input("Ad Soyad")
            with col3:
                new_password = st.text_input("Şifre", type="password")
            with col4:
                new_role = st.selectbox("Yetki", ["user", "admin"])

            submitted = st.form_submit_button("Kullanıcı Ekle", width='stretch', type="primary")
            if submitted:
                if new_username.strip() and new_password and new_fullname.strip():
                    if len(new_password) < 6:
                        st.error("Şifre en az 6 karakter olmalı!")
                    elif db.add_user(new_username.strip(), new_password, new_fullname.strip(), new_role):
                        db.add_log(st.session_state.user['id'], st.session_state.user['username'],
                                  'Kullanıcı Eklendi', new_username.strip())
                        st.success(f"'{new_username.strip()}' kullanıcısı eklendi!")
                        st.rerun()
                    else:
                        st.error("Bu kullanıcı adı zaten mevcut!")
                else:
                    st.error("Tüm alanları doldurun!")

        st.markdown("---")
        st.markdown("#### Mevcut Kullanıcılar")
        users = db.get_users()
        for user in users:
            with st.expander(f"{user['full_name']} (@{user['username']}) - {user['role'].upper()}"):
                col1, col2 = st.columns(2)
                with col1:
                    new_pw = st.text_input("Yeni Şifre", type="password", key=f"pw_{user['id']}")
                    if st.button("Şifre Değiştir", key=f"chpw_{user['id']}"):
                        if new_pw and len(new_pw) >= 6:
                            db.update_user_password(user['id'], new_pw)
                            db.add_log(st.session_state.user['id'], st.session_state.user['username'],
                                      'Şifre Değişikliği', f"{user['username']} kullanıcısının şifresi değiştirildi")
                            st.success("Şifre değiştirildi!")
                        else:
                            st.error("Şifre en az 6 karakter olmalı.")
                with col2:
                    new_role = st.selectbox("Yetki", ["user", "admin"],
                                           index=0 if user['role'] == 'user' else 1,
                                           key=f"role_{user['id']}")
                    if st.button("Yetki Güncelle", key=f"role_btn_{user['id']}"):
                        db.update_user_role(user['id'], new_role)
                        st.success("Yetki güncellendi!")

                if user['username'] != 'admin':
                    if st.button("🗑️ Kullanıcıyı Sil", key=f"del_user_{user['id']}"):
                        db.delete_user(user['id'])
                        db.add_log(st.session_state.user['id'], st.session_state.user['username'],
                                  'Kullanıcı Silindi', user['username'])
                        st.rerun()

        st.markdown("---")
        st.markdown("#### Kendi Şifreni Değiştir")
        with st.form("change_own_pw"):
            old_pw = st.text_input("Mevcut Şifre", type="password")
            new_pw = st.text_input("Yeni Şifre", type="password")
            new_pw2 = st.text_input("Yeni Şifre (Tekrar)", type="password")
            submitted = st.form_submit_button("Şifremi Değiştir", width='stretch', type="primary")
            if submitted:
                if db.verify_password(old_pw, st.session_state.user['password_hash']):
                    if new_pw == new_pw2 and len(new_pw) >= 6:
                        db.update_user_password(st.session_state.user['id'], new_pw)
                        user = db.authenticate(st.session_state.user['username'], new_pw)
                        st.session_state.user = user
                        db.add_log(st.session_state.user['id'], st.session_state.user['username'],
                                  'Şifre Değişikliği', 'Kullanıcı kendi şifresini değiştirdi')
                        st.success("Şifreniz değiştirildi!")
                    else:
                        st.error("Yeni şifreler eşleşmiyor veya çok kısa (min 6 karakter)!")
                else:
                    st.error("Mevcut şifre yanlış!")

    # --- Log ---
    with tab9:
        logs = db.get_logs(200)
        if logs:
            log_data = [{
                'Tarih/Saat': log['created_at'][:19].replace('T', ' '),
                'Kullanıcı': log.get('username', '-'),
                'İşlem': log['action'],
                'Detay': log.get('details', '-')
            } for log in logs]
            st.dataframe(pd.DataFrame(log_data), width='stretch', hide_index=True)
        else:
            st.info("Henüz log kaydı yok.")


# ======================================================================
# EXCEL İŞLEMLERİ & YEDEKLEME
# ======================================================================
def page_excel_yedek():
    page_header("📤 Excel İşlemleri & Yedekleme", "Toplu dışa/içe aktarma ve tam sistem yedeği")

    if not HAS_EXCEL:
        st.warning("Excel desteği için `openpyxl` kütüphanesi kurulu değil.")

    st.markdown("### 📥 Siparişleri Excel Olarak İndir")
    orders = db.get_orders_overview(assigned=None, include_completed=True)
    if orders and HAS_EXCEL:
        rows = []
        for o in orders:
            row = {
                'ID': o['id'], 'Ürün Adı': o['model_name'], 'Cinsiyet': o['gender'], 'Ürün Grubu': o.get('urun_grubu'),
                'Üretici': o.get('manufacturer_name') or '-', 'Durum': o['status'],
                'Sipariş Tarihi': tr_date(o.get('siparis_tarihi')), 'Termin': tr_date(o.get('deadline')),
                'Paket': o['package_size'], 'Toplam Kutu': o['total_boxes'], 'Toplam Adet': o['total_quantity'],
                'Kumaş (kg)': o['kumas_kg_total'], 'Lastik (cm)': o['lastik_cm_total'],
                'Lastik (adet-Raporlu)': o['lastik_adet_total'],
            }
            for s in o['sizes']:
                row[f"Adet_{s['size']}"] = s['box_qty'] * o['package_size']
            rows.append(row)
        df = pd.DataFrame(rows)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        st.download_button("📥 Tüm Siparişleri Excel İndir", output.getvalue(),
                          file_name=f"siparisler_{date.today().isoformat()}.xlsx",
                          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    elif not orders:
        st.info("Henüz sipariş yok.")

    st.markdown("---")
    st.markdown("### 📤 Excel'den Toplu Sipariş Yükle")
    st.caption("Excel dosyanızda en az şu kolonlar bulunmalı: Ürün Adı, Cinsiyet, Ürün Grubu, Paket, Adet_XS...Adet_3XL. "
              "Yüklenen siparişler reçete kütüphanesine göre otomatik hesaplanarak 'Planlama Bekleyen' listesine eklenir.")
    uploaded = st.file_uploader("Excel Dosyası (.xlsx)", type=['xlsx'])
    if uploaded and HAS_EXCEL:
        try:
            df_up = pd.read_excel(uploaded)
            st.dataframe(df_up.head(10), width='stretch')
            st.info(f"{len(df_up)} satır bulundu.")
            if st.button("✅ İçe Aktar"):
                success_count = 0
                for _, row in df_up.iterrows():
                    try:
                        gender = str(row.get('Cinsiyet', 'Kadın'))
                        urun_grubu = str(row.get('Ürün Grubu', ''))
                        package_size = int(row.get('Paket', 1) or 1)
                        recipe = db.get_recipe(gender, urun_grubu)
                        sizes_payload = []
                        for size in SIZES:
                            col = f"Adet_{size}"
                            adet = int(row.get(col, 0) or 0)
                            box_qty = adet // package_size if package_size else adet
                            if box_qty > 0:
                                r = recipe.get(size, {'kumas_gr': 0, 'lastik_adet': 0, 'lastik_cm': 0})
                                sizes_payload.append({'size': size, 'box_qty': box_qty, **r})
                        if not sizes_payload:
                            continue
                        db.add_order(str(row.get('Ürün Adı', 'Yeni Model')), gender, urun_grubu, package_size,
                                    3.0, 3.0, str(row.get('Termin', date.today().isoformat())),
                                    date.today().isoformat(), sizes_payload,
                                    [{}], [{}], [], st.session_state.user['id'])
                        success_count += 1
                    except Exception as e:
                        st.warning(f"Satır atlandı: {e}")
                db.add_log(st.session_state.user['id'], st.session_state.user['username'],
                          'Excel İçe Aktarma', f"{success_count} sipariş eklendi")
                st.success(f"{success_count} sipariş içe aktarıldı!")
                st.rerun()
        except Exception as e:
            st.error(f"Excel okuma hatası: {e}")

    st.markdown("---")
    st.markdown("### 💾 Tüm Verileri Yedekle (Zip)")
    st.caption("Tüm veritabanı tablolarını JSON olarak ve yüklenmiş fotoğrafları içeren bir zip dosyası indirin.")
    if st.button("💾 Yedek Zip Oluştur"):
        data = db.export_all_data()
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as z:
            for table, table_rows in data.items():
                z.writestr(f"data/{table}.json", json.dumps(table_rows, ensure_ascii=False, indent=2, default=str))
            if os.path.isdir(UPLOAD_DIR):
                for fname in os.listdir(UPLOAD_DIR):
                    fpath = os.path.join(UPLOAD_DIR, fname)
                    if os.path.isfile(fpath):
                        z.write(fpath, arcname=f"uploads/{fname}")
        zip_buffer.seek(0)
        st.download_button("⬇️ Zip İndir", zip_buffer.getvalue(),
                          file_name=f"paulkenzie_yedek_{date.today().isoformat()}.zip", mime="application/zip")


# ======================================================================
# NAVİGASYON
# ======================================================================
def main():
    if not st.session_state.logged_in:
        page_login()
        return

    user = st.session_state.user

    with st.sidebar:
        st.markdown(
            '<div class="pk-brand"><div class="pk-logo">🧥 PAUL KENZIE</div>'
            '<div class="pk-sub">Üretim ERP</div></div>', unsafe_allow_html=True)

        st.markdown(
            f'<div class="pk-user-card">👤 <b>{user["full_name"]}</b><br>'
            f'<span style="font-size:0.75rem;">{user["role"].upper()}</span></div>',
            unsafe_allow_html=True)

        pages = {
            '📊 Dashboard': 'dashboard',
            '📝 Yeni Sipariş': 'yeni_siparis',
            '🏭 Atanmış Siparişler': 'atanmis',
            '🗂️ Planlama Bekleyen': 'planlama',
            '✅ Tamamlanmış Siparişler': 'tamamlanmis',
        }
        if user['role'] == 'admin':
            pages['📤 Excel & Yedek'] = 'excel_yedek'
            pages['⚙️ Ayarlar'] = 'ayarlar'

        selection = st.radio("Sayfalar", list(pages.keys()), label_visibility="collapsed")
        page_key = pages[selection]

        st.markdown("---")
        if st.button("🚪 Çıkış Yap", width='stretch'):
            db.add_log(user['id'], user['username'], 'Çıkış', f'{user["full_name"]} sistemden çıktı')
            st.session_state.logged_in = False
            st.session_state.user = None
            st.session_state.edit_order_id = None
            reset_wizard()
            st.rerun()

    if st.session_state.edit_order_id and user['role'] == 'admin':
        page_order_edit(st.session_state.edit_order_id)
        return

    if page_key == 'dashboard':
        page_dashboard()
    elif page_key == 'yeni_siparis':
        page_yeni_siparis()
    elif page_key == 'atanmis':
        page_atanmis_siparisler()
    elif page_key == 'planlama':
        page_planlama_bekleyen()
    elif page_key == 'tamamlanmis':
        page_tamamlanmis_siparisler()
    elif page_key == 'excel_yedek':
        if user['role'] == 'admin':
            page_excel_yedek()
        else:
            st.error("Bu sayfaya erişim yetkiniz yok!")
    elif page_key == 'ayarlar':
        if user['role'] == 'admin':
            page_ayarlar()
        else:
            st.error("Bu sayfaya erişim yetkiniz yok!")


if __name__ == '__main__':
    main()
