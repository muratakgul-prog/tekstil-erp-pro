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
    DEFAULT_PRODUCT_GROUPS, CURRENCIES, IRSALIYE_KATEGORILERI,
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

# ======================================================================
# GÖRSEL TEMA / CSS
# ======================================================================
CUSTOM_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

/* Sadece güvenli metin elemanlarına özel font uygula - ikon fontlarını (chevron/ok simgeleri)
   ASLA ezme! Genel span/div üzerine kural koymak Streamlit'in materyal ikon
   font ligature'larını bozup "keyboard_arrow_right" gibi ham metin olarak
   görünmesine (yazı çakışmasına) sebep oluyordu. */
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
/* İkon fontlarını (Material Symbols) olduğu gibi bırak - kesinlikle ezme */
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

def get_urgency_color(deadline_str):
    if not deadline_str:
        return '#94a3b8'
    try:
        dl = datetime.strptime(deadline_str, '%Y-%m-%d').date()
        diff = (dl - date.today()).days
        if diff <= 3:
            return '#dc2626'
        elif diff <= 7:
            return '#d97706'
        elif diff <= 14:
            return '#ca8a04'
        else:
            return '#16a34a'
    except Exception:
        return '#94a3b8'


def get_urgency_label(deadline_str):
    if not deadline_str:
        return '⚪ Belirtilmedi'
    try:
        dl = datetime.strptime(deadline_str, '%Y-%m-%d').date()
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
    """Kumaş takip satırı arka plan rengi: 15+ gün yeşil, 1-14 turuncu, 0/geçmiş kırmızı."""
    if not termin_str:
        return '#ffffff'
    try:
        dl = datetime.strptime(termin_str, '%Y-%m-%d').date()
        diff = (dl - date.today()).days
        if diff <= 0:
            return '#fee2e2'
        elif diff <= 14:
            return '#ffedd5'
        else:
            return '#dcfce7'
    except Exception:
        return '#ffffff'


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
    c.drawString(20 * mm, y, f"Geliş Tarihi: {irsaliye.get('gelis_tarihi', '-')}")
    y -= 7 * mm
    c.drawString(20 * mm, y, f"Kategori: {irsaliye.get('kategori', '-')} | Model: {order.get('model_name','-')} "
                             f"({order.get('model_kodu','-')})")
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


def render_irsaliye_form(order_id, color_id, kategori, tedarikci_default, birim):
    key_prefix = f"irs_{kategori}_{order_id}_{color_id}"
    with st.form(f"form_{key_prefix}"):
        c1, c2, c3 = st.columns(3)
        with c1:
            irsaliye_no = st.text_input("İrsaliye No", value=f"IRS-{datetime.now().strftime('%Y%m%d')}-{order_id}",
                                        key=f"no_{key_prefix}")
            tedarikci = st.text_input("Tedarikçi", value=tedarikci_default, key=f"ted_{key_prefix}")
        with c2:
            miktar = st.number_input(f"Miktar ({birim})", min_value=0.0, value=0.0, key=f"mik_{key_prefix}")
            gelis_tarihi = st.date_input("Geliş Tarihi", value=date.today(), key=f"gt_{key_prefix}")
        with c3:
            aciklama = st.text_area("Açıklama", key=f"ac_{key_prefix}")
        if st.form_submit_button("💾 İrsaliyeyi Kaydet", type="primary"):
            irs_id = db.add_irsaliye(order_id, color_id, kategori, irsaliye_no, tedarikci, miktar, birim,
                                     gelis_tarihi.isoformat(), aciklama)
            db.add_log(st.session_state.user['id'], st.session_state.user['username'],
                      'İrsaliye Oluşturuldu', f"{kategori} - {irsaliye_no}")
            st.session_state[f'last_irs_{key_prefix}'] = irs_id
            st.success("İrsaliye kaydedildi!")
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


def render_fabric_tracking_table(colors):
    rows_html = ""
    for col in colors:
        bg = fabric_row_color(col.get('kumas_termin_tarihi'))
        rows_html += f"""<tr style="background:{bg};">
            <td>{col['color_name']}</td>
            <td>{col.get('kumas_adi') or '-'}</td>
            <td>{col['kumas_kg']}</td>
            <td>{col.get('kumas_siparis_kg') or 0}</td>
            <td>{col.get('kumas_gelen_kg') or 0}</td>
            <td>{col['kumas_kalan_kg']}</td>
            <td>{col.get('kumas_siparis_tarihi') or '-'}</td>
            <td>{col.get('kumas_termin_tarihi') or '-'}</td>
        </tr>"""
    html = f"""<table class="pk-track-table">
        <tr><th>Renk</th><th>Kumaş</th><th>Gerekli (kg)</th><th>Sipariş (kg)</th>
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


def render_steps(active):
    html = '<div class="pk-steps">'
    for i, label in enumerate(STEP_LABELS, 1):
        cls = 'done' if i < active else ('active' if i == active else '')
        icon = '✓ ' if i < active else ''
        html += f'<div class="pk-step {cls}">{icon}{i}. {label}</div>'
    html += '</div>'
    st.markdown(html, unsafe_allow_html=True)


def reset_wizard():
    st.session_state.order_step = 1
    st.session_state.order_data = {}
    for k in list(st.session_state.keys()):
        if k.startswith('size_box_') or k.startswith('col_'):
            del st.session_state[k]


def page_yeni_siparis():
    page_header("📝 Yeni Sipariş", "Ürün, sipariş adedi ve malzeme bilgilerini girin")

    if 'order_step' not in st.session_state:
        st.session_state.order_step = 1
    if 'order_data' not in st.session_state:
        st.session_state.order_data = {}

    step = st.session_state.order_step
    render_steps(step)

    fabrics = db.get_fabrics()
    elastics = db.get_elastics()

    # ---------------- STEP 1: Temel Bilgiler ----------------
    if step == 1:
        c1, c2 = st.columns(2)
        with c1:
            model_name = st.text_input("ÜRÜN ADI *", key="w_model_name")
            gender = st.radio("CİNSİYET *", GENDERS, horizontal=True, key="w_gender")
            groups = db.get_product_groups(gender)
            if not groups:
                st.warning(f"'{gender}' için Ayarlar > Ürün Grupları bölümünden ürün grubu tanımlamalısınız.")
                urun_grubu = None
            else:
                urun_grubu = st.selectbox("ÜRÜN GRUBU *", [g['name'] for g in groups], key="w_urun_grubu")
            package_size = st.selectbox("KUTU İÇİ ADEDİ *", PACKAGE_SIZES,
                                         format_func=lambda x: f"{x}'li paket", key="w_package_size")
        with c2:
            deadline = st.date_input("TERMİN TARİHİ *", min_value=date.today(), key="w_deadline")
            kumas_fire = st.number_input("KUMAŞ FİRE ORANI (%)", min_value=0.0, max_value=50.0, value=3.0, step=0.5,
                                          key="w_kumas_fire")
            lastik_fire = st.number_input("LASTİK FİRE ORANI (%)", min_value=0.0, max_value=50.0, value=3.0, step=0.5,
                                           key="w_lastik_fire")
            urun_foto = st.file_uploader("ÜRÜN FOTOĞRAFI (opsiyonel)", type=['png', 'jpg', 'jpeg'], key="w_urun_foto")

        settings = db.get_settings()
        manufacturers = db.get_manufacturers()
        with st.expander("💰 Maliyet Parametreleri (opsiyonel)"):
            c3, c4 = st.columns(2)
            with c3:
                para_birimi = st.selectbox("Para Birimi", CURRENCIES,
                                           index=CURRENCIES.index(settings['varsayilan_para']), key="w_para")
                iscilik_birim = st.number_input("İşçilik (birim/parça)", min_value=0.0, value=0.0, step=0.5, key="w_iscilik")
            with c4:
                genel_gider_yuzde = st.number_input("Genel Gider (%)", min_value=0.0, value=10.0, step=1.0, key="w_gg")
                kar_yuzde = st.number_input("Kâr Marjı (%)", min_value=0.0, value=30.0, step=1.0, key="w_kar")
            st.markdown("**Kutu**")
            c5, c6 = st.columns(2)
            with c5:
                kutu_mf_names, kutu_mf_ids = mf_options(manufacturers, "— Belirtilmedi —")
                kutu_mf_sel = st.selectbox("Kutu Üreticisi", kutu_mf_names, key="w_kutu_mf")
                kutu_manufacturer_id = kutu_mf_ids[kutu_mf_names.index(kutu_mf_sel)]
            with c6:
                kutu_fiyat = st.number_input("Kutu Fiyatı (adet başına)", min_value=0.0, value=0.0, step=0.5, key="w_kutu_fiyat")

        if st.button("Devam Et →", width='stretch', type="primary"):
            if not model_name.strip():
                st.error("Ürün adı boş olamaz!")
            elif not urun_grubu:
                st.error("Ürün grubu seçilmedi!")
            else:
                st.session_state.order_data.update({
                    'model_name': model_name.strip(),
                    'gender': gender,
                    'urun_grubu': urun_grubu,
                    'package_size': package_size,
                    'deadline': deadline.isoformat(),
                    'kumas_fire_orani': kumas_fire,
                    'lastik_fire_orani': lastik_fire,
                    '_urun_foto_file': urun_foto,
                    'para_birimi': para_birimi,
                    'iscilik_birim': iscilik_birim,
                    'genel_gider_yuzde': genel_gider_yuzde,
                    'kar_yuzde': kar_yuzde,
                    'kutu_manufacturer_id': kutu_manufacturer_id,
                    'kutu_fiyat': kutu_fiyat,
                })
                st.session_state.order_step = 2
                st.rerun()

    # ---------------- STEP 2: Sipariş Adedi ----------------
    elif step == 2:
        data = st.session_state.order_data
        st.markdown(f"**Ürün:** {data['model_name']}  |  **Cinsiyet:** {data['gender']}  |  "
                    f"**Ürün Grubu:** {data['urun_grubu']}  |  **Paket:** {data['package_size']}'li")

        recipe = db.get_recipe(data['gender'], data['urun_grubu'])
        has_recipe = any(recipe[s]['kumas_gr'] > 0 or recipe[s]['lastik_mt'] > 0 for s in SIZES)
        if not has_recipe:
            st.warning("Bu ürün grubu için Ayarlar > Reçeteler bölümünde henüz kumaş/lastik reçetesi tanımlanmamış. "
                       "Kutu adetlerini girebilirsiniz ama kumaş/lastik ihtiyacı 0 hesaplanacaktır.")

        st.markdown("#### Sipariş Adedi (Beden Bazlı Kutu Sayısı)")
        cols = st.columns(7)
        for i, size in enumerate(SIZES):
            with cols[i]:
                st.number_input(size, min_value=0, value=0, step=1, key=f"size_box_{size}", label_visibility="visible")
                r = recipe[size]
                st.caption(f"K:{r['kumas_gr']}gr L:{r['lastik_mt']}mt")

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
                            'lastik_mt': r['lastik_mt'],
                        })
                    st.session_state.order_data['sizes'] = sizes_payload
                    st.session_state.order_data['total_boxes'] = sum(s['box_qty'] for s in sizes_payload)
                    st.session_state.order_step = 3
                    st.rerun()

    # ---------------- STEP 3: Kumaş / Lastik / Aksesuar ----------------
    elif step == 3:
        data = st.session_state.order_data
        package_size = data['package_size']

        if package_size == 1:
            st.caption("Tekli ürün — tek renk/varyant bilgisi girin.")
        else:
            st.caption(f"Paket içinde {package_size} renk var — her renk için ayrı satır girin.")

        with st.form("order_step3"):
            colors_input = []

            st.markdown("### 🧵 Kumaş")
            fabric_names = ["— Tanımlı değil —"] + [f"{f['name']} ({f['kumas_turu']})" for f in fabrics]
            fabric_ids = [None] + [f['id'] for f in fabrics]
            fabric_info = {}
            for i in range(1, package_size + 1):
                st.markdown(f"**Renk {i}**" if package_size > 1 else "**Ürün**")
                r1, r2, r3 = st.columns(3)
                with r1:
                    color_name = st.text_input("Renk / Varyant Adı", key=f"col_name_{i}",
                                                value="Tek Renk" if package_size == 1 else "")
                with r2:
                    fsel = st.selectbox("Kumaş", fabric_names, key=f"col_fabric_{i}")
                    fabric_id = fabric_ids[fabric_names.index(fsel)]
                with r3:
                    kumas_renk = st.text_input("Kumaş Rengi", key=f"col_krenk_{i}")
                kumas_foto = st.file_uploader("Kumaş Fotoğrafı", type=['png', 'jpg', 'jpeg'], key=f"col_kfoto_{i}")
                fabric_info[i] = {'color_name': color_name, 'kumas_fabric_id': fabric_id, 'kumas_renk': kumas_renk,
                                   '_kumas_foto_file': kumas_foto}
                if i < package_size:
                    st.markdown("")

            st.markdown("---")
            st.markdown("### ➰ Lastik")
            elastic_labels = ["— Tanımlı değil —"] + [f"{e['tur']} - {e['ad']} ({e['boyut']})" for e in elastics]
            elastic_ids = [None] + [e['id'] for e in elastics]
            elastic_info = {}
            for i in range(1, package_size + 1):
                st.markdown(f"**Renk {i}**" if package_size > 1 else "**Ürün**")
                r1, r2 = st.columns(2)
                with r1:
                    esel = st.selectbox("Lastik", elastic_labels, key=f"col_elastic_{i}")
                    elastic_id = elastic_ids[elastic_labels.index(esel)]
                with r2:
                    lastik_renk = st.text_input("Lastik Rengi", key=f"col_lrenk_{i}")
                lastik_foto = st.file_uploader("Lastik Fotoğrafı", type=['png', 'jpg', 'jpeg'], key=f"col_lfoto_{i}")
                elastic_info[i] = {'lastik_elastic_id': elastic_id, 'lastik_renk': lastik_renk,
                                    '_lastik_foto_file': lastik_foto}

            st.markdown("---")
            st.markdown("### 🔘 Aksesuar (opsiyonel)")
            aksesuar_info = {}
            for i in range(1, package_size + 1):
                st.markdown(f"**Renk {i}**" if package_size > 1 else "**Ürün**")
                r1, r2 = st.columns(2)
                with r1:
                    aksesuar_adi = st.text_input("Aksesuar Adı", key=f"col_aadi_{i}")
                with r2:
                    aksesuar_renk = st.text_input("Aksesuar Rengi", key=f"col_arenk_{i}")
                aksesuar_foto = st.file_uploader("Aksesuar Fotoğrafı", type=['png', 'jpg', 'jpeg'], key=f"col_afoto_{i}")
                aksesuar_info[i] = {'aksesuar_adi': aksesuar_adi, 'aksesuar_renk': aksesuar_renk,
                                     '_aksesuar_foto_file': aksesuar_foto}

            for i in range(1, package_size + 1):
                colors_input.append({**fabric_info[i], **elastic_info[i], **aksesuar_info[i]})

            cb1, cb2 = st.columns(2)
            with cb1:
                back = st.form_submit_button("← Geri", width='stretch')
            with cb2:
                next_btn = st.form_submit_button("Devam Et →", width='stretch', type="primary")

            if back:
                st.session_state.order_step = 2
                st.rerun()
            if next_btn:
                st.session_state.order_data['colors'] = colors_input
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
        st.markdown(f"**Cinsiyet:** {data['gender']}  |  **Termin:** {data['deadline']}  |  "
                    f"**Toplam Adet (parça):** {total_qty}  |  "
                    f"**Kumaş Fire:** %{data['kumas_fire_orani']}  |  **Lastik Fire:** %{data['lastik_fire_orani']}")

        st.markdown("---")
        st.markdown("**Beden Dağılımı**")
        st.dataframe(pd.DataFrame([
            {'Beden': s['size'], 'Kutu Adedi': s['box_qty'], 'Kumaş (gr/parça)': s['kumas_gr'],
             'Lastik (adet/parça)': s['lastik_adet'], 'Lastik (mt/parça)': s['lastik_mt']}
            for s in data['sizes']
        ]), width='stretch', hide_index=True)

        kumas_fire_mult = 1 + data['kumas_fire_orani'] / 100.0
        lastik_fire_mult = 1 + data['lastik_fire_orani'] / 100.0
        kumas_gr_total = sum(s['box_qty'] * s['kumas_gr'] for s in data['sizes']) * kumas_fire_mult
        lastik_mt_total = sum(s['box_qty'] * s['lastik_mt'] for s in data['sizes']) * lastik_fire_mult

        st.markdown("---")
        st.markdown("**Renk Bazlı Malzeme İhtiyacı**")
        fabrics_by_id = {f['id']: f for f in fabrics}
        elastics_by_id = {e['id']: e for e in elastics}
        color_rows = []
        for col in data['colors']:
            fab = fabrics_by_id.get(col.get('kumas_fabric_id'))
            ela = elastics_by_id.get(col.get('lastik_elastic_id'))
            color_rows.append({
                'Renk': col['color_name'],
                'Kumaş': f"{fab['name']} / {fab['icerik']} / {col['kumas_renk']}" if fab else (col['kumas_renk'] or '—'),
                'Kumaş (kg)': round(kumas_gr_total / 1000.0, 3),
                'Lastik': f"{ela['tur']} {ela['ad']} ({ela['boyut']}) / {col['lastik_renk']}" if ela else (col['lastik_renk'] or '—'),
                'Lastik (mt)': round(lastik_mt_total, 1),
                'Aksesuar': col['aksesuar_adi'] or '—',
            })
        st.dataframe(pd.DataFrame(color_rows), width='stretch', hide_index=True)

        st.markdown("---")
        cb1, cb2 = st.columns(2)
        with cb1:
            if st.button("← Geri", width='stretch'):
                st.session_state.order_step = 3
                st.rerun()
        with cb2:
            if st.button("✅ Siparişi Oluştur", width='stretch', type="primary"):
                urun_foto_path = None
                if data.get('_urun_foto_file') is not None:
                    urun_foto_path = None  # set after we know order_id

                order_id, color_ids = db.add_order(
                    data['model_name'], data['gender'], data['urun_grubu'], data['package_size'],
                    data['kumas_fire_orani'], data['lastik_fire_orani'],
                    data['deadline'], data['sizes'], data['colors'], st.session_state.user['id'],
                    para_birimi=data.get('para_birimi', 'TL'),
                    iscilik_birim=data.get('iscilik_birim', 0),
                    genel_gider_yuzde=data.get('genel_gider_yuzde', 10),
                    kar_yuzde=data.get('kar_yuzde', 30),
                    kutu_manufacturer_id=data.get('kutu_manufacturer_id'),
                    kutu_fiyat=data.get('kutu_fiyat', 0),
                )

                if data.get('_urun_foto_file') is not None:
                    path = save_uploaded_file(data['_urun_foto_file'], f"order_{order_id}_urun")
                    db.set_order_photo(order_id, path)

                for i, col in enumerate(data['colors']):
                    color_id = color_ids[i]
                    kf = col.get('_kumas_foto_file')
                    lf = col.get('_lastik_foto_file')
                    af = col.get('_aksesuar_foto_file')
                    kpath = save_uploaded_file(kf, f"order_{order_id}_color_{color_id}_kumas") if kf else None
                    lpath = save_uploaded_file(lf, f"order_{order_id}_color_{color_id}_lastik") if lf else None
                    apath = save_uploaded_file(af, f"order_{order_id}_color_{color_id}_aksesuar") if af else None
                    if kpath or lpath or apath:
                        db.set_color_photos(color_id, kumas_foto=kpath, lastik_foto=lpath, aksesuar_foto=apath)

                db.add_log(st.session_state.user['id'], st.session_state.user['username'],
                          'Yeni Sipariş', f'{data["model_name"]} siparişi oluşturuldu (ID: {order_id})')
                st.success("Sipariş başarıyla oluşturuldu! 'Planlama Bekleyen Siparişler' listesinde görünecek.")
                reset_wizard()
                st.rerun()


# ======================================================================
# SİPARİŞ DETAY GÖRÜNÜMÜ (salt okunur)
# ======================================================================
def render_order_detail_readonly(o):
    st.caption(f"🏷️ Model Kodu: **{o.get('model_kodu') or '—'}**")
    d1, d2, d3, d4 = st.columns(4)
    d1.metric("Toplam Kutu", o['total_boxes'])
    d2.metric("Toplam Adet", o['total_quantity'])
    d3.metric("Kumaş (kg)", o['kumas_kg_total'])
    d4.metric("Lastik (mt)", o['lastik_mt_total'])

    if o.get('urun_foto'):
        show_photo(o['urun_foto'], "Ürün Fotoğrafı")

    costs = db.compute_order_costs(o)
    st.markdown("**💰 Maliyet & Kârlılık**")
    e1, e2, e3, e4 = st.columns(4)
    e1.metric("Birim Maliyet", f"{costs['birim_maliyet_pb']:.2f} {costs['para_birimi']}")
    e2.metric("Satış Fiyatı (birim)", f"{costs['satis_birim_pb']:.2f} {costs['para_birimi']}")
    e3.metric("Toplam Maliyet", f"{costs['toplam_maliyet_pb']:,.0f} {costs['para_birimi']}")
    e4.metric("Toplam Kâr", f"{costs['kar_toplam_pb']:,.0f} {costs['para_birimi']}")
    with st.expander("Maliyet Detayı"):
        st.write(f"Kumaş: {costs['kumas_maliyet']:.2f} · Lastik: {costs['lastik_maliyet']:.2f} · "
                f"Kutu: {costs['kutu_maliyet']:.2f} · Aksesuar: {costs['aksesuar_maliyet']:.2f} · "
                f"İşçilik: {costs['iscilik_maliyet']:.2f} ({costs['para_birimi']})")
        if costs['para_birimi'] != 'TL':
            st.caption(f"TL karşılığı — Maliyet: {costs['toplam_maliyet_tl']:,.0f} ₺ · "
                      f"Satış: {costs['satis_toplam_tl']:,.0f} ₺ · Kâr: {costs['kar_toplam_tl']:,.0f} ₺")

    st.markdown("**📦 Kutu Takibi**")
    kutu_bg = fabric_row_color(o.get('kutu_termin_tarihi'))
    kutu_kalan = (o.get('kutu_siparis_adet') or 0) - (o.get('kutu_gelen_adet') or 0)
    st.markdown(f"""<div style="background:{kutu_bg}; padding:10px 14px; border-radius:8px; font-size:0.86rem;">
        Üretici: {o.get('kutu_manufacturer_name') or '—'} &nbsp;|&nbsp;
        Gerekli: {o.get('total_boxes', 0)} kutu &nbsp;|&nbsp;
        Sipariş: {o.get('kutu_siparis_adet') or 0} &nbsp;|&nbsp;
        Gelen: {o.get('kutu_gelen_adet') or 0} &nbsp;|&nbsp;
        Kalan: {kutu_kalan} &nbsp;|&nbsp;
        Termin: {o.get('kutu_termin_tarihi') or '-'}
        </div>""", unsafe_allow_html=True)

    st.markdown("**Beden Dağılımı**")
    st.dataframe(pd.DataFrame([
        {'Beden': s['size'], 'Kutu': s['box_qty'], 'Kumaş (gr/parça)': s['kumas_gr'],
         'Lastik (adet/parça)': s['lastik_adet'], 'Lastik (mt/parça)': s['lastik_mt']}
        for s in o['sizes']
    ]), width='stretch', hide_index=True)

    st.markdown("**Renk / Malzeme Detayı**")
    for col in o['colors']:
        st.markdown(
            f"🎨 **{col['color_name']}** &nbsp; "
            f"Kumaş: {material_badge(col['kumas_status'])} &nbsp; "
            f"Lastik: {material_badge(col['lastik_status'])} &nbsp; "
            f"Aksesuar: {material_badge(col['aksesuar_status'])}",
            unsafe_allow_html=True)
        cc1, cc2, cc3 = st.columns(3)
        with cc1:
            st.caption(f"**Kumaş:** {col.get('kumas_adi') or '-'} ({col.get('kumas_turu') or '-'}) · "
                       f"{col.get('kumas_icerik') or '-'} · {col.get('kumas_renk') or '-'}\n\n"
                       f"Üretici: {col.get('kumas_manufacturer_name') or '—'} · {col['kumas_kg']} kg gerekli")
            show_photo(col.get('kumas_foto'))
        with cc2:
            st.caption(f"**Lastik:** {col.get('lastik_tur') or '-'} {col.get('lastik_adi') or ''} · "
                       f"{col.get('lastik_genislik') or '-'} · {col.get('lastik_renk') or '-'}\n\n"
                       f"Üretici: {col.get('lastik_manufacturer_name') or '—'} · {col['lastik_mt_toplam']} mt")
            show_photo(col.get('lastik_foto'))
        with cc3:
            if col.get('aksesuar_adi'):
                st.caption(f"**Aksesuar:** {col['aksesuar_adi']} · {col.get('aksesuar_renk') or '-'}\n\n"
                           f"Üretici: {col.get('aksesuar_manufacturer_name') or '—'}")
                show_photo(col.get('aksesuar_foto'))
            else:
                st.caption("Aksesuar yok")

    st.markdown("**Kumaş Tedarik Takibi**")
    render_fabric_tracking_table(o['colors'])


def order_list_page(title, subtitle, assigned_filter, is_admin, include_completed=False, year_filter=False):
    page_header(title, subtitle)

    orders = None
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
                dl_value = datetime.strptime(o['deadline'], '%Y-%m-%d').date() if o.get('deadline') else date.today()
            except Exception:
                dl_value = date.today()
            deadline = st.date_input("Termin Tarihi", value=dl_value)
            kumas_fire = st.number_input("Kumaş Fire Oranı (%)", min_value=0.0, max_value=50.0,
                                          value=float(o['kumas_fire_orani']), step=0.5)
            lastik_fire = st.number_input("Lastik Fire Oranı (%)", min_value=0.0, max_value=50.0,
                                           value=float(o['lastik_fire_orani']), step=0.5)
        if st.form_submit_button("💾 Temel Bilgileri Kaydet", type="primary"):
            db.update_order_basic(order_id, model_name, gender, urun_grubu, deadline.isoformat(),
                                  kumas_fire, lastik_fire)
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

        color_updates = []
        if not full_service:
            st.markdown("**Renk Bazlı Üretici Ataması**")
            for col in o['colors']:
                st.markdown(f"*{col['color_name']}*")
                cc1, cc2, cc3 = st.columns(3)
                with cc1:
                    k_idx = ids.index(col['kumas_manufacturer_id']) if col['kumas_manufacturer_id'] in ids else 0
                    k_sel = st.selectbox("Kumaş Üreticisi", names, index=k_idx, key=f"k_{order_id}_{col['id']}")
                    k_id = ids[names.index(k_sel)]
                with cc2:
                    l_idx = ids.index(col['lastik_manufacturer_id']) if col['lastik_manufacturer_id'] in ids else 0
                    l_sel = st.selectbox("Lastik Üreticisi", names, index=l_idx, key=f"l_{order_id}_{col['id']}")
                    l_id = ids[names.index(l_sel)]
                with cc3:
                    a_idx = ids.index(col['aksesuar_manufacturer_id']) if col['aksesuar_manufacturer_id'] in ids else 0
                    a_sel = st.selectbox("Aksesuar Üreticisi", names, index=a_idx, key=f"a_{order_id}_{col['id']}")
                    a_id = ids[names.index(a_sel)]
                color_updates.append({'id': col['id'], 'kumas_manufacturer_id': k_id,
                                       'lastik_manufacturer_id': l_id, 'aksesuar_manufacturer_id': a_id})

        if st.form_submit_button("💾 Atamayı Kaydet", type="primary"):
            db.update_order_assignment(order_id, new_mf_id, full_service, color_updates)
            db.update_order_status(order_id, new_status)
            db.add_log(st.session_state.user['id'], st.session_state.user['username'],
                      'Sipariş Ataması Güncellendi', f"{o['model_name']} (ID: {order_id})")
            st.success("Kaydedildi!")
            st.rerun()

    # --- Maliyet parametreleri & Kutu takibi ---
    st.markdown('<div class="pk-section-title">💰 Maliyet Parametreleri & 📦 Kutu Takibi</div>', unsafe_allow_html=True)
    manufacturers = db.get_manufacturers()
    with st.form(f"edit_costs_{order_id}"):
        c1, c2, c3 = st.columns(3)
        with c1:
            para_birimi = st.selectbox("Para Birimi", CURRENCIES,
                                       index=CURRENCIES.index(o.get('para_birimi', 'TL')))
            iscilik_birim = st.number_input("İşçilik (birim/parça)", min_value=0.0, value=float(o.get('iscilik_birim') or 0), step=0.5)
        with c2:
            genel_gider_yuzde = st.number_input("Genel Gider (%)", min_value=0.0, value=float(o.get('genel_gider_yuzde') or 10), step=1.0)
            kar_yuzde = st.number_input("Kâr Marjı (%)", min_value=0.0, value=float(o.get('kar_yuzde') or 30), step=1.0)
        with c3:
            names, ids = mf_options(manufacturers, "— Belirtilmedi —")
            k_idx = ids.index(o.get('kutu_manufacturer_id')) if o.get('kutu_manufacturer_id') in ids else 0
            kutu_mf_sel = st.selectbox("Kutu Üreticisi", names, index=k_idx)
            kutu_manufacturer_id = ids[names.index(kutu_mf_sel)]
            kutu_fiyat = st.number_input("Kutu Fiyatı (adet başına)", min_value=0.0, value=float(o.get('kutu_fiyat') or 0), step=0.5)

        st.markdown("**Kutu Sevkiyat Takibi**")
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
            kutu_siparis_tarihi = st.date_input("Sipariş Tarihi", value=ks_val, key="kutu_st")
        with t4:
            try:
                kt_val = datetime.strptime(o['kutu_termin_tarihi'], '%Y-%m-%d').date() if o.get('kutu_termin_tarihi') else date.today()
            except Exception:
                kt_val = date.today()
            kutu_termin_tarihi = st.date_input("Termin Tarihi", value=kt_val, key="kutu_tt")

        if st.form_submit_button("💾 Maliyet & Kutu Bilgisini Kaydet", type="primary"):
            conn = db.get_conn()
            conn.execute('''UPDATE orders SET para_birimi=?, iscilik_birim=?, genel_gider_yuzde=?, kar_yuzde=?
                            WHERE id=?''', (para_birimi, iscilik_birim, genel_gider_yuzde, kar_yuzde, order_id))
            conn.commit()
            conn.close()
            db.update_order_kutu(order_id, kutu_manufacturer_id, kutu_fiyat, kutu_siparis_adet, kutu_gelen_adet,
                                 kutu_siparis_tarihi.isoformat(), kutu_termin_tarihi.isoformat())
            db.add_log(st.session_state.user['id'], st.session_state.user['username'],
                      'Maliyet/Kutu Güncellendi', f"{o['model_name']} (ID: {order_id})")
            st.success("Kaydedildi!")
            st.rerun()

    if st.button("🚚 Kutu İrsaliyesi Oluştur (PDF)", key=f"kutu_irs_{order_id}"):
        st.session_state[f'show_irs_kutu_{order_id}'] = True
    if st.session_state.get(f'show_irs_kutu_{order_id}'):
        render_irsaliye_form(order_id, None, "Kutu", o.get('kutu_manufacturer_name') or '', "adet")

    # --- Renk bazlı malzeme / fotoğraf / kumaş takip ---
    st.markdown('<div class="pk-section-title">Renk / Malzeme / Kumaş Takibi</div>', unsafe_allow_html=True)
    render_fabric_tracking_table(o['colors'])

    fabrics = db.get_fabrics()
    elastics = db.get_elastics()
    fabric_labels = ["— Tanımlı değil —"] + [f"{f['name']} ({f['kumas_turu']})" for f in fabrics]
    fabric_ids = [None] + [f['id'] for f in fabrics]
    elastic_labels = ["— Tanımlı değil —"] + [f"{e['tur']} - {e['ad']} ({e['boyut']})" for e in elastics]
    elastic_ids = [None] + [e['id'] for e in elastics]

    for col in o['colors']:
        with st.expander(f"🎨 {col['color_name']} — malzeme ve kumaş takip detayı"):
            with st.form(f"edit_color_{col['id']}"):
                st.markdown("**Kumaş**")
                r1, r2 = st.columns(2)
                with r1:
                    f_idx = fabric_ids.index(col.get('kumas_fabric_id')) if col.get('kumas_fabric_id') in fabric_ids else 0
                    fsel = st.selectbox("Kumaş", fabric_labels, index=f_idx, key=f"ef_{col['id']}")
                    new_fabric_id = fabric_ids[fabric_labels.index(fsel)]
                    kumas_renk = st.text_input("Kumaş Rengi", value=col.get('kumas_renk') or '', key=f"ekr_{col['id']}")
                with r2:
                    kumas_foto_file = st.file_uploader("Kumaş Fotoğrafı (değiştir)", type=['png', 'jpg', 'jpeg'],
                                                        key=f"ekf_{col['id']}")
                    show_photo(col.get('kumas_foto'), "Mevcut")

                st.markdown("**Kumaş Tedarik Takibi**")
                t1, t2, t3, t4 = st.columns(4)
                with t1:
                    siparis_kg = st.number_input("Sipariş (kg)", min_value=0.0, value=float(col.get('kumas_siparis_kg') or 0),
                                                  key=f"esk_{col['id']}")
                with t2:
                    gelen_kg = st.number_input("Gelen (kg)", min_value=0.0, value=float(col.get('kumas_gelen_kg') or 0),
                                                key=f"egk_{col['id']}")
                with t3:
                    try:
                        st_val = datetime.strptime(col['kumas_siparis_tarihi'], '%Y-%m-%d').date() if col.get('kumas_siparis_tarihi') else date.today()
                    except Exception:
                        st_val = date.today()
                    siparis_tarihi = st.date_input("Sipariş Tarihi", value=st_val, key=f"est_{col['id']}")
                with t4:
                    try:
                        tt_val = datetime.strptime(col['kumas_termin_tarihi'], '%Y-%m-%d').date() if col.get('kumas_termin_tarihi') else date.today()
                    except Exception:
                        tt_val = date.today()
                    termin_tarihi = st.date_input("Termin Tarihi", value=tt_val, key=f"ett_{col['id']}")

                st.markdown("---")
                st.markdown("**Lastik**")
                r3, r4 = st.columns(2)
                with r3:
                    e_idx = elastic_ids.index(col.get('lastik_elastic_id')) if col.get('lastik_elastic_id') in elastic_ids else 0
                    esel = st.selectbox("Lastik", elastic_labels, index=e_idx, key=f"ee_{col['id']}")
                    new_elastic_id = elastic_ids[elastic_labels.index(esel)]
                    lastik_renk = st.text_input("Lastik Rengi", value=col.get('lastik_renk') or '', key=f"elr_{col['id']}")
                with r4:
                    lastik_foto_file = st.file_uploader("Lastik Fotoğrafı (değiştir)", type=['png', 'jpg', 'jpeg'],
                                                         key=f"elf_{col['id']}")
                    show_photo(col.get('lastik_foto'), "Mevcut")

                st.markdown("**Lastik Tedarik Takibi**")
                lt1, lt2, lt3, lt4, lt5 = st.columns(5)
                with lt1:
                    lastik_fiyat = st.number_input("Fiyat (mt)", min_value=0.0, value=float(col.get('lastik_fiyat_mt') or 0), key=f"elp_{col['id']}")
                with lt2:
                    lastik_siparis_mt = st.number_input("Sipariş (mt)", min_value=0.0, value=float(col.get('lastik_siparis_mt') or 0), key=f"elsm_{col['id']}")
                with lt3:
                    lastik_gelen_mt = st.number_input("Gelen (mt)", min_value=0.0, value=float(col.get('lastik_gelen_mt') or 0), key=f"elgm_{col['id']}")
                with lt4:
                    try:
                        ls_val = datetime.strptime(col['lastik_siparis_tarihi'], '%Y-%m-%d').date() if col.get('lastik_siparis_tarihi') else date.today()
                    except Exception:
                        ls_val = date.today()
                    lastik_siparis_tarihi = st.date_input("Sipariş T.", value=ls_val, key=f"elst_{col['id']}")
                with lt5:
                    try:
                        lt_val = datetime.strptime(col['lastik_termin_tarihi'], '%Y-%m-%d').date() if col.get('lastik_termin_tarihi') else date.today()
                    except Exception:
                        lt_val = date.today()
                    lastik_termin_tarihi = st.date_input("Termin", value=lt_val, key=f"eltt_{col['id']}")

                st.markdown("---")
                st.markdown("**Aksesuar**")
                r5, r6 = st.columns(2)
                with r5:
                    aksesuar_adi = st.text_input("Aksesuar Adı", value=col.get('aksesuar_adi') or '', key=f"eaa_{col['id']}")
                    aksesuar_renk = st.text_input("Aksesuar Rengi", value=col.get('aksesuar_renk') or '', key=f"ear_{col['id']}")
                with r6:
                    aksesuar_foto_file = st.file_uploader("Aksesuar Fotoğrafı (değiştir)", type=['png', 'jpg', 'jpeg'],
                                                           key=f"eaf_{col['id']}")
                    show_photo(col.get('aksesuar_foto'), "Mevcut")

                st.markdown("**Aksesuar Tedarik Takibi**")
                at1, at2, at3, at4, at5 = st.columns(5)
                with at1:
                    aksesuar_fiyat = st.number_input("Fiyat (adet)", min_value=0.0, value=float(col.get('aksesuar_fiyat') or 0), key=f"eap_{col['id']}")
                with at2:
                    aksesuar_siparis_adet = st.number_input("Sipariş (adet)", min_value=0.0, value=float(col.get('aksesuar_siparis_adet') or 0), key=f"easa_{col['id']}")
                with at3:
                    aksesuar_gelen_adet = st.number_input("Gelen (adet)", min_value=0.0, value=float(col.get('aksesuar_gelen_adet') or 0), key=f"eaga_{col['id']}")
                with at4:
                    try:
                        as_val = datetime.strptime(col['aksesuar_siparis_tarihi'], '%Y-%m-%d').date() if col.get('aksesuar_siparis_tarihi') else date.today()
                    except Exception:
                        as_val = date.today()
                    aksesuar_siparis_tarihi = st.date_input("Sipariş T.", value=as_val, key=f"east_{col['id']}")
                with at5:
                    try:
                        att_val = datetime.strptime(col['aksesuar_termin_tarihi'], '%Y-%m-%d').date() if col.get('aksesuar_termin_tarihi') else date.today()
                    except Exception:
                        att_val = date.today()
                    aksesuar_termin_tarihi = st.date_input("Termin", value=att_val, key=f"eatt_{col['id']}")

                if st.form_submit_button("💾 Bu Rengi Kaydet", type="primary"):
                    conn = db.get_conn()
                    conn.execute('''UPDATE order_colors SET kumas_fabric_id=?, kumas_renk=?, lastik_elastic_id=?,
                                    lastik_renk=?, aksesuar_adi=?, aksesuar_renk=? WHERE id=?''',
                                (new_fabric_id, kumas_renk, new_elastic_id, lastik_renk,
                                 aksesuar_adi, aksesuar_renk, col['id']))
                    conn.commit()
                    conn.close()
                    db.update_fabric_tracking(col['id'], siparis_kg, gelen_kg,
                                              siparis_tarihi.isoformat(), termin_tarihi.isoformat())
                    if new_elastic_id:
                        db.update_color_lastik_tracking(col['id'], lastik_fiyat, lastik_siparis_mt, lastik_gelen_mt,
                                                        lastik_siparis_tarihi.isoformat(), lastik_termin_tarihi.isoformat())
                    else:
                        db.update_color_lastik_tracking(col['id'], 0, lastik_siparis_mt, lastik_gelen_mt,
                                                        lastik_siparis_tarihi.isoformat(), lastik_termin_tarihi.isoformat())
                    db.update_color_aksesuar_tracking(col['id'], aksesuar_fiyat, aksesuar_siparis_adet, aksesuar_gelen_adet,
                                                      aksesuar_siparis_tarihi.isoformat(), aksesuar_termin_tarihi.isoformat())
                    kpath = save_uploaded_file(kumas_foto_file, f"order_{order_id}_color_{col['id']}_kumas") if kumas_foto_file else None
                    lpath = save_uploaded_file(lastik_foto_file, f"order_{order_id}_color_{col['id']}_lastik") if lastik_foto_file else None
                    apath = save_uploaded_file(aksesuar_foto_file, f"order_{order_id}_color_{col['id']}_aksesuar") if aksesuar_foto_file else None
                    if kpath or lpath or apath:
                        db.set_color_photos(col['id'], kumas_foto=kpath, lastik_foto=lpath, aksesuar_foto=apath)
                    db.add_log(st.session_state.user['id'], st.session_state.user['username'],
                              'Malzeme Güncellendi', f"{o['model_name']} - {col['color_name']} (ID: {order_id})")
                    st.success("Kaydedildi!")
                    st.rerun()

            irs1, irs2, irs3 = st.columns(3)
            with irs1:
                if st.button("🚚 Kumaş İrsaliyesi", key=f"irs_k_{col['id']}"):
                    st.session_state[f'show_irs_kumas_{col["id"]}'] = True
            with irs2:
                if st.button("🚚 Lastik İrsaliyesi", key=f"irs_l_{col['id']}"):
                    st.session_state[f'show_irs_lastik_{col["id"]}'] = True
            with irs3:
                if st.button("🚚 Aksesuar İrsaliyesi", key=f"irs_a_{col['id']}"):
                    st.session_state[f'show_irs_aksesuar_{col["id"]}'] = True
            if st.session_state.get(f'show_irs_kumas_{col["id"]}'):
                render_irsaliye_form(order_id, col['id'], "Kumaş", col.get('kumas_manufacturer_name') or '', "kg")
            if st.session_state.get(f'show_irs_lastik_{col["id"]}'):
                render_irsaliye_form(order_id, col['id'], "Lastik", col.get('lastik_manufacturer_name') or '', "mt")
            if st.session_state.get(f'show_irs_aksesuar_{col["id"]}'):
                render_irsaliye_form(order_id, col['id'], "Aksesuar", col.get('aksesuar_manufacturer_name') or '', "adet")

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
# EXCEL İŞLEMLERİ & YEDEKLEME
# ======================================================================
def page_excel_yedek():
    page_header("📤 Excel İşlemleri & Yedekleme", "Toplu dışa/içe aktarma ve tam sistem yedeği")

    if not HAS_EXCEL:
        st.warning("Excel desteği için `openpyxl` kütüphanesi kurulu değil. requirements.txt içinde mevcuttur, "
                  "ortamınızda `pip install openpyxl` çalıştırmanız gerekebilir.")

    st.markdown("### 📥 Siparişleri Excel Olarak İndir")
    orders = db.get_orders_overview(assigned=None, include_completed=True)
    if orders and HAS_EXCEL:
        rows = []
        for o in orders:
            row = {
                'ID': o['id'], 'Model Kodu': o.get('model_kodu'), 'Ürün Adı': o['model_name'],
                'Cinsiyet': o['gender'], 'Ürün Grubu': o.get('urun_grubu'),
                'Üretici': o.get('manufacturer_name') or '-', 'Durum': o['status'],
                'Termin': o.get('deadline'), 'Paket': o['package_size'],
                'Toplam Kutu': o['total_boxes'], 'Toplam Adet': o['total_quantity'],
                'Kumaş (kg)': o['kumas_kg_total'], 'Lastik (mt)': o['lastik_mt_total'],
                'Para Birimi': o.get('para_birimi'),
            }
            for s in o['sizes']:
                row[f"Adet_{s['size']}"] = s['box_qty'] * o['package_size']
            costs = db.compute_order_costs(o)
            row['Toplam Maliyet'] = costs['toplam_maliyet_pb']
            row['Satış Fiyatı'] = costs['satis_toplam_pb']
            row['Kâr'] = costs['kar_toplam_pb']
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
                                r = recipe.get(size, {'kumas_gr': 0, 'lastik_adet': 0, 'lastik_mt': 0})
                                sizes_payload.append({'size': size, 'box_qty': box_qty, **r})
                        if not sizes_payload:
                            continue
                        colors_payload = [{'color_name': str(row.get('Renk', 'Genel'))}]
                        db.add_order(str(row.get('Ürün Adı', 'Yeni Model')), gender, urun_grubu, package_size,
                                    3.0, 3.0, str(row.get('Termin', date.today().isoformat())),
                                    sizes_payload, colors_payload, st.session_state.user['id'])
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
            for table, rows in data.items():
                z.writestr(f"data/{table}.json", json.dumps(rows, ensure_ascii=False, indent=2, default=str))
            if os.path.isdir(UPLOAD_DIR):
                for fname in os.listdir(UPLOAD_DIR):
                    fpath = os.path.join(UPLOAD_DIR, fname)
                    if os.path.isfile(fpath):
                        z.write(fpath, arcname=f"uploads/{fname}")
        zip_buffer.seek(0)
        st.download_button("⬇️ Zip İndir", zip_buffer.getvalue(),
                          file_name=f"paulkenzie_yedek_{date.today().isoformat()}.zip", mime="application/zip")


# ======================================================================
# AYARLAR (Admin only)
# ======================================================================
def page_ayarlar():
    page_header("⚙️ Ayarlar", "Ana verileri (üretici, ürün grubu, kumaş, lastik, reçete) yönetin")

    tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs(
        ["🏭 Üreticiler", "🏷️ Ürün Grupları", "🧵 Kumaşlar", "➰ Lastikler",
         "📐 Reçeteler", "💱 Kur & Firma", "👤 Kullanıcılar", "🧾 Log"])

    # --- Üreticiler ---
    with tab1:
        with st.form("add_manufacturer"):
            c1, c2 = st.columns(2)
            with c1:
                new_mf = st.text_input("Üretici Adı")
                contact_person = st.text_input("İletişim Kişisi")
            with c2:
                phone = st.text_input("Telefon")
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
                        c1, c2 = st.columns(2)
                        with c1:
                            cp = st.text_input("İletişim Kişisi", value=mf.get('contact_person') or '', key=f"mfcp_{mf['id']}")
                            ph = st.text_input("Telefon", value=mf.get('phone') or '', key=f"mfph_{mf['id']}")
                        with c2:
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
            for gender, names in DEFAULT_PRODUCT_GROUPS.items():
                for name in names:
                    db.add_product_group(gender, name)
            st.success("Varsayılan ürün grupları eklendi!")
            st.rerun()

        with st.form("add_product_group"):
            c1, c2 = st.columns(2)
            with c1:
                pg_gender = st.radio("Cinsiyet", GENDERS, horizontal=True)
            with c2:
                pg_name = st.text_input("Ürün Grubu Adı (ör. String, Boxer)")
            if st.form_submit_button("Ekle", width='stretch', type="primary"):
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

    # --- Kumaşlar ---
    with tab3:
        with st.form("add_fabric"):
            c1, c2, c3 = st.columns(3)
            with c1:
                f_name = st.text_input("Kumaş Adı")
                f_turu = st.selectbox("Kumaş Türü", FABRIC_TYPES)
            with c2:
                f_icerik = st.text_input("İçerik", placeholder="ör. %95 Pamuk %5 Elastan")
                f_en = st.number_input("En (cm)", min_value=0.0, value=0.0, step=1.0)
            with c3:
                f_grm2 = st.number_input("Gramaj (gr/m²)", min_value=0.0, value=0.0, step=1.0)
                f_fiyat = st.number_input("Fiyat (kg başına)", min_value=0.0, value=0.0, step=1.0)
            if st.form_submit_button("Ekle", width='stretch', type="primary"):
                if f_name.strip():
                    db.add_fabric(f_name.strip(), f_icerik.strip(), f_turu, f_en, f_grm2)
                    fabrics_tmp = db.get_fabrics()
                    new_fab = next((f for f in fabrics_tmp if f['name'] == f_name.strip()), None)
                    if new_fab:
                        conn = db.get_conn()
                        conn.execute("UPDATE fabrics SET fiyat=? WHERE id=?", (f_fiyat, new_fab['id']))
                        conn.commit(); conn.close()
                    st.success(f"'{f_name.strip()}' eklendi!")
                    st.rerun()
                else:
                    st.error("Kumaş adı gerekli!")

        st.markdown("---")
        st.markdown("#### Mevcut Kumaşlar")
        fabrics = db.get_fabrics()
        if fabrics:
            df = pd.DataFrame([{'Ad': f['name'], 'Tür': f['kumas_turu'], 'İçerik': f['icerik'],
                                'En (cm)': f['en'], 'Gramaj (gr/m²)': f['gr_m2'],
                                'Fiyat (kg)': f.get('fiyat', 0)} for f in fabrics])
            st.dataframe(df, width='stretch', hide_index=True)
            fsel = st.selectbox("Fiyat güncellenecek/silinecek kumaş", ["—"] + [f['name'] for f in fabrics])
            if fsel != "—":
                fid = next(f['id'] for f in fabrics if f['name'] == fsel)
                cur_fiyat = next(f['fiyat'] for f in fabrics if f['id'] == fid)
                c1, c2 = st.columns(2)
                with c1:
                    new_fiyat = st.number_input("Yeni Fiyat (kg başına)", min_value=0.0, value=float(cur_fiyat or 0), key="fab_new_fiyat")
                    if st.button("💾 Fiyatı Güncelle"):
                        conn = db.get_conn()
                        conn.execute("UPDATE fabrics SET fiyat=? WHERE id=?", (new_fiyat, fid))
                        conn.commit(); conn.close()
                        st.success("Güncellendi!")
                        st.rerun()
                with c2:
                    st.write("")
                    st.write("")
                    if st.button("🗑️ Seçili Kumaşı Sil"):
                        db.delete_fabric(fid)
                        st.rerun()
        else:
            st.info("Henüz kumaş tanımlanmamış.")

    # --- Lastikler ---
    with tab4:
        with st.form("add_elastic"):
            c1, c2, c3 = st.columns(3)
            with c1:
                e_tur = st.selectbox("Lastik Türü", ELASTIC_TYPES)
                e_ad = st.text_input("Lastik Adı")
            with c2:
                e_boyut = st.text_input("Boyut", placeholder="ör. 2 cm")
                e_fiyat = st.number_input("Fiyat (mt başına)", min_value=0.0, value=0.0, step=0.1)
            with c3:
                all_groups = db.get_product_groups()
                e_urun_grubu = st.selectbox("Model / Ürün Grubu (opsiyonel)",
                                            ["—"] + [g['name'] for g in all_groups])
            if st.form_submit_button("Ekle", width='stretch', type="primary"):
                if e_ad.strip():
                    db.add_elastic(e_tur, e_ad.strip(), e_boyut.strip(),
                                   '' if e_urun_grubu == '—' else e_urun_grubu)
                    elastics_tmp = db.get_elastics()
                    new_el = next((e for e in elastics_tmp if e['ad'] == e_ad.strip() and e['tur'] == e_tur), None)
                    if new_el:
                        conn = db.get_conn()
                        conn.execute("UPDATE elastics SET fiyat=? WHERE id=?", (e_fiyat, new_el['id']))
                        conn.commit(); conn.close()
                    st.success(f"'{e_ad.strip()}' eklendi!")
                    st.rerun()
                else:
                    st.error("Lastik adı gerekli!")

        st.markdown("---")
        st.markdown("#### Mevcut Lastikler")
        elastics = db.get_elastics()
        if elastics:
            df = pd.DataFrame([{'Tür': e['tur'], 'Ad': e['ad'], 'Boyut': e['boyut'],
                                'Fiyat (mt)': e.get('fiyat', 0),
                                'Model/Ürün Grubu': e['urun_grubu'] or '—'} for e in elastics])
            st.dataframe(df, width='stretch', hide_index=True)
            esel = st.selectbox("Fiyat güncellenecek/silinecek lastik", ["—"] + [f"{e['tur']} - {e['ad']}" for e in elastics])
            if esel != "—":
                eid = next(e['id'] for e in elastics if f"{e['tur']} - {e['ad']}" == esel)
                cur_fiyat = next(e['fiyat'] for e in elastics if e['id'] == eid)
                c1, c2 = st.columns(2)
                with c1:
                    new_fiyat = st.number_input("Yeni Fiyat (mt başına)", min_value=0.0, value=float(cur_fiyat or 0), key="ela_new_fiyat")
                    if st.button("💾 Fiyatı Güncelle"):
                        conn = db.get_conn()
                        conn.execute("UPDATE elastics SET fiyat=? WHERE id=?", (new_fiyat, eid))
                        conn.commit(); conn.close()
                        st.success("Güncellendi!")
                        st.rerun()
                with c2:
                    st.write("")
                    st.write("")
                    if st.button("🗑️ Seçili Lastiği Sil"):
                        db.delete_elastic(eid)
                        st.rerun()
        else:
            st.info("Henüz lastik tanımlanmamış.")

    # --- Reçeteler ---
    with tab5:
        st.caption("Cinsiyet + Ürün Grubu + Beden bazında kumaş (gr) ve lastik (adet/mt) reçetesi tanımlayın. "
                  "Yeni sipariş oluşturulurken bu değerler otomatik uygulanır.")
        r_gender = st.radio("Cinsiyet", GENDERS, horizontal=True, key="recipe_gender")
        groups = db.get_product_groups(r_gender)
        if not groups:
            st.warning(f"Önce '{r_gender}' için Ürün Grupları sekmesinden bir ürün grubu ekleyin.")
        else:
            r_urun_grubu = st.selectbox("Ürün Grubu", [g['name'] for g in groups], key="recipe_ug")
            existing = db.get_recipe(r_gender, r_urun_grubu)

            with st.form(f"recipe_form_{r_gender}_{r_urun_grubu}"):
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
                        lm = st.number_input("Lastik (mt)", min_value=0.0, value=float(existing[size]['lastik_mt']),
                                             step=0.05, key=f"rm_{size}")
                        size_vals[size] = (kg, la, lm)
                if st.form_submit_button("💾 Reçeteyi Kaydet", type="primary"):
                    for size, (kg, la, lm) in size_vals.items():
                        db.upsert_recipe(r_gender, r_urun_grubu, size, kg, la, lm)
                    db.add_log(st.session_state.user['id'], st.session_state.user['username'],
                              'Reçete Güncellendi', f"{r_gender} - {r_urun_grubu}")
                    st.success("Reçete kaydedildi!")
                    st.rerun()

        st.markdown("---")
        st.markdown("#### Tanımlı Reçeteler")
        recipe_groups = db.get_recipe_groups()
        if recipe_groups:
            for rg in recipe_groups:
                with st.expander(f"{rg['gender']} — {rg['urun_grubu']}"):
                    r = db.get_recipe(rg['gender'], rg['urun_grubu'])
                    df = pd.DataFrame([{'Beden': s, 'Kumaş (gr)': r[s]['kumas_gr'],
                                        'Lastik (adet)': r[s]['lastik_adet'], 'Lastik (mt)': r[s]['lastik_mt']}
                                       for s in SIZES])
                    st.dataframe(df, width='stretch', hide_index=True)
                    if st.button("🗑️ Bu Reçeteyi Sil", key=f"delrec_{rg['gender']}_{rg['urun_grubu']}"):
                        db.delete_recipe_group(rg['gender'], rg['urun_grubu'])
                        st.rerun()
        else:
            st.info("Henüz reçete tanımlanmamış.")

    # --- Kur & Firma ---
    with tab6:
        settings = db.get_settings()
        with st.form("settings_form"):
            c1, c2, c3 = st.columns(3)
            with c1:
                usd_kur = st.number_input("USD Kur (₺)", min_value=0.0, value=float(settings['usd_kur']), step=0.1)
                eur_kur = st.number_input("EUR Kur (₺)", min_value=0.0, value=float(settings['eur_kur']), step=0.1)
            with c2:
                varsayilan_para = st.selectbox("Varsayılan Para Birimi", CURRENCIES,
                                               index=CURRENCIES.index(settings['varsayilan_para']))
                firma_adi = st.text_input("Firma Adı", value=settings['firma_adi'])
            with c3:
                barkod_prefix = st.text_input("Barkod / Model Kodu Öneki", value=settings['barkod_prefix'])
                st.caption("Model kodu şu formatta oluşturulur: PREFIX-001-RENK")
            if st.form_submit_button("💾 Ayarları Kaydet", type="primary"):
                db.update_settings(usd_kur, eur_kur, varsayilan_para, firma_adi.strip(), barkod_prefix.strip())
                db.add_log(st.session_state.user['id'], st.session_state.user['username'],
                          'Ayarlar Güncellendi', 'Kur ve firma bilgileri')
                st.success("Kaydedildi!")
                st.rerun()

    # --- Kullanıcılar ---
    with tab7:
        with st.form("add_user"):
            col1, col2 = st.columns(2)
            with col1:
                new_username = st.text_input("Kullanıcı Adı")
                new_fullname = st.text_input("Ad Soyad")
            with col2:
                new_password = st.text_input("Şifre", type="password")
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
    with tab8:
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

    # Sipariş düzenleme, seçilen sayfa ne olursa olsun önceliklidir (ayrı sayfa gibi davranır)
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
