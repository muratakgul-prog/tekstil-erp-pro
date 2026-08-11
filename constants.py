# --- Sabitler / Constants ---

# Beden aralığı: sadece XS - 3XL arası
SIZES = ['XS', 'S', 'M', 'L', 'XL', 'XXL', '3XL']

# Ürün cinsiyeti
GENDERS = ['Kadın', 'Erkek']

# Kutu içi (paket) adedi seçenekleri
PACKAGE_SIZES = [1, 3, 5, 7, 10]

# Lastik çeşitleri
ELASTIC_TYPES = ['Baskılı', 'Jakarlı', 'Raporlu']

# Kumaş türleri
FABRIC_TYPES = ['Örme', 'Dokuma']

# Örnek ürün grubu önerileri (Ayarlar'da serbest girilebilir, bunlar sadece ilk kurulum önerisi)
DEFAULT_PRODUCT_GROUPS = {
    'Kadın': ['String', 'Cheeky', 'Slip', 'Boyshort', 'Short', 'Pantolon', 'Triangle Bra'],
    'Erkek': ['Boxer', 'Slip', 'Short', 'Pantolon'],
}

# Sipariş üretim durumu
ORDER_STATUSES = ['Planlama', 'Kumaş Bekleniyor', 'Lastik Bekleniyor', 'Kesimde', 'Dikimde',
                   'Ütü-Paket', 'Kutu Bekleniyor', 'Kargoda', 'Tamamlandı', 'İptal']
# Bu durumlar "aktif" sayılmaz (dashboard/liste filtrelerinde)
CLOSED_STATUSES = ['Tamamlandı', 'İptal']

# Para birimleri
CURRENCIES = ['TL', 'USD', 'EUR']

# İrsaliye kategorileri
IRSALIYE_KATEGORILERI = ['Kumaş', 'Lastik', 'Kutu', 'Aksesuar']

# Malzeme tamamlanma durumu
MATERIAL_STATUS_DONE = 'Tamamlandı'
MATERIAL_STATUS_PENDING = 'Bekliyor'
MATERIAL_STATUS_NA = 'Yok'

# Atama tipi
ASSIGNMENT_FULL_SERVICE = 'tam_hizmet'   # üretici tüm süreçlerle ilgilenir
ASSIGNMENT_PARTIAL = 'kismi'             # üretici sadece ürünle ilgilenir, malzemeler ayrı atanır
ASSIGNMENT_UNPLANNED = 'planlama'        # henüz üretici atanmamış

STATUS_COLORS = {
    'Planlama': '#94a3b8',
    'Üretimde': '#2563eb',
    'Tamamlandı': '#16a34a',
}

MATERIAL_STATUS_COLORS = {
    MATERIAL_STATUS_DONE: '#16a34a',
    MATERIAL_STATUS_PENDING: '#d97706',
    MATERIAL_STATUS_NA: '#94a3b8',
}
