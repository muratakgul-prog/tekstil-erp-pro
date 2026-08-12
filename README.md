# 🧥 Paul Kenzie ERP — Üretim & Sipariş Planlama Sistemi

Tekstil üretimi için uçtan uca sipariş, reçete, malzeme tedariği ve maliyet/kârlılık takip sistemi. Streamlit tabanlı, SQLite veritabanı kullanır.

## ✨ Özellikler

- **Kullanıcı yönetimi**: Admin / kullanıcı rolleri, şifreli giriş.
- **Ürün Adları**: Model isimleri (ör. "Zebra Boxer") tanımlanır; Yeni Sipariş'te seçildiğinde cinsiyet ve ürün grubu otomatik dolar.
- **Reçete kütüphanesi**: Cinsiyet + Ürün Grubu + Beden bazında kumaş (gr) ve lastik (adet/cm) reçetesi — yeni siparişte otomatik uygulanır, düzenlenebilir/silinebilir.
- **Ürün grupları**: Kadın (string, cheeky, slip, boyshort, short, pantolon, triangle bra) ve Erkek (boxer, slip, short, pantolon) için ayrı ayrı tanımlanabilir.
- **Kumaş & Lastik ana verisi**: Kumaş türü (Örme/Dokuma), lastik adı (Baskılı/Jakarlı/Raporlu Lastik), boyut, ürün adı bağlantısı.
- **4 adımlı sipariş sihirbazı**: Temel Bilgiler → Sipariş Adedi (beden bazlı) → Kumaş/Lastik/Aksesuar (birbirinden bağımsız, dinamik satırlar) → Onay.
- **Kumaş, Lastik, Aksesuar birbirinden tamamen bağımsız**: Her biri varsayılan 1 satırla başlar, "+ Satır Ekle" ile farklı renk/malzeme kullanıldığında çoğaltılabilir.
- **Raporlu Lastik seçiminde otomatik adet takibi**, Baskılı/Jakarlı lastikte cm takibi.
- **Üretici ataması sadece düzenleme sayfasında**: "Tüm süreçlerle ilgilenecek" işaretlenirse kumaş/lastik/aksesuar otomatik o üreticiye atanır ve tamamlandı sayılır.
- **Kumaş, Lastik, Aksesuar & Kutu tedarik takibi**: Gerekli/Sipariş/Gelen/Kalan miktar, sipariş/termin tarihi, gün bazlı renk kodlama (15+ gün yeşil, 1-14 turuncu, 0/geçmiş kırmızı).
- **İrsaliye + PDF**: Kumaş/Lastik/Kutu/Aksesuar için gelen mal kaydı girildiğinde kalan miktar otomatik düşer; PDF irsaliye indirilebilir.
- **Fotoğraf yükleme**: Ürün, kumaş, lastik, aksesuar için ayrı ayrı görsel.
- **Excel içe/dışa aktarma** ve **tam sistem yedeği (zip)**.
- **Dashboard**: Aktif/Atanmış/Planlama Bekleyen/Acil/Toplam sipariş sayıları + yıllık tamamlanan sipariş özeti.
- **Sipariş kartları**: Toplam Adet, Toplam Kutu, Kumaş, Lastik kutucukları — tamamlanma ve termin durumuna göre yeşil/sarı/kırmızı, hafif gölgeli.
- **Türkçe tarih formatı** (GG.AA.YYYY) tüm ekranlarda.

## 🔐 Varsayılan Giriş

```
Kullanıcı adı: admin
Şifre: admin123
```

**Önemli:** İlk girişten sonra Ayarlar > Kullanıcılar bölümünden admin şifresini değiştirin.

## 🚀 Kurulum ve Çalıştırma

### 1) Yerel (Local)

```bash
pip install -r requirements.txt
streamlit run app.py
```

Tarayıcıda `http://localhost:8501` açılır.

### 2) Streamlit Cloud (ücretsiz, önerilen)

1. Bu depoyu GitHub'a yükleyin (`app.py`, `database.py`, `constants.py`, `requirements.txt` aynı klasörde).
2. https://share.streamlit.io adresinden "New app" ile bu repoyu seçin.
3. Deploy edin — birkaç dakika içinde `https://uygulamaniz.streamlit.app` linkiniz hazır olur.

> Not: Streamlit Cloud'un dosya sistemi geçicidir (redeploy'da sıfırlanabilir). Düzenli olarak Ayarlar > Excel & Yedek sayfasından **Yedek Zip** alıp saklamanız önerilir.

### 3) Docker

```bash
docker build -t paulkenzie-erp .
docker run -p 8501:8501 -v $(pwd)/data:/app/data -v $(pwd)/uploads:/app/uploads paulkenzie-erp
```

### 4) Kendi Sunucunuz (VPS)

```bash
nohup streamlit run app.py --server.port 8501 --server.address 0.0.0.0 &
```

## 📁 Proje Yapısı

```
paulkenzie_erp/
├── app.py                  # Streamlit arayüzü (tüm sayfalar)
├── database.py              # SQLite veri katmanı (tüm CRUD + hesaplama fonksiyonları)
├── constants.py              # Sabitler (bedenler, cinsiyet, durum listeleri vb.)
├── requirements.txt
├── Dockerfile
├── README.md
├── .streamlit/
│   └── config.toml           # Tema ve sunucu ayarları
├── app_data.db                # SQLite veritabanı (ilk çalıştırmada otomatik oluşur)
└── uploads/                   # Yüklenen fotoğraflar (ilk çalıştırmada otomatik oluşur)
```

## 💾 Veri Kalıcılığı

- Tüm veriler `app_data.db` (SQLite) dosyasında tutulur.
- Yüklenen fotoğraflar `uploads/` klasöründe saklanır.
- Bu iki konum GitHub'a **yüklenmemelidir** (`.gitignore` içinde hariç tutulmuştur) — her ortamda kendi verisiyle başlar.
- Düzenli yedek almak için: Sol menü → **Excel & Yedek** → **Yedek Zip Oluştur**.

## 🛠️ Teknoloji

- [Streamlit](https://streamlit.io) — arayüz
- SQLite — veritabanı
- pandas, openpyxl — Excel işlemleri
- reportlab — PDF irsaliye oluşturma

## 📄 Lisans

Bu proje şirket içi kullanım için özel olarak geliştirilmiştir.
