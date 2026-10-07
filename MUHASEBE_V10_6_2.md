# Muhasebe V10.6.2 — Otomatik çıkış ve planlı yedekleme

[Verileri koruyan güncellemeyi indir](https://raw.githubusercontent.com/kampusarena-lgtm/kamp-s2/main/KampusArena_V10_6_2_VERI_KORUYAN_GUNCELLEME.zip)

V10.6 veya V10.6.1 üzerine uygulanır; toplu gün kilidi açma da dahildir.

## Güncelleme

1. Muhasebe sunucusunu tamamen kapatın; yalnız tarayıcıyı kapatmak yeterli değildir.
2. ZIP'i mevcut programdan ayrı bir klasöre çıkarın.
3. GUNCELLE.bat çalıştırıp mevcut muhasebe klasörünü seçin.
4. Mevcut dosyalar ve veriler yedeklenir. Yalnız beş program dosyası güncellenir; veritabanı, Excel ve anahtarlar korunur.
5. Mevcut klasörden BASLAT.bat ile açıp tekrar giriş yapın.

**KURULUM.bat çalıştırmayın.** Elle güncelleme alternatifi ZIP içindeki ONCE_OKUYUN_GUNCELLEME.txt dosyasında.

## Yönetici oturumu

10 dakika gerçek kullanıcı etkinliği olmazsa otomatik çıkış yapılır. Süre sunucuda da kontrol edilir; otomatik log yenilemesi ve durum sorguları süreyi uzatmaz. Klavye, fare, form girişi ve kaydırma etkinliği süreyi yeniler. Aynı oturumdaki sekmeler ortak süreyi kullanır. Sunucu yeniden başlatılınca mevcut yönetici oturumları kapanır.

## Yedekleme

Program açıkken her saat ve Türkiye saatiyle her gün **09:05** yedek alınır. Son **24 saatlik** ve **30 günlük** yedek `backups/planli` altında tutulur; eski sürümlerin yedekleri bu temizliğe dahil edilmez.

Program kapalıyken yedek alınamaz. Açıldığında son günlük slot için o anın yedeği alınır; eski 09:05 anı yeniden oluşturulmaz.

SQLite backup API ile WAL dahil tutarlı veritabanı snapshot alınır ve yeni arşiv doğrulanır. Excel senkronu çalışmasa da veritabanı yedeği alınır. Varsa son şifreli Excel ve kurtarma anahtarları eklenir; asıl veri veritabanıdır.

**Ayarlar → Planlı yedekleme** ekranından manuel yedek alınabilir ve yedekler indirilip ayrı bir diske kopyalanabilir. Bu işlemler yalnız ana yöneticiye açıktır. Manuel yedekler otomatik temizlenmez.

## Doğrulama

22 muhasebe/güvenlik/yedekleme testi geçti. Tarayıcıda gerçek klavye etkinliği, manuel yedek indirme ve hızlandırılmış otomatik çıkış doğrulandı.

Windows PowerShell güncelleme betiği ve Excel/PanCafe bileşenleri Linux ortamında çalıştırılmadı; manuel güncelleme adımları pakette verilmiştir.
