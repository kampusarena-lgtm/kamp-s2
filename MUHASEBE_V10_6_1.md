# Muhasebe V10.6.1 — Toplu gün kilidi açma

[Verileri koruyan güncelleme ZIP'ini indir](https://raw.githubusercontent.com/kampusarena-lgtm/kamp-s2/main/KampusArena_V10_6_1_VERI_KORUYAN_GUNCELLEME.zip)

Bu paket mevcut muhasebe kurulumu içindir. **KURULUM.bat çalıştırmayın.**

1. Mevcut muhasebe sunucusunu tamamen kapatın; yalnız tarayıcıyı kapatmak yeterli değildir.
2. ZIP'i mevcut muhasebe klasöründen farklı bir klasöre çıkarın.
3. `GUNCELLE.bat` çalıştırın ve mevcut muhasebe klasörünü seçin.
4. Program kapalıysa önce eski program dosyaları, veritabanı, Excel ve mevcut anahtarlar yedeklenir. Yalnız `app.py` ve `templates/day_close.html` güncellenir.
5. Mevcut klasörünüzdeki `BASLAT.bat` ile açın.

Elle güncelleme ve geri dönüş adımları paketteki `ONCE_OKUYUN_GUNCELLEME.txt` dosyasında.

## Kullanım

Ana yönetici hesabıyla **Gün kilidi → Toplu gün kilidi açma** bölümünden başlangıç, bitiş ve açma nedenini girin. Başlangıç/bitiş dahil en fazla 366 geçmiş gün seçilebilir.

Açma öncesinde SQLite backup API ile veritabanı yedeği alınır; yedekleme başarısızsa kilitler değişmez. Finansal ve personel kayıtları korunur. Kilit değişiklikleri ve gün bazında işlem geçmişi birlikte kaydedilir; hata halinde tüm işlem geri alınır.

Açılan günler yalnız mevcut muhasebe günü boyunca açık kalır; ertesi muhasebe günü 09:00'da yeniden kilitlenir. Bugün zaten açık olan günler atlanır.

## Doğrulama

17 muhasebe/güvenlik testi geçti. Kayıtların birebir korunduğu, yedek, yetki/CSRF, hatada geri alma ve süresi geçen açmanın yenilenmesi kontrol edildi. Masaüstü ve telefon tarayıcısında gün sayısı, onay ve toplu açma kontrol edildi.

Windows PowerShell güncelleme betiği Linux ortamında çalıştırılmadı; manuel güncelleme alternatifi pakette anlatılıyor.
