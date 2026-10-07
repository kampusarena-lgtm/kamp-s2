# Kampüs Arena Muhasebe V10.6

Windows için güncellenmiş muhasebe uygulaması.

## İndir

[Kurulum paketini indir](https://raw.githubusercontent.com/kampusarena-lgtm/kamp-s2/main/KampusArena_Muhasebe_V10_6.zip)

## Kurulum

1. ZIP'i yeni ve boş bir klasöre çıkarın.
2. Python 3.9+ ve masaüstü Microsoft Excel kurulu olmalıdır.
3. KURULUM.bat dosyasını çalıştırıp Excel yedek şifresini belirleyin.
4. BASLAT.bat dosyasını çalıştırın.
5. Aynı bilgisayarda http://127.0.0.1:8788 adresinden yönetici hesabı oluşturun.
6. Ayarlar'dan personel ve maaş bilgilerini ekleyin.

Ayrıntılı açıklamalar paketteki ONCE_BUNU_OKUYUN.txt dosyasında.

## Değişiklikler

- Eşzamanlı maaş ve yemek ödemelerinde çift kayıt koruması.
- Geçersiz tutar doğrulaması ve veritabanında günlük kasa tekilliği.
- Kurulum başına oturum anahtarı, CSRF koruması ve veritabanından yönetici yetkisi kontrolü.
- Pasif personelin geçmiş hesapları ve aylara göre maaş tabanı.
- Telefon uyumlu yan menü, işlem sekmeleri, kasa toplamı ön izlemesi ve ilk kullanım yönlendirmesi.

## Doğrulama

13 otomatik muhasebe/güvenlik testi ve masaüstü/telefon tarayıcı kontrolleri geçti.
Testler: `py tests\test_application.py`.

Gerçek PanCafe/Firebird bağlantısı ve Windows/Excel şifrelemesi Linux ortamında çalıştırılmadı; Windows bilgisayarında doğrulanmalıdır.

Paket temiz kurulum içindir; işletme veritabanı, eski Excel anahtarı, oturum anahtarı ve test hesapları içermez.

ZIP SHA-256: `d91252bd6a40053bd4a554c0755fd1c9bd3f3d730c5099b8c78df056209871b7`
