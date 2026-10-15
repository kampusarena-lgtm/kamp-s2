# Muhasebe V10.6.3 — Kasa 09:00, kayıt kilidi 12:00

[Güncellemeyi indir](https://raw.githubusercontent.com/kampusarena-lgtm/kamp-s2/main/KampusArena_V10_6_3_VERI_KORUYAN_GUNCELLEME.zip)

Kasa ve PanCafe dönemi **09:00–09:00** olarak kalır. Kapanan güne ait kayıtlar aynı gün **12:00'ye kadar** girilir/düzeltilir; saat 12:00'de kilitlenir.

Örnek: 15 Ekim 09:00'da tamamlanan 14 Ekim kasası, 15 Ekim 12:00'ye kadar girilebilir.

**Önceden kilitlenmiş geçmiş günler açılmaz.** Mevcut kayıtlar korunur. Ana yöneticinin açıkça yeniden açtığı günler için mevcut izin devam eder.

09:00–12:00 arasında kasa formunda kapanan günün tarihi seçilir; yeni günün gider/veresiye formları kendi tarihinde kalır.

## Güncelleme

V10.6, V10.6.1 veya V10.6.2 üzerine uygulanır. Önceki toplu gün açma, otomatik yönetici çıkışı ve planlı yedekleme özellikleri dahildir.

1. Muhasebe sunucusunu tamamen kapatın.
2. ZIP'i mevcut programdan **ayrı bir klasöre** çıkarın.
3. GUNCELLE.bat çalıştırıp mevcut muhasebe klasörünü seçin.
4. Önce program dosyaları, veritabanı, Excel ve mevcut anahtarlar yedeklenir.
5. Yalnız yedi program dosyası güncellenir; mevcut veritabanı dosyasının SHA-256 değeri önce/sonra karşılaştırılır.
6. Mevcut klasörünüzdeki BASLAT.bat ile yeniden açın.

**KURULUM.bat çalıştırmayın; mevcut muhasebe klasörünü silmeyin.**

Elle güncelleme ve geri dönüş açıklamaları ZIP içindeki ONCE_OKUYUN_GUNCELLEME.txt dosyasındadır. Saatler mevcut Windows bilgisayarının yerel saatine göre çalışır.

## Doğrulama

[GitHub Actions kontrolleri](https://github.com/kampusarena-lgtm/kamp-s2/actions/workflows/muhasebe-noon-lock.yml): yeni saat kuralı için 12, mevcut muhasebe işlemleri için 13 test.

09:00/11:59:59/12:00 sınırları, 09:00–09:00 PanCafe dönemi, geç kayıt ve düzenleme, bekleyen PanCafe, eski kilitlerin korunması, ay/yıl sınırları kontrol edildi. Yeni bir şema veya geçmiş günleri açan veri göçü eklenmedi.

Windows güncelleme betiği ile gerçek PanCafe/Excel bağlantısı bu Linux kontrollerinde çalıştırılmadı.
