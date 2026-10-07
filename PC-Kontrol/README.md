# PC Kontrol — Katılımsız Windows 11 kontrolü

## v1.2 güncellemesi — sesli ve görünür alarm

Mevcut kurulumu güncellemek için yeni ZIP'i çıkart ve `PC-Kontrol/Guncelle.cmd` dosyasını **alıcı ve gönderici bilgisayarlarda** çalıştır. Kurulu gönderim anahtarları ve ayarlar korunur; yeniden eşleştirme gerekmez. Alıcı yeniden başlatılır, göndericide yeni kontrol yapılır. Alıcıdaki `Alici/Alarm-Test.cmd` dosyası yerel ses ve ekran alarmını denetir; ağa test raporu göndermez.

Yeni, kaydedilmiş sorun raporunda alıcıda bilgisayar adını ve sorunları gösteren bir pencere açılır; onaylanana kadar 15 saniyede bir ses çalar. Aynı raporun yeniden gönderimi ikinci alarm üretmez. Alarm ayrı süreçte çalışır; pencere açıkken alıcı yeni raporları almaya devam eder. Windows kullanıcı oturumunun açık olması gerekir; ses sessize alınmışsa görünür uyarı yine çıkar.

Ses aygıtı COM arayüzünün kimliği düzeltildi. Klavye/fare algılamasında PnP yanında Raw Input kontrolü eklendi. Uzak oturumda fiziksel giriş cihazı doğrulanamıyorsa `check_unavailable` gösterilir. Varsayılan ses çıkışı sağlıklıysa bilerek devre dışı bırakılan isteğe bağlı ses adaptörünün yalnızca **22** kodu alarm oluşturmaz; diğer hata kodları korunur ve raporda görünür.

Windows oturumu açılınca arka planda kısa CPU/GPU ve cihaz/sürücü kontrolleri yapar. **Sorun yoksa rapor göndermez.** Sorun varsa `http://192.168.1.77:8777/pc-kontrol/rapor/` adresine gönderir. Alıcı kapalıysa rapor yerelde kalır ve sonraki çalıştırmada tekrar denenir.

## İndir ve kur

1. [ZIP dosyasını indir](https://github.com/kampusarena-lgtm/kamp-s2/archive/refs/heads/main.zip) ve tamamen çıkart; içindeki **PC-Kontrol** klasörünü aç.
2. **Bir kez `Eslesmeyi-Hazirla.cmd` dosyasını çalıştır.** Bu adım gönderici ve alıcı için aynı özel gönderim anahtarını oluşturur. Oluşan anahtarın kendisi gösterilmez.
3. Hazırladığın **aynı klasörden** `Alici` klasörünü **192.168.1.77** bilgisayarına kopyala. `Aliciyi-Kur.cmd` dosyasını çalıştır. Windows, HTTP adres kaydı ve yerel ağdan TCP 8777 portuna giriş için yönetici izni ister.
4. Hazırladığın **aynı klasörden** `Gonderici` klasörünü kontrol edilecek Windows 11 bilgisayarlara kopyala. `Kur-ve-Otomatik-Baslat.cmd` dosyasını çalıştır.

Gönderici ve alıcı klasörlerini ayrı indirip bağımsız anahtar oluşturma; eşleşmeleri gerekir. Kurulumdan sonra kullanıcı oturumlarında otomatik çalışırlar. Gerçek Windows servisi değildir; alıcı kullanıcısı oturum açmış olmalıdır. Ayrıntılar: [KULLANIM.txt](KULLANIM.txt).

## Kontroller

- **CPU:** 3 saniyelik tam sayı ve SHA-256 doğrulaması; en fazla 2 iş parçacığı.
- **GPU:** Varsayılan donanımsal Direct3D 11 adaptöründe küçük doku yükleme/kopyalama/okuma ve veri karşılaştırması. Yazılımsal WARP'a geri dönüş yapılmaz.
- **Klavye/fare:** Windows'ta şu anda mevcut aygıtlar ve sürücü hataları.
- **Ses:** Etkin ve varsayılan Windows ses çıkışı; ses aygıtı sürücü hataları.

Fiziksel tuşların, fare düğmelerinin ve kulaklık sesinin gerçekten çalışması katılımsız kesin doğrulanamaz. Test edilmedikleri için bu işlevler otomatik sağlam veya arızalı sayılmaz. Sıcaklık ve kapsamlı kararlılık/VRAM testi yapılmaz. `check_unavailable` testin çalıştırılamadığını gösterir; fiziksel arıza teşhisi değildir.

## Raporlar ve durdurma

Alıcıdaki raporlar `Belgeler\PC-Kontrol-Gelen-Raporlar` altında HTML ve JSON olarak kaydedilir. Gönderilemeyen raporlar göndericide `%LOCALAPPDATA%\PC-Kontrol\Bekleyen-Raporlar` altında kalır. Tekrarlanan aynı rapor ikinci kez kaydedilmez.

Göndericide `Otomatik-Baslatmayi-Kapat.cmd`, alıcıda `Aliciyi-Durdur.cmd` otomatik çalışmayı kapatır. Hazırlanmış `Ayarlar.json` dosyaları özel anahtar içerir; herkese açık GitHub'a geri yükleme. Depodaki şablonların anahtar alanı bilerek boştur. HTTP alıcısı yerel ağ için hazırlanmıştır; HTTPS sunucusu değildir.

## Doğrulama

HTTP alıcı/gönderici, sorunsuz durumda gönderim yapılmaması, kuyruk ve yeniden deneme, yanlış onayda raporun korunması, çift rapor önleme, HTML kaçışları, CPU doğrulaması ve bozulmuş GPU desen karşılaştırması PowerShell 7/Linux üzerinde kontrol edildi. Kaynak C# 5 sözdizimiyle derlendi. Windows 11 Direct3D/CoreAudio/PnP, Windows açılış kurulumu ve gerçek `192.168.1.77` erişimi henüz fiziksel Windows donanımında doğrulanmadı.
