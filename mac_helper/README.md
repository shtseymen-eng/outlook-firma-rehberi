# Mac Mail Rehberi yardımcısı

Microsoft uygulama kimliği gerektirmeden bu Mac'teki Mail hesaplarını okur. Outlook uygulamasına özel hesapları okuyamaz; hesabın Apple Mail'e eklenmiş olması gerekir. Mail'de erişilemeyen veya henüz yüklenmemiş mesajlar için tam tarama garantisi yoktur.

## Kullanım

1. Web sayfasındaki `Mac-Mail-Rehberi.zip` dosyasını açın.
2. `Mac Mail Rehberi.app` uygulamasını açın. macOS 12+ ve Python 3.9+ gerekir (python.org veya Homebrew kurulumu). Hazır paket Apple Silicon (M1 ve sonrası) Mac içindir.
3. Uygulama henüz Apple tarafından noter onaylı değildir. macOS engellerse yalnızca bu depodan aldığınız sürüm için Sistem Ayarları > Gizlilik ve Güvenlik > Yine de Aç adımını kendiniz tamamlayın.
4. Açılan tarayıcı ekranında **Mac Mail hesaplarını göster** düğmesine basın. macOS Mail erişim izni sorarsa kendiniz izin verin. Reddedildiyse Sistem Ayarları > Gizlilik ve Güvenlik > Otomasyon bölümünden kontrol edin.
5. Klasörleri seçin, **Taramayı Başlat** düğmesine basın. İsteğe bağlı imza taraması firma/telefon bilgilerini de arar ve daha uzun sürer.
6. Sonucu **Excel'i Farklı Kaydet** ile `.xlsx` kaydedin. Kişiler aynı e-posta adresine göre tekilleştirilir. Kendi hesap adresleriniz dışlanır. Yanlış çıkarılan firma/ad bilgileri önizlemeden düzenlenebilir.
7. Sonra devam etmek için çalışma yedeği kaydedin. Sekmenin belleği kapanınca temizlenir. Yeniden açmak için yardımcı penceresindeki **Rehberi Aç** düğmesini kullanın.
8. İşiniz bitince **Yardımcıyı Kapat** düğmesine basın; yerel sunucu kapanır.

## Veri ve tarama davranışı

- Yalnızca 127.0.0.1 üzerinde rastgele port açılır; şirket ağına veya internete dinleme yapılmaz.
- Rehber ekranının tamamı uygulamaya dahildir ve yerel sunulur. Posta verisi GitHub'a gönderilmez.
- Her açılışta üretilen gizli anahtar URL parçasından belleğe alınır; anahtar disk, yedek veya günlükte tutulmaz. API yalnızca tam yerel Origin/Host ve anahtarla yanıt verir. Harici web sitelerine CORS yoktur.
- Mail'e yalnızca sabit okuma komutları gönderilir. Posta gönderme, silme, okundu durumunu değiştirme veya hesap ayarı değiştirme yoktur.
- Her tarama klasördeki mevcut mesaj kimliklerinin anlık listesini alıp 25'erli okur. Aynı klasör tekrar taranır; aynı e-posta adresleri birleştirilir. Yerel taramada Microsoft delta ilerleme kaydı kullanılmaz.
- Tamamlanamayan/iptal edilen klasörün yeni sonuçları rehbere katılmaz. Önceki tamamlanan klasörler korunur. Eksik çıktı açıkça işaretlenir.
- Rehber birikimlidir: Mail'den silinen mesajlar daha önce çıkarılmış kişileri kendiliğinden silmez. Temiz rehber için oturumdaki rehberi temizleyip yeniden tarayın.

## Geliştirme

Web bağımlılıklarını kurduktan sonra `python3 mac_helper/build_app.py <çıktı-klasörü>` ile paketleyin. Xcode command line tools gerekir. Derleme çalıştığı Mac’in mimarisini üretir ve ad-hoc imza atar; bu Apple noter onayı değildir.

Kontroller: `npm test --prefix web`, `python3 -m unittest discover -s mac_helper -p 'test_*.py'`. Python testleri sahte Mail okuyucusu kullanır. Gerçek Mail otomasyon testi ayrıca kullanıcının macOS iznini gerektirir.
