# Outlook Firma Rehberi v5

Windows ve macOS için masaüstü uygulaması. Kullanıcının izin verdiği Microsoft 365, Outlook.com ve Hotmail hesaplarındaki seçili posta klasörlerini tarar; e-posta adreslerini tekilleştirir ve firma bazlı bir Excel rehberi oluşturur.

## Neler yapar?

- Bir veya daha fazla Microsoft hesabını aynı taramada kullanır.
- Gelen, Gönderilen, Silinen ve özel klasörleri; alt klasörleri ne kadar derin olursa olsun listeler.
- Seçilen klasörlerdeki eski ve yeni tüm mesajları tarar. Sonraki taramalarda güvenli ilerleme kaydını kullanır.
- Gönderen, Alıcı, Bilgi ve Gizli alıcı adreslerini işler.
- Kullanıcının kendi adreslerini ve tanımlı takma adreslerini rehbere eklemez.
- Ad, unvan, firma ve telefon bilgilerini gönderenin güncel imzasından zenginleştirir.
- Aynı e-posta adresini hesaplar ve klasörler arasında yalnızca bir kez yazar.
- İşlem tamamlanınca önce `Tarama tamamlandı` mesajını, ardından dosya adı ve konumu seçilebilen yerel `Farklı Kaydet` penceresini açar.

Uygulama e-posta göndermez, silmez, taşımaz veya değiştirmez. Yalnızca Microsoft Graph `Mail.Read` ve oturum profili için `User.Read` yetkilerini ister. Microsoft, `Mail.Read` yetkisini oturum açan kullanıcının posta kutusunu okumaya yönelik bir temsilci izni olarak tanımlar; hem kişisel hem iş/okul hesaplarında kullanılabilir. [Microsoft Graph izin açıklaması](https://learn.microsoft.com/en-us/graph/permissions-reference#mailread)

## Desteklenen hesaplar

Kaynak posta kutusu olarak Microsoft 365 iş/okul hesapları, Outlook.com ve Hotmail desteklenir. Gmail, Yahoo, Yandex ve benzeri adresler kaynak hesap olarak bağlanmaz; mesajlarda bulunan kişisel adresler olarak `Kişisel Mailler` sayfasına ayrılır.

## Excel çıktısı

Tek bir `.xlsx` dosyası oluşturulur:

1. `Özet`
2. `Polisan-Poliport`
3. `Gümrük-YGM`
4. `Diğer Kurumsal`
5. `Kişisel Mailler`

Detay sayfalarında sütunlar `Ad Soyad`, `Görev / Unvan`, `Domain`, `Firma`, `E-posta`, `Cep Telefonu`, `Sabit Telefon`, `Kaynak Hesap(lar)` ve `Kaynak Klasör(ler)` şeklindedir. Aynı domain/firma adresleri art arda sıralanır. Çok satırlı firma bloklarında Firma ve uygun olduğunda Domain hücreleri dikey birleştirilir. Polisan ve Poliport aynı çalışma sayfasında, ayrı firma blokları olarak yer alır.

## Kurulum ve kullanım

### Windows 10/11 x64

1. [Releases](../../releases/latest) sayfasından `OutlookFirmaRehberi_v..._Windows_x64_Setup.exe` dosyasını indirin.
2. Kurulumu çalıştırın ve masaüstü veya Başlat menüsü kısayolundan açın.
3. `Hesap Ekle` ile Microsoft oturumunu açın.
4. Hesapları ve taranacak klasörleri işaretleyin.
5. `Taramayı Başlat` düğmesine basın.
6. Tarama tamamlanınca Excel dosyasının adını ve kaydedileceği klasörü seçin.

### macOS 12 ve sonrası

1. Apple Silicon Mac için `macOS_AppleSilicon.dmg`, Intel Mac için `macOS_Intel.dmg` dosyasını indirin.
2. DMG'yi açın ve `Outlook Firma Rehberi.app` uygulamasını Uygulamalar klasörüne taşıyın.
3. Uygulamayı açıp Windows ile aynı hesap, klasör ve tarama adımlarını izleyin.

İmzasız deneme paketlerinde macOS Gatekeeper veya Windows SmartScreen uyarı gösterebilir. Resmî, uyarısız dağıtım için yayın yöneticisinin kod imzalama ve Apple noter onayı bilgilerini GitHub'a eklemesi gerekir. İmzalanmamış bir dosyanın güvenli olduğu varsayılmamalıdır; SHA-256 dosyasını yayın sayfasındaki değerle karşılaştırın.

## Tarama kesilirse

`İptal Et` ile tarama güvenli sınırda durdurulur. Hata alan veya yarım kalan klasörün ilerleme kaydı değiştirilmez ve otomatik Excel kaydı açılmaz. Tamamlanmış klasörlerin yerel verisi korunur; sorunlu klasör daha sonra yeniden taranabilir.

## v4 verisi

Uygulama, yanında bulduğu eski `rehber_data.json` dosyasını yalnızca bir kez ve dosyayı silmeden içe aktarır. İstenirse `Eski v4 Verisini İçe Aktar` düğmesiyle dosya elle seçilebilir. Eski Windows COM uygulaması artık varsayılan giriş noktası değildir; `uygulama.py` yalnızca v5 uyumluluk başlatıcısıdır. v4.6 kaynakları Git geçmişinde korunur.

## Gizlilik ve yönetici kurulumu

- [Gizlilik ve yerel veri kullanımı](docs/PRIVACY.md)
- [Microsoft Entra uygulama kaydı ve yayın ayarları](docs/MICROSOFT_SETUP.md)

## Geliştirme

Python 3.11 veya sonrası gerekir.

```bash
python -m venv .venv
.venv/bin/python -m pip install -e ".[test,build]"
QT_QPA_PLATFORM=offscreen .venv/bin/python -m pytest
```

Windows PowerShell'de sanal ortam yürütülebilir yolu `.venv\Scripts\python.exe` olur.

Yerel kaynak açılış kontrolü:

```bash
.venv/bin/python scripts/smoke_launch.py --source
```

GitHub Actions, `v*` etiketi gönderildiğinde testlerden sonra Windows x64 kurulumunu, Apple Silicon ve Intel macOS DMG dosyalarını ve SHA-256 özetlerini üretir. Yayın için önce [Microsoft kurulum belgesindeki](docs/MICROSOFT_SETUP.md) depo değişkeni yapılandırılmalıdır.
