# Gizlilik ve yerel veri kullanımı

Outlook Firma Rehberi, kullanıcının açıkça bağladığı Microsoft hesaplarında seçtiği posta klasörlerini okur ve yerel bir Excel rehberi oluşturur.

## Erişim sınırı

- Uygulama yalnızca temsilci `Mail.Read` ve `User.Read` izinlerini ister.
- Uygulama e-posta göndermez, silmez, taşımaz, işaretlemez veya değiştirmez.
- Bir hesap ancak kullanıcı Microsoft oturum ekranında seçip izin verdiğinde bağlanır.
- Birden fazla hesap bağlandığında her hesap ayrı oturum olarak listelenir ve kullanıcı taranacak klasörleri seçer.

Microsoft'un `Mail.Read` tanımı, uygulamanın oturum açan kullanıcının posta kutusunu okuyabilmesine izin verir. Bu izin yazma veya gönderme izni değildir. [Microsoft Graph izin başvurusu](https://learn.microsoft.com/en-us/graph/permissions-reference#mailread)

## İşlenen veriler

Tarama sırasında mesajın göndereni, Alıcı/Bilgi/Gizli alıcıları, tarihleri ve metin/HTML gövdesi bellekte işlenir. Tam mesaj gövdesi veya ekler yerel veritabanına kaydedilmez. Rehber için aşağıdaki türetilmiş bilgiler saklanabilir:

- normalize edilmiş e-posta adresi;
- ad, unvan ve firma;
- cep ve sabit telefon numaraları;
- ayıklanan imza metni;
- kaynak hesap ve klasör bilgisi;
- klasörün güvenli ilerleme bağlantısı.

Bu yerel kayıt, işletim sisteminin uygulama verisi klasöründeki `rehber.sqlite3` dosyasında tutulur. Uygulama telemetri, reklam veya kendi sunucusuna veri gönderme kodu içermez. Ağ trafiği Microsoft oturumu ve Microsoft Graph posta okuma istekleri içindir.

## Oturum bilgileri

Microsoft erişim/yenileme belirteçlerinin MSAL önbelleği düz dosyada veya SQLite veritabanında tutulmaz. Şifreli işletim sistemi kimlik deposu kullanılır:

- macOS Keychain;
- Windows Credential Manager.

Uygulama Microsoft parolasını görmez veya saklamaz. Oturum açma sistem tarayıcısında Microsoft tarafından yürütülür.

## Excel dosyası

Excel dosyası yalnızca kullanıcı `Farklı Kaydet` penceresinde bir konum seçtiğinde yazılır. Dosya seçilen klasörde kalır; otomatik olarak karşıya yüklenmez. Excel'in paylaşılması veya bulut klasörüne kaydedilmesi kullanıcının kontrolündedir.

## Veriyi temizleme

- `Yerel Rehberi Temizle`, çıkarılmış kişi verisini ve klasör ilerleme kayıtlarını siler; hesap bağlantıları ve kendi adresi hariç tutma ayarları korunur.
- `Seçili Hesabı Ayır`, o Microsoft hesabını uygulamanın oturum önbelleğinden kaldırır.
- Tam kaldırma için uygulamayı kaldırdıktan sonra işletim sisteminin uygulama verisi klasöründeki `OutlookFirmaRehberi` verisini ve gerekirse işletim sistemi kimlik deposundaki `OutlookFirmaRehberi` girdisini kullanıcı veya yönetici silebilir.

## v4 içe aktarımı

Eski `rehber_data.json` dosyası bulunduğunda içe aktarım bir kez yapılır. Kaynak dosya değiştirilmez veya silinmez. Kullanıcı dosyayı ayrıca saklamak ya da silmek konusunda karar verir.

## Kurumsal kullanım

Kuruluş yöneticileri Entra kullanıcı onayı, koşullu erişim ve veri saklama politikalarını uygulayabilir. Bu uygulamanın yerel çıktısı da kuruluşun kişisel veri ve bilgi güvenliği politikalarına göre korunmalıdır.
