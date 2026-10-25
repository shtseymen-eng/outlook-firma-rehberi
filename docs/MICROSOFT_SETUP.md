# Microsoft Entra ve yayın kurulumu

Bu belge uygulamayı kendi kullanıcılarına dağıtacak yönetici içindir. Son kullanıcıların Microsoft parolasını uygulamaya veya GitHub'a vermesi gerekmez.

## 1. Microsoft Entra uygulamasını kaydedin

1. [Microsoft Entra yönetim merkezinde](https://entra.microsoft.com/) `Identity > Applications > App registrations > New registration` bölümünü açın.
2. Uygulamayı `Outlook Firma Rehberi` adıyla kaydedin.
3. Herkesin kullanacağı genel sürüm için kuruluş hesaplarıyla birlikte kişisel Microsoft hesaplarını da destekleyen hesap türünü seçin. Yalnız şirket içi kullanımda kuruluş politikanıza uygun tek kiracılı seçeneği kullanabilirsiniz.
4. `Authentication > Add a platform > Mobile and desktop applications` bölümünden sistem tarayıcısı yönlendirmesi olarak tam `http://localhost` adresini ekleyin. MSAL Python'ın etkileşimli masaüstü oturumu bu yönlendirmeyi kullanır. [Microsoft masaüstü uygulaması yapılandırması](https://learn.microsoft.com/en-us/entra/identity-platform/scenario-desktop-app-registration)
5. `Allow public client flows` seçeneğini etkinleştirin. Bu uygulama gizli anahtar taşıyan bir sunucu uygulaması değil, kullanıcının cihazında çalışan bir genel istemcidir. [MSAL Python genel istemci açıklaması](https://learn.microsoft.com/en-us/entra/msal/python/getting-started/client-applications)
6. `Overview` sayfasındaki `Application (client) ID` değerini kopyalayın. Client ID bir uygulama tanımlayıcısıdır; istemci sırrı değildir.

## 2. Yalnızca gerekli temsilci izinlerini ekleyin

`API permissions > Add a permission > Microsoft Graph > Delegated permissions` altında şunları ekleyin:

- `Mail.Read`: oturum açan kullanıcının posta kutusunu okumak için.
- `User.Read`: oturum açmak ve kullanıcının temel profilini/hesap adresini okumak için.

`Mail.ReadWrite`, `Mail.Send`, `Mail.Send.Shared` veya Application permissions eklemeyin. Bu programın posta silme, değiştirme ya da gönderme özelliği yoktur. `Mail.Read` temsilci izni kişisel Microsoft hesapları ile iş/okul hesaplarında geçerlidir. [Microsoft Graph izin başvurusu](https://learn.microsoft.com/en-us/graph/permissions-reference#mailread)

Kuruluşunuz kullanıcı onayını kapattıysa Entra yöneticisi bu iki temsilci izni için yönetici onayı vermelidir. Uygulama bu durumda kullanıcıya yönetici onayı gerektiğini bildirir.

## 3. Client ID'yi yapılandırın

### GitHub yayınları

GitHub deposunda `Settings > Secrets and variables > Actions > Variables` altında şu değişkeni ekleyin:

- `OUTLOOK_REHBERI_CLIENT_ID`: Entra'dan kopyalanan Application (client) ID.

Yayın hattı bu değeri paket içindeki `app_config.json` dosyasına yazar. Depoda istemci sırrı tutulmaz. Değişken eksikse yayın işi bilerek durur; hesap bağlanamayan bir paket oluşturmaz.

### Yerel geliştirme

macOS/Linux:

```bash
export OUTLOOK_REHBERI_CLIENT_ID="00000000-0000-0000-0000-000000000000"
python -m outlook_rehberi
```

Windows PowerShell:

```powershell
$env:OUTLOOK_REHBERI_CLIENT_ID = "00000000-0000-0000-0000-000000000000"
python -m outlook_rehberi
```

## 4. İsteğe bağlı imzalama bilgileri

İmzalama bilgileri yoksa iş akışı test edilmiş fakat imzasız paket üretir. Üretim dağıtımı için aşağıdaki GitHub Actions secrets değerlerini ekleyin.

Windows:

- `WINDOWS_CERTIFICATE_BASE64`
- `WINDOWS_CERTIFICATE_PASSWORD`

macOS:

- `APPLE_CERTIFICATE_BASE64`
- `APPLE_CERTIFICATE_PASSWORD`
- `APPLE_SIGNING_IDENTITY`
- `APPLE_ID`
- `APPLE_TEAM_ID`
- `APPLE_APP_SPECIFIC_PASSWORD`

İmzalama sertifikalarını veya parolaları repository variable olarak ya da kaynak dosyaya yazmayın.

## 5. Yayın oluşturun

1. Değişken ve gerekiyorsa secrets değerlerini kaydedin.
2. `v5.0.0` gibi bir etiket gönderin.
3. GitHub Actions üç ortamda testleri çalıştırır: Windows x64, macOS Apple Silicon ve macOS Intel.
4. Test ve pencere açmadan başlatma kontrolleri geçerse `.exe`, iki `.dmg` ve bunların `.sha256` dosyaları GitHub Release'e eklenir.

GitHub'ın güncel standart koşucu tablosunda `macos-15` ARM64, `macos-15-intel` Intel olarak tanımlanır. [GitHub-hosted runners başvurusu](https://docs.github.com/en/actions/reference/runners/github-hosted-runners)

## Sorun giderme

- `AADSTS65001`, `consent_required` veya `interaction_required`: Kiracı yöneticisi izin onayını kısıtlıyor olabilir.
- Oturum tarayıcısı dönmüyor: Uygulama kaydında Mobile and desktop platformunda tam `http://localhost` yönlendirmesini kontrol edin.
- Paket yapılandırma eksik diyor: GitHub variable veya yerel ortam değişkenindeki Client ID boş olabilir.
- Bir hesap görünmüyor: `Hesap Ekle` ile o hesabı ayrıca seçin; uygulama başka bir kullanıcının posta kutusuna kendiliğinden erişmez.
