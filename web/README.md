# Outlook Firma Rehberi — Web

Var olan masaüstü uygulamasının yanında çalışan, kurulum gerektirmeyen tarayıcı sürümü.

## Özellikler

- Microsoft 365, Outlook.com ve Hotmail ile Microsoft'un resmi giriş penceresi.
- Birden fazla hesap ve bütün derinliklerde posta klasörü seçimi.
- Kendi e-posta / takma adreslerini hariç tutma.
- Gönderen, Alıcı, Bilgi ve Gizli alıcı adreslerini tekilleştirme.
- Yalnız güncel gönderen imzasından ad, unvan, firma ve telefon çıkarma; alıntı mesajları dışlama.
- Klasör tamamlanmadan sonuç ve delta ilerleme kaydını tamamlandı saymama.
- Rehber önizlemesi ve ad, unvan, firma düzeltme.
- Beş sayfalı, firma/domain grupları birleştirilmiş gerçek .xlsx çıktısı.
- Eski v4 JSON dosyası ve web çalışma yedeği içe aktarma.

## Gizlilik ve masaüstünden farkı

Posta içerikleri ve rehber yalnızca bu sekmenin belleğinde işlenir. GitHub'a veya bir uygulama sunucusuna gönderilmez. Microsoft parola ve erişim belirteçleri Excel veya çalışma yedeğine eklenmez. MSAL erişim belirteçlerini memoryStorage içinde tutar. Sayfa yenilendiğinde hesapları yeniden bağlamak gerekir; kullanıcı çalışma yedeğini yerel JSON dosyası olarak alarak rehberini ve tarama ilerlemesini sonraki oturuma taşıyabilir. Tamamlanmış klasörlerdeki artımlı tarama bu yedekle de devam eder. Yedek, kişi bilgileri içerdiği için kullanıcının kendi cihazında korunmalıdır.

Bu sürüm paylaşılan sunucu veritabanı veya arka planda zamanlanmış tarama içermez. Tarama sırasında sekme açık kalmalıdır. Yarım tarama sonucu otomatik olarak tam çıktı sayılmaz; kısmi çıktı için ekrandaki seçim gerekir.

## Microsoft'un bir kez yapılacak bağlantı ayarı

1. Microsoft Entra uygulama kaydı oluşturun veya mevcut masaüstü kaydını kullanın.
2. Hesap türünü kişisel hesaplarla kuruluş hesaplarını destekleyecek şekilde seçin.
3. Authentication altında **Single-page application (SPA)** platformuna tam şu dönüş adresini ekleyin:

   `https://shtseymen-eng.github.io/outlook-firma-rehberi/auth.html`

4. Temsilci izinleri **User.Read** ve **Mail.Read** olmalıdır. Gönderme veya posta değiştirme izni gerekmez. Kurum politikası gerektiriyorsa yöneticisi onaylamalıdır.
5. Application (client) ID değerini sitedeki **Microsoft bağlantı kurulumu** alanına girin. Bu değer bir uygulama tanımlayıcısıdır; gizli anahtar değildir. Microsoft parolası veya Client Secret girilmez.
6. İsterseniz GitHub Actions repository variable `OUTLOOK_REHBERI_CLIENT_ID` ile yayın sırasında varsayılan kimlik verin.

Microsoft ayarı tamamlanmadan site hesap bağlayamaz; bunu ekranda açıkça gösterir. Hesap girişi ve Microsoft onay ekranları kullanıcı tarafından tamamlanır.

Kaynaklar: https://learn.microsoft.com/en-us/entra/msal/javascript/browser/initialization ve https://learn.microsoft.com/en-us/graph/api/message-delta?view=graph-rest-1.0

## Geliştirme

`npm ci` → `npm test` → `npm run build`

Vite geliştirme sayfası: `http://127.0.0.1:8767/outlook-firma-rehberi/`

MSAL v5 için auth.html ve resmi redirect-bridge paketi uygulamayla aynı kaynaktan sunulur. Auth yanıtı URL fragment'inde işlenir. Kimlik köprüsünde üçüncü taraf CDN veya analiz kodu bulunmaz.
