# Outlook Firma Rehberi web sürümü

Amaç: Var olan masaüstü programının temel hesabı/klasörü seçme, posta adreslerini tekilleştirme, imzadan kişi bilgisi çıkarma ve beş sayfalı Excel oluşturma akışını kurulum gerektirmeyen bir bağlantıya taşımak.

Karar: GitHub Pages üzerinde tarayıcı uygulaması. Microsoft resmi MSAL kütüphanesiyle OAuth / PKCE ve Graph yalnızca Mail.Read + User.Read temsilci izinleri. Microsoft uygulama kimliği ve SPA yönlendirmesi tamamlanmadan hesap bağlantısı tamamlandı sayılmaz. Masaüstü dosyaları korunur.

Alternatifler: Sunucuda Python uygulaması barındırmak kaynakları tekrar kullanır fakat sunucu ve posta verisi depolama ihtiyacı yaratır. Seçilen tarayıcı sürümü postaları doğrudan Microsoft'tan okur; üçüncü taraf sunucuya veya GitHub'a rehber göndermez.

Birden çok Microsoft hesabı, tüm derinliklerde klasör ve gizli klasör listesi, hesap takma adreslerini dışlama, sayfalama, istek sınırı için bekleme, iptal, tamamlanan klasörlerde artımlı yeniden tarama, rehber düzenleme, eski v4 JSON ve web çalışma yedeği içe aktarma. Klasör tamamen bitmeden ilerleme ve kişi kayıtları işlenmiş olarak kaydedilmez. Eksik sonuç açıkça işaretlenir.

Erişim belirteçleri ve rehber bellekte tutulur. Sayfa yenilendiğinde hesaplar tekrar bağlanır; kullanıcı isterse yerel çalışma yedeği alarak kişi/ilerleme verilerini korur. Yedeklerde erişim belirteci veya parola bulunmaz. Microsoft giriş ayarı dışında hiçbir veri localStorage'a yazılmaz.

Excel: Özet, Polisan-Poliport, Gümrük-YGM, Diğer Kurumsal, Kişisel Mailler. Ad Soyad, Görev/Unvan, Domain, Firma, E-posta, Cep Telefonu, Sabit Telefon, Kaynak Hesaplar, Kaynak Klasörler. Firma ve domain hücreleri gruplar halinde birleştirilir. Tüm dış kaynak metinleri Excel'de metin olarak yazılır.

Eksik dış yapılandırma: Client ID ve web yönlendirmesinin Microsoft Entra'da tanımlanması. Gerçek posta kutusu taraması kullanıcı oturumu / izin onayı olmadan test edilmez.
