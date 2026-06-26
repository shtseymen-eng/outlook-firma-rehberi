# 📋 Outlook Firma Rehberi

**Tüm Outlook klasörlerini tarayarak firma bazlı kişi rehberi oluşturur.**

Poliport dahili kullanım — S.SEYMEN

---

## ✅ Özellikler

| Özellik | Açıklama |
|---------|----------|
| **Tüm klasörler** | Gelen + Gönderilen + Silinmiş + tüm alt klasörler |
| **Çok hesap** | Outlook'taki tüm hesaplar aynı anda taranır |
| **İmza analizi** | Firma adı, Ad Soyad, Görev, Cep Tel, Sabit Tel |
| **Firma tespiti** | İmzada varsa oradan (Ltd/A.Ş.), yoksa @domain'den |
| **Tekrarsız** | Aynı mail adresi yalnızca 1 kez — bilgiler zenginleştirilir |
| **Artımlı tarama** | Son tarihten bu yana gelen mailler eklenir |
| **Kişisel ayrımı** | Gmail/Hotmail vb. ayrı bölümde gösterilir |
| **Excel çıktısı** | Özet + Kurumsal + Kişisel (3 sayfa) |

---

## 🚀 Kurulum

### Son sürümü indir
**[Releases](../../releases/latest)** sayfasından `OutlookFirmaRehberi_vX.X_Setup.exe` indirin.

1. `Setup.exe` → çift tıkla → kurulum sihirbazı
2. Masaüstü kısayolundan aç
3. **Tarama Yap** butonuna tıkla

> Outlook'un açık ve oturum açık olması gerekiyor.

---

## 🖥️ Ekran Görüntüsü

```
┌─────────────────────────────────────────────────────────────────────┐
│  📋 Outlook Firma Rehberi            [Firma:42] [Kişi:318] [Toplam:401]│
├──────────────────┬──────────────────────────────────────────────────┤
│ KURUMSAL FİRMALAR│  AKDENIZ LOJISTIK  ·  12 kişi  ·  akdeniz.com.tr│
│                  ├──────┬──────────────────┬──────────┬─────────────┤
│ AKDENIZ          │ İSİM │ GÖREV            │ E-POSTA  │ TELEFON     │
│ BORUSAN          │ A.K  │ Satın Alma Müd.  │ a@akd... │ 📱0532...  │
│ COSCO            │ M.Y  │ Lojistik Uzmanı  │ m@akd... │ ☎ 0212...  │
│ DHL              │ ...  │ ...              │ ...      │ ...         │
│ EKOL             ├──────┴──────────────────┴──────────┴─────────────┤
│ ...              │                                                   │
└──────────────────┴───────────────────────────────────────────────────┘
```

---

## 🔨 Build (Geliştiriciler için)

### GitHub Actions (Otomatik)
```bash
git tag v4.1
git push origin v4.1
```
→ Actions otomatik `Setup.exe` build eder, Releases'a yükler.

### Manuel Build
```bash
pip install pywin32 openpyxl pyinstaller
pyinstaller OutlookFirmaRehberi.spec --clean
# Sonra Inno Setup ile setup.iss derle
```

---

## 📁 Dosya Yapısı

```
outlook-firma-rehberi/
├── uygulama.py                    ← Ana uygulama (motor + GUI)
├── OutlookFirmaRehberi.spec       ← PyInstaller build ayarları
├── setup.iss                      ← Inno Setup kurulum scripti
├── version_info.txt               ← EXE metadata
├── assets/
│   └── icon.ico                   ← (Actions tarafından oluşturulur)
└── .github/workflows/
    └── build.yml                  ← CI/CD pipeline
```

---

*Poliport — S.SEYMEN*
