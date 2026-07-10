# Outlook Mail Prioritizer

Tatilden/uzun bir aradan döndüğünde Outlook gelen kutunda yüzlerce okunmamış
mail birikmiş oluyor. Bu script, okunmamış mailleri **sana ne kadar doğrudan
hitap ettiklerine** göre otomatik olarak renkli Outlook kategorilerine ayırıp
önceliklendirmeni sağlar — hiçbir mail silinmez ya da taşınmaz, sadece
etiketlenir.

After a vacation (or any long absence), your Outlook inbox is flooded with
hundreds of unread emails. This script automatically tags each unread email
with a color-coded Outlook category based on **how directly it's addressed
to you** — nothing is ever deleted or moved, only tagged.

## Nasıl önceliklendiriyor? / How it prioritizes

| Kategori / Category | Ne zaman? / When | Öncelik / Priority |
|---|---|---|
| 🔴 `1-@Isim` | Gerçekten **@mention** edilmişsin / You were genuinely **@mentioned** | En yüksek / Highest |
| 🟠 `2-Isim Geciyor` | Konu/gövdede adın düz metin olarak geçiyor / Your name appears as plain text | Yüksek / High |
| 🟡 `3-Bana (To)` | **To** alanındasın / You're in the **To** field | Orta / Medium |
| 🔵 `4-Bilgi (CC)` | **CC** alanındasın / You're in the **CC** field | En düşük / Lowest |

Outlook'ta gelen kutusunu **"Arrange by → Categories"** ile gruplarsan
öncelik sırasına göre düzenlenmiş halini görürsün:

Grouping your inbox by **"Arrange by → Categories"** in Outlook shows them
sorted by priority:

![1-@Isim ve 2-Isim Geciyor kategorileri](images/oncelik-kategorileri.png)
![3-Bana (To) kategorisi](images/to-kategorisi.png)
![4-Bilgi (CC) kategorisi](images/cc-kategorisi.png)

## Neden "gerçek" @mention tespiti önemli? / Why "real" mention detection matters

Outlook'ta birini `@isim` yazıp kişi seçiciden seçtiğinde, gönderilen mailde
`@` karakteri kaybolur; Outlook bunu senin adresine giden bir `mailto:` linkine
çevirir. Bu yüzden script düz metinde `@isim` aramak yerine, gövdedeki
`mailto:senin-adresin` linkini arayarak gerçek mention'ları yakalıyor
(elle yazılan `@isim` metnini de ayrıca destekliyor).

When you type `@name` in Outlook and pick someone from the people-picker, the
`@` character disappears from the sent message — Outlook turns it into a
`mailto:` link pointing to that person's address instead. So instead of
searching for literal `@name` text, the script detects real mentions via that
`mailto:` link in the HTML body (manually typed `@name` text is also still
supported as a fallback).

## Kurulum / Setup

```bash
pip install pywin32
```

1. Outlook masaüstü uygulamasını aç ve hesabına giriş yap.
   Open desktop Outlook and sign in to your account.
2. `outlook_prioritize.py` içindeki **KULLANICI AYARLARI / USER SETTINGS**
   bölümüne kendi adını(varyasyonlarını) yaz.
   Fill in your own name (and variants) in the **USER SETTINGS** section.
3. Önce `DRY_RUN = True` ile çalıştır, özet çıktısını kontrol et.
   Run once with `DRY_RUN = True` first and review the summary.
4. Sonuç doğruysa `DRY_RUN = False` yapıp tekrar çalıştır.
   If it looks right, set `DRY_RUN = False` and run again.

```bash
python outlook_prioritize.py
```

Windows'ta çift tıklayarak çalıştırmak için `Calistir.bat` dosyasını
kullanabilirsin (Python'un PATH'te olması yeterli, ekstra kurulum gerekmez).

On Windows you can just double-click `Calistir.bat` to run it (only requires
Python to be on PATH, no extra setup).

## Güvenlik notları / Safety notes

- Script hiçbir maili **silmez veya taşımaz**, sadece Outlook kategorisi
  ekler. Geri almak için kategoriyi maildeon kaldırman yeterli.
  The script never **deletes or moves** anything — it only adds an Outlook
  category. To undo, just remove the category from the email.
- Kurumsal/Exchange hesaplarında Outlook, bir programın posta verilerine
  eriştiğini fark ettiğinde bir izin penceresi açabilir — bu pencere ana
  pencerenin arkasında gizli kalabilir, script "donmuş" gibi görünürse
  kontrol et.
  On corporate/Exchange accounts, Outlook may pop up a security prompt when
  a program accesses mail data — this window can hide behind others; check
  for it if the script seems to hang.
- Kurumsal antivirüs/EDR yazılımları derlenmiş `.exe` dosyalarını
  engelleyebilir; bu repo sadece kaynak kodu (`.py`) içerir, exe'yi kendin
  derlemen gerekir (örn. [PyInstaller](https://pyinstaller.org/)).
  Corporate antivirus/EDR may block compiled `.exe` files; this repo only
  ships source code — build your own exe if needed (e.g. with
  [PyInstaller](https://pyinstaller.org/)).

## Gereksinimler / Requirements

- Windows + Outlook masaüstü uygulaması (klasik/desktop app, yeni Outlook
  değil) / Windows + classic desktop Outlook (not the new Outlook)
- Python 3.9+
- [pywin32](https://pypi.org/project/pywin32/)

## Lisans / License

[MIT](LICENSE)
