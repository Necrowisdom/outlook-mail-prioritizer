"""
Outlook okunmamis mailleri onceliklendirme scripti.
Prioritize unread Outlook emails by relevance to you.

Ne yapar / What it does:
  - Inbox'taki OKUNMAMIS mailleri tarar / Scans UNREAD emails in your Inbox
  - Her mail icin en yuksek eslesen onceligi bulur ve Outlook KATEGORISI atar
    Assigns an Outlook CATEGORY based on the highest-priority match:
      1-@Isim        -> gercekten @mention edilmissin (en yuksek oncelik)
                         you were genuinely @mentioned (highest priority)
      2-Isim Geciyor -> konu/govdede duz metin olarak isim gecen
                         your name appears as plain text in subject/body
      3-Bana (To)    -> sen To alaninda / you are in the To field
      4-Bilgi (CC)   -> sen CC alaninda / you are in the CC field
  - Hicbir maili SILMEZ / TASIMAZ, sadece kategori ekler. Geri almak icin
    Outlook'ta ilgili maillerden kategoriyi kaldirman yeterli.
    Never deletes/moves anything - only tags categories (fully reversible).

Kurulum / Setup:
  pip install pywin32

Calistirmadan once / Before running:
  1) Outlook masaustu uygulamasini ac ve hesabina giris yap.
     Open desktop Outlook and make sure you're signed in.
  2) Asagidaki "KULLANICI AYARLARI" bolumunu kendi bilgilerinle doldur.
     Fill in the "USER SETTINGS" section below with your own name.
  3) Once DRY_RUN = True ile calistir, ozet ciktisini kontrol et.
     Run once with DRY_RUN = True and check the summary output.
  4) Sonuc dogruysa DRY_RUN = False yapip tekrar calistir.
     If it looks right, set DRY_RUN = False and run again.

Calistirma / Run:
  python outlook_prioritize.py
"""

import re
import win32com.client

# ====================== KULLANICI AYARLARI / USER SETTINGS ======================
# Ismin ve olasi varyasyonlarin (buyuk/kucuk harf onemli degil).
# Your name and its likely variants (case-insensitive).
# Ornek / Example: ["Ada", "Ada Lovelace"]
NAME_VARIANTS = ["Isminiz", "Isim Soyisim"]

# Outlook'tan otomatik email adresi bulunamazsa veya farkli bir adres
# kullanmak istersen buraya yaz, yoksa None birak.
# Leave as None to auto-detect from Outlook, or override with your address.
USER_EMAIL_OVERRIDE = None

# True: hicbir sey degistirmez, sadece ne yapacagini yazdirir (once bunu dene)
# False: kategorileri gercekten Outlook'a yazar
# True: dry-run, changes nothing, just prints a summary (try this first)
# False: actually writes categories to Outlook
DRY_RUN = True
# ===================================================================================

# (kategori adi, Outlook renk kodu) / (category name, Outlook color code)
# Renk kodlari / Color codes: 1=Kirmizi/Red 2=Turuncu/Orange 3=Seftali/Peach
# 4=Sari/Yellow 5=Yesil/Green 6=Turkuaz/Teal 7=Zeytin/Olive 8=Mavi/Blue
# 9=Mor/Purple 10=Bordo/Maroon 11=Celik/Steel 12=Koyu Celik/Dark Steel 13=Gri/Gray
CATEGORIES = [
    ("1-@Isim", 1),
    ("2-Isim Geciyor", 2),
    ("3-Bana (To)", 4),
    ("4-Bilgi (CC)", 8),
]

TR_MAP = str.maketrans("İIŞĞÜÖÇ", "iışğüöç")


def tr_lower(s):
    return s.translate(TR_MAP).lower()


def ensure_categories(session):
    existing = {c.Name for c in session.Categories}
    for name, color in CATEGORIES:
        if name not in existing:
            session.Categories.Add(name, color)


def get_user_email(app):
    if USER_EMAIL_OVERRIDE:
        return USER_EMAIL_OVERRIDE.lower()
    session = app.Session
    try:
        return session.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress.lower()
    except Exception:
        return session.CurrentUser.Address.lower()


def classify(mail, user_email):
    subject = mail.Subject or ""
    try:
        body = mail.Body or ""
    except Exception:
        body = ""
    text = tr_lower(subject + "\n" + body[:2000])

    # Gercek Outlook @mention tespiti: Outlook, kisi secilip mention eklendiginde
    # "@" karakterini kaldirir ve govdeye seninle e-postani bir mailto: linki
    # olarak gomer. Duz metin "@isim" arama sadece manuel yazanlari yakalar,
    # bu yuzden asil tespiti HTMLBody icindeki mailto linkinden yapiyoruz.
    #
    # Real Outlook @mention detection: once you pick a contact for a mention,
    # Outlook strips the "@" character and embeds your address as a mailto:
    # link in the body instead. Plain-text "@name" search only catches manual
    # typers, so we detect real mentions via the mailto: link in HTMLBody.
    try:
        html = mail.HTMLBody or ""
    except Exception:
        html = ""
    if user_email and f"mailto:{user_email}".lower() in html.lower():
        return 0

    for variant in NAME_VARIANTS:
        v = tr_lower(variant)
        if v and ("@" + v) in text:
            return 0

    for variant in NAME_VARIANTS:
        v = tr_lower(variant)
        if v and re.search(r"\b" + re.escape(v) + r"\b", text):
            return 1

    is_to, is_cc = False, False
    try:
        for r in mail.Recipients:
            addr = None
            try:
                addr = r.AddressEntry.GetExchangeUser().PrimarySmtpAddress
            except Exception:
                addr = getattr(r, "Address", None)
            if addr and addr.lower() == user_email:
                if r.Type == 1:
                    is_to = True
                elif r.Type == 2:
                    is_cc = True
    except Exception:
        pass

    if is_to:
        return 2
    if is_cc:
        return 3
    return None


def main():
    app = win32com.client.Dispatch("Outlook.Application")
    session = app.Session
    ensure_categories(session)
    user_email = get_user_email(app)

    inbox = session.GetDefaultFolder(6)  # olFolderInbox
    items = inbox.Items.Restrict("[Unread] = true")

    total = items.Count
    print(f"Toplam okunmamis mail: {total}")
    print(f"Eslesme icin kullanilan e-posta: {user_email}")
    print(f"DRY_RUN = {DRY_RUN}\n")

    counts = {0: 0, 1: 0, 2: 0, 3: 0, None: 0}
    processed = 0

    for mail in list(items):
        try:
            level = classify(mail, user_email)
        except Exception as e:
            print(f"Hata (atlandi): {e}")
            continue

        counts[level] = counts.get(level, 0) + 1

        if level is not None:
            cat_name = CATEGORIES[level][0]
            existing_cats = [c.strip() for c in (mail.Categories or "").split(";") if c.strip()]
            if cat_name not in existing_cats:
                new_cats = existing_cats + [cat_name]
                if not DRY_RUN:
                    mail.Categories = ";".join(new_cats)
                    mail.Save()

        processed += 1
        if processed % 50 == 0:
            print(f"{processed}/{total} islendi...")

    print("\n--- Ozet ---")
    print(f"@Isim ile bahsedilen : {counts[0]}")
    print(f"Isim geciyor         : {counts[1]}")
    print(f"To alaninda          : {counts[2]}")
    print(f"CC alaninda          : {counts[3]}")
    print(f"Eslesmeyen           : {counts[None]}")

    if DRY_RUN:
        print("\n(DRY_RUN=True oldugu icin hicbir kategori kaydedilmedi.)")
        print("Ozet dogru gorunuyorsa DRY_RUN = False yapip tekrar calistir.")


if __name__ == "__main__":
    main()
