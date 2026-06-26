"""
Outlook Firma Rehberi v4.0 — Poliport / S.SEYMEN
Tüm Outlook klasörleri + alt klasörler, tekrarsız mail, imza analizi
Gereksinim: pip install pywin32 openpyxl
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading, json, os, datetime, re
from pathlib import Path
from html.parser import HTMLParser

# ─── RENKLER ────────────────────────────────────────────────
BG         = "#F0F4F8"
PANEL      = "#FFFFFF"
HDR_BG     = "#1E3A5F"
HDR_FG     = "#FFFFFF"
ACC        = "#2563EB"
ACC_HOV    = "#1D4ED8"
SUCCESS    = "#059669"
WARN       = "#D97706"
BORDER     = "#CBD5E1"
TEXT       = "#1E293B"
TEXT2      = "#475569"
TEXT3      = "#94A3B8"
ROW_ALT    = "#F8FAFC"
ROW_HOV    = "#EFF6FF"
FIRMA_BG   = "#EFF6FF"
FIRMA_FG   = "#1E40AF"
KIS_BG     = "#FFF7ED"
KIS_FG     = "#C2410C"
TAG_G_BG   = "#DBEAFE"; TAG_G_FG = "#1D4ED8"
TAG_S_BG   = "#D1FAE5"; TAG_S_FG = "#065F46"

F_TITLE = ("Segoe UI", 13, "bold")
F_HDR   = ("Segoe UI", 10, "bold")
F_BODY  = ("Segoe UI",  9)
F_SMALL = ("Segoe UI",  8)
F_MONO  = ("Consolas",  9)

DATA_FILE = Path(__file__).parent / "rehber_data.json"

# Kişisel mail domain'leri
KISISEL_DOMAINS = {
    "gmail.com","hotmail.com","hotmail.com.tr","yahoo.com","yahoo.com.tr",
    "outlook.com","live.com","live.com.tr","icloud.com","msn.com",
    "googlemail.com","yandex.com","yandex.ru","protonmail.com","me.com",
    "yandex.com.tr","windowslive.com","aim.com","aol.com",
}

# ─── VERİ KATMANI ───────────────────────────────────────────
class HTMLText(HTMLParser):
    def __init__(self):
        super().__init__()
        self.parts = []; self._skip = False
    def handle_starttag(self, tag, _):
        if tag in ("script","style"): self._skip = True
        if tag in ("br","p","div","tr","li","td","th"): self.parts.append("\n")
    def handle_endtag(self, tag):
        if tag in ("script","style"): self._skip = False
    def handle_data(self, d):
        if not self._skip: self.parts.append(d)
    def get_text(self): return "".join(self.parts)

def html2txt(h):
    try:
        p = HTMLText(); p.feed(h or ""); return p.get_text()
    except:
        return h or ""

RE_EMAIL = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}", re.I)
RE_GSM   = re.compile(
    r"(?<!\d)(\+?90[\s.\-]?)?0?5[0-9]{2}[\s.\-]?[0-9]{3}[\s.\-]?[0-9]{2}[\s.\-]?[0-9]{2}(?!\d)")
RE_SABIT = re.compile(
    r"(?<!\d)(?:0\s*(?:212|216|232|242|262|312|322|332|352|362|392"
    r"|412|422|432|442|452|462|472|482)|0\s*[2-4][0-9]{2})"
    r"[\s.\-]?[0-9]{3}[\s.\-]?[0-9]{2}[\s.\-]?[0-9]{2}(?!\d)")

TITLE_KW = [
    "müdür","uzman","şef","direktör","koordinatör","asistan","yetkili","sorumlu",
    "operatör","operator","lojistik","satın alma","muhasebe","finans",
    "insan kaynakları","pazarlama","satış","genel müdür","yönetici","mühendis",
    "tekniker","danışman","analist","planlama","depo","sevkiyat","ihracat",
    "ithalat","operasyon","ulaştırma","taşıma","temsilci","purchasing",
    "manager","director","chief","officer","specialist","coordinator","assistant",
    "analyst","engineer","consultant","supervisor","executive","head of",
    "logistics","procurement","operations","sales","marketing",
    "finance","accounting","transport","shipping","export","import","supply chain",
]

SIG_NOISE = [
    "to:","from:","cc:","bcc:","kime:","kimden:","bilgi:",
    "subject:","konu:","date:","tarih:",
    "<",">"
    "http","www.","fax","faks",
]

def _norm(n): return re.sub(r"[^\d+]","",n)

def _is_sig_noise(line):
    l = line.strip().lower()
    return any(kw in l for kw in SIG_NOISE)

def _title(blk):
    for line in blk.splitlines():
        l = line.strip()
        if not l or _is_sig_noise(l): continue
        ll = l.lower()
        for kw in TITLE_KW:
            if kw in ll: return l
    return ""

def _name(blk, disp=""):
    # Önce sender display name
    if disp and " " in disp and "@" not in disp:
        c = re.sub(r"['\"]","",disp).strip()
        c = re.sub(r"\(.*?\)","",c).strip()
        words = c.split()
        if 2 <= len(words) <= 4 and not any(kw in c.lower() for kw in TITLE_KW):
            return c
    # İmza bloğunda ara
    for line in blk.splitlines():
        l = line.strip()
        if not l or "@" in l: continue
        if _is_sig_noise(l): continue
        if re.search(r"[\(\)\+\\/<>:;]", l): continue
        if re.search(r"\d", l): continue
        w = l.split()
        if 2 <= len(w) <= 4:
            if all(x[0].isupper() for x in w if x):
                if not any(kw in l.lower() for kw in TITLE_KW):
                    return l
    return ""

def _firma_from_sig(blk):
    """İmzada geçen firma adını bulmaya çalış — BÜYÜK HARF veya Ltd/A.Ş. içeren satır."""
    patterns = [
        re.compile(r"\b(a\.?ş\.?|ltd\.?|şti\.?|inc\.?|co\.?|llc\.?|gmbh)\b", re.I),
        re.compile(r"^[A-ZÇĞİÖŞÜ][A-ZÇĞİÖŞÜ\s\.\-&]{3,}$"),
    ]
    for line in blk.splitlines():
        l = line.strip()
        if not l or len(l) < 3 or len(l) > 80: continue
        if _is_sig_noise(l): continue
        if "@" in l or re.search(r"[\d]{4,}", l): continue
        for pat in patterns:
            if pat.search(l):
                return l.strip()
    return ""

def _sig_block(body):
    """İmza bloğunu bul — son 20 satır fallback."""
    markers = [
        r"^\s*-{2,}\s*$",
        r"(?i)^\s*(iyi çalışmalar|saygılarımla|regards|best regards|thanks|"
        r"teşekkürler|kind regards|br,?|cheers|sincerely|iyi günler|good day)\s*[,.]?\s*$",
        r"(?i)^\s*(iyi çalışmalar|best regards|thank you|saygılar|"
        r"iyi günler|good day)\b",
    ]
    lines = body.splitlines()
    for i, line in enumerate(lines):
        for pat in markers:
            if re.match(pat, line):
                return "\n".join(lines[i:])
    return "\n".join(lines[-20:]) if len(lines) > 20 else body

def parse_mail(msg):
    """Bir Outlook mail öğesini ayrıştır, kişi bilgilerini döndür."""
    try:
        em = (msg.SenderEmailAddress or "").strip()
        nm = (msg.SenderName or "").strip()

        # Exchange internal adresleri çöz
        if "SMTP:" in em:
            em = em.split("SMTP:")[-1].strip()

        em = em.lower()
        if not em or em.startswith("/o=") or not RE_EMAIL.match(em):
            return None

        dom = em.split("@")[-1].lower()
        kisisel = dom in KISISEL_DOMAINS

        try:
            body = html2txt(msg.HTMLBody) if msg.BodyFormat == 2 else (msg.Body or "")
        except:
            body = ""

        sig = _sig_block(body)

        # Telefon numaraları
        gsm = set()
        sabit = set()
        for m in RE_GSM.findall(sig):
            n = _norm(m[0] if isinstance(m, tuple) else m)
            digits = re.sub(r"\D","",n)
            if 10 <= len(digits) <= 13:
                gsm.add(digits)
        for m in RE_SABIT.findall(sig):
            n = _norm(m[0] if isinstance(m, tuple) else m)
            digits = re.sub(r"\D","",n)
            if digits:
                sabit.add(digits)

        name  = _name(sig, nm)
        title = _title(sig)

        # Firma: imzadan çek, yoksa @sonrası
        firma_sig = _firma_from_sig(sig)
        firma_display = firma_sig if firma_sig else dom.split(".")[0].upper()

        return {
            "email"       : em,
            "domain"      : dom,
            "firma"       : dom.split(".")[0].upper(),   # gruplama için
            "firma_display": firma_display,              # gösterim için
            "name"        : name,
            "title"       : title,
            "gsm"         : list(gsm),
            "sabit"       : list(sabit),
            "kisisel"     : kisisel,
        }
    except:
        return None

def walk(folder, results, last_scan, stats, log_cb=None):
    """Klasörü ve tüm alt klasörleri recursive tara."""
    try:
        items = folder.Items
        try:
            items.Sort("[ReceivedTime]", True)
        except:
            pass

        for item in items:
            try:
                if item.Class != 43: continue  # Sadece MailItem
                # Artımlı tarama: son tarihten eski mailları atla
                if last_scan:
                    try:
                        rcv = str(item.ReceivedTime)[:19].replace("T"," ")
                        ls  = last_scan.strftime("%Y-%m-%d %H:%M:%S")
                        if rcv <= ls:
                            break
                    except:
                        pass

                d = parse_mail(item)
                if not d or not d["email"]:
                    continue

                stats["tarandi"] += 1
                em = d["email"]

                if em not in results:
                    results[em] = d
                    stats["yeni"] += 1
                else:
                    # Mevcut kaydı zenginleştir
                    ex = results[em]
                    if not ex.get("title") and d.get("title"):
                        ex["title"] = d["title"]
                    if not ex.get("name") and d.get("name"):
                        ex["name"] = d["name"]
                    if not ex.get("firma_display") and d.get("firma_display"):
                        ex["firma_display"] = d["firma_display"]
                    for g in d.get("gsm", []):
                        if g not in ex.get("gsm", []):
                            ex.setdefault("gsm",[]).append(g)
                    for s in d.get("sabit", []):
                        if s not in ex.get("sabit", []):
                            ex.setdefault("sabit",[]).append(s)

            except:
                continue

        # Alt klasörleri tara
        for sub in folder.Folders:
            try:
                if log_cb:
                    log_cb(f"      📂 {sub.Name}")
                walk(sub, results, last_scan, stats, log_cb)
            except:
                continue
    except:
        pass

def load_data():
    if not DATA_FILE.exists():
        return {}, None
    try:
        with open(DATA_FILE, encoding="utf-8") as f:
            d = json.load(f)
        last = d.get("last_scan")
        if last:
            last = datetime.datetime.fromisoformat(last)
        return d.get("contacts", {}), last
    except:
        return {}, None

def save_data(contacts):
    payload = {
        "last_scan": datetime.datetime.now().isoformat(),
        "contacts": contacts
    }
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

def do_scan(log_cb, done_cb):
    try:
        import win32com.client
        contacts, last_scan = load_data()
        results = dict(contacts)
        stats   = {"tarandi": 0, "yeni": 0}

        outlook = win32com.client.Dispatch("Outlook.Application")
        ns      = outlook.GetNamespace("MAPI")

        # Tüm hesapları tara
        for i in range(1, ns.Folders.Count + 1):
            try:
                acc = ns.Folders.Item(i)
                log_cb(f"📧  Hesap: {acc.Name}")
                # Her hesabın tüm klasörlerini tara (Gelen, Gönderilen, Silinmiş + alt)
                for j in range(1, acc.Folders.Count + 1):
                    try:
                        sub = acc.Folders.Item(j)
                        log_cb(f"  📁 {sub.Name}")
                        walk(sub, results, last_scan, stats, log_cb)
                    except:
                        continue
            except:
                continue

        save_data(results)
        done_cb(results, stats)
    except Exception as e:
        done_cb(None, None, str(e))

def group_by_firma(contacts):
    """Kurumsal / kişisel ayır, domain'e göre grupla."""
    firms   = {}
    kisisel = {}
    for em, d in contacts.items():
        dom   = d.get("domain","")
        firma = dom.split(".")[0].upper() if dom else "DİĞER"
        if d.get("kisisel") or dom in KISISEL_DOMAINS:
            kisisel.setdefault(firma, []).append(d)
        else:
            firms.setdefault(firma, []).append(d)
    for f in firms:   firms[f].sort(key=lambda x: x.get("name","").lower())
    for f in kisisel: kisisel[f].sort(key=lambda x: x.get("name","").lower())
    return dict(sorted(firms.items())), dict(sorted(kisisel.items()))

def build_excel(contacts, path):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    firms, kisisel = group_by_firma(contacts)
    now = datetime.datetime.now()
    ts  = now.strftime("%d.%m.%Y %H:%M")
    wb  = Workbook()

    thin = Side(style="thin", color="CBD5E1")
    def brd(): return Border(left=thin, right=thin, top=thin, bottom=thin)

    def H(cell, val, bg="1E3A5F", fg="FFFFFF", sz=10, bold=True):
        cell.value = val
        cell.font  = Font(name="Calibri", size=sz, bold=bold, color=fg)
        cell.fill  = PatternFill("solid", fgColor=bg)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = brd()

    def D(cell, val, bg="FFFFFF", bold=False, left=False):
        cell.value = val
        cell.font  = Font(name="Calibri", size=9, bold=bold)
        cell.fill  = PatternFill("solid", fgColor=bg)
        cell.alignment = Alignment(
            horizontal="left" if left else "left",
            vertical="center", wrap_text=True)
        cell.border = brd()

    def write_sheet(title, data_firms, is_first=False, accent="1E3A5F"):
        if is_first:
            ws = wb.active; ws.title = title[:31]
        else:
            ws = wb.create_sheet(title[:31])
        ws.sheet_view.showGridLines = False

        ws.merge_cells("A1:H1")
        H(ws["A1"], f"OUTLOOK TARAMA — {title.upper()}", bg=accent, sz=12)
        ws.row_dimensions[1].height = 28
        ws.merge_cells("A2:H2")
        H(ws["A2"], f"Oluşturulma: {ts}  |  Toplam: {sum(len(v) for v in data_firms.values())} kişi",
          bg="2D4A6E", sz=9, bold=False)
        ws.row_dimensions[2].height = 14

        cols = ["FİRMA (İMZA)","DOMAIN","AD SOYAD","GÖREV","E-POSTA","CEP TEL","SABİT TEL","NOTLAR"]
        wids = [22, 20, 22, 28, 32, 16, 16, 20]
        for ci,(h,w) in enumerate(zip(cols, wids), start=1):
            H(ws.cell(row=3, column=ci), h, bg=accent)
            ws.column_dimensions[get_column_letter(ci)].width = w
        ws.row_dimensions[3].height = 20

        rn = 4
        for fi,(firma,people) in enumerate(data_firms.items()):
            dom = people[0].get("domain","") if people else ""
            for pi, p in enumerate(people):
                bg = "EFF6FF" if fi%2==0 else "FFFFFF"
                firma_disp = p.get("firma_display","") or firma
                gsm_str    = "  /  ".join(p.get("gsm",[]))
                sabit_str  = "  /  ".join(p.get("sabit",[]))
                vals = [
                    firma_disp if pi==0 else "",
                    dom        if pi==0 else "",
                    p.get("name",""),
                    p.get("title",""),
                    p.get("email",""),
                    gsm_str,
                    sabit_str,
                    "",
                ]
                for ci,val in enumerate(vals, start=1):
                    D(ws.cell(row=rn, column=ci), val, bg=bg,
                      bold=(ci<=2 and pi==0))
                ws.row_dimensions[rn].height = 16
                rn += 1
            # Ayırıcı satır
            for ci in range(1, 9):
                ws.cell(row=rn, column=ci).fill = PatternFill("solid", fgColor="E2E8F0")
            ws.row_dimensions[rn].height = 3
            rn += 1

        ws.freeze_panes = "A4"
        ws.auto_filter.ref = f"A3:{get_column_letter(8)}{rn-1}"
        return ws

    write_sheet("Kurumsal Firmalar", firms,   is_first=True,  accent="1E3A5F")
    write_sheet("Kisisel Mailler",   kisisel, is_first=False, accent="7C3AED")

    # ── Özet sayfası ──
    ws2 = wb.create_sheet("Ozet", 0)
    ws2.sheet_view.showGridLines = False
    ws2.merge_cells("A1:F1")
    H(ws2["A1"], "OUTLOOK FİRMA REHBERİ — ÖZET", sz=14)
    ws2.row_dimensions[1].height = 34
    ws2.merge_cells("A2:F2")
    H(ws2["A2"], f"Oluşturulma: {ts}", bg="2D4A6E", sz=9, bold=False)
    ws2.row_dimensions[2].height = 16
    ws2.row_dimensions[3].height = 8

    total_kis = sum(len(v) for v in kisisel.values())
    for ci,(lbl,val) in enumerate([
        ("KURUMSAL FİRMA",   len(firms)),
        ("KURUMSAL KİŞİ",    sum(len(v) for v in firms.values())),
        ("KİŞİSEL MAİL",     total_kis),
        ("TOPLAM BENZERSİZ", len(contacts)),
    ], start=1):
        cl = get_column_letter(ci)
        ws2.merge_cells(f"{cl}4:{cl}4"); ws2.merge_cells(f"{cl}5:{cl}5")
        lc = ws2[f"{cl}4"]; vc = ws2[f"{cl}5"]
        lc.value = lbl
        lc.font  = Font(name="Calibri", size=8, bold=True, color="64748B")
        lc.alignment = Alignment(horizontal="center")
        vc.value = val
        vc.font  = Font(name="Calibri", size=20, bold=True, color="2563EB")
        vc.alignment = Alignment(horizontal="center")
    ws2.row_dimensions[4].height = 16
    ws2.row_dimensions[5].height = 30
    ws2.row_dimensions[6].height = 10

    for ci,(h,w) in enumerate(zip(
        ["FİRMA","DOMAIN","KİŞİ SAYISI","CEP TEL (1. Kişi)","ÖRNEK MAİLLER"],
        [22,25,12,18,50]
    ), start=1):
        H(ws2.cell(row=7, column=ci), h)
        ws2.column_dimensions[get_column_letter(ci)].width = w
    ws2.row_dimensions[7].height = 20

    for idx,(firma,people) in enumerate(firms.items(), start=1):
        r   = 7 + idx
        bg  = "EFF6FF" if idx%2==0 else "FFFFFF"
        dom = people[0].get("domain","") if people else ""
        gsm = people[0].get("gsm",[""])[0] if people else ""
        preview = "; ".join(p["email"] for p in people[:4])
        for ci,val in enumerate([firma, dom, len(people), gsm, preview], start=1):
            D(ws2.cell(row=r, column=ci), val, bg=bg)
        ws2.row_dimensions[r].height = 16

    ws2.freeze_panes = "A8"
    wb.save(path)


# ─── ÖZEL BUTON ─────────────────────────────────────────────
class FlatBtn(tk.Frame):
    def __init__(self, parent, text, command, bg=ACC, fg="white",
                 width=140, font=F_HDR, **kw):
        super().__init__(parent, bg=parent["bg"] if "bg" not in kw else kw.pop("bg"))
        self._cmd = command; self._bg = bg; self._text = text
        self._lbl = tk.Label(
            self, text=text, bg=bg, fg=fg, font=font,
            width=width//7, padx=10, pady=7,
            cursor="hand2", relief="flat"
        )
        self._lbl.pack()
        self._lbl.bind("<Enter>",    lambda e: self._lbl.config(bg=self._dk(bg)))
        self._lbl.bind("<Leave>",    lambda e: self._lbl.config(bg=self._bg))
        self._lbl.bind("<Button-1>", lambda e: self._cmd() if self._cmd else None)

    def _dk(self, h):
        r=int(h[1:3],16); g=int(h[3:5],16); b=int(h[5:7],16); f=0.82
        return f"#{int(r*f):02x}{int(g*f):02x}{int(b*f):02x}"

    def set_text(self, t):
        self._lbl.config(text=t)

    def set_state(self, on):
        if on:
            self._bg = ACC
            self._lbl.config(bg=ACC, cursor="hand2")
            self._lbl.bind("<Button-1>", lambda e: self._cmd() if self._cmd else None)
        else:
            self._bg = "#94A3B8"
            self._lbl.config(bg="#94A3B8", cursor="")
            self._lbl.unbind("<Button-1>")


# ─── ANA UYGULAMA ───────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Outlook Firma Rehberi v4.0 — Poliport")
        self.geometry("1280x760")
        self.minsize(960, 580)
        self.configure(bg=BG)
        try:
            self.state("zoomed")
        except:
            pass

        self._contacts = {}
        self._firms    = {}
        self._kisisel  = {}
        self._scanning = False
        self._mode     = "kurumsal"

        self._build_ui()
        self._load_existing()

    # ── UI ──────────────────────────────────────────────────
    def _build_ui(self):
        self._build_header()
        self._build_toolbar()
        self._build_body()
        self._build_status()

    def _build_header(self):
        hdr = tk.Frame(self, bg=HDR_BG, height=76)
        hdr.pack(fill="x"); hdr.pack_propagate(False)

        L = tk.Frame(hdr, bg=HDR_BG); L.pack(side="left", padx=20, pady=12)
        tk.Label(L, text="📋", font=("Segoe UI",20), bg=HDR_BG, fg=HDR_FG
                 ).pack(side="left", padx=(0,10))
        T = tk.Frame(L, bg=HDR_BG); T.pack(side="left")
        tk.Label(T, text="Outlook Firma Rehberi",
                 font=("Segoe UI",14,"bold"), bg=HDR_BG, fg=HDR_FG).pack(anchor="w")
        tk.Label(T, text="Tüm hesaplar · Gelen + Gönderilen + Silinmiş + Alt Klasörler · Tekrarsız",
                 font=("Segoe UI",8), bg=HDR_BG, fg="#93C5FD").pack(anchor="w")

        R = tk.Frame(hdr, bg=HDR_BG); R.pack(side="right", padx=20)
        self._stats = {}
        for lbl, key, col in [
            ("Firma",   "firma",   "#1E40AF"),
            ("Kişi",    "kisi",    "#1E40AF"),
            ("Kişisel", "kisisel", "#7C3AED"),
            ("Toplam",  "toplam",  "#059669"),
        ]:
            sf = tk.Frame(R, bg=col, padx=14, pady=6); sf.pack(side="left", padx=3)
            n  = tk.Label(sf, text="0", font=("Segoe UI",16,"bold"), bg=col, fg="#DBEAFE")
            n.pack()
            tk.Label(sf, text=lbl, font=("Segoe UI",8), bg=col, fg="#BFDBFE").pack()
            self._stats[key] = n

        self._last_lbl = tk.Label(hdr, text="Son tarama: henüz yapılmadı",
                                  font=("Segoe UI",8), bg=HDR_BG, fg="#93C5FD")
        self._last_lbl.place(relx=0.5, rely=0.92, anchor="s")

    def _build_toolbar(self):
        bar = tk.Frame(self, bg=PANEL, pady=9,
                       highlightbackground=BORDER, highlightthickness=1)
        bar.pack(fill="x")
        inner = tk.Frame(bar, bg=PANEL); inner.pack(fill="x", padx=16)

        # Arama kutusu
        sf = tk.Frame(inner, bg=PANEL, highlightbackground=BORDER, highlightthickness=1)
        sf.pack(side="left", fill="x", expand=True)
        tk.Label(sf, text="🔍", bg=PANEL, fg=TEXT3, font=("Segoe UI",10)
                 ).pack(side="left", padx=(8,0))
        self._sv = tk.StringVar()
        self._sv.trace_add("write", lambda *a: self._filter(self._sv.get()))
        tk.Entry(sf, textvariable=self._sv, bg=PANEL, fg=TEXT, font=F_BODY,
                 bd=0, relief="flat", insertbackground=ACC
                 ).pack(side="left", fill="x", expand=True, padx=6, pady=7)

        # Butonlar
        BF = tk.Frame(inner, bg=PANEL); BF.pack(side="right", padx=(10,0))

        self._btn_kur = FlatBtn(BF, "🏢 Kurumsal",
                                lambda: self._set_mode("kurumsal"), bg=ACC, width=115)
        self._btn_kur.pack(side="left", padx=2)

        self._btn_kis = FlatBtn(BF, "✉ Kişisel",
                                lambda: self._set_mode("kisisel"), bg="#7C3AED", width=105)
        self._btn_kis.pack(side="left", padx=2)

        tk.Frame(BF, bg=BORDER, width=1, height=28).pack(side="left", padx=8)

        FlatBtn(BF, "📥 Excel İçe Aktar", self._import_excel,
                bg="#0891B2", width=148).pack(side="left", padx=2)

        self._scan_btn = FlatBtn(BF, "🔄 Tarama Yap", self._start_scan,
                                 bg=ACC, width=130)
        self._scan_btn.pack(side="left", padx=2)

        FlatBtn(BF, "📤 Excel İndir", self._export_excel,
                bg=SUCCESS, width=125).pack(side="left", padx=2)

        FlatBtn(BF, "🗑 Sıfırla", self._reset_data,
                bg=WARN, width=95).pack(side="left", padx=2)

    def _build_body(self):
        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=12, pady=(8,0))

        # Sol — firma listesi
        LEFT = tk.Frame(body, bg=PANEL, width=240,
                        highlightbackground=BORDER, highlightthickness=1)
        LEFT.pack(side="left", fill="y", padx=(0,8))
        LEFT.pack_propagate(False)

        self._firma_title = tk.Label(LEFT, text="KURUMSAL FİRMALAR",
                                     font=("Segoe UI",8,"bold"), bg=PANEL, fg=TEXT3)
        self._firma_title.pack(anchor="w", padx=12, pady=(10,4))

        self._firma_var = tk.StringVar(value=[])
        self._firma_lb  = tk.Listbox(
            LEFT, listvariable=self._firma_var, font=F_BODY,
            bg=PANEL, fg=TEXT, selectbackground=ACC, selectforeground="white",
            activestyle="none", bd=0, highlightthickness=0, relief="flat"
        )
        sb = ttk.Scrollbar(LEFT, orient="vertical", command=self._firma_lb.yview)
        self._firma_lb.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self._firma_lb.pack(fill="both", expand=True, padx=4, pady=(0,8))
        self._firma_lb.bind("<<ListboxSelect>>", self._on_firma_sel)

        # Sağ — kişi tablosu
        RIGHT = tk.Frame(body, bg=PANEL,
                         highlightbackground=BORDER, highlightthickness=1)
        RIGHT.pack(side="left", fill="both", expand=True)

        self._det_hdr = tk.Frame(RIGHT, bg=FIRMA_BG, pady=10)
        self._det_hdr.pack(fill="x")
        self._det_name = tk.Label(self._det_hdr, text="← Soldan bir firma seçin",
                                  font=F_TITLE, bg=FIRMA_BG, fg=FIRMA_FG)
        self._det_name.pack(side="left", padx=16)
        self._det_cnt = tk.Label(self._det_hdr, text="",
                                 font=F_SMALL, bg=FIRMA_BG, fg=TEXT2)
        self._det_cnt.pack(side="right", padx=16)

        # Kolon başlıkları
        CH = tk.Frame(RIGHT, bg="#E2E8F0", pady=5); CH.pack(fill="x")
        for txt, w in [("Ad Soyad",22),("Görev / Unvan",26),("E-posta",28),
                        ("Cep Tel",14),("Sabit Tel",13)]:
            tk.Label(CH, text=txt.upper(), font=("Segoe UI",8,"bold"),
                     bg="#E2E8F0", fg=TEXT3, width=w, anchor="w"
                     ).pack(side="left", padx=8)

        # Scrollable tablo
        CF = tk.Frame(RIGHT, bg=PANEL); CF.pack(fill="both", expand=True)
        self._cv  = tk.Canvas(CF, bg=PANEL, highlightthickness=0)
        vsb = ttk.Scrollbar(CF, orient="vertical", command=self._cv.yview)
        self._cv.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self._cv.pack(side="left", fill="both", expand=True)
        self._tbl  = tk.Frame(self._cv, bg=PANEL)
        self._twin = self._cv.create_window((0,0), window=self._tbl, anchor="nw")
        self._tbl.bind("<Configure>",
                       lambda e: self._cv.configure(scrollregion=self._cv.bbox("all")))
        self._cv.bind("<Configure>",
                      lambda e: self._cv.itemconfig(self._twin, width=e.width))
        self._cv.bind_all("<MouseWheel>",
                          lambda e: self._cv.yview_scroll(-1*(e.delta//120), "units"))

    def _build_status(self):
        sf = tk.Frame(self, bg="#E2E8F0", height=26)
        sf.pack(fill="x", side="bottom")
        sf.pack_propagate(False)
        self._st = tk.Label(sf, text="Hazır", font=F_SMALL, bg="#E2E8F0", fg=TEXT2)
        self._st.pack(side="left", padx=12)
        self._prog = tk.Label(sf, text="", font=F_SMALL, bg="#E2E8F0", fg=TEXT3)
        self._prog.pack(side="right", padx=12)

    # ── Mod ─────────────────────────────────────────────────
    def _set_mode(self, mode):
        self._mode = mode
        firms  = self._firms if mode=="kurumsal" else self._kisisel
        hdr_bg = FIRMA_BG  if mode=="kurumsal" else KIS_BG
        hdr_fg = FIRMA_FG  if mode=="kurumsal" else KIS_FG
        lbl    = "KURUMSAL FİRMALAR" if mode=="kurumsal" else "KİŞİSEL MAİLLER"
        self._firma_title.config(text=lbl)
        self._det_hdr.config(bg=hdr_bg)
        self._det_name.config(bg=hdr_bg, fg=hdr_fg)
        self._det_cnt.config(bg=hdr_bg)
        self._firma_var.set(list(firms.keys()))
        if firms:
            self._firma_lb.selection_set(0)
            self._show_firma(list(firms.keys())[0])
        else:
            for w in self._tbl.winfo_children(): w.destroy()
            self._det_name.config(text="Bu bölümde veri yok")
            self._det_cnt.config(text="")

    # ── Veri yükleme ────────────────────────────────────────
    def _load_existing(self):
        self._contacts, last = load_data()
        self._firms, self._kisisel = group_by_firma(self._contacts)
        self._refresh_list()
        self._update_stats()
        if last:
            self._last_lbl.config(text=f"Son tarama: {last.strftime('%d.%m.%Y  %H:%M:%S')}")
        self._status(
            "Veriler yüklendi — Tarama Yap ile güncelleyin." if self._contacts
            else "Veri yok — 'Tarama Yap' butonuna tıklayın."
        )

    def _refresh_list(self):
        firms = self._firms if self._mode=="kurumsal" else self._kisisel
        self._firma_var.set(list(firms.keys()))
        if firms:
            self._firma_lb.selection_set(0)
            self._show_firma(list(firms.keys())[0])

    def _update_stats(self):
        self._stats["firma"].config(text=str(len(self._firms)))
        self._stats["kisi"].config(
            text=str(sum(len(v) for v in self._firms.values())))
        self._stats["kisisel"].config(
            text=str(sum(len(v) for v in self._kisisel.values())))
        self._stats["toplam"].config(text=str(len(self._contacts)))

    # ── Firma seçimi & tablo ────────────────────────────────
    def _on_firma_sel(self, event):
        sel = self._firma_lb.curselection()
        if not sel: return
        self._show_firma(self._firma_lb.get(sel[0]))

    def _show_firma(self, firma):
        for w in self._tbl.winfo_children(): w.destroy()
        firms  = self._firms if self._mode=="kurumsal" else self._kisisel
        people = firms.get(firma, [])
        dom    = people[0].get("domain","") if people else ""
        self._det_name.config(text=firma)
        self._det_cnt.config(text=f"{len(people)} kişi  ·  {dom}")

        for i, p in enumerate(people):
            bg  = PANEL if i%2==0 else ROW_ALT
            row = tk.Frame(self._tbl, bg=bg); row.pack(fill="x")
            row.bind("<Enter>", lambda e, r=row: r.config(bg=ROW_HOV))
            row.bind("<Leave>", lambda e, r=row, b=bg: r.config(bg=b))

            nm  = p.get("name","") or ""
            tl  = p.get("title","") or ""
            em  = p.get("email","") or ""
            gsm_list   = p.get("gsm",[])
            sabit_list = p.get("sabit",[])

            # Ad Soyad
            tk.Label(row, text=nm, bg=bg, fg=TEXT,
                     font=("Segoe UI",9,"bold") if nm else F_BODY,
                     width=22, anchor="w", wraplength=155).pack(side="left", padx=8, pady=6)
            # Görev
            tk.Label(row, text=tl, bg=bg, fg=TEXT2, font=F_BODY,
                     width=26, anchor="w", wraplength=175).pack(side="left", padx=8, pady=6)
            # E-posta
            tk.Label(row, text=em, bg=bg, fg=ACC, font=F_BODY,
                     width=28, anchor="w", wraplength=195).pack(side="left", padx=8, pady=6)

            # Telefon badge'leri
            ph = tk.Frame(row, bg=bg); ph.pack(side="left", padx=4, pady=4)
            for g in gsm_list:
                g_fmt = self._fmt_gsm(g)
                tk.Label(ph, text=f"📱 {g_fmt}", bg=TAG_G_BG, fg=TAG_G_FG,
                         font=F_SMALL, padx=5, pady=2).pack(side="left", padx=2)
            for s in sabit_list:
                s_fmt = self._fmt_sabit(s)
                tk.Label(ph, text=f"☎ {s_fmt}", bg=TAG_S_BG, fg=TAG_S_FG,
                         font=F_SMALL, padx=5, pady=2).pack(side="left", padx=2)

            tk.Frame(self._tbl, bg=BORDER, height=1).pack(fill="x")

        self._cv.yview_moveto(0)

    def _fmt_gsm(self, digits):
        """05XX XXX XX XX formatı"""
        d = re.sub(r"\D","",digits)
        if d.startswith("90") and len(d)==12: d = "0" + d[2:]
        if len(d)==10: return f"0{d[1:4]} {d[4:7]} {d[7:9]} {d[9:11]}"
        return digits

    def _fmt_sabit(self, digits):
        """0XXX XXX XX XX"""
        d = re.sub(r"\D","",digits)
        if len(d)==10: return f"0{d[1:4]} {d[4:7]} {d[7:9]} {d[9:11]}"
        return digits

    # ── Arama ───────────────────────────────────────────────
    def _filter(self, q=""):
        q = q.strip().lower()
        firms = self._firms if self._mode=="kurumsal" else self._kisisel
        items = [
            f for f, people in firms.items()
            if not q
            or q in f.lower()
            or any(
                q in (p.get("name","") + p.get("title","") +
                      p.get("email","") + p.get("firma_display","")).lower()
                for p in people
            )
        ]
        self._firma_var.set(items)
        if items:
            self._firma_lb.selection_set(0)
            self._show_firma(items[0])

    # ── Excel içe aktar ─────────────────────────────────────
    def _import_excel(self):
        path = filedialog.askopenfilename(
            filetypes=[("Excel Dosyası","*.xlsx *.xls")],
            title="Mail Listesi Excel Dosyasını Seç")
        if not path: return
        try:
            self._status("Excel içe aktarılıyor...")
            from openpyxl import load_workbook
            wb   = load_workbook(path, read_only=True)
            yeni = 0
            contacts = dict(self._contacts)
            for sheet in wb.sheetnames:
                ws = wb[sheet]
                for row in ws.iter_rows(values_only=True):
                    if not row or not row[0]: continue
                    email_val = None
                    for cell in row:
                        if cell and RE_EMAIL.match(str(cell).strip()):
                            email_val = str(cell).strip().lower(); break
                    if not email_val: continue
                    if email_val in contacts: continue
                    dom = email_val.split("@")[-1].lower()
                    contacts[email_val] = {
                        "email": email_val, "domain": dom,
                        "firma": dom.split(".")[0].upper(),
                        "firma_display": dom.split(".")[0].upper(),
                        "name":"","title":"","gsm":[],"sabit":[],
                        "kisisel": dom in KISISEL_DOMAINS,
                    }
                    yeni += 1
            save_data(contacts)
            self._contacts = contacts
            self._firms, self._kisisel = group_by_firma(contacts)
            self._refresh_list(); self._update_stats()
            self._status(f"✅  Excel aktarıldı — {yeni} yeni e-posta eklendi.")
            messagebox.showinfo("İçe Aktarma Tamam",
                f"{yeni} yeni e-posta eklendi.\nToplam: {len(contacts)}")
        except Exception as e:
            messagebox.showerror("Hata", f"Excel okunamadı:\n{e}")
            self._status("İçe aktarma başarısız.")

    # ── Tarama ──────────────────────────────────────────────
    def _start_scan(self):
        if self._scanning: return
        self._scanning = True
        self._scan_btn.set_state(False)
        self._scan_btn.set_text("⏳ Taranıyor...")
        self._status("Outlook bağlantısı kuruluyor, lütfen bekleyin...")

        def log(msg):
            self.after(0, lambda: self._prog.config(text=msg[-60:]))

        def done(contacts, stats, err=None):
            self.after(0, lambda: self._scan_done(contacts, stats, err))

        threading.Thread(target=do_scan, args=(log, done), daemon=True).start()

    def _scan_done(self, contacts, stats, err):
        self._scanning = False
        self._scan_btn.set_state(True)
        self._scan_btn.set_text("🔄 Tarama Yap")
        self._prog.config(text="")

        if err:
            messagebox.showerror("Tarama Hatası",
                f"Outlook'a bağlanılamadı:\n\n{err}\n\n"
                "Outlook'un açık ve oturum açık olduğundan emin olun.")
            self._status("Tarama başarısız."); return

        self._contacts = contacts
        self._firms, self._kisisel = group_by_firma(contacts)
        self._refresh_list(); self._update_stats()
        now = datetime.datetime.now()
        self._last_lbl.config(text=f"Son tarama: {now.strftime('%d.%m.%Y  %H:%M:%S')}")
        self._status(
            f"✅  {stats['tarandi']} mail okundu, {stats['yeni']} yeni kişi eklendi")
        messagebox.showinfo("Tarama Tamamlandı",
            f"Tarama başarıyla tamamlandı!\n\n"
            f"Okunan mail    : {stats['tarandi']}\n"
            f"Yeni kişi      : {stats['yeni']}\n"
            f"Toplam kişi    : {len(contacts)}\n"
            f"Kurumsal firma : {len(self._firms)}\n"
            f"Kişisel mail   : {sum(len(v) for v in self._kisisel.values())}")

    # ── Excel dışa aktar ────────────────────────────────────
    def _export_excel(self):
        if not self._contacts:
            messagebox.showwarning("Veri Yok",
                "Önce tarama yapın veya Excel içe aktarın."); return
        ts      = datetime.datetime.now().strftime("%d.%m.%Y_%H-%M")
        default = f"Outlook_Firma_Rehberi_{ts}.xlsx"
        path    = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Dosyası","*.xlsx")],
            initialfile=default,
            title="Excel Kaydet")
        if not path: return
        try:
            self._status("Excel oluşturuluyor...")
            build_excel(self._contacts, path)
            self._status(f"✅  Kaydedildi: {Path(path).name}")
            if messagebox.askyesno("Excel Hazır",
                f"Dosya kaydedildi:\n{path}\n\nŞimdi açmak ister misiniz?"):
                os.startfile(path)
        except Exception as e:
            messagebox.showerror("Hata", f"Excel oluşturulamadı:\n{e}")

    # ── Veri sıfırla ────────────────────────────────────────
    def _reset_data(self):
        if not messagebox.askyesno("Sıfırla",
            "Tüm tarama verileri silinecek.\nEmin misiniz?",
            icon="warning"): return
        if DATA_FILE.exists():
            DATA_FILE.unlink()
        self._contacts = {}; self._firms = {}; self._kisisel = {}
        self._refresh_list(); self._update_stats()
        self._last_lbl.config(text="Son tarama: henüz yapılmadı")
        self._status("Veriler sıfırlandı. Yeni tarama yapabilirsiniz.")

    def _status(self, msg):
        self._st.config(text=msg)


if __name__ == "__main__":
    app = App()
    app.mainloop()
