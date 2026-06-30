"""
Outlook Firma Rehberi v4.1 — Poliport / S.SEYMEN
Tüm Outlook klasörleri + alt klasörler, tekrarsız mail, imza analizi, düzenleme
Gereksinim: pip install pywin32 openpyxl
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading, json, os, datetime, re
from pathlib import Path
from html.parser import HTMLParser

# ─── RENKLER ────────────────────────────────────────────────
BG       = "#F0F4F8"
PANEL    = "#FFFFFF"
HDR_BG   = "#1E3A5F"
HDR_FG   = "#FFFFFF"
ACC      = "#2563EB"
SUCCESS  = "#059669"
WARN     = "#D97706"
BORDER   = "#CBD5E1"
TEXT     = "#1E293B"
TEXT2    = "#475569"
TEXT3    = "#94A3B8"
ROW_ALT  = "#F8FAFC"
ROW_HOV  = "#EFF6FF"
FIRMA_BG = "#EFF6FF"
FIRMA_FG = "#1E40AF"
KIS_BG   = "#FFF7ED"
KIS_FG   = "#C2410C"
TAG_G_BG = "#DBEAFE"; TAG_G_FG = "#1D4ED8"
TAG_S_BG = "#D1FAE5"; TAG_S_FG = "#065F46"
EDIT_BG  = "#FFF9C4"

F_TITLE = ("Segoe UI", 13, "bold")
F_HDR   = ("Segoe UI", 10, "bold")
F_BODY  = ("Segoe UI",  9)
F_SMALL = ("Segoe UI",  8)

DATA_FILE = Path(__file__).parent / "rehber_data.json"

KISISEL_DOMAINS = {
    "gmail.com","hotmail.com","hotmail.com.tr","yahoo.com","yahoo.com.tr",
    "outlook.com","live.com","live.com.tr","icloud.com","msn.com",
    "googlemail.com","yandex.com","yandex.ru","protonmail.com","me.com",
    "yandex.com.tr","windowslive.com","aim.com","aol.com",
}

# ─── HTML → TEXT ────────────────────────────────────────────
class HTMLText(HTMLParser):
    def __init__(self):
        super().__init__(); self.parts = []; self._skip = False
    def handle_starttag(self, tag, _):
        if tag in ("script","style"): self._skip = True
        if tag in ("br","p","div","tr","li","td","th"): self.parts.append("\n")
    def handle_endtag(self, tag):
        if tag in ("script","style"): self._skip = False
    def handle_data(self, d):
        if not self._skip: self.parts.append(d)
    def get_text(self): return "".join(self.parts)

def html2txt(h):
    try: p = HTMLText(); p.feed(h or ""); return p.get_text()
    except: return h or ""

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
SIG_NOISE = ["to:","from:","cc:","bcc:","kime:","kimden:","bilgi:",
             "subject:","konu:","date:","tarih:","http","www.","fax","faks","<",">"]

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
    if disp and " " in disp and "@" not in disp:
        c = re.sub(r"['\"]","",disp).strip()
        c = re.sub(r"\(.*?\)","",c).strip()
        words = c.split()
        if 2 <= len(words) <= 4 and not any(kw in c.lower() for kw in TITLE_KW):
            return c
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
            if pat.search(l): return l.strip()
    return ""

def _sig_block(body):
    markers = [
        r"^\s*-{2,}\s*$",
        r"(?i)^\s*(iyi çalışmalar|saygılarımla|regards|best regards|thanks|"
        r"teşekkürler|kind regards|br,?|cheers|sincerely|iyi günler|good day)\s*[,.]?\s*$",
    ]
    lines = body.splitlines()
    for i, line in enumerate(lines):
        for pat in markers:
            if re.match(pat, line):
                return "\n".join(lines[i:])
    return "\n".join(lines[-20:]) if len(lines) > 20 else body

def parse_mail(msg):
    try:
        em = (msg.SenderEmailAddress or "").strip()
        nm = (msg.SenderName or "").strip()
        if "SMTP:" in em: em = em.split("SMTP:")[-1].strip()
        em = em.lower()
        if not em or em.startswith("/o=") or not RE_EMAIL.match(em): return None
        dom = em.split("@")[-1].lower()
        kisisel = dom in KISISEL_DOMAINS
        try:
            body = html2txt(msg.HTMLBody) if msg.BodyFormat == 2 else (msg.Body or "")
        except: body = ""
        sig = _sig_block(body)
        gsm = set(); sabit = set()
        for m in RE_GSM.findall(sig):
            n = _norm(m[0] if isinstance(m, tuple) else m)
            digits = re.sub(r"\D","",n)
            if 10 <= len(digits) <= 13: gsm.add(digits)
        for m in RE_SABIT.findall(sig):
            n = _norm(m[0] if isinstance(m, tuple) else m)
            digits = re.sub(r"\D","",n)
            if digits: sabit.add(digits)
        name        = _name(sig, nm)
        firma_sig   = _firma_from_sig(sig)
        firma_disp  = firma_sig if firma_sig else dom.split(".")[0].upper()
        return {
            "email": em, "domain": dom,
            "firma": dom.split(".")[0].upper(),
            "firma_display": firma_disp,
            "name": name, "title": _title(sig),
            "gsm": list(gsm), "sabit": list(sabit),
            "kisisel": kisisel,
        }
    except: return None

def walk(folder, results, last_scan, stats, log_cb=None):
    try:
        items = folder.Items
        try: items.Sort("[ReceivedTime]", True)
        except: pass
        for item in items:
            try:
                if item.Class != 43: continue
                if last_scan:
                    try:
                        rcv = str(item.ReceivedTime)[:19].replace("T"," ")
                        ls  = last_scan.strftime("%Y-%m-%d %H:%M:%S")
                        if rcv <= ls: break
                    except: pass
                d = parse_mail(item)
                if not d or not d["email"]: continue
                stats["tarandi"] += 1
                em = d["email"]
                if em not in results:
                    results[em] = d; stats["yeni"] += 1
                else:
                    ex = results[em]
                    for fld in ["name","title","firma_display"]:
                        if not ex.get(fld) and d.get(fld):
                            ex[fld] = d[fld]
                    for g in d.get("gsm",[]):
                        if g not in ex.get("gsm",[]): ex.setdefault("gsm",[]).append(g)
                    for s in d.get("sabit",[]):
                        if s not in ex.get("sabit",[]): ex.setdefault("sabit",[]).append(s)
            except: continue
        for sub in folder.Folders:
            try:
                if log_cb: log_cb(f"  klasor: {sub.Name}")
                walk(sub, results, last_scan, stats, log_cb)
            except: continue
    except: pass

def load_data():
    if not DATA_FILE.exists(): return {}, None
    try:
        with open(DATA_FILE, encoding="utf-8") as f:
            d = json.load(f)
        last = d.get("last_scan")
        if last: last = datetime.datetime.fromisoformat(last)
        return d.get("contacts", {}), last
    except: return {}, None

def save_data(contacts):
    payload = {"last_scan": datetime.datetime.now().isoformat(), "contacts": contacts}
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

# ---- Klasor agaci (secim penceresi icin) ----
def get_folder_tree():
    """Outlook'taki hesap + klasor + alt klasor agacini doner."""
    import win32com.client
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns      = outlook.GetNamespace("MAPI")
    tree = []

    def collect(folder, depth, path_list):
        node = {"name": folder.Name, "depth": depth, "path": path_list, "children": []}
        try:
            for sub in folder.Folders:
                try:
                    node["children"].append(collect(sub, depth+1, path_list + [sub.Name]))
                except: continue
        except: pass
        return node

    for i in range(1, ns.Folders.Count + 1):
        try:
            acc = ns.Folders.Item(i)
            acc_node = {"name": acc.Name, "depth": 0, "path": [acc.Name], "children": []}
            for j in range(1, acc.Folders.Count + 1):
                try:
                    sub = acc.Folders.Item(j)
                    acc_node["children"].append(collect(sub, 1, [acc.Name, sub.Name]))
                except: continue
            tree.append(acc_node)
        except: continue
    return tree

def _resolve_folder(ns, path):
    """path = [HesapAdi, Klasor, AltKlasor, ...] -> Outlook Folder objesi"""
    try:
        acc = None
        for i in range(1, ns.Folders.Count + 1):
            f = ns.Folders.Item(i)
            if f.Name == path[0]: acc = f; break
        if acc is None: return None
        cur = acc
        for name in path[1:]:
            found = None
            for k in range(1, cur.Folders.Count + 1):
                sub = cur.Folders.Item(k)
                if sub.Name == name: found = sub; break
            if found is None: return None
            cur = found
        return cur
    except:
        return None

def do_scan(log_cb, done_cb, selected_paths=None):
    """selected_paths None ise tum klasorler taranir (eski davranis).
    Liste verilirse sadece bu path'ler (+ alt klasorleri) taranir.
    Her path: [HesapAdi, KlasorAdi, ...]"""
    try:
        import win32com.client
        contacts, last_scan = load_data()
        results = dict(contacts)
        stats   = {"tarandi": 0, "yeni": 0}
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns      = outlook.GetNamespace("MAPI")

        if selected_paths:
            for path in selected_paths:
                folder = _resolve_folder(ns, path)
                if folder is None: continue
                log_cb(f"klasor: {' / '.join(path)}")
                walk(folder, results, last_scan, stats, log_cb)
        else:
            for i in range(1, ns.Folders.Count + 1):
                try:
                    acc = ns.Folders.Item(i)
                    log_cb(f"hesap: {acc.Name}")
                    for j in range(1, acc.Folders.Count + 1):
                        try:
                            sub = acc.Folders.Item(j)
                            log_cb(f"  klasor: {sub.Name}")
                            walk(sub, results, last_scan, stats, log_cb)
                        except: continue
                except: continue

        save_data(results)
        done_cb(results, stats)
    except Exception as e:
        done_cb(None, None, str(e))

def group_by_firma(contacts):
    firms = {}; kisisel = {}
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
    now = datetime.datetime.now(); ts = now.strftime("%d.%m.%Y %H:%M")
    wb  = Workbook()
    thin = Side(style="thin", color="CBD5E1")
    def brd(): return Border(left=thin, right=thin, top=thin, bottom=thin)
    def H(cell, val, bg="1E3A5F", fg="FFFFFF", sz=10, bold=True):
        cell.value = val
        cell.font  = Font(name="Calibri", size=sz, bold=bold, color=fg)
        cell.fill  = PatternFill("solid", fgColor=bg)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = brd()
    def D(cell, val, bg="FFFFFF", bold=False):
        cell.value = val
        cell.font  = Font(name="Calibri", size=9, bold=bold)
        cell.fill  = PatternFill("solid", fgColor=bg)
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        cell.border = brd()
    def write_sheet(title, data_firms, is_first=False, accent="1E3A5F"):
        ws = wb.active if is_first else wb.create_sheet(title[:31])
        if is_first: ws.title = title[:31]
        ws.sheet_view.showGridLines = False
        ws.merge_cells("A1:H1")
        H(ws["A1"], f"OUTLOOK TARAMA — {title.upper()}", bg=accent, sz=12)
        ws.row_dimensions[1].height = 28
        ws.merge_cells("A2:H2")
        H(ws["A2"], f"Olusturulma: {ts}  |  Toplam: {sum(len(v) for v in data_firms.values())} kisi",
          bg="2D4A6E", sz=9, bold=False)
        ws.row_dimensions[2].height = 14
        cols = ["FİRMA (İMZA)","DOMAIN","AD SOYAD","GÖREV","E-POSTA","CEP TEL","SABİT TEL","NOT"]
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
                vals = [firma_disp if pi==0 else "", dom if pi==0 else "",
                        p.get("name",""), p.get("title",""), p.get("email",""),
                        gsm_str, sabit_str, p.get("not","")]
                for ci,val in enumerate(vals, start=1):
                    D(ws.cell(row=rn, column=ci), val, bg=bg, bold=(ci<=2 and pi==0))
                ws.row_dimensions[rn].height = 16; rn += 1
            for ci in range(1, 9):
                ws.cell(row=rn, column=ci).fill = PatternFill("solid", fgColor="E2E8F0")
            ws.row_dimensions[rn].height = 3; rn += 1
        ws.freeze_panes = "A4"
        ws.auto_filter.ref = f"A3:{get_column_letter(8)}{rn-1}"
    write_sheet("Kurumsal Firmalar", firms,   is_first=True,  accent="1E3A5F")
    write_sheet("Kisisel Mailler",   kisisel, is_first=False, accent="7C3AED")
    # Özet sayfası
    ws2 = wb.create_sheet("Ozet", 0)
    ws2.sheet_view.showGridLines = False
    ws2.merge_cells("A1:F1")
    H(ws2["A1"], "OUTLOOK FİRMA REHBERİ — ÖZET", sz=14)
    ws2.row_dimensions[1].height = 34
    total_kis = sum(len(v) for v in kisisel.values())
    for ci,(lbl,val) in enumerate([
        ("KURUMSAL FİRMA", len(firms)),
        ("KURUMSAL KİŞİ",  sum(len(v) for v in firms.values())),
        ("KİŞİSEL MAİL",  total_kis),
        ("TOPLAM",         len(contacts)),
    ], start=1):
        cl = get_column_letter(ci)
        lc = ws2[f"{cl}4"]; vc = ws2[f"{cl}5"]
        lc.value = lbl; lc.font = Font(name="Calibri", size=8, bold=True, color="64748B")
        lc.alignment = Alignment(horizontal="center")
        vc.value = val; vc.font = Font(name="Calibri", size=20, bold=True, color="2563EB")
        vc.alignment = Alignment(horizontal="center")
    ws2.row_dimensions[4].height = 16; ws2.row_dimensions[5].height = 30
    wb.save(path)


# ─── ÖZEL BUTON ─────────────────────────────────────────────
class FlatBtn(tk.Frame):
    def __init__(self, parent, text, command, bg=ACC, fg="white", width=140, font=F_HDR, **kw):
        super().__init__(parent, bg=parent["bg"] if "bg" not in kw else kw.pop("bg"))
        self._cmd = command; self._bg = bg
        self._lbl = tk.Label(self, text=text, bg=bg, fg=fg, font=font,
                             width=width//7, padx=10, pady=7, cursor="hand2", relief="flat")
        self._lbl.pack()
        self._lbl.bind("<Enter>",    lambda e: self._lbl.config(bg=self._dk(bg)))
        self._lbl.bind("<Leave>",    lambda e: self._lbl.config(bg=self._bg))
        self._lbl.bind("<Button-1>", lambda e: self._cmd() if self._cmd else None)
    def _dk(self, h):
        r=int(h[1:3],16); g=int(h[3:5],16); b=int(h[5:7],16); f=0.82
        return f"#{int(r*f):02x}{int(g*f):02x}{int(b*f):02x}"
    def set_text(self, t): self._lbl.config(text=t)
    def set_state(self, on):
        if on:
            self._bg = ACC; self._lbl.config(bg=ACC, cursor="hand2")
            self._lbl.bind("<Button-1>", lambda e: self._cmd() if self._cmd else None)
        else:
            self._bg = "#94A3B8"; self._lbl.config(bg="#94A3B8", cursor="")
            self._lbl.unbind("<Button-1>")


# ─── DÜZENLEME PENCERESİ ────────────────────────────────────
class EditDialog(tk.Toplevel):
    """Tek kişi kaydını düzenleme penceresi"""
    def __init__(self, parent, person: dict, on_save):
        super().__init__(parent)
        self.title("Kişi Düzenle")
        self.resizable(False, False)
        self.grab_set()
        self.configure(bg=PANEL)
        self._person  = person
        self._on_save = on_save
        self._vars    = {}
        self._build()
        # Ortala
        self.update_idletasks()
        px = parent.winfo_rootx() + parent.winfo_width()//2  - self.winfo_width()//2
        py = parent.winfo_rooty() + parent.winfo_height()//2 - self.winfo_height()//2
        self.geometry(f"+{px}+{py}")

    def _build(self):
        # Başlık
        hdr = tk.Frame(self, bg=HDR_BG, pady=12); hdr.pack(fill="x")
        tk.Label(hdr, text="✏️  Kişi Bilgilerini Düzenle",
                 font=("Segoe UI",11,"bold"), bg=HDR_BG, fg=HDR_FG).pack(padx=16)
        em = self._person.get("email","")
        tk.Label(hdr, text=em, font=("Segoe UI",8), bg=HDR_BG, fg="#93C5FD").pack()

        # Form alanları
        form = tk.Frame(self, bg=PANEL, padx=24, pady=16); form.pack(fill="x")

        fields = [
            ("Ad Soyad",         "name"),
            ("Görev / Unvan",    "title"),
            ("Firma Adı (İmza)", "firma_display"),
            ("Cep Tel (virgülle ayır)", "gsm_str"),
            ("Sabit Tel (virgülle ayır)","sabit_str"),
            ("Not",              "not"),
        ]

        # GSM ve sabit listelerini string'e çevir
        p = self._person
        p["gsm_str"]   = ", ".join(p.get("gsm",[]))
        p["sabit_str"] = ", ".join(p.get("sabit",[]))

        for row_i, (label, key) in enumerate(fields):
            tk.Label(form, text=label, font=("Segoe UI",8,"bold"),
                     bg=PANEL, fg=TEXT2, anchor="w").grid(
                row=row_i*2, column=0, sticky="w", pady=(8,1))

            var = tk.StringVar(value=p.get(key,"") or "")
            self._vars[key] = var

            entry = tk.Entry(form, textvariable=var, font=F_BODY,
                             bg=EDIT_BG, fg=TEXT, relief="solid", bd=1,
                             insertbackground=ACC, width=42)
            entry.grid(row=row_i*2+1, column=0, sticky="ew", ipady=5)

        form.columnconfigure(0, weight=1)

        # Mail adresi değiştirme (ayrı blok)
        tk.Frame(form, bg=BORDER, height=1).grid(
            row=len(fields)*2, column=0, sticky="ew", pady=12)

        tk.Label(form, text="E-posta Adresi  (değiştirirseniz ayrı kayıt oluşur)",
                 font=("Segoe UI",8,"bold"), bg=PANEL, fg=WARN, anchor="w").grid(
            row=len(fields)*2+1, column=0, sticky="w", pady=(0,2))

        self._email_var = tk.StringVar(value=p.get("email",""))
        tk.Entry(form, textvariable=self._email_var, font=F_BODY,
                 bg="#FFF9C4", fg=TEXT, relief="solid", bd=1,
                 insertbackground=ACC, width=42).grid(
            row=len(fields)*2+2, column=0, sticky="ew", ipady=5)

        # Butonlar
        bf = tk.Frame(self, bg="#F1F5F9", pady=12); bf.pack(fill="x")
        FlatBtn(bf, "💾  Kaydet",  self._save,   bg=SUCCESS, width=120).pack(side="right", padx=8)
        FlatBtn(bf, "✕  İptal",   self.destroy, bg="#64748B", width=100).pack(side="right", padx=4)
        FlatBtn(bf, "🗑  Kaydı Sil", self._delete, bg="#DC2626", width=120).pack(side="left", padx=8)

    def _save(self):
        p     = self._person
        new_email = self._email_var.get().strip().lower()

        # Doğrulama
        if new_email and not RE_EMAIL.match(new_email):
            messagebox.showerror("Hata", "Geçersiz e-posta adresi!", parent=self); return

        # Alanları güncelle
        p["name"]          = self._vars["name"].get().strip()
        p["title"]         = self._vars["title"].get().strip()
        p["firma_display"] = self._vars["firma_display"].get().strip()
        p["not"]           = self._vars["not"].get().strip()

        # Telefon listelerine çevir
        def parse_phones(s):
            parts = [x.strip() for x in re.split(r"[,;/]", s) if x.strip()]
            return [re.sub(r"\D","",x) for x in parts if re.sub(r"\D","",x)]

        p["gsm"]   = parse_phones(self._vars["gsm_str"].get())
        p["sabit"] = parse_phones(self._vars["sabit_str"].get())
        p["email"] = new_email or p["email"]

        self._on_save(p, new_email)
        self.destroy()

    def _delete(self):
        if messagebox.askyesno("Kaydı Sil",
            f"Bu kişi silinecek:\n{self._person.get('email','')}\n\nEmin misiniz?",
            icon="warning", parent=self):
            self._on_save(None, self._person.get("email",""))
            self.destroy()


# ─── KLASÖR SEÇİM PENCERESİ ─────────────────────────────────
class FolderPickerDialog(tk.Toplevel):
    """Outlook hesap/klasör ağacını checkbox ile gösterir;
    kullanıcı taranacak klasörleri işaretler."""
    def __init__(self, parent, tree, on_confirm):
        super().__init__(parent)
        self.title("Taranacak Klasörleri Seçin")
        self.geometry("560x620")
        self.minsize(480, 420)
        self.configure(bg=PANEL)
        self.grab_set()
        self._tree       = tree
        self._on_confirm = on_confirm
        self._check_vars = {}   # path_tuple -> tk.BooleanVar
        self._row_map    = []   # (path_tuple, frame) sıralı liste — toplu işlemler için
        self._build()
        self.update_idletasks()
        px = parent.winfo_rootx() + parent.winfo_width()//2  - self.winfo_width()//2
        py = parent.winfo_rooty() + parent.winfo_height()//2 - self.winfo_height()//2
        self.geometry(f"+{px}+{py}")

    def _build(self):
        hdr = tk.Frame(self, bg=HDR_BG, pady=12); hdr.pack(fill="x")
        tk.Label(hdr, text="📂  Taranacak Mail Klasörlerini Seçin",
                 font=("Segoe UI",11,"bold"), bg=HDR_BG, fg=HDR_FG).pack(padx=16)
        tk.Label(hdr, text="İşaretlediğiniz klasörler ve TÜM alt klasörleri taranır",
                 font=("Segoe UI",8), bg=HDR_BG, fg="#93C5FD").pack()

        # Hızlı seçim butonları
        qf = tk.Frame(self, bg=PANEL, pady=8); qf.pack(fill="x", padx=16)
        FlatBtn(qf, "✅ Tümünü Seç",    self._select_all,   bg="#0891B2", width=120).pack(side="left", padx=2)
        FlatBtn(qf, "⬜ Tümünü Kaldır", self._select_none,  bg="#94A3B8", width=130).pack(side="left", padx=2)
        FlatBtn(qf, "📥 Standart Klasörler", self._select_standard, bg="#7C3AED", width=170).pack(side="left", padx=2)

        # Scrollable ağaç
        CF = tk.Frame(self, bg=PANEL, highlightbackground=BORDER, highlightthickness=1)
        CF.pack(fill="both", expand=True, padx=16, pady=(0,8))
        cv  = tk.Canvas(CF, bg=PANEL, highlightthickness=0)
        vsb = ttk.Scrollbar(CF, orient="vertical", command=cv.yview)
        cv.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        cv.pack(side="left", fill="both", expand=True)
        inner = tk.Frame(cv, bg=PANEL)
        win = cv.create_window((0,0), window=inner, anchor="nw")
        inner.bind("<Configure>", lambda e: cv.configure(scrollregion=cv.bbox("all")))
        cv.bind("<Configure>", lambda e: cv.itemconfig(win, width=e.width))
        cv.bind_all("<MouseWheel>", lambda e: cv.yview_scroll(-1*(e.delta//120), "units"))

        for acc_node in self._tree:
            self._render_node(inner, acc_node, is_account=True)

        # Alt butonlar
        bf = tk.Frame(self, bg="#F1F5F9", pady=12); bf.pack(fill="x")
        self._count_lbl = tk.Label(bf, text="0 klasör seçili",
                                   font=F_SMALL, bg="#F1F5F9", fg=TEXT2)
        self._count_lbl.pack(side="left", padx=16)
        FlatBtn(bf, "🔄  Taramayı Başlat", self._confirm, bg=SUCCESS, width=170).pack(side="right", padx=8)
        FlatBtn(bf, "✕  İptal", self.destroy, bg="#64748B", width=100).pack(side="right", padx=4)

    def _render_node(self, parent_frame, node, is_account=False):
        path = tuple(node["path"])
        depth = node["depth"]

        row = tk.Frame(parent_frame, bg=PANEL)
        row.pack(fill="x", padx=(14*depth, 4), pady=1)

        var = tk.BooleanVar(value=False)
        self._check_vars[path] = var

        cb = tk.Checkbutton(
            row, variable=var, bg=PANEL, activebackground=PANEL,
            command=lambda p=path, v=var: self._on_toggle(p, v)
        )
        cb.pack(side="left")

        icon = "🏢" if is_account else ("📁" if depth==1 else "📂")
        font = ("Segoe UI",9,"bold") if is_account else F_BODY
        color = FIRMA_FG if is_account else TEXT
        tk.Label(row, text=f"{icon} {node['name']}", font=font, bg=PANEL, fg=color,
                 anchor="w").pack(side="left", padx=2)

        for child in node.get("children", []):
            self._render_node(parent_frame, child, is_account=False)

    def _on_toggle(self, path, var):
        # Bir klasör işaretlenince alt klasörlerini de aynı duruma getir
        val = var.get()
        for p, v in self._check_vars.items():
            if len(p) > len(path) and p[:len(path)] == path:
                v.set(val)
        self._update_count()

    def _update_count(self):
        n = sum(1 for v in self._check_vars.values() if v.get())
        self._count_lbl.config(text=f"{n} klasör seçili")

    def _select_all(self):
        for v in self._check_vars.values(): v.set(True)
        self._update_count()

    def _select_none(self):
        for v in self._check_vars.values(): v.set(False)
        self._update_count()

    def _select_standard(self):
        """Gelen Kutusu, Gönderilmiş Öğeler, Silinmiş Öğeler (TR/EN) işaretle."""
        standard_names = {
            "gelen kutusu","inbox","gönderilmiş öğeler","gonderilmis ogeler",
            "sent items","silinmiş öğeler","silinmis ogeler","deleted items",
        }
        for v in self._check_vars.values(): v.set(False)
        for path, var in self._check_vars.items():
            if len(path) >= 2 and path[-1].lower() in standard_names:
                var.set(True)
                self._on_toggle(path, var)
        self._update_count()

    def _confirm(self):
        selected = [list(p) for p, v in self._check_vars.items() if v.get()]
        if not selected:
            messagebox.showwarning("Seçim Yok", "En az bir klasör seçmelisiniz.", parent=self)
            return
        # Üst düzey seçimleri koru, alt klasörleri ayrı path olarak da göndermeye gerek yok
        # çünkü walk() zaten recursive iniyor — ama kullanıcı sadece bir alt klasörü
        # seçip üstünü seçmemiş olabilir, bu yüzden tüm işaretli path'leri gönderiyoruz.
        # Üst klasör seçiliyse, onun altındaki seçili alt klasörleri filtrele (tekrar taramayı önle).
        selected_sorted = sorted(selected, key=len)
        final = []
        for p in selected_sorted:
            if not any(len(f) < len(p) and p[:len(f)] == f for f in final):
                final.append(p)
        self.destroy()
        self._on_confirm(final)


# ─── ANA UYGULAMA ───────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Outlook Firma Rehberi v4.1 — Poliport")
        self.geometry("1280x760"); self.minsize(960, 580)
        self.configure(bg=BG)
        try: self.state("zoomed")
        except: pass
        self._contacts = {}; self._firms = {}; self._kisisel = {}
        self._scanning = False; self._mode = "kurumsal"
        self._build_ui(); self._load_existing()

    # ── UI ──────────────────────────────────────────────────
    def _build_ui(self):
        self._build_header(); self._build_toolbar()
        self._build_body();   self._build_status()

    def _build_header(self):
        hdr = tk.Frame(self, bg=HDR_BG, height=76)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        L = tk.Frame(hdr, bg=HDR_BG); L.pack(side="left", padx=20, pady=12)
        tk.Label(L, text="📋", font=("Segoe UI",20), bg=HDR_BG, fg=HDR_FG).pack(side="left", padx=(0,10))
        T = tk.Frame(L, bg=HDR_BG); T.pack(side="left")
        tk.Label(T, text="Outlook Firma Rehberi",
                 font=("Segoe UI",14,"bold"), bg=HDR_BG, fg=HDR_FG).pack(anchor="w")
        tk.Label(T, text="Tüm hesaplar · Gelen + Gönderilen + Silinmiş + Alt Klasörler · Tekrarsız · Düzenlenebilir",
                 font=("Segoe UI",8), bg=HDR_BG, fg="#93C5FD").pack(anchor="w")
        R = tk.Frame(hdr, bg=HDR_BG); R.pack(side="right", padx=20)
        self._stats = {}
        for lbl, key, col in [("Firma","firma","#1E40AF"),("Kişi","kisi","#1E40AF"),
                               ("Kişisel","kisisel","#7C3AED"),("Toplam","toplam","#059669")]:
            sf = tk.Frame(R, bg=col, padx=14, pady=6); sf.pack(side="left", padx=3)
            n  = tk.Label(sf, text="0", font=("Segoe UI",16,"bold"), bg=col, fg="#DBEAFE"); n.pack()
            tk.Label(sf, text=lbl, font=("Segoe UI",8), bg=col, fg="#BFDBFE").pack()
            self._stats[key] = n
        self._last_lbl = tk.Label(hdr, text="Son tarama: henüz yapılmadı",
                                  font=("Segoe UI",8), bg=HDR_BG, fg="#93C5FD")
        self._last_lbl.place(relx=0.5, rely=0.92, anchor="s")

    def _build_toolbar(self):
        bar = tk.Frame(self, bg=PANEL, pady=9,
                       highlightbackground=BORDER, highlightthickness=1); bar.pack(fill="x")
        inner = tk.Frame(bar, bg=PANEL); inner.pack(fill="x", padx=16)
        sf = tk.Frame(inner, bg=PANEL, highlightbackground=BORDER, highlightthickness=1)
        sf.pack(side="left", fill="x", expand=True)
        tk.Label(sf, text="🔍", bg=PANEL, fg=TEXT3, font=("Segoe UI",10)).pack(side="left", padx=(8,0))
        self._sv = tk.StringVar()
        self._sv.trace_add("write", lambda *a: self._filter(self._sv.get()))
        tk.Entry(sf, textvariable=self._sv, bg=PANEL, fg=TEXT, font=F_BODY,
                 bd=0, relief="flat", insertbackground=ACC).pack(
            side="left", fill="x", expand=True, padx=6, pady=7)
        BF = tk.Frame(inner, bg=PANEL); BF.pack(side="right", padx=(10,0))
        self._btn_kur = FlatBtn(BF, "🏢 Kurumsal", lambda: self._set_mode("kurumsal"), bg=ACC, width=115)
        self._btn_kur.pack(side="left", padx=2)
        self._btn_kis = FlatBtn(BF, "✉ Kişisel", lambda: self._set_mode("kisisel"), bg="#7C3AED", width=105)
        self._btn_kis.pack(side="left", padx=2)
        tk.Frame(BF, bg=BORDER, width=1, height=28).pack(side="left", padx=8)
        FlatBtn(BF, "📥 Excel İçe Aktar", self._import_excel, bg="#0891B2", width=148).pack(side="left", padx=2)
        self._scan_btn = FlatBtn(BF, "🔄 Tarama Yap", self._start_scan, bg=ACC, width=130)
        self._scan_btn.pack(side="left", padx=2)
        FlatBtn(BF, "📤 Excel İndir", self._export_excel, bg=SUCCESS, width=125).pack(side="left", padx=2)
        FlatBtn(BF, "🗑 Sıfırla", self._reset_data, bg=WARN, width=95).pack(side="left", padx=2)

    def _build_body(self):
        body = tk.Frame(self, bg=BG); body.pack(fill="both", expand=True, padx=12, pady=(8,0))
        # Sol — firma listesi
        LEFT = tk.Frame(body, bg=PANEL, width=240,
                        highlightbackground=BORDER, highlightthickness=1)
        LEFT.pack(side="left", fill="y", padx=(0,8)); LEFT.pack_propagate(False)
        self._firma_title = tk.Label(LEFT, text="KURUMSAL FİRMALAR",
                                     font=("Segoe UI",8,"bold"), bg=PANEL, fg=TEXT3)
        self._firma_title.pack(anchor="w", padx=12, pady=(10,4))
        self._firma_var = tk.StringVar(value=[])
        self._firma_lb  = tk.Listbox(LEFT, listvariable=self._firma_var, font=F_BODY,
                                     bg=PANEL, fg=TEXT, selectbackground=ACC,
                                     selectforeground="white", activestyle="none",
                                     bd=0, highlightthickness=0, relief="flat")
        sb = ttk.Scrollbar(LEFT, orient="vertical", command=self._firma_lb.yview)
        self._firma_lb.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self._firma_lb.pack(fill="both", expand=True, padx=4, pady=(0,8))
        self._firma_lb.bind("<<ListboxSelect>>", self._on_firma_sel)
        # Sağ — kişi tablosu
        RIGHT = tk.Frame(body, bg=PANEL, highlightbackground=BORDER, highlightthickness=1)
        RIGHT.pack(side="left", fill="both", expand=True)
        self._det_hdr = tk.Frame(RIGHT, bg=FIRMA_BG, pady=10); self._det_hdr.pack(fill="x")
        self._det_name = tk.Label(self._det_hdr, text="← Soldan bir firma seçin",
                                  font=F_TITLE, bg=FIRMA_BG, fg=FIRMA_FG)
        self._det_name.pack(side="left", padx=16)
        self._det_cnt = tk.Label(self._det_hdr, text="", font=F_SMALL, bg=FIRMA_BG, fg=TEXT2)
        self._det_cnt.pack(side="right", padx=16)
        # Kolon başlıkları
        CH = tk.Frame(RIGHT, bg="#E2E8F0", pady=5); CH.pack(fill="x")
        for txt, w in [("Ad Soyad",20),("Görev / Unvan",24),("E-posta",26),
                        ("Cep Tel",13),("Sabit Tel",12),("Not",12),("",5)]:
            tk.Label(CH, text=txt.upper(), font=("Segoe UI",8,"bold"),
                     bg="#E2E8F0", fg=TEXT3, width=w, anchor="w").pack(side="left", padx=8)
        # Scrollable tablo
        CF = tk.Frame(RIGHT, bg=PANEL); CF.pack(fill="both", expand=True)
        self._cv = tk.Canvas(CF, bg=PANEL, highlightthickness=0)
        vsb = ttk.Scrollbar(CF, orient="vertical", command=self._cv.yview)
        self._cv.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y"); self._cv.pack(side="left", fill="both", expand=True)
        self._tbl  = tk.Frame(self._cv, bg=PANEL)
        self._twin = self._cv.create_window((0,0), window=self._tbl, anchor="nw")
        self._tbl.bind("<Configure>",
                       lambda e: self._cv.configure(scrollregion=self._cv.bbox("all")))
        self._cv.bind("<Configure>",
                      lambda e: self._cv.itemconfig(self._twin, width=e.width))
        self._cv.bind_all("<MouseWheel>",
                          lambda e: self._cv.yview_scroll(-1*(e.delta//120), "units"))

    def _build_status(self):
        sf = tk.Frame(self, bg="#E2E8F0", height=26); sf.pack(fill="x", side="bottom")
        sf.pack_propagate(False)
        self._st   = tk.Label(sf, text="Hazır", font=F_SMALL, bg="#E2E8F0", fg=TEXT2)
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
            self._firma_lb.selection_set(0); self._show_firma(list(firms.keys())[0])
        else:
            for w in self._tbl.winfo_children(): w.destroy()
            self._det_name.config(text="Bu bölümde veri yok"); self._det_cnt.config(text="")

    # ── Veri ────────────────────────────────────────────────
    def _load_existing(self):
        self._contacts, last = load_data()
        self._firms, self._kisisel = group_by_firma(self._contacts)
        self._refresh_list(); self._update_stats()
        if last: self._last_lbl.config(text=f"Son tarama: {last.strftime('%d.%m.%Y  %H:%M:%S')}")
        self._status("Veriler yüklendi." if self._contacts else "Veri yok — 'Tarama Yap' butonuna tıklayın.")

    def _refresh_list(self):
        firms = self._firms if self._mode=="kurumsal" else self._kisisel
        self._firma_var.set(list(firms.keys()))
        if firms:
            self._firma_lb.selection_set(0); self._show_firma(list(firms.keys())[0])

    def _update_stats(self):
        self._stats["firma"].config(text=str(len(self._firms)))
        self._stats["kisi"].config(text=str(sum(len(v) for v in self._firms.values())))
        self._stats["kisisel"].config(text=str(sum(len(v) for v in self._kisisel.values())))
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

            nm  = p.get("name","")  or ""
            tl  = p.get("title","") or ""
            em  = p.get("email","") or ""
            nt  = p.get("not","")   or ""
            gsm_list   = p.get("gsm",[])
            sabit_list = p.get("sabit",[])

            tk.Label(row, text=nm, bg=bg, fg=TEXT,
                     font=("Segoe UI",9,"bold") if nm else F_BODY,
                     width=20, anchor="w", wraplength=145).pack(side="left", padx=8, pady=6)
            tk.Label(row, text=tl, bg=bg, fg=TEXT2, font=F_BODY,
                     width=24, anchor="w", wraplength=165).pack(side="left", padx=8, pady=6)
            tk.Label(row, text=em, bg=bg, fg=ACC, font=F_BODY,
                     width=26, anchor="w", wraplength=185).pack(side="left", padx=8, pady=6)

            # Telefon
            ph = tk.Frame(row, bg=bg); ph.pack(side="left", padx=4, pady=4)
            for g in gsm_list:
                tk.Label(ph, text=f"📱 {self._fmt_gsm(g)}", bg=TAG_G_BG, fg=TAG_G_FG,
                         font=F_SMALL, padx=5, pady=2).pack(side="left", padx=2)
            for s in sabit_list:
                tk.Label(ph, text=f"☎ {self._fmt_sabit(s)}", bg=TAG_S_BG, fg=TAG_S_FG,
                         font=F_SMALL, padx=5, pady=2).pack(side="left", padx=2)

            # Not
            if nt:
                tk.Label(row, text=f"📝 {nt[:30]}", bg="#FFFBEB", fg="#92400E",
                         font=F_SMALL, padx=6, pady=2).pack(side="left", padx=4)

            # ✏️ Düzenle butonu — her satırda
            edit_btn = tk.Label(row, text="✏️", font=("Segoe UI",11),
                                bg=bg, fg=ACC, cursor="hand2", padx=6)
            edit_btn.pack(side="right", padx=6)
            edit_btn.bind("<Button-1>", lambda e, person=p, f=firma: self._open_edit(person, f))
            edit_btn.bind("<Enter>", lambda e, btn=edit_btn, b=bg: btn.config(bg=EDIT_BG))
            edit_btn.bind("<Leave>", lambda e, btn=edit_btn, b=bg: btn.config(bg=b))

            tk.Frame(self._tbl, bg=BORDER, height=1).pack(fill="x")

        self._cv.yview_moveto(0)

    # ── Düzenleme ────────────────────────────────────────────
    def _open_edit(self, person, firma):
        def on_save(updated, new_email):
            old_email = person.get("email","")
            if updated is None:
                # Sil
                if old_email in self._contacts:
                    del self._contacts[old_email]
                self._status(f"🗑  {old_email} silindi.")
            else:
                # Güncelle
                if new_email and new_email != old_email:
                    # Mail adresi değişti — eski kaydı sil, yeniyi ekle
                    if old_email in self._contacts:
                        del self._contacts[old_email]
                    updated["email"] = new_email
                    updated["domain"] = new_email.split("@")[-1].lower()
                    updated["firma"]  = updated["domain"].split(".")[0].upper()
                    self._contacts[new_email] = updated
                else:
                    self._contacts[old_email] = updated
                self._status(f"✅  {updated.get('email','')} güncellendi.")

            save_data(self._contacts)
            self._firms, self._kisisel = group_by_firma(self._contacts)
            self._update_stats()
            # Aynı firmayı yeniden göster
            firms = self._firms if self._mode=="kurumsal" else self._kisisel
            self._firma_var.set(list(firms.keys()))
            if firma in firms:
                idx = list(firms.keys()).index(firma)
                self._firma_lb.selection_clear(0, tk.END)
                self._firma_lb.selection_set(idx)
                self._show_firma(firma)
            elif firms:
                self._firma_lb.selection_set(0)
                self._show_firma(list(firms.keys())[0])

        EditDialog(self, dict(person), on_save)

    def _fmt_gsm(self, digits):
        d = re.sub(r"\D","",digits)
        if d.startswith("90") and len(d)==12: d = "0" + d[2:]
        if len(d)==10: return f"0{d[1:4]} {d[4:7]} {d[7:9]} {d[9:11]}"
        return digits

    def _fmt_sabit(self, digits):
        d = re.sub(r"\D","",digits)
        if len(d)==10: return f"0{d[1:4]} {d[4:7]} {d[7:9]} {d[9:11]}"
        return digits

    # ── Arama ───────────────────────────────────────────────
    def _filter(self, q=""):
        q = q.strip().lower()
        firms = self._firms if self._mode=="kurumsal" else self._kisisel
        items = [
            f for f, people in firms.items()
            if not q or q in f.lower()
            or any(q in (p.get("name","") + p.get("title","") +
                         p.get("email","") + p.get("firma_display","") +
                         p.get("not","")).lower() for p in people)
        ]
        self._firma_var.set(items)
        if items:
            self._firma_lb.selection_set(0); self._show_firma(items[0])

    # ── Excel içe aktar ─────────────────────────────────────
    def _import_excel(self):
        path = filedialog.askopenfilename(
            filetypes=[("Excel Dosyası","*.xlsx *.xls")], title="Mail Listesi Excel Dosyasını Seç")
        if not path: return
        try:
            self._status("Excel içe aktarılıyor...")
            from openpyxl import load_workbook
            wb = load_workbook(path, read_only=True); yeni = 0
            contacts = dict(self._contacts)
            for sheet in wb.sheetnames:
                ws = wb[sheet]
                for row in ws.iter_rows(values_only=True):
                    if not row or not row[0]: continue
                    email_val = None
                    for cell in row:
                        if cell and RE_EMAIL.match(str(cell).strip()):
                            email_val = str(cell).strip().lower(); break
                    if not email_val or email_val in contacts: continue
                    dom = email_val.split("@")[-1].lower()
                    contacts[email_val] = {
                        "email": email_val, "domain": dom,
                        "firma": dom.split(".")[0].upper(),
                        "firma_display": dom.split(".")[0].upper(),
                        "name":"","title":"","gsm":[],"sabit":[],"not":"",
                        "kisisel": dom in KISISEL_DOMAINS,
                    }; yeni += 1
            save_data(contacts)
            self._contacts = contacts
            self._firms, self._kisisel = group_by_firma(contacts)
            self._refresh_list(); self._update_stats()
            self._status(f"✅  Excel aktarıldı — {yeni} yeni e-posta eklendi.")
            messagebox.showinfo("İçe Aktarma Tamam", f"{yeni} yeni e-posta eklendi.\nToplam: {len(contacts)}")
        except Exception as e:
            messagebox.showerror("Hata", f"Excel okunamadı:\n{e}")

    # ── Tarama ──────────────────────────────────────────────
    def _start_scan(self):
        if self._scanning: return
        self._scanning = True
        self._scan_btn.set_state(False); self._scan_btn.set_text("⏳ Klasörler okunuyor...")
        self._status("Outlook klasör listesi alınıyor, lütfen bekleyin...")

        def fetch_tree():
            try:
                tree = get_folder_tree()
                self.after(0, lambda: self._show_folder_picker(tree))
            except Exception as e:
                self.after(0, lambda: self._scan_tree_error(str(e)))

        threading.Thread(target=fetch_tree, daemon=True).start()

    def _scan_tree_error(self, err):
        self._scanning = False
        self._scan_btn.set_state(True); self._scan_btn.set_text("🔄 Tarama Yap")
        messagebox.showerror("Outlook Hatası",
            f"Outlook'a bağlanılamadı:\n\n{err}\n\nOutlook'un açık olduğundan emin olun.")
        self._status("Klasör listesi alınamadı.")

    def _show_folder_picker(self, tree):
        self._scan_btn.set_state(True); self._scan_btn.set_text("🔄 Tarama Yap")
        self._scanning = False
        if not tree:
            messagebox.showwarning("Klasör Yok", "Outlook'ta hiç hesap/klasör bulunamadı.")
            self._status("Klasör bulunamadı."); return

        def on_confirm(selected_paths):
            self._run_scan(selected_paths)

        FolderPickerDialog(self, tree, on_confirm)

    def _run_scan(self, selected_paths):
        if self._scanning: return
        self._scanning = True
        self._scan_btn.set_state(False); self._scan_btn.set_text("⏳ Taranıyor...")
        self._status(f"{len(selected_paths)} klasör taranıyor, lütfen bekleyin...")

        def log(msg): self.after(0, lambda: self._prog.config(text=msg[-60:]))
        def done(contacts, stats, err=None): self.after(0, lambda: self._scan_done(contacts, stats, err))
        threading.Thread(target=do_scan, args=(log, done, selected_paths), daemon=True).start()

    def _scan_done(self, contacts, stats, err):
        self._scanning = False
        self._scan_btn.set_state(True); self._scan_btn.set_text("🔄 Tarama Yap")
        self._prog.config(text="")
        if err:
            messagebox.showerror("Tarama Hatası",
                f"Outlook'a bağlanılamadı:\n\n{err}\n\nOutlook'un açık olduğundan emin olun.")
            self._status("Tarama başarısız."); return
        self._contacts = contacts
        self._firms, self._kisisel = group_by_firma(contacts)
        self._refresh_list(); self._update_stats()
        now = datetime.datetime.now()
        self._last_lbl.config(text=f"Son tarama: {now.strftime('%d.%m.%Y  %H:%M:%S')}")
        self._status(f"✅  {stats['tarandi']} mail okundu, {stats['yeni']} yeni kişi eklendi")
        messagebox.showinfo("Tarama Tamamlandı",
            f"Tarama tamamlandı!\n\nOkunan mail  : {stats['tarandi']}\n"
            f"Yeni kişi    : {stats['yeni']}\nToplam kişi  : {len(contacts)}\n"
            f"Kurumsal     : {len(self._firms)}\nKişisel      : {sum(len(v) for v in self._kisisel.values())}")

    # ── Excel dışa aktar ────────────────────────────────────
    def _export_excel(self):
        if not self._contacts:
            messagebox.showwarning("Veri Yok", "Önce tarama yapın."); return
        ts = datetime.datetime.now().strftime("%d.%m.%Y_%H-%M")
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel Dosyası","*.xlsx")],
            initialfile=f"Outlook_Firma_Rehberi_{ts}.xlsx", title="Excel Kaydet")
        if not path: return
        try:
            self._status("Excel oluşturuluyor...")
            build_excel(self._contacts, path)
            self._status(f"✅  Kaydedildi: {Path(path).name}")
            if messagebox.askyesno("Excel Hazır", f"Kaydedildi:\n{path}\n\nŞimdi açmak ister misiniz?"):
                os.startfile(path)
        except Exception as e:
            messagebox.showerror("Hata", f"Excel oluşturulamadı:\n{e}")

    # ── Sıfırla ─────────────────────────────────────────────
    def _reset_data(self):
        if not messagebox.askyesno("Sıfırla", "Tüm veriler silinecek.\nEmin misiniz?", icon="warning"): return
        if DATA_FILE.exists(): DATA_FILE.unlink()
        self._contacts = {}; self._firms = {}; self._kisisel = {}
        self._refresh_list(); self._update_stats()
        self._last_lbl.config(text="Son tarama: henüz yapılmadı")
        self._status("Veriler sıfırlandı.")

    def _status(self, msg): self._st.config(text=msg)


if __name__ == "__main__":
    App().mainloop()
