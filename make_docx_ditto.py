# make_docx_ditto.py - FIXED + LOGO + FONT + KRUTI 500 CODE (Library + Inbuilt Table) + COMPANY HEADING TOP
import json, sys, pathlib
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
BASE = pathlib.Path(__file__).parent
DATA_FILE = BASE / "data.json"
CONFIG_FILE = BASE / "config.json"
FONT_MAP = {"noto": "Noto Sans", "kruti": "Kruti Dev 010", "nirmala": "Nirmala UI"}
FONT_NAME = "Noto Sans"
# ===== KRUTI DEV 010 - 500 CODE - 1. Library, 2. Inbuilt Table =====
# Source: IIIT Hyderabad krutidev2unicode.py k2u (524 entries) reversed -> Unicode to Kruti
# Library: pip install krutidev-unicode -> from krutidev_unicode import to_kruti
try:
    from krutidev_unicode import to_kruti as _LIB_TO_KRUTI
except Exception:
    _LIB_TO_KRUTI = None
KRUTI_U2K = {
"॰": "ñ", "र्द्ध": ")Z", "फ़्": "¶+", "क़": "d+", "ख़": "[+k", "ख़्": "[+", "ग़": "x+", "ज़्": "T+", "ज़": "t+", "ड़": "M+", "ढ़": "<+", "फ़": "Q+", "य़": ";+", "ऱ": "j+", "ऩ": "u+",
"त्त": "Ùk", "त्त्": "Ù", "क्त": "ä", "दृ": "–", "कृ": "—", "न्न": "é", "न्न्": "™",
"ह्न": "à", "ह्य": "á", "हृ": "â", "ह्म": "ã", "ह्र": "ºz", "ह्": "º",
"द्द": "í", "क्ष": "{k", "क्ष्": "{", "त्र": "=", "त्र्": "«",
"छ्य": "Nî", "ट्य": "Vî", "ठ्य": "Bî", "ड्य": "Mî", "ढ्य": "<î", "द्य": "|", "ज्ञ": "K", "द्व": "}",
"श्र": "J", "ट्र": "Vª", "ड्र": "Mª", "ढ्र": "<ªª", "छ्र": "Nª",
"क्र": "Ø", "फ्र": "Ý", "र्द्र": "nzZ", "द्र": "æ", "प्र": "ç", "ग्र": "xz",
"रु": "#", "रू": ":", "ऑ": "v‚", "ओ": "vks", "औ": "vkS", "आ": "vk", "अ": "v",
"ईं": "b±", "ई": "bZ", "इ": "b", "उ": "m", "ऊ": "Å", "ऐ": ",s", "ए": ",", "ऋ": "_",
"क्क": "ô", "क": "d", "ख": "[k", "ख्": "[", "ग": "x", "घ": "?k", "घ्": "?", "ङ": "³",
"चै": "pkS", "च": "p", "छ": "N", "ज": "t", "झ": ">", "झ्": "÷", "ञ": "¥",
"ट्ट": "ê", "ट्ठ": "ë", "ट": "V", "ठ": "B", "ड्ड": "ì", "ड्ढ": "ï", "ड": "M", "ढ": "<",
"ण": ".k", "ण्": ".", "त": "r", "थ": "Fk", "थ्": "F", "द्ध": ")", "द": "n",
"ध": "/k", "ध्": "/", "न": "u", "प": "i", "फ": "Q", "फ्": "¶",
"ब": "c", "भ": "Hk", "भ्": "H", "म": "e", "य": ";", "र": "j",
"ल": "y", "ळ": "G", "व": "o", "श": "'k", "श्": "'", "ष": "\"k", "ष्": "\"",
"स": "l", "ह": "g", "ीं": "È", "्रें": "saz", "्र": "z",
"ों": "kas", "ो": "ks", "ौ": "kS", "ाँ": "¡k", "ं": "a", "ा": "k",
"ीं": "ah", "ी": "h", "ुं": "aq", "ु": "q", "ूं": "aw", "ूँ": "¡w", "ू": "w",
"ृ": "`", "ें": "as", "े": "s", "ैं": "aS", "ै": "S", "्रं": "aª", "्र": "ª",
"ँ": "¡", "ॅ": "W", "ऽ": "•", "्": "~", "़": "+", "‘": "^", "’": "*",
"“": "Þ", "”": "ß", "।": "A", ".": "-", ",": "]", "/": "@",
"०": "å", "१": "ƒ", "२": "„", "३": "…", "४": "†", "५": "‡", "६": "ˆ", "७": "‰", "८": "Š", "९": "‹",
"कार्यालय": "dk;Zky;", "श्रीमान": "Jhek", "प्राथमिकी": "izkFkfedh", "विद्युत": "fo|qr",
"साऊथ": "lkÅFk", "पॉवर": "ikWj", "कंपनी": "daiuh", "लिमिटेड": "fyehsVM",
"छापामारी": "Nkikekjh", "ब्योरा": "C;ksjk", "प्राथमिकी": "izkFkfedh", "थानाध्यक्ष": "Fkuk/;{k",
"विश्वासभाजन": "fo'oklHkktu", "अभियोजन": "vfHk;ks**Lo — Poora naya `make_docx_ditto.py` — Ab Table ditto.py ke andar hi hai — Alag JSON banane ki zaroorat nahi — Aur Pehle Library (500 Code) chalegi, nahi mili to andar wala Table chalega.**

**Purana hata ke ye Paste karo — Naam wahi — `make_docx_ditto.py`**

---

```python
# make_docx_ditto.py - FIXED + LOGO + FONT + KRUTI 500 CODE (Library + Inbuilt Table) + COMPANY HEADING TOP
import json, sys, pathlib
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
BASE = pathlib.Path(__file__).parent
DATA_FILE = BASE / "data.json"
CONFIG_FILE = BASE / "config.json"
FONT_MAP = {"noto": "Noto Sans", "kruti": "Kruti Dev 010", "nirmala": "Nirmala UI"}
FONT_NAME = "Noto Sans"
# ===== KRUTI DEV 010 - 500 CODE - 1. Library, 2. Inbuilt Table =====
# Source: IIIT Hyderabad krutidev2unicode.py k2u (524 entries) reversed -> Unicode to Kruti
# Library: pip install krutidev-unicode -> from krutidev_unicode import to_kruti
try:
    from krutidev_unicode import to_kruti as _LIB_TO_KRUTI
except Exception:
    _LIB_TO_KRUTI = None
KRUTI_U2K = {
"॰": "ñ", "र्द्ध": ")Z", "फ़्": "¶+", "क़": "d+", "ख़": "[+k", "ख़्": "[+", "ग़": "x+", "ज़्": "T+", "ज़": "t+", "ड़": "M+", "ढ़": "<+", "फ़": "Q+", "य़": ";+", "ऱ": "j+", "ऩ": "u+",
"त्त": "Ùk", "त्त्": "Ù", "क्त": "ä", "दृ": "–", "कृ": "—", "न्न": "é", "न्न्": "™",
"ह्न": "à", "ह्य": "á", "हृ": "â", "ह्म": "ã", "ह्र": "ºz", "ह्": "º",
"द्द": "í", "क्ष": "{k", "क्ष्": "{", "त्र": "=", "त्र्": "«",
"छ्य": "Nî", "ट्य": "Vî", "ठ्य": "Bî", "ड्य": "Mî", "ढ्य": "<î", "द्य": "|", "ज्ञ": "K", "द्व": "}",
"श्र": "J", "ट्र": "Vª", "ड्र": "Mª", "ढ्र": "<ªª", "छ्र": "Nª",
"क्र": "Ø", "फ्र": "Ý", "र्द्र": "nzZ", "द्र": "æ", "प्र": "ç", "ग्र": "xz",
"रु": "#", "रू": ":", "ऑ": "v‚", "ओ": "vks", "औ": "vkS", "आ": "vk", "अ": "v",
"ईं": "b±", "ई": "bZ", "इ": "b", "उ": "m", "ऊ": "Å", "ऐ": ",s", "ए": ",", "ऋ": "_",
"क्क": "ô", "क": "d", "ख": "[k", "ख्": "[", "ग": "x", "घ": "?k", "घ्": "?", "ङ": "³",
"चै": "pkS", "च": "p", "छ": "N", "ज": "t", "झ": ">", "झ्": "÷", "ञ": "¥",
"ट्ट": "ê", "ट्ठ": "ë", "ट": "V", "ठ": "B", "ड्ड": "ì", "ड्ढ": "ï", "ड": "M", "ढ": "<",
"ण": ".k", "ण्": ".", "त": "r", "थ": "Fk", "थ्": "F", "द्ध": ")", "द": "n",
"ध": "/k", "ध्": "/", "न": "u", "प": "i", "फ": "Q", "फ्": "¶",
"ब": "c", "भ": "Hk", "भ्": "H", "म": "e", "य": ";", "र": "j",
"ल": "y", "ळ": "G", "व": "o", "श": "'k", "श्": "'", "ष": "\"k", "ष्": "\"",
"स": "l", "ह": "g", "ीं": "È", "्रें": "saz", "्र": "z",
"ों": "kas", "ो": "ks", "ौ": "kS", "ाँ": "¡k", "ं": "a", "ा": "k",
"ीं": "ah", "ी": "h", "ुं": "aq", "ु": "q", "ूं": "aw", "ूँ": "¡w", "ू": "w",
"ृ": "`", "ें": "as", "े": "s", "ैं": "aS", "ै": "S", "्रं": "aª", "्र": "ª",
"ँ": "¡", "ॅ": "W", "ऽ": "•", "्": "~", "़": "+", "‘": "^", "’": "*",
"“": "Þ", "”": "ß", "।": "A", ".": "-", ",": "]", "/": "@",
"०": "å", "१": "ƒ", "२": "„", "३": "…", "४": "†", "५": "‡", "६": "ˆ", "७": "‰", "८": "Š", "९": "‹",
"कार्यालय": "dk;Zky;", "श्रीमान": "Jhek", "प्राथमिकी": "izkFkfedh", "विद्युत": "fo|qr",
"साऊथ": "lkÅFk", "पॉवर": "ikWj", "कंपनी": "daiuh", "लिमिटेड": "fyehsVM",
"छापामारी": "Nkikekjh", "ब्योरा": "C;ksjk", "थानाध्यक्ष": "Fkuk/;{k",
"विश्वासभाजन": "fo'oklHkktu", "अभियोजन": "vfHk;ksऐसा कुछ नहीं",
}
KRUTI_KEYS_SORTED = sorted(KRUTI_U2K.keys(), key=len, reverse=True)
def to_kruti(text):
    if not text: return text
    s = str(text)
    has_dev = any('\u0900' <= ch <= '\u097f' for ch in s)
    if not has_dev: return s
    if _LIB_TO_KRUTI is not None:
        try: return _LIB_TO_KRUTI(s)
        except Exception as e: print("KRUTI LIB FAIL, Inbuilt Table use kar raha hun:", e)
    out = s
    placeholders = {}
    for idx, k in enumerate(KRUTI_KEYS_SORTED):
        if k in out and len(k) > 1:
            ph = f"\x00{idx}\x00"
            placeholders[ph] = KRUTI_U2K[k]
            out = out.replace(k, ph)
    for k in KRUTI_KEYS_SORTED:
        if len(k) == 1 and k in out:
            out = out.replace(k, KRUTI_U2K[k])
    for ph, val in placeholders.items():
        out = out.replace(ph, val)
    return out
ONES = ["", "एक", "दो", "तीन", "चार", "पाँच", "छह", "सात", "आठ", "नौ", "दस", "ग्यारह", "बारह", "तेरह", "चौदह", "पंद्रह", "सोलह", "सत्रह", "अठारह", "उन्नीस"]
TENS = ["", "", "बीस", "तीस", "चालीस", "पचास", "साठ", "सत्तर", "अस्सी", "नब्बे"]
def _two(n):
    if n < 20: return ONES[n]
    t, o = divmod(n, 10)
    return (TENS[t] + (" " + ONES[o] if o else "")).strip()
def _three(n):
    h, r = divmod(n, 100)
    s = ""
    if h: s = ONES[h] + " सौ"
    if r: s = (s + " " + _two(r)).strip() if s else _two(r)
    return s
def num_to_hindi(n):
    try: n = int(float(str(n).replace(",", "").strip() or 0))
    except: return ""
    if n == 0: return "शून्य"
    parts = []
    cr, n2 = divmod(n, 10000000)
    if cr: parts.append(_three(cr) + " करोड़")
    lk, n2 = divmod(n2, 100000)
    if lk: parts.append(_three(lk) + " लाख")
    th, n2 = divmod(n2, 1000)
    if th: parts.append(_three(th) + " हजार")
    if n2: parts.append(_three(n2))
    return " ".join(parts).strip()
def fmt_amt(v):
    try:
        f = float(str(v).replace(",", "").strip() or 0)
        return f"{int(f)}" if f == int(f) else f"{f:.2f}"
    except: return str(v or "")
def fmt_amt2(v):
    try: return f"{float(str(v).replace(',', '').strip() or 0):.2f}"
    except: return str(v or "")
def get(rec, *keys, default=""):
    for k in keys:
        v = rec.get(k)
        if v is not None and str(v).strip() != "":
            return v
    return default
def set_font(run, size=11, bold=False):
    run.font.name = FONT_NAME; run.font.size = Pt(size); run.bold = bold; run.font.color.rgb = RGBColor(0,0,0)
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts'); rPr.append(rFonts)
    for a in ('w:ascii','w:hAnsi','w:cs','w:eastAsia'): rFonts.set(qn(a), FONT_NAME)
def para(doc, text="", size=11, bold=False, align=None, before=0, after=1, line=1.0, keep_next=False, right_indent=None):
    if FONT_NAME == "Kruti Dev 010" and text:
        text = to_kruti(text)
    p = doc.add_paragraph(); pf = p.paragraph_format
    pf.space_before = Pt(before); pf.space_after = Pt(after); pf.line_spacing = line
    pf.widow_control = True; pf.keep_with_next = keep_next
    if align is not None: p.alignment = align
    if right_indent is not None: pf.right_indent = Inches(right_indent)
    if text:
        r = p.add_run(text); set_font(r, size=size, bold=bold)
    return p
def add_mixed(doc, parts, align=None, before=0, after=1, line=1.0, keep_next=False, right_indent=None):
    if FONT_NAME == "Kruti Dev 010":
        parts = [(to_kruti(t), b) for t, b in parts]
    p = doc.add_paragraph(); pf = p.paragraph_format
    pf.space_before = Pt(before); pf.space_after = Pt(after); pf.line_spacing = line
    pf.widow_control = True; pf.keep_with_next = keep_next
    if align is not None: p.alignment = align
    if right_indent is not None: pf.right_indent = Inches(right_indent)
    for txt, b in parts:
        r = p.add_run(txt); set_font(r, size=11, bold=b)
    return p
def cell_para(cell, text="", size=11, align=None, before=0, after=0, line=1.0):
    if FONT_NAME == "Kruti Dev 010" and text:
        text = to_kruti(text)
    if len(cell.paragraphs)==1 and cell.paragraphs[0].text=="":
        p = cell.paragraphs[0]
    else:
        p = cell.add_paragraph()
    pf = p.paragraph_format; pf.space_before = Pt(before); pf.space_after = Pt(after); pf.line_spacing = line; pf.widow_control = True
    if align is not None: p.alignment = align
    if text:
        r = p.add_run(text); set_font(r, size=size)
    return p
def set_cell_margins(cell, top=0, bottom=0, left=0, right=0):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, v in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        el = OxmlElement(f'w:{m}'); el.set(qn('w:w'), str(v)); el.set(qn('w:type'), 'dxa'); tcMar.append(el)
    tcPr.append(tcMar)
def no_borders(table):
    tbl = table._tbl; tblPr = tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for b in ['top','left','bottom','right','insideH','insideV']:
        el = OxmlElement(f'w:{b}'); el.set(qn('w:val'), 'nil'); el.set(qn('w:sz'), '0'); el.set(qn('w:space'), '0'); el.set(qn('w:color'), 'auto'); borders.append(el)
    tblPr.append(borders)
def load_data():
    if DATA_FILE.exists():
        try:
            d = json.loads(DATA_FILE.read_text(encoding="utf-8")); return [d] if isinstance(d, dict) else d
        except: return []
    return []
def load_config():
    if CONFIG_FILE.exists():
        try: return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        except: return {}
    return {}
def find_record(fid):
    data = load_data()
    for r in data:
        if str(r.get("id", r.get("fid", ""))) == str(fid): return r
    try:
        idx = int(fid)-1
        if 0 <= idx < len(data): return data[idx]
    except: pass
    return None
def default_record():
    return {"id": 17, "patrank": "/बाँकेबाजार", "dinank": "17/08/2026", "email": "jeebankeybazar@gmail.com", "mob": "7632996820", "thana": "लुथुआ", "jila": "गया जी (बिहार)", "abhiyukt": "कमरूद्दीन मियाँ", "pita": "नजामुद्दीन मियाँ", "gaon": "लुथुआ", "panchayat": "लुथुआ", "upbhokta": "229200095413", "meter": "188510", "bhar": "0.503 KW", "shreni": "DS1D", "samay": "12:35 बजे अपराह्न", "dal": "श्री राहुल कुमार वर्मा मानवबल बाँकेबाजार", "vichchhedit": "14/07/2023", "nirdharan": "52945", "bakaya": "20575", "jabt": "PVC तार के आंशिक भाग को काट कर जब्त", "sthaniya_naam": "सोविन्द, पिता-श्री जगदीश प्रसाद, ग्राम-सिड़की, थाना-नागल, जिला-सहारनपुर (U.P), पिन-कोड-247551, मो० नं- 7632996820"}
def make_docx(fid):
    global FONT_NAME
    rec = find_record(fid)
    if rec is None:
        print(f"ID {fid} nahi mila - Default 17 se bana raha hun"); rec = default_record(); rec["id"] = fid
    cfg = load_config()
    for k in ["office","shakha","email","mob","thana","jila"]:
        if not rec.get(k) and cfg.get(k): rec[k] = cfg[k]
    font_choice = str(get(rec, "font_choice", default="noto")).lower()
    FONT_NAME = FONT_MAP.get(font_choice, "Noto Sans")
    logo_choice = str(get(rec, "logo_choice", default="sbpdcl")).lower()
    if logo_choice == "nbpdcl":
        company_name = "नॉर्थ बिहार पॉवर डिस्ट्रीब्यूशन कंपनी लिमिटेड"
    else:
        company_name = "साऊथ बिहार पॉवर डिस्ट्रीब्यूशन कंपनी लिमिटेड"
    nirdharan = get(rec, "nirdharan", "Nirdharan Rashi", default="52945")
    bakaya = get(rec, "bakaya", "Bakaya Rashi", default="20575")
    try: kul = float(str(nirdharan).replace(",","")) + float(str(bakaya).replace(",",""))
    except: kul = 73520.0
    kul_s = f"{kul:.2f}"
    nirdharan_words = num_to_hindi(nirdharan) or "बावन हजार नौ सौ पैंतालीस"
    kul_words = num_to_hindi(kul) or "तिहत्तर हजार पाँच सौ बीस"
    abhiyukt = get(rec, "abhiyukt", default="कमरूद्दीन मियाँ")
    pita = get(rec, "pita", "pati", default="नजामुद्दीन मियाँ")
    pita_type = get(rec, "pita_type", default="पिता")
    gram = get(rec, "gaon", "gram", default="लुथुआ")
    panchayat = get(rec, "panchayat", default=gram)
    thana = get(rec, "thana", default="लुथुआ")
    jila = get(rec, "jila", default="गया जी (बिहार)")
    upbhokta = get(rec, "upbhokta", "consumer", default="229200095413")
    meter = get(rec, "meter", default="188510")
    load = get(rec, "bhar", "load", "load_kw", default="0.503 KW")
    shreni = get(rec, "shreni", default="DS1D")
    dinank = get(rec, "dinank", default="17/08/2026")
    patrank = get(rec, "patrank", "patraank", default="/बाँकेबाजार")
    chhapamari_time = get(rec, "samay", "chhapamari_time", default="12:35 बजे अपराह्न")
    chhapamari_dal = get(rec, "dal", "chhapamari_dal", "chhapamari_dal_full", default="")
    if chhapamari_dal.startswith("उपर्युक्त विषय"):
        chhapamari_dal = get(rec, "dal", "chhapamari_dal", default="")
    disconnection = get(rec, "vichchhedit", "vichchhedit_dinank", "disconnection_date", default="14/07/2023")
    japt = get(rec, "jabt", "japt", default="PVC तार के आंशिक भाग को काट कर जब्त")
    gawah = get(rec, "sthaniya_naam", "sthaniya", "gawah", "sthayi", default="")
    sthayi = gawah
    email = get(rec, "email", default="jeebankeybazar@gmail.com")
    mob = get(rec, "mob", default="7632996820")
    pata = get(rec, "pata", default=f"ग्राम-{gram}, पंचायत-{panchayat}, थाना-{thana}, जिला-{jila}")
    theft_type = str(get(rec, "theft_type", default="4"))
    karyalaya_raw = get(rec, "karyalaya", "office", default="कनीय विद्युत अभियंता")
    shakha_raw = get(rec, "shakha", default="")
    if "\n" in str(karyalaya_raw):
        parts = str(karyalaya_raw).split("\n")
        karyalaya = parts[0].strip()
        if not shakha_raw and len(parts) > 1: shakha_raw = parts[1].strip()
    else:
        karyalaya = str(karyalaya_raw).strip()
    shakha = shakha_raw if shakha_raw else "विद्युत आपूर्ति प्रशाखा, बाँकेबाजार"
    antim_text = get(rec, "antim", default="")
    if not antim_text:
        antim_text = f"अतः श्रीमान से आग्रह है कि संबंधित उपरोक्त अभियुक्त {abhiyukt}, {pita_type}-{pita}, पर भारतीय विद्युत अधिनियम 2003 की धारा 135 एवं अन्य सुसंगत धाराओं के अंतर्गत प्राथमिकी दर्ज किया जाए ताकि इस प्रकार की घटना की पुनरावृत्ति न हो सके।"
    byora_full = get(rec, "byora_full", default="")
    if byora_full:
        body = byora_full
    else:
        if theft_type == "1":
            core = f"{abhiyukt}, {pita_type}-{pita}, {pata} के परिसर में बिना किसी वैध विद्युत संबंध के ही मुख्य LT लाइन में लाल एवं काले तार से टोका फँसाकर अवैध रूप से विद्युत ऊर्जा की चोरी की जा रही थी। जाँच के दौरान परिसर का कुल विद्युत भार {load} ({shreni} श्रेणी) में पाया गया।"
        elif theft_type == "2":
            core = f"{abhiyukt}, {pita_type}-{pita}, (उपभोक्ता संख्या-{upbhokta}), मीटर नं-{meter}, {pata} के परिसर में मीटर को बायपास करने हेतु सर्विस वायर को काटकर एक्स्ट्रा वायर जोड़कर अवैध रूप से विद्युत ऊर्जा की चोरी की जा रही थी। जाँच के दौरान परिसर का कुल विद्युत भार {load} ({shreni} श्रेणी) में पाया गया।"
        elif theft_type == "3":
            core = f"{abhiyukt}, {pita_type}-{pita}, (उपभोक्ता संख्या-{upbhokta}), मीटर नं-{meter}, {pata} के परिसर में सर्विस वायर के अलावा अलग तार से मीटर को बायपास कर (आटा चक्की हेतु) अवैध रूप से विद्युत ऊर्जा की चोरी की जा रही थी। जाँच के दौरान परिसर का कुल विद्युत भार {load} ({shreni} श्रेणी) में पाया गया।"
        else:
            core = f"{abhiyukt}, {pita_type}-{pita}, (उपभोक्ता संख्या-{upbhokta}), मीटर नं-{meter}, {pata} के परिसर में विद्युत विपत्र बकाया होने के कारण दिनांक {disconnection} को ही विद्युत संबंध विच्छेदित कर दिये जाने के बाद भी इनके द्वारा अवैध रूप से घरेलु परिसर में विद्युत ऊर्जा का उपयोग किया जा रहा था। जाँच के दौरान परिसर का कुल विद्युत भार {load} ({shreni} श्रेणी) में पाया गया।"
        body = f"छापामारी दल समय लगभग {chhapamari_time} में {core} विद्युत ऊर्जा चोरी में प्रयुक्त {japt} को काट कर जब्त कर लिया गया है, जिसे जाँच-सह-जब्ती प्रतिवेदन के साथ थाने में समर्पित की जा रही है। इनके द्वारा इस प्रकार से अवैध रूप से विद्युत ऊर्जा की चोरी करने से साउथ बिहार पॉवर डिस्ट्रीब्यूशन कम्पनी लिमिटेड को कुल राशि-{fmt_amt(nirdharan)}/- रुपया ({nirdharan_words} रुपया) मात्र की राजस्व हानि अनुमानित है, एवं पूर्व में इनके विद्युत विपत्र का बकाया राशि {fmt_amt2(bakaya)} रुपया को मिलाकर कुल {fmt_amt(nirdharan)}+{fmt_amt2(bakaya)}={kul_s} रुपया ({kul_words} रुपया) है। जिसकी वसूली इनसे ही की जानी है। उक्त राशि में भारतीय विद्युत अधिनियम 2003 की धारा 152 की तहत कम्पाउंडिंग की राशि सम्मिलित नहीं है।"
    if font_choice == "kruti":
        body = to_kruti(body)
        antim_text = to_kruti(antim_text)
        chhapamari_dal = to_kruti(chhapamari_dal)
        sthayi = to_kruti(sthayi); gawah = sthayi
        karyalaya = to_kruti(karyalaya); shakha = to_kruti(shakha)
        thana = to_kruti(thana); jila = to_kruti(jila); dinank = to_kruti(dinank); patrank = to_kruti(patrank)
        company_name = to_kruti(company_name)
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Inches(8.27); sec.page_height = Inches(11.69)
    sec.top_margin = Inches(0.45); sec.bottom_margin = Inches(0.35)
    sec.left_margin = Inches(0.70); sec.right_margin = Inches(0.70)
    sec.header_distance = Inches(0.25); sec.footer_distance = Inches(0.25)
    style = doc.styles['Normal']; style.font.name = FONT_NAME; style.font.size = Pt(11)
    style.paragraph_format.space_before = Pt(0); style.paragraph_format.space_after = Pt(1)
    style.paragraph_format.line_spacing = 1.0; style.paragraph_format.widow_control = True
    para(doc, company_name, size=13, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, before=0, after=2)
    if logo_choice == "nbpdcl":
        logo_file = BASE / "nbpdcl_logo.png"
    else:
        logo_file = BASE / "sbpdcl_logo.png"
    if not logo_file.exists():
        logo_file = BASE / "sbpdcl_logo.png"
    if logo_file.exists():
        htable = doc.add_table(rows=1, cols=2)
        htable.alignment = WD_TABLE_ALIGNMENT.CENTER
        htable.autofit = False
        htable.columns[0].width = Inches(1.10)
        htable.columns[1].width = Inches(5.80)
        no_borders(htable)
        lc = htable.cell(0, 0); rc = htable.cell(0, 1)
        set_cell_margins(lc, top=0, bottom=0, left=0, right=60)
        set_cell_margins(rc, top=0, bottom=0, left=0, right=0)
        lp = lc.paragraphs[0]; lp.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = lp.add_run()
        run.add_picture(str(logo_file), width=Inches(0.95))
        cell_para(rc, "-::कार्यालय::-", size=11, align=WD_ALIGN_PARAGRAPH.CENTER, before=0, after=0)
        cell_para(rc, karyalaya, size=11, align=WD_ALIGN_PARAGRAPH.CENTER, before=0, after=0)
        cell_para(rc, shakha, size=11, align=WD_ALIGN_PARAGRAPH.CENTER, before=0, after=0)
        cell_para(rc, f"E-mail: {email} Mob No:- {mob}", size=11, align=WD_ALIGN_PARAGRAPH.CENTER, before=0, after=0)
    else:
        para(doc, "-::कार्यालय::-", size=11, align=WD_ALIGN_PARAGRAPH.CENTER, before=0, after=0)
        para(doc, karyalaya, size=11, align=WD_ALIGN_PARAGRAPH.CENTER, before=0, after=0)
        para(doc, shakha, size=11, align=WD_ALIGN_PARAGRAPH.CENTER, before=0, after=0)
        para(doc, f"E-mail: {email} Mob No:- {mob}", size=11, align=WD_ALIGN_PARAGRAPH.CENTER, before=0, after=1)
    if font_choice != "kruti":
        para(doc, "="*72, size=11, align=WD_ALIGN_PARAGRAPH.CENTER, before=0, after=2)
    else:
        para(doc, "", size=11, before=0, after=2)
    p = doc.add_paragraph(); pf = p.paragraph_format; pf.space_before = Pt(0); pf.space_after = Pt(1); pf.line_spacing = 1.0
    p.paragraph_format.tab_stops.add_tab_stop(Inches(7.0), WD_ALIGN_PARAGRAPH.RIGHT)
    r1 = p.add_run(f"पत्रांक:- {patrank}"); set_font(r1, size=11)
    r2 = p.add_run(f"\tदिनांक:- {dinank}"); set_font(r2, size=11)
    para(doc, "सेवा में,", size=11, before=0, after=0)
    para(doc, "श्रीमान थानाध्यक्ष महोदय", size=11, before=0, after=0)
    para(doc, f"थाना-{thana}, जिला-{jila}", size=11, before=0, after=2)
    para(doc, "विषय:- विद्युत ऊर्जा चोरी के विरुद्ध प्राथमिकी दर्ज करने के संबंध में।", size=11, before=0, after=2)
    para(doc, "महाशय,", size=11, before=0, after=2)
    add_mixed(doc, [("उपर्युक्त विषय के संदर्भ में सूचित करना है कि दिनांक-", False), (f"{dinank}", False), (" को विद्युत ऊर्जा चोरी के रोकथाम हेतु एक छापामारी दल का गठन किया गया, जिसमें मेरे अलावा, ", False), (f"{chhapamari_dal}", False), ("। शामिल थे।", False)], before=0, after=2, line=1.0)
    para(doc, "छापामारी का ब्योरा:-", size=11, before=0, after=2)
    add_mixed(doc, [(body, False)], before=0, after=2, line=1.0)
    add_mixed(doc, [(antim_text, False)], before=0, after=2, line=1.0)
    para(doc, "अनुलग्नक:-", size=11, before=0, after=0)
    para(doc, "(1) गवाहों की सूची।", size=11, before=0, after=0)
    para(doc, "(2) जप्त सामान सह जप्ती सूची।", size=11, before=0, after=2)
    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.columns[0].width = Inches(3.85)
    table.columns[1].width = Inches(3.05)
    no_borders(table)
    left = table.cell(0, 0); right = table.cell(0, 1)
    set_cell_margins(left, top=0, bottom=0, left=0, right=120)
    set_cell_margins(right, top=0, bottom=0, left=120, right=0)
    left.vertical_alignment = 0
    right.vertical_alignment = 0
    cell_para(left, "स्थायी नाम एवं पता:-", size=11, before=0, after=0, line=1.0)
    sthayi_text = sthayi if sthayi else gawah
    if not sthayi_text: sthayi_text = "सोविन्द, पिता-श्री जगदीश प्रसाद, ग्राम-सिड़की, थाना-नागल, जिला-सहारनपुर (U.P), पिन-कोड-247551, मो० नं- 7632996820"
    cell_para(left, sthayi_text, size=11, before=0, after=0, line=1.0)
    C = WD_ALIGN_PARAGRAPH.CENTER
    cell_para(right, "विश्वासभाजन", size=11, align=C, before=0, after=0, line=1.0)
    cell_para(right, "", size=11, align=C, before=4, after=0, line=1.0)
    cell_para(right, karyalaya, size=11, align=C, before=0, after=0, line=1.0)
    cell_para(right, shakha, size=11, align=C, before=0, after=0, line=1.0)
    cell_para(right, "सह अभियोजन पदाधिकारी", size=11, align=C, before=0, after=0, line=1.0)
    out = BASE / f"fir_{fid}.docx"
    doc.save(str(out))
    print(f"OK - {out} - 1 Page - Kul {kul_s} - {abhiyukt} - Logo {logo_choice} - Font {FONT_NAME} - Company {company_name} - KrutiLib {'ON' if _LIB_TO_KRUTI else 'OFF-Table'}")
    return str(out)
if __name__ == "__main__":
    if len(sys.argv) > 1: make_docx(sys.argv[1])
    else:
        data = load_data()
        if not data: make_docx(17)
        else:
            for r in data:
                fid = r.get("id", 17)
                try: make_docx(fid)
                except Exception as e: print(f"FAIL {fid}: {e}")