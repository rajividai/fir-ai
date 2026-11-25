# make_docx_ditto.py - FIXED + LOGO + FONT + KRUTI V2 + COMPANY HEADING TOP + STATIC LABELS KRUTI
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
KRUTI_CONS = {
 "क":"d","ख":"[k","ग":"x","घ":"?k","ङ":"u","च":"p","छ":"N","ज":"t","झ":"]","ञ":"U",
 "ट":"V","ठ":"B","ड":"M","ढ":"<","ण":".k","त":"r","थ":"Fk","द":"n","ध":"/k","न":"u",
 "प":"i","फ":"Q","ब":"c","भ":"H","म":"e","य":";","र":"j","ल":"y","व":"o",
 "श":"'k","ष":"\"k","स":"l","ह":"g","क्ष":"{k","त्र":"=","ज्ञ":"K",
}
KRUTI_VOWEL = {"अ":"v","आ":"vk","इ":"b","ई":"bZ","उ":"m","ऊ":"mZ","ऋ":"`","ए":"vks","ऐ":"vkS","ओ":"vksS","औ":"vkS"}
KRUTI_MATRA = {"ा":"k","ी":"h","ु":"q","ू":"w","ृ":"`","े":"s","ै":"S","ो":"ks","ौ":"kS","ं":"a","ः":"%","ँ":"a","़":".","।":"|"}
def to_kruti(text):
    if not text: return text
    s = str(text)
    # Agar pehle se Kruti hai (ASCII hi hai, koi Devanagari nahi) to waisa hi wapas - Double Convert se bachao
    has_dev = any('\u0900' <= ch <= '\u097f' for ch in s)
    if not has_dev: return s
    out = ""
    i = 0
    while i < len(s):
        ch = s[i]
        if ch == "ि":
            out += "f"; i += 1; continue
        if ch == "्":
            out += "~"; i += 1; continue
        if s[i:i+2] in ["क्ष","त्र","ज्ञ"]:
            out += KRUTI_CONS[s[i:i+2]]; i += 2; continue
        if ch in KRUTI_CONS:
            if i+1 < len(s) and s[i+1] == "ि":
                out += "f" + KRUTI_CONS[ch]; i += 2; continue
            out += KRUTI_CONS[ch]; i += 1; continue
        if ch in KRUTI_VOWEL:
            out += KRUTI_VOWEL[ch]; i += 1; continue
        if ch in KRUTI_MATRA:
            out += KRUTI_MATRA[ch]; i += 1; continue
        out += ch; i += 1
    out = out.replace("df", "fd").replace("f[", "[f")
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
    for b in ['top','left','bottom','right','inside