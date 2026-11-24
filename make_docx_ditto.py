# make_docx_ditto_TABLE2.py - Vishwasbhajan CENTER of Right Box + Right Box Aligned
# Rename karke use karo: cp make_docx_ditto_TABLE2.py make_docx_ditto.py
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
def set_font(run, size=11, bold=False):
    run.font.name = "Nirmala UI"; run.font.size = Pt(size); run.bold = bold; run.font.color.rgb = RGBColor(0,0,0)
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts'); rPr.append(rFonts)
    for a in ('w:ascii','w:hAnsi','w:cs','w:eastAsia'): rFonts.set(qn(a), 'Nirmala UI')
def para(doc, text="", size=11, bold=False, align=None, before=0, after=1, line=1.0, keep_next=False, right_indent=None):
    p = doc.add_paragraph(); pf = p.paragraph_format
    pf.space_before = Pt(before); pf.space_after = Pt(after); pf.line_spacing = line
    pf.widow_control = True; pf.keep_with_next = keep_next
    if align is not None: p.alignment = align
    if right_indent is not None: pf.right_indent = Inches(right_indent)
    if text:
        r = p.add_run(text); set_font(r, size=size, bold=bold)
    return p
def add_mixed(doc, parts, align=None, before=0, after=1, line=1.0, keep_next=False, right_indent=None):
    p = doc.add_paragraph(); pf = p.paragraph_format
    pf.space_before = Pt(before); pf.space_after = Pt(after); pf.line_spacing = line
    pf.widow_control = True; pf.keep_with_next = keep_next
    if align is not None: p.alignment = align
    if right_indent is not None: pf.right_indent = Inches(right_indent)
    for txt, b in parts:
        r = p.add_run(txt); set_font(r, size=11, bold=b)
    return p
def cell_para(cell, text="", size=11, align=None, before=0, after=0, line=1.0):
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
    return {"id": 17, "patrank": "/बाँकेबाजार", "dinank": "17/08/2026", "email": "jeebankeybazar@gmail.com", "mob": "7632996820", "thana": "लुथुआ", "jila": "गया जी (बिहार)", "abhiyukt": "कमरूद्दीन मियाँ", "pita": "नजामुद्दीन मियाँ", "gram": "लुथुआ", "panchayat": "लुथुआ", "upbhokta": "229200095413", "meter": "188510", "load": "0.503", "shreni": "DS1D", "chhapamari_time": "12:35 बजे अपराह्न", "chhapamari_dal": "श्री राहुल कुमार वर्मा मानवबल बाँकेबाजार, श्री अमित कुमार, तकनीशियन कोटि 03, बाँकेबाजार, कमरूद्दीन बिलिंग सुपरवाइजर बाँकेबाजार एवं श्री लोकेश कुमार मानवबल बाँकेबाजार, राजीव झा, सहायक विद्युत अभियंता इमामगंज", "disconnection_date": "14/07/2023", "nirdharan": "52945", "bakaya": "20575", "japt": "PVC तार के आंशिक भाग को काट कर जब्त", "gawah": "सोविन्द, पिता-श्री जगदीश प्रसाद, ग्राम-सिड़की, थाना-नागल, जिला-सहारनपुर (U.P), पिन-कोड-247551, मो० नं- 7632996820", "sthayi": "सोविन्द, पिता-श्री जगदीश प्रसाद, ग्राम-सिड़की, थाना-नागल, जिला-सहारनपुर (U.P), पिन-कोड-247551, मो० नं- 7632996820"}
def make_docx(fid):
    rec = find_record(fid)
    if rec is None:
        print(f"ID {fid} nahi mila - Default 17 se bana raha hun"); rec = default_record(); rec["id"] = fid
    cfg = load_config()
    for k in ["office","shakha","email","mob","thana","jila"]:
        if not rec.get(k) and cfg.get(k): rec[k] = cfg[k]
    nirdharan = rec.get("nirdharan") or rec.get("Nirdharan Rashi") or "52945"
    bakaya = rec.get("bakaya") or rec.get("Bakaya Rashi") or "20575"
    try: kul = float(str(nirdharan).replace(",","")) + float(str(bakaya).replace(",",""))
    except: kul = 73520.0
    kul_s = f"{kul:.2f}"
    nirdharan_words = num_to_hindi(nirdharan) or "बावन हजार नौ सौ पैंतालीस"
    kul_words = num_to_hindi(kul) or "तिहत्तर हजार पाँच सौ बीस"
    abhiyukt = rec.get("abhiyukt") or "कमरूद्दीन मियाँ"; pita = rec.get("pita") or "नजामुद्दीन मियाँ"
    gram = rec.get("gram") or "लुथुआ"; panchayat = rec.get("panchayat") or gram
    thana = rec.get("thana") or "लुथुआ"; jila = rec.get("jila") or "गया जी (बिहार)"
    upbhokta = rec.get("upbhokta") or "229200095413"; meter = rec.get("meter") or "188510"
    load = rec.get("load") or "0.503"; shreni = rec.get("shreni") or "DS1D"
    dinank = rec.get("dinank") or "17/08/2026"; patrank = rec.get("patrank") or "/बाँकेबाजार"
    chhapamari_time = rec.get("chhapamari_time") or "12:35 बजे अपराह्न"
    chhapamari_dal = rec.get("chhapamari_dal") or "श्री राहुल कुमार वर्मा मानवबल बाँकेबाजार, श्री अमित कुमार, तकनीशियन कोटि 03, बाँकेबाजार, कमरूद्दीन बिलिंग सुपरवाइजर बाँकेबाजार एवं श्री लोकेश कुमार मानवबल बाँकेबाजार, राजीव झा, सहायक विद्युत अभियंता इमामगंज"
    disconnection = rec.get("disconnection_date") or "14/07/2023"
    japt = rec.get("japt") or "PVC तार के आंशिक भाग को काट कर जब्त"
    gawah = rec.get("gawah") or rec.get("sthayi") or ""; sthayi = rec.get("sthayi") or gawah
    email = rec.get("email") or "jeebankeybazar@gmail.com"; mob = rec.get("mob") or "7632996820"
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Inches(8.27); sec.page_height = Inches(11.69)
    sec.top_margin = Inches(0.45); sec.bottom_margin = Inches(0.35)
    sec.left_margin = Inches(0.70); sec.right_margin = Inches(0.70)
    sec.header_distance = Inches(0.25); sec.footer_distance = Inches(0.25)
    style = doc.styles['Normal']; style.font.name = 'Nirmala UI'; style.font.size = Pt(11)
    style.paragraph_format.space_before = Pt(0); style.paragraph_format.space_after = Pt(1)
    style.paragraph_format.line_spacing = 1.0; style.paragraph_format.widow_control = True
    para(doc, "-::कार्यालय::-", size=11, align=WD_ALIGN_PARAGRAPH.CENTER, before=0, after=0)
    para(doc, "कनीय विद्युत अभियंता", size=11, align=WD_ALIGN_PARAGRAPH.CENTER, before=0, after=0)
    para(doc, "विद्युत आपूर्ति प्रशाखा, बाँकेबाजार", size=11, align=WD_ALIGN_PARAGRAPH.CENTER, before=0, after=0)
    para(doc, f"E-mail: {email} Mob No:- {mob}", size=11, align=WD_ALIGN_PARAGRAPH.CENTER, before=0, after=1)
    para(doc, "="*72, size=11, align=WD_ALIGN_PARAGRAPH.CENTER, before=0, after=2)
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
    body = (f"छापामारी दल समय लगभग {chhapamari_time} में {abhiyukt}, पिता-{pita} अजमेरी, (उपभोक्ता संख्या-{upbhokta}), " f"मीटर नं-{meter}, ग्राम-{gram}, पंचायत-{panchayat}, थाना-{thana}, जिला-{jila} के परिसर में विद्युत विपत्र बकाया होने के कारण " f"दिनांक {disconnection} को ही विद्युत संबंध विच्छेदित कर दिये जाने के बाद भी इनके द्वारा अवैध रूप से घरेलु परिसर में विद्युत ऊर्जा का उपयोग किया जा रहा था। " f"जाँच के दौरान परिसर का कुल विद्युत भार {load} KW ({shreni} श्रेणी) में पाया गया। " f"विद्युत ऊर्जा चोरी में प्रयुक्त {japt} कर लिया गया है, जिसे जाँच-सह-जब्ती प्रतिवेदन के साथ थाने में समर्पित की जा रही है। " f"इनके द्वारा इस प्रकार से अवैध रूप से विद्युत ऊर्जा की चोरी करने से साउथ बिहार पॉवर डिस्ट्रीब्यूशन कम्पनी लिमिटेड को कुल राशि-{fmt_amt(nirdharan)}/- रुपया ({nirdharan_words} रुपया) मात्र की राजस्व हानि अनुमानित है, " f"एवं पूर्व में इनके विद्युत विपत्र का बकाया राशि {fmt_amt2(bakaya)} रुपया को मिलाकर कुल {fmt_amt(nirdharan)}+{fmt_amt2(bakaya)}={kul_s} रुपया ({kul_words} रुपया) है। " f"जिसकी वसूली इनसे ही की जानी है। उक्त राशि में भारतीय विद्युत अधिनियम 2003 की धारा 152 की तहत कम्पाउंडिंग की राशि सम्मिलित नहीं है।")
    add_mixed(doc, [(body, False)], before=0, after=2, line=1.0)
    add_mixed(doc, [("अतः श्रीमान से आग्रह है कि संबंधित उपरोक्त अभियुक्त पर भारतीय विद्युत अधिनियम 2003 की धारा 135 एवं अन्य सुसंगत धाराओं के अंतर्गत प्राथमिकी दर्ज किया जाए ताकि इस प्रकार की घटना की पुनरावृत्ति न हो सके।", False)], before=0, after=2, line=1.0)
    para(doc, "अनुलग्नक:-", size=11, before=0, after=0)
    para(doc, "(1) गवाहों की सूची।", size=11, before=0, after=0)
    para(doc, "(2) जप्त सामान सह जप्ती सूची।", size=11, before=0, after=2)
    # BOTTOM TABLE - Left Sthayi | Right Signature - SAME LEVEL
    # FIX: Right Box ka saara text CENTER - Vishwasbhajan Right Box ke MIDDLE me aayega
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
    # RIGHT BOX - SAB CENTER - Aligned
    C = WD_ALIGN_PARAGRAPH.CENTER
    cell_para(right, "विश्वासभाजन", size=11, align=C, before=0, after=0, line=1.0)
    cell_para(right, "", size=11, align=C, before=4, after=0, line=1.0)
    cell_para(right, "कनीय विद्युत अभियंता", size=11, align=C, before=0, after=0, line=1.0)
    cell_para(right, "विद्युत आपूर्ति प्रशाखा, बाँकेबाजार", size=11, align=C, before=0, after=0, line=1.0)
    cell_para(right, "सह अभियोजन पदाधिकारी", size=11, align=C, before=0, after=0, line=1.0)
    out = BASE / f"fir_{fid}.docx"
    doc.save(str(out))
    print(f"OK - {out} - 1 Page - Kul {kul_s} - {abhiyukt}")
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
