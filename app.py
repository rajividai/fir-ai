import json
import os, pathlib, subprocess, sys, urllib.parse, os
from http.server import BaseHTTPRequestHandler, HTTPServer
BASE = pathlib.Path(__file__).parent
DATA = BASE / "data.json"
def hindi_words(n):
    try: n = int(float(n))
    except: return ""
    ones = ["","एक","दो","तीन","चार","पाँच","छह","सात","आठ","नौ","दस","ग्यारह","बारह","तेरह","चौदह","पंद्रह","सोलह","सत्रह","अठारह","उन्नीस","बीस","इक्कीस","बाईस","तेईस","चौबीस","पच्चीस","छब्बीस","सत्ताईस","अट्ठाईस","उनतीस","तीस","इकतीस","बत्तीस","तैंतीस","चौंतीस","पैंतीस","छत्तीस","सैंतीस","अड़तीस","उनतालीस","चालीस","इकतालीस","बयालीस","तैंतालीस","चवालीस","पैंतालीस","छियालीस","सैंतालीस","अड़तालीस","उनचास","पचास","इक्यावन","बावन","तिरपन","चौवन","पचपन","छप्पन","सत्तावन","अट्ठावन","उनसठ","साठ","इकसठ","बासठ","तिरसठ","चौंसठ","पैंसठ","छियासठ","सड़सठ","अड़सठ","उनहत्तर","सत्तर","इकहत्तर","बहत्तर","तिहत्तर","चौहत्तर","पचहत्तर","छिहत्तर","सतहत्तर","अठहत्तर","उनासी","अस्सी","इक्यासी","बयासी","तिरासी","चौरासी","पचासी","छियासी","सत्तासी","अट्ठासी","नवासी","नब्बे","इक्यानवे","बानवे","तिरानवे","चौरानवे","पचानवे","छियानवे","सत्तानवे","अट्ठानवे","निन्यानवे"]
    if n < 100: return ones[n]
    if n < 1000:
        s = ones[n//100] + " सौ"
        if n%100: s += " " + ones[n%100]
        return s
    if n < 100000:
        s = hindi_words(n//1000) + " हजार"
        if n%1000: s += " " + hindi_words(n%1000)
        return s
    return str(n)
class H(BaseHTTPRequestHandler):
    def do_GET(self):
        if self.path == "/" or self.path == "/index.html":
            html = (BASE / "index.html").read_text(encoding="utf-8")
            self.send_response(200); self.send_header("Content-Type","text/html; charset=utf-8"); self.end_headers()
            self.wfile.write(html.encode()); return
        if self.path == "/list":
            data = json.loads(DATA.read_text(encoding="utf-8")) if DATA.exists() else []
            out = [{"id":d.get("id"), "abhiyukt":d.get("abhiyukt"), "theft_type":d.get("theft_type"), "thana":d.get("thana")} for d in data]
            self.send_response(200); self.send_header("Content-Type","application/json; charset=utf-8"); self.end_headers()
            self.wfile.write(json.dumps(out, ensure_ascii=False).encode()); return
        if self.path.startswith("/download"):
            q = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
            fid = q.get("id",[""])[0]
            p1 = BASE / f"fir_{fid}.docx"
            files = [p1] if p1.exists() else list(BASE.glob(f"fir_{fid}_*.docx"))
            if not files: files = list(BASE.glob(f"fir_{fid}*.docx"))
            if files:
                data_bytes = files[0].read_bytes()
                self.send_response(200)
                self.send_header("Content-Type","application/vnd.openxmlformats-officed.wordprocessingml.document")
                self.send_header("Content-Disposition", f"attachment; filename=fir_{fid}.docx")
                self.send_header("Content-Length", str(len(data_bytes)))
                self.end_headers(); self.wfile.write(data_bytes); return
            self.send_response(404); self.send_header("Content-Type","text/plain; charset=utf-8"); self.end_headers()
            self.wfile.write(f"fir_{fid} wali .docx nahi mili".encode()); return
        self.send_response(404); self.end_headers()
    def do_POST(self):
        if self.path != "/generate":
            self.send_response(404); self.end_headers(); return
        ln = int(self.headers.get("Content-Length",0))
        body = json.loads(self.rfile.read(ln).decode())
        logo_choice = body.get("logo_choice","sbpdcl")
        font_choice = body.get("font_choice","noto")
        karyalaya = body.get("karyalaya","कनीय विद्युत अभियंता")
        shakha = body.get("shakha","विद्युत आपूर्ति प्रशाखा, बाँकेबाजार")
        zila = body.get("zila","गया जी")
        thana = body.get("thana","लुधुआ")
        patrank = body.get("patrank","/बाँकेबाजार")
        dinank = body.get("dinank","17/08/2026")
        email = body.get("email","")
        mob = body.get("mob","")
        hastakshar = body.get("hastakshar","")
        theft_type = str(body.get("theft_type","4"))
        abhiyukt = body.get("abhiyukt","")
        pita_type = body.get("pita_type","पिता")
        pita = body.get("pita","")
        gaon = body.get("gaon","")
        panchayat = body.get("panchayat","")
        upbhokta = body.get("upbhokta","")
        meter = body.get("meter","")
        pata = body.get("pata", f"ग्राम-{gaon}, पंचायत-{panchayat}, थाना-{thana}, जिला-{zila} (बिहार)")
        chhapamari_dinank = body.get("chhapamari_dinank", dinank)
        samay = body.get("samay","")
        dal = body.get("dal","")
        bhar = body.get("bhar","")
        shreni = body.get("shreni","DS1D")
        vichchhedit = body.get("vichchhedit","")
        jabt = body.get("jabt","PVC तार के आंशिक भाग को काट कर जब्त")
        sthaniya = body.get("sthaniya","")
        compounding = body.get("compounding","")
        antim_text = body.get("antim","")
        n = float(body.get("nirdharan",0) or 0)
        b = float(body.get("bakaya",0) or 0)
        kul = n + b
        if theft_type == "1":
            byora_core = f"{abhiyukt}, {pita_type}-{pita}, {pata} के परिसर में बिना किसी वैध विद्युत संबंध के ही मुख्य LT लाइन में लाल एवं काले तार से टोका फँसाकर अवैध रूप से विद्युत ऊर्जा की चोरी की जा रही थी। जाँच के दौरान परिसर का कुल विद्युत भार {bhar} ({shreni} श्रेणी) में पाया गया।"
        elif theft_type == "2":
            byora_core = f"{abhiyukt}, {pita_type}-{pita}, (उपभोक्ता संख्या-{upbhokta}), मीटर नं-{meter}, {pata} के परिसर में मीटर को बायपास करने हेतु सर्विस वायर को काटकर एक्स्ट्रा वायर जोड़कर अवैध रूप से विद्युत ऊर्जा की चोरी की जा रही थी। जाँच के दौरान परिसर का कुल विद्युत भार {bhar} ({shreni} श्रेणी) में पाया गया।"
        elif theft_type == "3":
            byora_core = f"{abhiyukt}, {pita_type}-{pita}, (उपभोक्ता संख्या-{upbhokta}), मीटर नं-{meter}, {pata} के परिसर में सर्विस वायर के अलावा अलग तार से मीटर को बायपास कर (आटा चक्की हेतु) अवैध रूप से विद्युत ऊर्जा की चोरी की जा रही थी। जाँच के दौरान परिसर का कुल विद्युत भार {bhar} ({shreni} श्रेणी) में पाया गया।"
        else:
            byora_core = f"{abhiyukt}, {pita_type}-{pita}, (उपभोक्ता संख्या-{upbhokta}), मीटर नं-{meter}, {pata} के परिसर में विद्युत विपत्र बकाया होने के कारण दिनांक {vichchhedit} को ही विद्युत संबंध विच्छेदित कर दिये जाने के बाद भी इनके द्वारा अवैध रूप से घरेलु परिसर में विद्युत ऊर्जा का उपयोग किया जा रहा था। जाँच के दौरान परिसर का कुल विद्युत भार {bhar} ({shreni} श्रेणी) में पाया गया।"
        byora = f"छापामारी दल समय लगभग {samay} में {byora_core} विद्युत ऊर्जा चोरी में प्रयुक्त {jabt} को काट कर जब्त कर लिया गया है, जिसे जाँच-सह-जब्ती प्रतिवेदन के साथ थाने में समर्पित की जा रही है। इनके द्वारा इस प्रकार से अवैध रूप से विद्युत ऊर्जा की चोरी करने से साउथ बिहार पॉवर डिस्ट्रीब्यूशन कम्पनी लिमिटेड को कुल राशि-{int(n)}/- रुपया ({hindi_words(n)} रुपया) मात्र की राजस्व हानि अनुमानित है, एवं पूर्व में इनके विद्युत विपत्र का बकाया राशि {b:.2f} रुपया को मिलाकर कुल {int(n)}+{int(b)}={kul:.2f} रुपया ({hindi_words(kul)} रुपया) है। जिसकी वसूली इनसे ही की जानी है। उक्त राशि में भारतीय विद्युत अधिनियम 2003 की धारा 152 की तहत कम्पाउंडिंग की राशि सम्मिलित नहीं है।"
        if compounding: byora += f" कम्पाउंडिंग राशि {compounding} रुपया अलग से है।"
        dal_full = f"उपर्युक्त विषय के संदर्भ में सूचित करना है कि दिनांक-{chhapamari_dinank} को विद्युत ऊर्जा चोरी के रोकथाम हेतु एक छापामारी दल का गठन किया गया, जिसमें मेरे अलावा, {dal} शामिल थे।"
        if not antim_text: antim_text = f"अतः श्रीमान से आग्रह है कि संबंधित उपरोक्त अभियुक्त {abhiyukt}, {pita_type}-{pita}, पर भारतीय विद्युत अधिनियम 2003 की धारा 135 एवं अन्य सुसंगत धाराओं के अंतर्गत प्राथमिकी दर्ज किया जाए ताकि इस प्रकार की घटना की पुनरावृत्ति न हो सके।"
        if not hastakshar: hastakshar = f"विश्वासभाजन\n{karyalaya}, {shakha}, सह अभियोजन पदाधिकारी"
        data = json.loads(DATA.read_text(encoding="utf-8")) if DATA.exists() else []
        new_id = (max([d.get("id",0) for d in data], default=0)+1) if data else 1
        entry = {"id": new_id, "logo_choice": logo_choice, "font_choice": font_choice, "theft_type": theft_type, "karyalaya": f"{karyalaya}\n{shakha}", "shakha": shakha, "zila": zila, "thana": thana, "patrank": patrank, "abhiyukt": abhiyukt, "pita": pita, "pita_type": pita_type, "gaon": gaon, "panchayat": panchayat, "upbhokta": upbhokta, "meter": meter, "pata": pata, "email": email, "mob": mob, "dinank": dinank, "chhapamari_dinank": chhapamari_dinank, "samay": samay, "dal": dal, "bhar": bhar, "shreni": shreni, "vichchhedit": vichchhedit, "jabt": jabt, "nirdharan": n, "bakaya": b, "kul": kul, "compounding": compounding, "chhapamari_dal_full": dal_full, "byora_full": byora, "antim": antim_text, "hastakshar": hastakshar, "sthaniya_naam": sthaniya}
        data.append(entry)
        DATA.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        r = subprocess.run([sys.executable, str(BASE/"make_docx_ditto.py"), str(new_id)], cwd=str(BASE), capture_output=True, text=True)
        print("STDOUT:", r.stdout)
        print("STDERR:", r.stderr)
        self.send_response(200); self.send_header("Content-Type","application/json; charset=utf-8"); self.end_headers()
        self.wfile.write(json.dumps({"ok": r.returncode==0, "id": new_id, "kul": kul, "theft_type": theft_type, "msg": r.stdout[-300:], "err": r.stderr[-300:]}, ensure_ascii=False).encode())
    def log_message(self, *a): pass
PORT = int(os.environ.get("PORT", 8000))
print(f"FIR App chal raha hai - http://localhost:{PORT}")
print("Folder:", str(BASE))
HTTPServer(("0.0.0.0", PORT), H).serve_forever()