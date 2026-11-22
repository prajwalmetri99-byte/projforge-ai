import io
import json
import os
import socket
import threading
import time
import webbrowser
import zipfile
from http.server import BaseHTTPRequestHandler, HTTPServer
from socketserver import ThreadingMixIn
from xml.sax.saxutils import escape

DB_FILE = "users_db.json"
LAUNCH_FILE = "launch_config.json"
FREE_DAYS = 2
LOCK = threading.Lock()


def get_launch_end_ts():
    if os.path.exists(LAUNCH_FILE):
        try:
            with open(LAUNCH_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if "end_ts" in data:
                    return float(data["end_ts"])
        except Exception:
            pass
    end_ts = time.time() + (FREE_DAYS * 86400)
    with open(LAUNCH_FILE, "w", encoding="utf-8") as f:
        json.dump({"end_ts": end_ts}, f, indent=2)
    return end_ts


LAUNCH_END_TS = get_launch_end_ts()


def is_free_period():
    return time.time() < LAUNCH_END_TS


def load_users():
    if os.path.exists(DB_FILE):
        try:
            with open(DB_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_users(data):
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)


def generate_project_bundle(mode, title, branch, budget, extra_notes):
    budget = max(200, int(budget or 600))
    t_clean = title.strip() or "Smart Automated IoT System"
    low = t_clean.lower()

    is_software = any(k in low for k in ("python", "ai", "ml", "app", "web", "attendance", "detection", "data", "software"))
    is_network = any(k in low for k in ("rc", "rlc", "filter", "network", "circuit", "amplifier", "rectifier", "aelic", "voltage"))

    if is_software:
        comps = [
            {"name": "PC / Laptop with Python 3.x & VS Code", "cost": 0},
            {"name": "Open-Source Python Libraries (CSV, Math, JSON)", "cost": 0},
            {"name": "Report & Documentation Setup", "cost": min(budget, 100)},
        ]
        sim_type = "software"
        code = (
            f"# Ready-to-Run Python Implementation: {t_clean}\n"
            f"# Branch: {branch} | Budget: Rs.{budget}\n"
            "import json\n\n"
            "def run_system(values):\n"
            f"    print('Running {t_clean}...')\n"
            "    avg = round(sum(values) / max(1, len(values)), 2)\n"
            "    status = 'ALERT TRIGGERED' if avg > 50 else 'NORMAL'\n"
            "    return {'project': '" + t_clean.replace("'", "") + "', 'reading': avg, 'status': status}\n\n"
            "if __name__ == '__main__':\n"
            "    print(json.dumps(run_system([45, 72, 68, 55]), indent=2))\n"
        )
    elif is_network:
        r_cost = min(150, int(budget * 0.25))
        c_cost = min(250, int(budget * 0.45))
        comps = [
            {"name": "Precision Resistor Kit (1kΩ – 10kΩ)", "cost": r_cost},
            {"name": "Electrolytic / Ceramic Capacitor Kit (10µF)", "cost": c_cost},
            {"name": "Breadboard, Jumper Wires & Probes", "cost": max(50, budget - r_cost - c_cost)},
        ]
        sim_type = "network"
        code = (
            f"# Transient & Frequency Response Analyzer: {t_clean}\n"
            "import math\n\n"
            "R, C, Vin = 10000.0, 10e-6, 5.0\n"
            "tau = R * C\n"
            "fc = 1.0 / (2.0 * math.pi * R * C)\n"
            "print(f'Tau = {tau:.3f}s | Cutoff Freq = {fc:.2f} Hz')\n"
            "for ms in range(0, 500, 100):\n"
            "    t = ms / 1000.0\n"
            "    print(f't={t:.2f}s -> Vout = {Vin * (1 - math.exp(-t / tau)):.3f} V')\n"
        )
    else:
        mcu = ("Arduino Nano + USB Cable", 240) if budget <= 650 else ("Arduino Uno R3 + USB Cable", 420)
        rem = max(120, budget - mcu[1])
        s1 = int(rem * 0.55)
        comps = [
            {"name": mcu[0], "cost": mcu[1]},
            {"name": f"Sensor Module for {t_clean}", "cost": s1},
            {"name": "Breadboard, LEDs, 220Ω Resistors & Wires", "cost": rem - s1},
        ]
        sim_type = "arduino"
        code = (
            f"// Ready-to-Dump Arduino Code: {t_clean}\n"
            f"// Budget Target: Rs.{budget}\n"
            "const int sensorPin = A0;\n"
            "const int triggerPin = 2;\n"
            "const int outPin = 9;\n\n"
            "void setup() {\n"
            "  pinMode(triggerPin, INPUT);\n"
            "  pinMode(outPin, OUTPUT);\n"
            "  Serial.begin(9600);\n"
            "}\n\n"
            "void loop() {\n"
            "  int val = analogRead(sensorPin);\n"
            "  if (val < 450 || digitalRead(triggerPin) == LOW) {\n"
            "    digitalWrite(outPin, HIGH);\n"
            "  } else {\n"
            "    digitalWrite(outPin, LOW);\n"
            "  }\n"
            "  delay(150);\n"
            "}\n"
        )

    total_cost = sum(c["cost"] for c in comps)
    spec = f" Custom Specifications: {extra_notes}." if extra_notes.strip() else ""

    return {
        "title": t_clean,
        "branch": branch,
        "budget": budget,
        "total_cost": total_cost,
        "sim_type": sim_type,
        "components": comps,
        "viva": {
            "kn": f"'{t_clean}' ಪ್ರಾಜೆಕ್ಟ್ ಅನ್ನು ₹{budget} ಬಜೆಟ್ ಒಳಗೆ ವಿನ್ಯಾಸಗೊಳಿಸಲಾಗಿದೆ. ಸೆನ್ಸಾರ್/ಇನ್‌ಪುಟ್ ನಿಂದ ಡೇಟಾ ಪಡೆದು ಆಟೋಮ್ಯಾಟಿಕ್ ಆಗಿ ಔಟ್‌ಪುಟ್ ನಿಯಂತ್ರಿಸುತ್ತದೆ.{spec}",
            "en": f"'{t_clean}' is engineered within a ₹{budget} budget. It reads input signals, checks threshold conditions, and drives outputs automatically.{spec}",
            "hi": f"'{t_clean}' प्रोजेक्ट को ₹{budget} के बजट में तैयार किया गया है। यह इनपुट सिग्नल प्रोसेस करके आउटपुट को नियंत्रित करता है।"
        },
        "intro": f"This project presents the design and implementation of '{t_clean}' in the {branch} domain, optimized within a student budget of Rs.{budget}.{spec}",
        "working": f"The '{t_clean}' system captures input signals, processes threshold logic via the controller, and triggers the output stage with minimal power consumption.",
        "conclusion": f"The '{t_clean}' prototype was successfully simulated and verified within Rs.{total_cost}.",
        "code": code
    }


def build_docx_bytes(proj, cfg):
    font = cfg.get("font", "Times New Roman")
    h_size = int(cfg.get("h_size", 14)) * 2
    b_size = int(cfg.get("b_size", 12)) * 2
    h_bold = bool(cfg.get("h_bold", True))
    line_sp = "360" if cfg.get("spacing") == "1.5" else "276"
    order = cfg.get("order", ["intro", "components", "working", "code", "conclusion"])

    def p_xml(text, sz, bold=False, center=False, is_code=False):
        f_name = "Consolas" if is_code else font
        b_tag = "<w:b/>" if bold else ""
        jc_tag = '<w:jc w:val="center"/>' if center else '<w:jc w:val="both"/>'
        runs = []
        for idx, line in enumerate(str(text).split("\n")):
            br = "<w:br/>" if idx > 0 else ""
            runs.append(
                f'{br}<w:r><w:rPr><w:rFonts w:ascii="{escape(f_name)}" w:hAnsi="{escape(f_name)}"/>'
                f'{b_tag}<w:sz w:val="{sz}"/></w:rPr><w:t xml:space="preserve">{escape(line)}</w:t></w:r>'
            )
        return f'<w:p><w:pPr>{jc_tag}<w:spacing w:after="140" w:line="{line_sp}" w:lineRule="auto"/></w:pPr>{"".join(runs)}</w:p>'

    body = [
        p_xml(cfg.get("college", "ENGINEERING COLLEGE").upper(), h_size + 4, bold=True, center=True),
        p_xml(f"Project Report — {proj['branch']}", b_size, bold=True, center=True),
        p_xml(proj["title"].upper(), h_size + 2, bold=True, center=True),
        p_xml(f"By: {cfg.get('student', 'Student')} (USN: {cfg.get('usn', '1AM23XX000')}) | Guide: {cfg.get('guide', 'Prof. Faculty')}", b_size, bold=True, center=True),
    ]

    for idx, sec in enumerate(order, start=1):
        if sec == "intro":
            body.extend([p_xml(f"{idx}. INTRODUCTION", h_size, bold=h_bold), p_xml(proj["intro"], b_size)])
        elif sec == "components":
            body.extend([
                p_xml(f"{idx}. COMPONENTS & BUDGET (TOTAL: Rs.{proj['total_cost']})", h_size, bold=h_bold),
                p_xml("\n".join(f"• {c['name']} — Rs.{c['cost']}" for c in proj["components"]), b_size)
            ])
        elif sec == "working":
            body.extend([p_xml(f"{idx}. CIRCUIT & WORKING PRINCIPLE", h_size, bold=h_bold), p_xml(proj["working"], b_size)])
        elif sec == "code":
            body.extend([p_xml(f"{idx}. SOURCE CODE", h_size, bold=h_bold), p_xml(proj["code"], max(18, b_size - 2), is_code=True)])
        elif sec == "conclusion":
            body.extend([p_xml(f"{idx}. CONCLUSION", h_size, bold=h_bold), p_xml(proj["conclusion"], b_size)])

    doc_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + "".join(body) + "</w:body></w:document>"
    ct_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
    rels_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", ct_xml)
        zf.writestr("_rels/.rels", rels_xml)
        zf.writestr("word/document.xml", doc_xml)
    return buf.getvalue()


MANIFEST_JSON = json.dumps({
    "name": "ProjForge AI Studio",
    "short_name": "ProjForge",
    "start_url": "/",
    "display": "standalone",
    "background_color": "#0b1120",
    "theme_color": "#0284c7",
    "icons": [{
        "src": "data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 192 192'%3E%3Crect width='192' height='192' rx='40' fill='%230284c7'/%3E%3Ctext x='96' y='125' font-size='95' text-anchor='middle' fill='white'%3E⚡%3C/text%3E%3C/svg%3E",
        "sizes": "192x192",
        "type": "image/svg+xml"
    }]
})

SW_JS = "self.addEventListener('fetch', function(e) {});"

ROBOTS_TXT = "User-agent: *\nAllow: /\nSitemap: https://projforge-ai.onrender.com/sitemap.xml\n"

SITEMAP_XML = """<?xml version="1.0" encoding="UTF-8"?>
<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">
  <url>
    <loc>https://projforge-ai.onrender.com/</loc>
    <changefreq>daily</changefreq>
    <priority>1.0</priority>
  </url>
</urlset>"""

HTML_PAGE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<meta name="google-site-verification" content="SCartIuz0KkP830B_WEA5VUsTsnhn86LXdXn18LBKJQ" />
<meta name="description" content="ProjForge AI - Instant Engineering Mini Project Generator, Budget Circuit Simulation, Ready-to-Dump Arduino/Python Code & Exact MS Word (.docx) Report Maker for VTU & Engineering Students." />
<meta name="keywords" content="ProjForge AI, engineering project report generator, VTU mini project word format, Arduino code generator, PBL project ideas, Kannada project explanation" />
<meta name="robots" content="index, follow" />
<meta name="theme-color" content="#0284c7" />
<link rel="manifest" href="/manifest.json" />
<title>ProjForge AI — Engineering Project & Exact MS Word (.docx) Report Generator</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; font-family: 'Segoe UI', Arial, sans-serif; }
  body { background: #0b1120; color: #f8fafc; min-height: 100vh; }
  .banner { background: linear-gradient(90deg, #10b981, #059669); color: #fff; text-align: center; padding: 9px; font-weight: 800; font-size: 13px; }
  .banner.paid { background: linear-gradient(90deg, #f59e0b, #ef4444); color: #111; }
  .nav { background: #1e293b; padding: 12px 18px; display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #334155; flex-wrap: wrap; gap: 8px; }
  .nav h1 { font-size: 18px; color: #38bdf8; }
  .pill { background: #0f172a; border: 1px solid #38bdf8; color: #38bdf8; padding: 5px 10px; border-radius: 20px; font-size: 12px; font-weight: 700; }
  .install-btn { background: #f59e0b; color: #111; border: none; padding: 6px 12px; border-radius: 20px; font-weight: 800; font-size: 12px; cursor: pointer; display: none; }
  .wrap { max-width: 1180px; margin: 16px auto; padding: 0 12px; display: grid; grid-template-columns: 1.1fr 1fr; gap: 16px; }
  .card { background: #1e293b; border: 1px solid #334155; border-radius: 14px; padding: 16px; }
  .card h2 { font-size: 15px; color: #38bdf8; margin-bottom: 10px; display: flex; justify-content: space-between; align-items: center; }
  .tabs { display: flex; gap: 6px; margin-bottom: 10px; }
  .tab { flex: 1; padding: 8px 4px; border-radius: 8px; border: 1px solid #475569; background: #0f172a; color: #cbd5e1; font-size: 11.5px; font-weight: 700; cursor: pointer; }
  .tab.active { background: #38bdf8; color: #0f172a; border-color: #38bdf8; }
  label { display: block; font-size: 12px; color: #94a3b8; margin: 8px 0 4px; font-weight: 600; }
  input, select, textarea { width: 100%; padding: 9px 10px; border-radius: 8px; border: 1px solid #475569; background: #0f172a; color: #fff; font-size: 13px; outline: none; }
  .grid2 { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }
  .btn-ai { width: 100%; margin-top: 12px; padding: 11px; border: none; border-radius: 9px; background: linear-gradient(90deg, #0284c7, #2563eb); color: #fff; font-weight: 800; font-size: 14px; cursor: pointer; }
  .btn-pay { width: 100%; margin-top: 14px; padding: 13px; border: none; border-radius: 10px; background: linear-gradient(90deg, #10b981, #059669); color: #fff; font-size: 14.5px; font-weight: 800; cursor: pointer; }
  .sim-box { background: #020617; border: 2px solid #0ea5e9; border-radius: 10px; padding: 8px; margin-top: 10px; text-align: center; }
  .lang-row { display: flex; gap: 6px; margin: 8px 0; }
  .lbtn { padding: 5px 10px; border-radius: 6px; border: 1px solid #475569; background: #0f172a; color: #cbd5e1; cursor: pointer; font-size: 12px; font-weight: 700; }
  .lbtn.active { background: #10b981; color: #052e16; border-color: #10b981; }
  .viva-box { background: #0f172a; padding: 10px; border-radius: 8px; font-size: 13px; line-height: 1.5; border-left: 4px solid #10b981; }
  .bom { margin: 8px 0 0 18px; font-size: 12.5px; color: #e2e8f0; line-height: 1.55; }
  .code-box { background: #020617; color: #a7f3d0; padding: 10px; border-radius: 8px; font-family: monospace; font-size: 12px; white-space: pre-wrap; margin-top: 8px; max-height: 160px; overflow-y: auto; border: 1px solid #334155; }
  .modal { position: fixed; inset: 0; background: rgba(0,0,0,0.88); display: flex; justify-content: center; align-items: center; z-index: 99; padding: 12px; }
  .mcard { background: #1e293b; border: 2px solid #38bdf8; border-radius: 16px; padding: 22px; width: 370px; max-width: 100%; text-align: center; }
  .wipe-alert { background: #064e3b; border: 1px solid #10b981; color: #a7f3d0; padding: 8px 10px; border-radius: 8px; font-size: 12px; margin-top: 10px; display: none; }
  @keyframes flow { 0% { stroke-dashoffset: 24; } 100% { stroke-dashoffset: 0; } }
  .wire { stroke-dasharray: 6 6; animation: flow 0.8s linear infinite; }
  @media (max-width: 820px) { .wrap { grid-template-columns: 1fr; } }
</style>
</head>
<body>
  <div class="banner" id="topBanner">🎉 LAUNCH OFFER: 100% FREE for First 2 Days! (₹49 after 2 days) 🎉</div>
  <div class="nav">
    <h1>⚡ ProjForge AI</h1>
    <div style="display:flex;gap:6px;align-items:center;">
      <button class="install-btn" id="installBtn" onclick="installPWA()">📲 Install App</button>
      <span class="pill" id="userBadge">Login Required</span>
      <button class="lbtn" onclick="logout()">Switch</button>
    </div>
  </div>

  <!-- Login Modal -->
  <div class="modal" id="loginModal">
    <div class="mcard">
      <h2 style="color:#38bdf8;margin-bottom:6px;">Student Login</h2>
      <p style="font-size:13px;color:#94a3b8;margin-bottom:14px;">Login with Email or Mobile Number to access ProjForge AI.</p>
      <input id="loginContact" placeholder="Email or 10-digit Mobile Number" onkeydown="if(event.key==='Enter') doLogin()" />
      <button class="btn-pay" style="margin-top:12px;" onclick="doLogin()">Open Studio</button>
    </div>
  </div>

  <!-- Payment Modal -->
  <div class="modal" id="payModal" style="display:none;" onclick="if(event.target===this) this.style.display='none'">
    <div class="mcard" style="border-color:#f59e0b;">
      <div style="color:#f59e0b;font-weight:800;font-size:12px;">2-DAY FREE LAUNCH ENDED — STANDARD PLAN</div>
      <h2 style="margin:6px 0;font-size:26px;">Pay ₹49</h2>
      <p style="font-size:12.5px;color:#cbd5e1;margin-bottom:12px;">Downloads MS Word (.docx) + Dump Code & auto-wipes your project details for 100% privacy.</p>
      <div style="background:#fff;padding:8px;border-radius:10px;display:inline-block;margin-bottom:10px;">
        <img id="upiQr" src="https://api.qrserver.com/v1/create-qr-code/?size=140x140&data=upi://pay?pa=projforge@upi&pn=ProjForge&am=49&cu=INR" width="140" height="140" />
      </div>
      <button class="btn-pay" style="margin-top:4px;" onclick="executeDownloadAndWipe()">Confirm & Download .docx</button>
      <button style="margin-top:8px;background:transparent;border:none;color:#94a3b8;cursor:pointer;" onclick="document.getElementById('payModal').style.display='none'">Cancel</button>
    </div>
  </div>

  <div class="wrap">
    <div class="card">
      <h2>1. Select / Enter Project <span style="font-size:11px;color:#10b981;">ZERO-TRACE</span></h2>
      <div class="tabs">
        <button class="tab active" id="t_trend" onclick="setMode('trending')">🔥 Trending</button>
        <button class="tab" id="t_name" onclick="setMode('name_only')">✍️ Only Name</button>
        <button class="tab" id="t_custom" onclick="setMode('custom')">⚙️ Full Custom</button>
      </div>

      <div id="trendingBox">
        <label>Pick Trending Project:</label>
        <select id="trendSelect" onchange="applyTrending()">
          <option value="Automatic Smart Street Light & Traffic Control|AELIC / Embedded">Smart Street Light & Traffic Control (Arduino / AELIC)</option>
          <option value="ESP32 IoT Patient Heart & Temp Monitor|IoT / Biomedical">ESP32 IoT Patient Health Monitor (Trending)</option>
          <option value="AI Smart Attendance & 75% Shortage Alert|Python / CSE">AI Smart Attendance & Shortage Alert (Python)</option>
          <option value="RC Low-Pass Noise Filter & Transient Analyzer|Network Analysis">RC Noise Filter & Transient Analyzer (Network Analysis)</option>
          <option value="Smart Ultrasonic Blind Walking Stick|PBL / 1st Year">Smart Ultrasonic Blind Walking Stick (PBL)</option>
        </select>
      </div>

      <label>Project Name:</label>
      <input id="projTitle" value="Automatic Smart Street Light & Traffic Control" />

      <div class="grid2">
        <div><label>Subject / Branch:</label><input id="projBranch" value="AELIC / Embedded" /></div>
        <div><label>Max Budget (₹ INR):</label><input id="projBudget" type="number" value="600" min="200" step="50" /></div>
      </div>

      <div id="customNotesBox" style="display:none;">
        <label>Extra Technical Details:</label>
        <textarea id="extraNotes" placeholder="e.g., Buzzer alert, LCD display..."></textarea>
      </div>

      <button class="btn-ai" onclick="compileProject()">⚡ AI Compile Circuit & Budget (Free)</button>

      <div class="sim-box">
        <div style="font-size:11px;color:#38bdf8;font-weight:700;margin-bottom:4px;">LIVE CONNECTION SIMULATION</div>
        <svg id="simSvg" width="100%" height="145" viewBox="0 0 460 145"></svg>
      </div>

      <div class="lang-row">
        <button class="lbtn active" id="l_kn" onclick="setLang('kn')">ಕನ್ನಡ</button>
        <button class="lbtn" id="l_en" onclick="setLang('en')">English</button>
        <button class="lbtn" id="l_hi" onclick="setLang('hi')">हिन्दी</button>
      </div>
      <div class="viva-box" id="vivaText"></div>
      <ul class="bom" id="bomList"></ul>
    </div>

    <div class="card">
      <h2>2. Madam's Exact Word (.docx) Rules</h2>
      <div class="grid2">
        <div><label>Student Name</label><input id="stName" value="Prajwal Metri" /></div>
        <div><label>USN</label><input id="stUsn" value="1AM23CS045" /></div>
      </div>
      <div class="grid2">
        <div><label>College Name</label><input id="clgName" value="AMC Engineering College" /></div>
        <div><label>Guide / Madam Name</label><input id="guideName" value="Prof. Anitha K" /></div>
      </div>
      <div class="grid2">
        <div>
          <label>Font Name</label>
          <select id="fontName">
            <option value="Times New Roman">Times New Roman</option>
            <option value="Arial">Arial</option>
            <option value="Calibri">Calibri</option>
            <option value="Georgia">Georgia</option>
            <option value="Verdana">Verdana</option>
            <option value="Cambria">Cambria</option>
          </select>
        </div>
        <div>
          <label>Heading Style</label>
          <select id="hBold">
            <option value="true">Bold Headings</option>
            <option value="false">Normal Headings</option>
          </select>
        </div>
      </div>
      <div class="grid2">
        <div>
          <label>Heading / Body Size</label>
          <div class="grid2">
            <select id="hSize"><option value="14">14 pt</option><option value="16">16 pt</option></select>
            <select id="bSize"><option value="12">12 pt</option><option value="11">11 pt</option></select>
          </div>
        </div>
        <div>
          <label>Line Spacing</label>
          <select id="lineSpacing"><option value="1.15">1.15</option><option value="1.5">1.5 (VTU)</option></select>
        </div>
      </div>

      <label>Exact Section Order:</label>
      <select id="secOrder">
        <option value="intro,components,working,code,conclusion">1. Intro → 2. Components → 3. Working → 4. Code → 5. Conclusion</option>
        <option value="intro,working,components,code,conclusion">1. Intro → 2. Working → 3. Components → 4. Code → 5. Conclusion</option>
        <option value="intro,components,code,working,conclusion">1. Intro → 2. Components → 3. Code → 4. Working → 5. Conclusion</option>
      </select>

      <button class="btn-pay" id="mainPayBtn" onclick="handleDownloadClick()">Download Word (.docx) + Dump Code — FREE (2-Day Launch)</button>
      <div class="wipe-alert" id="wipeAlert">🛡️ Downloaded & project details wiped from server memory!</div>
      <div id="codeSection" style="display:none;">
        <label style="color:#10b981;">✅ Ready-to-Dump Code:</label>
        <div class="code-box" id="codeOut"></div>
      </div>
    </div>
  </div>

<script>
if ("serviceWorker" in navigator) navigator.serviceWorker.register("/sw.js");
let deferredPrompt = null;
window.addEventListener("beforeinstallprompt", (e) => {
  e.preventDefault();
  deferredPrompt = e;
  document.getElementById("installBtn").style.display = "inline-block";
});
function installPWA() {
  if (deferredPrompt) { deferredPrompt.prompt(); deferredPrompt = null; }
  else alert("Browser menu (⋮) -> 'Add to Home Screen' / 'Install App' click maadi!");
}

let currentUser = "", isFreeNow = true, hoursLeft = 48, curMode = "trending", curLang = "kn", compiledProj = null;

function setMode(m) {
  curMode = m;
  ["trend", "name", "custom"].forEach(k => document.getElementById("t_" + k).classList.toggle("active", m.startsWith(k)));
  document.getElementById("trendingBox").style.display = m === "trending" ? "block" : "none";
  document.getElementById("customNotesBox").style.display = m === "custom" ? "block" : "none";
  if (m === "trending") applyTrending();
}
function applyTrending() {
  const [t, b] = document.getElementById("trendSelect").value.split("|");
  document.getElementById("projTitle").value = t;
  document.getElementById("projBranch").value = b;
  compileProject();
}
function setLang(l) {
  curLang = l;
  ["kn", "en", "hi"].forEach(x => document.getElementById("l_" + x).classList.toggle("active", x === l));
  if (compiledProj) document.getElementById("vivaText").textContent = compiledProj.viva[curLang];
}
function updatePriceUI() {
  const banner = document.getElementById("topBanner");
  const btn = document.getElementById("mainPayBtn");
  if (isFreeNow) {
    banner.classList.remove("paid");
    banner.textContent = `🎉 LAUNCH OFFER: 100% FREE for ${hoursLeft} hrs left! (₹49 after 2 days) 🎉`;
    btn.textContent = "Download Word (.docx) + Dump Code — FREE (Launch Offer)";
  } else {
    banner.classList.add("paid");
    banner.textContent = "⚡ 2-DAY FREE LAUNCH ENDED — Exact Word (.docx) + Dump Code at ₹49 ⚡";
    btn.textContent = "Unlock & Download Word (.docx) + Code — ₹49";
  }
}
async function doLogin() {
  const contact = document.getElementById("loginContact").value.trim();
  if (contact.length < 5) { alert("Valid Email or Mobile Number haaki!"); return; }
  const res = await fetch("/api/login", { method: "POST", body: JSON.stringify({ contact }) });
  const data = await res.json();
  currentUser = contact;
  isFreeNow = data.is_free;
  hoursLeft = data.hours_left;
  document.getElementById("userBadge").textContent = currentUser;
  document.getElementById("loginModal").style.display = "none";
  updatePriceUI();
  compileProject();
}
function logout() { currentUser = ""; document.getElementById("loginModal").style.display = "flex"; }
async function compileProject() {
  const res = await fetch("/api/compile", {
    method: "POST",
    body: JSON.stringify({
      mode: curMode,
      title: document.getElementById("projTitle").value,
      branch: document.getElementById("projBranch").value,
      budget: document.getElementById("projBudget").value,
      extra_notes: document.getElementById("extraNotes").value
    })
  });
  compiledProj = await res.json();
  document.getElementById("vivaText").textContent = compiledProj.viva[curLang];
  document.getElementById("bomList").innerHTML = compiledProj.components.map(c => `<li><b>${c.name}</b> — ₹${c.cost}</li>`).join("") + `<li><b style="color:#10b981;">Total Cost: ₹${compiledProj.total_cost}</b></li>`;
  document.getElementById("codeOut").textContent = compiledProj.code;
  const svg = document.getElementById("simSvg");
  svg.innerHTML = `<rect x="20" y="35" width="115" height="75" rx="8" fill="#0369a1" stroke="#38bdf8" stroke-width="2"/><text x="77" y="77" fill="#fff" font-size="12" font-weight="bold" text-anchor="middle">CONTROLLER</text><rect x="190" y="35" width="100" height="75" rx="8" fill="#1e293b" stroke="#f59e0b" stroke-width="2"/><text x="240" y="77" fill="#fbbf24" font-size="12" font-weight="bold" text-anchor="middle">INPUT/SENSOR</text><rect x="345" y="35" width="95" height="75" rx="8" fill="#064e3b" stroke="#10b981" stroke-width="2"/><text x="392" y="77" fill="#6ee7b7" font-size="12" font-weight="bold" text-anchor="middle">OUTPUT</text><line x1="135" y1="72" x2="190" y2="72" stroke="#fbbf24" stroke-width="3" class="wire"/><line x1="290" y1="72" x2="345" y2="72" stroke="#10b981" stroke-width="3" class="wire"/>`;
}
function handleDownloadClick() {
  if (!currentUser) { document.getElementById("loginModal").style.display = "flex"; return; }
  if (isFreeNow) {
    executeDownloadAndWipe();
  } else {
    document.getElementById("payModal").style.display = "flex";
  }
}
async function executeDownloadAndWipe() {
  document.getElementById("payModal").style.display = "none";
  const cfg = {
    contact: currentUser, project: compiledProj,
    student: document.getElementById("stName").value, usn: document.getElementById("stUsn").value,
    college: document.getElementById("clgName").value, guide: document.getElementById("guideName").value,
    font: document.getElementById("fontName").value, h_bold: document.getElementById("hBold").value === "true",
    h_size: document.getElementById("hSize").value, b_size: document.getElementById("bSize").value,
    spacing: document.getElementById("lineSpacing").value, order: document.getElementById("secOrder").value.split(",")
  };
  const res = await fetch("/api/unlock_and_wipe", { method: "POST", body: JSON.stringify(cfg) });
  const blob = await res.blob();
  const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = `${cfg.usn}_Report.docx`; a.click();
  document.getElementById("codeSection").style.display = "block";
  document.getElementById("wipeAlert").style.display = "block";
}
</script>
</body>
</html>"""


class ThreadingHTTPServer(ThreadingMixIn, HTTPServer):
    daemon_threads = True


class StudioHandler(BaseHTTPRequestHandler):
    def log_message(self, *args):
        pass

    def send_json(self, obj):
        raw = json.dumps(obj).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def do_GET(self):
        if self.path == "/manifest.json":
            raw = MANIFEST_JSON.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/manifest+json")
            self.send_header("Content-Length", str(len(raw)))
            self.end_headers()
            self.wfile.write(raw)
        elif self.path == "/sw.js":
            raw = SW_JS.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/javascript")
            self.send_header("Content-Length", str(len(raw)))
            self.end_headers()
            self.wfile.write(raw)
        elif self.path == "/robots.txt":
            raw = ROBOTS_TXT.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/plain")
            self.send_header("Content-Length", str(len(raw)))
            self.end_headers()
            self.wfile.write(raw)
        elif self.path == "/sitemap.xml":
            raw = SITEMAP_XML.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/xml")
            self.send_header("Content-Length", str(len(raw)))
            self.end_headers()
            self.wfile.write(raw)
        else:
            raw = HTML_PAGE.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(raw)))
            self.end_headers()
            self.wfile.write(raw)

    def do_POST(self):
        n = int(self.headers.get("Content-Length", 0))
        body = json.loads(self.rfile.read(n).decode("utf-8")) if n else {}
        if self.path == "/api/login":
            contact = body.get("contact", "").strip().lower()
            with LOCK:
                users = load_users()
                if contact not in users:
                    users[contact] = {"joined_at": time.time()}
                    save_users(users)
                rem_sec = max(0, int(LAUNCH_END_TS - time.time()))
                self.send_json({
                    "is_free": is_free_period(),
                    "hours_left": max(1, round(rem_sec / 3600))
                })
        elif self.path == "/api/compile":
            self.send_json(generate_project_bundle(body.get("mode"), body.get("title", ""), body.get("branch", "Engineering"), body.get("budget", 600), body.get("extra_notes", "")))
        elif self.path == "/api/unlock_and_wipe":
            docx_bytes = build_docx_bytes(body["project"], body)
            del body
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            self.send_header("Content-Length", str(len(docx_bytes)))
            self.end_headers()
            self.wfile.write(docx_bytes)


if __name__ == "__main__":
    env_port = os.environ.get("PORT")
    port = int(env_port) if env_port else 8095
    if not env_port:
        for p in range(8095, 8120):
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                if s.connect_ex(("127.0.0.1", p)) != 0:
                    port = p
                    break
    srv = ThreadingHTTPServer(("0.0.0.0", port), StudioHandler)
    print(f"ProjForge running on port {port}")
    if not env_port:
        threading.Timer(0.6, lambda: webbrowser.open(f"http://127.0.0.1:{port}")).start()
    srv.serve_forever()