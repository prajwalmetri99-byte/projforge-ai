import io
import os
import re
import json
from datetime import datetime
from flask import Flask, render_template_string, request, jsonify, send_file
import requests
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

app = Flask(__name__)

# ==================== CONFIGURATION ====================
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "").strip()


# ==================== 1. GEMINI AI ENGINE (DIRECT REST API) ====================
def generate_project_blueprint(topic, branch, level):
    if GEMINI_API_KEY:
        try:
            url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={GEMINI_API_KEY}"
            prompt = f"""
You are a senior Indian Engineering Professor (VTU / Autonomous curriculum) and Principal Systems Architect.
Generate a comprehensive engineering mini-project blueprint in strict JSON format for:
- Project Title / Topic: "{topic}"
- Engineering Branch: "{branch}"
- Academic Level: "{level}"

Return ONLY valid JSON matching this exact schema (no markdown fences, no extra text):
{{
  "title": "Formal Engineering Project Title",
  "tagline": "One-line technical summary",
  "abstract": "Detailed 250-word academic abstract covering problem statement, proposed methodology, core technologies, and measurable outcomes.",
  "problem_statement": "Detailed 160-word problem definition explaining existing drawbacks in real-world scenarios.",
  "objectives": ["Objective 1", "Objective 2", "Objective 3", "Objective 4", "Objective 5"],
  "literature_survey": [
    {{"paper": "Author et al. (2024) - Title", "summary": "Key finding and limitation addressed by our system."}},
    {{"paper": "Author et al. (2025) - Title", "summary": "Key finding and limitation addressed by our system."}},
    {{"paper": "Author et al. (2026) - Title", "summary": "Key finding and limitation addressed by our system."}}
  ],
  "hardware_requirements": [
    {{"name": "ESP32 / Arduino Uno R3", "spec": "Main Microcontroller Board", "cost_inr": 450}},
    {{"name": "Primary Sensor Module", "spec": "Input Transducer / Sensor", "cost_inr": 250}},
    {{"name": "Actuator / Relay / Display", "spec": "Output Control Unit", "cost_inr": 180}},
    {{"name": "Breadboard & Jumper Wires", "spec": "830-Point Solderless Breadboard + M-M/M-F Wires", "cost_inr": 140}}
  ],
  "software_requirements": [
    "OS: Windows 11 / Ubuntu 22.04 LTS",
    "IDE / Tool: VS Code / Arduino IDE / Python 3.11",
    "Libraries / Frameworks: Relevant libraries for this project"
  ],
  "block_diagram_nodes": [
    "Input / Sensor Layer",
    "Signal Conditioning",
    "Core Controller / AI Unit",
    "Decision & Control Logic",
    "Output Actuators / Cloud UI"
  ],
  "pin_connections": [
    {{"from_pin": "Arduino D2 / GPIO 4", "to_pin": "Sensor OUT Pin", "signal": "Digital / Analog Input", "wire_color": "#38bdf8"}},
    {{"from_pin": "Arduino D8 / GPIO 5", "to_pin": "Relay / Actuator IN", "signal": "PWM / Digital Control", "wire_color": "#f59e0b"}},
    {{"from_pin": "Arduino 5V Rail", "to_pin": "Breadboard +VCC Bus", "signal": "5V DC Power", "wire_color": "#ef4444"}},
    {{"from_pin": "Arduino GND Pin", "to_pin": "Breadboard -GND Bus", "signal": "Common Ground", "wire_color": "#94a3b8"}}
  ],
  "ppt_slides": [
    {{"slide": "Slide 1: Title & Team", "points": "Project Title, USN, Guide Name & Department"}},
    {{"slide": "Slide 2: Problem Statement", "points": "Limitations of existing manual systems & motivation"}},
    {{"slide": "Slide 3: Objectives", "points": "Core engineering goals & target accuracy/latency"}},
    {{"slide": "Slide 4: Block Diagram & Hardware", "points": "Controller, Sensors, Breadboard wiring & BOM"}},
    {{"slide": "Slide 5: Working & Algorithm", "points": "Step-by-step data flow & embedded/AI logic"}},
    {{"slide": "Slide 6: Results & Future Scope", "points": "Prototype test results, applications & scalability"}}
  ],
  "working_principle": "Detailed 350-word step-by-step explanation of system architecture, data flow, algorithm/circuit operation, and fail-safe mechanisms.",
  "source_code_language": "Python or C++",
  "source_code": "Complete, clean, well-commented working source code (minimum 65-95 lines) ready to compile or execute.",
  "results_and_discussion": "Detailed 200-word analysis of testing parameters, accuracy/latency metrics, and real-world performance observations.",
  "advantages": ["Advantage 1", "Advantage 2", "Advantage 3", "Advantage 4"],
  "applications": ["Application 1", "Application 2", "Application 3", "Application 4"],
  "future_scope": "150-word discussion on scalability, IoT/Cloud upgrades, and industrial deployment.",
  "references": [
    "[1] IEEE Standard / Research Paper 1 relevant to topic",
    "[2] Technical Reference Book / Datasheet 2",
    "[3] IEEE / Springer Journal Article 3 (2025)"
  ],
  "viva_qa": [
    {{"q_en": "Viva Question 1 in English?", "a_en": "Clear technical answer in English.", "q_kn": "ಕನ್ನಡದಲ್ಲಿ ವೈವಾ ಪ್ರಶ್ನೆ 1?", "a_kn": "ಕನ್ನಡದಲ್ಲಿ ಸ್ಪಷ್ಟ ಉತ್ತರ."}},
    {{"q_en": "Viva Question 2?", "a_en": "Answer 2", "q_kn": "ಕನ್ನಡ ಪ್ರಶ್ನೆ 2?", "a_kn": "ಕನ್ನಡ ಉತ್ತರ 2"}},
    {{"q_en": "Viva Question 3?", "a_en": "Answer 3", "q_kn": "ಕನ್ನಡ ಪ್ರಶ್ನೆ 3?", "a_kn": "ಕನ್ನಡ ಉತ್ತರ 3"}},
    {{"q_en": "Viva Question 4?", "a_en": "Answer 4", "q_kn": "ಕನ್ನಡ ಪ್ರಶ್ನೆ 4?", "a_kn": "ಕನ್ನಡ ಉತ್ತರ 4"}},
    {{"q_en": "Viva Question 5?", "a_en": "Answer 5", "q_kn": "ಕನ್ನಡ ಪ್ರಶ್ನೆ 5?", "a_kn": "ಕನ್ನಡ ಉತ್ತರ 5"}},
    {{"q_en": "Viva Question 6?", "a_en": "Answer 6", "q_kn": "ಕನ್ನಡ ಪ್ರಶ್ನೆ 6?", "a_kn": "ಕನ್ನಡ ಉತ್ತರ 6"}},
    {{"q_en": "Viva Question 7?", "a_en": "Answer 7", "q_kn": "ಕನ್ನಡ ಪ್ರಶ್ನೆ 7?", "a_kn": "ಕನ್ನಡ ಉತ್ತರ 7"}},
    {{"q_en": "Viva Question 8?", "a_en": "Answer 8", "q_kn": "ಕನ್ನಡ ಪ್ರಶ್ನೆ 8?", "a_kn": "ಕನ್ನಡ ಉತ್ತರ 8"}},
    {{"q_en": "Viva Question 9?", "a_en": "Answer 9", "q_kn": "ಕನ್ನಡ ಪ್ರಶ್ನೆ 9?", "a_kn": "ಕನ್ನಡ ಉತ್ತರ 9"}},
    {{"q_en": "Viva Question 10?", "a_en": "Answer 10", "q_kn": "ಕನ್ನಡ ಪ್ರಶ್ನೆ 10?", "a_kn": "ಕನ್ನಡ ಉತ್ತರ 10"}}
  ]
}}
"""
            payload = {
                "contents": [{"parts": [{"text": prompt}]}],
                "generationConfig": {"responseMimeType": "application/json", "temperature": 0.4}
            }
            resp = requests.post(url, json=payload, timeout=45)
            if resp.status_code == 200:
                res_json = resp.json()
                raw_text = res_json["candidates"][0]["content"]["parts"][0]["text"].strip()
                raw_text = re.sub(r"^```json\s*|\s*```$", "", raw_text, flags=re.MULTILINE)
                return json.loads(raw_text)
        except Exception as e:
            print("Gemini API fallback:", e)

    return _fallback_blueprint(topic, branch, level)


def _fallback_blueprint(topic, branch, level):
    clean_title = topic.strip().title() or "Smart Automated Embedded & AI Monitoring System"
    return {
        "title": clean_title,
        "tagline": f"An intelligent {branch} engineering solution for {level} academic implementation.",
        "abstract": (
            f"This project presents the design, simulation, and prototype implementation of '{clean_title}' "
            f"tailored for {branch} engineering applications. Conventional systems suffer from manual latency, "
            "high power overhead, and lack of real-time telemetry. To overcome these limitations, the proposed "
            "architecture integrates precision sensing, deterministic microcontroller/algorithmic processing, and "
            "automated actuation with real-time status logging. Experimental validation confirms a 94.6% operational "
            "accuracy with sub-200ms response latency, making it cost-effective and scalable for campus and industrial deployment."
        ),
        "problem_statement": (
            f"Existing manual and legacy approaches related to {clean_title} require continuous human supervision, "
            "are prone to human error, and lack predictive fault alerts. Developing an automated, low-cost, and "
            "modular system bridges the gap between theoretical engineering principles and practical reliability."
        ),
        "objectives": [
            f"To design a modular system architecture for {clean_title}.",
            "To minimize power consumption and hardware footprint while maintaining high accuracy.",
            "To implement real-time threshold detection and automated alert/control mechanisms.",
            "To validate system performance under varying load and environmental test cases.",
            "To prepare standardized VTU/Autonomous technical documentation and reproducible source code."
        ],
        "literature_survey": [
            {"paper": "R. Kumar et al. (2024) - IEEE Access", "summary": "Proposed a baseline monitoring model; lacked low-power edge optimization."},
            {"paper": "S. Kulkarni & P. Rao (2025) - Springer", "summary": "Demonstrated sensor fusion techniques; high BOM cost limited student prototyping."},
            {"paper": "A. Sharma et al. (2026) - Elsevier", "summary": "Introduced real-time analytics; our work simplifies deployment with open-source firmware."}
        ],
        "hardware_requirements": [
            {"name": "Arduino Uno R3 / ESP32 Board", "spec": "ATmega328P / 32-bit Dual-Core MCU", "cost_inr": 480},
            {"name": "Primary Sensor Module", "spec": "High-precision digital/analog sensor", "cost_inr": 260},
            {"name": "5V Relay & Status Indicator LED", "spec": "Optocoupler relay + high-bright LED/Buzzer", "cost_inr": 190},
            {"name": "830-Point Breadboard & Jumpers", "spec": "Solderless prototyping board + M-M wires", "cost_inr": 150}
        ],
        "software_requirements": [
            "Operating System: Windows 10/11 or Ubuntu Linux",
            "Development Environment: Arduino IDE 2.3 / VS Code / Python 3.11",
            "Libraries: Wire.h, Adafruit_SSD1306, Serial Monitor"
        ],
        "block_diagram_nodes": [
            "Regulated Power (5V/3.3V)",
            "Multi-Sensor Input Layer",
            "Core MCU / AI Processor",
            "Threshold Control Engine",
            "Relay, LED & Actuator"
        ],
        "pin_connections": [
            {"from_pin": "Arduino Digital Pin D2", "to_pin": "Sensor Module OUT (Breadboard Row 12)", "signal": "Digital Input Signal", "wire_color": "#38bdf8"},
            {"from_pin": "Arduino Digital Pin D8", "to_pin": "Relay / Actuator IN (Breadboard Row 24)", "signal": "Active-HIGH Control", "wire_color": "#f59e0b"},
            {"from_pin": "Arduino 5V Pin", "to_pin": "Breadboard +VCC Power Rail", "signal": "+5V Regulated Supply", "wire_color": "#ef4444"},
            {"from_pin": "Arduino GND Pin", "to_pin": "Breadboard -GND Ground Rail", "signal": "Common Ground (0V)", "wire_color": "#94a3b8"}
        ],
        "ppt_slides": [
            {"slide": "Slide 1: Title & Introduction", "points": f"{clean_title} — Team Details & Guide"},
            {"slide": "Slide 2: Problem Statement", "points": "Drawbacks of manual monitoring & need for automation"},
            {"slide": "Slide 3: Proposed Architecture", "points": "Block diagram, Arduino/ESP32 & Breadboard interfacing"},
            {"slide": "Slide 4: Hardware & Software", "points": "Component specs, budget (~Rs. 1,080) & firmware stack"},
            {"slide": "Slide 5: Working & Simulation", "points": "Real-time sensor filtering & threshold actuation"},
            {"slide": "Slide 6: Results & Conclusion", "points": "94.6% accuracy, <200ms latency & future IoT scope"}
        ],
        "working_principle": (
            f"The working of '{clean_title}' is organized into three synchronized stages: Data Acquisition, "
            "Signal Processing, and Automated Actuation. All sensors and output modules are mounted on an 830-point "
            "solderless breadboard with shared +5V and GND power rails from the microcontroller. In the continuous loop, "
            "raw sensor values are sampled every 100 ms, filtered using a moving-average algorithm, and compared against "
            "pre-set thresholds to trigger the relay and status indicators immediately."
        ),
        "source_code_language": "C++ (Arduino / Embedded Firmware)",
        "source_code": f"""// ============================================================================
// Project: {clean_title}
// Branch : {branch} | Level: {level}
// Generated by ProjForge AI
// ============================================================================

#define SENSOR_PIN 2
#define RELAY_PIN  8
#define LED_PIN    13

void setup() {{
  Serial.begin(115200);
  pinMode(SENSOR_PIN, INPUT);
  pinMode(RELAY_PIN, OUTPUT);
  pinMode(LED_PIN, OUTPUT);
  Serial.println("SYSTEM READY: {clean_title}");
}}

void loop() {{
  int sensorState = digitalRead(SENSOR_PIN);
  if (sensorState == HIGH) {{
    digitalWrite(RELAY_PIN, HIGH);
    digitalWrite(LED_PIN, HIGH);
    Serial.println("[ACTIVE] Trigger Detected -> Output ON");
  }} else {{
    digitalWrite(RELAY_PIN, LOW);
    digitalWrite(LED_PIN, LOW);
    Serial.println("[IDLE] Monitoring Normal");
  }}
  delay(200);
}}
""",
        "results_and_discussion": (
            "The breadboard prototype was validated across 50 test cycles. Actuation latency stayed consistently "
            "under 180 ms with clean signal transitions and zero false triggers."
        ),
        "advantages": [
            "Low hardware cost and easy breadboard assembly",
            "Fast response latency (< 200 ms)",
            "Modular pin mapping extendable to IoT cloud",
            "Clear visual LED and Serial telemetry"
        ],
        "applications": [
            "Smart Campus & Laboratory Automation",
            "Industrial Safety & Preventive Maintenance",
            "Energy Conservation & Resource Monitoring",
            "Academic Mini-Project Demonstrations"
        ],
        "future_scope": (
            "Can be upgraded with ESP32 Wi-Fi MQTT telemetry, mobile push alerts, and custom PCB fabrication."
        ),
        "references": [
            "[1] IEEE Std 1451-2024, Standard for Smart Transducer Interface for Sensors and Actuators.",
            "[2] Muhammad Ali Mazidi, 'Embedded Systems & Microcontroller Architecture', Pearson Education.",
            "[3] VTU Academic Project Guidelines & Technical Documentation Standards (2025-26)."
        ],
        "viva_qa": [
            {"q_en": "What is the primary objective of your project?", "a_en": f"To automate and optimize {clean_title} with high accuracy and low cost.", "q_kn": "ನಿಮ್ಮ ಪ್ರಾಜೆಕ್ಟ್‌ನ ಮುಖ್ಯ ಉದ್ದೇಶವೇನು?", "a_kn": f"{clean_title} ಅನ್ನು ಕಡಿಮೆ ವೆಚ್ಚದಲ್ಲಿ ಸ್ವಯಂಚಾಲಿತಗೊಳಿಸುವುದು."},
            {"q_en": "Why do we use a breadboard in prototyping?", "a_en": "It allows solderless, reusable circuit connections and easy pin debugging.", "q_kn": "ಪ್ರೊಟೊಟೈಪಿಂಗ್‌ನಲ್ಲಿ ಬ್ರೆಡ್‌ಬೋರ್ಡ್ ಏಕೆ ಬಳಸುತ್ತೇವೆ?", "a_kn": "ಸೋಲ್ಡರಿಂಗ್ ಇಲ್ಲದೆ ಸುಲಭವಾಗಿ ಕನೆಕ್ಷನ್ ಬದಲಾಯಿಸಲು ಮತ್ತು ಪರೀಕ್ಷಿಸಲು ಬಳಸುತ್ತೇವೆ."},
            {"q_en": "What is the function of the +5V and GND rails on the breadboard?", "a_en": "They distribute common regulated power and ground reference to all connected modules.", "q_kn": "ಬ್ರೆಡ್‌ಬೋರ್ಡ್‌ನಲ್ಲಿ +5V ಮತ್ತು GND ಲೈನ್‌ಗಳ ಕೆಲಸವೇನು?", "a_kn": "ಎಲ್ಲಾ ಸೆನ್ಸರ್ ಮತ್ತು ಮಾಡ್ಯೂಲ್‌ಗಳಿಗೆ ಸಮಾನ ವಿದ್ಯುತ್ ಪೂರೈಸಲು ಬಳಸಲಾಗುತ್ತದೆ."},
            {"q_en": "What is the total cost of your prototype?", "a_en": "Approximately Rs. 1,080 including microcontroller, sensor, relay, and breadboard.", "q_kn": "ನಿಮ್ಮ ಪ್ರಾಜೆಕ್ಟ್ ಮಾದರಿಯ ಒಟ್ಟು ವೆಚ್ಚ ಎಷ್ಟು?", "a_kn": "ಸುಮಾರು ₹1,080 ವೆಚ್ಚವಾಗುತ್ತದೆ."},
            {"q_en": "What is the response time of your system?", "a_en": "Less than 200 milliseconds from sensor trigger to output actuation.", "q_kn": "ನಿಮ್ಮ ಸಿಸ್ಟಮ್‌ನ ರೆಸ್ಪಾನ್ಸ್ ಸಮಯ ಎಷ್ಟು?", "a_kn": "200 ಮಿಲಿಸೆಕೆಂಡ್‌ಗಳಿಗಿಂತ ಕಡಿಮೆ ಸಮಯದಲ್ಲಿ ಪ್ರತಿಕ್ರಿಯಿಸುತ್ತದೆ."},
            {"q_en": "How can this project be scaled for real-world industry?", "a_en": "By fabricating a custom PCB, adding IP65 enclosure, and integrating cloud logging.", "q_kn": "ಇದನ್ನು ಕೈಗಾರಿಕಾ ಮಟ್ಟಕ್ಕೆ ಹೇಗೆ ವಿಸ್ತರಿಸಬಹುದು?", "a_kn": "ಕಸ್ಟಮ್ PCB ಮತ್ತು IoT ಕ್ಲೌಡ್ ಡ್ಯಾಶ್‌ಬೋರ್ಡ್ ಸೇರಿಸುವ ಮೂಲಕ ವಿಸ್ತರಿಸಬಹುದು."}
        ]
    }


# ==================== 2. 15+ PAGE VTU DOCX REPORT BUILDER ====================
def _set_cell_shading(cell, fill_hex):
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading)


def _add_page_border(section):
    pgBorders = parse_xml(
        f'<w:pgBorders {nsdecls("w")} w:offsetFrom="page">'
        '<w:top w:val="single" w:sz="12" w:space="24" w:color="1E3A8A"/>'
        '<w:left w:val="single" w:sz="12" w:space="24" w:color="1E3A8A"/>'
        '<w:bottom w:val="single" w:sz="12" w:space="24" w:color="1E3A8A"/>'
        '<w:right w:val="single" w:sz="12" w:space="24" w:color="1E3A8A"/>'
        '</w:pgBorders>'
    )
    section._sectPr.append(pgBorders)


def _add_heading(doc, text, level=1):
    p = doc.add_heading(text, level=level)
    for run in p.runs:
        run.font.name = "Times New Roman"
        run.font.color.rgb = RGBColor(20, 45, 110) if level == 1 else RGBColor(30, 64, 140)
    return p


def _add_body(doc, text, bold=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_after = Pt(8)
    run = p.add_run(text)
    run.font.name = "Times New Roman"
    run.font.size = Pt(12)
    run.bold = bold
    return p


def build_vtu_docx(blueprint, student_name, usn, college, branch, guide_name):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.2)
        section.right_margin = Inches(1.0)
        _add_page_border(section)

    title = blueprint.get("title", "Engineering Mini Project")
    year_str = datetime.now().strftime("%Y-%m")

    # PAGE 1: COVER PAGE
    for txt, sz, bld, itl, clr in [
        ("VISVESVARAYA TECHNOLOGICAL UNIVERSITY / AUTONOMOUS\n", 14, True, False, RGBColor(15, 35, 95)),
        (f"{college.upper()}\nDepartment of {branch}\n\n", 13, True, False, None),
        ("A MINI-PROJECT REPORT ON\n\n", 12, False, True, None),
        (f'"{title.upper()}"\n\n', 18, True, False, RGBColor(180, 30, 30)),
        (f"Submitted in partial fulfillment of the requirements for the award of the degree of\nBACHELOR OF ENGINEERING IN {branch.upper()}\n\nSubmitted By:\n", 12, False, False, None),
        (f"{student_name.upper()}   (USN: {usn.upper()})\n\n", 13, True, False, None),
        (f"Under the Guidance of:\nProf. {guide_name}\nDepartment of {branch}\n\nAcademic Year: {year_str}", 12, False, False, None),
    ]:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(txt)
        r.font.name = "Times New Roman"
        r.font.size = Pt(sz)
        r.bold = bld
        r.italic = itl
        if clr:
            r.font.color.rgb = clr
    doc.add_page_break()

    # PAGE 2: CERTIFICATE
    _add_heading(doc, f"{college.upper()}", 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
    _add_heading(doc, f"Department of {branch}", 2).alignment = WD_ALIGN_PARAGRAPH.CENTER
    _add_heading(doc, "BONAFIDE CERTIFICATE", 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
    _add_body(
        doc,
        f'Certified that the Mini-Project work entitled "{title}" is a bonafide work carried out by '
        f"{student_name} (USN: {usn}), a bonafide student of {college}, in partial fulfillment for the award "
        f"of Bachelor of Engineering in {branch} during the academic year {year_str}."
    )
    doc.add_paragraph("\n\n")
    sig_table = doc.add_table(rows=2, cols=3)
    sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, (h, n) in enumerate([
        ("Signature of Guide", f"Prof. {guide_name}"),
        ("Signature of HOD", f"HOD, Dept. of {branch}"),
        ("Signature of Principal", f"Principal, {college}")
    ]):
        sig_table.cell(0, i).text = h
        sig_table.cell(0, i).paragraphs[0].runs[0].bold = True
        sig_table.cell(1, i).text = n
    doc.add_page_break()

    # PAGE 3: DECLARATION & ACKNOWLEDGEMENT
    _add_heading(doc, "DECLARATION", 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
    _add_body(
        doc,
        f'I, {student_name} (USN: {usn}), student of {branch}, {college}, hereby declare that the Mini-Project '
        f'entitled "{title}" has been independently carried out by me under the supervision of Prof. {guide_name}.'
    )
    _add_body(doc, f"\nPlace: Karnataka\nDate: {datetime.now().strftime('%d-%m-%Y')}                                   Signature: ({student_name})", bold=True)
    _add_heading(doc, "ACKNOWLEDGEMENT", 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
    _add_body(
        doc,
        f"I express my sincere gratitude to my guide Prof. {guide_name}, HOD, Principal, and faculty of {college} "
        "for their continuous guidance and lab support throughout this mini-project."
    )
    doc.add_page_break()

    # PAGE 4: ABSTRACT & TABLE OF CONTENTS
    _add_heading(doc, "ABSTRACT", 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
    _add_body(doc, blueprint.get("abstract", ""))
    _add_heading(doc, "TABLE OF CONTENTS", 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
    toc_items = [
        ("1.", "CHAPTER 1: INTRODUCTION & LITERATURE SURVEY", "05"),
        ("2.", "CHAPTER 2: SYSTEM ARCHITECTURE & CIRCUIT DESIGN", "08"),
        ("3.", "CHAPTER 3: HARDWARE, SOFTWARE & BUDGET ESTIMATION", "10"),
        ("4.", "CHAPTER 4: METHODOLOGY & SOURCE CODE IMPLEMENTATION", "12"),
        ("5.", "CHAPTER 5: RESULTS, APPLICATIONS, FUTURE SCOPE & REFERENCES", "15"),
        ("6.", "APPENDIX-A: TOP VIVA-VOCE QUESTIONS & ANSWERS (EN & KN)", "17"),
    ]
    toc_tbl = doc.add_table(rows=len(toc_items) + 1, cols=3)
    toc_tbl.style = "Table Grid"
    for idx, h in enumerate(["Sl.", "Chapter / Section Title", "Page"]):
        c = toc_tbl.cell(0, idx)
        c.text = h
        _set_cell_shading(c, "1E3A8A")
        c.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
        c.paragraphs[0].runs[0].bold = True
    for r_idx, (num, sec_name, pg) in enumerate(toc_items, start=1):
        toc_tbl.cell(r_idx, 0).text = num
        toc_tbl.cell(r_idx, 1).text = sec_name
        toc_tbl.cell(r_idx, 2).text = pg
    doc.add_page_break()

    # CHAPTER 1
    _add_heading(doc, "CHAPTER 1: INTRODUCTION & LITERATURE SURVEY", 1)
    _add_heading(doc, "1.1 Problem Statement", 2)
    _add_body(doc, blueprint.get("problem_statement", ""))
    _add_heading(doc, "1.2 Project Objectives", 2)
    for obj in blueprint.get("objectives", []):
        p = doc.add_paragraph(obj, style="List Bullet")
        p.runs[0].font.name = "Times New Roman"
    _add_heading(doc, "1.3 Literature Survey", 2)
    lit_list = blueprint.get("literature_survey", [])
    if lit_list:
        lt = doc.add_table(rows=len(lit_list) + 1, cols=2)
        lt.style = "Table Grid"
        for idx, h in enumerate(["Reference Paper", "Key Findings & Research Gap"]):
            c = lt.cell(0, idx)
            c.text = h
            _set_cell_shading(c, "1E3A8A")
            c.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
            c.paragraphs[0].runs[0].bold = True
        for i, item in enumerate(lit_list, start=1):
            lt.cell(i, 0).text = item.get("paper", "")
            lt.cell(i, 1).text = item.get("summary", "")
    doc.add_page_break()

    # CHAPTER 2
    _add_heading(doc, "CHAPTER 2: SYSTEM ARCHITECTURE & CIRCUIT DESIGN", 1)
    _add_heading(doc, "2.1 Functional Block Diagram Flow", 2)
    nodes = blueprint.get("block_diagram_nodes", [])
    for idx, node in enumerate(nodes):
        bt = doc.add_table(rows=1, cols=1)
        bt.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = bt.cell(0, 0)
        cell.text = f"[ STAGE {idx + 1} ]   {node.upper()}"
        _set_cell_shading(cell, "DBEAFE")
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        cell.paragraphs[0].runs[0].bold = True
        if idx < len(nodes) - 1:
            arr = doc.add_paragraph("⬇")
            arr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _add_heading(doc, "2.2 Pin-to-Pin Interfacing Table", 2)
    pins = blueprint.get("pin_connections", [])
    if pins:
        pt = doc.add_table(rows=len(pins) + 1, cols=3)
        pt.style = "Table Grid"
        for idx, h in enumerate(["Controller Pin", "Target Module Pin", "Signal / Protocol"]):
            c = pt.cell(0, idx)
            c.text = h
            _set_cell_shading(c, "1E3A8A")
            c.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
            c.paragraphs[0].runs[0].bold = True
        for i, pin in enumerate(pins, start=1):
            pt.cell(i, 0).text = pin.get("from_pin", "")
            pt.cell(i, 1).text = pin.get("to_pin", "")
            pt.cell(i, 2).text = pin.get("signal", "")
    doc.add_page_break()

    # CHAPTER 3
    _add_heading(doc, "CHAPTER 3: HARDWARE, SOFTWARE & BUDGET ESTIMATION", 1)
    hw_list = blueprint.get("hardware_requirements", [])
    total_cost = 0
    if hw_list:
        ht = doc.add_table(rows=len(hw_list) + 2, cols=4)
        ht.style = "Table Grid"
        for idx, h in enumerate(["Sl.", "Component", "Specification", "Est. Cost (INR)"]):
            c = ht.cell(0, idx)
            c.text = h
            _set_cell_shading(c, "1E3A8A")
            c.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
            c.paragraphs[0].runs[0].bold = True
        for i, hw in enumerate(hw_list, start=1):
            cost = int(hw.get("cost_inr", 0))
            total_cost += cost
            ht.cell(i, 0).text = str(i)
            ht.cell(i, 1).text = hw.get("name", "")
            ht.cell(i, 2).text = hw.get("spec", "")
            ht.cell(i, 3).text = f"Rs. {cost}"
        ht.cell(len(hw_list) + 1, 1).text = "TOTAL ESTIMATED COST"
        ht.cell(len(hw_list) + 1, 3).text = f"Rs. {total_cost}"
    _add_heading(doc, "3.2 Software Requirements", 2)
    for sw in blueprint.get("software_requirements", []):
        doc.add_paragraph(sw, style="List Bullet")
    doc.add_page_break()

    # CHAPTER 4
    _add_heading(doc, "CHAPTER 4: WORKING & COMPLETE SOURCE CODE", 1)
    _add_body(doc, blueprint.get("working_principle", ""))
    _add_heading(doc, f"Source Code ({blueprint.get('source_code_language', 'Code')})", 2)
    ct = doc.add_table(rows=1, cols=1)
    ct.style = "Table Grid"
    cc = ct.cell(0, 0)
    _set_cell_shading(cc, "F3F4F6")
    r = cc.paragraphs[0].add_run(blueprint.get("source_code", ""))
    r.font.name = "Consolas"
    r.font.size = Pt(9.5)
    doc.add_page_break()

    # CHAPTER 5
    _add_heading(doc, "CHAPTER 5: RESULTS, APPLICATIONS, FUTURE SCOPE & REFERENCES", 1)
    _add_body(doc, blueprint.get("results_and_discussion", ""))
    _add_heading(doc, "Advantages & Applications", 2)
    for adv in blueprint.get("advantages", []):
        doc.add_paragraph("Advantage: " + adv, style="List Bullet")
    for ap in blueprint.get("applications", []):
        doc.add_paragraph("Application: " + ap, style="List Bullet")
    _add_heading(doc, "Future Scope", 2)
    _add_body(doc, blueprint.get("future_scope", ""))
    _add_heading(doc, "References", 2)
    for ref in blueprint.get("references", []):
        _add_body(doc, ref, align=WD_ALIGN_PARAGRAPH.LEFT)
    doc.add_page_break()

    # APPENDIX A: VIVA Q&A
    _add_heading(doc, "APPENDIX-A: VIVA-VOCE QUESTIONS & ANSWERS (ENGLISH & KANNADA)", 1)
    for idx, qa in enumerate(blueprint.get("viva_qa", []), start=1):
        _add_body(doc, f"Q{idx} (EN): {qa.get('q_en', '')}", bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)
        _add_body(doc, f"Ans (EN): {qa.get('a_en', '')}", align=WD_ALIGN_PARAGRAPH.LEFT)
        if qa.get("q_kn"):
            _add_body(doc, f"ಪ್ರಶ್ನೆ {idx} (KN): {qa.get('q_kn', '')}", bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)
            _add_body(doc, f"ಉತ್ತರ (KN): {qa.get('a_kn', '')}", align=WD_ALIGN_PARAGRAPH.LEFT)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


# ==================== 3. FRONTEND HTML (MULTI-PAGE + PROFILE DROPDOWN + 3D HARDWARE LAB) ====================
HTML_PAGE = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
  <title>ProjForge AI — 3D Hardware & VTU Project Studio</title>
  <script src="https://cdn.tailwindcss.com"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/three.js/r128/three.min.js"></script>
  <link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;700&family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;600&display=swap" rel="stylesheet">
  <style>
    * { box-sizing: border-box; }
    html, body {
      max-width: 100%; overflow-x: hidden;
      font-family: 'Inter', sans-serif; background: #050811; color: #f1f5f9;
    }
    h1, h2, h3, .brand-font { font-family: 'Space Grotesk', sans-serif; }
    code, pre { font-family: 'JetBrains Mono', monospace; }

    /* 4-Sec Cinematic Splash */
    #splash-screen {
      position: fixed; inset: 0; z-index: 9999;
      background: radial-gradient(circle at center, #0f172a 0%, #020617 100%);
      display: flex; flex-direction: column; align-items: center; justify-content: center;
      transition: opacity 0.6s ease, visibility 0.6s ease;
    }
    .cyber-ring {
      width: 104px; height: 104px; border-radius: 50%;
      border: 3px solid transparent; border-top-color: #38bdf8; border-right-color: #818cf8;
      animation: spinRing 1.1s linear infinite;
      box-shadow: 0 0 35px rgba(56, 189, 248, 0.45);
      display: flex; align-items: center; justify-content: center;
    }
    .cyber-core {
      width: 64px; height: 64px; border-radius: 16px;
      background: linear-gradient(135deg, #2563eb, #7c3aed);
      display: flex; align-items: center; justify-content: center;
      font-weight: 800; font-size: 24px; color: white;
      animation: pulseCore 1.4s ease-in-out infinite alternate;
    }
    @keyframes spinRing { 100% { transform: rotate(360deg); } }
    @keyframes pulseCore { from { transform: scale(0.92); } to { transform: scale(1.06); } }
    .flash-bar { width: 220px; height: 4px; background: #1e293b; border-radius: 999px; overflow: hidden; margin-top: 20px; }
    .flash-progress { height: 100%; width: 0%; background: linear-gradient(90deg, #38bdf8, #818cf8, #c084fc); animation: loadBar 3.5s ease forwards; }
    @keyframes loadBar { 0% { width: 0%; } 60% { width: 78%; } 100% { width: 100%; } }

    .glass { background: rgba(15, 23, 42, 0.8); backdrop-filter: blur(16px); border: 1px solid rgba(148, 163, 184, 0.14); }
    .page-section { display: none; animation: fadeIn 0.25s ease; }
    .page-section.active { display: block; }
    @keyframes fadeIn { from { opacity: 0; transform: translateY(6px); } to { opacity: 1; transform: translateY(0); } }
  </style>
</head>
<body class="pb-20 md:pb-8">

  <!-- 1. SPLASH SCREEN -->
  <div id="splash-screen">
    <div class="cyber-ring"><div class="cyber-core">PF</div></div>
    <h1 class="text-2xl md:text-4xl font-bold mt-5 tracking-tight bg-gradient-to-r from-sky-400 via-indigo-400 to-purple-400 bg-clip-text text-transparent">
      ProjForge AI 3D
    </h1>
    <p id="splash-status" class="text-[11px] text-slate-400 mt-2 tracking-widest uppercase">
      Loading 3D Breadboard &amp; VTU Report Engine...
    </p>
    <div class="flash-bar"><div class="flash-progress"></div></div>
  </div>

  <!-- 2. LOGIN MODAL -->
  <div id="login-modal" class="fixed inset-0 z-50 hidden items-center justify-center bg-black/80 backdrop-blur-md p-4">
    <div class="glass max-w-md w-full rounded-2xl p-6 shadow-2xl border border-sky-500/30">
      <div class="flex items-center gap-3 mb-4">
        <div class="w-10 h-10 rounded-xl bg-gradient-to-tr from-blue-600 to-violet-600 flex items-center justify-center font-bold text-white">PF</div>
        <div>
          <h2 class="text-lg font-bold text-white">Student Workspace Sign-In</h2>
          <p class="text-xs text-slate-400">Auto-fills your Name &amp; USN on VTU .docx cover pages</p>
        </div>
      </div>

      <div id="referred-banner" class="hidden mb-3 px-3 py-2 rounded-xl bg-emerald-500/15 border border-emerald-500/30 text-xs text-emerald-300">
        🎉 Invited by a friend (<span id="ref-by-code" class="font-mono font-bold"></span>)!
      </div>

      <form id="login-form" class="space-y-3 mt-3">
        <div>
          <label class="text-xs text-slate-300">Student Email ID *</label>
          <input id="login-email" type="email" required placeholder="student@college.edu.in"
            class="w-full mt-1 px-3.5 py-2.5 rounded-xl bg-slate-900 border border-slate-700 text-sm focus:border-sky-400 focus:outline-none" />
        </div>
        <div class="grid grid-cols-2 gap-3">
          <div>
            <label class="text-xs text-slate-300">Full Name *</label>
            <input id="login-name" type="text" required placeholder="Rahul Sharma"
              class="w-full mt-1 px-3.5 py-2.5 rounded-xl bg-slate-900 border border-slate-700 text-sm focus:border-sky-400 focus:outline-none" />
          </div>
          <div>
            <label class="text-xs text-slate-300">USN / Roll No *</label>
            <input id="login-usn" type="text" required placeholder="1VT23CS045"
              class="w-full mt-1 px-3.5 py-2.5 rounded-xl bg-slate-900 border border-slate-700 text-sm uppercase focus:border-sky-400 focus:outline-none" />
          </div>
        </div>
        <div>
          <label class="text-xs text-slate-300">College Name</label>
          <input id="login-college" type="text" placeholder="VTU Affiliated Engineering College"
            class="w-full mt-1 px-3.5 py-2.5 rounded-xl bg-slate-900 border border-slate-700 text-sm focus:border-sky-400 focus:outline-none" />
        </div>
        <button type="submit"
          class="w-full py-3 rounded-xl bg-gradient-to-r from-sky-500 to-indigo-600 font-semibold text-white text-sm shadow-lg">
          Enter ProjForge Studio →
        </button>
      </form>
      <button onclick="skipLogin()" class="w-full mt-3 text-xs text-slate-400 hover:text-slate-200 underline">
        Continue as Guest
      </button>
    </div>
  </div>

  <!-- 3. REFER FRIENDS MODAL -->
  <div id="refer-modal" class="fixed inset-0 z-50 hidden items-center justify-center bg-black/80 backdrop-blur-md p-4">
    <div class="glass max-w-md w-full rounded-2xl p-6 shadow-2xl border border-indigo-500/40 relative">
      <button onclick="closeReferModal()" class="absolute top-4 right-4 text-slate-400 hover:text-white text-sm">✕</button>
      <div class="flex items-center gap-3 mb-3">
        <div class="w-10 h-10 rounded-xl bg-gradient-to-br from-emerald-500 to-teal-600 flex items-center justify-center text-lg">🎁</div>
        <div>
          <h2 class="text-base font-bold text-white">Refer Friends &amp; Classmates</h2>
          <p class="text-xs text-slate-400">Share ProjForge AI with your engineering batchmates</p>
        </div>
      </div>

      <div class="bg-slate-900 border border-slate-800 rounded-xl p-3 mt-3">
        <span class="text-[10px] uppercase tracking-wider text-slate-400">Your Referral Code</span>
        <div class="flex items-center justify-between mt-1">
          <span id="my-ref-code" class="font-mono font-bold text-sm text-sky-400">PF-STUDENT</span>
          <button onclick="copyRefCode()" class="text-xs px-2.5 py-1 rounded-lg bg-slate-800 text-slate-200">Copy</button>
        </div>
      </div>

      <div class="bg-slate-900 border border-slate-800 rounded-xl p-3 mt-3">
        <span class="text-[10px] uppercase tracking-wider text-slate-400">Invite Link</span>
        <input id="my-ref-link" type="text" readonly
          class="w-full mt-1 px-2.5 py-2 rounded-lg bg-slate-950 border border-slate-800 text-xs text-slate-300 font-mono" />
      </div>

      <div class="grid grid-cols-2 gap-2.5 mt-4">
        <button onclick="shareOnWhatsApp()" class="py-2.5 px-3 rounded-xl bg-emerald-600 font-semibold text-white text-xs">
          💬 WhatsApp Share
        </button>
        <button onclick="copyInviteLink()" class="py-2.5 px-3 rounded-xl bg-indigo-600 font-semibold text-white text-xs">
          🔗 Copy Link
        </button>
      </div>
    </div>
  </div>

  <!-- 4. CLEAN TOP NAVBAR WITH PROFILE DROPDOWN (NO OVERFLOW) -->
  <header class="sticky top-0 z-40 glass border-b border-slate-800/90">
    <div class="max-w-6xl mx-auto px-4 py-3 flex items-center justify-between">
      <div class="flex items-center gap-2.5 cursor-pointer" onclick="navigatePage('page-home')">
        <div class="w-9 h-9 rounded-xl bg-gradient-to-br from-sky-500 to-indigo-600 flex items-center justify-center font-bold text-white shadow">PF</div>
        <div>
          <span class="brand-font font-bold text-base sm:text-lg text-white">ProjForge <span class="text-sky-400">AI</span></span>
          <span class="hidden sm:inline-block ml-2 text-[10px] px-2 py-0.5 rounded-full bg-sky-500/15 text-sky-300 border border-sky-500/30">3D Studio</span>
        </div>
      </div>

      <!-- Desktop Page Nav -->
      <nav class="hidden md:flex items-center gap-1.5 bg-slate-900/90 p-1 rounded-xl border border-slate-800">
        <button onclick="navigatePage('page-home')" id="nav-d-home" class="nav-d-btn px-3 py-1.5 rounded-lg text-xs font-semibold bg-sky-600 text-white">🏠 Create</button>
        <button onclick="navigatePage('page-report')" id="nav-d-report" class="nav-d-btn px-3 py-1.5 rounded-lg text-xs font-semibold text-slate-300 hover:text-white">📄 Report &amp; PPT</button>
        <button onclick="navigatePage('page-3d')" id="nav-d-3d" class="nav-d-btn px-3 py-1.5 rounded-lg text-xs font-semibold text-slate-300 hover:text-white">🧊 3D Breadboard Lab</button>
        <button onclick="navigatePage('page-code')" id="nav-d-code" class="nav-d-btn px-3 py-1.5 rounded-lg text-xs font-semibold text-slate-300 hover:text-white">💻 Code &amp; Viva</button>
      </nav>

      <!-- Profile Avatar + Dropdown Menu -->
      <div class="relative">
        <button onclick="toggleProfileMenu()" id="profile-trigger"
          class="flex items-center gap-2 px-3 py-1.5 rounded-xl bg-slate-900 hover:bg-slate-800 border border-slate-700 text-xs text-slate-200">
          <span id="avatar-initial" class="w-6 h-6 rounded-full bg-gradient-to-tr from-sky-500 to-indigo-600 flex items-center justify-center font-bold text-[11px] text-white">G</span>
          <span id="profile-short-name" class="max-w-[90px] truncate font-medium">Guest</span>
          <span class="text-slate-400 text-[10px]">▼</span>
        </button>

        <div id="profile-dropdown" class="hidden absolute right-0 mt-2 w-56 rounded-2xl glass border border-slate-700 shadow-2xl py-2 z-50">
          <div class="px-4 py-2 border-b border-slate-800">
            <p id="menu-full-name" class="text-xs font-bold text-white truncate">Guest Student</p>
            <p id="menu-usn" class="text-[11px] text-sky-400 font-mono">1VT23CS001</p>
          </div>
          <button onclick="closeProfileMenu(); openLoginModal();" class="w-full text-left px-4 py-2.5 text-xs text-slate-200 hover:bg-slate-800/80 flex items-center gap-2">
            <span>✏️</span> Edit Student Profile
          </button>
          <button onclick="closeProfileMenu(); openReferModal();" class="w-full text-left px-4 py-2.5 text-xs text-indigo-300 hover:bg-slate-800/80 flex items-center gap-2">
            <span>🎁</span> Refer Friends &amp; Share
          </button>
          <div class="border-t border-slate-800 my-1"></div>
          <button onclick="closeProfileMenu(); logoutUser();" class="w-full text-left px-4 py-2.5 text-xs text-rose-400 hover:bg-rose-500/15 font-semibold flex items-center gap-2">
            <span>🚪</span> Logout
          </button>
        </div>
      </div>
    </div>
  </header>

  <!-- ==================== PAGE 1: GENERATOR HOME ==================== -->
  <main class="max-w-5xl mx-auto px-4 pt-6">
    <section id="page-home" class="page-section active space-y-6">
      <div class="text-center max-w-2xl mx-auto">
        <span class="text-[11px] uppercase tracking-widest px-3 py-1 rounded-full bg-sky-500/10 text-sky-300 border border-sky-500/20">
          AI Engineering Workspace
        </span>
        <h1 class="text-2xl sm:text-4xl font-bold mt-3 leading-tight bg-gradient-to-r from-white via-sky-200 to-indigo-300 bg-clip-text text-transparent">
          Design Projects, 3D Breadboard Circuits &amp; 15+ Page VTU Reports
        </h1>
        <p class="text-slate-400 text-xs sm:text-sm mt-2">
          Enter any engineering topic below. Each section opens in its own dedicated clean workspace.
        </p>
      </div>

      <div class="glass rounded-2xl p-5 sm:p-7 max-w-3xl mx-auto shadow-xl border border-slate-800 space-y-4">
        <div>
          <label class="text-xs font-semibold uppercase tracking-wider text-sky-400">Project Topic or Problem Statement *</label>
          <input id="topic-input" type="text"
            placeholder="e.g., IoT Smart Helmet with Alcohol & Accident Detection"
            class="w-full mt-1.5 px-4 py-3.5 rounded-xl bg-slate-950 border border-slate-700 text-white text-sm focus:border-sky-400 focus:outline-none" />
        </div>

        <div class="grid grid-cols-1 sm:grid-cols-3 gap-3.5">
          <div>
            <label class="text-xs text-slate-400">Engineering Branch</label>
            <select id="branch-input" class="w-full mt-1 px-3 py-2.5 rounded-xl bg-slate-950 border border-slate-700 text-xs sm:text-sm text-white">
              <option>Computer Science &amp; Engineering (CSE)</option>
              <option>Electronics &amp; Communication (ECE)</option>
              <option>Artificial Intelligence &amp; ML (AIML)</option>
              <option>Electrical &amp; Electronics (EEE)</option>
              <option>Information Science (ISE)</option>
              <option>Mechanical Engineering</option>
              <option>Civil Engineering</option>
            </select>
          </div>
          <div>
            <label class="text-xs text-slate-400">Academic Level</label>
            <select id="level-input" class="w-full mt-1 px-3 py-2.5 rounded-xl bg-slate-950 border border-slate-700 text-xs sm:text-sm text-white">
              <option>2nd / 3rd Year Mini-Project</option>
              <option>Final Year Major Capstone Project</option>
              <option>Hackathon / Innovation Prototype</option>
            </select>
          </div>
          <div>
            <label class="text-xs text-slate-400">Guide Name (For Certificate)</label>
            <input id="guide-input" type="text" placeholder="Dr. S. K. Kulkarni"
              class="w-full mt-1 px-3 py-2.5 rounded-xl bg-slate-950 border border-slate-700 text-xs sm:text-sm text-white" />
          </div>
        </div>

        <div class="flex flex-wrap items-center gap-2 pt-1">
          <span class="text-xs text-slate-400">Quick Topics:</span>
          <button type="button" onclick="setSample('Smart Automated Railway Gate Control using IR & Arduino')" class="text-xs px-2.5 py-1 rounded-lg bg-slate-900 hover:bg-slate-800 text-sky-300 border border-slate-700">🚂 Smart Railway Gate</button>
          <button type="button" onclick="setSample('AI-Powered Leaf Disease Detection using CNN & Python')" class="text-xs px-2.5 py-1 rounded-lg bg-slate-900 hover:bg-slate-800 text-sky-300 border border-slate-700">🌿 Crop Disease AI</button>
          <button type="button" onclick="setSample('Smart Fire & Gas Leakage Alert System with Relay Cutoff')" class="text-xs px-2.5 py-1 rounded-lg bg-slate-900 hover:bg-slate-800 text-sky-300 border border-slate-700">🔥 Gas &amp; Fire Alert</button>
        </div>

        <button id="generate-btn" onclick="generateProject()"
          class="w-full py-3.5 rounded-xl bg-gradient-to-r from-sky-500 via-blue-600 to-indigo-600 font-bold text-white text-sm sm:text-base shadow-lg shadow-blue-600/25">
          ⚡ Generate Complete Project &amp; Open 3D Lab
        </button>
      </div>

      <!-- Quick Shortcut Cards after Generation -->
      <div id="quick-cards" class="hidden grid grid-cols-1 sm:grid-cols-3 gap-4 max-w-3xl mx-auto">
        <div onclick="navigatePage('page-3d')" class="glass p-4 rounded-2xl border border-sky-500/40 cursor-pointer hover:bg-slate-900 transition">
          <div class="text-2xl mb-1">🧊</div>
          <h3 class="font-bold text-sm text-sky-300">Open 3D Breadboard Lab →</h3>
          <p class="text-xs text-slate-400 mt-1">Interactive 3D Arduino, Breadboard, Sensor &amp; Live Wire Simulation</p>
        </div>
        <div onclick="navigatePage('page-report')" class="glass p-4 rounded-2xl border border-emerald-500/40 cursor-pointer hover:bg-slate-900 transition">
          <div class="text-2xl mb-1">📄</div>
          <h3 class="font-bold text-sm text-emerald-300">VTU Report &amp; PPT →</h3>
          <p class="text-xs text-slate-400 mt-1">Download 15+ Page Word (.docx) + Seminar PPT Slides</p>
        </div>
        <div onclick="navigatePage('page-code')" class="glass p-4 rounded-2xl border border-purple-500/40 cursor-pointer hover:bg-slate-900 transition">
          <div class="text-2xl mb-1">🎓</div>
          <h3 class="font-bold text-sm text-purple-300">Code &amp; Viva Q&amp;A →</h3>
          <p class="text-xs text-slate-400 mt-1">Ready-to-run Source Code + English &amp; Kannada Viva</p>
        </div>
      </div>
    </section>

    <!-- ==================== PAGE 2: REPORT & PPT PAGE ==================== -->
    <section id="page-report" class="page-section space-y-5">
      <div id="empty-report" class="glass rounded-2xl p-8 text-center">
        <p class="text-sm text-slate-400">No project generated yet. Go to <b>Create</b> tab and enter your topic first!</p>
        <button onclick="navigatePage('page-home')" class="mt-3 px-4 py-2 rounded-xl bg-sky-600 text-xs font-semibold text-white">Go to Create →</button>
      </div>

      <div id="content-report" class="hidden space-y-5">
        <div class="glass rounded-2xl p-5 flex flex-col sm:flex-row sm:items-center justify-between gap-4 border border-emerald-500/30">
          <div>
            <span class="text-[11px] uppercase tracking-wider text-emerald-400 font-semibold">VTU / Autonomous Documentation</span>
            <h2 id="rep-title" class="text-lg sm:text-2xl font-bold text-white mt-0.5"></h2>
            <p id="rep-tagline" class="text-xs text-slate-400 mt-1"></p>
          </div>
          <button onclick="downloadDocxDirect()" id="download-docx-btn"
            class="px-5 py-3 rounded-xl bg-gradient-to-r from-emerald-500 to-teal-600 font-bold text-white text-xs sm:text-sm shadow-lg shrink-0">
            📥 Download 15+ Page .docx
          </button>
        </div>

        <div class="glass rounded-2xl p-5 space-y-4">
          <h3 class="text-base font-bold text-sky-300">Abstract &amp; Working</h3>
          <p id="rep-abstract" class="text-xs sm:text-sm text-slate-300 leading-relaxed"></p>
          <p id="rep-working" class="text-xs sm:text-sm text-slate-300 leading-relaxed border-t border-slate-800 pt-3"></p>
        </div>

        <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
          <div class="glass rounded-2xl p-5">
            <h3 class="text-sm font-bold text-emerald-300 mb-2">💰 Hardware Bill of Materials (BOM)</h3>
            <div id="rep-bom" class="text-xs text-slate-300 space-y-2"></div>
          </div>
          <div class="glass rounded-2xl p-5">
            <h3 class="text-sm font-bold text-indigo-300 mb-2">📊 Seminar PPT Slide Outline</h3>
            <div id="rep-ppt" class="text-xs text-slate-300 space-y-2"></div>
          </div>
        </div>
      </div>
    </section>

    <!-- ==================== PAGE 3: SEPARATE 3D HARDWARE & BREADBOARD LAB ==================== -->
    <section id="page-3d" class="page-section space-y-5">
      <div class="glass rounded-2xl p-4 sm:p-5 border border-sky-500/30 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
        <div>
          <span class="text-[11px] uppercase tracking-wider text-sky-400 font-bold">Dedicated 3D Hardware Workbench</span>
          <h2 id="lab3d-title" class="text-lg sm:text-xl font-bold text-white">3D Breadboard, Arduino &amp; Sensor Circuit Lab</h2>
          <p class="text-xs text-slate-400">Drag / Touch to rotate 360° • Scroll / Pinch to zoom • Toggle Live Power Simulation</p>
        </div>
        <div class="flex flex-wrap gap-2">
          <button onclick="toggle3DSimulation()" id="sim-toggle-btn"
            class="px-4 py-2.5 rounded-xl bg-emerald-600 hover:bg-emerald-500 font-bold text-white text-xs shadow">
            ⚡ Live Simulation: ON
          </button>
          <button onclick="reset3DCamera()"
            class="px-3.5 py-2.5 rounded-xl bg-slate-800 hover:bg-slate-700 text-xs text-slate-200 border border-slate-700">
            🔄 Reset 3D View
          </button>
        </div>
      </div>

      <!-- 3D Interactive Canvas Container -->
      <div class="glass rounded-2xl p-2 border border-slate-800 relative overflow-hidden">
        <div id="threejs-container" class="w-full h-[360px] sm:h-[440px] rounded-xl bg-slate-950 cursor-grab active:cursor-grabbing"></div>
        <!-- Legend Overlay -->
        <div class="flex flex-wrap gap-3 px-3 py-2.5 text-[11px] text-slate-300 bg-slate-950/90 border-t border-slate-800 rounded-b-xl">
          <span class="flex items-center gap-1.5"><span class="w-3 h-3 rounded-sm bg-blue-600 inline-block"></span> Arduino / ESP32 MCU</span>
          <span class="flex items-center gap-1.5"><span class="w-3 h-3 rounded-sm bg-slate-200 inline-block"></span> 830-Pt Breadboard</span>
          <span class="flex items-center gap-1.5"><span class="w-3 h-3 rounded-sm bg-emerald-500 inline-block"></span> Input Sensor</span>
          <span class="flex items-center gap-1.5"><span class="w-3 h-3 rounded-sm bg-amber-500 inline-block"></span> Relay / Actuator</span>
          <span class="flex items-center gap-1.5"><span class="w-3 h-3 rounded-full bg-red-500 inline-block"></span> +5V / GND / Signal Jumper Wires</span>
        </div>
      </div>

      <!-- Pin-to-Pin Wiring Guide Below 3D Model -->
      <div class="glass rounded-2xl p-5">
        <h3 class="text-sm font-bold text-sky-300 mb-3">📌 Hardware Pin-to-Pin Breadboard Wiring Guide</h3>
        <div class="overflow-x-auto">
          <table class="w-full text-xs text-left border border-slate-800">
            <thead class="bg-slate-900 text-sky-300">
              <tr>
                <th class="p-2.5 border-b border-slate-800">Controller Pin</th>
                <th class="p-2.5 border-b border-slate-800">Breadboard / Module Pin</th>
                <th class="p-2.5 border-b border-slate-800">Signal / Wire</th>
              </tr>
            </thead>
            <tbody id="lab3d-pins" class="divide-y divide-slate-800 text-slate-300">
              <tr>
                <td class="p-2.5 font-mono text-sky-300">Arduino D2</td>
                <td class="p-2.5 font-mono text-indigo-300">Sensor OUT (Row 12)</td>
                <td class="p-2.5">Digital Input (Blue Wire)</td>
              </tr>
              <tr>
                <td class="p-2.5 font-mono text-sky-300">Arduino D8</td>
                <td class="p-2.5 font-mono text-indigo-300">Relay IN (Row 24)</td>
                <td class="p-2.5">Control Output (Yellow Wire)</td>
              </tr>
              <tr>
                <td class="p-2.5 font-mono text-sky-300">Arduino 5V &amp; GND</td>
                <td class="p-2.5 font-mono text-indigo-300">Breadboard + / - Rails</td>
                <td class="p-2.5">Power Supply (Red / Grey Wires)</td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>
    </section>

    <!-- ==================== PAGE 4: CODE & BILINGUAL VIVA PAGE ==================== -->
    <section id="page-code" class="page-section space-y-5">
      <div id="empty-code" class="glass rounded-2xl p-8 text-center">
        <p class="text-sm text-slate-400">Generate a project first to view its working source code and Kannada + English Viva Q&amp;A.</p>
        <button onclick="navigatePage('page-home')" class="mt-3 px-4 py-2 rounded-xl bg-sky-600 text-xs font-semibold text-white">Go to Create →</button>
      </div>

      <div id="content-code" class="hidden space-y-5">
        <div class="glass rounded-2xl p-5 space-y-3">
          <div class="flex items-center justify-between">
            <h3 id="code-lang-title" class="text-base font-bold text-emerald-300">💻 Working Source Code</h3>
            <button onclick="copySourceCode()" class="px-3 py-1.5 rounded-lg bg-slate-800 text-xs text-sky-300 border border-slate-700">📋 Copy Code</button>
          </div>
          <pre id="code-block" class="bg-slate-950 p-4 rounded-xl border border-slate-800 text-xs text-emerald-300 overflow-x-auto leading-relaxed"></pre>
        </div>

        <div class="glass rounded-2xl p-5 space-y-3">
          <h3 class="text-base font-bold text-purple-300">🎓 Top Viva-Voce Questions &amp; Answers (English + ಕನ್ನಡ)</h3>
          <div id="viva-grid" class="grid grid-cols-1 md:grid-cols-2 gap-3.5"></div>
        </div>
      </div>
    </section>
  </main>

  <!-- MOBILE BOTTOM NAVIGATION BAR (CLEAN 4 TABS) -->
  <nav class="md:hidden fixed bottom-0 inset-x-0 z-40 glass border-t border-slate-800 grid grid-cols-4 py-1.5 px-1">
    <button onclick="navigatePage('page-home')" id="nav-m-home" class="nav-m-btn flex flex-col items-center py-1 text-sky-400 font-bold">
      <span class="text-base">🏠</span><span class="text-[10px]">Create</span>
    </button>
    <button onclick="navigatePage('page-report')" id="nav-m-report" class="nav-m-btn flex flex-col items-center py-1 text-slate-400">
      <span class="text-base">📄</span><span class="text-[10px]">Report</span>
    </button>
    <button onclick="navigatePage('page-3d')" id="nav-m-3d" class="nav-m-btn flex flex-col items-center py-1 text-slate-400">
      <span class="text-base">🧊</span><span class="text-[10px]">3D Lab</span>
    </button>
    <button onclick="navigatePage('page-code')" id="nav-m-code" class="nav-m-btn flex flex-col items-center py-1 text-slate-400">
      <span class="text-base">🎓</span><span class="text-[10px]">Code/Viva</span>
    </button>
  </nav>

  <script>
    let currentBlueprint = null;
    let studentProfile = JSON.parse(localStorage.getItem("pf_student") || "null");

    // Splash & Init
    window.addEventListener("DOMContentLoaded", () => {
      const urlParams = new URLSearchParams(window.location.search);
      const refParam = urlParams.get("ref");
      if (refParam) {
        localStorage.setItem("pf_referred_by", refParam);
        document.getElementById("ref-by-code").innerText = refParam;
        document.getElementById("referred-banner").classList.remove("hidden");
      }

      setTimeout(() => {
        const s = document.getElementById("splash-screen");
        s.style.opacity = "0";
        s.style.visibility = "hidden";
        if (!studentProfile || !studentProfile.email) {
          openLoginModal();
        } else {
          updateProfileUI();
        }
      }, 3400);

      initThreeJS();
    });

    // Multi-Page Navigation
    function navigatePage(pageId) {
      document.querySelectorAll(".page-section").forEach(sec => sec.classList.remove("active"));
      document.getElementById(pageId).classList.add("active");

      // Update Desktop Nav
      const key = pageId.replace("page-", "");
      document.querySelectorAll(".nav-d-btn").forEach(b => {
        b.classList.remove("bg-sky-600", "text-white");
        b.classList.add("text-slate-300");
      });
      const dBtn = document.getElementById("nav-d-" + key);
      if (dBtn) { dBtn.classList.add("bg-sky-600", "text-white"); dBtn.classList.remove("text-slate-300"); }

      // Update Mobile Nav
      document.querySelectorAll(".nav-m-btn").forEach(b => {
        b.classList.remove("text-sky-400", "font-bold");
        b.classList.add("text-slate-400");
      });
      const mBtn = document.getElementById("nav-m-" + key);
      if (mBtn) { mBtn.classList.add("text-sky-400", "font-bold"); mBtn.classList.remove("text-slate-400"); }

      if (pageId === "page-3d") {
        setTimeout(resizeThreeJS, 60);
      }
      window.scrollTo({ top: 0, behavior: "smooth" });
    }

    // Profile Dropdown & Auth
    function toggleProfileMenu() {
      document.getElementById("profile-dropdown").classList.toggle("hidden");
    }
    function closeProfileMenu() {
      document.getElementById("profile-dropdown").classList.add("hidden");
    }
    window.addEventListener("click", (e) => {
      if (!e.target.closest("#profile-trigger") && !e.target.closest("#profile-dropdown")) {
        closeProfileMenu();
      }
    });

    function openLoginModal() {
      if (studentProfile) {
        document.getElementById("login-email").value = studentProfile.email || "";
        document.getElementById("login-name").value = studentProfile.name === "Guest Student" ? "" : (studentProfile.name || "");
        document.getElementById("login-usn").value = studentProfile.usn === "1VT23CS001" ? "" : (studentProfile.usn || "");
        document.getElementById("login-college").value = studentProfile.college || "";
      }
      const m = document.getElementById("login-modal");
      m.classList.remove("hidden"); m.classList.add("flex");
    }
    function closeLoginModal() {
      const m = document.getElementById("login-modal");
      m.classList.add("hidden"); m.classList.remove("flex");
    }
    function skipLogin() {
      studentProfile = { email: "", name: "Guest Student", usn: "1VT23CS001", college: "VTU Affiliated Engineering College" };
      updateProfileUI();
      closeLoginModal();
    }
    function logoutUser() {
      localStorage.removeItem("pf_student");
      studentProfile = null;
      currentBlueprint = null;
      document.getElementById("login-form").reset();
      document.getElementById("quick-cards").classList.add("hidden");
      document.getElementById("content-report").classList.add("hidden");
      document.getElementById("empty-report").classList.remove("hidden");
      document.getElementById("content-code").classList.add("hidden");
      document.getElementById("empty-code").classList.remove("hidden");
      updateProfileUI();
      navigatePage("page-home");
      openLoginModal();
    }

    document.getElementById("login-form").addEventListener("submit", (e) => {
      e.preventDefault();
      studentProfile = {
        email: document.getElementById("login-email").value.trim(),
        name: document.getElementById("login-name").value.trim(),
        usn: document.getElementById("login-usn").value.trim().toUpperCase(),
        college: document.getElementById("login-college").value.trim() || "VTU Affiliated Engineering College"
      };
      localStorage.setItem("pf_student", JSON.stringify(studentProfile));
      updateProfileUI();
      closeLoginModal();
    });

    function updateProfileUI() {
      const name = (studentProfile && studentProfile.name) || "Guest Student";
      const usn = (studentProfile && studentProfile.usn) || "1VT23CS001";
      document.getElementById("avatar-initial").innerText = name.charAt(0).toUpperCase();
      document.getElementById("profile-short-name").innerText = name.split(" ")[0];
      document.getElementById("menu-full-name").innerText = name;
      document.getElementById("menu-usn").innerText = usn;
    }

    // Refer Friends
    function getReferralCode() {
      if (studentProfile && studentProfile.usn && studentProfile.usn !== "1VT23CS001") {
        return "PF-" + studentProfile.usn.replace(/[^A-Z0-9]/gi, "");
      }
      let code = localStorage.getItem("pf_my_ref_code");
      if (!code) {
        code = "PF-" + Math.floor(1000 + Math.random() * 9000);
        localStorage.setItem("pf_my_ref_code", code);
      }
      return code;
    }
    function openReferModal() {
      const code = getReferralCode();
      document.getElementById("my-ref-code").innerText = code;
      document.getElementById("my-ref-link").value = `${window.location.origin}/?ref=${code}`;
      const m = document.getElementById("refer-modal");
      m.classList.remove("hidden"); m.classList.add("flex");
    }
    function closeReferModal() {
      const m = document.getElementById("refer-modal");
      m.classList.add("hidden"); m.classList.remove("flex");
    }
    function copyRefCode() {
      navigator.clipboard.writeText(document.getElementById("my-ref-code").innerText);
      alert("Referral code copied!");
    }
    function copyInviteLink() {
      navigator.clipboard.writeText(document.getElementById("my-ref-link").value);
      alert("Invite link copied!");
    }
    function shareOnWhatsApp() {
      const link = document.getElementById("my-ref-link").value;
      const msg = `🚀 Build your complete Engineering Project, 3D Breadboard Circuit & 15+ Page VTU Report (.docx) with Kannada + English Viva Q&A for FREE!\n\nTry ProjForge AI:\n${link}`;
      window.open("https://wa.me/?text=" + encodeURIComponent(msg), "_blank");
    }

    function setSample(txt) { document.getElementById("topic-input").value = txt; }

    // Generate Project
    async function generateProject() {
      const topic = document.getElementById("topic-input").value.trim();
      const branch = document.getElementById("branch-input").value;
      const level = document.getElementById("level-input").value;
      if (!topic) return alert("Please enter your project topic first!");

      const btn = document.getElementById("generate-btn");
      btn.disabled = true;
      btn.innerText = "⏳ Building 3D Circuit, Report & Viva Q&A...";

      try {
        const res = await fetch("/api/generate", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ topic, branch, level })
        });
        const data = await res.json();
        if (data.error) throw new Error(data.error);
        currentBlueprint = data.blueprint;
        populateAllPages(currentBlueprint);
        document.getElementById("quick-cards").classList.remove("hidden");
        navigatePage("page-3d");
      } catch (err) {
        alert("Error: " + err.message);
      } finally {
        btn.disabled = false;
        btn.innerText = "⚡ Generate Complete Project & Open 3D Lab";
      }
    }

    function populateAllPages(bp) {
      // Report Page
      document.getElementById("empty-report").classList.add("hidden");
      document.getElementById("content-report").classList.remove("hidden");
      document.getElementById("rep-title").innerText = bp.title || "";
      document.getElementById("rep-tagline").innerText = bp.tagline || "";
      document.getElementById("rep-abstract").innerText = bp.abstract || "";
      document.getElementById("rep-working").innerText = bp.working_principle || "";

      let total = 0;
      document.getElementById("rep-bom").innerHTML = (bp.hardware_requirements || []).map(h => {
        total += Number(h.cost_inr || 0);
        return `<div class="flex justify-between border-b border-slate-800 pb-1"><span>${h.name}</span><span class="text-emerald-400 font-semibold">₹${h.cost_inr}</span></div>`;
      }).join("") + `<div class="flex justify-between pt-1 font-bold text-white"><span>Total Est. Cost</span><span>₹${total}</span></div>`;

      document.getElementById("rep-ppt").innerHTML = (bp.ppt_slides || []).map(s => `
        <div class="border-b border-slate-800 pb-1"><span class="text-sky-300 font-semibold">${s.slide}:</span> ${s.points}</div>
      `).join("");

      // 3D Lab Page Pins
      document.getElementById("lab3d-title").innerText = `3D Hardware Lab: ${bp.title || "Circuit"}`;
      document.getElementById("lab3d-pins").innerHTML = (bp.pin_connections || []).map(p => `
        <tr>
          <td class="p-2.5 font-mono text-sky-300">${p.from_pin}</td>
          <td class="p-2.5 font-mono text-indigo-300">${p.to_pin}</td>
          <td class="p-2.5">${p.signal}</td>
        </tr>
      `).join("");

      // Code & Viva Page
      document.getElementById("empty-code").classList.add("hidden");
      document.getElementById("content-code").classList.remove("hidden");
      document.getElementById("code-lang-title").innerText = `💻 ${bp.source_code_language || "Source Code"}`;
      document.getElementById("code-block").innerText = bp.source_code || "";
      document.getElementById("viva-grid").innerHTML = (bp.viva_qa || []).map((v, i) => `
        <div class="bg-slate-900/90 p-4 rounded-xl border border-slate-800 space-y-1.5">
          <p class="text-xs font-bold text-sky-300">Q${i + 1}: ${v.q_en}</p>
          <p class="text-xs text-slate-300">${v.a_en}</p>
          ${v.q_kn ? `<p class="text-xs font-bold text-amber-300 pt-1">ಪ್ರಶ್ನೆ: ${v.q_kn}</p><p class="text-xs text-slate-300">${v.a_kn}</p>` : ""}
        </div>
      `).join("");
    }

    function copySourceCode() {
      navigator.clipboard.writeText(document.getElementById("code-block").innerText);
      alert("Source code copied!");
    }

    async function downloadDocxDirect() {
      if (!currentBlueprint) return;
      const btn = document.getElementById("download-docx-btn");
      btn.disabled = true; btn.innerText = "⏳ Generating DOCX...";
      try {
        const payload = {
          blueprint: currentBlueprint,
          student_name: (studentProfile && studentProfile.name) || "Student Name",
          usn: (studentProfile && studentProfile.usn) || "1VT23CS001",
          college: (studentProfile && studentProfile.college) || "VTU Affiliated Engineering College",
          branch: document.getElementById("branch-input").value,
          guide_name: document.getElementById("guide-input").value.trim() || "Project Guide"
        };
        const res = await fetch("/api/download-docx", {
          method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload)
        });
        if (!res.ok) throw new Error("Failed");
        const blob = await res.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement("a"); a.href = url; a.download = "ProjForge_VTU_Project_Report.docx";
        document.body.appendChild(a); a.click(); a.remove();
      } catch (err) {
        alert("Download error: " + err.message);
      } finally {
        btn.disabled = false; btn.innerText = "📥 Download 15+ Page .docx";
      }
    }

    // ==================== 3D INTERACTIVE BREADBOARD & ARDUINO ENGINE (THREE.JS) ====================
    let scene, camera, renderer, circuitGroup, statusLedMesh, signalPulseMesh;
    let simActive = true;
    let isDragging = false, prevMouse = { x: 0, y: 0 };

    function initThreeJS() {
      const container = document.getElementById("threejs-container");
      if (!container || typeof THREE === "undefined") return;

      scene = new THREE.Scene();
      scene.background = new THREE.Color(0x030712);

      camera = new THREE.PerspectiveCamera(45, container.clientWidth / container.clientHeight, 0.1, 1000);
      camera.position.set(0, 7.5, 10.5);
      camera.lookAt(0, 0, 0);

      renderer = new THREE.WebGLRenderer({ antialias: true });
      renderer.setSize(container.clientWidth, container.clientHeight);
      container.innerHTML = "";
      container.appendChild(renderer.domElement);

      // Lights
      scene.add(new THREE.AmbientLight(0xffffff, 0.65));
      const dirLight = new THREE.DirectionalLight(0x38bdf8, 1.1);
      dirLight.position.set(6, 12, 8);
      scene.add(dirLight);

      // Lab Grid Floor
      const grid = new THREE.GridHelper(24, 24, 0x1e3a8a, 0x0f172a);
      grid.position.y = -0.35;
      scene.add(grid);

      circuitGroup = new THREE.Group();
      scene.add(circuitGroup);

      // 1. Arduino / ESP32 Microcontroller Board (Left)
      const arduino = new THREE.Mesh(
        new THREE.BoxGeometry(3.4, 0.25, 2.3),
        new THREE.MeshStandardMaterial({ color: 0x1d4ed8, roughness: 0.4 })
      );
      arduino.position.set(-3.1, 0, 0);
      circuitGroup.add(arduino);

      // USB Port & MCU Chip
      const usb = new THREE.Mesh(
        new THREE.BoxGeometry(0.7, 0.4, 0.6),
        new THREE.MeshStandardMaterial({ color: 0xcbd5e1, metalness: 0.7 })
      );
      usb.position.set(-4.6, 0.25, -0.4);
      circuitGroup.add(usb);

      const chip = new THREE.Mesh(
        new THREE.BoxGeometry(1.3, 0.15, 0.45),
        new THREE.MeshStandardMaterial({ color: 0x0f172a })
      );
      chip.position.set(-2.8, 0.18, 0.1);
      circuitGroup.add(chip);

      // 2. 830-Point Solderless Breadboard (Right)
      const breadboard = new THREE.Mesh(
        new THREE.BoxGeometry(5.2, 0.28, 2.6),
        new THREE.MeshStandardMaterial({ color: 0xf8fafc, roughness: 0.5 })
      );
      breadboard.position.set(2.2, 0, 0);
      circuitGroup.add(breadboard);

      // Power Rails (+5V Red & -GND Blue stripes on Breadboard)
      const redRail = new THREE.Mesh(new THREE.BoxGeometry(4.9, 0.02, 0.08), new THREE.MeshBasicMaterial({ color: 0xef4444 }));
      redRail.position.set(2.2, 0.15, -1.1);
      circuitGroup.add(redRail);

      const blueRail = new THREE.Mesh(new THREE.BoxGeometry(4.9, 0.02, 0.08), new THREE.MeshBasicMaterial({ color: 0x2563eb }));
      blueRail.position.set(2.2, 0.15, 1.1);
      circuitGroup.add(blueRail);

      // 3. Input Sensor Module mounted on Breadboard
      const sensor = new THREE.Mesh(
        new THREE.BoxGeometry(1.1, 0.55, 0.8),
        new THREE.MeshStandardMaterial({ color: 0x10b981, roughness: 0.3 })
      );
      sensor.position.set(1.0, 0.4, -0.35);
      circuitGroup.add(sensor);

      // 4. Relay / Output Actuator Module mounted on Breadboard
      const relay = new THREE.Mesh(
        new THREE.BoxGeometry(1.2, 0.65, 0.9),
        new THREE.MeshStandardMaterial({ color: 0xf59e0b, roughness: 0.3 })
      );
      relay.position.set(3.4, 0.45, 0.25);
      circuitGroup.add(relay);

      // Live Blinking LED on Breadboard
      statusLedMesh = new THREE.Mesh(
        new THREE.SphereGeometry(0.16, 16, 16),
        new THREE.MeshBasicMaterial({ color: 0x22c55e })
      );
      statusLedMesh.position.set(2.2, 0.45, -0.3);
      circuitGroup.add(statusLedMesh);

      // 5. 3D Curved Jumper Wires connecting Arduino <-> Breadboard
      function add3DWire(p1, p2, p3, hexColor) {
        const curve = new THREE.CatmullRomCurve3([
          new THREE.Vector3(...p1),
          new THREE.Vector3(...p2),
          new THREE.Vector3(...p3)
        ]);
        const geo = new THREE.TubeGeometry(curve, 24, 0.045, 8, false);
        const mat = new THREE.MeshStandardMaterial({ color: hexColor });
        circuitGroup.add(new THREE.Mesh(geo, mat));
      }

      add3DWire([-1.7, 0.15, -0.8], [-0.4, 1.4, -1.0], [0.4, 0.16, -1.1], 0xef4444); // +5V Red
      add3DWire([-1.7, 0.15, 0.8], [-0.4, 1.3, 1.0], [0.4, 0.16, 1.1], 0x64748b);  // GND Grey
      add3DWire([-1.7, 0.15, -0.3], [-0.3, 1.6, -0.3], [1.0, 0.3, -0.35], 0x38bdf8); // Sensor Signal Blue
      add3DWire([-1.7, 0.15, 0.3], [0.6, 1.8, 0.3], [3.4, 0.3, 0.25], 0xfacc15);   // Relay Control Yellow

      // Moving Signal Pulse Sphere
      signalPulseMesh = new THREE.Mesh(
        new THREE.SphereGeometry(0.11, 12, 12),
        new THREE.MeshBasicMaterial({ color: 0x38bdf8 })
      );
      circuitGroup.add(signalPulseMesh);

      // Mouse & Touch Rotation Controls
      const dom = renderer.domElement;
      dom.addEventListener("mousedown", (e) => { isDragging = true; prevMouse = { x: e.clientX, y: e.clientY }; });
      window.addEventListener("mouseup", () => { isDragging = false; });
      window.addEventListener("mousemove", (e) => {
        if (!isDragging) return;
        circuitGroup.rotation.y += (e.clientX - prevMouse.x) * 0.01;
        circuitGroup.rotation.x += (e.clientY - prevMouse.y) * 0.008;
        prevMouse = { x: e.clientX, y: e.clientY };
      });

      dom.addEventListener("touchstart", (e) => {
        if (e.touches.length === 1) {
          isDragging = true;
          prevMouse = { x: e.touches[0].clientX, y: e.touches[0].clientY };
        }
      }, { passive: true });
      dom.addEventListener("touchmove", (e) => {
        if (!isDragging || e.touches.length !== 1) return;
        circuitGroup.rotation.y += (e.touches[0].clientX - prevMouse.x) * 0.012;
        circuitGroup.rotation.x += (e.touches[0].clientY - prevMouse.y) * 0.008;
        prevMouse = { x: e.touches[0].clientX, y: e.touches[0].clientY };
      }, { passive: true });
      dom.addEventListener("touchend", () => { isDragging = false; });

      dom.addEventListener("wheel", (e) => {
        e.preventDefault();
        camera.position.z = Math.max(5, Math.min(18, camera.position.z + e.deltaY * 0.01));
      }, { passive: false });

      animate3D();
    }

    let clockTick = 0;
    function animate3D() {
      requestAnimationFrame(animate3D);
      if (simActive && circuitGroup) {
        clockTick += 0.04;
        if (!isDragging) circuitGroup.rotation.y += 0.003;
        const pulse = (Math.sin(clockTick * 2) + 1) / 2;
        statusLedMesh.material.color.setHex(pulse > 0.5 ? 0x22c55e : 0x064e3b);
        signalPulseMesh.position.set(-1.7 + pulse * 5.1, 0.8 + Math.sin(pulse * Math.PI) * 0.9, 0.25);
      }
      if (renderer && scene && camera) renderer.render(scene, camera);
    }

    function toggle3DSimulation() {
      simActive = !simActive;
      const btn = document.getElementById("sim-toggle-btn");
      btn.innerText = simActive ? "⚡ Live Simulation: ON" : "⏸️ Simulation: PAUSED";
      btn.className = simActive
        ? "px-4 py-2.5 rounded-xl bg-emerald-600 hover:bg-emerald-500 font-bold text-white text-xs shadow"
        : "px-4 py-2.5 rounded-xl bg-amber-600 hover:bg-amber-500 font-bold text-white text-xs shadow";
    }

    function reset3DCamera() {
      if (circuitGroup) { circuitGroup.rotation.set(0, 0, 0); }
      if (camera) { camera.position.set(0, 7.5, 10.5); camera.lookAt(0, 0, 0); }
    }

    function resizeThreeJS() {
      const container = document.getElementById("threejs-container");
      if (!container || !renderer || !camera) return;
      camera.aspect = container.clientWidth / container.clientHeight;
      camera.updateProjectionMatrix();
      renderer.setSize(container.clientWidth, container.clientHeight);
    }
    window.addEventListener("resize", resizeThreeJS);
  </script>
</body>
</html>
"""


# ==================== 4. FLASK API ROUTES ====================
@app.route("/")
def index():
    return render_template_string(HTML_PAGE)


@app.route("/api/generate", methods=["POST"])
def api_generate():
    data = request.get_json(force=True)
    topic = data.get("topic", "").strip()
    branch = data.get("branch", "Computer Science & Engineering").strip()
    level = data.get("level", "3rd Year Mini-Project").strip()
    if not topic:
        return jsonify({"error": "Please enter a project topic."}), 400
    blueprint = generate_project_blueprint(topic, branch, level)
    return jsonify({"status": "ok", "blueprint": blueprint})


@app.route("/api/download-docx", methods=["POST"])
def download_docx():
    payload = request.get_json(force=True)
    blueprint = payload.get("blueprint", {})
    docx_buffer = build_vtu_docx(
        blueprint,
        payload.get("student_name") or "Student Name",
        payload.get("usn") or "1VT23CS001",
        payload.get("college") or "VTU Affiliated Engineering College",
        payload.get("branch") or "Engineering",
        payload.get("guide_name") or "Project Guide",
    )
    safe_name = "".join(c if c.isalnum() else "_" for c in blueprint.get("title", "Project_Report"))[:40]
    return send_file(
        docx_buffer,
        as_attachment=True,
        download_name=f"{safe_name}_VTU_Report.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))