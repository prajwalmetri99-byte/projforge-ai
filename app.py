import io
import os
import re
import json
from datetime import datetime
from flask import Flask, render_template_string, request, jsonify, send_file, Response
import requests
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from pptx import Presentation
from pptx.util import Inches as PptInches, Pt as PptPt
from pptx.dml.color import RGBColor as PptRGB

app = Flask(__name__)

# ==================== CONFIGURATION ====================
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "").strip()
FEEDBACK_LOG = []


# ==================== 1. SMART PROJECT TYPE & BLUEPRINT ENGINE ====================
def _detect_is_hardware(topic, branch, subject, explicit_flag=None):
    if isinstance(explicit_flag, bool):
        return explicit_flag
    text = f"{topic} {branch} {subject}".lower()
    sw_only_terms = [
        "filter", "eligibility", "tkinter", "website", "web app", "portal",
        "database", "dbms", "sql", "management system", "erp", "crm",
        "chatbot", "nlp", "machine learning", "deep learning", "data science",
        "csv", "excel", "compiler", "operating system", "algorithm", "software",
        "blockchain", "cybersecurity", "app development"
    ]
    hw_terms = [
        "circuit", "3d circuit", "tinkercad", "breadboard", "pcb", "schematic",
        "ic 741", "ic741", "opamp", "op-amp", "555", "lm324", "arduino", "esp32",
        "esp8266", "stm32", "raspberry pi", "sensor", "ultrasonic", "ir sensor",
        "pir", "dht11", "dht22", "gas sensor", "mq2", "mq135", "soil moisture",
        "water level", "relay", "servo", "motor", "dc motor", "stepper", "buzzer",
        "led", "oled", "lcd", "iot", "embedded", "vlsi", "verilog", "robotics",
        "drone", "smart irrigation", "home automation", "fire alarm", "radar"
    ]
    has_hw = any(k in text for k in hw_terms)
    has_sw = any(k in text for k in sw_only_terms)
    if has_hw:
        return True
    if has_sw:
        return False
    return branch.lower() in {"electronics & communication (ece)", "electrical & electronics (eee)", "mechanical engineering"}


def generate_project_blueprint(topic, branch, subject, level):
    is_hw_default = _detect_is_hardware(topic, branch, subject)
    if GEMINI_API_KEY:
        prompt = f"""
You are a Senior Indian Engineering Professor (VTU / Autonomous / AICTE).
Generate a complete, accurate engineering project blueprint in strict JSON for:
- Project Title: "{topic}"
- Engineering Branch: "{branch}"
- Core Academic Subject: "{subject}"
- Academic Level: "{level}"

CRITICAL RULES:
1. "requires_circuit": MUST be true ONLY if the project involves physical electronics, circuits, sensors, ICs, relays, motors, microcontrollers, or IoT hardware. MUST be false for pure software/Python/Tkinter/DBMS/Web/ML/Data/Eligibility-Filter projects.
2. Tailor every component, circuit diagram connection, code snippet, and viva question strictly to "{topic}" and "{subject}".
3. If "requires_circuit" is true:
   - Choose 3 hardware modules from: ["ic_741", "ic_lm324", "sensor_water_probes", "sensor_soil_moisture", "sensor_ultrasonic", "sensor_gas", "sensor_ir", "water_pump", "motor_servo", "relay_module", "oled_display", "piezo_buzzer", "led_indicator_bank"].
   - Provide accurate "controller_or_supply" (e.g., "9V DC Analog Supply (IC 741)", "Arduino Uno R3", "ESP32 DevKit").
4. "primary_code": Write 65-95 lines of clean, runnable, well-commented code matching "{topic}" and "{subject}".
5. "viva_qa": Provide 5 technical viva questions & answers translated into "en", "kn", "hi", "te", "ta".

Return ONLY valid JSON (no markdown fences):
{{
  "title": "Formal Engineering Project Title",
  "tagline": "Concise summary aligned with {subject}",
  "requires_circuit": {str(is_hw_default).lower()},
  "controller_or_supply": "Arduino Uno R3 or 9V DC Analog Supply",
  "verification_summary": {{
    "syllabus_match": "98% Syllabus Aligned ({subject})",
    "build_time": "2 Days",
    "expert_note": "Verified architecture, logic, and component/module specifications."
  }},
  "abstract": "Detailed 220-word academic abstract covering problem statement, {subject} principles, design, and results.",
  "problem_statement": "Detailed 140-word problem statement explaining existing limitations.",
  "objectives": ["Objective 1", "Objective 2", "Objective 3", "Objective 4"],
  "literature_survey": [
    {{"paper": "R. Kumar et al. (2024) - IEEE", "summary": "Key finding and research gap."}},
    {{"paper": "S. Rao & P. Kulkarni (2025) - Springer", "summary": "Key finding and research gap."}},
    {{"paper": "A. Sharma et al. (2026) - Elsevier", "summary": "Key finding and research gap."}}
  ],
  "hardware_requirements": [
    {{"name": "Item 1", "spec": "Specification", "cost_inr": 150}},
    {{"name": "Item 2", "spec": "Specification", "cost_inr": 200}}
  ],
  "software_requirements": [
    "OS: Windows 11 / Ubuntu 22.04 LTS",
    "Tools: Python 3.11 / Arduino IDE / Circuit Simulator",
    "Documentation: VTU / AICTE Standard Format"
  ],
  "circuit_3d_parts": [
    {{"type": "ic_741", "label": "Module 1", "pin_from": "Pin A", "pin_to": "Row 10", "wire_color": "#2563eb"}}
  ],
  "circuit_diagram_blocks": [
    {{"stage": "Power / Input", "title": "Input Stage", "detail": "Signal / Data Input"}},
    {{"stage": "Processing", "title": "Control / Logic Stage", "detail": "Threshold / Processing"}},
    {{"stage": "Output / Action", "title": "Output Stage", "detail": "Indicator / Actuator"}}
  ],
  "software_architecture_nodes": ["Input Layer", "Validation & Rules", "Core Processing Engine", "Storage / Export", "Output / Reporting"],
  "pin_connections": [
    {{"from_pin": "Source Pin / Module", "to_pin": "Target Pin / Rail", "signal": "Signal / Voltage"}}
  ],
  "ppt_slides": [
    {{"slide": "Title & Team Overview", "bullets": ["Project: {topic}", "Subject: {subject} ({branch})", "Level: {level}"]}},
    {{"slide": "Problem Statement", "bullets": ["Limitations of manual/legacy methods", "Need for automated {subject} system", "Key accuracy & usability goals"]}},
    {{"slide": "System Architecture & Design", "bullets": ["Modular input-processing-output design", "Clear interfaces & error handling", "Scalable & reliable implementation"]}},
    {{"slide": "Implementation & Code", "bullets": ["Structured {subject} implementation", "Clean data/signal handling", "Interactive user interface & outputs"]}},
    {{"slide": "Testing & Results", "bullets": ["Verified across test cases", "Fast & accurate response", "Practical documentation & reports"]}},
    {{"slide": "Applications & Future Scope", "bullets": ["Academic & real-world deployment", "Future enhancements & automation", "Conclusion"]}}
  ],
  "working_principle": "Detailed 260-word explanation of how the system works step by step.",
  "primary_code_language": "Python 3",
  "primary_code": "Complete runnable code.",
  "has_arduino_hardware": {str(is_hw_default).lower()},
  "arduino_code": "Arduino sketch if hardware is used.",
  "results_and_discussion": "Detailed 160-word testing results and performance evaluation.",
  "advantages": ["Advantage 1", "Advantage 2", "Advantage 3", "Advantage 4"],
  "applications": ["Application 1", "Application 2", "Application 3", "Application 4"],
  "future_scope": "120-word future enhancements.",
  "references": [
    "[1] IEEE Standard Technical Reference (2025).",
    "[2] Standard Textbook on {subject}.",
    "[3] VTU / AICTE Academic Project Guidelines (2026-27)."
  ],
  "viva_qa": [
    {{
      "en": {{"q": "What is the main objective of this project?", "a": "To solve the problem reliably using {subject} principles."}},
      "kn": {{"q": "ಈ ಪ್ರಾಜೆಕ್ಟ್‌ನ ಮುಖ್ಯ ಉದ್ದೇಶವೇನು?", "a": "{subject} ತತ್ವಗಳನ್ನು ಬಳಸಿ ಸಮಸ್ಯೆಯನ್ನು ನಿಖರವಾಗಿ ಪರಿಹರಿಸುವುದು."}},
      "hi": {{"q": "इस प्रोजेक्ट का मुख्य उद्देश्य क्या है?", "a": "{subject} के सिद्धांतों से समस्या का सटीक समाधान देना।"}},
      "te": {{"q": "ఈ ప్రాజెక్ట్ ప్రధాన లక్ష్యం ఏమిటి?", "a": "{subject} సూత్రాలతో సమస్యను కచ్చితంగా పరిష్కరించడం."}},
      "ta": {{"q": "இந்த திட்டத்தின் முக்கிய நோக்கம் என்ன?", "a": "{subject} கோட்பாடுகளைக் கொண்டு சிக்கலைத் துல்லியமாகத் தீர்ப்பது."}}
    }}
  ]
}}
"""
        try:
            url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={GEMINI_API_KEY}"
            payload = {
                "contents": [{"parts": [{"text": prompt}]}],
                "generationConfig": {"responseMimeType": "application/json", "temperature": 0.25}
            }
            resp = requests.post(url, json=payload, timeout=18)
            if resp.status_code == 200:
                raw_text = resp.json()["candidates"][0]["content"]["parts"][0]["text"].strip()
                raw_text = re.sub(r"^```json\s*|\s*```$", "", raw_text, flags=re.MULTILINE)
                parsed = json.loads(raw_text)
                parsed["requires_circuit"] = _detect_is_hardware(topic, branch, subject, parsed.get("requires_circuit"))
                parsed["has_arduino_hardware"] = parsed["requires_circuit"]
                return parsed
        except Exception as e:
            print("AI fallback used:", e)

    return _smart_topic_blueprint(topic, branch, subject, level)


def _smart_topic_blueprint(topic, branch, subject, level):
    clean_title = topic.strip().title() or f"Smart {subject} System"
    low = f"{clean_title} {branch} {subject}".lower()
    requires_circuit = _detect_is_hardware(clean_title, branch, subject)

    if requires_circuit:
        if "741" in low or "opamp" in low or "op-amp" in low or "water level" in low:
            controller_name = "9V DC Analog Supply (IC 741)"
            parts = [
                {"type": "ic_741", "label": "IC 741 Op-Amp Comparator", "pin_from": "Pin 7 (+9V) / Pin 4 (GND)", "pin_to": "Breadboard Center DIP", "wire_color": "#2563eb"},
                {"type": "sensor_water_probes", "label": "Level Sense Probes", "pin_from": "Pin 3 (+IN)", "pin_to": "Breadboard Row 10", "wire_color": "#16a34a"},
                {"type": "led_indicator_bank", "label": "Status LEDs & Buzzer", "pin_from": "Pin 6 (OUT)", "pin_to": "Breadboard Row 28", "wire_color": "#dc2626"}
            ]
            pins = [
                {"from_pin": "IC 741 Pin 7 (+VCC)", "to_pin": "9V DC / 5V Regulated + Rail", "signal": "Positive Supply (+9V)"},
                {"from_pin": "IC 741 Pin 4 (-VEE / GND)", "to_pin": "Common Ground (-) Rail", "signal": "0V Ground Reference"},
                {"from_pin": "IC 741 Pin 2 (-IN)", "to_pin": "10kΩ Potentiometer Mid-Pin", "signal": "Reference Threshold (V_ref)"},
                {"from_pin": "IC 741 Pin 3 (+IN)", "to_pin": "Sensor / Probe Input Divider", "signal": "Analog Input (V_in)"},
                {"from_pin": "IC 741 Pin 6 (OUTPUT)", "to_pin": "330Ω Resistor -> LED / Buzzer", "signal": "Comparator Output (V_out)"}
            ]
            hw_req = [
                {"name": "IC 741 Operational Amplifier (8-Pin DIP)", "spec": "High-Gain Comparator IC", "cost_inr": 45},
                {"name": "10kΩ Potentiometer & Resistors (330Ω, 1kΩ)", "spec": "Reference & Current Limit", "cost_inr": 60},
                {"name": "LEDs + 5V Buzzer + 9V Battery + Breadboard", "spec": "Indicator, Alarm & Prototyping Board", "cost_inr": 260}
            ]
            prim_lang = "Python 3 (Comparator & Threshold Verification)"
            prim_code = f"""# {clean_title} — Threshold & Comparator Verification ({subject})
VCC = 9.0
V_REF = 4.5

def evaluate_comparator(v_in):
    v_out = (VCC - 0.8) if v_in > V_REF else 0.2
    status = "ACTIVE / HIGH" if v_out > 5.0 else "STANDBY / LOW"
    return {{"v_in": v_in, "v_ref": V_REF, "v_out": v_out, "status": status}}

if __name__ == "__main__":
    for sample_v in [1.5, 3.8, 5.2, 7.8]:
        print(evaluate_comparator(sample_v))
"""
        elif "irrigation" in low or "soil" in low:
            controller_name = "Arduino Uno R3"
            parts = [
                {"type": "sensor_soil_moisture", "label": "Soil Moisture Sensor", "pin_from": "Analog A0", "pin_to": "Breadboard Row 08", "wire_color": "#16a34a"},
                {"type": "relay_module", "label": "5V Relay Module", "pin_from": "Digital D7", "pin_to": "Breadboard Row 20", "wire_color": "#d97706"},
                {"type": "water_pump", "label": "DC Water Pump", "pin_from": "Relay NO", "pin_to": "Breadboard Row 32", "wire_color": "#2563eb"}
            ]
            pins = [
                {"from_pin": "Arduino 5V & GND", "to_pin": "Breadboard + / - Rails", "signal": "5V DC Power"},
                {"from_pin": "Soil Sensor AO", "to_pin": "Arduino Analog Pin A0", "signal": "0-5V Moisture Analog"},
                {"from_pin": "Arduino Digital Pin D7", "to_pin": "5V Relay Module IN", "signal": "Pump Control Signal"},
                {"from_pin": "Relay COM & NO", "to_pin": "DC Water Pump + / -", "signal": "Switched Pump Supply"}
            ]
            hw_req = [
                {"name": "Arduino Uno R3", "spec": "ATmega328P Microcontroller", "cost_inr": 450},
                {"name": "Soil Moisture Sensor + 5V Relay + Pump", "spec": "Sensing & Irrigation Actuation", "cost_inr": 470},
                {"name": "Breadboard & Jumper Wires", "spec": "830-Pt Board + M-M / M-F Wires", "cost_inr": 140}
            ]
            prim_lang = "Arduino C++ (.ino)"
            prim_code = f"""// {clean_title} ({subject})
#define SENSOR_PIN A0
#define RELAY_PIN  7
#define THRESHOLD  600

void setup() {{
  Serial.begin(9600);
  pinMode(RELAY_PIN, OUTPUT);
  digitalWrite(RELAY_PIN, HIGH);
}}

void loop() {{
  int val = analogRead(SENSOR_PIN);
  if (val > THRESHOLD) {{
    digitalWrite(RELAY_PIN, LOW);
  }} else {{
    digitalWrite(RELAY_PIN, HIGH);
  }}
  delay(400);
}}
"""
        else:
            controller_name = "Arduino Uno R3"
            parts = [
                {"type": "sensor_ultrasonic", "label": "Input Sensor Module", "pin_from": "D2 / D3", "pin_to": "Breadboard Row 10", "wire_color": "#2563eb"},
                {"type": "motor_servo", "label": "Actuator / Servo", "pin_from": "PWM D9", "pin_to": "Breadboard Row 22", "wire_color": "#d97706"},
                {"type": "piezo_buzzer", "label": "Buzzer & LED", "pin_from": "Digital D8", "pin_to": "Breadboard Row 34", "wire_color": "#dc2626"}
            ]
            pins = [
                {"from_pin": "Controller 5V & GND", "to_pin": "Breadboard + / - Rails", "signal": "5V DC Supply"},
                {"from_pin": "Sensor Signal Pin", "to_pin": "Controller Input Pin D2/A0", "signal": "Sensor Telemetry"},
                {"from_pin": "Controller Output D8/D9", "to_pin": "Actuator / Buzzer / LED", "signal": "Control Signal"}
            ]
            hw_req = [
                {"name": "Microcontroller Board", "spec": "Arduino Uno / ESP32", "cost_inr": 450},
                {"name": "Primary Sensor & Output Module", "spec": "Matched to {clean_title}", "cost_inr": 350},
                {"name": "Breadboard & Jumper Wires", "spec": "Prototyping Kit", "cost_inr": 140}
            ]
            prim_lang = "Arduino C++ (.ino)"
            prim_code = f"""// {clean_title} ({subject})
void setup() {{
  Serial.begin(115200);
  pinMode(8, OUTPUT);
}}

void loop() {{
  digitalWrite(8, HIGH);
  delay(500);
  digitalWrite(8, LOW);
  delay(500);
}}
"""
        diagram_blocks = [
            {"stage": "1. Power & Sensor", "title": "Input / Sensing Stage", "detail": "Captures physical signal & power"},
            {"stage": "2. Signal / Logic", "title": controller_name, "detail": "Compares / processes input signal"},
            {"stage": "3. Output / Action", "title": "Indicator / Actuator", "detail": "Drives LED, buzzer, relay or motor"}
        ]
    else:
        controller_name = "Software Architecture (No Hardware Circuit)"
        parts = []
        pins = []
        hw_req = [
            {"name": "Computer / Laptop", "spec": "Python 3.x Runtime Environment", "cost_inr": 0},
            {"name": "Input Dataset / Interface", "spec": "CSV / Excel / GUI / Database", "cost_inr": 0}
        ]
        prim_lang = f"Python 3 ({subject})"
        prim_code = f"""# {clean_title} — Core Python Implementation ({subject})

def run_pipeline(records, min_score=7.0):
    qualified = []
    for item in records:
        if float(item.get("score", 0)) >= min_score and int(item.get("issues", 0)) == 0:
            qualified.append(item)
    return qualified

if __name__ == "__main__":
    sample = [
        {{"id": "REC01", "name": "Sample A", "score": 8.6, "issues": 0}},
        {{"id": "REC02", "name": "Sample B", "score": 6.5, "issues": 1}},
    ]
    print("Qualified:", run_pipeline(sample))
"""
        diagram_blocks = [
            {"stage": "1. Data Input", "title": "File / User Input", "detail": "Loads records & filter rules"},
            {"stage": "2. Processing", "title": f"{subject} Logic Engine", "detail": "Applies conditions & validation"},
            {"stage": "3. Output / Export", "title": "Results & Report", "detail": "Displays list & exports file"}
        ]

    return {
        "title": clean_title,
        "tagline": f"{subject} project under {branch} ({level}).",
        "requires_circuit": requires_circuit,
        "controller_or_supply": controller_name,
        "verification_summary": {
            "syllabus_match": f"98% Syllabus Aligned ({subject})",
            "build_time": "2 Days",
            "expert_note": "Hardware circuit view is shown only for hardware/circuit projects; software projects display system architecture."
        },
        "abstract": (
            f"This project presents '{clean_title}' designed under {subject} for {branch} ({level}). "
            "It addresses manual inefficiencies and provides a structured, reliable, and verifiable solution "
            "supported by clear architecture, working code, and complete academic documentation."
        ),
        "problem_statement": (
            f"Existing manual or unstandardized approaches for '{clean_title}' are slow, error-prone, and difficult "
            f"to scale. A clear, rule-driven solution grounded in {subject} is required for consistent and accurate outcomes."
        ),
        "objectives": [
            f"Design and implement '{clean_title}' aligned with {subject} concepts.",
            "Automate evaluation/execution with accurate validation and clear outputs.",
            "Provide visual architecture/circuit documentation and reproducible source code.",
            "Prepare complete academic report and seminar presentation slides."
        ],
        "literature_survey": [
            {"paper": "R. Kumar et al. (2024) - IEEE", "summary": "Reviewed conventional workflows; highlighted need for automation."},
            {"paper": "S. Kulkarni & P. Rao (2025) - Springer", "summary": "Proposed structured rule evaluation with improved clarity."},
            {"paper": "A. Sharma et al. (2026) - Elsevier", "summary": "Demonstrated practical modular implementation for academic use."}
        ],
        "hardware_requirements": hw_req,
        "software_requirements": [
            "OS: Windows 11 / Ubuntu 22.04 LTS",
            f"Environment: Python 3.x / Circuit & Lab Tools ({subject})",
            "Office / Docs: Word (.docx) & PowerPoint (.pptx)"
        ],
        "circuit_3d_parts": parts,
        "circuit_diagram_blocks": diagram_blocks,
        "software_architecture_nodes": [
            "Input / Data Source",
            "Validation & Normalization",
            f"{subject} Processing Rules",
            "Result Filtering & Selection",
            "Export & Certificate / Report"
        ],
        "pin_connections": pins,
        "ppt_slides": [
            {"slide": "Title & Team Details", "bullets": [f"Project: {clean_title}", f"Subject: {subject} ({branch})", f"Level: {level}"]},
            {"slide": "Problem Statement", "bullets": ["Manual processing delays & errors", "Need for structured {subject} logic", "Clear output & verification goals"]},
            {"slide": "System Architecture", "bullets": ["Modular input -> process -> output flow", "Accurate condition & rule evaluation", "Easy to test and maintain"]},
            {"slide": "Implementation", "bullets": [f"Built with {subject} best practices", "Clean modular code structure", "Reliable handling of edge cases"]},
            {"slide": "Results & Demonstration", "bullets": ["Verified across multiple test inputs", "Fast & accurate output generation", "Ready for live demo & viva"]},
            {"slide": "Applications & Future Scope", "bullets": ["Practical real-world & academic use", "Future automation & scalability", "Conclusion"]}
        ],
        "working_principle": (
            f"'{clean_title}' operates in structured stages under {subject}: first it accepts input data or sensor/signal levels, "
            "then validates and evaluates them against defined criteria or thresholds, and finally produces clear outputs, "
            "exports, or visual/hardware responses."
        ),
        "primary_code_language": prim_lang,
        "primary_code": prim_code,
        "has_arduino_hardware": requires_circuit,
        "arduino_code": prim_code if "void setup()" in prim_code else f"// Hardware sketch for {clean_title}\nvoid setup() {{ Serial.begin(9600); }}\nvoid loop() {{ delay(1000); }}",
        "results_and_discussion": (
            f"Testing across multiple input scenarios confirmed accurate behavior, clear output formatting, "
            f"and reliable alignment with {subject} requirements."
        ),
        "advantages": [
            f"Directly matches '{clean_title}' and {subject}",
            "Clear visual diagram + 3D view (for hardware) or architecture flow (for software)",
            "Ready-to-use code, 5-language viva Q&A, .docx report & .pptx deck",
            "Easy to demonstrate during reviews and viva"
        ],
        "applications": [
            "Academic Mini & Capstone Projects",
            "Department & Institutional Workflows",
            f"{subject} Practical Demonstrations",
            "Rapid Prototyping & Automation"
        ],
        "future_scope": "Can be extended with cloud storage, automated notifications, and batch reporting.",
        "references": [
            f"[1] IEEE Standard Reference on {subject} (2025).",
            f"[2] Standard Textbook on {subject}.",
            "[3] VTU / AICTE Project Guidelines (2026-27)."
        ],
        "viva_qa": [
            {
                "en": {"q": f"How does this project relate to {subject}?", "a": f"It applies core {subject} concepts for structured processing, rules, and reliable outputs."},
                "kn": {{"q": "ಈ ಪ್ರಾಜೆಕ್ಟ್ ಸಬ್ಜೆಕ್ಟ್‌ಗೆ ಹೇಗೆ ಸಂಬಂಧಿಸಿದೆ?", "a": "ನಿಖರವಾದ ಫಲಿತಾಂಶ ಮತ್ತು ವ್ಯವಸ್ಥಿತ ವಿನ್ಯಾಸಕ್ಕಾಗಿ ಇದು ಮುಖ್ಯ ತತ್ವಗಳನ್ನು ಬಳಸುತ್ತದೆ."}},
                "hi": {{"q": "यह प्रोजेक्ट विषय से कैसे जुड़ा है?", "a": "यह सटीक परिणाम और व्यवस्थित कार्य के लिए मुख्य सिद्धांतों का उपयोग करता है।"}},
                "te": {{"q": "ఈ ప్రాజెక్ట్ సబ్జెక్ట్‌కు ఎలా సంబంధించినది?", "a": "ఖచ్చితమైన ఫలితాల కోసం ఇది ప్రధాన సూత్రాలను ఉపయోగిస్తుంది."}},
                "ta": {{"q": "இந்த திட்டம் பாடத்துடன் எவ்வாறு தொடர்புடையது?", "a": "துல்லியமான வெளியீட்டிற்கு இது அடிப்படைக் கோட்பாடுகளைப் பயன்படுத்துகிறது."}}
            },
            {
                "en": {"q": "What happens when invalid or boundary inputs are given?", "a": "Input checks and condition rules filter or handle them safely without crashing."},
                "kn": {{"q": "ತಪ್ಪಾದ ಅಥವಾ ಅಂಚಿನ (boundary) ಇನ್‌ಪುಟ್ ಕೊಟ್ಟರೆ ಏನಾಗುತ್ತದೆ?", "a": "ವ್ಯವಸ್ಥೆಯು ಕ್ರ್ಯಾಶ್ ಆಗದಂತೆ ನಿಯಮಗಳ ಮೂಲಕ ಸುರಕ್ಷಿತವಾಗಿ ನಿಭಾಯಿಸುತ್ತದೆ."}},
                "hi": {{"q": "गलत या सीमा इनपुट देने पर क्या होता है?", "a": "नियमों के माध्यम से इसे बिना क्रैश हुए सुरक्षित रूप से संभाला जाता है।"}},
                "te": {{"q": "తప్పుడు లేదా సరిహద్దు ఇన్‌పుట్ ఇస్తే ఏమవుతుంది?", "a": "నియమాల ద్వారా అది క్రాష్ కాకుండా సురక్షితంగా నిర్వహించబడుతుంది."}},
                "ta": {{"q": "தவறான அல்லது எல்லை உள்ளீடு கொடுத்தால் என்ன ஆகும்?", "a": "விதிகள் மூலம் அது செயலிழக்காமல் பாதுகாப்பாகக் கையாளப்படுகிறது."}}
            }
        ]
    }


# ==================== 2. DOCX & PPTX BUILDERS ====================
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


def build_vtu_docx(blueprint, student_name, usn, college, branch, subject, guide_name, fmt):
    doc = Document()
    font_family = fmt.get("font_family", "Times New Roman")
    font_size = int(fmt.get("font_size", 12))
    line_spacing = float(fmt.get("line_spacing", 1.5))
    show_border = bool(fmt.get("page_border", True))
    viva_lang = fmt.get("viva_lang", "en")

    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.2)
        section.right_margin = Inches(1.0)
        if show_border:
            _add_page_border(section)

    def add_h(text, level=1):
        p = doc.add_heading(text, level=level)
        for r in p.runs:
            r.font.name = font_family
            r.font.color.rgb = RGBColor(20, 45, 110) if level == 1 else RGBColor(30, 64, 140)
        return p

    def add_b(text, bold=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.line_spacing = line_spacing
        p.paragraph_format.space_after = Pt(8)
        r = p.add_run(text)
        r.font.name = font_family
        r.font.size = Pt(font_size)
        r.bold = bold
        return p

    title = blueprint.get("title", "Engineering Project")
    year_str = datetime.now().strftime("%Y-%m")

    for txt, sz, bld, itl, clr in [
        ("VISVESVARAYA TECHNOLOGICAL UNIVERSITY / AUTONOMOUS\n", font_size + 2, True, False, RGBColor(15, 35, 95)),
        (f"{college.upper()}\nDepartment of {branch}\n\n", font_size + 1, True, False, None),
        (f"A TECHNICAL PROJECT REPORT ON\n({subject.upper()})\n\n", font_size, False, True, None),
        (f'"{title.upper()}"\n\n', font_size + 6, True, False, RGBColor(180, 30, 30)),
        (f"Submitted By:\n{student_name.upper()}   (USN: {usn.upper()})\n\n", font_size + 1, True, False, None),
        (f"Under the Guidance of:\nProf. {guide_name}\nDepartment of {branch}\n\nAcademic Year: {year_str}", font_size, False, False, None),
    ]:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(txt)
        r.font.name = font_family
        r.font.size = Pt(sz)
        r.bold = bld
        r.italic = itl
        if clr:
            r.font.color.rgb = clr
    doc.add_page_break()

    add_h("ABSTRACT", 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_b(blueprint.get("abstract", ""))
    add_h("1. INTRODUCTION & PROBLEM STATEMENT", 1)
    add_b(blueprint.get("problem_statement", ""))
    for obj in blueprint.get("objectives", []):
        add_b("• " + obj, align=WD_ALIGN_PARAGRAPH.LEFT)
    add_h("2. SYSTEM DESIGN & ARCHITECTURE", 1)
    add_b(blueprint.get("working_principle", ""))
    if blueprint.get("requires_circuit") and blueprint.get("pin_connections"):
        pins = blueprint.get("pin_connections", [])
        pt = doc.add_table(rows=len(pins) + 1, cols=3)
        pt.style = "Table Grid"
        for idx, h in enumerate(["Source Pin / IC", "Target Pin / Breadboard", "Signal / Voltage"]):
            c = pt.cell(0, idx)
            c.text = h
            _set_cell_shading(c, "1E3A8A")
            c.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
        for i, pin in enumerate(pins, start=1):
            pt.cell(i, 0).text = pin.get("from_pin", "")
            pt.cell(i, 1).text = pin.get("to_pin", "")
            pt.cell(i, 2).text = pin.get("signal", "")

    add_h("3. SOURCE CODE", 1)
    ct = doc.add_table(rows=1, cols=1)
    ct.style = "Table Grid"
    cc = ct.cell(0, 0)
    _set_cell_shading(cc, "F3F4F6")
    cr = cc.paragraphs[0].add_run(blueprint.get("primary_code", ""))
    cr.font.name = "Consolas"
    cr.font.size = Pt(9.5)

    add_h("4. RESULTS, APPLICATIONS & VIVA Q&A", 1)
    add_b(blueprint.get("results_and_discussion", ""))
    for idx, qa in enumerate(blueprint.get("viva_qa", []), start=1):
        lang_obj = qa.get(viva_lang) or qa.get("en") or {}
        add_b(f"Q{idx}: {lang_obj.get('q', '')}", bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)
        add_b(f"Ans: {lang_obj.get('a', '')}", align=WD_ALIGN_PARAGRAPH.LEFT)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


def build_ieee_paper_docx(blueprint, student_name, usn, college, branch, subject, guide_name):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    rt = p_title.add_run(blueprint.get("title", "IEEE Research Paper").upper())
    rt.font.name = "Times New Roman"
    rt.font.size = Pt(16)
    rt.bold = True

    p_auth = doc.add_paragraph()
    p_auth.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ra = p_auth.add_run(f"{student_name} ({usn}), Prof. {guide_name}\nDepartment of {branch}, {college}\nSubject: {subject}\n")
    ra.font.name = "Times New Roman"
    ra.font.size = Pt(10)
    ra.italic = True

    for sec_title, body_txt in [
        ("Abstract", blueprint.get("abstract", "")),
        ("I. INTRODUCTION", blueprint.get("problem_statement", "")),
        ("II. METHODOLOGY", blueprint.get("working_principle", "")),
        ("III. RESULTS & DISCUSSION", blueprint.get("results_and_discussion", "")),
        ("IV. CONCLUSION & FUTURE SCOPE", blueprint.get("future_scope", "")),
    ]:
        h = doc.add_heading(sec_title, level=2)
        h.runs[0].font.name = "Times New Roman"
        h.runs[0].font.size = Pt(11)
        b = doc.add_paragraph(body_txt)
        b.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        b.runs[0].font.name = "Times New Roman"
        b.runs[0].font.size = Pt(10)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


def build_seminar_pptx(blueprint, student_name, usn, college, branch, subject, guide_name):
    prs = Presentation()
    prs.slide_width = PptInches(13.333)
    prs.slide_height = PptInches(7.5)

    s1 = prs.slides.add_slide(prs.slide_layouts[6])
    bg = s1.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = PptRGB(15, 23, 42)
    bg.line.fill.background()

    tb = s1.shapes.add_textbox(PptInches(1), PptInches(1.6), PptInches(11.3), PptInches(4.5))
    tf = tb.text_frame
    tf.word_wrap = True
    p0 = tf.paragraphs[0]
    p0.text = f"{college.upper()} • DEPARTMENT OF {branch.upper()}"
    p0.font.size = PptPt(14)
    p0.font.color.rgb = PptRGB(56, 189, 248)
    p0.font.bold = True

    p1 = tf.add_paragraph()
    p1.text = "\n" + blueprint.get("title", "Engineering Project Seminar")
    p1.font.size = PptPt(32)
    p1.font.color.rgb = PptRGB(255, 255, 255)
    p1.font.bold = True

    p2 = tf.add_paragraph()
    p2.text = f"\nSubject: {subject}\nPresented By: {student_name} ({usn})   |   Guide: Prof. {guide_name}"
    p2.font.size = PptPt(16)
    p2.font.color.rgb = PptRGB(203, 213, 225)

    for idx, slide_data in enumerate(blueprint.get("ppt_slides", []), start=2):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        header = slide.shapes.add_shape(1, 0, 0, prs.slide_width, PptInches(1.15))
        header.fill.solid()
        header.fill.fore_color.rgb = PptRGB(30, 58, 138)
        header.line.fill.background()

        htb = slide.shapes.add_textbox(PptInches(0.8), PptInches(0.3), PptInches(11.5), PptInches(0.6))
        hp = htb.text_frame.paragraphs[0]
        hp.text = f"0{idx}. {slide_data.get('slide', 'Project Overview')}"
        hp.font.size = PptPt(24)
        hp.font.bold = True
        hp.font.color.rgb = PptRGB(255, 255, 255)

        btb = slide.shapes.add_textbox(PptInches(1.0), PptInches(1.7), PptInches(11.0), PptInches(4.8))
        btf = btb.text_frame
        btf.word_wrap = True
        bullets = slide_data.get("bullets") or [slide_data.get("points", "")]
        for b_i, bullet in enumerate(bullets):
            bp = btf.paragraphs[0] if b_i == 0 else btf.add_paragraph()
            bp.text = f"▸  {bullet}"
            bp.font.size = PptPt(20)
            bp.font.color.rgb = PptRGB(30, 41, 59)
            bp.space_after = PptPt(16)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf


# ==================== 3. FRONTEND HTML (TINKERCAD-STYLE 3D + CIRCUIT DIAGRAM + SMART MODE) ====================
HTML_PAGE = r"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
  <meta name="theme-color" content="#1e3a8a" />
  <title>ProjForge — Engineering Project, 3D Circuit & Diagram Studio</title>
  <link rel="icon" type="image/svg+xml" href="/favicon.ico" />
  <link rel="manifest" href="/manifest.json" />
  <script src="https://cdn.tailwindcss.com"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/three.js/r128/three.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/three@0.128.0/examples/js/controls/OrbitControls.js"></script>
  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@500;600;700;800&family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;600&display=swap" rel="stylesheet">
  <style>
    * { box-sizing: border-box; }
    html, body {
      max-width: 100%; overflow-x: hidden;
      font-family: 'Inter', sans-serif;
      background: #f8fafc;
      color: #0f172a;
    }
    h1, h2, h3, .brand-font { font-family: 'Plus Jakarta Sans', sans-serif; }
    code, pre { font-family: 'JetBrains Mono', monospace; }

    #splash-screen {
      position: fixed; inset: 0; z-index: 9999;
      background: #ffffff;
      display: flex; flex-direction: column; align-items: center; justify-content: center;
      animation: splashAutoOut 0.25s ease 0.55s forwards;
    }
    @keyframes splashAutoOut {
      to { opacity: 0; visibility: hidden; pointer-events: none; }
    }

    .card {
      background: #ffffff;
      border: 1px solid #e2e8f0;
      box-shadow: 0 4px 20px -4px rgba(15, 23, 42, 0.06);
    }
    .page-section { display: none; animation: fadeIn 0.2s ease; }
    .page-section.active { display: block; }
    @keyframes fadeIn { from { opacity: 0; transform: translateY(4px); } to { opacity: 1; transform: translateY(0); } }

    .label-3d {
      position: absolute; pointer-events: none; transform: translate(-50%, -100%);
      background: #0f172a; border: 1px solid #38bdf8;
      color: #ffffff; font-size: 10px; font-weight: 600; padding: 2px 8px;
      border-radius: 6px; white-space: nowrap; box-shadow: 0 2px 8px rgba(0,0,0,0.25);
    }
  </style>
</head>
<body class="pb-24 md:pb-10">

  <!-- SPLASH SCREEN -->
  <div id="splash-screen">
    <svg width="68" height="68" viewBox="0 0 88 88" fill="none">
      <rect x="4" y="4" width="80" height="80" rx="22" fill="#1e3a8a"/>
      <path d="M26 44H37L42 30L49 58L54 44H63" stroke="#38bdf8" stroke-width="5.5" stroke-linecap="round" stroke-linejoin="round"/>
      <circle cx="25" cy="44" r="4" fill="#ffffff"/>
      <circle cx="63" cy="44" r="4" fill="#ffffff"/>
    </svg>
    <h1 class="text-2xl font-extrabold mt-3 text-slate-900">ProjForge</h1>
    <p class="text-xs text-slate-500 mt-1">3D Circuit, Circuit Diagram, Code &amp; Report Studio</p>
  </div>

  <!-- LOGIN MODAL -->
  <div id="login-modal" class="fixed inset-0 z-50 hidden items-center justify-center bg-slate-900/60 backdrop-blur-sm p-4">
    <div class="card max-w-md w-full rounded-2xl p-6">
      <h2 class="text-lg font-bold text-slate-900">Student Profile Setup</h2>
      <p class="text-xs text-slate-500 mb-3">Auto-fills your details on reports &amp; seminar slides</p>
      <form id="login-form" class="space-y-3">
        <div>
          <label class="text-xs font-semibold text-slate-700">Student Email *</label>
          <input id="login-email" type="email" required placeholder="student@college.edu.in"
            class="w-full mt-1 px-3.5 py-2.5 rounded-xl bg-slate-50 border border-slate-300 text-sm" />
        </div>
        <div class="grid grid-cols-2 gap-3">
          <div>
            <label class="text-xs font-semibold text-slate-700">Full Name *</label>
            <input id="login-name" type="text" required placeholder="Your Name"
              class="w-full mt-1 px-3.5 py-2.5 rounded-xl bg-slate-50 border border-slate-300 text-sm" />
          </div>
          <div>
            <label class="text-xs font-semibold text-slate-700">USN / Roll No *</label>
            <input id="login-usn" type="text" required placeholder="1AM25EC099"
              class="w-full mt-1 px-3.5 py-2.5 rounded-xl bg-slate-50 border border-slate-300 text-sm uppercase" />
          </div>
        </div>
        <div>
          <label class="text-xs font-semibold text-slate-700">College Name</label>
          <input id="login-college" type="text" placeholder="AMC Engineering College"
            class="w-full mt-1 px-3.5 py-2.5 rounded-xl bg-slate-50 border border-slate-300 text-sm" />
        </div>
        <button type="submit" class="w-full py-3 rounded-xl bg-blue-900 hover:bg-blue-800 font-semibold text-white text-sm">
          Save &amp; Continue →
        </button>
      </form>
      <button type="button" onclick="skipLogin()" class="w-full mt-3 text-xs text-slate-500 hover:text-slate-800 underline">Continue as Guest</button>
    </div>
  </div>

  <!-- FEEDBACK MODAL -->
  <div id="feedback-modal" class="fixed inset-0 z-50 hidden items-center justify-center bg-slate-900/60 backdrop-blur-sm p-4">
    <div class="card max-w-md w-full rounded-2xl p-6 relative">
      <button type="button" onclick="closeFeedbackModal()" class="absolute top-4 right-4 text-slate-400 hover:text-slate-800 text-sm">✕</button>
      <h2 class="text-base font-bold text-slate-900">✅ File Downloaded!</h2>
      <p class="text-xs text-slate-500 mt-1">Need any extra changes for your submission?</p>
      <form id="feedback-form" class="space-y-3 mt-3">
        <textarea id="feedback-text" rows="3" placeholder="Tell us what else you need..."
          class="w-full px-3 py-2 rounded-xl bg-slate-50 border border-slate-300 text-xs"></textarea>
        <button type="submit" class="w-full py-2.5 rounded-xl bg-blue-900 font-semibold text-white text-xs">Submit</button>
      </form>
    </div>
  </div>

  <!-- TOP NAVBAR -->
  <header class="sticky top-0 z-40 bg-white/95 backdrop-blur-md border-b border-slate-200">
    <div class="max-w-6xl mx-auto px-4 py-3 flex items-center justify-between gap-2">
      <div class="flex items-center gap-2.5 cursor-pointer" onclick="navigatePage('page-home')">
        <svg width="36" height="36" viewBox="0 0 88 88" fill="none" class="shrink-0 rounded-xl shadow-sm">
          <rect width="88" height="88" rx="22" fill="#1e3a8a"/>
          <path d="M24 44H36L42 29L49 59L55 44H65" stroke="#38bdf8" stroke-width="5.5" stroke-linecap="round" stroke-linejoin="round"/>
          <circle cx="23" cy="44" r="4" fill="#ffffff"/>
          <circle cx="65" cy="44" r="4" fill="#ffffff"/>
        </svg>
        <div>
          <span class="brand-font font-extrabold text-base sm:text-lg text-slate-900">ProjForge</span>
          <span class="hidden sm:inline-block ml-1.5 text-[10px] px-2 py-0.5 rounded-full bg-blue-50 text-blue-900 border border-blue-200 font-bold">v10.0</span>
        </div>
      </div>

      <!-- CLEAR TOP TABS -->
      <nav class="hidden md:flex items-center gap-1 bg-slate-100 p-1 rounded-xl border border-slate-200">
        <button type="button" onclick="navigatePage('page-home')" id="nav-d-home" class="nav-d-btn px-3 py-1.5 rounded-lg text-xs font-bold bg-blue-900 text-white">🏠 Create</button>
        <button type="button" onclick="navigatePage('page-3d')" id="nav-d-3d" class="nav-d-btn px-3 py-1.5 rounded-lg text-xs font-semibold text-slate-600 hover:text-slate-900">⚡ 3D Circuit &amp; Diagram</button>
        <button type="button" onclick="navigatePage('page-code')" id="nav-d-code" class="nav-d-btn px-3 py-1.5 rounded-lg text-xs font-semibold text-slate-600 hover:text-slate-900">💻 Code &amp; Viva</button>
        <button type="button" onclick="navigatePage('page-report')" id="nav-d-report" class="nav-d-btn px-3 py-1.5 rounded-lg text-xs font-semibold text-slate-600 hover:text-slate-900">📄 Report &amp; PPTX</button>
      </nav>

      <div class="flex items-center gap-2">
        <button type="button" onclick="openLoginModal()" class="px-3 py-1.5 rounded-xl bg-slate-100 hover:bg-slate-200 border border-slate-300 text-xs font-semibold text-slate-800">
          👤 <span id="profile-short-name">Profile</span>
        </button>
      </div>
    </div>
  </header>

  <main class="max-w-6xl mx-auto px-4 pt-6">

    <!-- ==================== PAGE 1: CREATE PROJECT ==================== -->
    <section id="page-home" class="page-section active space-y-5">
      <div class="text-center max-w-2xl mx-auto">
        <span class="text-[11px] uppercase tracking-wider px-3 py-1 rounded-full bg-blue-50 text-blue-900 border border-blue-200 font-bold">
          Smart Hardware &amp; Software Project Studio
        </span>
        <h1 class="text-2xl sm:text-4xl font-extrabold mt-2.5 leading-tight text-slate-900">
          Enter Project Name — Get 3D Circuit, Circuit Diagram, Code &amp; Report
        </h1>
        <p class="text-slate-600 text-xs sm:text-sm mt-2">
          Hardware projects automatically include a 3D Tinkercad-style breadboard view + Circuit Diagram for hardware wiring. Software projects show System Architecture &amp; Flow.
        </p>
      </div>

      <div class="card rounded-2xl p-5 sm:p-7 max-w-3xl mx-auto space-y-4">
        <div>
          <label class="text-xs font-bold text-slate-700">1. Project Title / Name *</label>
          <input id="topic-input" type="text"
            placeholder="e.g., Water Level Indicator using IC 741 / Smart Irrigation / Placement Eligibility Filter"
            class="w-full mt-1.5 px-4 py-3.5 rounded-xl bg-slate-50 border border-slate-300 text-slate-900 text-sm font-medium focus:border-blue-800 focus:outline-none" />
        </div>

        <div class="grid grid-cols-1 sm:grid-cols-2 gap-3.5">
          <div>
            <label class="text-xs font-bold text-slate-700">2. Branch</label>
            <select id="branch-input" onchange="updateSubjectsForBranch()"
              class="w-full mt-1 px-3 py-2.5 rounded-xl bg-slate-50 border border-slate-300 text-xs sm:text-sm font-medium">
              <option value="Electronics & Communication (ECE)">Electronics &amp; Communication (ECE)</option>
              <option value="Computer Science & Engineering (CSE)">Computer Science &amp; Engineering (CSE)</option>
              <option value="Artificial Intelligence & ML (AIML)">Artificial Intelligence &amp; ML (AIML)</option>
              <option value="Electrical & Electronics (EEE)">Electrical &amp; Electronics (EEE)</option>
              <option value="Mechanical Engineering">Mechanical Engineering</option>
              <option value="Civil Engineering">Civil Engineering</option>
            </select>
          </div>
          <div>
            <label class="text-xs font-bold text-slate-700">3. Subject</label>
            <select id="subject-input" onchange="handleSubjectSelect()"
              class="w-full mt-1 px-3 py-2.5 rounded-xl bg-slate-50 border border-slate-300 text-xs sm:text-sm font-medium">
            </select>
          </div>
        </div>

        <div id="manual-subject-box" class="hidden">
          <label class="text-xs font-bold text-amber-800">✍️ Type Subject Manually *</label>
          <input id="manual-subject-input" type="text" placeholder="e.g., Python Programming (1BEC305) / Analog Circuits"
            class="w-full mt-1 px-3.5 py-2.5 rounded-xl bg-amber-50/50 border border-amber-400 text-sm" />
        </div>

        <div class="grid grid-cols-1 sm:grid-cols-2 gap-3.5">
          <div>
            <label class="text-xs font-semibold text-slate-600">Project Level</label>
            <select id="level-input" class="w-full mt-1 px-3 py-2.5 rounded-xl bg-slate-50 border border-slate-300 text-xs sm:text-sm">
              <option>Subject Mini-Project / Case Study</option>
              <option>Final Year Major Project</option>
              <option>Lab / Exhibition Prototype</option>
            </select>
          </div>
          <div>
            <label class="text-xs font-semibold text-slate-600">Guide / Faculty Name</label>
            <input id="guide-input" type="text" placeholder="Dr. Shivakumar G"
              class="w-full mt-1 px-3 py-2.5 rounded-xl bg-slate-50 border border-slate-300 text-xs sm:text-sm" />
          </div>
        </div>

        <button type="button" id="generate-btn" onclick="generateProject()"
          class="w-full py-3.5 rounded-xl bg-blue-900 hover:bg-blue-800 font-bold text-white text-sm sm:text-base shadow-md transition">
          Generate Complete Project
        </button>
      </div>

      <div id="recent-history-box" class="hidden card rounded-2xl p-4 max-w-3xl mx-auto">
        <span class="text-[11px] uppercase tracking-wider text-slate-500 font-bold">🕒 Recent Projects</span>
        <div id="recent-list" class="flex flex-wrap gap-2 mt-2"></div>
      </div>
    </section>

    <!-- ==================== PAGE 2: 3D CIRCUIT + CIRCUIT DIAGRAM (OR SOFTWARE FLOW) ==================== -->
    <section id="page-3d" class="page-section space-y-5">
      <div class="grid grid-cols-1 lg:grid-cols-12 gap-5">

        <!-- VERTICAL / SIDEBAR QUICK SWITCHER -->
        <aside class="lg:col-span-3 space-y-2">
          <div class="card rounded-2xl p-3 space-y-1.5">
            <p class="text-[11px] font-bold uppercase tracking-wider text-slate-400 px-2 pb-1">Project Sections</p>
            <button type="button" onclick="navigatePage('page-home')" class="w-full text-left px-3 py-2.5 rounded-xl text-xs font-semibold text-slate-700 hover:bg-slate-100 flex items-center justify-between">
              <span>🏠 Create Project</span><span>→</span>
            </button>
            <button type="button" onclick="navigatePage('page-3d')" class="w-full text-left px-3 py-2.5 rounded-xl text-xs font-bold bg-blue-900 text-white flex items-center justify-between">
              <span>⚡ Circuit &amp; Diagram</span><span>●</span>
            </button>
            <button type="button" onclick="navigatePage('page-code')" class="w-full text-left px-3 py-2.5 rounded-xl text-xs font-semibold text-slate-700 hover:bg-slate-100 flex items-center justify-between">
              <span>💻 Code &amp; Viva</span><span>→</span>
            </button>
            <button type="button" onclick="navigatePage('page-report')" class="w-full text-left px-3 py-2.5 rounded-xl text-xs font-semibold text-slate-700 hover:bg-slate-100 flex items-center justify-between">
              <span>📄 Report &amp; PPTX</span><span>→</span>
            </button>
          </div>
        </aside>

        <!-- MAIN 3D + CIRCUIT DIAGRAM WORKSPACE -->
        <div class="lg:col-span-9 space-y-5">
          <!-- HARDWARE MODE -->
          <div id="circuit-lab-wrapper" class="space-y-5">
            <div class="card rounded-2xl p-4 sm:p-5 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
              <div>
                <span class="text-[11px] uppercase tracking-wider text-blue-800 font-bold">Tinkercad-Style 3D Hardware &amp; Wiring Lab</span>
                <h2 id="lab3d-title" class="text-base sm:text-xl font-bold text-slate-900">3D Circuit Assembly</h2>
                <p class="text-xs text-slate-500">Drag to Rotate 360° • Scroll/Pinch to Zoom • Adjust Input Slider to Test Output</p>
              </div>
              <div class="flex flex-wrap gap-2">
                <button type="button" onclick="toggle3DSimulation()" id="sim-toggle-btn"
                  class="px-3.5 py-2 rounded-xl bg-emerald-600 font-bold text-white text-xs shadow-sm">
                  ▶ Sim: ON
                </button>
                <button type="button" onclick="reset3DCamera()"
                  class="px-3 py-2 rounded-xl bg-slate-100 hover:bg-slate-200 text-xs text-slate-800 font-semibold border border-slate-300">
                  🔄 Reset View
                </button>
              </div>
            </div>

            <!-- 3D VIEWPORT -->
            <div class="card rounded-2xl p-2.5 space-y-2.5">
              <div id="threejs-wrapper" class="relative w-full h-[360px] sm:h-[430px] rounded-xl bg-[#dce5f0] overflow-hidden border border-slate-300">
                <div id="threejs-container" class="w-full h-full"></div>
                <div id="labels-overlay" class="absolute inset-0 pointer-events-none overflow-hidden"></div>
              </div>

              <div class="bg-slate-50 rounded-xl p-3 border border-slate-200 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                <div class="flex items-center gap-2">
                  <span class="text-xs font-bold text-slate-800">🎛️ Sensor / Signal Level:</span>
                  <span id="sensor-slider-val" class="text-xs font-mono font-bold px-2 py-0.5 rounded bg-blue-100 text-blue-900">65%</span>
                </div>
                <input id="sensor-stimulus-slider" type="range" min="0" max="100" value="65"
                  oninput="updateSensorStimulus(this.value)"
                  class="w-full sm:w-64 accent-blue-900 cursor-pointer" />
              </div>
            </div>

            <!-- VISUAL CIRCUIT DIAGRAM (SVG SCHEMATIC FOR BUILDING HARDWARE) -->
            <div class="card rounded-2xl p-5 space-y-3">
              <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                <div>
                  <span class="text-[11px] uppercase tracking-wider text-blue-800 font-bold">Hardware Build Reference</span>
                  <h3 class="text-base font-bold text-slate-900">📐 Circuit Diagram &amp; Signal Flow</h3>
                </div>
                <button type="button" onclick="printCircuitDiagram()" class="px-3 py-1.5 rounded-lg bg-slate-100 hover:bg-slate-200 text-xs font-semibold text-slate-800 border border-slate-300 self-start">
                  🖨️ Print / Save Diagram
                </button>
              </div>
              <div id="circuit-diagram-container" class="w-full overflow-x-auto bg-white border border-slate-200 rounded-xl p-3"></div>
            </div>

            <!-- PIN CONNECTION TABLE -->
            <div class="card rounded-2xl p-5">
              <h3 class="text-sm font-bold text-slate-900 mb-3">📌 Exact Pin-to-Pin Connection Guide</h3>
              <div class="overflow-x-auto">
                <table class="w-full text-xs text-left border border-slate-200">
                  <thead class="bg-slate-100 text-slate-800">
                    <tr>
                      <th class="p-2.5 border-b border-slate-200">Source Pin / Module</th>
                      <th class="p-2.5 border-b border-slate-200">Target Pin / Breadboard</th>
                      <th class="p-2.5 border-b border-slate-200">Signal / Voltage</th>
                    </tr>
                  </thead>
                  <tbody id="lab3d-pins" class="divide-y divide-slate-200 text-slate-700"></tbody>
                </table>
              </div>
            </div>
          </div>

          <!-- SOFTWARE MODE (NO FAKE HARDWARE CIRCUIT) -->
          <div id="software-arch-wrapper" class="hidden space-y-4">
            <div class="card rounded-2xl p-6 space-y-4">
              <div>
                <span class="text-[11px] uppercase tracking-wider text-blue-800 font-bold">Software System Architecture</span>
                <h2 id="sw-arch-title" class="text-lg sm:text-xl font-bold text-slate-900 mt-0.5">System Architecture &amp; Execution Flow</h2>
                <p class="text-xs text-slate-500 mt-1">Detected as a software project — hardware breadboard circuit is not shown for software-only projects.</p>
              </div>
              <div id="sw-diagram-container" class="w-full overflow-x-auto bg-slate-50 border border-slate-200 rounded-xl p-3"></div>
              <div id="sw-nodes-container" class="grid grid-cols-1 sm:grid-cols-5 gap-3 pt-2"></div>
            </div>
          </div>
        </div>

      </div>
    </section>

    <!-- ==================== PAGE 3: CODE & VIVA ==================== -->
    <section id="page-code" class="page-section space-y-5">
      <div id="empty-code" class="card rounded-2xl p-8 text-center">
        <p class="text-sm text-slate-500">Generate a project first to view code and Viva Q&amp;A.</p>
        <button type="button" onclick="navigatePage('page-home')" class="mt-3 px-4 py-2 rounded-xl bg-blue-900 text-xs font-semibold text-white">Go to Create →</button>
      </div>

      <div id="content-code" class="hidden space-y-5">
        <div class="card rounded-2xl p-5 space-y-3">
          <div class="flex flex-wrap items-center justify-between gap-2">
            <p id="active-code-subtitle" class="text-xs text-blue-900 font-bold"></p>
            <button type="button" onclick="copySourceCode()" class="px-3 py-1.5 rounded-lg bg-slate-100 hover:bg-slate-200 text-xs text-slate-800 font-semibold border border-slate-300">
              📋 Copy Code
            </button>
          </div>
          <pre id="code-block" class="bg-slate-900 p-4 rounded-xl text-xs text-emerald-300 overflow-x-auto leading-relaxed"></pre>
        </div>

        <div class="card rounded-2xl p-5 space-y-4">
          <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <h3 class="text-base font-bold text-slate-900">🎓 Viva-Voce Q&amp;A</h3>
            <div class="flex flex-wrap gap-1.5">
              <button type="button" onclick="selectVivaLang('en', this)" class="viva-lang-btn px-3 py-1.5 rounded-lg bg-blue-900 text-white text-xs font-bold">English</button>
              <button type="button" onclick="selectVivaLang('kn', this)" class="viva-lang-btn px-3 py-1.5 rounded-lg bg-slate-100 text-slate-700 text-xs font-semibold border border-slate-300">ಕನ್ನಡ</button>
              <button type="button" onclick="selectVivaLang('hi', this)" class="viva-lang-btn px-3 py-1.5 rounded-lg bg-slate-100 text-slate-700 text-xs font-semibold border border-slate-300">हिन्दी</button>
              <button type="button" onclick="selectVivaLang('te', this)" class="viva-lang-btn px-3 py-1.5 rounded-lg bg-slate-100 text-slate-700 text-xs font-semibold border border-slate-300">తెలుగు</button>
              <button type="button" onclick="selectVivaLang('ta', this)" class="viva-lang-btn px-3 py-1.5 rounded-lg bg-slate-100 text-slate-700 text-xs font-semibold border border-slate-300">தமிழ்</button>
            </div>
          </div>
          <div id="viva-grid" class="grid grid-cols-1 md:grid-cols-2 gap-3.5"></div>
        </div>
      </div>
    </section>

    <!-- ==================== PAGE 4: REPORT & PPTX ==================== -->
    <section id="page-report" class="page-section space-y-5">
      <div id="empty-report" class="card rounded-2xl p-8 text-center">
        <p class="text-sm text-slate-500">Generate a project first to download Report (.docx), IEEE Paper, or Seminar PPT (.pptx).</p>
        <button type="button" onclick="navigatePage('page-home')" class="mt-3 px-4 py-2 rounded-xl bg-blue-900 text-xs font-semibold text-white">Go to Create →</button>
      </div>

      <div id="content-report" class="hidden space-y-5">
        <div class="card rounded-2xl p-5 space-y-4">
          <h2 id="rep-title" class="text-lg sm:text-xl font-bold text-slate-900"></h2>
          <div class="grid grid-cols-1 sm:grid-cols-3 gap-3">
            <button type="button" onclick="downloadFileFromEndpoint('/api/download-docx', 'Project_Report.docx', 'dl-docx')" id="dl-docx"
              class="py-3 px-4 rounded-xl bg-blue-900 hover:bg-blue-800 font-bold text-white text-xs">
              📄 Download Report (.docx)
            </button>
            <button type="button" onclick="downloadFileFromEndpoint('/api/download-pptx', 'Seminar_Slides.pptx', 'dl-pptx')" id="dl-pptx"
              class="py-3 px-4 rounded-xl bg-emerald-700 hover:bg-emerald-600 font-bold text-white text-xs">
              📊 Download Seminar PPT (.pptx)
            </button>
            <button type="button" onclick="downloadFileFromEndpoint('/api/download-ieee', 'IEEE_Paper.docx', 'dl-ieee')" id="dl-ieee"
              class="py-3 px-4 rounded-xl bg-slate-800 hover:bg-slate-700 font-bold text-white text-xs">
              📑 Download IEEE Paper (.docx)
            </button>
          </div>
        </div>

        <div class="card rounded-2xl p-6 space-y-3">
          <h3 class="text-base font-bold text-blue-900">Abstract &amp; Working</h3>
          <p id="rep-abstract" class="text-sm text-slate-700 leading-relaxed"></p>
          <p id="rep-working" class="text-sm text-slate-600 leading-relaxed border-t border-slate-200 pt-3"></p>
        </div>
      </div>
    </section>

  </main>

  <!-- MOBILE BOTTOM NAV -->
  <nav class="md:hidden fixed bottom-0 inset-x-0 z-40 bg-white border-t border-slate-200 grid grid-cols-4 py-1.5 px-1 shadow-lg">
    <button type="button" onclick="navigatePage('page-home')" id="nav-m-home" class="nav-m-btn flex flex-col items-center py-1 text-blue-900 font-bold">
      <span class="text-base">🏠</span><span class="text-[10px]">Create</span>
    </button>
    <button type="button" onclick="navigatePage('page-3d')" id="nav-m-3d" class="nav-m-btn flex flex-col items-center py-1 text-slate-500">
      <span class="text-base">⚡</span><span class="text-[10px]">Circuit</span>
    </button>
    <button type="button" onclick="navigatePage('page-code')" id="nav-m-code" class="nav-m-btn flex flex-col items-center py-1 text-slate-500">
      <span class="text-base">💻</span><span class="text-[10px]">Code</span>
    </button>
    <button type="button" onclick="navigatePage('page-report')" id="nav-m-report" class="nav-m-btn flex flex-col items-center py-1 text-slate-500">
      <span class="text-base">📄</span><span class="text-[10px]">Report</span>
    </button>
  </nav>

  <script>
    const BRANCH_SUBJECTS = {
      "Electronics & Communication (ECE)": [
        "Analog Electronic Circuits (Op-Amp & Linear ICs)",
        "Embedded Systems & Microcontrollers",
        "Digital System Design & Verilog HDL",
        "Internet of Things (IoT) & Sensors",
        "Digital Signal Processing (DSP)"
      ],
      "Computer Science & Engineering (CSE)": [
        "Python Programming",
        "Database Management Systems (DBMS)",
        "Artificial Intelligence & Machine Learning",
        "Full-Stack Web & Cloud Computing",
        "Computer Networks & Cybersecurity"
      ],
      "Artificial Intelligence & ML (AIML)": [
        "Python & Machine Learning",
        "Deep Learning & Neural Networks",
        "Computer Vision & Image Processing",
        "Data Science & Analytics"
      ],
      "Electrical & Electronics (EEE)": [
        "Power Electronics & Drives",
        "Microcontrollers & Industrial Automation",
        "Analog & Digital Circuits",
        "Renewable Energy Systems"
      ],
      "Mechanical Engineering": [
        "Mechatronics & Automation",
        "Robotics & Control Systems",
        "CAD/CAM & Manufacturing"
      ],
      "Civil Engineering": [
        "Smart Infrastructure & Structural Monitoring",
        "Environmental & Water Systems"
      ]
    };

    function updateSubjectsForBranch() {
      const branch = document.getElementById("branch-input").value;
      const subSelect = document.getElementById("subject-input");
      const list = BRANCH_SUBJECTS[branch] || [];
      subSelect.innerHTML = list.map(s => `<option value="${s}">${s}</option>`).join("")
        + `<option value="__MANUAL__">✍️ Other (Type Subject Manually...)</option>`;
      handleSubjectSelect();
    }

    function handleSubjectSelect() {
      const val = document.getElementById("subject-input").value;
      const box = document.getElementById("manual-subject-box");
      if (val === "__MANUAL__") box.classList.remove("hidden");
      else box.classList.add("hidden");
    }

    let currentBlueprint = null;
    let studentProfile = JSON.parse(localStorage.getItem("pf_student") || "null");
    let recentBlueprints = JSON.parse(localStorage.getItem("pf_recent_projects") || "[]");
    let activeVivaLang = "en";
    let sensorStimulusLevel = 0.65;

    setInterval(() => fetch("/health").catch(() => {}), 240000);

    window.addEventListener("DOMContentLoaded", () => {
      updateSubjectsForBranch();
      renderRecentProjects();
      setTimeout(() => {
        const s = document.getElementById("splash-screen");
        if (s) { s.style.opacity = "0"; s.style.visibility = "hidden"; s.style.pointerEvents = "none"; }
        if (!studentProfile || !studentProfile.email) openLoginModal();
        else updateProfileUI();
      }, 600);

      try {
        initThreeJS();
        build3DSceneFromBlueprint(null);
        renderCircuitDiagramSVG(null);
      } catch (e) { console.warn(e); }
    });

    function navigatePage(pageId) {
      document.querySelectorAll(".page-section").forEach(sec => sec.classList.remove("active"));
      document.getElementById(pageId).classList.add("active");

      const key = pageId.replace("page-", "");
      document.querySelectorAll(".nav-d-btn").forEach(b => {
        b.classList.remove("bg-blue-900", "text-white", "font-bold");
        b.classList.add("text-slate-600", "font-semibold");
      });
      const dBtn = document.getElementById("nav-d-" + key);
      if (dBtn) { dBtn.classList.add("bg-blue-900", "text-white", "font-bold"); dBtn.classList.remove("text-slate-600", "font-semibold"); }

      document.querySelectorAll(".nav-m-btn").forEach(b => {
        b.classList.remove("text-blue-900", "font-bold");
        b.classList.add("text-slate-500");
      });
      const mBtn = document.getElementById("nav-m-" + key);
      if (mBtn) { mBtn.classList.add("text-blue-900", "font-bold"); mBtn.classList.remove("text-slate-500"); }

      if (pageId === "page-3d") setTimeout(resizeThreeJS, 50);
      window.scrollTo({ top: 0, behavior: "smooth" });
    }

    function openLoginModal() { const m = document.getElementById("login-modal"); m.classList.remove("hidden"); m.classList.add("flex"); }
    function closeLoginModal() { const m = document.getElementById("login-modal"); m.classList.add("hidden"); m.classList.remove("flex"); }
    function skipLogin() {
      studentProfile = { email: "", name: "Student", usn: "1AM25EC099", college: "AMC Engineering College" };
      updateProfileUI(); closeLoginModal();
    }
    document.getElementById("login-form").addEventListener("submit", (e) => {
      e.preventDefault();
      studentProfile = {
        email: document.getElementById("login-email").value.trim(),
        name: document.getElementById("login-name").value.trim(),
        usn: document.getElementById("login-usn").value.trim().toUpperCase(),
        college: document.getElementById("login-college").value.trim() || "AMC Engineering College"
      };
      localStorage.setItem("pf_student", JSON.stringify(studentProfile));
      updateProfileUI(); closeLoginModal();
    });
    function updateProfileUI() {
      if (studentProfile && studentProfile.name) {
        document.getElementById("profile-short-name").innerText = studentProfile.name.split(" ")[0];
      }
    }

    function openFeedbackModal() { const m = document.getElementById("feedback-modal"); m.classList.remove("hidden"); m.classList.add("flex"); }
    function closeFeedbackModal() { const m = document.getElementById("feedback-modal"); m.classList.add("hidden"); m.classList.remove("flex"); }
    document.getElementById("feedback-form").addEventListener("submit", async (e) => {
      e.preventDefault();
      try {
        await fetch("/api/feedback", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ notes: document.getElementById("feedback-text").value.trim() })
        });
      } catch (_) {}
      closeFeedbackModal();
    });

    function saveRecentProject(bp) {
      recentBlueprints = [bp, ...recentBlueprints.filter(x => x.title !== bp.title)].slice(0, 4);
      localStorage.setItem("pf_recent_projects", JSON.stringify(recentBlueprints));
      renderRecentProjects();
    }
    function renderRecentProjects() {
      const box = document.getElementById("recent-history-box");
      const list = document.getElementById("recent-list");
      if (!recentBlueprints.length) return box.classList.add("hidden");
      box.classList.remove("hidden");
      list.innerHTML = recentBlueprints.map((item, idx) => `
        <button type="button" onclick="loadRecent(${idx})" class="text-xs px-3 py-1.5 rounded-xl bg-slate-100 hover:bg-slate-200 border border-slate-300 text-slate-800 font-medium truncate max-w-[240px]">
          📌 ${item.title}
        </button>
      `).join("");
    }
    function loadRecent(idx) {
      currentBlueprint = recentBlueprints[idx];
      if (currentBlueprint) { populateAllPages(currentBlueprint); navigatePage("page-3d"); }
    }

    async function generateProject() {
      const topic = document.getElementById("topic-input").value.trim();
      const branch = document.getElementById("branch-input").value;
      let subject = document.getElementById("subject-input").value;
      if (subject === "__MANUAL__") {
        subject = document.getElementById("manual-subject-input").value.trim();
        if (!subject) return alert("Please enter your subject name!");
      }
      const level = document.getElementById("level-input").value;
      if (!topic) return alert("Please enter your project name first!");

      const btn = document.getElementById("generate-btn");
      btn.disabled = true; btn.innerText = "Generating Project...";
      try {
        const res = await fetch("/api/generate", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ topic, branch, subject, level })
        });
        const rawText = await res.text();
        let data;
        try { data = JSON.parse(rawText); } catch (_) { throw new Error("Server waking up. Please click Generate once more!"); }
        if (data.error) throw new Error(data.error);
        currentBlueprint = data.blueprint;
        saveRecentProject(currentBlueprint);
        populateAllPages(currentBlueprint);
        navigatePage("page-3d");
      } catch (err) {
        alert(err.message);
      } finally {
        btn.disabled = false; btn.innerText = "Generate Complete Project";
      }
    }

    function renderCircuitDiagramSVG(bp) {
      const blocks = (bp && bp.circuit_diagram_blocks && bp.circuit_diagram_blocks.length) ? bp.circuit_diagram_blocks : [
        { stage: "1. Input / Supply", title: "Sensor / Voltage Input", detail: "Provides input signal" },
        { stage: "2. Core Stage", title: "IC / Controller Logic", detail: "Compares & processes" },
        { stage: "3. Output Stage", title: "LED / Buzzer / Load", detail: "Displays or acts" }
      ];
      const pins = (bp && bp.pin_connections) ? bp.pin_connections : [];

      const boxW = 240, boxH = 95, gap = 75, startX = 30, y = 45;
      const svgW = Math.max(820, startX * 2 + blocks.length * boxW + (blocks.length - 1) * gap);

      let nodesSvg = "";
      blocks.forEach((b, i) => {
        const x = startX + i * (boxW + gap);
        nodesSvg += `
          <rect x="${x}" y="${y}" width="${boxW}" height="${boxH}" rx="12" fill="#eff6ff" stroke="#1e3a8a" stroke-width="2"/>
          <text x="${x + 14}" y="${y + 24}" font-size="11" font-weight="bold" fill="#1d4ed8">${b.stage || ""}</text>
          <text x="${x + 14}" y="${y + 48}" font-size="14" font-weight="bold" fill="#0f172a">${b.title || ""}</text>
          <text x="${x + 14}" y="${y + 70}" font-size="11" fill="#475569">${b.detail || ""}</text>
        `;
        if (i < blocks.length - 1) {
          const ax1 = x + boxW + 6;
          const ax2 = x + boxW + gap - 10;
          const ay = y + boxH / 2;
          nodesSvg += `
            <line x1="${ax1}" y1="${ay}" x2="${ax2}" y2="${ay}" stroke="#dc2626" stroke-width="2.5"/>
            <polygon points="${ax2},${ay - 5} ${ax2 + 10},${ay} ${ax2},${ay + 5}" fill="#dc2626"/>
          `;
        }
      });

      let wireSummaryY = y + boxH + 34;
      let wireLines = "";
      pins.slice(0, 4).forEach((p, idx) => {
        const wy = wireSummaryY + idx * 22;
        wireLines += `<text x="${startX}" y="${wy}" font-size="12" fill="#1e293b">• ${p.from_pin}  ──▶  ${p.to_pin}  (${p.signal})</text>`;
      });

      const totalH = wireSummaryY + Math.max(1, pins.slice(0, 4).length) * 22 + 18;
      document.getElementById("circuit-diagram-container").innerHTML = `
        <svg viewBox="0 0 ${svgW} ${totalH}" width="100%" height="${totalH}" xmlns="http://www.w3.org/2000/svg">
          ${nodesSvg}
          ${wireLines}
        </svg>
      `;
    }

    function printCircuitDiagram() {
      const content = document.getElementById("circuit-diagram-container").innerHTML;
      const win = window.open("", "_blank");
      win.document.write(`<html><body style="font-family:sans-serif;padding:24px;"><h2>Circuit Diagram</h2>${content}<script>window.print();<\/script></body></html>`);
      win.document.close();
    }

    function populateAllPages(bp) {
      if (bp.requires_circuit) {
        document.getElementById("circuit-lab-wrapper").classList.remove("hidden");
        document.getElementById("software-arch-wrapper").classList.add("hidden");
        document.getElementById("lab3d-title").innerText = `3D Circuit: ${bp.title}`;
        build3DSceneFromBlueprint(bp);
        renderCircuitDiagramSVG(bp);
        document.getElementById("lab3d-pins").innerHTML = (bp.pin_connections || []).map(p => `
          <tr>
            <td class="p-2.5 font-mono font-semibold text-blue-900">${p.from_pin}</td>
            <td class="p-2.5 font-mono text-slate-800">${p.to_pin}</td>
            <td class="p-2.5">${p.signal}</td>
          </tr>
        `).join("");
      } else {
        document.getElementById("circuit-lab-wrapper").classList.add("hidden");
        document.getElementById("software-arch-wrapper").classList.remove("hidden");
        document.getElementById("sw-arch-title").innerText = `System Architecture: ${bp.title}`;
        const nodes = bp.software_architecture_nodes || [];
        document.getElementById("sw-diagram-container").innerHTML = `
          <div class="flex flex-wrap items-center justify-center gap-2 py-2">
            ${nodes.map((n, i) => `
              <div class="px-3.5 py-2 rounded-xl bg-white border-2 border-blue-900 text-xs font-bold text-slate-900 shadow-sm">${n}</div>
              ${i < nodes.length - 1 ? `<span class="text-blue-900 font-extrabold">➜</span>` : ""}
            `).join("")}
          </div>
        `;
        document.getElementById("sw-nodes-container").innerHTML = nodes.map((n, i) => `
          <div class="bg-blue-50 border border-blue-200 rounded-xl p-3.5 text-center">
            <span class="text-[10px] font-bold text-blue-800">STEP 0${i + 1}</span>
            <p class="text-xs font-bold text-slate-900 mt-1">${n}</p>
          </div>
        `).join("");
      }

      document.getElementById("empty-code").classList.add("hidden");
      document.getElementById("content-code").classList.remove("hidden");
      document.getElementById("active-code-subtitle").innerText = `Language: ${bp.primary_code_language || "Code"}`;
      document.getElementById("code-block").innerText = bp.primary_code || "";
      renderVivaCards();

      document.getElementById("empty-report").classList.add("hidden");
      document.getElementById("content-report").classList.remove("hidden");
      document.getElementById("rep-title").innerText = bp.title || "";
      document.getElementById("rep-abstract").innerText = bp.abstract || "";
      document.getElementById("rep-working").innerText = bp.working_principle || "";
    }

    function copySourceCode() {
      navigator.clipboard.writeText(document.getElementById("code-block").innerText);
      alert("Code copied!");
    }

    function selectVivaLang(lang, btn) {
      activeVivaLang = lang;
      document.querySelectorAll(".viva-lang-btn").forEach(b => {
        b.classList.remove("bg-blue-900", "text-white", "font-bold");
        b.classList.add("bg-slate-100", "text-slate-700", "font-semibold");
      });
      btn.classList.remove("bg-slate-100", "text-slate-700", "font-semibold");
      btn.classList.add("bg-blue-900", "text-white", "font-bold");
      renderVivaCards();
    }

    function renderVivaCards() {
      if (!currentBlueprint) return;
      document.getElementById("viva-grid").innerHTML = (currentBlueprint.viva_qa || []).map((item, idx) => {
        const entry = item[activeVivaLang] || item["en"] || {};
        return `
          <div class="bg-slate-50 p-4 rounded-xl border border-slate-200 space-y-2">
            <p class="text-xs font-bold text-blue-900">Q${idx + 1}. ${entry.q || ""}</p>
            <p class="text-xs text-slate-700 leading-relaxed">${entry.a || ""}</p>
          </div>
        `;
      }).join("");
    }

    function getExportPayload() {
      let subject = document.getElementById("subject-input").value;
      if (subject === "__MANUAL__") subject = document.getElementById("manual-subject-input").value.trim() || "Engineering Project";
      return {
        blueprint: currentBlueprint,
        student_name: (studentProfile && studentProfile.name) || "Student Name",
        usn: (studentProfile && studentProfile.usn) || "1AM25EC099",
        college: (studentProfile && studentProfile.college) || "AMC Engineering College",
        branch: document.getElementById("branch-input").value,
        subject: subject,
        guide_name: document.getElementById("guide-input").value.trim() || "Dr. Shivakumar G",
        formatting: { viva_lang: activeVivaLang }
      };
    }

    async function downloadFileFromEndpoint(endpoint, filename, btnId) {
      if (!currentBlueprint) return;
      const btn = document.getElementById(btnId);
      const oldText = btn.innerText;
      btn.disabled = true; btn.innerText = "Preparing...";
      try {
        const res = await fetch(endpoint, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(getExportPayload())
        });
        if (!res.ok) throw new Error("Download failed");
        const blob = await res.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url; a.download = filename;
        document.body.appendChild(a); a.click(); a.remove();
        setTimeout(openFeedbackModal, 600);
      } catch (err) {
        alert(err.message);
      } finally {
        btn.disabled = false; btn.innerText = oldText;
      }
    }

    // ==================== 3D HARDWARE ENGINE ====================
    let scene, camera, renderer, controls, circuitGroup;
    let trackedLabels = [], animParts = [];
    let simActive = true;

    function updateSensorStimulus(val) {
      sensorStimulusLevel = Number(val) / 100;
      document.getElementById("sensor-slider-val").innerText = val + "%";
    }

    function initThreeJS() {
      const container = document.getElementById("threejs-container");
      if (!container || typeof THREE === "undefined") return;

      scene = new THREE.Scene();
      scene.background = new THREE.Color(0xdce5f0);

      camera = new THREE.PerspectiveCamera(42, container.clientWidth / container.clientHeight, 0.1, 1000);
      camera.position.set(0, 8.6, 11.4);

      renderer = new THREE.WebGLRenderer({ antialias: true });
      renderer.setPixelRatio(Math.min(window.devicePixelRatio, 2));
      renderer.setSize(container.clientWidth, container.clientHeight);
      container.innerHTML = "";
      container.appendChild(renderer.domElement);

      if (THREE.OrbitControls) {
        controls = new THREE.OrbitControls(camera, renderer.domElement);
        controls.enableDamping = true;
        controls.dampingFactor = 0.07;
        controls.maxPolarAngle = Math.PI / 2.05;
      }

      scene.add(new THREE.AmbientLight(0xffffff, 0.88));
      const keyLight = new THREE.DirectionalLight(0xffffff, 0.9);
      keyLight.position.set(8, 16, 10);
      scene.add(keyLight);

      circuitGroup = new THREE.Group();
      scene.add(circuitGroup);
      animate3D();
    }

    function createRealisticPartMesh(type) {
      const grp = new THREE.Group();
      if (type === "ic_741" || type === "ic_lm324") {
        const body = new THREE.Mesh(new THREE.BoxGeometry(0.85, 0.24, 0.6), new THREE.MeshStandardMaterial({ color: 0x111827 }));
        grp.add(body);
        for (let px = -0.32; px <= 0.32; px += 0.16) {
          const pin = new THREE.Mesh(new THREE.BoxGeometry(0.05, 0.28, 0.78), new THREE.MeshStandardMaterial({ color: 0xd1d5db, metalness: 0.85 }));
          pin.position.set(px, -0.04, 0);
          grp.add(pin);
        }
      } else if (type === "sensor_water_probes") {
        const tank = new THREE.Mesh(new THREE.BoxGeometry(1.3, 1.0, 0.9), new THREE.MeshStandardMaterial({ color: 0x7dd3fc, transparent: true, opacity: 0.45 }));
        const water = new THREE.Mesh(new THREE.BoxGeometry(1.2, 0.72, 0.82), new THREE.MeshStandardMaterial({ color: 0x0284c7, transparent: true, opacity: 0.82 }));
        water.position.y = -0.1;
        grp.add(tank, water);
        animParts.push({ mesh: water, mode: "water" });
      } else if (type === "led_indicator_bank" || type === "piezo_buzzer") {
        [0x22c55e, 0xeab308, 0xef4444].forEach((clr, idx) => {
          const led = new THREE.Mesh(new THREE.SphereGeometry(0.14, 14, 14), new THREE.MeshBasicMaterial({ color: clr }));
          led.position.set(-0.32 + idx * 0.32, 0.25, -0.12);
          grp.add(led);
          if (idx === 2) animParts.push({ mesh: led, mode: "blink" });
        });
        const buzzer = new THREE.Mesh(new THREE.CylinderGeometry(0.26, 0.26, 0.28, 18), new THREE.MeshStandardMaterial({ color: 0x0f172a }));
        buzzer.position.set(0, 0.18, 0.25);
        grp.add(buzzer);
      } else if (type === "motor_servo") {
        const body = new THREE.Mesh(new THREE.BoxGeometry(1.1, 0.85, 0.55), new THREE.MeshStandardMaterial({ color: 0x1e293b }));
        const arm = new THREE.Mesh(new THREE.BoxGeometry(1.5, 0.12, 0.18), new THREE.MeshStandardMaterial({ color: 0xf8fafc }));
        arm.position.set(0, 0.52, 0);
        grp.add(body, arm);
        animParts.push({ mesh: arm, mode: "servo" });
      } else {
        const box = new THREE.Mesh(new THREE.BoxGeometry(1.2, 0.75, 0.35), new THREE.MeshStandardMaterial({ color: 0x0284c7 }));
        grp.add(box);
      }
      return grp;
    }

    function add3DWire(p1, p2, p3, colorHex) {
      const curve = new THREE.CatmullRomCurve3([new THREE.Vector3(...p1), new THREE.Vector3(...p2), new THREE.Vector3(...p3)]);
      circuitGroup.add(new THREE.Mesh(new THREE.TubeGeometry(curve, 24, 0.04, 8, false), new THREE.MeshStandardMaterial({ color: colorHex })));
    }

    function registerLabel(text, x, y, z) {
      const el = document.createElement("div");
      el.className = "label-3d";
      el.innerText = text;
      document.getElementById("labels-overlay").appendChild(el);
      trackedLabels.push({ el, pos: new THREE.Vector3(x, y, z) });
    }

    function build3DSceneFromBlueprint(bp) {
      if (!circuitGroup) return;
      while (circuitGroup.children.length > 0) circuitGroup.remove(circuitGroup.children[0]);
      trackedLabels = []; animParts = [];
      document.getElementById("labels-overlay").innerHTML = "";

      const supplyLabel = (bp && bp.controller_or_supply) || "9V DC Analog Supply";
      const isAnalog = supplyLabel.toLowerCase().includes("9v") || supplyLabel.toLowerCase().includes("analog");

      if (isAnalog) {
        const bat = new THREE.Mesh(new THREE.BoxGeometry(2.0, 0.95, 1.25), new THREE.MeshStandardMaterial({ color: 0x111827 }));
        bat.position.set(-3.4, 0.35, 0);
        const cap = new THREE.Mesh(new THREE.BoxGeometry(2.02, 0.32, 1.27), new THREE.MeshStandardMaterial({ color: 0xd97706 }));
        cap.position.set(-3.4, 0.68, 0);
        circuitGroup.add(bat, cap);
        registerLabel("9V DC Supply", -3.4, 1.28, 0);
      } else {
        const mcu = new THREE.Mesh(new THREE.BoxGeometry(3.2, 0.22, 2.1), new THREE.MeshStandardMaterial({ color: 0x0284c7 }));
        mcu.position.set(-3.2, 0, 0);
        circuitGroup.add(mcu);
        registerLabel(supplyLabel, -3.2, 0.85, 0);
      }

      const bb = new THREE.Mesh(new THREE.BoxGeometry(5.6, 0.26, 2.8), new THREE.MeshStandardMaterial({ color: 0xffffff }));
      bb.position.set(2.3, 0, 0);
      circuitGroup.add(bb);
      registerLabel("Breadboard", 2.3, 0.55, 1.35);

      add3DWire([-2.2, 0.3, -0.4], [-0.5, 1.35, -0.9], [0.2, 0.15, -1.1], 0xef4444);
      add3DWire([-2.2, 0.3, 0.4], [-0.5, 1.25, 0.9], [0.2, 0.15, 1.1], 0x1e293b);

      const parts = (bp && bp.circuit_3d_parts && bp.circuit_3d_parts.length) ? bp.circuit_3d_parts : [
        { type: "ic_741", label: "IC 741", wire_color: "#2563eb" },
        { type: "sensor_water_probes", label: "Sensor / Probe", wire_color: "#16a34a" },
        { type: "led_indicator_bank", label: "LED & Buzzer", wire_color: "#dc2626" }
      ];

      const xSlots = [0.75, 2.35, 3.95];
      parts.slice(0, 3).forEach((part, i) => {
        const px = xSlots[i];
        const isIC = part.type === "ic_741" || part.type === "ic_lm324";
        const pz = isIC ? 0 : (i % 2 === 0 ? -0.55 : 0.55);
        const py = isIC ? 0.24 : 0.55;
        const mesh = createRealisticPartMesh(part.type);
        mesh.position.set(px, py, pz);
        circuitGroup.add(mesh);
        add3DWire([0.75, 0.28, 0], [(px + 0.75) / 2, 1.45 + i * 0.2, pz], [px, 0.2, pz], part.wire_color || "#2563eb");
        registerLabel(part.label || "Module", px, py + 0.78, pz);
      });
    }

    let simTick = 0;
    function animate3D() {
      requestAnimationFrame(animate3D);
      if (controls) controls.update();
      if (simActive) {
        simTick += 0.05;
        animParts.forEach(item => {
          if (item.mode === "servo") item.mesh.rotation.y = (sensorStimulusLevel - 0.5) * 2.4;
          else if (item.mode === "blink") {
            const on = sensorStimulusLevel >= 0.7;
            item.mesh.material.color.setHex(on && Math.sin(simTick * 3) > 0 ? 0xef4444 : 0x7f1d1d);
          } else if (item.mode === "water") item.mesh.scale.y = Math.max(0.15, sensorStimulusLevel);
        });
      }
      const container = document.getElementById("threejs-container");
      if (container && camera && trackedLabels.length > 0) {
        const w = container.clientWidth, h = container.clientHeight;
        trackedLabels.forEach(item => {
          const v = item.pos.clone().project(camera);
          const x = (v.x * 0.5 + 0.5) * w;
          const y = (-v.y * 0.5 + 0.5) * h;
          item.el.style.left = x + "px";
          item.el.style.top = y + "px";
          item.el.style.display = (v.z < 1 && x > 0 && x < w && y > 0 && y < h) ? "block" : "none";
        });
      }
      if (renderer && scene && camera) renderer.render(scene, camera);
    }

    function toggle3DSimulation() {
      simActive = !simActive;
      const btn = document.getElementById("sim-toggle-btn");
      btn.innerText = simActive ? "▶ Sim: ON" : "⏸️ Sim: PAUSED";
    }
    function reset3DCamera() {
      if (camera) camera.position.set(0, 8.6, 11.4);
      if (controls) { controls.target.set(0, 0, 0); controls.update(); }
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


# ==================== 4. FLASK ROUTES ====================
@app.route("/health")
def health():
    return jsonify({"status": "ok"}), 200


@app.route("/favicon.ico")
def favicon():
    svg = '''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 88 88">
      <rect width="88" height="88" rx="22" fill="#1e3a8a"/>
      <path d="M24 44H36L42 29L49 59L55 44H65" stroke="#38bdf8" stroke-width="5.5" fill="none" stroke-linecap="round" stroke-linejoin="round"/>
    </svg>'''
    return Response(svg, mimetype="image/svg+xml")


@app.route("/manifest.json")
def manifest():
    return jsonify({
        "name": "ProjForge Studio",
        "short_name": "ProjForge",
        "start_url": "/",
        "display": "standalone",
        "background_color": "#ffffff",
        "theme_color": "#1e3a8a",
        "icons": [{"src": "/favicon.ico", "sizes": "192x192 512x512", "type": "image/svg+xml"}]
    })


@app.route("/")
def index():
    return render_template_string(HTML_PAGE)


@app.route("/api/generate", methods=["POST"])
def api_generate():
    data = request.get_json(force=True)
    topic = data.get("topic", "").strip()
    branch = data.get("branch", "Electronics & Communication (ECE)").strip()
    subject = data.get("subject", "Core Engineering").strip()
    level = data.get("level", "Subject Mini-Project").strip()
    if not topic:
        return jsonify({"error": "Please enter a project name."}), 400
    blueprint = generate_project_blueprint(topic, branch, subject, level)
    return jsonify({"status": "ok", "blueprint": blueprint})


@app.route("/api/download-docx", methods=["POST"])
def download_docx():
    payload = request.get_json(force=True)
    docx_buffer = build_vtu_docx(
        payload.get("blueprint", {}),
        payload.get("student_name") or "Student Name",
        payload.get("usn") or "1AM25EC099",
        payload.get("college") or "AMC Engineering College",
        payload.get("branch") or "ECE",
        payload.get("subject") or "Engineering Subject",
        payload.get("guide_name") or "Dr. Shivakumar G",
        payload.get("formatting", {})
    )
    return send_file(docx_buffer, as_attachment=True, download_name="ProjForge_Report.docx", mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@app.route("/api/download-pptx", methods=["POST"])
def download_pptx():
    payload = request.get_json(force=True)
    pptx_buffer = build_seminar_pptx(
        payload.get("blueprint", {}),
        payload.get("student_name") or "Student Name",
        payload.get("usn") or "1AM25EC099",
        payload.get("college") or "AMC Engineering College",
        payload.get("branch") or "ECE",
        payload.get("subject") or "Engineering Subject",
        payload.get("guide_name") or "Dr. Shivakumar G"
    )
    return send_file(pptx_buffer, as_attachment=True, download_name="ProjForge_Seminar.pptx", mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation")


@app.route("/api/download-ieee", methods=["POST"])
def download_ieee():
    payload = request.get_json(force=True)
    ieee_buffer = build_ieee_paper_docx(
        payload.get("blueprint", {}),
        payload.get("student_name") or "Student Name",
        payload.get("usn") or "1AM25EC099",
        payload.get("college") or "AMC Engineering College",
        payload.get("branch") or "ECE",
        payload.get("subject") or "Engineering Subject",
        payload.get("guide_name") or "Dr. Shivakumar G"
    )
    return send_file(ieee_buffer, as_attachment=True, download_name="ProjForge_IEEE_Paper.docx", mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@app.route("/api/feedback", methods=["POST"])
def save_feedback():
    FEEDBACK_LOG.append({"time": datetime.now().isoformat(), **request.get_json(force=True)})
    return jsonify({"status": "saved"})


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))