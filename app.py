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


# ==================== 1. GEMINI AI ENGINE (SUBJECT-AWARE + SMART 3D CIRCUIT) ====================
def generate_project_blueprint(topic, branch, subject, level):
    if GEMINI_API_KEY:
        try:
            url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={GEMINI_API_KEY}"
            prompt = f"""
You are a senior Indian Engineering Professor (VTU / Autonomous curriculum) and Principal Systems Architect.
Generate a deep, subject-aligned engineering project blueprint in strict JSON format for:
- Project Title / Topic: "{topic}"
- Engineering Branch: "{branch}"
- Core Academic Subject: "{subject}"
- Academic Level: "{level}"

CRITICAL RULES:
1. Subject Alignment: The entire methodology, algorithms/circuits, components, and viva questions MUST strictly align with the chosen subject "{subject}" under "{branch}".
2. "requires_circuit": Set to true ONLY IF the project involves electronics, analog/digital circuits (Op-Amp, 555 timer, logic gates), microcontrollers (Arduino/ESP32), sensors, motors, or IoT hardware. Set to false if it is a pure software/web/database/AI software-only project!
3. If "requires_circuit" is true:
   - Populate "circuit_3d_parts" with 3 to 5 specific components matching the exact project (choose types from: "ic_opamp", "sensor_water_probes", "sensor_ultrasonic", "sensor_gas", "sensor_ir", "motor_servo", "relay_module", "oled_display", "piezo_buzzer", "led_bar").
   - Set "controller_or_supply" to either "9V DC Regulated Supply (Analog)" (if pure analog/Op-Amp/555 without MCU) or "Arduino Uno R3 / ESP32".
4. If "requires_circuit" is false:
   - Populate "software_architecture_nodes" with 5 specific software/AI/database pipeline stages.
5. "primary_code_language" and "primary_code": Provide 75-100 lines of clean, runnable code matching the subject (e.g., Python/SQL/HTML-JS for software, or MATLAB/Python simulation / Arduino C++ for hardware).
6. "viva_qa": Provide 8 subject-focused viva questions & answers in all 5 languages: "en", "kn", "hi", "te", "ta".

Return ONLY valid JSON (no markdown fences):
{{
  "title": "Formal Engineering Project Title",
  "tagline": "One-line technical summary aligned with {subject}",
  "requires_circuit": true,
  "controller_or_supply": "Arduino Uno R3 or 9V DC Analog Supply",
  "abstract": "Detailed 280-word academic abstract covering problem statement, {subject} principles, proposed design, and results.",
  "problem_statement": "Detailed 180-word problem definition explaining existing drawbacks.",
  "objectives": ["Objective 1", "Objective 2", "Objective 3", "Objective 4", "Objective 5"],
  "literature_survey": [
    {{"paper": "R. Sharma et al. (2024) - IEEE Transactions", "summary": "Key finding and research gap addressed."}},
    {{"paper": "K. Kulkarni & S. Rao (2025) - Springer", "summary": "Key finding and research gap addressed."}},
    {{"paper": "A. Verma et al. (2026) - Elsevier", "summary": "Key finding and research gap addressed."}}
  ],
  "hardware_requirements": [
    {{"name": "Component 1", "spec": "Detailed spec", "cost_inr": 350}},
    {{"name": "Component 2", "spec": "Detailed spec", "cost_inr": 220}},
    {{"name": "Component 3", "spec": "Detailed spec", "cost_inr": 180}},
    {{"name": "830-Point Breadboard & Jumpers", "spec": "Solderless prototyping board", "cost_inr": 140}}
  ],
  "software_requirements": [
    "OS: Windows 11 / Ubuntu 22.04 LTS",
    "Tools: Subject-specific IDE / Simulator / Python 3.11",
    "Libraries: Relevant libraries"
  ],
  "circuit_3d_parts": [
    {{"type": "ic_opamp", "label": "LM324 / IC 741 Op-Amp", "pin_from": "VCC (+9V) / IN+", "pin_to": "Breadboard Center DIP", "wire_color": "#38bdf8"}},
    {{"type": "sensor_water_probes", "label": "Sensor / Probe Array", "pin_from": "Probe Sense Lead", "pin_to": "Breadboard Row 12", "wire_color": "#22c55e"}},
    {{"type": "led_bar", "label": "LED Level & Piezo Alert", "pin_from": "Output Pin 1 / Pin 7", "pin_to": "Breadboard Row 28", "wire_color": "#f59e0b"}}
  ],
  "software_architecture_nodes": [
    "User Input / Client UI",
    "API Gateway & Validation",
    "Core {subject} Engine",
    "Database / Model Storage",
    "Analytics & Output Dashboard"
  ],
  "pin_connections": [
    {{"from_pin": "Pin / Module A", "to_pin": "Breadboard / Target B", "signal": "Signal / Voltage"}},
    {{"from_pin": "Pin / Module C", "to_pin": "Breadboard / Target D", "signal": "Signal / Voltage"}},
    {{"from_pin": "Power +VCC", "to_pin": "Breadboard Red Rail", "signal": "+5V / +9V DC"}},
    {{"from_pin": "Ground GND", "to_pin": "Breadboard Blue Rail", "signal": "0V Reference"}}
  ],
  "ppt_slides": [
    {{"slide": "Slide 1: Title & Team", "points": "Project Title, Subject ({subject}), USN & Guide"}},
    {{"slide": "Slide 2: Problem Statement", "points": "Existing limitations & engineering need"}},
    {{"slide": "Slide 3: System Design", "points": "Architecture / Circuit schematic & working"}},
    {{"slide": "Slide 4: Implementation", "points": "Code / Breadboard prototype & testing"}},
    {{"slide": "Slide 5: Results & Future Scope", "points": "Performance metrics, applications & conclusion"}}
  ],
  "working_principle": "Detailed 380-word step-by-step technical working explanation grounded in {subject}.",
  "primary_code_language": "Python 3 / C++",
  "primary_code": "Complete 75-95 line runnable code.",
  "has_arduino_hardware": true,
  "arduino_code": "Complete Arduino (.ino) sketch if applicable.",
  "results_and_discussion": "Detailed 220-word analysis of testing parameters, accuracy, and observations.",
  "advantages": ["Advantage 1", "Advantage 2", "Advantage 3", "Advantage 4"],
  "applications": ["Application 1", "Application 2", "Application 3", "Application 4"],
  "future_scope": "160-word discussion on scalability and future upgrades.",
  "references": [
    "[1] IEEE Standard Technical Reference (2025).",
    "[2] Standard Textbook on {subject}, Pearson / McGraw Hill.",
    "[3] VTU Academic Curriculum & Project Guidelines (2025-26)."
  ],
  "viva_qa": [
    {{
      "en": {{"q": "Question 1?", "a": "Answer 1."}},
      "kn": {{"q": "ಪ್ರಶ್ನೆ 1?", "a": "ಉತ್ತರ 1."}},
      "hi": {{"q": "प्रश्न 1?", "a": "उत्तर 1।"}},
      "te": {{"q": "ప్రశ్న 1?", "a": "సమాధానం 1."}},
      "ta": {{"q": "கேள்வி 1?", "a": "பதில் 1."}}
    }}
  ]
}}
"""
            payload = {
                "contents": [{"parts": [{"text": prompt}]}],
                "generationConfig": {"responseMimeType": "application/json", "temperature": 0.4}
            }
            resp = requests.post(url, json=payload, timeout=50)
            if resp.status_code == 200:
                res_json = resp.json()
                raw_text = res_json["candidates"][0]["content"]["parts"][0]["text"].strip()
                raw_text = re.sub(r"^```json\s*|\s*```$", "", raw_text, flags=re.MULTILINE)
                return json.loads(raw_text)
        except Exception as e:
            print("Gemini API fallback:", e)

    return _fallback_blueprint(topic, branch, subject, level)


def _fallback_blueprint(topic, branch, subject, level):
    clean_title = topic.strip().title() or f"Smart {subject} System"
    low = (clean_title + " " + subject + " " + branch).lower()

    # Determine if project requires physical circuit simulation
    pure_software_keywords = ["website", "web app", "database", "dbms", "sql", "management system", "chatbot", "nlp", "data science", "cloud", "blockchain", "cybersecurity", "android app", "software"]
    hardware_keywords = ["opamp", "op-amp", "ic", "arduino", "esp32", "sensor", "iot", "motor", "servo", "relay", "water level", "gate", "robot", "embedded", "vlsi", "ece", "eee", "circuit", "analog", "digital", "555"]

    has_hw = any(k in low for k in hardware_keywords)
    is_pure_sw = any(k in low for k in pure_software_keywords) and not has_hw
    requires_circuit = not is_pure_sw

    is_opamp = "opamp" in low or "op-amp" in low or "analog" in low or "water level" in low

    if is_opamp:
        controller_name = "9V DC Regulated Supply (Analog)"
        parts = [
            {"type": "ic_opamp", "label": "LM324 Quad Op-Amp IC", "pin_from": "Pin 4 (+9V) / Pin 11 (GND)", "pin_to": "Breadboard Center DIP", "wire_color": "#38bdf8"},
            {"type": "sensor_water_probes", "label": "Water Tank Level Probes", "pin_from": "Low/Mid/Full Probes", "pin_to": "Op-Amp Non-Inv Pins (3,5,10)", "wire_color": "#22c55e"},
            {"type": "led_bar", "label": "3-Stage LED & Piezo Buzzer", "pin_from": "Op-Amp Output Pins (1,7,8)", "pin_to": "Breadboard Row 28", "wire_color": "#f59e0b"}
        ]
        pins = [
            {"from_pin": "LM324 Pin 3, 5, 10 (+IN)", "to_pin": "Water Level Sensor Probes", "signal": "Analog Sense Voltage"},
            {"from_pin": "LM324 Pin 2, 6, 9 (-IN)", "to_pin": "10kΩ Reference Potentiometer", "signal": "V_ref Threshold"},
            {"from_pin": "LM324 Pin 1, 7, 8 (OUT)", "to_pin": "Green/Yellow/Red LEDs + Piezo", "signal": "Comparator Output"},
            {"from_pin": "9V DC Battery / Supply", "to_pin": "Breadboard +VCC & -GND Rails", "signal": "9V Regulated DC"}
        ]
        prim_lang = "Python 3 (Op-Amp Comparator Voltage Simulator)"
        prim_code = f"""# ============================================================================
# Project: {clean_title}
# Subject: {subject} ({branch})
# Op-Amp Comparator Threshold Simulation
# ============================================================================

V_SUPPLY = 9.0
V_REF_LOW = 2.2
V_REF_MID = 4.5
V_REF_FULL = 7.2

def evaluate_opamp_comparators(probe_voltage):
    low_led = "ON" if probe_voltage >= V_REF_LOW else "OFF"
    mid_led = "ON" if probe_voltage >= V_REF_MID else "OFF"
    full_led = "ON" if probe_voltage >= V_REF_FULL else "OFF"
    buzzer = "ACTIVE_BEEP" if probe_voltage >= V_REF_FULL else "SILENT"
    print(f"Probe={{probe_voltage:.1f}}V -> LOW:{{low_led}} | MID:{{mid_led}} | FULL:{{full_led}} | Buzzer:{{buzzer}}")

if __name__ == "__main__":
    print("=== {clean_title} Simulation ===")
    for v in [0.5, 2.6, 5.1, 7.8]:
        evaluate_opamp_comparators(v)
"""
    elif requires_circuit:
        controller_name = "Arduino Uno R3"
        parts = [
            {"type": "sensor_ultrasonic", "label": "Ultrasonic / IR Sensor", "pin_from": "Arduino D2 / D3", "pin_to": "Breadboard Row 10", "wire_color": "#38bdf8"},
            {"type": "motor_servo", "label": "SG90 Servo / Relay Unit", "pin_from": "Arduino PWM D9", "pin_to": "Breadboard Row 22", "wire_color": "#f59e0b"},
            {"type": "piezo_buzzer", "label": "Piezo Buzzer & Status LED", "pin_from": "Arduino D8", "pin_to": "Breadboard Row 34", "wire_color": "#ef4444"}
        ]
        pins = [
            {"from_pin": "Arduino Digital D2/D3", "to_pin": "Sensor Trig/Echo (Row 10)", "signal": "5V TTL Digital"},
            {"from_pin": "Arduino PWM Pin D9", "to_pin": "Actuator Control (Row 22)", "signal": "50Hz PWM Signal"},
            {"from_pin": "Arduino Digital D8", "to_pin": "Piezo Buzzer + LED (Row 34)", "signal": "Alert Output"},
            {"from_pin": "Arduino 5V & GND", "to_pin": "Breadboard + / - Rails", "signal": "5V DC Power"}
        ]
        prim_lang = "Arduino C++ (.ino) & Python Telemetry"
        prim_code = f"""// ============================================================================
// Project: {clean_title}
// Subject: {subject} ({branch})
// ============================================================================

#define SENSOR_PIN 2
#define ACTUATOR_PIN 9
#define BUZZER_PIN 8

void setup() {{
  Serial.begin(115200);
  pinMode(SENSOR_PIN, INPUT);
  pinMode(ACTUATOR_PIN, OUTPUT);
  pinMode(BUZZER_PIN, OUTPUT);
  Serial.println("READY: {clean_title}");
}}

void loop() {{
  if (digitalRead(SENSOR_PIN) == HIGH) {{
    digitalWrite(ACTUATOR_PIN, HIGH);
    digitalWrite(BUZZER_PIN, HIGH);
  }} else {{
    digitalWrite(ACTUATOR_PIN, LOW);
    digitalWrite(BUZZER_PIN, LOW);
  }}
  delay(200);
}}
"""
    else:
        controller_name = "Software-Only Architecture"
        parts = []
        pins = []
        prim_lang = f"Python 3 ({subject} Application)"
        prim_code = f"""# ============================================================================
# Project: {clean_title}
# Branch : {branch} | Subject: {subject}
# Pure Software Implementation (No External Hardware Circuit Required)
# ============================================================================

from datetime import datetime

class ProjectEngine:
    def __init__(self, module_name):
        self.module_name = module_name
        self.records = []

    def process_transaction(self, user_id, payload):
        entry = {{
            "id": len(self.records) + 1,
            "user": user_id,
            "data": payload,
            "timestamp": datetime.now().isoformat(),
            "status": "VERIFIED"
        }}
        self.records.append(entry)
        return entry

if __name__ == "__main__":
    engine = ProjectEngine("{clean_title}")
    print(engine.process_transaction("USR_101", {{"subject": "{subject}", "score": 96.4}}))
"""

    return {
        "title": clean_title,
        "tagline": f"Subject-aligned {subject} engineering project for {branch} ({level}).",
        "requires_circuit": requires_circuit,
        "controller_or_supply": controller_name,
        "abstract": (
            f"This project presents the design, analysis, and implementation of '{clean_title}' grounded in the core principles "
            f"of {subject} under {branch}. Existing conventional approaches suffer from manual latency, poor scalability, and "
            "lack of automated real-time feedback. By applying systematic engineering design methodologies taught in "
            f"{subject}, the proposed system achieves high reliability, low operational cost, and 95.2% functional accuracy."
        ),
        "problem_statement": (
            f"Conventional methods used in {clean_title} rely heavily on manual monitoring and outdated architectures, "
            "leading to delayed response times and human errors. There is a strong need for an automated, accurate, "
            f"and cost-effective solution leveraging {subject} concepts."
        ),
        "objectives": [
            f"To apply core {subject} concepts to design '{clean_title}'.",
            "To eliminate manual intervention through automated real-time processing.",
            "To optimize response latency, accuracy, and overall system reliability.",
            "To validate system performance across boundary and stress test cases.",
            "To prepare complete VTU-compliant technical documentation."
        ],
        "literature_survey": [
            {"paper": "R. Kumar et al. (2024) - IEEE Access", "summary": f"Explored foundational {subject} techniques; lacked cost optimization."},
            {"paper": "S. Kulkarni & P. Rao (2025) - Springer", "summary": "Demonstrated prototype validation; higher complexity for student deployment."},
            {"paper": "A. Sharma et al. (2026) - Elsevier", "summary": "Introduced modern automation; our implementation improves efficiency and clarity."}
        ],
        "hardware_requirements": [
            {"name": controller_name, "spec": "Primary Processing / Power Unit", "cost_inr": 350},
            {"name": parts[0]["label"] if parts else "Cloud / Local Server", "spec": "Primary Input Subsystem", "cost_inr": 220},
            {"name": parts[1]["label"] if len(parts) > 1 else "Database Engine", "spec": "Core Logic / Actuation", "cost_inr": 190},
            {"name": "830-Point Breadboard & Jumpers" if requires_circuit else "Development Workstation", "spec": "Prototyping Environment", "cost_inr": 140}
        ],
        "software_requirements": [
            "Operating System: Windows 11 / Ubuntu 22.04 LTS",
            f"Subject Tools: Python 3.11 / MATLAB / Multisim / VS Code ({subject})",
            "Documentation: MS Word / LaTeX VTU Format"
        ],
        "circuit_3d_parts": parts,
        "software_architecture_nodes": [
            "Input Layer & Data Capture",
            "Validation & Pre-Processing",
            f"Core {subject} Processing Engine",
            "Storage & Analytics Layer",
            "Interactive Output & Reporting"
        ],
        "pin_connections": pins,
        "ppt_slides": [
            {"slide": "Slide 1: Title & Team", "points": f"{clean_title} — {subject} ({branch})"},
            {"slide": "Slide 2: Problem Statement", "points": "Limitations of existing approaches & project motivation"},
            {"slide": "Slide 3: Proposed Design", "points": "Architecture / Circuit schematic & component selection"},
            {"slide": "Slide 4: Implementation", "points": "Working methodology, code/breadboard setup & testing"},
            {"slide": "Slide 5: Results & Conclusion", "points": "Experimental accuracy (95.2%), applications & future scope"}
        ],
        "working_principle": (
            f"The architecture of '{clean_title}' is structured around core {subject} principles. In Stage 1, input signals/data "
            "are acquired and conditioned to remove noise or invalid entries. In Stage 2, the core comparator/algorithm evaluates "
            "the conditioned inputs against calibrated reference thresholds. In Stage 3, deterministic output actions (visual LEDs, "
            "audio alerts, or dashboard updates) are triggered immediately with sub-200ms latency."
        ),
        "primary_code_language": prim_lang,
        "primary_code": prim_code,
        "has_arduino_hardware": requires_circuit,
        "arduino_code": f"// Hardware Firmware for {clean_title}\nvoid setup() {{ Serial.begin(115200); }}\nvoid loop() {{ Serial.println(\"Running {clean_title}\"); delay(500); }}",
        "results_and_discussion": (
            f"Comprehensive testing was conducted under {subject} lab conditions across 50 trials. The system demonstrated "
            "consistent threshold stability, zero false alarms, and 95.2% overall accuracy."
        ),
        "advantages": [
            f"Directly demonstrates practical {subject} engineering principles",
            "High reliability and sub-200ms response latency",
            "Low implementation cost and modular architecture",
            "Easy to test, debug, and demonstrate in lab vivas"
        ],
        "applications": [
            "Industrial & Commercial Automation",
            "Smart Campus & Residential Safety Systems",
            f"{subject} Laboratory Demonstrations",
            "Scalable Enterprise & IoT Deployments"
        ],
        "future_scope": "Can be extended with cloud telemetry, mobile app alerts, and PCB miniaturization.",
        "references": [
            f"[1] IEEE Standard Reference on {subject} Systems (2025).",
            f"[2] Standard University Textbook on {subject}, Pearson Education.",
            "[3] VTU Academic Project Report Guidelines (2025-26)."
        ],
        "viva_qa": [
            {
                "en": {"q": f"How does this project apply {subject} concepts?", "a": f"It directly implements core {subject} principles for signal/data processing and threshold decision-making."},
                "kn": {"q": f"ಈ ಪ್ರಾಜೆಕ್ಟ್ {subject} ತತ್ವಗಳನ್ನು ಹೇಗೆ ಬಳಸುತ್ತದೆ?", "a": f"ಇದು ಸಿಗ್ನಲ್/ಡೇಟಾ ಸಂಸ್ಕರಣೆ ಮತ್ತು ನಿರ್ಧಾರ ತೆಗೆದುಕೊಳ್ಳಲು {subject} ತತ್ವಗಳನ್ನು ನೇರವಾಗಿ ಬಳಸುತ್ತದೆ."},
                "hi": {"q": f"यह प्रोजेक्ट {subject} के सिद्धांतों को कैसे लागू करता है?", "a": f"यह सिग्नल/डेटा प्रोसेसिंग और निर्णय लेने के लिए {subject} के सिद्धांतों का उपयोग करता है।"},
                "te": {"q": f"ఈ ప్రాజెక్ట్ {subject} సూత్రాలను ఎలా ఉపయోగిస్తుంది?", "a": f"ఇది సిగ్నల్/డేటా ప్రాసెసింగ్ మరియు నిర్ణయం తీసుకోవడానికి {subject} సూత్రాలను ఉపయోగిస్తుంది."},
                "ta": {"q": f"இந்த திட்டம் {subject} கோட்பாடுகளை எவ்வாறு பயன்படுத்துகிறது?", "a": f"இது சிக்னல்/தரவு செயலாக்கம் மற்றும் முடிவெடுப்பதற்கு {subject} கோட்பாடுகளைப் பயன்படுத்துகிறது."}
            },
            {
                "en": {"q": "What is the main advantage of your proposed design?", "a": "High accuracy (95.2%), low cost, and fast real-time response under 200 ms."},
                "kn": {"q": "ನಿಮ್ಮ ವಿನ್ಯಾಸದ ಮುಖ್ಯ ಪ್ರಯೋಜನವೇನು?", "a": "ಹೆಚ್ಚಿನ ನಿಖರತೆ (95.2%), ಕಡಿಮೆ ವೆಚ್ಚ ಮತ್ತು 200 ms ಗಿಂತ ವೇಗವಾದ ಪ್ರತಿಕ್ರಿಯೆ."},
                "hi": {"q": "आपके डिज़ाइन का मुख्य लाभ क्या है?", "a": "उच्च सटीकता (95.2%), कम लागत और 200 ms से कम का तेज़ रिस्पॉन्स।"},
                "te": {"q": "మీ డిజైన్ యొక్క ప్రధాన ప్రయోజనం ఏమిటి?", "a": "అధిక ఖచ్చితత్వం (95.2%), తక్కువ ఖర్చు మరియు 200 ms కంటే వేగవంతమైన స్పందన."},
                "ta": {"q": "உங்கள் வடிவமைப்பின் முக்கிய நன்மை என்ன?", "a": "அதிக துல்லியம் (95.2%), குறைந்த செலவு மற்றும் 200 ms க்கும் குறைவான வேகமான பதில்."}
            },
            {
                "en": {"q": "How did you verify and test your prototype?", "a": "By running 50 repeated test cases and measuring output stability against reference thresholds."},
                "kn": {"q": "ನಿಮ್ಮ ಮಾದರಿಯನ್ನು ನೀವು ಹೇಗೆ ಪರೀಕ್ಷಿಸಿದ್ದೀರಿ?", "a": "50 ಬಾರಿ ಪರೀಕ್ಷೆ ನಡೆಸಿ ರೆಫರೆನ್ಸ್ ಥ್ರೆಶೋಲ್ಡ್‌ಗಳೊಂದಿಗೆ ಫಲಿತಾಂಶವನ್ನು ದೃಢಪಡಿಸಲಾಗಿದೆ."},
                "hi": {"q": "आपने अपने प्रोटोटाइप का परीक्षण कैसे किया?", "a": "50 परीक्षण चलाकर और रेफरेंस थ्रेशोल्ड के साथ आउटपुट की जांच करके।"},
                "te": {"q": "మీరు మీ ప్రోటోటైప్‌ను ఎలా పరీక్షించారు?", "a": "50 పరీక్షలు నిర్వహించి రిఫరెన్స్ థ్రెషోల్డ్‌లతో అవుట్‌పుట్‌ను సరిపోల్చడం ద్వారా."},
                "ta": {"q": "உங்கள் முன்மாதிரியை எவ்வாறு சோதித்தீர்கள்?", "a": "50 சோதனைகளை நடத்தி வெளியீட்டு நிலைத்தன்மையை உறுதி செய்தோம்."}
            }
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

    # PAGE 1: COVER PAGE
    for txt, sz, bld, itl, clr in [
        ("VISVESVARAYA TECHNOLOGICAL UNIVERSITY / AUTONOMOUS\n", font_size + 2, True, False, RGBColor(15, 35, 95)),
        (f"{college.upper()}\nDepartment of {branch}\n\n", font_size + 1, True, False, None),
        (f"A TECHNICAL PROJECT REPORT ON\n({subject.upper()})\n\n", font_size, False, True, None),
        (f'"{title.upper()}"\n\n', font_size + 6, True, False, RGBColor(180, 30, 30)),
        (f"Submitted in partial fulfillment of the requirements for the award of the degree of\nBACHELOR OF ENGINEERING IN {branch.upper()}\n\nSubmitted By:\n", font_size, False, False, None),
        (f"{student_name.upper()}   (USN: {usn.upper()})\n\n", font_size + 1, True, False, None),
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

    # PAGE 2: BONAFIDE CERTIFICATE
    add_h(f"{college.upper()}", 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_h(f"Department of {branch}", 2).alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_h("BONAFIDE CERTIFICATE", 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_b(
        f'Certified that the Project work entitled "{title}" under the subject {subject} is a bonafide work carried out by '
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
    add_h("DECLARATION", 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_b(
        f'I, {student_name} (USN: {usn}), student of {branch}, {college}, hereby declare that the Project '
        f'entitled "{title}" has been independently carried out by me under the supervision of Prof. {guide_name}.'
    )
    add_b(f"\nPlace: Karnataka\nDate: {datetime.now().strftime('%d-%m-%Y')}                                   Signature: ({student_name})", bold=True)
    add_h("ACKNOWLEDGEMENT", 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_b(
        f"I express my sincere gratitude to my guide Prof. {guide_name}, HOD, Principal, and faculty of {college} "
        f"for their continuous guidance and laboratory support in {subject} throughout this project."
    )
    doc.add_page_break()

    # PAGE 4: ABSTRACT & TOC
    add_h("ABSTRACT", 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_b(blueprint.get("abstract", ""))
    add_h("TABLE OF CONTENTS", 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
    toc_items = [
        ("1.", "CHAPTER 1: INTRODUCTION & PROBLEM DEFINITION", "05"),
        ("2.", "CHAPTER 2: LITERATURE SURVEY", "07"),
        ("3.", "CHAPTER 3: SYSTEM DESIGN & CIRCUIT / ARCHITECTURE", "09"),
        ("4.", "CHAPTER 4: REQUIREMENTS & BILL OF MATERIALS", "11"),
        ("5.", "CHAPTER 5: METHODOLOGY & IMPLEMENTATION CODE", "12"),
        ("6.", "CHAPTER 6: RESULTS, APPLICATIONS & REFERENCES", "14"),
        ("7.", "APPENDIX-A: VIVA-VOCE QUESTIONS & ANSWERS", "16"),
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

    # CHAPTER 1: INTRODUCTION
    add_h("CHAPTER 1: INTRODUCTION & PROBLEM DEFINITION", 1)
    add_h("1.1 Overview", 2)
    add_b(blueprint.get("abstract", ""))
    add_h("1.2 Problem Statement", 2)
    add_b(blueprint.get("problem_statement", ""))
    add_h("1.3 Project Objectives", 2)
    for obj in blueprint.get("objectives", []):
        p = doc.add_paragraph(obj, style="List Bullet")
        p.runs[0].font.name = font_family
        p.runs[0].font.size = Pt(font_size)
    doc.add_page_break()

    # CHAPTER 2: LITERATURE SURVEY
    add_h("CHAPTER 2: LITERATURE SURVEY", 1)
    lit_list = blueprint.get("literature_survey", [])
    if lit_list:
        lt = doc.add_table(rows=len(lit_list) + 1, cols=2)
        lt.style = "Table Grid"
        for idx, h in enumerate(["Reference Paper / Author", "Key Findings & Research Gap"]):
            c = lt.cell(0, idx)
            c.text = h
            _set_cell_shading(c, "1E3A8A")
            c.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
        for i, item in enumerate(lit_list, start=1):
            lt.cell(i, 0).text = item.get("paper", "")
            lt.cell(i, 1).text = item.get("summary", "")
    doc.add_page_break()

    # CHAPTER 3: SYSTEM DESIGN
    add_h("CHAPTER 3: SYSTEM DESIGN & INTERFACING", 1)
    if blueprint.get("requires_circuit") and blueprint.get("pin_connections"):
        add_h("3.1 Hardware Pin-to-Pin Connections", 2)
        pins = blueprint.get("pin_connections", [])
        pt = doc.add_table(rows=len(pins) + 1, cols=3)
        pt.style = "Table Grid"
        for idx, h in enumerate(["Source Pin / Module", "Target Pin / Breadboard", "Signal"]):
            c = pt.cell(0, idx)
            c.text = h
            _set_cell_shading(c, "1E3A8A")
            c.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
        for i, pin in enumerate(pins, start=1):
            pt.cell(i, 0).text = pin.get("from_pin", "")
            pt.cell(i, 1).text = pin.get("to_pin", "")
            pt.cell(i, 2).text = pin.get("signal", "")
    else:
        add_h("3.1 Software Pipeline Architecture", 2)
        for idx, node in enumerate(blueprint.get("software_architecture_nodes", []), start=1):
            add_b(f"Stage {idx}: {node}", bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)
    doc.add_page_break()

    # CHAPTER 4: REQUIREMENTS
    add_h("CHAPTER 4: HARDWARE & SOFTWARE REQUIREMENTS", 1)
    for hw in blueprint.get("hardware_requirements", []):
        add_b(f"• {hw.get('name', '')} ({hw.get('spec', '')}) — Est. Rs. {hw.get('cost_inr', 0)}", align=WD_ALIGN_PARAGRAPH.LEFT)
    for sw in blueprint.get("software_requirements", []):
        add_b(f"• {sw}", align=WD_ALIGN_PARAGRAPH.LEFT)
    doc.add_page_break()

    # CHAPTER 5: WORKING & CODE
    add_h("CHAPTER 5: METHODOLOGY & IMPLEMENTATION CODE", 1)
    add_b(blueprint.get("working_principle", ""))
    add_h(f"5.1 Source Code ({blueprint.get('primary_code_language', 'Code')})", 2)
    ct = doc.add_table(rows=1, cols=1)
    ct.style = "Table Grid"
    cc = ct.cell(0, 0)
    _set_cell_shading(cc, "F3F4F6")
    r = cc.paragraphs[0].add_run(blueprint.get("primary_code", ""))
    r.font.name = "Consolas"
    r.font.size = Pt(9.5)
    doc.add_page_break()

    # CHAPTER 6: RESULTS & REFERENCES
    add_h("CHAPTER 6: RESULTS, APPLICATIONS & REFERENCES", 1)
    add_b(blueprint.get("results_and_discussion", ""))
    for adv in blueprint.get("advantages", []):
        add_b("✓ Advantage: " + adv, align=WD_ALIGN_PARAGRAPH.LEFT)
    for ap in blueprint.get("applications", []):
        add_b("✓ Application: " + ap, align=WD_ALIGN_PARAGRAPH.LEFT)
    add_b("Future Scope: " + blueprint.get("future_scope", ""))
    add_h("References", 2)
    for ref in blueprint.get("references", []):
        add_b(ref, align=WD_ALIGN_PARAGRAPH.LEFT)
    doc.add_page_break()

    # APPENDIX A: VIVA Q&A
    add_h("APPENDIX-A: VIVA-VOCE QUESTIONS & ANSWERS", 1)
    for idx, qa in enumerate(blueprint.get("viva_qa", []), start=1):
        lang_obj = qa.get(viva_lang) or qa.get("en") or {}
        add_b(f"Q{idx}: {lang_obj.get('q', '')}", bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)
        add_b(f"Ans: {lang_obj.get('a', '')}", align=WD_ALIGN_PARAGRAPH.LEFT)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


# ==================== 3. FRONTEND HTML (BRANCH->SUBJECTS + TINKERCAD 3D + SMART VIEW) ====================
HTML_PAGE = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
  <title>ProjForge AI — Subject-Aware 3D Circuit & VTU Studio</title>
  <script src="https://cdn.tailwindcss.com"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/three.js/r128/three.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/three@0.128.0/examples/js/controls/OrbitControls.js"></script>
  <link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;700&family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;600&display=swap" rel="stylesheet">
  <style>
    * { box-sizing: border-box; }
    html, body {
      max-width: 100%; overflow-x: hidden;
      font-family: 'Inter', sans-serif; background: #050811; color: #f1f5f9;
    }
    h1, h2, h3, .brand-font { font-family: 'Space Grotesk', sans-serif; }
    code, pre { font-family: 'JetBrains Mono', monospace; }

    #splash-screen {
      position: fixed; inset: 0; z-index: 9999;
      background: radial-gradient(circle at center, #0f172a 0%, #020617 100%);
      display: flex; flex-direction: column; align-items: center; justify-content: center;
      transition: opacity 0.5s ease, visibility 0.5s ease;
    }
    .cyber-ring {
      width: 96px; height: 96px; border-radius: 50%;
      border: 3px solid transparent; border-top-color: #38bdf8; border-right-color: #818cf8;
      animation: spinRing 1s linear infinite;
      display: flex; align-items: center; justify-content: center;
    }
    .cyber-core {
      width: 60px; height: 60px; border-radius: 14px;
      background: linear-gradient(135deg, #2563eb, #7c3aed);
      display: flex; align-items: center; justify-content: center;
      font-weight: 800; font-size: 22px; color: white;
    }
    @keyframes spinRing { 100% { transform: rotate(360deg); } }

    .glass { background: rgba(15, 23, 42, 0.85); backdrop-filter: blur(14px); border: 1px solid rgba(148, 163, 184, 0.14); }
    .page-section { display: none; animation: fadeIn 0.2s ease; }
    .page-section.active { display: block; }
    @keyframes fadeIn { from { opacity: 0; transform: translateY(5px); } to { opacity: 1; transform: translateY(0); } }

    .label-3d {
      position: absolute; pointer-events: none; transform: translate(-50%, -100%);
      background: rgba(15, 23, 42, 0.92); border: 1px solid rgba(56, 189, 248, 0.6);
      color: #f8fafc; font-size: 10px; font-weight: 600; padding: 2px 8px;
      border-radius: 6px; white-space: nowrap; box-shadow: 0 3px 10px rgba(0,0,0,0.45);
    }
  </style>
</head>
<body class="pb-20 md:pb-8">

  <!-- 1. SPLASH SCREEN -->
  <div id="splash-screen">
    <div class="cyber-ring"><div class="cyber-core">PF</div></div>
    <h1 class="text-2xl font-bold mt-4 bg-gradient-to-r from-sky-400 to-indigo-400 bg-clip-text text-transparent">ProjForge AI Studio</h1>
    <p class="text-[11px] text-slate-400 mt-1 tracking-widest uppercase">Subject-Aware Engineering &amp; 3D Circuit Lab...</p>
  </div>

  <!-- 2. LOGIN MODAL -->
  <div id="login-modal" class="fixed inset-0 z-50 hidden items-center justify-center bg-black/80 backdrop-blur-md p-4">
    <div class="glass max-w-md w-full rounded-2xl p-6 shadow-2xl border border-sky-500/30">
      <h2 class="text-lg font-bold text-white">Student Workspace Sign-In</h2>
      <p class="text-xs text-slate-400 mb-3">Auto-fills your Name &amp; USN on VTU .docx reports</p>
      <form id="login-form" class="space-y-3">
        <div>
          <label class="text-xs text-slate-300">Student Email *</label>
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
        <button type="submit" class="w-full py-3 rounded-xl bg-gradient-to-r from-sky-500 to-indigo-600 font-semibold text-white text-sm">
          Save &amp; Continue →
        </button>
      </form>
      <button onclick="skipLogin()" class="w-full mt-3 text-xs text-slate-400 hover:text-slate-200 underline">Continue as Guest</button>
    </div>
  </div>

  <!-- 3. REFER FRIENDS MODAL -->
  <div id="refer-modal" class="fixed inset-0 z-50 hidden items-center justify-center bg-black/80 backdrop-blur-md p-4">
    <div class="glass max-w-md w-full rounded-2xl p-6 shadow-2xl border border-indigo-500/40 relative">
      <button onclick="closeReferModal()" class="absolute top-4 right-4 text-slate-400 hover:text-white text-sm">✕</button>
      <h2 class="text-base font-bold text-white">🎁 Refer Friends &amp; Classmates</h2>
      <p class="text-xs text-slate-400 mt-1">Share ProjForge AI with your batchmates</p>
      <div class="bg-slate-900 border border-slate-800 rounded-xl p-3 mt-3">
        <span class="text-[10px] text-slate-400">Your Referral Code:</span>
        <div class="flex items-center justify-between mt-1">
          <span id="my-ref-code" class="font-mono font-bold text-sm text-sky-400">PF-STUDENT</span>
          <button onclick="copyRefCode()" class="text-xs px-2.5 py-1 rounded-lg bg-slate-800 text-slate-200">Copy</button>
        </div>
      </div>
      <div class="grid grid-cols-2 gap-2.5 mt-4">
        <button onclick="shareOnWhatsApp()" class="py-2.5 px-3 rounded-xl bg-emerald-600 font-semibold text-white text-xs">💬 WhatsApp Share</button>
        <button onclick="copyInviteLink()" class="py-2.5 px-3 rounded-xl bg-indigo-600 font-semibold text-white text-xs">🔗 Copy Link</button>
      </div>
    </div>
  </div>

  <!-- 4. TOP NAVBAR -->
  <header class="sticky top-0 z-40 glass border-b border-slate-800/90">
    <div class="max-w-6xl mx-auto px-4 py-3 flex items-center justify-between">
      <div class="flex items-center gap-2.5 cursor-pointer" onclick="navigatePage('page-home')">
        <div class="w-9 h-9 rounded-xl bg-gradient-to-br from-sky-500 to-indigo-600 flex items-center justify-center font-bold text-white shadow">PF</div>
        <span class="brand-font font-bold text-base sm:text-lg text-white">ProjForge <span class="text-sky-400">AI</span></span>
      </div>

      <nav class="hidden md:flex items-center gap-1.5 bg-slate-900/90 p-1 rounded-xl border border-slate-800">
        <button onclick="navigatePage('page-home')" id="nav-d-home" class="nav-d-btn px-3 py-1.5 rounded-lg text-xs font-semibold bg-sky-600 text-white">🏠 Create</button>
        <button onclick="navigatePage('page-3d')" id="nav-d-3d" class="nav-d-btn px-3 py-1.5 rounded-lg text-xs font-semibold text-slate-300">🧊 Circuit / Architecture</button>
        <button onclick="navigatePage('page-code')" id="nav-d-code" class="nav-d-btn px-3 py-1.5 rounded-lg text-xs font-semibold text-slate-300">💻 Code &amp; 5-Lang Viva</button>
        <button onclick="navigatePage('page-report')" id="nav-d-report" class="nav-d-btn px-3 py-1.5 rounded-lg text-xs font-semibold text-slate-300">📄 Report Studio</button>
      </nav>

      <div class="relative">
        <button onclick="toggleProfileMenu()" id="profile-trigger"
          class="flex items-center gap-2 px-3 py-1.5 rounded-xl bg-slate-900 hover:bg-slate-800 border border-slate-700 text-xs text-slate-200">
          <span id="avatar-initial" class="w-6 h-6 rounded-full bg-gradient-to-tr from-sky-500 to-indigo-600 flex items-center justify-center font-bold text-[11px] text-white">G</span>
          <span id="profile-short-name" class="max-w-[85px] truncate font-medium">Guest</span>
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
            <span>🎁</span> Refer Friends
          </button>
          <div class="border-t border-slate-800 my-1"></div>
          <button onclick="closeProfileMenu(); logoutUser();" class="w-full text-left px-4 py-2.5 text-xs text-rose-400 hover:bg-rose-500/15 font-semibold flex items-center gap-2">
            <span>🚪</span> Logout
          </button>
        </div>
      </div>
    </div>
  </header>

  <main class="max-w-5xl mx-auto px-4 pt-6">

    <!-- ==================== PAGE 1: BRANCH + SUBJECT + TOPIC GENERATOR ==================== -->
    <section id="page-home" class="page-section active space-y-5">
      <div class="text-center max-w-2xl mx-auto">
        <h1 class="text-2xl sm:text-4xl font-bold leading-tight bg-gradient-to-r from-white via-sky-200 to-indigo-300 bg-clip-text text-transparent">
          Subject-Specific Engineering Project &amp; 3D Circuit Studio
        </h1>
        <p class="text-slate-400 text-xs sm:text-sm mt-2">
          Select your Branch &amp; Subject first—projects, 3D hardware simulations, code &amp; VTU reports adapt automatically.
        </p>
      </div>

      <div class="glass rounded-2xl p-5 sm:p-7 max-w-3xl mx-auto shadow-xl border border-slate-800 space-y-4">
        <div class="grid grid-cols-1 sm:grid-cols-2 gap-3.5">
          <div>
            <label class="text-xs font-semibold text-sky-400">1. Select Engineering Branch *</label>
            <select id="branch-input" onchange="updateSubjectsForBranch()"
              class="w-full mt-1 px-3 py-2.5 rounded-xl bg-slate-950 border border-slate-700 text-xs sm:text-sm text-white focus:border-sky-400 focus:outline-none">
              <option value="Computer Science & Engineering (CSE)">Computer Science &amp; Engineering (CSE)</option>
              <option value="Electronics & Communication (ECE)">Electronics &amp; Communication (ECE)</option>
              <option value="Artificial Intelligence & ML (AIML)">Artificial Intelligence &amp; ML (AIML)</option>
              <option value="Electrical & Electronics (EEE)">Electrical &amp; Electronics (EEE)</option>
              <option value="Information Science (ISE)">Information Science (ISE)</option>
              <option value="Mechanical Engineering">Mechanical Engineering</option>
              <option value="Civil Engineering">Civil Engineering</option>
            </select>
          </div>

          <div>
            <label class="text-xs font-semibold text-emerald-400">2. Select Branch Subject *</label>
            <select id="subject-input"
              class="w-full mt-1 px-3 py-2.5 rounded-xl bg-slate-950 border border-slate-700 text-xs sm:text-sm text-white focus:border-emerald-400 focus:outline-none">
            </select>
          </div>
        </div>

        <div>
          <label class="text-xs font-semibold uppercase tracking-wider text-sky-400">3. Enter Project Title / Problem Statement *</label>
          <input id="topic-input" type="text"
            placeholder="e.g., Water Level Indicator using Opamp / Smart Attendance System"
            class="w-full mt-1.5 px-4 py-3.5 rounded-xl bg-slate-950 border border-slate-700 text-white text-sm focus:border-sky-400 focus:outline-none" />
        </div>

        <div class="grid grid-cols-1 sm:grid-cols-2 gap-3.5">
          <div>
            <label class="text-xs text-slate-400">Project Level</label>
            <select id="level-input" class="w-full mt-1 px-3 py-2.5 rounded-xl bg-slate-950 border border-slate-700 text-xs sm:text-sm text-white">
              <option>2nd / 3rd Year Subject Mini-Project</option>
              <option>Final Year Major Capstone Project</option>
              <option>Lab / Hackathon Prototype</option>
            </select>
          </div>
          <div>
            <label class="text-xs text-slate-400">Guide Name (For Certificate)</label>
            <input id="guide-input" type="text" placeholder="Dr. S. K. Kulkarni"
              class="w-full mt-1 px-3 py-2.5 rounded-xl bg-slate-950 border border-slate-700 text-xs sm:text-sm text-white" />
          </div>
        </div>

        <div class="flex flex-wrap items-center gap-2 pt-1">
          <span class="text-xs text-slate-400">Quick Examples:</span>
          <button type="button" onclick="applyQuickExample('Electronics & Communication (ECE)', 'Analog Electronic Circuits (Op-Amp & Timers)', 'Water Level Indicator using Opamp')"
            class="text-xs px-2.5 py-1 rounded-lg bg-slate-900 hover:bg-slate-800 text-emerald-300 border border-slate-700">💧 Water Level Op-Amp (Circuit)</button>
          <button type="button" onclick="applyQuickExample('Computer Science & Engineering (CSE)', 'Database Management Systems (DBMS)', 'Hospital Patient Records & Appointment Management System')"
            class="text-xs px-2.5 py-1 rounded-lg bg-slate-900 hover:bg-slate-800 text-sky-300 border border-slate-700">🏥 Hospital DBMS (Software)</button>
          <button type="button" onclick="applyQuickExample('Electronics & Communication (ECE)', 'Embedded Systems & Microcontrollers', 'Smart Automated Railway Gate using Ultrasonic & Servo')"
            class="text-xs px-2.5 py-1 rounded-lg bg-slate-900 hover:bg-slate-800 text-amber-300 border border-slate-700">🚂 Railway Gate (Arduino)</button>
        </div>

        <button id="generate-btn" onclick="generateProject()"
          class="w-full py-3.5 rounded-xl bg-gradient-to-r from-sky-500 via-blue-600 to-indigo-600 font-bold text-white text-sm sm:text-base shadow-lg">
          ⚡ Generate Subject-Aligned Project
        </button>
      </div>
    </section>

    <!-- ==================== PAGE 2: CONDITIONAL TINKERCAD 3D LAB OR SOFTWARE ARCHITECTURE ==================== -->
    <section id="page-3d" class="page-section space-y-4">
      <!-- Case A: Hardware / Circuit Project -> Tinkercad-Style 3D Lab -->
      <div id="circuit-lab-wrapper" class="space-y-4">
        <div class="glass rounded-2xl p-4 sm:p-5 border border-sky-500/30 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div>
            <span class="text-[11px] uppercase tracking-wider text-sky-400 font-bold">Tinkercad-Style 3D Circuit Workbench</span>
            <h2 id="lab3d-title" class="text-base sm:text-xl font-bold text-white">3D Breadboard &amp; Component Assembly</h2>
            <p class="text-xs text-slate-400">Drag to Rotate 360° • Pinch/Scroll to Zoom • Click Palette below to add/test parts</p>
          </div>
          <div class="flex flex-wrap gap-2">
            <button onclick="toggle3DSimulation()" id="sim-toggle-btn"
              class="px-3.5 py-2 rounded-xl bg-emerald-600 font-bold text-white text-xs shadow">
              ▶ Start Simulation: ON
            </button>
            <button onclick="reset3DCamera()"
              class="px-3 py-2 rounded-xl bg-slate-800 text-xs text-slate-200 border border-slate-700">
              🔄 Reset View
            </button>
          </div>
        </div>

        <!-- Tinkercad Light-Lab 3D Viewport -->
        <div class="glass rounded-2xl p-2 border border-slate-700 relative overflow-hidden">
          <div id="threejs-wrapper" class="relative w-full h-[380px] sm:h-[460px] rounded-xl bg-[#e2e8f0] overflow-hidden">
            <div id="threejs-container" class="w-full h-full"></div>
            <div id="labels-overlay" class="absolute inset-0 pointer-events-none overflow-hidden"></div>
          </div>

          <!-- Tinkercad Component Quick-Add Strip -->
          <div class="flex flex-wrap items-center gap-2 pt-2.5 px-1">
            <span class="text-[11px] font-semibold text-slate-400">Add Component:</span>
            <button onclick="addExtra3DPart('sensor_ultrasonic', 'Ultrasonic HC-SR04')" class="text-[11px] px-2.5 py-1 rounded-lg bg-slate-900 hover:bg-slate-800 border border-slate-700 text-sky-300">+ Ultrasonic</button>
            <button onclick="addExtra3DPart('motor_servo', 'Micro Servo Motor')" class="text-[11px] px-2.5 py-1 rounded-lg bg-slate-900 hover:bg-slate-800 border border-slate-700 text-amber-300">+ Servo Motor</button>
            <button onclick="addExtra3DPart('piezo_buzzer', 'Piezo Buzzer')" class="text-[11px] px-2.5 py-1 rounded-lg bg-slate-900 hover:bg-slate-800 border border-slate-700 text-rose-300">+ Piezo Buzzer</button>
            <button onclick="addExtra3DPart('led_bar', 'Status LEDs')" class="text-[11px] px-2.5 py-1 rounded-lg bg-slate-900 hover:bg-slate-800 border border-slate-700 text-emerald-300">+ LEDs</button>
          </div>
        </div>

        <!-- Pin-to-Pin Wiring Table -->
        <div class="glass rounded-2xl p-5">
          <h3 class="text-sm font-bold text-sky-300 mb-3">📌 Exact Pin-to-Pin Breadboard Connections</h3>
          <div class="overflow-x-auto">
            <table class="w-full text-xs text-left border border-slate-800">
              <thead class="bg-slate-900 text-sky-300">
                <tr>
                  <th class="p-2.5 border-b border-slate-800">Source Pin / IC</th>
                  <th class="p-2.5 border-b border-slate-800">Breadboard / Target</th>
                  <th class="p-2.5 border-b border-slate-800">Signal / Voltage</th>
                </tr>
              </thead>
              <tbody id="lab3d-pins" class="divide-y divide-slate-800 text-slate-300"></tbody>
            </table>
          </div>
        </div>
      </div>

      <!-- Case B: Pure Software / Non-Circuit Project -> Clean Software Architecture View -->
      <div id="software-arch-wrapper" class="hidden space-y-4">
        <div class="glass rounded-2xl p-6 border border-indigo-500/30 space-y-4">
          <div class="flex items-center justify-between">
            <div>
              <span class="text-[11px] uppercase tracking-wider text-indigo-400 font-bold">Pure Software / Algorithmic Project</span>
              <h2 id="sw-arch-title" class="text-lg sm:text-xl font-bold text-white mt-0.5">System Architecture &amp; Data Pipeline</h2>
              <p class="text-xs text-slate-400 mt-1">No hardware circuit is needed for this software subject—here is the complete system flow:</p>
            </div>
          </div>

          <div id="sw-nodes-container" class="grid grid-cols-1 sm:grid-cols-5 gap-3 pt-2"></div>
        </div>
      </div>
    </section>

    <!-- ==================== PAGE 3: CODE & 5-LANG VIVA ==================== -->
    <section id="page-code" class="page-section space-y-5">
      <div id="empty-code" class="glass rounded-2xl p-8 text-center">
        <p class="text-sm text-slate-400">Generate a project first to view subject-specific source code and 5-language Viva Q&amp;A.</p>
        <button onclick="navigatePage('page-home')" class="mt-3 px-4 py-2 rounded-xl bg-sky-600 text-xs font-semibold text-white">Go to Create →</button>
      </div>

      <div id="content-code" class="hidden space-y-5">
        <div class="glass rounded-2xl p-5 space-y-3">
          <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
            <div class="flex flex-wrap gap-2">
              <button onclick="switchCodeMode('primary')" id="btn-code-primary"
                class="px-3.5 py-2 rounded-xl bg-emerald-600 text-white text-xs font-bold">
                💻 Main Project Code
              </button>
              <button onclick="switchCodeMode('arduino')" id="btn-code-arduino"
                class="px-3.5 py-2 rounded-xl bg-slate-800 text-slate-300 text-xs font-semibold border border-slate-700">
                🔌 Optional Arduino Dump Code (.ino)
              </button>
            </div>
            <button onclick="copySourceCode()" class="px-3 py-1.5 rounded-lg bg-slate-800 text-xs text-sky-300 border border-slate-700 self-start">
              📋 Copy Code
            </button>
          </div>
          <p id="active-code-subtitle" class="text-xs text-emerald-300 font-semibold"></p>
          <pre id="code-block" class="bg-slate-950 p-4 rounded-xl border border-slate-800 text-xs text-emerald-300 overflow-x-auto leading-relaxed"></pre>
        </div>

        <div class="glass rounded-2xl p-5 space-y-4">
          <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <div>
              <h3 class="text-base font-bold text-purple-300">🎓 Viva-Voce Q&amp;A (Select One Language)</h3>
              <p class="text-xs text-slate-400">Tap any language to view focused questions &amp; answers</p>
            </div>
            <div class="flex flex-wrap gap-1.5">
              <button onclick="selectVivaLang('en', this)" class="viva-lang-btn px-3 py-1.5 rounded-lg bg-purple-600 text-white text-xs font-bold">English</button>
              <button onclick="selectVivaLang('kn', this)" class="viva-lang-btn px-3 py-1.5 rounded-lg bg-slate-800 text-slate-300 text-xs font-semibold">ಕನ್ನಡ</button>
              <button onclick="selectVivaLang('hi', this)" class="viva-lang-btn px-3 py-1.5 rounded-lg bg-slate-800 text-slate-300 text-xs font-semibold">हिन्दी</button>
              <button onclick="selectVivaLang('te', this)" class="viva-lang-btn px-3 py-1.5 rounded-lg bg-slate-800 text-slate-300 text-xs font-semibold">తెలుగు</button>
              <button onclick="selectVivaLang('ta', this)" class="viva-lang-btn px-3 py-1.5 rounded-lg bg-slate-800 text-slate-300 text-xs font-semibold">தமிழ்</button>
            </div>
          </div>
          <div id="viva-grid" class="grid grid-cols-1 md:grid-cols-2 gap-3.5"></div>
        </div>
      </div>
    </section>

    <!-- ==================== PAGE 4: REPORT STUDIO WITH FORMATTING TOOLS ==================== -->
    <section id="page-report" class="page-section space-y-5">
      <div id="empty-report" class="glass rounded-2xl p-8 text-center">
        <p class="text-sm text-slate-400">Generate a project first to customize and download your 15+ page VTU .docx report.</p>
        <button onclick="navigatePage('page-home')" class="mt-3 px-4 py-2 rounded-xl bg-sky-600 text-xs font-semibold text-white">Go to Create →</button>
      </div>

      <div id="content-report" class="hidden space-y-5">
        <div class="glass rounded-2xl p-5 border border-emerald-500/30 space-y-4">
          <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <div>
              <span class="text-[11px] uppercase tracking-wider text-emerald-400 font-bold">VTU / College Report Formatting Tools</span>
              <h2 id="rep-title" class="text-lg sm:text-xl font-bold text-white mt-0.5"></h2>
            </div>
            <button onclick="downloadDocxDirect()" id="download-docx-btn"
              class="px-5 py-3 rounded-xl bg-gradient-to-r from-emerald-500 to-teal-600 font-bold text-white text-xs sm:text-sm shadow-lg shrink-0">
              📥 Download 15+ Page .docx Report
            </button>
          </div>

          <div class="grid grid-cols-2 sm:grid-cols-4 gap-3 pt-2 border-t border-slate-800">
            <div>
              <label class="text-[11px] text-slate-400">Font Style</label>
              <select id="fmt-font-family" onchange="updateLivePreviewStyle()" class="w-full mt-1 px-2.5 py-2 rounded-lg bg-slate-950 border border-slate-700 text-xs text-white">
                <option value="Times New Roman">Times New Roman (VTU)</option>
                <option value="Arial">Arial</option>
                <option value="Calibri">Calibri</option>
              </select>
            </div>
            <div>
              <label class="text-[11px] text-slate-400">Body Font Size</label>
              <select id="fmt-font-size" onchange="updateLivePreviewStyle()" class="w-full mt-1 px-2.5 py-2 rounded-lg bg-slate-950 border border-slate-700 text-xs text-white">
                <option value="12" selected>12 pt (Standard)</option>
                <option value="11">11 pt (Compact)</option>
                <option value="14">14 pt (Large)</option>
              </select>
            </div>
            <div>
              <label class="text-[11px] text-slate-400">Line Spacing</label>
              <select id="fmt-line-spacing" onchange="updateLivePreviewStyle()" class="w-full mt-1 px-2.5 py-2 rounded-lg bg-slate-950 border border-slate-700 text-xs text-white">
                <option value="1.5" selected>1.5x Spacing (VTU)</option>
                <option value="1.15">1.15x Spacing</option>
                <option value="2.0">2.0x Double Space</option>
              </select>
            </div>
            <div>
              <label class="text-[11px] text-slate-400">Page Border</label>
              <select id="fmt-page-border" class="w-full mt-1 px-2.5 py-2 rounded-lg bg-slate-950 border border-slate-700 text-xs text-white">
                <option value="true" selected>Classic Blue Border</option>
                <option value="false">No Border (Clean)</option>
              </select>
            </div>
          </div>
        </div>

        <div id="live-doc-preview" class="glass rounded-2xl p-6 space-y-4 border-2 border-blue-900/60">
          <h3 class="text-base font-bold text-sky-300">Live Document Preview</h3>
          <p id="rep-abstract" class="text-sm text-slate-200 leading-relaxed"></p>
          <p id="rep-working" class="text-sm text-slate-300 leading-relaxed border-t border-slate-800 pt-3"></p>
        </div>

        <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
          <div class="glass rounded-2xl p-5">
            <h3 class="text-sm font-bold text-emerald-300 mb-2">💰 Bill of Materials (BOM)</h3>
            <div id="rep-bom" class="text-xs text-slate-300 space-y-2"></div>
          </div>
          <div class="glass rounded-2xl p-5">
            <h3 class="text-sm font-bold text-indigo-300 mb-2">📊 Seminar PPT Slide Outline</h3>
            <div id="rep-ppt" class="text-xs text-slate-300 space-y-2"></div>
          </div>
        </div>
      </div>
    </section>
  </main>

  <!-- MOBILE BOTTOM NAVIGATION -->
  <nav class="md:hidden fixed bottom-0 inset-x-0 z-40 glass border-t border-slate-800 grid grid-cols-4 py-1.5 px-1">
    <button onclick="navigatePage('page-home')" id="nav-m-home" class="nav-m-btn flex flex-col items-center py-1 text-sky-400 font-bold">
      <span class="text-base">🏠</span><span class="text-[10px]">Create</span>
    </button>
    <button onclick="navigatePage('page-3d')" id="nav-m-3d" class="nav-m-btn flex flex-col items-center py-1 text-slate-400">
      <span class="text-base">🧊</span><span class="text-[10px]">Circuit/Flow</span>
    </button>
    <button onclick="navigatePage('page-code')" id="nav-m-code" class="nav-m-btn flex flex-col items-center py-1 text-slate-400">
      <span class="text-base">💻</span><span class="text-[10px]">Code/Viva</span>
    </button>
    <button onclick="navigatePage('page-report')" id="nav-m-report" class="nav-m-btn flex flex-col items-center py-1 text-slate-400">
      <span class="text-base">📄</span><span class="text-[10px]">Report</span>
    </button>
  </nav>

  <script>
    // Branch -> Subjects Mapping
    const BRANCH_SUBJECTS = {
      "Computer Science & Engineering (CSE)": [
        "Artificial Intelligence & Machine Learning",
        "Database Management Systems (DBMS)",
        "Computer Networks & Cybersecurity",
        "Operating Systems & System Software",
        "Full-Stack Web & Cloud Computing",
        "Internet of Things (IoT) & Embedded Systems",
        "Digital Image Processing & Computer Vision"
      ],
      "Electronics & Communication (ECE)": [
        "Analog Electronic Circuits (Op-Amp & Timers)",
        "Embedded Systems & Microcontrollers",
        "Digital Signal Processing (DSP)",
        "VLSI Design & Verilog HDL",
        "IoT & Wireless Sensor Networks",
        "Control Systems & Robotics",
        "Digital Communication Systems"
      ],
      "Artificial Intelligence & ML (AIML)": [
        "Deep Learning & Neural Networks (CNN/RNN)",
        "Natural Language Processing (NLP)",
        "Computer Vision & Video Analytics",
        "Data Science & Predictive Modeling",
        "Edge AI & TinyML Hardware"
      ],
      "Electrical & Electronics (EEE)": [
        "Power Electronics & Motor Drives",
        "Microcontrollers & PLC Automation",
        "Renewable Energy & Smart Grid",
        "Analog & Digital Instrumentation",
        "Electric Vehicles (EV) & Battery Management"
      ],
      "Information Science (ISE)": [
        "Software Engineering & Web Applications",
        "Database Management Systems (DBMS)",
        "Cloud Computing & DevOps",
        "Machine Learning & Data Mining",
        "Information & Network Security"
      ],
      "Mechanical Engineering": [
        "Mechatronics & Industrial Automation",
        "Robotics & Kinematics",
        "CAD/CAM & Additive Manufacturing",
        "Thermal & Fluid Engineering Simulation"
      ],
      "Civil Engineering": [
        "Smart Cities & Structural Health Monitoring",
        "Environmental & Water Resource Engineering",
        "Transportation & Traffic Automation",
        "Geotechnical & Surveying Systems"
      ]
    };

    function updateSubjectsForBranch() {
      const branch = document.getElementById("branch-input").value;
      const subSelect = document.getElementById("subject-input");
      const list = BRANCH_SUBJECTS[branch] || ["Core Engineering Project"];
      subSelect.innerHTML = list.map(s => `<option value="${s}">${s}</option>`).join("");
    }

    function applyQuickExample(branch, subject, topic) {
      document.getElementById("branch-input").value = branch;
      updateSubjectsForBranch();
      document.getElementById("subject-input").value = subject;
      document.getElementById("topic-input").value = topic;
    }

    let currentBlueprint = null;
    let studentProfile = JSON.parse(localStorage.getItem("pf_student") || "null");
    let activeCodeMode = "primary";
    let activeVivaLang = "en";

    window.addEventListener("DOMContentLoaded", () => {
      updateSubjectsForBranch();
      setTimeout(() => {
        const s = document.getElementById("splash-screen");
        s.style.opacity = "0"; s.style.visibility = "hidden";
        if (!studentProfile || !studentProfile.email) openLoginModal();
        else updateProfileUI();
      }, 2000);

      initThreeJS();
      build3DSceneFromBlueprint(null);
    });

    function navigatePage(pageId) {
      document.querySelectorAll(".page-section").forEach(sec => sec.classList.remove("active"));
      document.getElementById(pageId).classList.add("active");

      const key = pageId.replace("page-", "");
      document.querySelectorAll(".nav-d-btn").forEach(b => {
        b.classList.remove("bg-sky-600", "text-white"); b.classList.add("text-slate-300");
      });
      const dBtn = document.getElementById("nav-d-" + key);
      if (dBtn) { dBtn.classList.add("bg-sky-600", "text-white"); dBtn.classList.remove("text-slate-300"); }

      document.querySelectorAll(".nav-m-btn").forEach(b => {
        b.classList.remove("text-sky-400", "font-bold"); b.classList.add("text-slate-400");
      });
      const mBtn = document.getElementById("nav-m-" + key);
      if (mBtn) { mBtn.classList.add("text-sky-400", "font-bold"); mBtn.classList.remove("text-slate-400"); }

      if (pageId === "page-3d") setTimeout(resizeThreeJS, 50);
      window.scrollTo({ top: 0, behavior: "smooth" });
    }

    // Profile & Auth
    function toggleProfileMenu() { document.getElementById("profile-dropdown").classList.toggle("hidden"); }
    function closeProfileMenu() { document.getElementById("profile-dropdown").classList.add("hidden"); }
    window.addEventListener("click", (e) => {
      if (!e.target.closest("#profile-trigger") && !e.target.closest("#profile-dropdown")) closeProfileMenu();
    });

    function openLoginModal() {
      const m = document.getElementById("login-modal"); m.classList.remove("hidden"); m.classList.add("flex");
    }
    function closeLoginModal() {
      const m = document.getElementById("login-modal"); m.classList.add("hidden"); m.classList.remove("flex");
    }
    function skipLogin() {
      studentProfile = { email: "", name: "Guest Student", usn: "1VT23CS001", college: "VTU Affiliated Engineering College" };
      updateProfileUI(); closeLoginModal();
    }
    function logoutUser() {
      localStorage.removeItem("pf_student");
      studentProfile = null; currentBlueprint = null;
      document.getElementById("login-form").reset();
      updateProfileUI(); navigatePage("page-home"); openLoginModal();
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
      updateProfileUI(); closeLoginModal();
    });
    function updateProfileUI() {
      const name = (studentProfile && studentProfile.name) || "Guest Student";
      const usn = (studentProfile && studentProfile.usn) || "1VT23CS001";
      document.getElementById("avatar-initial").innerText = name.charAt(0).toUpperCase();
      document.getElementById("profile-short-name").innerText = name.split(" ")[0];
      document.getElementById("menu-full-name").innerText = name;
      document.getElementById("menu-usn").innerText = usn;
    }

    function getReferralCode() {
      if (studentProfile && studentProfile.usn && studentProfile.usn !== "1VT23CS001") return "PF-" + studentProfile.usn;
      return "PF-STUDENT";
    }
    function openReferModal() {
      document.getElementById("my-ref-code").innerText = getReferralCode();
      const m = document.getElementById("refer-modal"); m.classList.remove("hidden"); m.classList.add("flex");
    }
    function closeReferModal() {
      const m = document.getElementById("refer-modal"); m.classList.add("hidden"); m.classList.remove("flex");
    }
    function copyRefCode() { navigator.clipboard.writeText(getReferralCode()); alert("Code copied!"); }
    function copyInviteLink() { navigator.clipboard.writeText(`${window.location.origin}/?ref=${getReferralCode()}`); alert("Invite link copied!"); }
    function shareOnWhatsApp() {
      const link = `${window.location.origin}/?ref=${getReferralCode()}`;
      window.open("https://wa.me/?text=" + encodeURIComponent(`🚀 Build Subject-Aligned Engineering Projects, 3D Circuits & VTU Reports:\n${link}`), "_blank");
    }

    async function generateProject() {
      const topic = document.getElementById("topic-input").value.trim();
      const branch = document.getElementById("branch-input").value;
      const subject = document.getElementById("subject-input").value;
      const level = document.getElementById("level-input").value;
      if (!topic) return alert("Please enter your project topic first!");

      const btn = document.getElementById("generate-btn");
      btn.disabled = true; btn.innerText = "⏳ Building Subject-Aligned Project...";
      try {
        const res = await fetch("/api/generate", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ topic, branch, subject, level })
        });
        const data = await res.json();
        if (data.error) throw new Error(data.error);
        currentBlueprint = data.blueprint;
        populateAllPages(currentBlueprint);
        navigatePage("page-3d");
      } catch (err) {
        alert("Error: " + err.message);
      } finally {
        btn.disabled = false; btn.innerText = "⚡ Generate Subject-Aligned Project";
      }
    }

    function populateAllPages(bp) {
      // Conditional Circuit vs Software Architecture
      if (bp.requires_circuit) {
        document.getElementById("circuit-lab-wrapper").classList.remove("hidden");
        document.getElementById("software-arch-wrapper").classList.add("hidden");
        document.getElementById("lab3d-title").innerText = `3D Circuit: ${bp.title}`;
        build3DSceneFromBlueprint(bp);
        document.getElementById("lab3d-pins").innerHTML = (bp.pin_connections || []).map(p => `
          <tr>
            <td class="p-2.5 font-mono text-sky-300">${p.from_pin}</td>
            <td class="p-2.5 font-mono text-indigo-300">${p.to_pin}</td>
            <td class="p-2.5">${p.signal}</td>
          </tr>
        `).join("");
      } else {
        document.getElementById("circuit-lab-wrapper").classList.add("hidden");
        document.getElementById("software-arch-wrapper").classList.remove("hidden");
        document.getElementById("sw-arch-title").innerText = `System Flow: ${bp.title}`;
        document.getElementById("sw-nodes-container").innerHTML = (bp.software_architecture_nodes || []).map((n, i) => `
          <div class="bg-slate-900 border border-indigo-500/40 rounded-xl p-4 text-center">
            <span class="text-[10px] font-bold text-sky-400">STAGE 0${i + 1}</span>
            <p class="text-xs font-semibold text-white mt-1">${n}</p>
          </div>
        `).join("");
      }

      // Code & Viva
      document.getElementById("empty-code").classList.add("hidden");
      document.getElementById("content-code").classList.remove("hidden");
      const ardBtn = document.getElementById("btn-code-arduino");
      if (bp.has_arduino_hardware) ardBtn.classList.remove("hidden");
      else ardBtn.classList.add("hidden");
      switchCodeMode("primary");
      renderVivaCards();

      // Report Page
      document.getElementById("empty-report").classList.add("hidden");
      document.getElementById("content-report").classList.remove("hidden");
      document.getElementById("rep-title").innerText = bp.title || "";
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
    }

    function switchCodeMode(mode) {
      if (!currentBlueprint) return;
      activeCodeMode = mode;
      const pBtn = document.getElementById("btn-code-primary");
      const aBtn = document.getElementById("btn-code-arduino");
      if (mode === "primary") {
        pBtn.className = "px-3.5 py-2 rounded-xl bg-emerald-600 text-white text-xs font-bold";
        aBtn.className = "px-3.5 py-2 rounded-xl bg-slate-800 text-slate-300 text-xs font-semibold border border-slate-700";
        document.getElementById("active-code-subtitle").innerText = `Language: ${currentBlueprint.primary_code_language || "Code"}`;
        document.getElementById("code-block").innerText = currentBlueprint.primary_code || "";
      } else {
        aBtn.className = "px-3.5 py-2 rounded-xl bg-sky-600 text-white text-xs font-bold";
        pBtn.className = "px-3.5 py-2 rounded-xl bg-slate-800 text-slate-300 text-xs font-semibold border border-slate-700";
        document.getElementById("active-code-subtitle").innerText = "Optional Hardware Firmware (.ino) for Arduino / ESP32";
        document.getElementById("code-block").innerText = currentBlueprint.arduino_code || "";
      }
    }

    function copySourceCode() {
      navigator.clipboard.writeText(document.getElementById("code-block").innerText);
      alert("Code copied!");
    }

    function selectVivaLang(lang, btn) {
      activeVivaLang = lang;
      document.querySelectorAll(".viva-lang-btn").forEach(b => {
        b.classList.remove("bg-purple-600", "text-white", "font-bold");
        b.classList.add("bg-slate-800", "text-slate-300", "font-semibold");
      });
      btn.classList.remove("bg-slate-800", "text-slate-300", "font-semibold");
      btn.classList.add("bg-purple-600", "text-white", "font-bold");
      renderVivaCards();
    }

    function renderVivaCards() {
      if (!currentBlueprint) return;
      const list = currentBlueprint.viva_qa || [];
      document.getElementById("viva-grid").innerHTML = list.map((item, idx) => {
        const entry = item[activeVivaLang] || item["en"] || {};
        return `
          <div class="bg-slate-900/95 p-4 rounded-xl border border-slate-800 space-y-1.5">
            <p class="text-xs font-bold text-sky-300">Q${idx + 1}. ${entry.q || ""}</p>
            <p class="text-xs text-slate-200 leading-relaxed">${entry.a || ""}</p>
          </div>
        `;
      }).join("");
    }

    function updateLivePreviewStyle() {
      const box = document.getElementById("live-doc-preview");
      box.style.fontFamily = document.getElementById("fmt-font-family").value;
      box.style.fontSize = document.getElementById("fmt-font-size").value + "px";
      box.style.lineHeight = document.getElementById("fmt-line-spacing").value;
    }

    async function downloadDocxDirect() {
      if (!currentBlueprint) return;
      const btn = document.getElementById("download-docx-btn");
      btn.disabled = true; btn.innerText = "⏳ Building 15+ Page DOCX...";
      try {
        const payload = {
          blueprint: currentBlueprint,
          student_name: (studentProfile && studentProfile.name) || "Student Name",
          usn: (studentProfile && studentProfile.usn) || "1VT23CS001",
          college: (studentProfile && studentProfile.college) || "VTU Affiliated Engineering College",
          branch: document.getElementById("branch-input").value,
          subject: document.getElementById("subject-input").value,
          guide_name: document.getElementById("guide-input").value.trim() || "Project Guide",
          formatting: {
            font_family: document.getElementById("fmt-font-family").value,
            font_size: document.getElementById("fmt-font-size").value,
            line_spacing: document.getElementById("fmt-line-spacing").value,
            page_border: document.getElementById("fmt-page-border").value === "true",
            viva_lang: activeVivaLang
          }
        };
        const res = await fetch("/api/download-docx", {
          method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload)
        });
        if (!res.ok) throw new Error("Failed");
        const blob = await res.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement("a"); a.href = url; a.download = "ProjForge_15Page_VTU_Report.docx";
        document.body.appendChild(a); a.click(); a.remove();
      } catch (err) {
        alert("Download error: " + err.message);
      } finally {
        btn.disabled = false; btn.innerText = "📥 Download 15+ Page .docx Report";
      }
    }

    // ==================== TINKERCAD-STYLE LIGHT WORKBENCH 3D ENGINE ====================
    let scene, camera, renderer, controls, circuitGroup;
    let trackedLabels = [], animParts = [], extraSlotCount = 0;
    let simActive = true;

    function initThreeJS() {
      const container = document.getElementById("threejs-container");
      if (!container || typeof THREE === "undefined") return;

      scene = new THREE.Scene();
      scene.background = new THREE.Color(0xe2e8f0); // Clean Tinkercad Light Lab Canvas

      camera = new THREE.PerspectiveCamera(42, container.clientWidth / container.clientHeight, 0.1, 1000);
      camera.position.set(0, 8.5, 11.5);

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
        controls.minDistance = 5;
        controls.maxDistance = 20;
      }

      scene.add(new THREE.AmbientLight(0xffffff, 0.85));
      const keyLight = new THREE.DirectionalLight(0xffffff, 0.9);
      keyLight.position.set(8, 15, 10);
      scene.add(keyLight);

      const grid = new THREE.GridHelper(22, 22, 0x94a3b8, 0xcbd5e1);
      grid.position.y = -0.22;
      scene.add(grid);

      circuitGroup = new THREE.Group();
      scene.add(circuitGroup);
      animate3D();
    }

    function createPartMesh(type) {
      const grp = new THREE.Group();
      if (type === "ic_opamp") {
        // 14-Pin DIP IC Chip straddling center groove
        const body = new THREE.Mesh(new THREE.BoxGeometry(1.1, 0.22, 0.55), new THREE.MeshStandardMaterial({ color: 0x1e293b }));
        grp.add(body);
        for (let px = -0.45; px <= 0.45; px += 0.15) {
          const pin = new THREE.Mesh(new THREE.BoxGeometry(0.04, 0.25, 0.68), new THREE.MeshStandardMaterial({ color: 0x94a3b8 }));
          pin.position.set(px, -0.05, 0);
          grp.add(pin);
        }
      } else if (type === "sensor_water_probes") {
        // Transparent Water Tank + Submerged Copper Probes
        const tank = new THREE.Mesh(
          new THREE.BoxGeometry(1.3, 0.95, 0.9),
          new THREE.MeshStandardMaterial({ color: 0x38bdf8, transparent: true, opacity: 0.45 })
        );
        const water = new THREE.Mesh(
          new THREE.BoxGeometry(1.22, 0.55, 0.82),
          new THREE.MeshStandardMaterial({ color: 0x0284c7, transparent: true, opacity: 0.75 })
        );
        water.position.y = -0.15;
        grp.add(tank, water);
        animParts.push({ mesh: water, mode: "water" });
      } else if (type === "led_bar") {
        // Red, Yellow, Green Indicator LEDs + Piezo
        [0x22c55e, 0xeab308, 0xef4444].forEach((clr, idx) => {
          const led = new THREE.Mesh(new THREE.SphereGeometry(0.14, 14, 14), new THREE.MeshBasicMaterial({ color: clr }));
          led.position.set(-0.35 + idx * 0.35, 0.2, 0);
          grp.add(led);
          if (idx === 2) animParts.push({ mesh: led, mode: "blink" });
        });
      } else if (type === "sensor_ultrasonic") {
        const pcb = new THREE.Mesh(new THREE.BoxGeometry(1.4, 0.65, 0.12), new THREE.MeshStandardMaterial({ color: 0x0284c7 }));
        const eyeL = new THREE.Mesh(new THREE.CylinderGeometry(0.24, 0.24, 0.35, 16), new THREE.MeshStandardMaterial({ color: 0xe2e8f0, metalness: 0.7 }));
        eyeL.rotation.x = Math.PI / 2; eyeL.position.set(-0.36, 0, 0.18);
        const eyeR = eyeL.clone(); eyeR.position.set(0.36, 0, 0.18);
        grp.add(pcb, eyeL, eyeR);
      } else if (type === "motor_servo") {
        const body = new THREE.Mesh(new THREE.BoxGeometry(1.1, 0.85, 0.55), new THREE.MeshStandardMaterial({ color: 0x1e293b }));
        const arm = new THREE.Mesh(new THREE.BoxGeometry(1.5, 0.12, 0.18), new THREE.MeshStandardMaterial({ color: 0xf8fafc }));
        arm.position.set(0, 0.52, 0);
        grp.add(body, arm);
        animParts.push({ mesh: arm, mode: "servo" });
      } else if (type === "piezo_buzzer") {
        const drum = new THREE.Mesh(new THREE.CylinderGeometry(0.4, 0.4, 0.4, 20), new THREE.MeshStandardMaterial({ color: 0x1e293b }));
        grp.add(drum);
      } else {
        const mod = new THREE.Mesh(new THREE.BoxGeometry(1.1, 0.6, 0.65), new THREE.MeshStandardMaterial({ color: 0x0d9488 }));
        grp.add(mod);
      }
      return grp;
    }

    function add3DWire(p1, p2, p3, colorHex) {
      const curve = new THREE.CatmullRomCurve3([new THREE.Vector3(...p1), new THREE.Vector3(...p2), new THREE.Vector3(...p3)]);
      const geo = new THREE.TubeGeometry(curve, 24, 0.042, 8, false);
      const mat = new THREE.MeshStandardMaterial({ color: colorHex });
      circuitGroup.add(new THREE.Mesh(geo, mat));
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
      trackedLabels = []; animParts = []; extraSlotCount = 0;
      document.getElementById("labels-overlay").innerHTML = "";

      const supplyLabel = (bp && bp.controller_or_supply) || "Arduino Uno R3";
      const isAnalogBattery = supplyLabel.toLowerCase().includes("9v") || supplyLabel.toLowerCase().includes("analog");

      if (isAnalogBattery) {
        // 9V DC Battery Block (Tinkercad style)
        const bat = new THREE.Mesh(new THREE.BoxGeometry(2.1, 0.95, 1.3), new THREE.MeshStandardMaterial({ color: 0x1e293b }));
        bat.position.set(-3.4, 0.35, 0);
        const cap = new THREE.Mesh(new THREE.BoxGeometry(2.12, 0.35, 1.32), new THREE.MeshStandardMaterial({ color: 0xd97706 }));
        cap.position.set(-3.4, 0.68, 0);
        circuitGroup.add(bat, cap);
        registerLabel("9V DC Power Supply", -3.4, 1.25, 0);
      } else {
        // Arduino Uno R3 Blue PCB
        const mcu = new THREE.Mesh(new THREE.BoxGeometry(3.3, 0.22, 2.2), new THREE.MeshStandardMaterial({ color: 0x0284c7, roughness: 0.35 }));
        mcu.position.set(-3.2, 0, 0);
        const usb = new THREE.Mesh(new THREE.BoxGeometry(0.65, 0.36, 0.55), new THREE.MeshStandardMaterial({ color: 0x94a3b8, metalness: 0.7 }));
        usb.position.set(-4.65, 0.22, -0.45);
        const chip = new THREE.Mesh(new THREE.BoxGeometry(1.35, 0.14, 0.42), new THREE.MeshStandardMaterial({ color: 0x0f172a }));
        chip.position.set(-2.9, 0.16, 0.15);
        circuitGroup.add(mcu, usb, chip);
        registerLabel(supplyLabel, -3.2, 0.85, 0);
      }

      // 830-Point Breadboard
      const bb = new THREE.Mesh(new THREE.BoxGeometry(5.6, 0.26, 2.8), new THREE.MeshStandardMaterial({ color: 0xffffff, roughness: 0.4 }));
      bb.position.set(2.3, 0, 0);
      const groove = new THREE.Mesh(new THREE.BoxGeometry(5.4, 0.03, 0.24), new THREE.MeshStandardMaterial({ color: 0x94a3b8 }));
      groove.position.set(2.3, 0.135, 0);
      const railRed = new THREE.Mesh(new THREE.BoxGeometry(5.2, 0.02, 0.07), new THREE.MeshBasicMaterial({ color: 0xef4444 }));
      railRed.position.set(2.3, 0.14, -1.2);
      const railBlue = new THREE.Mesh(new THREE.BoxGeometry(5.2, 0.02, 0.07), new THREE.MeshBasicMaterial({ color: 0x2563eb }));
      railBlue.position.set(2.3, 0.14, 1.2);
      circuitGroup.add(bb, groove, railRed, railBlue);

      const holeMat = new THREE.MeshBasicMaterial({ color: 0x475569 });
      const holeGeo = new THREE.BoxGeometry(0.06, 0.02, 0.06);
      for (let gx = -0.1; gx <= 4.7; gx += 0.32) {
        for (let gz of [-0.85, -0.55, 0.55, 0.85]) {
          const h = new THREE.Mesh(holeGeo, holeMat);
          h.position.set(gx, 0.14, gz);
          circuitGroup.add(h);
        }
      }
      registerLabel("830-Pt Breadboard", 2.3, 0.55, 1.35);

      // Power Wires (+VCC Red & -GND Black)
      add3DWire([-2.2, 0.3, -0.4], [-0.5, 1.35, -0.9], [0.2, 0.15, -1.2], 0xef4444);
      add3DWire([-2.2, 0.3, 0.4], [-0.5, 1.25, 0.9], [0.2, 0.15, 1.2], 0x1e293b);

      const parts = (bp && bp.circuit_3d_parts && bp.circuit_3d_parts.length) ? bp.circuit_3d_parts : [
        { type: "ic_opamp", label: "LM324 Op-Amp IC", wire_color: "#38bdf8" },
        { type: "sensor_water_probes", label: "Level Sensor Probes", wire_color: "#22c55e" },
        { type: "led_bar", label: "LED & Buzzer Output", wire_color: "#f59e0b" }
      ];

      const xSlots = [0.7, 2.3, 3.9];
      parts.slice(0, 3).forEach((part, i) => {
        const px = xSlots[i];
        const pz = part.type === "ic_opamp" ? 0 : (i % 2 === 0 ? -0.55 : 0.55);
        const py = part.type === "ic_opamp" ? 0.24 : 0.55;
        const mesh = createPartMesh(part.type);
        mesh.position.set(px, py, pz);
        circuitGroup.add(mesh);

        add3DWire([0.7, 0.28, 0], [(px + 0.7) / 2, 1.45 + i * 0.2, pz], [px, 0.2, pz], part.wire_color || "#38bdf8");
        registerLabel(part.label || "Module", px, py + 0.75, pz);
      });
    }

    function addExtra3DPart(type, label) {
      if (!circuitGroup) return;
      const offsetX = 0.8 + (extraSlotCount % 4) * 1.1;
      extraSlotCount++;
      const mesh = createPartMesh(type);
      mesh.position.set(offsetX, 0.55, -1.85);
      circuitGroup.add(mesh);
      add3DWire([offsetX, 0.25, -1.85], [offsetX, 1.2, -1.5], [offsetX, 0.15, -1.2], "#10b981");
      registerLabel(label, offsetX, 1.25, -1.85);
    }

    let simTick = 0;
    function animate3D() {
      requestAnimationFrame(animate3D);
      if (controls) controls.update();
      if (simActive) {
        simTick += 0.05;
        animParts.forEach(item => {
          if (item.mode === "servo") item.mesh.rotation.y = Math.sin(simTick) * 0.9;
          else if (item.mode === "blink") item.mesh.material.color.setHex(Math.sin(simTick * 2.5) > 0 ? 0xef4444 : 0x7f1d1d);
          else if (item.mode === "water") item.mesh.scale.y = 0.65 + 0.35 * ((Math.sin(simTick * 0.8) + 1) / 2);
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
      btn.innerText = simActive ? "▶ Start Simulation: ON" : "⏸️ Simulation: PAUSED";
      btn.className = simActive
        ? "px-3.5 py-2 rounded-xl bg-emerald-600 font-bold text-white text-xs shadow"
        : "px-3.5 py-2 rounded-xl bg-amber-600 font-bold text-white text-xs shadow";
    }

    function reset3DCamera() {
      if (camera) camera.position.set(0, 8.5, 11.5);
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


# ==================== 4. FLASK API ROUTES ====================
@app.route("/")
def index():
    return render_template_string(HTML_PAGE)


@app.route("/api/generate", methods=["POST"])
def api_generate():
    data = request.get_json(force=True)
    topic = data.get("topic", "").strip()
    branch = data.get("branch", "Computer Science & Engineering (CSE)").strip()
    subject = data.get("subject", "Core Engineering").strip()
    level = data.get("level", "3rd Year Mini-Project").strip()
    if not topic:
        return jsonify({"error": "Please enter a project topic."}), 400
    blueprint = generate_project_blueprint(topic, branch, subject, level)
    return jsonify({"status": "ok", "blueprint": blueprint})


@app.route("/api/download-docx", methods=["POST"])
def download_docx():
    payload = request.get_json(force=True)
    blueprint = payload.get("blueprint", {})
    fmt = payload.get("formatting", {})
    docx_buffer = build_vtu_docx(
        blueprint,
        payload.get("student_name") or "Student Name",
        payload.get("usn") or "1VT23CS001",
        payload.get("college") or "VTU Affiliated Engineering College",
        payload.get("branch") or "Engineering",
        payload.get("subject") or "Engineering Subject",
        payload.get("guide_name") or "Project Guide",
        fmt
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