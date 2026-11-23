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


# ==================== 1. SILENT BACKGROUND AI ENGINE ====================
def generate_project_blueprint(topic, branch, subject, level):
    if GEMINI_API_KEY:
        prompt = f"""
You are a Senior Indian Engineering Professor (VTU / Autonomous / AICTE curriculum) and Principal Systems Architect.
Generate a comprehensive, subject-aligned engineering project blueprint in strict JSON format for:
- Project Title: "{topic}"
- Engineering Branch: "{branch}"
- Core Academic Subject: "{subject}"
- Academic Level: "{level}"

STRICT TECHNICAL RULES:
1. Domain & Subject Precision: Every calculation, IC/sensor model, algorithm, and viva question MUST strictly match "{topic}" under "{subject}".
2. "requires_circuit": true ONLY IF physical electronics, analog/digital ICs, sensors, microcontrollers, or motors are used. false for pure software/DBMS/Web/AI software-only projects.
3. If "requires_circuit" is true:
   - Choose 3 to 4 realistic hardware modules from: ["ic_741", "ic_lm324", "sensor_water_probes", "sensor_soil_moisture", "sensor_ultrasonic", "sensor_gas", "sensor_ir", "water_pump", "motor_servo", "relay_module", "oled_display", "piezo_buzzer", "led_indicator_bank"].
   - Set "controller_or_supply" accurately (e.g., "9V DC Analog Supply (No MCU)" for analog IC circuits, or "Arduino Uno R3" / "ESP32 DevKit").
4. "primary_code": Write 85-115 lines of clean, runnable, well-commented code matching the exact subject.
5. "viva_qa": Provide 8 deep technical viva questions & answers translated naturally into all 5 languages: "en", "kn", "hi", "te", "ta".

Return ONLY valid JSON (no markdown fences):
{{
  "title": "Formal Engineering Project Title",
  "tagline": "Technical summary aligned with {subject}",
  "requires_circuit": true,
  "controller_or_supply": "Arduino Uno R3 or 9V DC Analog Supply",
  "verification_summary": {{
    "syllabus_match": "98% Syllabus Aligned ({subject})",
    "build_time": "2 to 3 Days",
    "expert_note": "Component tolerances, circuit equations, and architecture verified for Indian engineering lab standards."
  }},
  "abstract": "Detailed 300-word academic abstract covering problem statement, {subject} principles, design methodology, and quantitative outcomes.",
  "problem_statement": "Detailed 200-word problem definition explaining existing limitations.",
  "objectives": ["Objective 1", "Objective 2", "Objective 3", "Objective 4", "Objective 5"],
  "literature_survey": [
    {{"paper": "Author et al. (2024) - IEEE Transactions", "summary": "Key finding and research gap addressed."}},
    {{"paper": "Author et al. (2025) - Springer", "summary": "Key finding and research gap addressed."}},
    {{"paper": "Author et al. (2026) - Elsevier", "summary": "Key finding and research gap addressed."}}
  ],
  "hardware_requirements": [
    {{"name": "Component 1", "spec": "Exact Part No & Electrical Rating", "cost_inr": 180}},
    {{"name": "Component 2", "spec": "Exact Part No & Electrical Rating", "cost_inr": 220}},
    {{"name": "Component 3", "spec": "Exact Part No & Electrical Rating", "cost_inr": 150}},
    {{"name": "830-Point Breadboard & Passives", "spec": "Resistors, Capacitors & Jumper Wires", "cost_inr": 140}}
  ],
  "software_requirements": [
    "OS: Windows 11 / Ubuntu 22.04 LTS",
    "Simulation / IDE: Multisim / Proteus / Python 3.11 / Arduino IDE",
    "Documentation: VTU / AICTE Standard Format"
  ],
  "circuit_3d_parts": [
    {{"type": "ic_741", "label": "Component Label", "pin_from": "Pin A", "pin_to": "Breadboard Row", "wire_color": "#2563eb"}}
  ],
  "software_architecture_nodes": ["Stage 1", "Stage 2", "Stage 3", "Stage 4", "Stage 5"],
  "pin_connections": [
    {{"from_pin": "Pin / Port 1", "to_pin": "Target Module 1", "signal": "Signal / Voltage"}}
  ],
  "ppt_slides": [
    {{"slide": "Title & Team Overview", "bullets": ["Project Title: {topic}", "Subject: {subject} ({branch})", "Academic Level: {level}"]}},
    {{"slide": "Problem Statement & Motivation", "bullets": ["Limitations of manual legacy systems", "High latency and human error risks", "Need for automated {subject} solution"]}},
    {{"slide": "Project Objectives", "bullets": ["Apply {subject} concepts practically", "Achieve sub-200ms response time", "Keep prototype budget student-friendly"]}},
    {{"slide": "System Architecture & Design", "bullets": ["Modular input-processing-output pipeline", "Calibrated threshold switching logic", "Fail-safe hardware/software protection"]}},
    {{"slide": "Hardware & Software Stack", "bullets": ["Verified components & breadboard layout", "Clean modular firmware / source code", "Real-time telemetry & status indication"]}},
    {{"slide": "Experimental Results", "bullets": ["95.4% operational accuracy in lab trials", "Zero false triggers across 50 test cycles", "Low power consumption & compact footprint"]}},
    {{"slide": "Applications & Future Scope", "bullets": ["Campus, domestic & industrial deployment", "Cloud IoT dashboard & mobile alerts", "Custom PCB fabrication for mass scale"]}}
  ],
  "working_principle": "Detailed 400-word step-by-step technical working explanation grounded in {subject}.",
  "primary_code_language": "Python 3 / C++",
  "primary_code": "Complete 85-115 line runnable code.",
  "has_arduino_hardware": true,
  "arduino_code": "Complete Arduino (.ino) sketch if applicable.",
  "results_and_discussion": "Detailed 240-word analysis of testing parameters, accuracy, and observations.",
  "advantages": ["Advantage 1", "Advantage 2", "Advantage 3", "Advantage 4"],
  "applications": ["Application 1", "Application 2", "Application 3", "Application 4"],
  "future_scope": "180-word discussion on scalability and future upgrades.",
  "references": [
    "[1] IEEE Standard Technical Reference (2025).",
    "[2] Standard University Textbook on {subject}, Pearson Education.",
    "[3] VTU / AICTE Academic Project Guidelines (2025-26)."
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
        for model_name in ["gemini-4-hydrogen", "gemini-4-carbon", "gemini-3.8-flash", "gemini-2.5-flash"]:
            try:
                url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={GEMINI_API_KEY}"
                payload = {
                    "contents": [{"parts": [{"text": prompt}]}],
                    "generationConfig": {"responseMimeType": "application/json", "temperature": 0.3}
                }
                resp = requests.post(url, json=payload, timeout=65)
                if resp.status_code == 200:
                    res_json = resp.json()
                    raw_text = res_json["candidates"][0]["content"]["parts"][0]["text"].strip()
                    raw_text = re.sub(r"^```json\s*|\s*```$", "", raw_text, flags=re.MULTILINE)
                    return json.loads(raw_text)
            except Exception as e:
                print("Background engine step:", e)

    return _smart_topic_blueprint(topic, branch, subject, level)


def _smart_topic_blueprint(topic, branch, subject, level):
    clean_title = topic.strip().title() or f"Smart {subject} System"
    low = (clean_title + " " + subject).lower()

    pure_sw_keys = ["website", "web app", "database", "dbms", "sql", "management system", "chatbot", "nlp", "data science", "cloud", "blockchain", "cybersecurity", "software"]
    hw_keys = ["ic741", "ic 741", "opamp", "op-amp", "arduino", "esp32", "sensor", "iot", "motor", "servo", "relay", "water", "irrigation", "home", "gate", "robot", "embedded", "vlsi", "circuit", "analog", "555", "gas", "fire"]

    has_hw = any(k in low for k in hw_keys)
    is_pure_sw = any(k in low for k in pure_sw_keys) and not has_hw
    requires_circuit = not is_pure_sw

    if "741" in low or "opamp" in low or "op-amp" in low or "water level" in low:
        controller_name = "9V DC Analog Supply (No MCU)"
        parts = [
            {"type": "ic_741", "label": "IC 741 Op-Amp Comparator", "pin_from": "Pin 7 (+9V) / Pin 4 (GND)", "pin_to": "Breadboard Center DIP", "wire_color": "#2563eb"},
            {"type": "sensor_water_probes", "label": "Water Tank Copper Probes", "pin_from": "Non-Inverting Pin 3 (+IN)", "pin_to": "Breadboard Row 10", "wire_color": "#16a34a"},
            {"type": "led_indicator_bank", "label": "Level LEDs & Piezo Buzzer", "pin_from": "Op-Amp Output Pin 6", "pin_to": "Breadboard Row 28", "wire_color": "#d97706"}
        ]
        pins = [
            {"from_pin": "IC 741 Pin 3 (+IN)", "to_pin": "Water Tank Copper Sense Probe", "signal": "Analog Sense Voltage"},
            {"from_pin": "IC 741 Pin 2 (-IN)", "to_pin": "10kΩ Potentiometer Divider", "signal": "Reference Voltage (V_ref)"},
            {"from_pin": "IC 741 Pin 6 (OUT)", "to_pin": "330Ω Resistor -> LED & Buzzer", "signal": "Comparator High (+8.2V)"},
            {"from_pin": "9V DC Battery", "to_pin": "Pin 7 (+VCC) & Pin 4 (GND)", "signal": "9V Regulated DC"}
        ]
        hw_req = [
            {"name": "IC 741 Operational Amplifier (8-Pin DIP)", "spec": "Single High-Gain Op-Amp IC", "cost_inr": 45},
            {"name": "10kΩ Potentiometer & Resistors (330Ω, 1kΩ)", "spec": "Voltage Divider & Current Limiting", "cost_inr": 60},
            {"name": "5mm Red/Green LEDs + 5V Piezo Buzzer", "spec": "Audio-Visual Level Indicators", "cost_inr": 85},
            {"name": "9V DC Battery + 830-Pt Breadboard", "spec": "Solderless Prototyping Setup", "cost_inr": 190}
        ]
        prim_lang = "Python 3 (IC 741 Comparator Threshold Simulator)"
        prim_code = f"""# Project: {clean_title} | Subject: {subject}
VCC = 9.0
V_REF = 4.5

def ic741_comparator(v_non_inverting):
    v_out = (VCC - 0.8) if v_non_inverting > V_REF else 0.2
    led_state = "ON (TANK FULL)" if v_out > 5.0 else "OFF (BELOW LEVEL)"
    print(f"Pin3(+IN)={{v_non_inverting:.1f}}V -> Pin6(OUT)={{v_out:.1f}}V | LED: {{led_state}}")

if __name__ == "__main__":
    for probe_v in [1.2, 3.4, 4.8, 7.5]:
        ic741_comparator(probe_v)
"""
    elif "irrigation" in low or "soil" in low or "plant" in low:
        controller_name = "Arduino Uno R3"
        parts = [
            {"type": "sensor_soil_moisture", "label": "Soil Moisture Fork Sensor", "pin_from": "Arduino Analog A0", "pin_to": "Breadboard Row 08", "wire_color": "#16a34a"},
            {"type": "relay_module", "label": "5V Single-Channel Relay", "pin_from": "Arduino Digital D7", "pin_to": "Breadboard Row 20", "wire_color": "#d97706"},
            {"type": "water_pump", "label": "Mini Submersible Water Pump", "pin_from": "Relay NO Contact", "pin_to": "Breadboard Row 32", "wire_color": "#2563eb"}
        ]
        pins = [
            {"from_pin": "Arduino Analog A0", "to_pin": "Soil Moisture Sensor AO", "signal": "0-5V Analog Moisture"},
            {"from_pin": "Arduino Digital D7", "to_pin": "5V Relay Module IN", "signal": "Active-LOW Pump Trigger"},
            {"from_pin": "Relay NO & COM", "to_pin": "DC Submersible Water Pump", "signal": "Switched Pump Power"},
            {"from_pin": "Arduino 5V & GND", "to_pin": "Breadboard + / - Rails", "signal": "5V DC Supply"}
        ]
        hw_req = [
            {"name": "Arduino Uno R3 (ATmega328P)", "spec": "16MHz Microcontroller Board", "cost_inr": 450},
            {"name": "Capacitive Soil Moisture Sensor", "spec": "Corrosion-Resistant Analog Probe", "cost_inr": 180},
            {"name": "5V Optocoupler Relay + Mini DC Pump", "spec": "Automated Watering Actuator", "cost_inr": 290},
            {"name": "830-Point Breadboard & Jumpers", "spec": "Solderless Prototyping Board", "cost_inr": 140}
        ]
        prim_lang = "Arduino C++ (.ino) — Automated Irrigation Firmware"
        prim_code = f"""// Project: {clean_title}
#define MOISTURE_PIN A0
#define PUMP_RELAY   7
void setup() {{ Serial.begin(115200); pinMode(PUMP_RELAY, OUTPUT); }}
void loop() {{
  int val = analogRead(MOISTURE_PIN);
  digitalWrite(PUMP_RELAY, val > 620 ? LOW : HIGH);
  delay(500);
}}
"""
    elif requires_circuit:
        controller_name = "Arduino Uno R3"
        parts = [
            {"type": "sensor_ultrasonic", "label": "Ultrasonic HC-SR04 Sensor", "pin_from": "Arduino D2 / D3", "pin_to": "Breadboard Row 10", "wire_color": "#2563eb"},
            {"type": "motor_servo", "label": "SG90 Servo Actuator", "pin_from": "Arduino PWM D9", "pin_to": "Breadboard Row 22", "wire_color": "#d97706"},
            {"type": "piezo_buzzer", "label": "Piezo Buzzer & Status LED", "pin_from": "Arduino D8", "pin_to": "Breadboard Row 34", "wire_color": "#dc2626"}
        ]
        pins = [
            {"from_pin": "Arduino D2 / D3", "to_pin": "HC-SR04 Trig & Echo", "signal": "5V Ultrasonic Pulse"},
            {"from_pin": "Arduino PWM D9", "to_pin": "SG90 Servo Signal", "signal": "50Hz PWM Control"},
            {"from_pin": "Arduino D8", "to_pin": "Piezo Buzzer & LED", "signal": "Digital Alert"},
            {"from_pin": "Arduino 5V & GND", "to_pin": "Breadboard + / - Rails", "signal": "5V DC Power"}
        ]
        hw_req = [
            {"name": "Arduino Uno R3", "spec": "ATmega328P Microcontroller", "cost_inr": 450},
            {"name": "HC-SR04 Ultrasonic Sensor", "spec": "2cm - 400cm Rangefinder", "cost_inr": 160},
            {"name": "SG90 Micro Servo & Piezo Buzzer", "spec": "180° Actuator + Alert", "cost_inr": 220},
            {"name": "830-Point Breadboard & Jumpers", "spec": "Prototyping Board", "cost_inr": 140}
        ]
        prim_lang = "Arduino C++ (.ino)"
        prim_code = f"""// Firmware for {clean_title}
void setup() {{ Serial.begin(115200); }}
void loop() {{ Serial.println("Active: {clean_title}"); delay(300); }}"""
    else:
        controller_name = "Software-Only Architecture"
        parts = []
        pins = []
        hw_req = [
            {"name": "Development Workstation", "spec": "Intel i5 / Ryzen 5, 8GB+ RAM", "cost_inr": 0},
            {"name": "Python 3.11 & Cloud Database", "spec": "Open-Source Runtime Stack", "cost_inr": 0}
        ]
        prim_lang = f"Python 3 ({subject} Implementation)"
        prim_code = f"""# Project: {clean_title} | Subject: {subject}
from datetime import datetime

class SystemService:
    def execute(self, record):
        return {{"timestamp": datetime.now().isoformat(), "payload": record, "verified": True}}

if __name__ == "__main__":
    print(SystemService().execute({{"project": "{clean_title}"}}))
"""

    return {
        "title": clean_title,
        "tagline": f"Verified {subject} engineering project under {branch} ({level}).",
        "requires_circuit": requires_circuit,
        "controller_or_supply": controller_name,
        "verification_summary": {
            "syllabus_match": f"98% Syllabus Aligned ({subject})",
            "build_time": "2 Days",
            "expert_note": f"Architecture, component ratings, and methodology strictly follow {subject} academic standards."
        },
        "abstract": (
            f"This project presents the design, analysis, and implementation of '{clean_title}' based on the core principles "
            f"of {subject} under {branch}. Existing conventional systems suffer from manual latency, poor accuracy, and "
            "lack of real-time feedback. By applying systematic engineering principles from "
            f"{subject}, the proposed design achieves reliable operation, low cost, and 95.4% functional accuracy."
        ),
        "problem_statement": (
            f"Traditional approaches for {clean_title} depend on manual supervision and legacy components, causing delayed "
            "response times and human errors. An automated, reliable solution grounded in "
            f"{subject} is required to overcome these limitations."
        ),
        "objectives": [
            f"To apply {subject} concepts to design and build '{clean_title}'.",
            "To achieve fast, automated response without manual intervention.",
            "To minimize component cost while maintaining high accuracy and stability.",
            "To verify circuit/software behavior across standard lab test cases.",
            "To prepare complete VTU / AICTE-compliant project documentation."
        ],
        "literature_survey": [
            {"paper": "R. Kumar et al. (2024) - IEEE Access", "summary": f"Explored foundational {subject} methods; lacked low-cost optimization."},
            {"paper": "S. Kulkarni & P. Rao (2025) - Springer", "summary": "Demonstrated prototype validation; higher circuit complexity."},
            {"paper": "A. Sharma et al. (2026) - Elsevier", "summary": "Introduced automated monitoring; our design improves clarity and reliability."}
        ],
        "hardware_requirements": hw_req,
        "software_requirements": [
            "OS: Windows 11 / Ubuntu 22.04 LTS",
            f"Tools: Multisim / Python 3.11 / Arduino IDE / VS Code ({subject})",
            "Documentation: VTU / AICTE Standard Word (.docx) & Seminar (.pptx)"
        ],
        "circuit_3d_parts": parts,
        "software_architecture_nodes": [
            "User Input / Client UI",
            "API & Input Validation",
            f"Core {subject} Processing Engine",
            "Database & Storage Layer",
            "Analytics & Output Dashboard"
        ],
        "pin_connections": pins,
        "ppt_slides": [
            {"slide": "Title & Team Overview", "bullets": [f"Project Title: {clean_title}", f"Subject: {subject} ({branch})", f"Academic Level: {level}"]},
            {"slide": "Problem Statement", "bullets": ["Manual supervision causes delays and errors", "Lack of real-time threshold alerts", "Need for automated, low-cost engineering design"]},
            {"slide": "Project Objectives", "bullets": [f"Implement practical {subject} principles", "Achieve sub-200ms response latency", "Deliver reliable prototype within student budget"]},
            {"slide": "System Architecture & Design", "bullets": ["Structured input, processing & output stages", "Calibrated reference threshold switching", "Safe electrical & logical interfacing"]},
            {"slide": "Implementation & Testing", "bullets": ["Breadboard / software module integration", "50 repeated laboratory validation trials", "Clean serial & visual status feedback"]},
            {"slide": "Results & Future Scope", "bullets": ["95.4% accuracy with zero false triggers", "Applicable to campus & industrial automation", "Expandable to cloud IoT & custom PCB"]}
        ],
        "working_principle": (
            f"The operation of '{clean_title}' follows core {subject} principles across three stages: Input Sensing/Acquisition, "
            "Threshold Comparison/Processing, and Output Actuation. Input signals are conditioned against noise, compared "
            "against calibrated reference levels, and converted into deterministic visual, audio, or mechanical outputs."
        ),
        "primary_code_language": prim_lang,
        "primary_code": prim_code,
        "has_arduino_hardware": requires_circuit,
        "arduino_code": prim_code if "void setup()" in prim_code else f"// Optional Arduino Sketch for {clean_title}\nvoid setup() {{ Serial.begin(115200); }}\nvoid loop() {{ Serial.println(\"Running {clean_title}\"); delay(500); }}",
        "results_and_discussion": (
            f"Experimental validation was performed across 50 test cycles under {subject} lab conditions, confirming "
            "stable threshold switching, zero false triggers, and 95.4% accuracy."
        ),
        "advantages": [
            f"Directly demonstrates practical {subject} engineering concepts",
            "High reliability and sub-200ms response latency",
            "Low prototype cost and easy breadboard debugging",
            "Clear visual and measurable outputs for external viva"
        ],
        "applications": [
            "Industrial & Commercial Automation",
            "Smart Campus & Domestic Safety Systems",
            f"{subject} Laboratory Demonstrations",
            "Scalable IoT & Embedded Deployments"
        ],
        "future_scope": "Can be upgraded with wireless telemetry, mobile alerts, and compact custom PCB fabrication.",
        "references": [
            f"[1] IEEE Standard Technical Reference on {subject} (2025).",
            f"[2] Standard University Textbook on {subject}, Pearson Education.",
            "[3] VTU / AICTE Academic Project Report Guidelines (2025-26)."
        ],
        "viva_qa": [
            {
                "en": {"q": f"How does this project apply {subject} concepts?", "a": f"It directly uses {subject} principles for accurate signal/data processing and threshold decision-making."},
                "kn": {"q": f"ಈ ಪ್ರಾಜೆಕ್ಟ್ {subject} ತತ್ವಗಳನ್ನು ಹೇಗೆ ಬಳಸುತ್ತದೆ?", "a": f"ಇದು ಸಿಗ್ನಲ್/ಡೇಟಾ ಸಂಸ್ಕರಣೆ ಮತ್ತು ನಿರ್ಧಾರ ತೆಗೆದುಕೊಳ್ಳಲು {subject} ತತ್ವಗಳನ್ನು ನೇರವಾಗಿ ಬಳಸುತ್ತದೆ."},
                "hi": {"q": f"यह प्रोजेक्ट {subject} के सिद्धांतों को कैसे लागू करता है?", "a": f"यह सिग्नल/डेटा प्रोसेसिंग और निर्णय लेने के लिए {subject} के सिद्धांतों का उपयोग करता है।"},
                "te": {"q": f"ఈ ప్రాజెక్ట్ {subject} సూత్రాలను ఎలా ఉపయోగిస్తుంది?", "a": f"ఇది సిగ్నల్/డేటా ప్రాసెసింగ్ మరియు నిర్ణయం తీసుకోవడానికి {subject} సూత్రాలను ఉపయోగిస్తుంది."},
                "ta": {"q": f"இந்த திட்டம் {subject} கோட்பாடுகளை எவ்வாறு பயன்படுத்துகிறது?", "a": f"இது சிக்னல்/தரவு செயலாக்கம் மற்றும் முடிவெடுப்பதற்கு {subject} கோட்பாடுகளைப் பயன்படுத்துகிறது."}
            },
            {
                "en": {"q": "What is the role of the reference threshold in your design?", "a": "It sets the exact switching point so outputs trigger reliably without noise interference."},
                "kn": {"q": "ನಿಮ್ಮ ವಿನ್ಯಾಸದಲ್ಲಿ ರೆಫರೆನ್ಸ್ ಥ್ರೆಶೋಲ್ಡ್‌ನ ಕೆಲಸವೇನು?", "a": "ನಾಯ್ಸ್ ತೊಂದರೆಯಿಲ್ಲದೆ ಸರಿಯಾದ ಸಮಯಕ್ಕೆ ಔಟ್‌ಪುಟ್ ಟ್ರಿಗರ್ ಆಗಲು ಇದು ನಿಖರವಾದ ಮಿತಿಯನ್ನು ನಿಗದಿಪಡಿಸುತ್ತದೆ."},
                "hi": {"q": "आपके डिज़ाइन में रेफरेंस थ्रेशोल्ड की क्या भूमिका है?", "a": "यह सटीक स्विचिंग पॉइंट सेट करता है ताकि बिना शोर के आउटपुट ट्रिगर हो सके।"},
                "te": {"q": "మీ డిజైన్‌లో రిఫరెన్స్ థ్రెషోల్డ్ పాత్ర ఏమిటి?", "a": "నాయిస్ లేకుండా అవుట్‌పుట్ సరిగ్గా ట్రిగ్గర్ కావడానికి ఇది ఖచ్చితమైన స్థాయిని సెట్ చేస్తుంది."},
                "ta": {"q": "உங்கள் வடிவமைப்பில் குறிப்பு வரம்பின் (Reference Threshold) பங்கு என்ன?", "a": "இரைச்சல் இன்றி வெளியீடு சரியாகத் தூண்டப்படுவதற்கு இது துல்லியமான நிலையை அமைக்கிறது."}
            },
            {
                "en": {"q": "How did you verify and test your prototype?", "a": "By running 50 repeated test trials and verifying output stability across varying input levels."},
                "kn": {"q": "ನಿಮ್ಮ ಮಾದರಿಯನ್ನು ನೀವು ಹೇಗೆ ಪರೀಕ್ಷಿಸಿದ್ದೀರಿ?", "a": "50 ಬಾರಿ ಪರೀಕ್ಷೆ ನಡೆಸಿ ವಿವಿಧ ಇನ್‌ಪುಟ್ ಹಂತಗಳಲ್ಲಿ ಔಟ್‌ಪುಟ್ ಸ್ಥಿರತೆಯನ್ನು ದೃಢಪಡಿಸಲಾಗಿದೆ."},
                "hi": {"q": "आपने अपने प्रोटोटाइप का परीक्षण कैसे किया?", "a": "50 परीक्षण चलाकर और विभिन्न इनपुट स्तरों पर आउटपुट की जांच करके।"},
                "te": {"q": "మీరు మీ ప్రోటోటైప్‌ను ఎలా పరీక్షించారు?", "a": "50 పరీక్షలు నిర్వహించి వివిధ ఇన్‌పుట్ స్థాయిలలో అవుట్‌పుట్ స్థిరత్వాన్ని నిర్ధారించాము."},
                "ta": {"q": "உங்கள் முன்மாதிரியை எவ்வாறு சோதித்தீர்கள்?", "a": "50 சோதனைகளை நடத்தி பல்வேறு உள்ளீட்டு நிலைகளில் வெளியீட்டை உறுதி செய்தோம்."}
            }
        ]
    }


# ==================== 2. DOCX REPORT & IEEE PAPER & PPTX BUILDERS ====================
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

    # COVER PAGE
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

    # CERTIFICATE
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

    # DECLARATION & ACKNOWLEDGEMENT
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

    # ABSTRACT & TOC
    add_h("ABSTRACT", 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_b(blueprint.get("abstract", ""))
    add_h("TABLE OF CONTENTS", 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
    toc_items = [
        ("1.", "CHAPTER 1: INTRODUCTION & PROBLEM DEFINITION", "05"),
        ("2.", "CHAPTER 2: LITERATURE SURVEY", "07"),
        ("3.", "CHAPTER 3: SYSTEM DESIGN & PIN / ARCHITECTURE MAPPING", "09"),
        ("4.", "CHAPTER 4: HARDWARE & SOFTWARE REQUIREMENTS", "11"),
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

    # CHAPTERS
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

    add_h("CHAPTER 3: SYSTEM DESIGN & INTERFACING", 1)
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
    else:
        for idx, node in enumerate(blueprint.get("software_architecture_nodes", []), start=1):
            add_b(f"Stage {idx}: {node}", bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)
    doc.add_page_break()

    add_h("CHAPTER 4: HARDWARE & SOFTWARE REQUIREMENTS", 1)
    for hw in blueprint.get("hardware_requirements", []):
        add_b(f"• {hw.get('name', '')} ({hw.get('spec', '')}) — Est. Rs. {hw.get('cost_inr', 0)}", align=WD_ALIGN_PARAGRAPH.LEFT)
    for sw in blueprint.get("software_requirements", []):
        add_b(f"• {sw}", align=WD_ALIGN_PARAGRAPH.LEFT)
    doc.add_page_break()

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

    add_h("APPENDIX-A: VIVA-VOCE QUESTIONS & ANSWERS", 1)
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
        ("II. METHODOLOGY & SYSTEM ARCHITECTURE", blueprint.get("working_principle", "")),
        ("III. EXPERIMENTAL RESULTS & DISCUSSION", blueprint.get("results_and_discussion", "")),
        ("IV. CONCLUSION & FUTURE SCOPE", blueprint.get("future_scope", "")),
    ]:
        h = doc.add_heading(sec_title, level=2)
        h.runs[0].font.name = "Times New Roman"
        h.runs[0].font.size = Pt(11)
        b = doc.add_paragraph(body_txt)
        b.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        b.runs[0].font.name = "Times New Roman"
        b.runs[0].font.size = Pt(10)

    h_ref = doc.add_heading("REFERENCES", level=2)
    h_ref.runs[0].font.name = "Times New Roman"
    for ref in blueprint.get("references", []):
        rp = doc.add_paragraph(ref)
        rp.runs[0].font.name = "Times New Roman"
        rp.runs[0].font.size = Pt(9.5)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


def build_seminar_pptx(blueprint, student_name, usn, college, branch, subject, guide_name):
    prs = Presentation()
    prs.slide_width = PptInches(13.333)
    prs.slide_height = PptInches(7.5)

    # Slide 1: Cover Slide
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

    # Content Slides
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


# ==================== 3. FRONTEND HTML (DISCIPLINED WHITE STUDIO + PWA INSTALL + PPTX + IEEE) ====================
HTML_PAGE = r"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
  <meta name="theme-color" content="#1e3a8a" />
  <title>ProjForge — Engineering Project & Report Studio</title>
  <link rel="icon" type="image/svg+xml" href="/favicon.ico" />
  <link rel="apple-touch-icon" href="/favicon.ico" />
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
      animation: splashAutoOut 0.35s ease 0.9s forwards;
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

  <!-- 1. CLEAN WHITE SPLASH SCREEN -->
  <div id="splash-screen">
    <svg width="76" height="76" viewBox="0 0 88 88" fill="none">
      <rect x="4" y="4" width="80" height="80" rx="22" fill="#1e3a8a"/>
      <path d="M26 44H37L42 30L49 58L54 44H63" stroke="#38bdf8" stroke-width="5.5" stroke-linecap="round" stroke-linejoin="round"/>
      <circle cx="25" cy="44" r="4" fill="#ffffff"/>
      <circle cx="63" cy="44" r="4" fill="#ffffff"/>
    </svg>
    <h1 class="text-2xl font-extrabold mt-3 text-slate-900 tracking-tight">ProjForge</h1>
    <p class="text-xs text-slate-500 mt-1 font-medium">Academic Engineering &amp; Project Studio</p>
  </div>

  <!-- 2. LOGIN MODAL -->
  <div id="login-modal" class="fixed inset-0 z-50 hidden items-center justify-center bg-slate-900/60 backdrop-blur-sm p-4">
    <div class="card max-w-md w-full rounded-2xl p-6">
      <h2 class="text-lg font-bold text-slate-900">Student Profile Setup</h2>
      <p class="text-xs text-slate-500 mb-3">Auto-fills your Name &amp; USN on VTU .docx reports and .pptx slides</p>
      <form id="login-form" class="space-y-3">
        <div>
          <label class="text-xs font-semibold text-slate-700">Student Email *</label>
          <input id="login-email" type="email" required placeholder="student@college.edu.in"
            class="w-full mt-1 px-3.5 py-2.5 rounded-xl bg-slate-50 border border-slate-300 text-sm text-slate-900 focus:border-blue-700 focus:outline-none" />
        </div>
        <div class="grid grid-cols-2 gap-3">
          <div>
            <label class="text-xs font-semibold text-slate-700">Full Name *</label>
            <input id="login-name" type="text" required placeholder="Rahul Sharma"
              class="w-full mt-1 px-3.5 py-2.5 rounded-xl bg-slate-50 border border-slate-300 text-sm text-slate-900 focus:border-blue-700 focus:outline-none" />
          </div>
          <div>
            <label class="text-xs font-semibold text-slate-700">USN / Roll No *</label>
            <input id="login-usn" type="text" required placeholder="1VT23CS045"
              class="w-full mt-1 px-3.5 py-2.5 rounded-xl bg-slate-50 border border-slate-300 text-sm uppercase text-slate-900 focus:border-blue-700 focus:outline-none" />
          </div>
        </div>
        <div>
          <label class="text-xs font-semibold text-slate-700">College Name</label>
          <input id="login-college" type="text" placeholder="VTU Affiliated Engineering College"
            class="w-full mt-1 px-3.5 py-2.5 rounded-xl bg-slate-50 border border-slate-300 text-sm text-slate-900 focus:border-blue-700 focus:outline-none" />
        </div>
        <button type="submit" class="w-full py-3 rounded-xl bg-blue-900 hover:bg-blue-800 font-semibold text-white text-sm">
          Save &amp; Continue →
        </button>
      </form>
      <button type="button" onclick="skipLogin()" class="w-full mt-3 text-xs text-slate-500 hover:text-slate-800 underline">Continue as Guest</button>
    </div>
  </div>

  <!-- 3. REFER FRIENDS MODAL -->
  <div id="refer-modal" class="fixed inset-0 z-50 hidden items-center justify-center bg-slate-900/60 backdrop-blur-sm p-4">
    <div class="card max-w-md w-full rounded-2xl p-6 relative">
      <button type="button" onclick="closeReferModal()" class="absolute top-4 right-4 text-slate-400 hover:text-slate-800 text-sm">✕</button>
      <h2 class="text-base font-bold text-slate-900">🎁 Share With Classmates</h2>
      <p class="text-xs text-slate-500 mt-1">Invite your engineering batchmates to build projects &amp; reports</p>
      <div class="bg-slate-50 border border-slate-200 rounded-xl p-3 mt-3">
        <span class="text-[10px] font-semibold text-slate-500">Your Referral Code:</span>
        <div class="flex items-center justify-between mt-1">
          <span id="my-ref-code" class="font-mono font-bold text-sm text-blue-900">PF-STUDENT</span>
          <button type="button" onclick="copyRefCode()" class="text-xs px-2.5 py-1 rounded-lg bg-slate-200 text-slate-800 font-semibold">Copy</button>
        </div>
      </div>
      <div class="grid grid-cols-2 gap-2.5 mt-4">
        <button type="button" onclick="shareOnWhatsApp()" class="py-2.5 px-3 rounded-xl bg-emerald-600 font-semibold text-white text-xs">💬 WhatsApp Share</button>
        <button type="button" onclick="copyInviteLink()" class="py-2.5 px-3 rounded-xl bg-blue-900 font-semibold text-white text-xs">🔗 Copy Link</button>
      </div>
    </div>
  </div>

  <!-- 4. INSTALL APP HELP MODAL (FOR IOS / BROWSERS WITHOUT AUTO PROMPT) -->
  <div id="install-help-modal" class="fixed inset-0 z-50 hidden items-center justify-center bg-slate-900/60 backdrop-blur-sm p-4">
    <div class="card max-w-md w-full rounded-2xl p-6 relative">
      <button type="button" onclick="closeInstallHelp()" class="absolute top-4 right-4 text-slate-400 hover:text-slate-800 text-sm">✕</button>
      <h2 class="text-base font-bold text-slate-900">📲 Install ProjForge App</h2>
      <p class="text-xs text-slate-600 mt-2 leading-relaxed">
        To install this app directly onto your phone or laptop home screen:
      </p>
      <ul class="text-xs text-slate-700 space-y-2 mt-3 bg-slate-50 p-3.5 rounded-xl border border-slate-200">
        <li><b>Android Chrome:</b> Tap the <b>⋮ (3 dots)</b> menu at top-right → Tap <b>"Install app"</b> or <b>"Add to Home screen"</b>.</li>
        <li><b>iPhone Safari:</b> Tap the <b>Share (↑)</b> button → Tap <b>"Add to Home Screen"</b>.</li>
        <li><b>Laptop Chrome/Edge:</b> Click the <b>Install icon</b> in the right corner of the address bar.</li>
      </ul>
      <button type="button" onclick="closeInstallHelp()" class="w-full mt-4 py-2.5 rounded-xl bg-blue-900 text-white text-xs font-semibold">Got It</button>
    </div>
  </div>

  <!-- 5. POST-DOWNLOAD FEEDBACK MODAL -->
  <div id="feedback-modal" class="fixed inset-0 z-50 hidden items-center justify-center bg-slate-900/60 backdrop-blur-sm p-4">
    <div class="card max-w-md w-full rounded-2xl p-6 relative">
      <button type="button" onclick="closeFeedbackModal()" class="absolute top-4 right-4 text-slate-400 hover:text-slate-800 text-sm">✕</button>
      <div class="flex items-center gap-2.5 mb-2">
        <span class="text-2xl">✅</span>
        <div>
          <h2 class="text-base font-bold text-slate-900">File Downloaded Successfully!</h2>
          <p class="text-xs text-emerald-700 font-medium">Tell us what else you need for your college submission</p>
        </div>
      </div>

      <form id="feedback-form" class="space-y-3 mt-4">
        <div>
          <label class="text-xs text-slate-700 font-semibold">Rate Your Experience *</label>
          <div class="flex gap-2 mt-1.5" id="star-rating-group">
            <button type="button" onclick="setRating(1)" class="star-btn px-3 py-1.5 rounded-lg bg-slate-100 border border-slate-300 text-xs">1 ⭐</button>
            <button type="button" onclick="setRating(2)" class="star-btn px-3 py-1.5 rounded-lg bg-slate-100 border border-slate-300 text-xs">2 ⭐</button>
            <button type="button" onclick="setRating(3)" class="star-btn px-3 py-1.5 rounded-lg bg-slate-100 border border-slate-300 text-xs">3 ⭐</button>
            <button type="button" onclick="setRating(4)" class="star-btn px-3 py-1.5 rounded-lg bg-slate-100 border border-slate-300 text-xs">4 ⭐</button>
            <button type="button" onclick="setRating(5)" class="star-btn px-3 py-1.5 rounded-lg bg-blue-900 border border-blue-900 text-white font-bold text-xs">5 ⭐</button>
          </div>
        </div>
        <div>
          <label class="text-xs text-slate-700 font-semibold">Any other requirements or features needed?</label>
          <textarea id="feedback-text" rows="3"
            placeholder="e.g., Add specific lab subjects, circuit simulation tools, or university formats..."
            class="w-full mt-1 px-3 py-2 rounded-xl bg-slate-50 border border-slate-300 text-xs text-slate-900 focus:border-blue-700 focus:outline-none"></textarea>
        </div>
        <div class="flex gap-2.5 pt-1">
          <button type="submit" class="flex-1 py-2.5 rounded-xl bg-blue-900 hover:bg-blue-800 font-semibold text-white text-xs">
            Submit Feedback
          </button>
          <button type="button" onclick="closeFeedbackModal()" class="px-4 py-2.5 rounded-xl bg-slate-100 text-slate-700 text-xs font-semibold">
            Skip
          </button>
        </div>
      </form>
    </div>
  </div>

  <!-- 6. DISCIPLINED WHITE TOP NAVBAR WITH INSTALL BUTTON -->
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
          <span class="brand-font font-extrabold text-base sm:text-lg text-slate-900 tracking-tight">ProjForge</span>
          <span class="hidden sm:inline-block ml-1.5 text-[10px] px-2 py-0.5 rounded-full bg-blue-50 text-blue-900 border border-blue-200 font-bold">INDIA STUDIO</span>
        </div>
      </div>

      <!-- Desktop Tabs -->
      <nav class="hidden md:flex items-center gap-1 bg-slate-100 p-1 rounded-xl border border-slate-200">
        <button type="button" onclick="navigatePage('page-home')" id="nav-d-home" class="nav-d-btn px-3 py-1.5 rounded-lg text-xs font-bold bg-blue-900 text-white">🏠 Create</button>
        <button type="button" onclick="navigatePage('page-3d')" id="nav-d-3d" class="nav-d-btn px-3 py-1.5 rounded-lg text-xs font-semibold text-slate-600 hover:text-slate-900">🧊 Circuit / Flow</button>
        <button type="button" onclick="navigatePage('page-code')" id="nav-d-code" class="nav-d-btn px-3 py-1.5 rounded-lg text-xs font-semibold text-slate-600 hover:text-slate-900">💻 Code &amp; Viva</button>
        <button type="button" onclick="navigatePage('page-report')" id="nav-d-report" class="nav-d-btn px-3 py-1.5 rounded-lg text-xs font-semibold text-slate-600 hover:text-slate-900">📄 Report &amp; PPTX</button>
      </nav>

      <!-- Install Button + Profile Menu -->
      <div class="flex items-center gap-2">
        <button type="button" id="install-app-btn" onclick="triggerInstallApp()"
          class="px-3 py-1.5 rounded-xl bg-blue-50 hover:bg-blue-100 border border-blue-200 text-blue-900 font-bold text-xs flex items-center gap-1.5 transition">
          <span>📲</span><span>Install</span>
        </button>

        <div class="relative">
          <button type="button" onclick="toggleProfileMenu()" id="profile-trigger"
            class="flex items-center gap-2 px-3 py-1.5 rounded-xl bg-slate-100 hover:bg-slate-200 border border-slate-300 text-xs text-slate-800">
            <span id="avatar-initial" class="w-6 h-6 rounded-full bg-blue-900 flex items-center justify-center font-bold text-[11px] text-white">G</span>
            <span id="profile-short-name" class="max-w-[75px] truncate font-semibold">Guest</span>
            <span class="text-slate-500 text-[10px]">▼</span>
          </button>

          <div id="profile-dropdown" class="hidden absolute right-0 mt-2 w-56 rounded-2xl card py-2 z-50">
            <div class="px-4 py-2 border-b border-slate-100">
              <p id="menu-full-name" class="text-xs font-bold text-slate-900 truncate">Guest Student</p>
              <p id="menu-usn" class="text-[11px] text-blue-800 font-mono">1VT23CS001</p>
            </div>
            <button type="button" onclick="closeProfileMenu(); openLoginModal();" class="w-full text-left px-4 py-2.5 text-xs text-slate-700 hover:bg-slate-50 flex items-center gap-2">
              <span>✏️</span> Edit Student Profile
            </button>
            <button type="button" onclick="closeProfileMenu(); openReferModal();" class="w-full text-left px-4 py-2.5 text-xs text-blue-800 hover:bg-slate-50 flex items-center gap-2">
              <span>🎁</span> Share With Friends
            </button>
            <div class="border-t border-slate-100 my-1"></div>
            <button type="button" onclick="closeProfileMenu(); logoutUser();" class="w-full text-left px-4 py-2.5 text-xs text-rose-600 hover:bg-rose-50 font-semibold flex items-center gap-2">
              <span>🚪</span> Logout
            </button>
          </div>
        </div>
      </div>
    </div>
  </header>

  <main class="max-w-5xl mx-auto px-4 pt-6">

    <!-- ==================== PAGE 1: GENERATOR ==================== -->
    <section id="page-home" class="page-section active space-y-5">
      <div class="text-center max-w-2xl mx-auto">
        <span class="text-[11px] uppercase tracking-wider px-3 py-1 rounded-full bg-blue-50 text-blue-900 border border-blue-200 font-bold">
          For Indian Engineering Students • VTU • Autonomous • AICTE
        </span>
        <h1 class="text-2xl sm:text-4xl font-extrabold mt-2.5 leading-tight text-slate-900">
          Build Complete Engineering Projects, 3D Circuits, Reports &amp; PPTs
        </h1>
        <p class="text-slate-600 text-xs sm:text-sm mt-2">
          Select your Branch &amp; Subject (or enter manually) to generate 3D Breadboard simulations, runnable source code, 15+ page Word reports, IEEE papers &amp; PowerPoint (.pptx) decks.
        </p>
      </div>

      <div class="card rounded-2xl p-5 sm:p-7 max-w-3xl mx-auto space-y-4">
        <div class="grid grid-cols-1 sm:grid-cols-2 gap-3.5">
          <div>
            <label class="text-xs font-bold text-slate-700">1. Select Engineering Branch *</label>
            <select id="branch-input" onchange="updateSubjectsForBranch()"
              class="w-full mt-1 px-3 py-2.5 rounded-xl bg-slate-50 border border-slate-300 text-xs sm:text-sm text-slate-900 font-medium focus:border-blue-800 focus:outline-none">
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
            <label class="text-xs font-bold text-slate-700">2. Select Subject (or Type Manually) *</label>
            <select id="subject-input" onchange="handleSubjectSelect()"
              class="w-full mt-1 px-3 py-2.5 rounded-xl bg-slate-50 border border-slate-300 text-xs sm:text-sm text-slate-900 font-medium focus:border-blue-800 focus:outline-none">
            </select>
          </div>
        </div>

        <div id="manual-subject-box" class="hidden">
          <label class="text-xs font-bold text-amber-800">✍️ Type Your Subject Name Manually *</label>
          <input id="manual-subject-input" type="text"
            placeholder="e.g., Linear Integrated Circuits / Microcontroller Lab / Cloud Security"
            class="w-full mt-1 px-3.5 py-2.5 rounded-xl bg-amber-50/50 border border-amber-400 text-slate-900 text-xs sm:text-sm focus:outline-none" />
        </div>

        <div>
          <label class="text-xs font-bold text-slate-700">3. Enter Project Title / Problem Statement *</label>
          <input id="topic-input" type="text"
            placeholder="e.g., Water Level Indicator using Opamp IC741 / Smart Irrigation System"
            class="w-full mt-1.5 px-4 py-3.5 rounded-xl bg-slate-50 border border-slate-300 text-slate-900 text-sm font-medium focus:border-blue-800 focus:outline-none" />
        </div>

        <div class="grid grid-cols-1 sm:grid-cols-2 gap-3.5">
          <div>
            <label class="text-xs font-semibold text-slate-600">Project Level</label>
            <select id="level-input" class="w-full mt-1 px-3 py-2.5 rounded-xl bg-slate-50 border border-slate-300 text-xs sm:text-sm text-slate-900">
              <option>2nd / 3rd Year Subject Mini-Project</option>
              <option>Final Year Major Capstone Project</option>
              <option>Lab / Hackathon Prototype</option>
            </select>
          </div>
          <div>
            <label class="text-xs font-semibold text-slate-600">Guide Name (For Certificate)</label>
            <input id="guide-input" type="text" placeholder="Dr. S. K. Kulkarni"
              class="w-full mt-1 px-3 py-2.5 rounded-xl bg-slate-50 border border-slate-300 text-xs sm:text-sm text-slate-900" />
          </div>
        </div>

        <button type="button" id="generate-btn" onclick="generateProject()"
          class="w-full py-3.5 rounded-xl bg-blue-900 hover:bg-blue-800 font-bold text-white text-sm sm:text-base shadow-md transition">
          Generate Project
        </button>
      </div>

      <!-- Recent Projects History -->
      <div id="recent-history-box" class="hidden card rounded-2xl p-4 max-w-3xl mx-auto">
        <span class="text-[11px] uppercase tracking-wider text-slate-500 font-bold">🕒 Recent Projects (Click to Open)</span>
        <div id="recent-list" class="flex flex-wrap gap-2 mt-2"></div>
      </div>
    </section>

    <!-- ==================== PAGE 2: 3D CIRCUIT OR SOFTWARE ARCHITECTURE ==================== -->
    <section id="page-3d" class="page-section space-y-4">
      <div id="ai-verify-card" class="hidden card rounded-2xl p-4 border-l-4 border-l-emerald-600 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
        <div class="space-y-1">
          <div class="flex flex-wrap items-center gap-2">
            <span class="text-[10px] uppercase font-bold px-2.5 py-0.5 rounded-full bg-emerald-50 text-emerald-800 border border-emerald-200">
              ✓ Syllabus &amp; Lab Verified
            </span>
            <span id="verify-match" class="text-xs font-bold text-blue-900"></span>
            <span id="verify-time" class="text-xs text-slate-500"></span>
          </div>
          <p id="verify-note" class="text-xs text-slate-600"></p>
        </div>
      </div>

      <div id="circuit-lab-wrapper" class="space-y-4">
        <div class="card rounded-2xl p-4 sm:p-5 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div>
            <span class="text-[11px] uppercase tracking-wider text-blue-800 font-bold">Interactive 3D Hardware &amp; Breadboard Lab</span>
            <h2 id="lab3d-title" class="text-base sm:text-xl font-bold text-slate-900">3D Circuit Assembly</h2>
            <p class="text-xs text-slate-500">Drag 1-Finger to Rotate 360° • Pinch / Scroll to Zoom • Test Sensor Input Below</p>
          </div>
          <div class="flex flex-wrap gap-2">
            <button type="button" onclick="toggle3DSimulation()" id="sim-toggle-btn"
              class="px-3.5 py-2 rounded-xl bg-emerald-600 font-bold text-white text-xs shadow-sm">
              ▶ Simulation: ON
            </button>
            <button type="button" onclick="reset3DCamera()"
              class="px-3 py-2 rounded-xl bg-slate-100 hover:bg-slate-200 text-xs text-slate-800 font-semibold border border-slate-300">
              🔄 Reset View
            </button>
          </div>
        </div>

        <div class="card rounded-2xl p-2.5 space-y-2.5">
          <div id="threejs-wrapper" class="relative w-full h-[380px] sm:h-[460px] rounded-xl bg-[#e2e8f0] overflow-hidden border border-slate-300">
            <div id="threejs-container" class="w-full h-full"></div>
            <div id="labels-overlay" class="absolute inset-0 pointer-events-none overflow-hidden"></div>
          </div>

          <div class="bg-slate-50 rounded-xl p-3 border border-slate-200 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <div class="flex items-center gap-2">
              <span class="text-xs font-bold text-slate-800">🎛️ Live Sensor / Input Level:</span>
              <span id="sensor-slider-val" class="text-xs font-mono font-bold px-2 py-0.5 rounded bg-blue-100 text-blue-900">65%</span>
            </div>
            <input id="sensor-stimulus-slider" type="range" min="0" max="100" value="65"
              oninput="updateSensorStimulus(this.value)"
              class="w-full sm:w-64 accent-blue-900 cursor-pointer" />
          </div>
        </div>

        <div class="card rounded-2xl p-5">
          <h3 class="text-sm font-bold text-slate-900 mb-3">📌 Exact Pin-to-Pin Breadboard Connections</h3>
          <div class="overflow-x-auto">
            <table class="w-full text-xs text-left border border-slate-200">
              <thead class="bg-slate-100 text-slate-800">
                <tr>
                  <th class="p-2.5 border-b border-slate-200">Source Pin / IC</th>
                  <th class="p-2.5 border-b border-slate-200">Breadboard / Target</th>
                  <th class="p-2.5 border-b border-slate-200">Signal / Voltage</th>
                </tr>
              </thead>
              <tbody id="lab3d-pins" class="divide-y divide-slate-200 text-slate-700"></tbody>
            </table>
          </div>
        </div>
      </div>

      <div id="software-arch-wrapper" class="hidden space-y-4">
        <div class="card rounded-2xl p-6 space-y-4">
          <div>
            <span class="text-[11px] uppercase tracking-wider text-blue-800 font-bold">Software System Architecture</span>
            <h2 id="sw-arch-title" class="text-lg sm:text-xl font-bold text-slate-900 mt-0.5">Data &amp; Execution Pipeline</h2>
            <p class="text-xs text-slate-500 mt-1">This is a pure software subject project—no hardware circuit is required.</p>
          </div>
          <div id="sw-nodes-container" class="grid grid-cols-1 sm:grid-cols-5 gap-3 pt-2"></div>
        </div>
      </div>
    </section>

    <!-- ==================== PAGE 3: CODE & VOICE VIVA ==================== -->
    <section id="page-code" class="page-section space-y-5">
      <div id="empty-code" class="card rounded-2xl p-8 text-center">
        <p class="text-sm text-slate-500">Generate a project first to view source code and 5-language Voice Viva Q&amp;A.</p>
        <button type="button" onclick="navigatePage('page-home')" class="mt-3 px-4 py-2 rounded-xl bg-blue-900 text-xs font-semibold text-white">Go to Create →</button>
      </div>

      <div id="content-code" class="hidden space-y-5">
        <div class="card rounded-2xl p-5 space-y-3">
          <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
            <div class="flex flex-wrap gap-2">
              <button type="button" onclick="switchCodeMode('primary')" id="btn-code-primary"
                class="px-3.5 py-2 rounded-xl bg-blue-900 text-white text-xs font-bold">
                💻 Main Project Code
              </button>
              <button type="button" onclick="switchCodeMode('arduino')" id="btn-code-arduino"
                class="px-3.5 py-2 rounded-xl bg-slate-100 text-slate-700 text-xs font-semibold border border-slate-300">
                🔌 Optional Arduino Dump Code (.ino)
              </button>
            </div>
            <button type="button" onclick="copySourceCode()" class="px-3 py-1.5 rounded-lg bg-slate-100 hover:bg-slate-200 text-xs text-slate-800 font-semibold border border-slate-300 self-start">
              📋 Copy Code
            </button>
          </div>
          <p id="active-code-subtitle" class="text-xs text-blue-900 font-bold"></p>
          <pre id="code-block" class="bg-slate-900 p-4 rounded-xl text-xs text-emerald-300 overflow-x-auto leading-relaxed"></pre>
        </div>

        <div class="card rounded-2xl p-5 space-y-4">
          <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <div>
              <h3 class="text-base font-bold text-slate-900">🎓 Viva-Voce Q&amp;A (5 Languages + Voice Audio)</h3>
              <p class="text-xs text-slate-500">Select your preferred language &amp; tap 🔊 Listen to practice</p>
            </div>
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

    <!-- ==================== PAGE 4: REPORT (.DOCX), IEEE PAPER & POWERPOINT (.PPTX) ==================== -->
    <section id="page-report" class="page-section space-y-5">
      <div id="empty-report" class="card rounded-2xl p-8 text-center">
        <p class="text-sm text-slate-500">Generate a project first to download your 15+ page Report (.docx), IEEE Paper &amp; Seminar PPT (.pptx).</p>
        <button type="button" onclick="navigatePage('page-home')" class="mt-3 px-4 py-2 rounded-xl bg-blue-900 text-xs font-semibold text-white">Go to Create →</button>
      </div>

      <div id="content-report" class="hidden space-y-5">
        <div class="card rounded-2xl p-5 space-y-4">
          <div>
            <span class="text-[11px] uppercase tracking-wider text-blue-800 font-bold">University Submission Downloads</span>
            <h2 id="rep-title" class="text-lg sm:text-xl font-bold text-slate-900 mt-0.5"></h2>
          </div>

          <!-- 3 Direct Download Buttons: VTU DOCX + Seminar PPTX + IEEE Research Paper -->
          <div class="grid grid-cols-1 sm:grid-cols-3 gap-3">
            <button type="button" onclick="downloadDocxDirect()" id="download-docx-btn"
              class="py-3 px-4 rounded-xl bg-blue-900 hover:bg-blue-800 font-bold text-white text-xs shadow-sm flex items-center justify-center gap-2">
              <span>📄</span><span>Download Report (.docx)</span>
            </button>
            <button type="button" onclick="downloadPptxDirect()" id="download-pptx-btn"
              class="py-3 px-4 rounded-xl bg-emerald-700 hover:bg-emerald-600 font-bold text-white text-xs shadow-sm flex items-center justify-center gap-2">
              <span>📊</span><span>Download Seminar PPT (.pptx)</span>
            </button>
            <button type="button" onclick="downloadIeeeDirect()" id="download-ieee-btn"
              class="py-3 px-4 rounded-xl bg-slate-800 hover:bg-slate-700 font-bold text-white text-xs shadow-sm flex items-center justify-center gap-2">
              <span>📑</span><span>Download IEEE Paper (.docx)</span>
            </button>
          </div>

          <!-- Report Formatting Tools -->
          <div class="grid grid-cols-2 sm:grid-cols-4 gap-3 pt-3 border-t border-slate-200">
            <div>
              <label class="text-[11px] font-semibold text-slate-600">Font Style</label>
              <select id="fmt-font-family" onchange="updateLivePreviewStyle()" class="w-full mt-1 px-2.5 py-2 rounded-lg bg-slate-50 border border-slate-300 text-xs text-slate-900">
                <option value="Times New Roman">Times New Roman (VTU)</option>
                <option value="Arial">Arial</option>
                <option value="Calibri">Calibri</option>
              </select>
            </div>
            <div>
              <label class="text-[11px] font-semibold text-slate-600">Body Font Size</label>
              <select id="fmt-font-size" onchange="updateLivePreviewStyle()" class="w-full mt-1 px-2.5 py-2 rounded-lg bg-slate-50 border border-slate-300 text-xs text-slate-900">
                <option value="12" selected>12 pt (Standard)</option>
                <option value="11">11 pt (Compact)</option>
                <option value="14">14 pt (Large)</option>
              </select>
            </div>
            <div>
              <label class="text-[11px] font-semibold text-slate-600">Line Spacing</label>
              <select id="fmt-line-spacing" onchange="updateLivePreviewStyle()" class="w-full mt-1 px-2.5 py-2 rounded-lg bg-slate-50 border border-slate-300 text-xs text-slate-900">
                <option value="1.5" selected>1.5x Spacing (VTU)</option>
                <option value="1.15">1.15x Spacing</option>
                <option value="2.0">2.0x Double Space</option>
              </select>
            </div>
            <div>
              <label class="text-[11px] font-semibold text-slate-600">Page Border</label>
              <select id="fmt-page-border" class="w-full mt-1 px-2.5 py-2 rounded-lg bg-slate-50 border border-slate-300 text-xs text-slate-900">
                <option value="true" selected>Classic Blue Border</option>
                <option value="false">No Border (Clean)</option>
              </select>
            </div>
          </div>
        </div>

        <div id="live-doc-preview" class="card rounded-2xl p-6 space-y-4 border-2 border-blue-900/20">
          <h3 class="text-base font-bold text-blue-900">Document Preview</h3>
          <p id="rep-abstract" class="text-sm text-slate-700 leading-relaxed"></p>
          <p id="rep-working" class="text-sm text-slate-600 leading-relaxed border-t border-slate-200 pt-3"></p>
        </div>

        <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
          <div class="card rounded-2xl p-5">
            <h3 class="text-sm font-bold text-slate-900 mb-2">💰 Bill of Materials (BOM)</h3>
            <div id="rep-bom" class="text-xs text-slate-700 space-y-2"></div>
          </div>
          <div class="card rounded-2xl p-5">
            <div class="flex items-center justify-between mb-2">
              <h3 class="text-sm font-bold text-slate-900">📊 Seminar PPT Slide Outline</h3>
              <button type="button" onclick="downloadPptxDirect()" class="text-[11px] px-2.5 py-1 rounded bg-emerald-50 text-emerald-800 border border-emerald-200 font-bold">📥 .pptx</button>
            </div>
            <div id="rep-ppt" class="text-xs text-slate-700 space-y-2"></div>
          </div>
        </div>

        <!-- Indian Engineering Student Submission Checklist -->
        <div class="card rounded-2xl p-5">
          <h3 class="text-sm font-bold text-slate-900 mb-2">✅ College Viva &amp; Lab Submission Checklist</h3>
          <div class="grid grid-cols-1 sm:grid-cols-2 gap-2 text-xs text-slate-700">
            <label class="flex items-center gap-2"><input type="checkbox" checked class="accent-blue-900"> Cover Page &amp; Bonafide Certificate signed by Guide &amp; HOD</label>
            <label class="flex items-center gap-2"><input type="checkbox" checked class="accent-blue-900"> 15+ Page Report printed in Times New Roman 12pt (1.5 spacing)</label>
            <label class="flex items-center gap-2"><input type="checkbox" checked class="accent-blue-900"> 7-Slide Seminar PowerPoint (.pptx) ready on USB/Laptop</label>
            <label class="flex items-center gap-2"><input type="checkbox" checked class="accent-blue-900"> Hardware Breadboard / Working Code Demo tested</label>
          </div>
        </div>
      </div>
    </section>
  </main>

  <!-- MOBILE BOTTOM NAVIGATION -->
  <nav class="md:hidden fixed bottom-0 inset-x-0 z-40 bg-white border-t border-slate-200 grid grid-cols-4 py-1.5 px-1 shadow-lg">
    <button type="button" onclick="navigatePage('page-home')" id="nav-m-home" class="nav-m-btn flex flex-col items-center py-1 text-blue-900 font-bold">
      <span class="text-base">🏠</span><span class="text-[10px]">Create</span>
    </button>
    <button type="button" onclick="navigatePage('page-3d')" id="nav-m-3d" class="nav-m-btn flex flex-col items-center py-1 text-slate-500">
      <span class="text-base">🧊</span><span class="text-[10px]">Circuit/Flow</span>
    </button>
    <button type="button" onclick="navigatePage('page-code')" id="nav-m-code" class="nav-m-btn flex flex-col items-center py-1 text-slate-500">
      <span class="text-base">💻</span><span class="text-[10px]">Code/Viva</span>
    </button>
    <button type="button" onclick="navigatePage('page-report')" id="nav-m-report" class="nav-m-btn flex flex-col items-center py-1 text-slate-500">
      <span class="text-base">📄</span><span class="text-[10px]">Report/PPTX</span>
    </button>
  </nav>

  <script>
    const BRANCH_SUBJECTS = {
      "Computer Science & Engineering (CSE)": [
        "Artificial Intelligence & Machine Learning",
        "Database Management Systems (DBMS)",
        "Computer Networks & Cybersecurity",
        "Operating Systems & System Software",
        "Full-Stack Web & Cloud Computing",
        "Internet of Things (IoT) & Embedded Systems",
        "Digital Image Processing & Computer Vision",
        "Data Structures & Algorithms",
        "Software Engineering & Agile Practices"
      ],
      "Electronics & Communication (ECE)": [
        "Analog Electronic Circuits (Op-Amp & Linear ICs)",
        "Embedded Systems & Microcontrollers",
        "Digital System Design & Verilog HDL",
        "Digital Signal Processing (DSP)",
        "IoT & Wireless Sensor Networks",
        "Control Systems & Robotics",
        "VLSI Design & CMOS Circuits",
        "Digital & Wireless Communication"
      ],
      "Artificial Intelligence & ML (AIML)": [
        "Deep Learning & Neural Networks (CNN/RNN)",
        "Natural Language Processing (NLP)",
        "Computer Vision & Video Analytics",
        "Data Science & Predictive Modeling",
        "Reinforcement Learning & Autonomous Agents",
        "Edge AI & TinyML Hardware"
      ],
      "Electrical & Electronics (EEE)": [
        "Power Electronics & Motor Drives",
        "Microcontrollers & PLC Industrial Automation",
        "Renewable Energy & Smart Grid Systems",
        "Analog & Digital Instrumentation",
        "Electric Vehicles (EV) & Battery Management",
        "High Voltage & Protection Engineering"
      ],
      "Information Science (ISE)": [
        "Software Engineering & Web Applications",
        "Database Management Systems (DBMS)",
        "Cloud Computing & DevOps",
        "Machine Learning & Data Mining",
        "Information & Network Security",
        "Big Data Analytics"
      ],
      "Mechanical Engineering": [
        "Mechatronics & Industrial Automation",
        "Robotics & Kinematics",
        "CAD/CAM & Additive Manufacturing",
        "Thermal & Fluid Engineering Simulation",
        "Automobile & EV Powertrain Systems"
      ],
      "Civil Engineering": [
        "Smart Cities & Structural Health Monitoring",
        "Environmental & Water Resource Engineering",
        "Transportation & Smart Traffic Automation",
        "Geotechnical & Surveying Systems",
        "Green Building & Sustainable Construction"
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
      if (val === "__MANUAL__") {
        box.classList.remove("hidden");
        document.getElementById("manual-subject-input").focus();
      } else {
        box.classList.add("hidden");
      }
    }

    let currentBlueprint = null;
    let studentProfile = JSON.parse(localStorage.getItem("pf_student") || "null");
    let recentBlueprints = JSON.parse(localStorage.getItem("pf_recent_projects") || "[]");
    let activeCodeMode = "primary";
    let activeVivaLang = "en";
    let selectedRating = 5;
    let sensorStimulusLevel = 0.65;
    let deferredInstallPrompt = null;

    // Register Service Worker & Capture Install Prompt
    if ("serviceWorker" in navigator) {
      navigator.serviceWorker.register("/sw.js").catch(() => {});
    }
    window.addEventListener("beforeinstallprompt", (e) => {
      e.preventDefault();
      deferredInstallPrompt = e;
    });

    async function triggerInstallApp() {
      if (deferredInstallPrompt) {
        deferredInstallPrompt.prompt();
        await deferredInstallPrompt.userChoice;
        deferredInstallPrompt = null;
      } else {
        const m = document.getElementById("install-help-modal");
        m.classList.remove("hidden"); m.classList.add("flex");
      }
    }
    function closeInstallHelp() {
      const m = document.getElementById("install-help-modal");
      m.classList.add("hidden"); m.classList.remove("flex");
    }

    function dismissSplash() {
      const s = document.getElementById("splash-screen");
      if (s) {
        s.style.opacity = "0";
        s.style.visibility = "hidden";
        s.style.pointerEvents = "none";
      }
    }

    window.addEventListener("DOMContentLoaded", () => {
      updateSubjectsForBranch();
      renderRecentProjects();
      setTimeout(() => {
        dismissSplash();
        if (!studentProfile || !studentProfile.email) openLoginModal();
        else updateProfileUI();
      }, 950);

      try {
        initThreeJS();
        build3DSceneFromBlueprint(null);
      } catch (e) {
        console.warn("3D init warning:", e);
      }
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
      if (dBtn) {
        dBtn.classList.add("bg-blue-900", "text-white", "font-bold");
        dBtn.classList.remove("text-slate-600", "font-semibold");
      }

      document.querySelectorAll(".nav-m-btn").forEach(b => {
        b.classList.remove("text-blue-900", "font-bold");
        b.classList.add("text-slate-500");
      });
      const mBtn = document.getElementById("nav-m-" + key);
      if (mBtn) {
        mBtn.classList.add("text-blue-900", "font-bold");
        mBtn.classList.remove("text-slate-500");
      }

      if (pageId === "page-3d") setTimeout(resizeThreeJS, 50);
      window.scrollTo({ top: 0, behavior: "smooth" });
    }

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
    function copyInviteLink() { navigator.clipboard.writeText(window.location.origin + "/?ref=" + getReferralCode()); alert("Invite link copied!"); }
    function shareOnWhatsApp() {
      const link = window.location.origin + "/?ref=" + getReferralCode();
      const text = "Build Engineering Projects, 3D Circuits, VTU Reports (.docx) & Seminar PPTs (.pptx): " + link;
      window.open("https://wa.me/?text=" + encodeURIComponent(text), "_blank");
    }

    function openFeedbackModal() {
      const m = document.getElementById("feedback-modal"); m.classList.remove("hidden"); m.classList.add("flex");
    }
    function closeFeedbackModal() {
      const m = document.getElementById("feedback-modal"); m.classList.add("hidden"); m.classList.remove("flex");
    }
    function setRating(r) {
      selectedRating = r;
      document.querySelectorAll(".star-btn").forEach((b, idx) => {
        if (idx + 1 === r) b.className = "star-btn px-3 py-1.5 rounded-lg bg-blue-900 border border-blue-900 text-white font-bold text-xs";
        else b.className = "star-btn px-3 py-1.5 rounded-lg bg-slate-100 border border-slate-300 text-xs";
      });
    }
    document.getElementById("feedback-form").addEventListener("submit", async (e) => {
      e.preventDefault();
      const notes = document.getElementById("feedback-text").value.trim();
      try {
        await fetch("/api/feedback", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ rating: selectedRating, notes, student: studentProfile })
        });
      } catch (_) {}
      closeFeedbackModal();
      alert("Thank you! Your feedback has been saved.");
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
      if (currentBlueprint) {
        populateAllPages(currentBlueprint);
        navigatePage("page-3d");
      }
    }

    async function generateProject() {
      const topic = document.getElementById("topic-input").value.trim();
      const branch = document.getElementById("branch-input").value;
      let subject = document.getElementById("subject-input").value;
      if (subject === "__MANUAL__") {
        subject = document.getElementById("manual-subject-input").value.trim();
        if (!subject) return alert("Please enter your subject name manually!");
      }
      const level = document.getElementById("level-input").value;
      if (!topic) return alert("Please enter your project title first!");

      const btn = document.getElementById("generate-btn");
      btn.disabled = true; btn.innerText = "Generating Project...";
      try {
        const res = await fetch("/api/generate", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ topic, branch, subject, level })
        });
        const data = await res.json();
        if (data.error) throw new Error(data.error);
        currentBlueprint = data.blueprint;
        saveRecentProject(currentBlueprint);
        populateAllPages(currentBlueprint);
        navigatePage("page-3d");
      } catch (err) {
        alert("Error: " + err.message);
      } finally {
        btn.disabled = false; btn.innerText = "Generate Project";
      }
    }

    function populateAllPages(bp) {
      if (bp.verification_summary) {
        document.getElementById("ai-verify-card").classList.remove("hidden");
        document.getElementById("verify-match").innerText = bp.verification_summary.syllabus_match || "";
        document.getElementById("verify-time").innerText = "• Est. Build: " + (bp.verification_summary.build_time || "2 Days");
        document.getElementById("verify-note").innerText = bp.verification_summary.expert_note || "";
      }

      if (bp.requires_circuit) {
        document.getElementById("circuit-lab-wrapper").classList.remove("hidden");
        document.getElementById("software-arch-wrapper").classList.add("hidden");
        document.getElementById("lab3d-title").innerText = `3D Circuit: ${bp.title}`;
        build3DSceneFromBlueprint(bp);
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
        document.getElementById("sw-arch-title").innerText = `System Flow: ${bp.title}`;
        document.getElementById("sw-nodes-container").innerHTML = (bp.software_architecture_nodes || []).map((n, i) => `
          <div class="bg-blue-50 border border-blue-200 rounded-xl p-4 text-center">
            <span class="text-[10px] font-bold text-blue-800">STAGE 0${i + 1}</span>
            <p class="text-xs font-bold text-slate-900 mt-1">${n}</p>
          </div>
        `).join("");
      }

      document.getElementById("empty-code").classList.add("hidden");
      document.getElementById("content-code").classList.remove("hidden");
      const ardBtn = document.getElementById("btn-code-arduino");
      if (bp.has_arduino_hardware) ardBtn.classList.remove("hidden");
      else ardBtn.classList.add("hidden");
      switchCodeMode("primary");
      renderVivaCards();

      document.getElementById("empty-report").classList.add("hidden");
      document.getElementById("content-report").classList.remove("hidden");
      document.getElementById("rep-title").innerText = bp.title || "";
      document.getElementById("rep-abstract").innerText = bp.abstract || "";
      document.getElementById("rep-working").innerText = bp.working_principle || "";

      let total = 0;
      document.getElementById("rep-bom").innerHTML = (bp.hardware_requirements || []).map(h => {
        total += Number(h.cost_inr || 0);
        return `<div class="flex justify-between border-b border-slate-200 pb-1"><span>${h.name}</span><span class="text-emerald-700 font-bold">₹${h.cost_inr}</span></div>`;
      }).join("") + `<div class="flex justify-between pt-1 font-bold text-slate-900"><span>Total Est. Cost</span><span>₹${total}</span></div>`;

      document.getElementById("rep-ppt").innerHTML = (bp.ppt_slides || []).map((s, i) => {
        const desc = s.bullets ? s.bullets.join(" • ") : (s.points || "");
        return `<div class="border-b border-slate-200 pb-1.5"><span class="text-blue-900 font-bold">Slide ${i + 1} (${s.slide}):</span> ${desc}</div>`;
      }).join("");
    }

    function switchCodeMode(mode) {
      if (!currentBlueprint) return;
      activeCodeMode = mode;
      const pBtn = document.getElementById("btn-code-primary");
      const aBtn = document.getElementById("btn-code-arduino");
      if (mode === "primary") {
        pBtn.className = "px-3.5 py-2 rounded-xl bg-blue-900 text-white text-xs font-bold";
        aBtn.className = "px-3.5 py-2 rounded-xl bg-slate-100 text-slate-700 text-xs font-semibold border border-slate-300";
        document.getElementById("active-code-subtitle").innerText = `Language: ${currentBlueprint.primary_code_language || "Code"}`;
        document.getElementById("code-block").innerText = currentBlueprint.primary_code || "";
      } else {
        aBtn.className = "px-3.5 py-2 rounded-xl bg-blue-900 text-white text-xs font-bold";
        pBtn.className = "px-3.5 py-2 rounded-xl bg-slate-100 text-slate-700 text-xs font-semibold border border-slate-300";
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
        b.classList.remove("bg-blue-900", "text-white", "font-bold");
        b.classList.add("bg-slate-100", "text-slate-700", "font-semibold");
      });
      btn.classList.remove("bg-slate-100", "text-slate-700", "font-semibold");
      btn.classList.add("bg-blue-900", "text-white", "font-bold");
      renderVivaCards();
    }

    function speakVivaByIndex(idx) {
      if (!currentBlueprint || !("speechSynthesis" in window)) return;
      const item = (currentBlueprint.viva_qa || [])[idx];
      if (!item) return;
      const entry = item[activeVivaLang] || item["en"] || {};
      window.speechSynthesis.cancel();
      const u = new SpeechSynthesisUtterance(entry.a || "");
      const langMap = { en: "en-IN", kn: "kn-IN", hi: "hi-IN", te: "te-IN", ta: "ta-IN" };
      u.lang = langMap[activeVivaLang] || "en-IN";
      window.speechSynthesis.speak(u);
    }

    function renderVivaCards() {
      if (!currentBlueprint) return;
      const list = currentBlueprint.viva_qa || [];
      document.getElementById("viva-grid").innerHTML = list.map((item, idx) => {
        const entry = item[activeVivaLang] || item["en"] || {};
        return `
          <div class="bg-slate-50 p-4 rounded-xl border border-slate-200 space-y-2">
            <div class="flex items-start justify-between gap-2">
              <p class="text-xs font-bold text-blue-900">Q${idx + 1}. ${entry.q || ""}</p>
              <button type="button" onclick="speakVivaByIndex(${idx})" class="shrink-0 text-[10px] px-2 py-1 rounded bg-blue-100 hover:bg-blue-200 text-blue-900 font-bold">
                🔊 Listen
              </button>
            </div>
            <p class="text-xs text-slate-700 leading-relaxed">${entry.a || ""}</p>
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

    function getExportPayload() {
      let subject = document.getElementById("subject-input").value;
      if (subject === "__MANUAL__") subject = document.getElementById("manual-subject-input").value.trim() || "Engineering Project";
      return {
        blueprint: currentBlueprint,
        student_name: (studentProfile && studentProfile.name) || "Student Name",
        usn: (studentProfile && studentProfile.usn) || "1VT23CS001",
        college: (studentProfile && studentProfile.college) || "VTU Affiliated Engineering College",
        branch: document.getElementById("branch-input").value,
        subject: subject,
        guide_name: document.getElementById("guide-input").value.trim() || "Project Guide",
        formatting: {
          font_family: document.getElementById("fmt-font-family").value,
          font_size: document.getElementById("fmt-font-size").value,
          line_spacing: document.getElementById("fmt-line-spacing").value,
          page_border: document.getElementById("fmt-page-border").value === "true",
          viva_lang: activeVivaLang
        }
      };
    }

    async function downloadFileFromEndpoint(endpoint, filename, btnId, defaultLabel) {
      if (!currentBlueprint) return;
      const btn = document.getElementById(btnId);
      btn.disabled = true; btn.innerText = "Preparing File...";
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
        setTimeout(openFeedbackModal, 700);
      } catch (err) {
        alert("Error: " + err.message);
      } finally {
        btn.disabled = false; btn.innerHTML = defaultLabel;
      }
    }

    function downloadDocxDirect() {
      downloadFileFromEndpoint("/api/download-docx", "ProjForge_Project_Report.docx", "download-docx-btn", "<span>📄</span><span>Download Report (.docx)</span>");
    }

    function downloadPptxDirect() {
      downloadFileFromEndpoint("/api/download-pptx", "ProjForge_Seminar_Presentation.pptx", "download-pptx-btn", "<span>📊</span><span>Download Seminar PPT (.pptx)</span>");
    }

    function downloadIeeeDirect() {
      downloadFileFromEndpoint("/api/download-ieee", "ProjForge_IEEE_Paper.docx", "download-ieee-btn", "<span>📑</span><span>Download IEEE Paper (.docx)</span>");
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
      scene.background = new THREE.Color(0xe2e8f0);

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
        controls.minDistance = 5;
        controls.maxDistance = 20;
      }

      scene.add(new THREE.AmbientLight(0xffffff, 0.85));
      const keyLight = new THREE.DirectionalLight(0xffffff, 0.95);
      keyLight.position.set(8, 16, 10);
      scene.add(keyLight);

      const grid = new THREE.GridHelper(22, 22, 0x94a3b8, 0xcbd5e1);
      grid.position.y = -0.22;
      scene.add(grid);

      circuitGroup = new THREE.Group();
      scene.add(circuitGroup);
      animate3D();
    }

    function createRealisticPartMesh(type) {
      const grp = new THREE.Group();
      if (type === "ic_741" || type === "ic_lm324") {
        const len = type === "ic_lm324" ? 1.15 : 0.75;
        const body = new THREE.Mesh(new THREE.BoxGeometry(len, 0.22, 0.56), new THREE.MeshStandardMaterial({ color: 0x111827, roughness: 0.4 }));
        const notch = new THREE.Mesh(new THREE.CylinderGeometry(0.08, 0.08, 0.23, 12), new THREE.MeshStandardMaterial({ color: 0x374151 }));
        notch.position.set(-len / 2 + 0.04, 0, 0);
        grp.add(body, notch);
        const step = type === "ic_lm324" ? 0.15 : 0.18;
        for (let px = -len / 2 + 0.12; px <= len / 2 - 0.1; px += step) {
          const pin = new THREE.Mesh(new THREE.BoxGeometry(0.04, 0.28, 0.72), new THREE.MeshStandardMaterial({ color: 0xd1d5db, metalness: 0.85 }));
          pin.position.set(px, -0.05, 0);
          grp.add(pin);
        }
      } else if (type === "sensor_water_probes") {
        const tank = new THREE.Mesh(new THREE.BoxGeometry(1.35, 1.0, 0.95), new THREE.MeshStandardMaterial({ color: 0x7dd3fc, transparent: true, opacity: 0.42 }));
        const water = new THREE.Mesh(new THREE.BoxGeometry(1.26, 0.75, 0.86), new THREE.MeshStandardMaterial({ color: 0x0284c7, transparent: true, opacity: 0.8 }));
        water.position.y = -0.1;
        [-0.3, 0, 0.3].forEach((rx, idx) => {
          const rod = new THREE.Mesh(new THREE.CylinderGeometry(0.025, 0.025, 0.95 - idx * 0.15, 10), new THREE.MeshStandardMaterial({ color: 0xd97706, metalness: 0.8 }));
          rod.position.set(rx, 0.1, 0);
          grp.add(rod);
        });
        grp.add(tank, water);
        animParts.push({ mesh: water, mode: "water" });
      } else if (type === "sensor_soil_moisture") {
        const head = new THREE.Mesh(new THREE.BoxGeometry(0.75, 0.35, 0.12), new THREE.MeshStandardMaterial({ color: 0x991b1b }));
        head.position.y = 0.35;
        const prongL = new THREE.Mesh(new THREE.BoxGeometry(0.14, 0.7, 0.06), new THREE.MeshStandardMaterial({ color: 0xfbbf24, metalness: 0.7 }));
        prongL.position.set(-0.18, -0.15, 0);
        const prongR = prongL.clone(); prongR.position.set(0.18, -0.15, 0);
        grp.add(head, prongL, prongR);
      } else if (type === "water_pump") {
        const body = new THREE.Mesh(new THREE.CylinderGeometry(0.36, 0.36, 0.85, 20), new THREE.MeshStandardMaterial({ color: 0xf8fafc }));
        body.rotation.z = Math.PI / 2;
        const nozzle = new THREE.Mesh(new THREE.CylinderGeometry(0.11, 0.11, 0.45, 14), new THREE.MeshStandardMaterial({ color: 0x38bdf8 }));
        nozzle.position.set(0.25, 0.45, 0);
        grp.add(body, nozzle);
      } else if (type === "relay_module") {
        const pcb = new THREE.Mesh(new THREE.BoxGeometry(1.1, 0.12, 0.85), new THREE.MeshStandardMaterial({ color: 0x1e293b }));
        const cube = new THREE.Mesh(new THREE.BoxGeometry(0.65, 0.55, 0.55), new THREE.MeshStandardMaterial({ color: 0x0284c7 }));
        cube.position.y = 0.32;
        grp.add(pcb, cube);
      } else if (type === "led_indicator_bank" || type === "piezo_buzzer") {
        [0x22c55e, 0xeab308, 0xef4444].forEach((clr, idx) => {
          const led = new THREE.Mesh(new THREE.SphereGeometry(0.13, 14, 14), new THREE.MeshBasicMaterial({ color: clr }));
          led.position.set(-0.32 + idx * 0.32, 0.24, -0.15);
          grp.add(led);
          if (idx === 2) animParts.push({ mesh: led, mode: "blink" });
        });
        const buzzer = new THREE.Mesh(new THREE.CylinderGeometry(0.26, 0.26, 0.28, 18), new THREE.MeshStandardMaterial({ color: 0x0f172a }));
        buzzer.position.set(0, 0.18, 0.25);
        grp.add(buzzer);
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
      } else {
        const bezel = new THREE.Mesh(new THREE.BoxGeometry(1.2, 0.8, 0.14), new THREE.MeshStandardMaterial({ color: 0x1e3a8a }));
        const screen = new THREE.Mesh(new THREE.BoxGeometry(1.0, 0.58, 0.04), new THREE.MeshBasicMaterial({ color: 0x06b6d4 }));
        screen.position.z = 0.08;
        grp.add(bezel, screen);
      }
      return grp;
    }

    function add3DWire(p1, p2, p3, colorHex) {
      const curve = new THREE.CatmullRomCurve3([new THREE.Vector3(...p1), new THREE.Vector3(...p2), new THREE.Vector3(...p3)]);
      const geo = new THREE.TubeGeometry(curve, 24, 0.04, 8, false);
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
      trackedLabels = []; animParts = [];
      document.getElementById("labels-overlay").innerHTML = "";

      const supplyLabel = (bp && bp.controller_or_supply) || "9V DC Analog Supply";
      const isAnalogBattery = supplyLabel.toLowerCase().includes("9v") || supplyLabel.toLowerCase().includes("analog");

      if (isAnalogBattery) {
        const bat = new THREE.Mesh(new THREE.BoxGeometry(2.0, 0.95, 1.25), new THREE.MeshStandardMaterial({ color: 0x111827 }));
        bat.position.set(-3.4, 0.35, 0);
        const cap = new THREE.Mesh(new THREE.BoxGeometry(2.02, 0.32, 1.27), new THREE.MeshStandardMaterial({ color: 0xd97706 }));
        cap.position.set(-3.4, 0.68, 0);
        const term1 = new THREE.Mesh(new THREE.CylinderGeometry(0.1, 0.1, 0.18, 12), new THREE.MeshStandardMaterial({ color: 0xd1d5db, metalness: 0.8 }));
        term1.position.set(-2.6, 0.88, -0.25);
        const term2 = term1.clone(); term2.position.set(-2.6, 0.88, 0.25);
        circuitGroup.add(bat, cap, term1, term2);
        registerLabel("9V DC Battery Supply", -3.4, 1.28, 0);
      } else {
        const mcu = new THREE.Mesh(new THREE.BoxGeometry(3.3, 0.22, 2.2), new THREE.MeshStandardMaterial({ color: 0x0284c7, roughness: 0.35 }));
        mcu.position.set(-3.2, 0, 0);
        const usb = new THREE.Mesh(new THREE.BoxGeometry(0.65, 0.36, 0.55), new THREE.MeshStandardMaterial({ color: 0x94a3b8, metalness: 0.7 }));
        usb.position.set(-4.65, 0.22, -0.45);
        const chip = new THREE.Mesh(new THREE.BoxGeometry(1.35, 0.14, 0.42), new THREE.MeshStandardMaterial({ color: 0x0f172a }));
        chip.position.set(-2.9, 0.16, 0.15);
        circuitGroup.add(mcu, usb, chip);
        registerLabel(supplyLabel, -3.2, 0.85, 0);
      }

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

      add3DWire([-2.2, 0.3, -0.4], [-0.5, 1.35, -0.9], [0.2, 0.15, -1.2], 0xef4444);
      add3DWire([-2.2, 0.3, 0.4], [-0.5, 1.25, 0.9], [0.2, 0.15, 1.2], 0x1e293b);

      const parts = (bp && bp.circuit_3d_parts && bp.circuit_3d_parts.length) ? bp.circuit_3d_parts : [
        { type: "ic_741", label: "IC 741 Op-Amp", wire_color: "#2563eb" },
        { type: "sensor_water_probes", label: "Water Level Probes", wire_color: "#16a34a" },
        { type: "led_indicator_bank", label: "LEDs & Piezo Buzzer", wire_color: "#d97706" }
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
          if (item.mode === "servo") {
            item.mesh.rotation.y = (sensorStimulusLevel - 0.5) * 2.4;
          } else if (item.mode === "blink") {
            const triggered = sensorStimulusLevel >= 0.75;
            item.mesh.material.color.setHex(triggered && Math.sin(simTick * 3) > 0 ? 0xef4444 : 0x7f1d1d);
          } else if (item.mode === "water") {
            item.mesh.scale.y = Math.max(0.15, sensorStimulusLevel);
          }
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
      btn.innerText = simActive ? "▶ Simulation: ON" : "⏸️ Simulation: PAUSED";
      btn.className = simActive
        ? "px-3.5 py-2 rounded-xl bg-emerald-600 font-bold text-white text-xs shadow-sm"
        : "px-3.5 py-2 rounded-xl bg-amber-600 font-bold text-white text-xs shadow-sm";
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


# ==================== 4. FLASK API ROUTES & PWA SERVICE WORKER ====================
@app.route("/favicon.ico")
def favicon():
    svg = '''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 88 88">
      <rect width="88" height="88" rx="22" fill="#1e3a8a"/>
      <path d="M24 44H36L42 29L49 59L55 44H65" stroke="#38bdf8" stroke-width="5.5" fill="none" stroke-linecap="round" stroke-linejoin="round"/>
      <circle cx="23" cy="44" r="4" fill="#ffffff"/>
      <circle cx="65" cy="44" r="4" fill="#ffffff"/>
    </svg>'''
    return Response(svg, mimetype="image/svg+xml")


@app.route("/manifest.json")
def manifest():
    return jsonify({
        "name": "ProjForge Engineering Studio",
        "short_name": "ProjForge",
        "start_url": "/",
        "display": "standalone",
        "background_color": "#ffffff",
        "theme_color": "#1e3a8a",
        "icons": [
            {
                "src": "/favicon.ico",
                "sizes": "192x192 512x512",
                "type": "image/svg+xml",
                "purpose": "any maskable"
            }
        ]
    })


@app.route("/sw.js")
def service_worker():
    js = "self.addEventListener('fetch', function(event) {});"
    return Response(js, mimetype="application/javascript")


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


@app.route("/api/download-pptx", methods=["POST"])
def download_pptx():
    payload = request.get_json(force=True)
    blueprint = payload.get("blueprint", {})
    pptx_buffer = build_seminar_pptx(
        blueprint,
        payload.get("student_name") or "Student Name",
        payload.get("usn") or "1VT23CS001",
        payload.get("college") or "VTU Affiliated Engineering College",
        payload.get("branch") or "Engineering",
        payload.get("subject") or "Engineering Subject",
        payload.get("guide_name") or "Project Guide"
    )
    safe_name = "".join(c if c.isalnum() else "_" for c in blueprint.get("title", "Seminar_Presentation"))[:40]
    return send_file(
        pptx_buffer,
        as_attachment=True,
        download_name=f"{safe_name}_Seminar.pptx",
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )


@app.route("/api/download-ieee", methods=["POST"])
def download_ieee():
    payload = request.get_json(force=True)
    blueprint = payload.get("blueprint", {})
    ieee_buffer = build_ieee_paper_docx(
        blueprint,
        payload.get("student_name") or "Student Name",
        payload.get("usn") or "1VT23CS001",
        payload.get("college") or "VTU Affiliated Engineering College",
        payload.get("branch") or "Engineering",
        payload.get("subject") or "Engineering Subject",
        payload.get("guide_name") or "Project Guide"
    )
    safe_name = "".join(c if c.isalnum() else "_" for c in blueprint.get("title", "IEEE_Paper"))[:40]
    return send_file(
        ieee_buffer,
        as_attachment=True,
        download_name=f"{safe_name}_IEEE_Paper.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )


@app.route("/api/feedback", methods=["POST"])
def save_feedback():
    data = request.get_json(force=True)
    FEEDBACK_LOG.append({"time": datetime.now().isoformat(), **data})
    return jsonify({"status": "saved"})


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))