import io
import os
import re
import json
from datetime import datetime, timezone, timedelta
from flask import Flask, render_template_string, request, jsonify, send_file
import razorpay
import requests
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

app = Flask(__name__)

# ==================== CONFIGURATION ====================
LAUNCH_START_STR = os.environ.get("LAUNCH_START_ISO", "2026-11-23T12:00:00+00:00")
FREE_HOURS = int(os.environ.get("FREE_LAUNCH_HOURS", "48"))
PRICE_INR = 49

RAZORPAY_KEY_ID = os.environ.get("RAZORPAY_KEY_ID", "").strip()
RAZORPAY_KEY_SECRET = os.environ.get("RAZORPAY_KEY_SECRET", "").strip()
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "").strip()

rzp_client = None
if RAZORPAY_KEY_ID and RAZORPAY_KEY_SECRET:
    rzp_client = razorpay.Client(auth=(RAZORPAY_KEY_ID, RAZORPAY_KEY_SECRET))


def is_free_launch_active():
    try:
        start_dt = datetime.fromisoformat(LAUNCH_START_STR)
        end_dt = start_dt + timedelta(hours=FREE_HOURS)
        now_dt = datetime.now(timezone.utc)
        if now_dt < end_dt:
            return True, int((end_dt - now_dt).total_seconds())
    except Exception:
        pass
    return False, 0


# ==================== 1. GEMINI AI ENGINE (VIA DIRECT REST API) ====================
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
    {{"name": "Component/Hardware 1", "spec": "Detailed technical specification", "cost_inr": 450}},
    {{"name": "Component/Hardware 2", "spec": "Detailed technical specification", "cost_inr": 250}},
    {{"name": "Component/Hardware 3", "spec": "Detailed technical specification", "cost_inr": 180}},
    {{"name": "Component/Hardware 4", "spec": "Detailed technical specification", "cost_inr": 120}}
  ],
  "software_requirements": [
    "OS: Windows 11 / Ubuntu 22.04 LTS",
    "IDE / Tool: VS Code / Arduino IDE / MATLAB / Python 3.11",
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
    {{"from_pin": "Controller Pin / Module A", "to_pin": "Sensor/Module B Pin", "signal": "VCC / GND / Digital / I2C / API"}},
    {{"from_pin": "Controller Pin / Module C", "to_pin": "Sensor/Module D Pin", "signal": "Signal Type"}},
    {{"from_pin": "Controller Pin / Module E", "to_pin": "Output/Actuator Pin", "signal": "PWM / UART / HTTP"}},
    {{"from_pin": "Power Rail (5V/3.3V)", "to_pin": "All Peripherals", "signal": "Regulated DC Power"}}
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
    {{"q_en": "Viva Question 10?", "a_en": "Answer 10", "q_kn": "ಕನ್ನಡ ಪ್ರಶ್ನೆ 10?", "a_kn": "ಕನ್ನಡ ಉತ್ತರ 10"}},
    {{"q_en": "Viva Question 11?", "a_en": "Answer 11", "q_kn": "ಕನ್ನಡ ಪ್ರಶ್ನೆ 11?", "a_kn": "ಕನ್ನಡ ಉತ್ತರ 11"}},
    {{"q_en": "Viva Question 12?", "a_en": "Answer 12", "q_kn": "ಕನ್ನಡ ಪ್ರಶ್ನೆ 12?", "a_kn": "ಕನ್ನಡ ಉತ್ತರ 12"}}
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
            {"name": "ESP32 / Arduino Uno R3 Microcontroller", "spec": "32-bit Dual-Core / ATmega328P, 5V/3.3V logic", "cost_inr": 480},
            {"name": "Primary Sensor / Transducer Module", "spec": "High-precision industrial grade sensor array", "cost_inr": 320},
            {"name": "0.96-inch I2C OLED Display & Buzzer", "spec": "128x64 SSD1306 display + 5V active piezo", "cost_inr": 220},
            {"name": "Relay Driver & Regulated Power Supply", "spec": "5V 2A SMPS adapter + optocoupler relay", "cost_inr": 250}
        ],
        "software_requirements": [
            "Operating System: Windows 10/11 or Ubuntu Linux",
            "Development Environment: Arduino IDE 2.3 / VS Code / Python 3.11",
            "Libraries: Wire.h, Adafruit_SSD1306, NumPy / Serial Monitor"
        ],
        "block_diagram_nodes": [
            "Regulated Power (5V/3.3V)",
            "Multi-Sensor Input Layer",
            "Core MCU / AI Processor",
            "Threshold Control Engine",
            "OLED, Buzzer & Actuator"
        ],
        "pin_connections": [
            {"from_pin": "MCU GPIO D2 / SDA", "to_pin": "Sensor Data / OLED SDA", "signal": "Digital / I2C Data"},
            {"from_pin": "MCU GPIO D3 / SCL", "to_pin": "OLED Clock Pin (SCL)", "signal": "I2C Clock"},
            {"from_pin": "MCU GPIO D8", "to_pin": "Relay Module IN1", "signal": "Active-HIGH Control"},
            {"from_pin": "MCU 5V & GND", "to_pin": "Sensor & Display VCC/GND", "signal": "DC Power Rail"}
        ],
        "working_principle": (
            f"The working of '{clean_title}' is organized into three synchronized stages: Data Acquisition, "
            "Signal Processing, and Automated Actuation. During initialization, the controller calibrates baseline "
            "sensor offsets and establishes I2C communication with the status display. In the continuous loop phase, "
            "raw physical parameters are sampled every 100 milliseconds, filtered using a moving-average noise-reduction "
            "algorithm, and compared against pre-calibrated safety thresholds. Whenever an anomaly or trigger condition "
            "is detected, the controller immediately energizes the output driver/actuator, triggers the audio-visual alert, "
            "and logs telemetry frames over the serial/cloud interface."
        ),
        "source_code_language": "C++ (Embedded Firmware)",
        "source_code": f"""// ============================================================================
// Project: {clean_title}
// Branch : {branch} | Level: {level}
// Generated by ProjForge AI
// ============================================================================

#include <Wire.h>

#define SENSOR_PIN A0
#define RELAY_PIN  8
#define BUZZER_PIN 9
#define THRESHOLD  520

float filteredValue = 0.0;
const float ALPHA = 0.25;

void setup() {{
  Serial.begin(115200);
  pinMode(SENSOR_PIN, INPUT);
  pinMode(RELAY_PIN, OUTPUT);
  pinMode(BUZZER_PIN, OUTPUT);

  digitalWrite(RELAY_PIN, LOW);
  digitalWrite(BUZZER_PIN, LOW);

  Serial.println("SYSTEM INITIALIZED: {clean_title}");
  delay(1000);
}}

void loop() {{
  int rawReading = analogRead(SENSOR_PIN);
  filteredValue = (ALPHA * rawReading) + ((1.0 - ALPHA) * filteredValue);

  Serial.print("Raw Sensor: ");
  Serial.print(rawReading);
  Serial.print(" | Filtered: ");
  Serial.println(filteredValue, 2);

  if (filteredValue > THRESHOLD) {{
    digitalWrite(RELAY_PIN, HIGH);
    digitalWrite(BUZZER_PIN, HIGH);
    Serial.println("[ALERT] Threshold Exceeded -> Actuator ON");
  }} else {{
    digitalWrite(RELAY_PIN, LOW);
    digitalWrite(BUZZER_PIN, LOW);
    Serial.println("[NORMAL] Parameters Within Safe Range");
  }}

  delay(250);
}}
""",
        "results_and_discussion": (
            "The prototype was tested across 50 consecutive operational cycles. The moving-average filter reduced "
            "sensor jitter by 78%, while actuation latency remained consistently under 180 ms. Total bill-of-materials "
            "cost was kept under Rs. 1,300, demonstrating strong feasibility for student lab demonstrations."
        ),
        "advantages": [
            "Low hardware cost and compact breadboard/PCB footprint",
            "Fast response latency (< 200 ms) with built-in noise filtering",
            "Modular architecture easily extendable to IoT cloud dashboards",
            "User-friendly real-time status visualization"
        ],
        "applications": [
            "Smart Campus & Laboratory Automation",
            "Industrial Safety & Preventive Maintenance",
            "Energy Conservation & Resource Monitoring",
            "Academic Research & Prototyping"
        ],
        "future_scope": (
            "In future iterations, the system can be upgraded with an MQTT cloud broker, mobile push notifications, "
            "and a lightweight TinyML anomaly prediction model trained on historical sensor logs."
        ),
        "references": [
            "[1] IEEE Std 1451-2024, Standard for Smart Transducer Interface for Sensors and Actuators.",
            "[2] Muhammad Ali Mazidi, 'Embedded Systems & Microcontroller Architecture', Pearson Education.",
            "[3] VTU Academic Project Guidelines & Technical Documentation Standards (2025-26)."
        ],
        "viva_qa": [
            {"q_en": "What is the primary objective of your project?", "a_en": f"To automate and optimize {clean_title} with high accuracy and low cost.", "q_kn": "ನಿಮ್ಮ ಪ್ರಾಜೆಕ್ಟ್‌ನ ಮುಖ್ಯ ಉದ್ದೇಶವೇನು?", "a_kn": f"{clean_title} ಅನ್ನು ಕಡಿಮೆ ವೆಚ್ಚದಲ್ಲಿ ಮತ್ತು ಹೆಚ್ಚಿನ ನಿಖರತೆಯೊಂದಿಗೆ ಸ್ವಯಂಚಾಲಿತಗೊಳಿಸುವುದು."},
            {"q_en": "Why did you choose this specific controller/algorithm?", "a_en": "It offers low power consumption, fast I/O response, and easy library support.", "q_kn": "ನೀವು ಈ ಕಂಟ್ರೋಲರ್/ಅಲ್ಗಾರಿದಮ್ ಅನ್ನು ಏಕೆ ಆಯ್ಕೆ ಮಾಡಿದ್ದೀರಿ?", "a_kn": "ಇದು ಕಡಿಮೆ ವಿದ್ಯುತ್ ಬಳಸುತ್ತದೆ ಮತ್ತು ವೇಗವಾಗಿ ಕಾರ್ಯನಿರ್ವಹಿಸುತ್ತದೆ."},
            {"q_en": "How do you reduce sensor noise in your system?", "a_en": "By applying a software moving-average filter before threshold comparison.", "q_kn": "ನಿಮ್ಮ ಸಿಸ್ಟಮ್‌ನಲ್ಲಿ ಸೆನ್ಸರ್ ನಾಯ್ಸ್ (Noise) ಹೇಗೆ ಕಡಿಮೆ ಮಾಡುತ್ತೀರಿ?", "a_kn": "ಸಾಫ್ಟ್‌ವೇರ್ ಮೂವಿಂಗ್-ಆವರೇಜ್ ಫಿಲ್ಟರ್ ಬಳಸುವ ಮೂಲಕ ನಾಯ್ಸ್ ಕಡಿಮೆ ಮಾಡಲಾಗುತ್ತದೆ."},
            {"q_en": "What is the total cost of your prototype?", "a_en": "Approximately Rs. 1,270 including microcontroller, sensors, and power supply.", "q_kn": "ನಿಮ್ಮ ಪ್ರಾಜೆಕ್ಟ್ ಮಾದರಿಯ ಒಟ್ಟು ವೆಚ್ಚ ಎಷ್ಟು?", "a_kn": "ಸುಮಾರು ₹1,270 ವೆಚ್ಚವಾಗುತ್ತದೆ."},
            {"q_en": "What happens if the power supply fluctuates?", "a_en": "A regulated 5V DC adapter with decoupling capacitors protects the logic circuit.", "q_kn": "ವಿದ್ಯುತ್ ಪೂರೈಕೆಯಲ್ಲಿ ಏರುಪೇರಾದರೆ ಏನಾಗುತ್ತದೆ?", "a_kn": "ರೆಗ್ಯುಲೇಟೆಡ್ 5V DC ಅಡಾಪ್ಟರ್ ಸರ್ಕ್ಯೂಟ್ ಅನ್ನು ಸುರಕ್ಷಿತವಾಗಿಡುತ್ತದೆ."},
            {"q_en": "How can this project be scaled for real-world industry?", "a_en": "By adding industrial-grade IP65 enclosures, LoRaWAN/Wi-Fi telemetry, and cloud logging.", "q_kn": "ಇದನ್ನು ಕೈಗಾರಿಕಾ ಮಟ್ಟಕ್ಕೆ ಹೇಗೆ ವಿಸ್ತರಿಸಬಹುದು?", "a_kn": "Wi-Fi/IoT ಕ್ಲೌಡ್ ಡ್ಯಾಶ್‌ಬೋರ್ಡ್ ಮತ್ತು ಕೈಗಾರಿಕಾ ಸೆನ್ಸರ್‌ಗಳನ್ನು ಸೇರಿಸುವ ಮೂಲಕ ವಿಸ್ತರಿಸಬಹುದು."},
            {"q_en": "What is the response time of your system?", "a_en": "Less than 200 milliseconds from sensor trigger to output actuation.", "q_kn": "ನಿಮ್ಮ ಸಿಸ್ಟಮ್‌ನ ರೆಸ್ಪಾನ್ಸ್ ಸಮಯ ಎಷ್ಟು?", "a_kn": "200 ಮಿಲಿಸೆಕೆಂಡ್‌ಗಳಿಗಿಂತ ಕಡಿಮೆ ಸಮಯದಲ್ಲಿ ಪ್ರತಿಕ್ರಿಯಿಸುತ್ತದೆ."},
            {"q_en": "Which communication protocol is used for the display?", "a_en": "I2C (Inter-Integrated Circuit) protocol using SDA and SCL lines.", "q_kn": "ಡಿಸ್‌ಪ್ಲೇಗಾಗಿ ಯಾವ ಕಮ್ಯುನಿಕೇಷನ್ ಪ್ರೋಟೋಕಾಲ್ ಬಳಸಲಾಗಿದೆ?", "a_kn": "SDA ಮತ್ತು SCL ಲೈನ್‌ಗಳನ್ನು ಹೊಂದಿರುವ I2C ಪ್ರೋಟೋಕಾಲ್ ಬಳಸಲಾಗಿದೆ."},
            {"q_en": "What are the main limitations of your current prototype?", "a_en": "It currently operates locally without long-range cellular backup.", "q_kn": "ಪ್ರಸ್ತುತ ಮಾದರಿಯ ಮುಖ್ಯ ಮಿತಿಗಳೇನು?", "a_kn": "ಪ್ರಸ್ತುತ ಇದು ಸ್ಥಳೀಯವಾಗಿ ಕಾರ್ಯನಿರ್ವಹಿಸುತ್ತದೆ; ದೀರ್ಘ-ಶ್ರೇಣಿಯ ಕ್ಲೌಡ್ ಬ್ಯಾಕಪ್ ಸೇರಿಸಬಹುದು."},
            {"q_en": "How did you test and validate the results?", "a_en": "By running 50 repeated test trials and comparing sensor logs against standard instruments.", "q_kn": "ಫಲಿತಾಂಶಗಳನ್ನು ನೀವು ಹೇಗೆ ಪರೀಕ್ಷಿಸಿದ್ದೀರಿ?", "a_kn": "50 ಬಾರಿ ಪರೀಕ್ಷೆ ನಡೆಸಿ ಸ್ಟ್ಯಾಂಡರ್ಡ್ ಮೀಟರ್‌ಗಳೊಂದಿಗೆ ಹೋಲಿಸಿ ದೃಢಪಡಿಸಲಾಗಿದೆ."}
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


# ==================== 3. FRONTEND HTML (4s FLASH + LOGIN + UI) ====================
HTML_PAGE = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>ProjForge AI — Engineering Project & VTU Report Studio</title>
  <script src="https://cdn.tailwindcss.com"></script>
  <script src="https://checkout.razorpay.com/v1/checkout.js"></script>
  <link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;700&family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;600&display=swap" rel="stylesheet">
  <style>
    body { font-family: 'Inter', sans-serif; background: #060913; color: #f1f5f9; overflow-x: hidden; }
    h1, h2, h3, .brand-font { font-family: 'Space Grotesk', sans-serif; }
    code, pre { font-family: 'JetBrains Mono', monospace; }

    #splash-screen {
      position: fixed; inset: 0; z-index: 9999;
      background: radial-gradient(circle at center, #0f172a 0%, #020617 100%);
      display: flex; flex-direction: column; align-items: center; justify-content: center;
      transition: opacity 0.7s ease, visibility 0.7s ease;
    }
    .cyber-ring {
      width: 112px; height: 112px; border-radius: 50%;
      border: 3px solid transparent; border-top-color: #38bdf8; border-right-color: #818cf8;
      animation: spinRing 1.1s linear infinite;
      box-shadow: 0 0 40px rgba(56, 189, 248, 0.5);
      display: flex; align-items: center; justify-content: center;
    }
    .cyber-core {
      width: 70px; height: 70px; border-radius: 18px;
      background: linear-gradient(135deg, #2563eb, #7c3aed);
      display: flex; align-items: center; justify-content: center;
      font-weight: 800; font-size: 25px; color: white;
      animation: pulseCore 1.4s ease-in-out infinite alternate;
    }
    @keyframes spinRing { 100% { transform: rotate(360deg); } }
    @keyframes pulseCore { from { transform: scale(0.92); } to { transform: scale(1.07); } }
    .flash-bar { width: 240px; height: 4px; background: #1e293b; border-radius: 999px; overflow: hidden; margin-top: 22px; }
    .flash-progress { height: 100%; width: 0%; background: linear-gradient(90deg, #38bdf8, #818cf8, #c084fc); animation: loadBar 3.7s ease forwards; }
    @keyframes loadBar { 0% { width: 0%; } 60% { width: 78%; } 100% { width: 100%; } }

    .glass { background: rgba(15, 23, 42, 0.75); backdrop-filter: blur(16px); border: 1px solid rgba(148, 163, 184, 0.15); }
    .watermark-box { position: relative; overflow: hidden; }
    .watermark-box::after {
      content: "PROJFORGE AI • PREVIEW • DOWNLOAD DOCX FOR FULL 15+ PAGES";
      position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%) rotate(-18deg);
      font-size: 1.1rem; font-weight: 800; letter-spacing: 2px;
      color: rgba(56, 189, 248, 0.11); white-space: nowrap; pointer-events: none;
    }
  </style>
</head>
<body>

  <!-- 1. 4-SEC FLASH ANIMATION -->
  <div id="splash-screen">
    <div class="cyber-ring"><div class="cyber-core">PF</div></div>
    <h1 class="text-3xl md:text-4xl font-bold mt-6 tracking-tight bg-gradient-to-r from-sky-400 via-indigo-400 to-purple-400 bg-clip-text text-transparent">
      ProjForge AI Studio
    </h1>
    <p id="splash-status" class="text-xs text-slate-400 mt-2 tracking-widest uppercase">
      Initializing Neural Circuit &amp; VTU Report Engine...
    </p>
    <div class="flash-bar"><div class="flash-progress"></div></div>
  </div>

  <!-- 2. SMART EMAIL LOGIN MODAL -->
  <div id="login-modal" class="fixed inset-0 z-50 hidden items-center justify-center bg-black/80 backdrop-blur-md p-4">
    <div class="glass max-w-md w-full rounded-2xl p-6 md:p-8 shadow-2xl border border-sky-500/30">
      <div class="flex items-center gap-3 mb-4">
        <div class="w-10 h-10 rounded-xl bg-gradient-to-tr from-blue-600 to-violet-600 flex items-center justify-center font-bold text-white">PF</div>
        <div>
          <h2 class="text-xl font-bold text-white">Student Workspace Sign-In</h2>
          <p class="text-xs text-slate-400">Auto-fills your Name &amp; USN on VTU .docx cover pages</p>
        </div>
      </div>

      <form id="login-form" class="space-y-3.5 mt-4">
        <div>
          <label class="text-xs text-slate-300 font-medium">Student Email ID *</label>
          <input id="login-email" type="email" required placeholder="prajwal@college.edu.in"
            class="w-full mt-1 px-3.5 py-2.5 rounded-xl bg-slate-900/90 border border-slate-700 text-sm focus:border-sky-400 focus:outline-none" />
        </div>
        <div class="grid grid-cols-2 gap-3">
          <div>
            <label class="text-xs text-slate-300 font-medium">Full Name *</label>
            <input id="login-name" type="text" required placeholder="Prajwal Metri"
              class="w-full mt-1 px-3.5 py-2.5 rounded-xl bg-slate-900/90 border border-slate-700 text-sm focus:border-sky-400 focus:outline-none" />
          </div>
          <div>
            <label class="text-xs text-slate-300 font-medium">USN / Roll No *</label>
            <input id="login-usn" type="text" required placeholder="1VT23CS045"
              class="w-full mt-1 px-3.5 py-2.5 rounded-xl bg-slate-900/90 border border-slate-700 text-sm uppercase focus:border-sky-400 focus:outline-none" />
          </div>
        </div>
        <div>
          <label class="text-xs text-slate-300 font-medium">College Name</label>
          <input id="login-college" type="text" placeholder="VTU Affiliated Engineering College"
            class="w-full mt-1 px-3.5 py-2.5 rounded-xl bg-slate-900/90 border border-slate-700 text-sm focus:border-sky-400 focus:outline-none" />
        </div>
        <button type="submit"
          class="w-full py-3 rounded-xl bg-gradient-to-r from-sky-500 to-indigo-600 hover:from-sky-400 hover:to-indigo-500 font-semibold text-white shadow-lg shadow-sky-500/25 transition">
          Enter ProjForge Studio →
        </button>
      </form>
      <button onclick="skipLogin()" class="w-full mt-3 text-xs text-slate-400 hover:text-slate-200 underline">
        Continue as Guest
      </button>
    </div>
  </div>

  <!-- 3. HEADER -->
  <header class="sticky top-0 z-40 glass border-b border-slate-800">
    <div class="max-w-7xl mx-auto px-4 py-3 flex items-center justify-between">
      <div class="flex items-center gap-2.5">
        <div class="w-9 h-9 rounded-xl bg-gradient-to-br from-sky-500 to-indigo-600 flex items-center justify-center font-bold text-white shadow">PF</div>
        <div>
          <span class="brand-font font-bold text-lg text-white">ProjForge <span class="text-sky-400">AI</span></span>
          <span class="hidden sm:inline-block ml-2 text-[11px] px-2 py-0.5 rounded-full bg-sky-500/15 text-sky-300 border border-sky-500/30">v2.0 Pro</span>
        </div>
      </div>

      <div class="flex items-center gap-3">
        {% if free_active %}
        <span class="text-xs px-3 py-1.5 rounded-full bg-emerald-500/15 text-emerald-300 border border-emerald-500/30 font-medium">
          🎉 2-Day Free Launch Active!
        </span>
        {% else %}
        <span class="text-xs px-3 py-1.5 rounded-full bg-amber-500/15 text-amber-300 border border-amber-500/30 font-medium">
          Full 15+ Page Pack: ₹{{ price_inr }}
        </span>
        {% endif %}
        <button onclick="openLoginModal()" id="user-badge"
          class="text-xs px-3 py-1.5 rounded-xl bg-slate-800 hover:bg-slate-700 border border-slate-700 text-slate-200">
          👤 Student Login
        </button>
      </div>
    </div>
  </header>

  <!-- 4. GENERATOR SECTION -->
  <main class="max-w-7xl mx-auto px-4 py-8">
    <section class="text-center max-w-3xl mx-auto mb-8">
      <h1 class="text-3xl md:text-5xl font-bold leading-tight bg-gradient-to-r from-white via-sky-200 to-indigo-300 bg-clip-text text-transparent">
        Build Engineering Projects, Circuits &amp; 15+ Page VTU Reports in Seconds
      </h1>
      <p class="text-slate-400 text-sm md:text-base mt-3">
        Gemini AI Engine • Auto Pin-to-Pin Circuit Diagrams • Working Source Code • English &amp; Kannada Viva Q&amp;A
      </p>
    </section>

    <section class="glass rounded-2xl p-5 md:p-7 max-w-4xl mx-auto shadow-2xl border border-slate-700/70">
      <div class="grid grid-cols-1 md:grid-cols-3 gap-4">
        <div class="md:col-span-3">
          <label class="text-xs font-semibold uppercase tracking-wider text-sky-400">Enter Your Project Topic / Idea *</label>
          <input id="topic-input" type="text"
            placeholder="e.g., IoT Smart Helmet with Alcohol & Accident Detection / AI Fake News Detector"
            class="w-full mt-1.5 px-4 py-3.5 rounded-xl bg-slate-950/90 border border-slate-700 text-white text-sm md:text-base focus:border-sky-400 focus:outline-none" />
        </div>

        <div>
          <label class="text-xs text-slate-400">Engineering Branch</label>
          <select id="branch-input" class="w-full mt-1 px-3.5 py-2.5 rounded-xl bg-slate-950 border border-slate-700 text-sm text-white">
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
          <label class="text-xs text-slate-400">Project Level</label>
          <select id="level-input" class="w-full mt-1 px-3.5 py-2.5 rounded-xl bg-slate-950 border border-slate-700 text-sm text-white">
            <option>2nd / 3rd Year Mini-Project</option>
            <option>Final Year Major Capstone Project</option>
            <option>Hackathon / Innovation Prototype</option>
          </select>
        </div>

        <div>
          <label class="text-xs text-slate-400">Guide Name (For Certificate)</label>
          <input id="guide-input" type="text" placeholder="Dr. S. K. Kulkarni"
            class="w-full mt-1 px-3.5 py-2.5 rounded-xl bg-slate-950 border border-slate-700 text-sm text-white" />
        </div>
      </div>

      <div class="flex flex-wrap items-center gap-2 mt-4">
        <span class="text-xs text-slate-400">Try:</span>
        <button type="button" onclick="setSample('Smart Automated Railway Gate Control using IR & ESP32')" class="text-xs px-2.5 py-1 rounded-lg bg-slate-800/90 hover:bg-slate-700 text-sky-300 border border-slate-700">🚂 Smart Railway Gate</button>
        <button type="button" onclick="setSample('AI-Powered Leaf Disease Detection using CNN & Python')" class="text-xs px-2.5 py-1 rounded-lg bg-slate-800/90 hover:bg-slate-700 text-sky-300 border border-slate-700">🌿 Crop Disease AI</button>
        <button type="button" onclick="setSample('RFID & Biometric Smart Attendance System with Cloud Log')" class="text-xs px-2.5 py-1 rounded-lg bg-slate-800/90 hover:bg-slate-700 text-sky-300 border border-slate-700">🪪 Smart Attendance</button>
      </div>

      <button id="generate-btn" onclick="generateProject()"
        class="w-full mt-5 py-3.5 rounded-xl bg-gradient-to-r from-sky-500 via-blue-600 to-indigo-600 hover:opacity-95 font-bold text-white text-base shadow-xl shadow-blue-600/25 transition">
        ⚡ Generate Complete Project, Circuit &amp; VTU Report
      </button>
    </section>

    <!-- 5. RESULTS WORKSPACE -->
    <section id="result-section" class="hidden mt-10 space-y-6">
      <div class="glass rounded-2xl p-5 flex flex-col md:flex-row items-start md:items-center justify-between gap-4 border border-sky-500/30">
        <div>
          <span class="text-xs uppercase tracking-wider text-sky-400 font-semibold">Generated Blueprint Ready</span>
          <h2 id="res-title" class="text-xl md:text-2xl font-bold text-white mt-0.5"></h2>
          <p id="res-tagline" class="text-xs text-slate-400 mt-1"></p>
        </div>

        <button onclick="initiateDocxDownload()" id="download-docx-btn"
          class="px-5 py-3 rounded-xl bg-gradient-to-r from-emerald-500 to-teal-600 hover:from-emerald-400 hover:to-teal-500 font-bold text-white text-sm shadow-lg shadow-emerald-600/25 shrink-0">
          {% if free_active %}
          📥 Download 15+ Page VTU Report (.docx) — FREE
          {% else %}
          📥 Unlock &amp; Download 15+ Page .docx (₹{{ price_inr }})
          {% endif %}
        </button>
      </div>

      <div class="flex flex-wrap gap-2 border-b border-slate-800 pb-3">
        <button onclick="switchTab('tab-report', this)" class="tab-btn px-4 py-2 rounded-xl bg-sky-600 text-white text-xs font-semibold">📄 Watermarked Report Preview</button>
        <button onclick="switchTab('tab-circuit', this)" class="tab-btn px-4 py-2 rounded-xl bg-slate-800 text-slate-300 text-xs font-semibold">🔌 Block &amp; Pin Circuit Diagram</button>
        <button onclick="switchTab('tab-code', this)" class="tab-btn px-4 py-2 rounded-xl bg-slate-800 text-slate-300 text-xs font-semibold">💻 Working Source Code</button>
        <button onclick="switchTab('tab-viva', this)" class="tab-btn px-4 py-2 rounded-xl bg-slate-800 text-slate-300 text-xs font-semibold">🎓 Viva Q&amp;A (English + ಕನ್ನಡ)</button>
      </div>

      <!-- TAB 1 -->
      <div id="tab-report" class="tab-panel glass watermark-box rounded-2xl p-6 space-y-5">
        <div class="flex items-center justify-between">
          <h3 class="text-lg font-bold text-sky-300">Abstract &amp; Executive Overview</h3>
          <span class="text-[11px] px-2.5 py-1 rounded bg-slate-800 text-slate-400">Preview Mode • Full Certificate &amp; Chapters in .docx</span>
        </div>
        <p id="res-abstract" class="text-sm text-slate-300 leading-relaxed"></p>
        <div class="grid grid-cols-1 md:grid-cols-2 gap-5">
          <div class="bg-slate-900/80 p-4 rounded-xl border border-slate-800">
            <h4 class="text-sm font-bold text-indigo-300 mb-2">🎯 Project Objectives</h4>
            <ul id="res-objectives" class="list-disc list-inside text-xs text-slate-300 space-y-1.5"></ul>
          </div>
          <div class="bg-slate-900/80 p-4 rounded-xl border border-slate-800">
            <h4 class="text-sm font-bold text-emerald-300 mb-2">💰 Estimated Hardware &amp; BOM Budget</h4>
            <div id="res-bom" class="text-xs text-slate-300 space-y-1.5"></div>
          </div>
        </div>
        <div class="bg-slate-900/80 p-4 rounded-xl border border-slate-800">
          <h4 class="text-sm font-bold text-sky-300 mb-2">⚙️ Working Methodology</h4>
          <p id="res-working" class="text-xs text-slate-300 leading-relaxed"></p>
        </div>
      </div>

      <!-- TAB 2 -->
      <div id="tab-circuit" class="tab-panel hidden glass rounded-2xl p-6 space-y-5">
        <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div>
            <h3 class="text-lg font-bold text-sky-300">Auto-Generated System Block &amp; Pin-Interfacing Diagram</h3>
            <p class="text-xs text-slate-400">Download high-resolution PNG diagram for your PPT or lab record</p>
          </div>
          <button onclick="downloadDiagramPNG()" class="px-4 py-2 rounded-xl bg-indigo-600 hover:bg-indigo-500 text-xs font-semibold text-white">
            🖼️ Download Diagram Image (.PNG)
          </button>
        </div>
        <div class="overflow-x-auto bg-slate-950 rounded-xl p-4 border border-slate-800 flex justify-center">
          <canvas id="circuit-canvas" width="860" height="360" class="max-w-full rounded-lg"></canvas>
        </div>
        <div>
          <h4 class="text-sm font-bold text-indigo-300 mb-2">📌 Pin-to-Pin Connection Table</h4>
          <div class="overflow-x-auto">
            <table class="w-full text-xs text-left border border-slate-800">
              <thead class="bg-slate-900 text-sky-300">
                <tr>
                  <th class="p-2.5 border-b border-slate-800">Controller / Source Pin</th>
                  <th class="p-2.5 border-b border-slate-800">Sensor / Target Module Pin</th>
                  <th class="p-2.5 border-b border-slate-800">Signal / Protocol</th>
                </tr>
              </thead>
              <tbody id="res-pins" class="divide-y divide-slate-800 text-slate-300"></tbody>
            </table>
          </div>
        </div>
      </div>

      <!-- TAB 3 -->
      <div id="tab-code" class="tab-panel hidden glass rounded-2xl p-6 space-y-4">
        <div class="flex items-center justify-between">
          <h3 id="res-code-lang" class="text-lg font-bold text-emerald-300">Source Code</h3>
          <button onclick="copySourceCode()" class="px-3.5 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-xs text-sky-300 border border-slate-700">
            📋 Copy Code
          </button>
        </div>
        <pre id="res-code" class="bg-slate-950 p-4 rounded-xl border border-slate-800 text-xs text-emerald-300 overflow-x-auto leading-relaxed"></pre>
      </div>

      <!-- TAB 4 -->
      <div id="tab-viva" class="tab-panel hidden glass rounded-2xl p-6 space-y-4">
        <h3 class="text-lg font-bold text-purple-300">🎓 Top External Viva-Voce Questions (English &amp; ಕನ್ನಡ)</h3>
        <div id="res-viva" class="grid grid-cols-1 md:grid-cols-2 gap-4"></div>
      </div>
    </section>
  </main>

  <script>
    let currentBlueprint = null;
    let studentProfile = JSON.parse(localStorage.getItem("pf_student") || "null");

    window.addEventListener("DOMContentLoaded", () => {
      const steps = ["Loading Gemini Neural Architecture...", "Calibrating Circuit Renderer...", "Preparing VTU 15+ Page DOCX Engine..."];
      let idx = 0;
      const iv = setInterval(() => {
        idx++;
        if (idx < steps.length) document.getElementById("splash-status").innerText = steps[idx];
      }, 1150);

      setTimeout(() => {
        clearInterval(iv);
        const s = document.getElementById("splash-screen");
        s.style.opacity = "0";
        s.style.visibility = "hidden";
        if (!studentProfile) openLoginModal();
        else updateUserBadge();
      }, 3800);
    });

    function openLoginModal() {
      const m = document.getElementById("login-modal");
      m.classList.remove("hidden"); m.classList.add("flex");
    }
    function closeLoginModal() {
      const m = document.getElementById("login-modal");
      m.classList.add("hidden"); m.classList.remove("flex");
    }
    function skipLogin() {
      if (!studentProfile) studentProfile = { email: "", name: "Student Name", usn: "1VT23CS001", college: "VTU Affiliated Engineering College" };
      closeLoginModal();
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
      updateUserBadge();
      closeLoginModal();
    });
    function updateUserBadge() {
      if (studentProfile && studentProfile.name) {
        document.getElementById("user-badge").innerText = `👤 ${studentProfile.name} (${studentProfile.usn})`;
      }
    }
    function setSample(txt) { document.getElementById("topic-input").value = txt; }
    function switchTab(tabId, btn) {
      document.querySelectorAll(".tab-panel").forEach(p => p.classList.add("hidden"));
      document.getElementById(tabId).classList.remove("hidden");
      document.querySelectorAll(".tab-btn").forEach(b => { b.classList.remove("bg-sky-600", "text-white"); b.classList.add("bg-slate-800", "text-slate-300"); });
      btn.classList.remove("bg-slate-800", "text-slate-300"); btn.classList.add("bg-sky-600", "text-white");
    }

    async function generateProject() {
      const topic = document.getElementById("topic-input").value.trim();
      const branch = document.getElementById("branch-input").value;
      const level = document.getElementById("level-input").value;
      if (!topic) return alert("Please enter your project topic first!");

      const btn = document.getElementById("generate-btn");
      btn.disabled = true;
      btn.innerText = "⏳ Generating Blueprint, Circuit & Viva Q&A...";

      try {
        const res = await fetch("/api/generate", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ topic, branch, level })
        });
        const data = await res.json();
        if (data.error) throw new Error(data.error);
        currentBlueprint = data.blueprint;
        renderBlueprint(currentBlueprint);
        document.getElementById("result-section").classList.remove("hidden");
        document.getElementById("result-section").scrollIntoView({ behavior: "smooth" });
      } catch (err) {
        alert("Error: " + err.message);
      } finally {
        btn.disabled = false;
        btn.innerText = "⚡ Generate Complete Project, Circuit & VTU Report";
      }
    }

    function renderBlueprint(bp) {
      document.getElementById("res-title").innerText = bp.title || "";
      document.getElementById("res-tagline").innerText = bp.tagline || "";
      document.getElementById("res-abstract").innerText = bp.abstract || "";
      document.getElementById("res-working").innerText = bp.working_principle || "";
      document.getElementById("res-objectives").innerHTML = (bp.objectives || []).map(o => `<li>${o}</li>`).join("");

      let total = 0;
      document.getElementById("res-bom").innerHTML = (bp.hardware_requirements || []).map(h => {
        total += Number(h.cost_inr || 0);
        return `<div class="flex justify-between border-b border-slate-800 pb-1"><span>${h.name}</span><span class="text-emerald-400 font-semibold">₹${h.cost_inr}</span></div>`;
      }).join("") + `<div class="flex justify-between pt-1 font-bold text-white"><span>Total Prototype Cost</span><span>₹${total}</span></div>`;

      document.getElementById("res-pins").innerHTML = (bp.pin_connections || []).map(p => `
        <tr><td class="p-2.5 font-mono text-sky-300">${p.from_pin}</td><td class="p-2.5 font-mono text-indigo-300">${p.to_pin}</td><td class="p-2.5">${p.signal}</td></tr>
      `).join("");

      document.getElementById("res-code-lang").innerText = `💻 ${bp.source_code_language || "Source Code"}`;
      document.getElementById("res-code").innerText = bp.source_code || "";

      document.getElementById("res-viva").innerHTML = (bp.viva_qa || []).map((v, i) => `
        <div class="bg-slate-900/90 p-4 rounded-xl border border-slate-800 space-y-1.5">
          <p class="text-xs font-bold text-sky-300">Q${i + 1}: ${v.q_en}</p>
          <p class="text-xs text-slate-300">${v.a_en}</p>
          ${v.q_kn ? `<p class="text-xs font-bold text-amber-300 pt-1">ಪ್ರಶ್ನೆ: ${v.q_kn}</p><p class="text-xs text-slate-300">${v.a_kn}</p>` : ""}
        </div>
      `).join("");

      drawCircuitDiagram(bp);
    }

    function drawCircuitDiagram(bp) {
      const canvas = document.getElementById("circuit-canvas");
      const ctx = canvas.getContext("2d");
      ctx.fillStyle = "#020617"; ctx.fillRect(0, 0, canvas.width, canvas.height);
      ctx.strokeStyle = "rgba(56, 189, 248, 0.08)";
      for (let x = 0; x < canvas.width; x += 25) { ctx.beginPath(); ctx.moveTo(x, 0); ctx.lineTo(x, canvas.height); ctx.stroke(); }
      for (let y = 0; y < canvas.height; y += 25) { ctx.beginPath(); ctx.moveTo(0, y); ctx.lineTo(canvas.width, y); ctx.stroke(); }

      ctx.fillStyle = "#38bdf8"; ctx.font = "bold 15px Space Grotesk, sans-serif";
      ctx.fillText("SYSTEM ARCHITECTURE: " + (bp.title || "").toUpperCase().slice(0, 52), 24, 30);

      const nodes = (bp.block_diagram_nodes || []).slice(0, 5);
      const boxW = 142, boxH = 110, startX = 24, gap = 25, y = 115;
      nodes.forEach((node, idx) => {
        const x = startX + idx * (boxW + gap);
        ctx.fillStyle = idx === 2 ? "#1e3a8a" : "#0f172a";
        ctx.strokeStyle = idx === 2 ? "#38bdf8" : "#6366f1";
        ctx.lineWidth = 2;
        ctx.beginPath(); ctx.roundRect(x, y, boxW, boxH, 12); ctx.fill(); ctx.stroke();

        ctx.fillStyle = "#94a3b8"; ctx.font = "bold 10px Inter, sans-serif";
        ctx.fillText(`STAGE 0${idx + 1}`, x + 12, y + 22);

        ctx.fillStyle = "#f8fafc"; ctx.font = "bold 11px Inter, sans-serif";
        const words = node.split(" ");
        let line = "", lineY = y + 46;
        words.forEach(w => {
          if ((line + w).length > 17) { ctx.fillText(line, x + 12, lineY); line = w + " "; lineY += 16; }
          else { line += w + " "; }
        });
        ctx.fillText(line, x + 12, lineY);

        if (idx < nodes.length - 1) {
          const ax1 = x + boxW + 3, ax2 = x + boxW + gap - 3, ay = y + boxH / 2;
          ctx.strokeStyle = "#38bdf8"; ctx.beginPath(); ctx.moveTo(ax1, ay); ctx.lineTo(ax2, ay); ctx.stroke();
          ctx.fillStyle = "#38bdf8"; ctx.beginPath(); ctx.moveTo(ax2, ay); ctx.lineTo(ax2 - 6, ay - 4); ctx.lineTo(ax2 - 6, ay + 4); ctx.fill();
        }
      });
      ctx.fillStyle = "#64748b"; ctx.font = "11px Inter, sans-serif";
      ctx.fillText("Generated by ProjForge AI Studio • Ready for VTU Report & PPT", 24, 335);
    }

    function downloadDiagramPNG() {
      const link = document.createElement("a");
      link.download = "ProjForge_Circuit_Diagram.png";
      link.href = document.getElementById("circuit-canvas").toDataURL("image/png");
      link.click();
    }
    function copySourceCode() {
      navigator.clipboard.writeText(document.getElementById("res-code").innerText);
      alert("Source code copied!");
    }

    async function initiateDocxDownload() {
      if (!currentBlueprint) return;
      const btn = document.getElementById("download-docx-btn");
      btn.disabled = true; btn.innerText = "⏳ Preparing 15+ Page VTU DOCX...";
      try {
        const orderRes = await fetch("/api/create-order", { method: "POST" });
        const orderData = await orderRes.json();
        if (orderData.free) {
          await triggerDocxFile({});
        } else {
          const rzp = new Razorpay({
            key: orderData.key_id, amount: orderData.amount, currency: orderData.currency,
            name: "ProjForge AI", description: "15+ Page VTU Report + Code + Bilingual Viva",
            order_id: orderData.order_id,
            handler: async (resp) => { await triggerDocxFile(resp); },
            theme: { color: "#2563eb" }
          });
          rzp.open();
        }
      } catch (err) {
        alert("Download error: " + err.message);
      } finally {
        btn.disabled = false; btn.innerText = "📥 Download 15+ Page VTU Report (.docx)";
      }
    }

    async function triggerDocxFile(paymentProof) {
      const payload = {
        ...paymentProof,
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
      if (!res.ok) { const e = await res.json(); throw new Error(e.error || "Failed"); }
      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a"); a.href = url; a.download = "ProjForge_VTU_Project_Report.docx";
      document.body.appendChild(a); a.click(); a.remove();
    }
  </script>
</body>
</html>
"""


# ==================== 4. FLASK API ROUTES ====================
@app.route("/")
def index():
    free_active, rem_seconds = is_free_launch_active()
    return render_template_string(
        HTML_PAGE,
        free_active=free_active,
        rem_seconds=rem_seconds,
        price_inr=PRICE_INR,
    )


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


@app.route("/api/create-order", methods=["POST"])
def create_order():
    free_active, _ = is_free_launch_active()
    if free_active:
        return jsonify({"free": True})
    if not rzp_client:
        return jsonify({"error": "Payment gateway keys not configured yet."}), 500
    order = rzp_client.order.create({
        "amount": PRICE_INR * 100,
        "currency": "INR",
        "payment_capture": 1,
        "notes": {"product": "ProjForge AI 15+ Page VTU Report & Code"}
    })
    return jsonify({
        "free": False,
        "order_id": order["id"],
        "amount": order["amount"],
        "currency": order["currency"],
        "key_id": RAZORPAY_KEY_ID
    })


@app.route("/api/download-docx", methods=["POST"])
def download_docx():
    payload = request.get_json(force=True)
    free_active, _ = is_free_launch_active()
    if not free_active and rzp_client:
        payment_id = payload.get("razorpay_payment_id", "")
        order_id = payload.get("razorpay_order_id", "")
        signature = payload.get("razorpay_signature", "")
        if not (payment_id and order_id and signature):
            return jsonify({"error": "Payment verification required."}), 402
        try:
            rzp_client.utility.verify_payment_signature({
                "razorpay_order_id": order_id,
                "razorpay_payment_id": payment_id,
                "razorpay_signature": signature
            })
        except Exception:
            return jsonify({"error": "Invalid payment signature."}), 402

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