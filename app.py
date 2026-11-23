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


# ==================== 1. GEMINI AI ENGINE (DYNAMIC 3D + MULTI-LANG + DUAL CODE) ====================
def generate_project_blueprint(topic, branch, level):
    if GEMINI_API_KEY:
        try:
            url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={GEMINI_API_KEY}"
            prompt = f"""
You are a senior Indian Engineering Professor (VTU / Autonomous curriculum) and Principal Systems Architect.
Generate a comprehensive engineering project blueprint in strict JSON format for:
- Project Title / Topic: "{topic}"
- Engineering Branch: "{branch}"
- Academic Level: "{level}"

IMPORTANT RULES:
1. "primary_code_language" and "primary_code": MUST match the project domain! If the project is AI/ML, Data Science, Image Processing, or Software, write complete runnable Python code (e.g., TensorFlow/OpenCV/Flask/Scikit-learn, 60-90 lines). If it is purely embedded/IoT, write C++/Python.
2. "has_arduino_hardware": true if sensors/motors/IoT/embedded hardware is applicable, else false.
3. "arduino_code": Always provide a complete, compilable Arduino (.ino) sketch if hardware/sensors can be interfaced with this project (or an optional IoT/hardware edge node sketch).
4. "circuit_3d_parts": Provide 3 to 5 specific 3D hardware modules tailored specifically to THIS project (choose types from: "sensor_ultrasonic", "sensor_gas", "sensor_ir", "sensor_temp", "camera_module", "motor_servo", "relay_module", "oled_display", "buzzer_alarm", "led_indicator").
5. "viva_qa": Provide 8 high-value viva questions & answers translated naturally into all 5 languages: English ("en"), Kannada ("kn"), Hindi ("hi"), Telugu ("te"), and Tamil ("ta").

Return ONLY valid JSON matching this exact schema (no markdown fences):
{{
  "title": "Formal Engineering Project Title",
  "tagline": "One-line technical summary",
  "abstract": "Detailed 240-word academic abstract covering problem statement, methodology, core tech, and results.",
  "problem_statement": "Detailed 150-word problem definition explaining drawbacks in existing systems.",
  "objectives": ["Objective 1", "Objective 2", "Objective 3", "Objective 4"],
  "literature_survey": [
    {{"paper": "Author et al. (2024) - Title", "summary": "Key finding and research gap."}},
    {{"paper": "Author et al. (2025) - Title", "summary": "Key finding and research gap."}},
    {{"paper": "Author et al. (2026) - Title", "summary": "Key finding and research gap."}}
  ],
  "hardware_requirements": [
    {{"name": "Component 1", "spec": "Specification", "cost_inr": 450}},
    {{"name": "Component 2", "spec": "Specification", "cost_inr": 250}},
    {{"name": "Component 3", "spec": "Specification", "cost_inr": 180}},
    {{"name": "Breadboard & Jumpers", "spec": "830-Point Breadboard + Wires", "cost_inr": 140}}
  ],
  "software_requirements": [
    "OS: Windows 11 / Ubuntu 22.04",
    "IDE: VS Code / Python 3.11 / Arduino IDE",
    "Libraries: Project specific libraries"
  ],
  "controller_board": "Arduino Uno R3 or ESP32 DevKit",
  "circuit_3d_parts": [
    {{"type": "camera_module", "label": "Vision Sensor / Cam", "pin_from": "GPIO 21 / USB", "pin_to": "Breadboard Row 10", "wire_color": "#38bdf8"}},
    {{"type": "oled_display", "label": "0.96 I2C OLED", "pin_from": "SDA/SCL (A4/A5)", "pin_to": "Breadboard Row 22", "wire_color": "#a855f7"}},
    {{"type": "buzzer_alarm", "label": "Alert Buzzer & LED", "pin_from": "Digital Pin D8", "pin_to": "Breadboard Row 34", "wire_color": "#f59e0b"}}
  ],
  "pin_connections": [
    {{"from_pin": "Controller Pin 1", "to_pin": "Module 1 Pin", "signal": "Signal Type"}},
    {{"from_pin": "Controller Pin 2", "to_pin": "Module 2 Pin", "signal": "Signal Type"}},
    {{"from_pin": "Controller 5V", "to_pin": "Breadboard +VCC Rail", "signal": "5V DC Power"}},
    {{"from_pin": "Controller GND", "to_pin": "Breadboard -GND Rail", "signal": "Common Ground"}}
  ],
  "ppt_slides": [
    {{"slide": "Slide 1: Title & Team", "points": "Project Title, USN, Guide Name & Department"}},
    {{"slide": "Slide 2: Problem Statement", "points": "Limitations of existing systems & motivation"}},
    {{"slide": "Slide 3: Proposed Architecture", "points": "System flow, algorithms & hardware/software stack"}},
    {{"slide": "Slide 4: Implementation", "points": "Code pipeline, circuit interfacing & testing"}},
    {{"slide": "Slide 5: Results & Conclusion", "points": "Accuracy metrics, applications & future scope"}}
  ],
  "working_principle": "Detailed 320-word step-by-step explanation of system architecture, data flow, and execution.",
  "primary_code_language": "Python 3 (AI / Application Code)",
  "primary_code": "Complete runnable Python or primary language source code (60-90 lines).",
  "has_arduino_hardware": true,
  "arduino_code": "Complete Arduino C++ (.ino) firmware code for hardware/sensor interfacing.",
  "results_and_discussion": "Detailed 180-word analysis of testing parameters, accuracy, and performance.",
  "advantages": ["Advantage 1", "Advantage 2", "Advantage 3"],
  "applications": ["Application 1", "Application 2", "Application 3"],
  "future_scope": "120-word discussion on scalability and future upgrades.",
  "references": [
    "[1] IEEE Standard / Research Paper 1",
    "[2] Reference Book / Documentation 2",
    "[3] IEEE / Springer Journal Article 3 (2025)"
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

    return _fallback_blueprint(topic, branch, level)


def _fallback_blueprint(topic, branch, level):
    clean_title = topic.strip().title() or "Smart Automated Engineering System"
    low = clean_title.lower()
    is_ai = any(k in low for k in ["ai", "cnn", "detection", "machine learning", "deep learning", "python", "fake", "disease", "image"])

    if is_ai:
        parts = [
            {"type": "camera_module", "label": "OV7670 / USB Vision Cam", "pin_from": "ESP32 GPIO 21", "pin_to": "Breadboard Row 08", "wire_color": "#38bdf8"},
            {"type": "oled_display", "label": "0.96 I2C OLED Status", "pin_from": "GPIO 22 (SCL)", "pin_to": "Breadboard Row 22", "wire_color": "#a855f7"},
            {"type": "buzzer_alarm", "label": "Alert Buzzer & LED", "pin_from": "GPIO 18 (OUT)", "pin_to": "Breadboard Row 35", "wire_color": "#f59e0b"}
        ]
        prim_lang = "Python 3 (OpenCV & CNN Inference)"
        prim_code = f"""# ============================================================================
# Project: {clean_title}
# Branch : {branch} | Level: {level}
# ============================================================================

import cv2
import numpy as np
import time

CONFIDENCE_THRESHOLD = 0.82
CLASSES = ["Healthy / Normal", "Anomaly / Alert Detected"]

def preprocess_frame(frame):
    resized = cv2.resize(frame, (224, 224))
    normalized = resized.astype("float32") / 255.0
    return np.expand_dims(normalized, axis=0)

def run_inference(tensor):
    # Simulated CNN feature activation score
    mean_intensity = float(np.mean(tensor))
    score = min(0.98, max(0.45, mean_intensity * 1.45))
    label = CLASSES[1] if score >= CONFIDENCE_THRESHOLD else CLASSES[0]
    return label, round(score * 100, 2)

def main():
    print("==================================================")
    print("INITIALIZING: {clean_title}")
    print("==================================================")
    sample_input = np.random.rand(480, 640, 3) * 255
    tensor = preprocess_frame(sample_input)
    prediction, confidence = run_inference(tensor)
    print(f"[RESULT] Prediction: {{prediction}} | Confidence: {{confidence}}%")

if __name__ == "__main__":
    main()
"""
    else:
        parts = [
            {"type": "sensor_gas" if "gas" in low or "fire" in low or "alcohol" in low else "sensor_ir", "label": "Primary Sensor Module", "pin_from": "Arduino D2 / A0", "pin_to": "Breadboard Row 10", "wire_color": "#38bdf8"},
            {"type": "motor_servo" if "gate" in low or "railway" in low or "door" in low else "relay_module", "label": "Servo / Relay Actuator", "pin_from": "Arduino PWM D9", "pin_to": "Breadboard Row 24", "wire_color": "#f59e0b"},
            {"type": "oled_display", "label": "128x64 I2C OLED Display", "pin_from": "Arduino A4/A5", "pin_to": "Breadboard Row 38", "wire_color": "#a855f7"}
        ]
        prim_lang = "Python 3 (Serial Telemetry & Dashboard)"
        prim_code = f"""# ============================================================================
# Project: {clean_title} — Python Telemetry Monitor
# ============================================================================
import time

THRESHOLD = 520

def process_reading(val):
    status = "ALERT_TRIGGERED" if val > THRESHOLD else "SAFE_NORMAL"
    print(f"[TELEMETRY] Sensor={{val}} | Status={{status}}")

if __name__ == "__main__":
    for sample in [310, 420, 615, 290]:
        process_reading(sample)
        time.sleep(0.3)
"""

    arduino_code = f"""// ============================================================================
// Optional Hardware Firmware (.ino) for: {clean_title}
// Board: Arduino Uno R3 / ESP32
// ============================================================================

#define SENSOR_PIN 2
#define ACTUATOR_PIN 9
#define BUZZER_PIN 8

void setup() {{
  Serial.begin(115200);
  pinMode(SENSOR_PIN, INPUT);
  pinMode(ACTUATOR_PIN, OUTPUT);
  pinMode(BUZZER_PIN, OUTPUT);
  Serial.println("HARDWARE READY: {clean_title}");
}}

void loop() {{
  int state = digitalRead(SENSOR_PIN);
  if (state == HIGH) {{
    digitalWrite(ACTUATOR_PIN, HIGH);
    digitalWrite(BUZZER_PIN, HIGH);
    Serial.println("[TRIGGER] Hardware Actuator ON");
  }} else {{
    digitalWrite(ACTUATOR_PIN, LOW);
    digitalWrite(BUZZER_PIN, LOW);
    Serial.println("[NORMAL] Monitoring...");
  }}
  delay(200);
}}
"""

    return {
        "title": clean_title,
        "tagline": f"An intelligent {branch} engineering solution for {level} academic implementation.",
        "abstract": (
            f"This project presents the design, simulation, and prototype implementation of '{clean_title}' "
            f"tailored for {branch} engineering applications. Conventional systems suffer from manual latency, "
            "high operational overhead, and lack of real-time telemetry. To overcome these limitations, the proposed "
            "architecture integrates precision sensing, algorithmic/AI processing, and automated actuation. "
            "Experimental validation confirms a 94.8% operational accuracy with sub-200ms response latency."
        ),
        "problem_statement": (
            f"Existing manual approaches related to {clean_title} require continuous human supervision, "
            "are prone to human error, and lack automated fault alerts."
        ),
        "objectives": [
            f"To design a modular architecture for {clean_title}.",
            "To minimize latency and hardware cost while maximizing accuracy.",
            "To implement real-time detection and automated alert/control mechanisms.",
            "To validate performance across multiple test scenarios."
        ],
        "literature_survey": [
            {"paper": "R. Kumar et al. (2024) - IEEE Access", "summary": "Proposed baseline monitoring; lacked real-time edge optimization."},
            {"paper": "S. Kulkarni & P. Rao (2025) - Springer", "summary": "Demonstrated sensor/AI fusion; high BOM cost limited student prototyping."},
            {"paper": "A. Sharma et al. (2026) - Elsevier", "summary": "Introduced automated analytics; our work simplifies deployment."}
        ],
        "hardware_requirements": [
            {"name": "Arduino Uno R3 / ESP32 DevKit", "spec": "Main Microcontroller Unit", "cost_inr": 480},
            {"name": parts[0]["label"], "spec": "Primary Input Module", "cost_inr": 280},
            {"name": parts[1]["label"], "spec": "Output / Control Module", "cost_inr": 220},
            {"name": "830-Point Breadboard & Jumpers", "spec": "Solderless Prototyping Board", "cost_inr": 140}
        ],
        "software_requirements": [
            "OS: Windows 11 / Ubuntu Linux",
            "IDE: VS Code / Python 3.11 / Arduino IDE 2.3",
            "Libraries: OpenCV / NumPy / Wire.h"
        ],
        "controller_board": "Arduino Uno R3 / ESP32",
        "circuit_3d_parts": parts,
        "pin_connections": [
            {"from_pin": parts[0]["pin_from"], "to_pin": f"{parts[0]['label']} ({parts[0]['pin_to']})", "signal": "Input Signal"},
            {"from_pin": parts[1]["pin_from"], "to_pin": f"{parts[1]['label']} ({parts[1]['pin_to']})", "signal": "Control Output"},
            {"from_pin": "Controller 5V", "to_pin": "Breadboard +VCC Red Rail", "signal": "+5V DC Power"},
            {"from_pin": "Controller GND", "to_pin": "Breadboard -GND Blue Rail", "signal": "Common Ground"}
        ],
        "ppt_slides": [
            {"slide": "Slide 1: Title & Team", "points": f"{clean_title} — Team USN & Guide"},
            {"slide": "Slide 2: Problem Statement", "points": "Drawbacks of existing manual systems"},
            {"slide": "Slide 3: System Architecture", "points": "3D Hardware layout, software pipeline & interfacing"},
            {"slide": "Slide 4: Implementation", "points": "Core algorithm/code & breadboard circuit testing"},
            {"slide": "Slide 5: Results & Future Scope", "points": "94.8% accuracy, applications & scalability"}
        ],
        "working_principle": (
            f"The working of '{clean_title}' operates in three synchronized stages: Data Acquisition, "
            "Intelligent Processing, and Automated Output Control. Input signals are captured, filtered against "
            "noise, evaluated by the core algorithm/controller, and converted into real-time alerts or actuation."
        ),
        "primary_code_language": prim_lang,
        "primary_code": prim_code,
        "has_arduino_hardware": True,
        "arduino_code": arduino_code,
        "results_and_discussion": (
            "The system was tested across 50 operational trials, achieving 94.8% accuracy with sub-200ms response time."
        ),
        "advantages": [
            "High accuracy and fast response time (< 200 ms)",
            "Low prototype cost and modular design",
            "Easy to test on 830-point solderless breadboard"
        ],
        "applications": [
            "Smart Campus & Lab Automation",
            "Industrial Monitoring & Safety",
            "Academic Mini & Major Projects"
        ],
        "future_scope": "Can be scaled with cloud dashboards, mobile alerts, and custom PCB fabrication.",
        "references": [
            "[1] IEEE Standard Technical Documentation (2025).",
            "[2] Muhammad Ali Mazidi, 'Embedded Systems & Architecture', Pearson.",
            "[3] VTU Academic Project Guidelines (2025-26)."
        ],
        "viva_qa": [
            {
                "en": {"q": "What is the main objective of your project?", "a": f"To automate and optimize {clean_title} with high accuracy and low cost."},
                "kn": {"q": "ನಿಮ್ಮ ಪ್ರಾಜೆಕ್ಟ್‌ನ ಮುಖ್ಯ ಉದ್ದೇಶವೇನು?", "a": f"{clean_title} ಅನ್ನು ಕಡಿಮೆ ವೆಚ್ಚದಲ್ಲಿ ಮತ್ತು ಹೆಚ್ಚಿನ ನಿಖರತೆಯೊಂದಿಗೆ ಸ್ವಯಂಚಾಲಿತಗೊಳಿಸುವುದು."},
                "hi": {"q": "आपके प्रोजेक्ट का मुख्य उद्देश्य क्या है?", "a": f"{clean_title} को कम लागत और उच्च सटीकता के साथ स्वचालित करना।"},
                "te": {"q": "మీ ప్రాజెక్ట్ యొక్క ముఖ్య ఉద్దేశ్యం ఏమిటి?", "a": f"{clean_title} ని తక్కువ ఖర్చుతో మరియు అధిక ఖచ్చితత్వంతో ఆటోమేట్ చేయడం."},
                "ta": {"q": "உங்கள் திட்டத்தின் முக்கிய நோக்கம் என்ன?", "a": f"{clean_title} ஐ குறைந்த செலவில் மற்றும் அதிக துல்லியத்துடன் தானியக்கமாக்குவது."}
            },
            {
                "en": {"q": "Why is a breadboard used in hardware prototyping?", "a": "It allows solderless circuit assembly, quick testing, and easy pin debugging."},
                "kn": {"q": "ಹಾರ್ಡ್‌ವೇರ್ ಪ್ರೊಟೊಟೈಪಿಂಗ್‌ನಲ್ಲಿ ಬ್ರೆಡ್‌ಬೋರ್ಡ್ ಏಕೆ ಬಳಸುತ್ತೇವೆ?", "a": "ಸೋಲ್ಡರಿಂಗ್ ಇಲ್ಲದೆ ಸುಲಭವಾಗಿ ಕನೆಕ್ಷನ್ ಬದಲಾಯಿಸಲು ಮತ್ತು ಪರೀಕ್ಷಿಸಲು ಬಳಸುತ್ತೇವೆ."},
                "hi": {"q": "हार्डवेयर प्रोटोटाइपिंग में ब्रेडबोर्ड का उपयोग क्यों किया जाता है?", "a": "यह बिना सोल्डरिंग के आसान सर्किट कनेक्शन और टेस्टिंग की सुविधा देता है।"},
                "te": {"q": "హార్డ్‌వేర్ ప్రోటోటైపింగ్‌లో బ్రెడ్‌బోర్డ్ ఎందుకు ఉపయోగిస్తారు?", "a": "సోల్డరింగ్ లేకుండా సులభంగా సర్క్యూట్ కనెక్షన్లు మరియు టెస్టింగ్ చేయడానికి ఉపయోగిస్తారు."},
                "ta": {"q": "வன்பொருள் முன்மாதிரியில் பிரெட்போர்டு ஏன் பயன்படுத்தப்படுகிறது?", "a": "சாலிடரிங் இல்லாமல் எளிதாக சர்க்யூட் இணைப்புகளை மாற்றவும் சோதிக்கவும் இது உதவுகிறது."}
            },
            {
                "en": {"q": "How do you improve accuracy and reduce false triggers?", "a": "By applying threshold filtering and signal normalization before decision logic."},
                "kn": {"q": "ನಿಖರತೆಯನ್ನು ಹೆಚ್ಚಿಸಲು ಮತ್ತು ತಪ್ಪು ಟ್ರಿಗರ್ ತಡೆಯಲು ಏನು ಮಾಡಿದ್ದೀರಿ?", "a": "ನಿರ್ಧಾರ ತೆಗೆದುಕೊಳ್ಳುವ ಮುನ್ನ ಫಿಲ್ಟರಿಂಗ್ ಮತ್ತು ಥ್ರೆಶೋಲ್ಡ್ ಪರಿಶೀಲನೆ ಬಳಸಲಾಗಿದೆ."},
                "hi": {"q": "सटीकता बढ़ाने और गलत ट्रिगर कम करने के लिए आप क्या करते हैं?", "a": "निर्णय लेने से पहले सिग्नल फ़िल्टरिंग और थ्रेशोल्ड जांच का उपयोग किया जाता है।"},
                "te": {"q": "ఖచ్చితత్వాన్ని పెంచడానికి మరియు తప్పుడు ట్రిగ్గర్‌లను తగ్గించడానికి మీరు ఏమి చేస్తారు?", "a": "నిర్ణయం తీసుకునే ముందు సిగ్నల్ ఫిల్టరింగ్ మరియు థ్రెషోల్డ్ తనిఖీని ఉపయోగిస్తాము."},
                "ta": {"q": "துல்லியத்தை அதிகரிக்கவும் தவறான தூண்டுதல்களைக் குறைக்கவும் என்ன செய்கிறீர்கள்?", "a": "முடிவெடுப்பதற்கு முன் சிக்னல் வடிகட்டுதல் மற்றும் வரம்பு சரிபார்ப்பு பயன்படுத்தப்படுகிறது."}
            },
            {
                "en": {"q": "What is the response time and estimated budget of your system?", "a": "Response time is under 200 ms and total prototype cost is around Rs. 1,120."},
                "kn": {"q": "ನಿಮ್ಮ ಸಿಸ್ಟಮ್‌ನ ರೆಸ್ಪಾನ್ಸ್ ಸಮಯ ಮತ್ತು ಒಟ್ಟು ವೆಚ್ಚ ಎಷ್ಟು?", "a": "ರೆಸ್ಪಾನ್ಸ್ ಸಮಯ 200 ms ಗಿಂತ ಕಡಿಮೆ ಇದ್ದು, ಒಟ್ಟು ವೆಚ್ಚ ಸುಮಾರು ₹1,120 ಆಗುತ್ತದೆ."},
                "hi": {"q": "आपके सिस्टम का रिस्पॉन्स टाइम और कुल बजट क्या है?", "a": "रिस्पॉन्स टाइम 200 ms से कम है और कुल लागत लगभग ₹1,120 है।"},
                "te": {"q": "మీ సిస్టమ్ యొక్క రెస్పాన్స్ సమయం మరియు మొత్తం బడ్జెట్ ఎంత?", "a": "రెస్పాన్స్ సమయం 200 ms కంటే తక్కువగా ఉంటుంది మరియు మొత్తం ఖర్చు సుమಾರು ₹1,120."},
                "ta": {"q": "உங்கள் அமைப்பின் பதில் நேரம் மற்றும் மொத்த செலவு என்ன?", "a": "பதில் நேரம் 200 ms க்கும் குறைவாக உள்ளது மற்றும் மொத்த செலவு சுமார் ₹1,120 ஆகும்."}
            },
            {
                "en": {"q": "How can this project be scaled for real-world deployment?", "a": "By integrating cloud IoT dashboards, mobile alerts, and custom PCB hardware."},
                "kn": {"q": "ಇದನ್ನು ನೈಜ-ಪ್ರಪಂಚದ ಬಳಕೆಗೆ ಹೇಗೆ ವಿಸ್ತರಿಸಬಹುದು?", "a": "ಕ್ಲೌಡ್ IoT ಡ್ಯಾಶ್‌ಬೋರ್ಡ್, ಮೊಬೈಲ್ ಅಲರ್ಟ್ ಮತ್ತು ಕಸ್ಟಮ್ PCB ಬಳಸುವ ಮೂಲಕ ವಿಸ್ತರಿಸಬಹುದು."},
                "hi": {"q": "इसे वास्तविक दुनिया में उपयोग के लिए कैसे बढ़ाया जा सकता है?", "a": "क्लाउड IoT डैशबोर्ड, मोबाइल अलर्ट और कस्टम PCB जोड़कर इसे बढ़ाया जा सकता है।"},
                "te": {"q": "దీనిని వాస్తవ ప్రపంచ వినియోగానికి ఎలా విస్తరించవచ్చు?", "a": "క్లౌడ్ IoT డాష్‌బోర్డ్, మొబైల్ అలర్ట్‌లు మరియు కస్టమ్ PCB జోడించడం ద్వారా విస్తరించవచ్చు."},
                "ta": {"q": "நடைமுறை பயன்பாட்டிற்கு இதை எவ்வாறு விரிவுபடுத்தலாம்?", "a": "கிளவுட் IoT டாஷ்போர்டு, மொபைல் எச்சரிக்கைகள் மற்றும் தனிப்பயன் PCB மூலம் விரிவுபடுத்தலாம்."}
            }
        ]
    }


# ==================== 2. CUSTOMIZABLE VTU DOCX REPORT BUILDER ====================
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


def build_vtu_docx(blueprint, student_name, usn, college, branch, guide_name, fmt):
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
        ("A PROJECT REPORT ON\n\n", font_size, False, True, None),
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

    # PAGE 2: CERTIFICATE
    add_h(f"{college.upper()}", 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_h(f"Department of {branch}", 2).alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_h("BONAFIDE CERTIFICATE", 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_b(
        f'Certified that the Project work entitled "{title}" is a bonafide work carried out by '
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

    # PAGE 3: ABSTRACT & TOC
    add_h("ABSTRACT", 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_b(blueprint.get("abstract", ""))
    add_h("TABLE OF CONTENTS", 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
    toc_items = [
        ("1.", "CHAPTER 1: INTRODUCTION & LITERATURE SURVEY", "04"),
        ("2.", "CHAPTER 2: HARDWARE ARCHITECTURE & PIN MAPPING", "06"),
        ("3.", "CHAPTER 3: BILL OF MATERIALS & SOFTWARE STACK", "08"),
        ("4.", "CHAPTER 4: METHODOLOGY & SOURCE CODE", "10"),
        ("5.", "CHAPTER 5: RESULTS, APPLICATIONS & REFERENCES", "13"),
        ("6.", "APPENDIX-A: VIVA-VOCE QUESTIONS & ANSWERS", "15"),
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
    add_h("CHAPTER 1: INTRODUCTION & LITERATURE SURVEY", 1)
    add_h("1.1 Problem Statement", 2)
    add_b(blueprint.get("problem_statement", ""))
    add_h("1.2 Objectives", 2)
    for obj in blueprint.get("objectives", []):
        p = doc.add_paragraph(obj, style="List Bullet")
        p.runs[0].font.name = font_family
        p.runs[0].font.size = Pt(font_size)
    add_h("1.3 Literature Survey", 2)
    for item in blueprint.get("literature_survey", []):
        add_b(f"• {item.get('paper', '')}: {item.get('summary', '')}", align=WD_ALIGN_PARAGRAPH.LEFT)
    doc.add_page_break()

    # CHAPTER 2 & 3
    add_h("CHAPTER 2: HARDWARE PIN MAPPING & BOM", 1)
    pins = blueprint.get("pin_connections", [])
    if pins:
        pt = doc.add_table(rows=len(pins) + 1, cols=3)
        pt.style = "Table Grid"
        for idx, h in enumerate(["Controller Pin", "Breadboard / Module Pin", "Signal"]):
            c = pt.cell(0, idx)
            c.text = h
            _set_cell_shading(c, "1E3A8A")
            c.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
        for i, pin in enumerate(pins, start=1):
            pt.cell(i, 0).text = pin.get("from_pin", "")
            pt.cell(i, 1).text = pin.get("to_pin", "")
            pt.cell(i, 2).text = pin.get("signal", "")

    add_h("CHAPTER 3: HARDWARE & SOFTWARE REQUIREMENTS", 1)
    for hw in blueprint.get("hardware_requirements", []):
        add_b(f"• {hw.get('name', '')} ({hw.get('spec', '')}) — Est. Rs. {hw.get('cost_inr', 0)}", align=WD_ALIGN_PARAGRAPH.LEFT)
    for sw in blueprint.get("software_requirements", []):
        add_b(f"• {sw}", align=WD_ALIGN_PARAGRAPH.LEFT)
    doc.add_page_break()

    # CHAPTER 4
    add_h("CHAPTER 4: WORKING & SOURCE CODE", 1)
    add_b(blueprint.get("working_principle", ""))
    add_h(f"4.1 {blueprint.get('primary_code_language', 'Primary Code')}", 2)
    ct = doc.add_table(rows=1, cols=1)
    ct.style = "Table Grid"
    cc = ct.cell(0, 0)
    _set_cell_shading(cc, "F3F4F6")
    r = cc.paragraphs[0].add_run(blueprint.get("primary_code", ""))
    r.font.name = "Consolas"
    r.font.size = Pt(9.5)
    doc.add_page_break()

    # CHAPTER 5 & APPENDIX
    add_h("CHAPTER 5: RESULTS, APPLICATIONS & REFERENCES", 1)
    add_b(blueprint.get("results_and_discussion", ""))
    for adv in blueprint.get("advantages", []):
        add_b("✓ Advantage: " + adv, align=WD_ALIGN_PARAGRAPH.LEFT)
    for ap in blueprint.get("applications", []):
        add_b("✓ Application: " + ap, align=WD_ALIGN_PARAGRAPH.LEFT)
    add_b("Future Scope: " + blueprint.get("future_scope", ""))
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


# ==================== 3. FRONTEND HTML (TINKERCAD 3D + 5-LANG VIVA + REPORT TOOLS) ====================
HTML_PAGE = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
  <title>ProjForge AI — 3D Circuit & VTU Project Studio</title>
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
      background: rgba(2, 6, 23, 0.88); border: 1px solid rgba(56, 189, 248, 0.5);
      color: #e0f2fe; font-size: 10px; font-weight: 600; padding: 2px 7px;
      border-radius: 6px; white-space: nowrap; box-shadow: 0 2px 8px rgba(0,0,0,0.6);
    }
  </style>
</head>
<body class="pb-20 md:pb-8">

  <!-- 1. SPLASH SCREEN -->
  <div id="splash-screen">
    <div class="cyber-ring"><div class="cyber-core">PF</div></div>
    <h1 class="text-2xl font-bold mt-4 bg-gradient-to-r from-sky-400 to-indigo-400 bg-clip-text text-transparent">ProjForge AI 3D</h1>
    <p class="text-[11px] text-slate-400 mt-1 tracking-widest uppercase">Initializing Tinkercad-Style 3D &amp; VTU Studio...</p>
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
        <button onclick="navigatePage('page-3d')" id="nav-d-3d" class="nav-d-btn px-3 py-1.5 rounded-lg text-xs font-semibold text-slate-300">🧊 3D Circuit Lab</button>
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

    <!-- ==================== PAGE 1: CREATE PROJECT ==================== -->
    <section id="page-home" class="page-section active space-y-5">
      <div class="text-center max-w-2xl mx-auto">
        <h1 class="text-2xl sm:text-4xl font-bold leading-tight bg-gradient-to-r from-white via-sky-200 to-indigo-300 bg-clip-text text-transparent">
          Engineering Project, Custom 3D Circuit &amp; VTU Report Studio
        </h1>
        <p class="text-slate-400 text-xs sm:text-sm mt-2">
          Generates topic-specific 3D Breadboard circuits, domain-accurate code (Python + optional Arduino), 5-language Viva Q&amp;A &amp; customizable VTU reports.
        </p>
      </div>

      <div class="glass rounded-2xl p-5 sm:p-7 max-w-3xl mx-auto shadow-xl border border-slate-800 space-y-4">
        <div>
          <label class="text-xs font-semibold uppercase tracking-wider text-sky-400">Project Topic / Problem Statement *</label>
          <input id="topic-input" type="text"
            placeholder="e.g., Smart Automated Railway Gate / AI Leaf Disease Detection"
            class="w-full mt-1.5 px-4 py-3.5 rounded-xl bg-slate-950 border border-slate-700 text-white text-sm focus:border-sky-400 focus:outline-none" />
        </div>

        <div class="grid grid-cols-1 sm:grid-cols-3 gap-3.5">
          <div>
            <label class="text-xs text-slate-400">Branch</label>
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
            <label class="text-xs text-slate-400">Project Level</label>
            <select id="level-input" class="w-full mt-1 px-3 py-2.5 rounded-xl bg-slate-950 border border-slate-700 text-xs sm:text-sm text-white">
              <option>2nd / 3rd Year Mini-Project</option>
              <option>Final Year Major Capstone Project</option>
              <option>Hackathon / Innovation Prototype</option>
            </select>
          </div>
          <div>
            <label class="text-xs text-slate-400">Guide Name</label>
            <input id="guide-input" type="text" placeholder="Dr. S. K. Kulkarni"
              class="w-full mt-1 px-3 py-2.5 rounded-xl bg-slate-950 border border-slate-700 text-xs sm:text-sm text-white" />
          </div>
        </div>

        <div class="flex flex-wrap items-center gap-2 pt-1">
          <span class="text-xs text-slate-400">Try:</span>
          <button type="button" onclick="setSample('Smart Automated Railway Gate Control using IR & Servo Motor')" class="text-xs px-2.5 py-1 rounded-lg bg-slate-900 hover:bg-slate-800 text-sky-300 border border-slate-700">🚂 Railway Gate (Servo+IR)</button>
          <button type="button" onclick="setSample('AI-Powered Leaf Disease Detection using CNN & Python')" class="text-xs px-2.5 py-1 rounded-lg bg-slate-900 hover:bg-slate-800 text-sky-300 border border-slate-700">🌿 Leaf Disease AI (Python)</button>
          <button type="button" onclick="setSample('IoT Gas Leakage & Fire Alert System with OLED & Buzzer')" class="text-xs px-2.5 py-1 rounded-lg bg-slate-900 hover:bg-slate-800 text-sky-300 border border-slate-700">🔥 Gas &amp; Fire Alert</button>
        </div>

        <button id="generate-btn" onclick="generateProject()"
          class="w-full py-3.5 rounded-xl bg-gradient-to-r from-sky-500 via-blue-600 to-indigo-600 font-bold text-white text-sm sm:text-base shadow-lg">
          ⚡ Generate Project, 3D Circuit &amp; Report
        </button>
      </div>
    </section>

    <!-- ==================== PAGE 2: TINKERCAD-STYLE 3D HARDWARE LAB ==================== -->
    <section id="page-3d" class="page-section space-y-4">
      <div class="glass rounded-2xl p-4 sm:p-5 border border-sky-500/30 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
        <div>
          <span class="text-[11px] uppercase tracking-wider text-sky-400 font-bold">Topic-Specific 3D Hardware Assembly</span>
          <h2 id="lab3d-title" class="text-base sm:text-xl font-bold text-white">3D Breadboard &amp; Microcontroller Circuit</h2>
          <p class="text-xs text-slate-400">1-Finger Drag to Rotate • 2-Finger Pinch / Scroll to Zoom • Smooth Orbit Controls</p>
        </div>
        <div class="flex flex-wrap gap-2">
          <button onclick="toggle3DSimulation()" id="sim-toggle-btn"
            class="px-3.5 py-2 rounded-xl bg-emerald-600 font-bold text-white text-xs shadow">
            ⚡ Power Simulation: ON
          </button>
          <button onclick="reset3DCamera()"
            class="px-3 py-2 rounded-xl bg-slate-800 text-xs text-slate-200 border border-slate-700">
            🔄 Reset Camera
          </button>
        </div>
      </div>

      <!-- 3D Viewport with Floating Component Labels -->
      <div class="glass rounded-2xl p-2 border border-slate-800 relative overflow-hidden">
        <div id="threejs-wrapper" class="relative w-full h-[380px] sm:h-[460px] rounded-xl bg-[#040814] overflow-hidden">
          <div id="threejs-container" class="w-full h-full"></div>
          <div id="labels-overlay" class="absolute inset-0 pointer-events-none overflow-hidden"></div>
        </div>
      </div>

      <!-- Pin-to-Pin Wiring Table -->
      <div class="glass rounded-2xl p-5">
        <h3 class="text-sm font-bold text-sky-300 mb-3">📌 Exact Pin-to-Pin Breadboard Connections</h3>
        <div class="overflow-x-auto">
          <table class="w-full text-xs text-left border border-slate-800">
            <thead class="bg-slate-900 text-sky-300">
              <tr>
                <th class="p-2.5 border-b border-slate-800">Controller Pin</th>
                <th class="p-2.5 border-b border-slate-800">Breadboard / Component</th>
                <th class="p-2.5 border-b border-slate-800">Signal / Bus</th>
              </tr>
            </thead>
            <tbody id="lab3d-pins" class="divide-y divide-slate-800 text-slate-300"></tbody>
          </table>
        </div>
      </div>
    </section>

    <!-- ==================== PAGE 3: CODE (PYTHON + OPTIONAL ARDUINO) & 5-LANG VIVA ==================== -->
    <section id="page-code" class="page-section space-y-5">
      <div id="empty-code" class="glass rounded-2xl p-8 text-center">
        <p class="text-sm text-slate-400">Generate a project first to view domain-specific source code and 5-language Viva Q&amp;A.</p>
        <button onclick="navigatePage('page-home')" class="mt-3 px-4 py-2 rounded-xl bg-sky-600 text-xs font-semibold text-white">Go to Create →</button>
      </div>

      <div id="content-code" class="hidden space-y-5">
        <!-- Code Switcher: Primary Code vs Optional Arduino Dump Code -->
        <div class="glass rounded-2xl p-5 space-y-3">
          <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
            <div class="flex flex-wrap gap-2">
              <button onclick="switchCodeMode('primary')" id="btn-code-primary"
                class="px-3.5 py-2 rounded-xl bg-emerald-600 text-white text-xs font-bold">
                🐍 Main Project Code
              </button>
              <button onclick="switchCodeMode('arduino')" id="btn-code-arduino"
                class="px-3.5 py-2 rounded-xl bg-slate-800 text-slate-300 text-xs font-semibold border border-slate-700">
                🔌 Optional Arduino Dump Code (.ino)
              </button>
            </div>
            <button onclick="copySourceCode()" class="px-3 py-1.5 rounded-lg bg-slate-800 text-xs text-sky-300 border border-slate-700 self-start">
              📋 Copy Active Code
            </button>
          </div>
          <p id="active-code-subtitle" class="text-xs text-emerald-300 font-semibold"></p>
          <pre id="code-block" class="bg-slate-950 p-4 rounded-xl border border-slate-800 text-xs text-emerald-300 overflow-x-auto leading-relaxed"></pre>
        </div>

        <!-- 5-Language Viva Selector -->
        <div class="glass rounded-2xl p-5 space-y-4">
          <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <div>
              <h3 class="text-base font-bold text-purple-300">🎓 Viva-Voce Q&amp;A (Select Preferred Language)</h3>
              <p class="text-xs text-slate-400">Choose one language to view clean, focused questions &amp; answers</p>
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
        <p class="text-sm text-slate-400">Generate a project first to customize and download your VTU .docx report.</p>
        <button onclick="navigatePage('page-home')" class="mt-3 px-4 py-2 rounded-xl bg-sky-600 text-xs font-semibold text-white">Go to Create →</button>
      </div>

      <div id="content-report" class="hidden space-y-5">
        <!-- Student Document Formatting Toolbar -->
        <div class="glass rounded-2xl p-5 border border-emerald-500/30 space-y-4">
          <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <div>
              <span class="text-[11px] uppercase tracking-wider text-emerald-400 font-bold">VTU / College Report Formatting Tools</span>
              <h2 id="rep-title" class="text-lg sm:text-xl font-bold text-white mt-0.5"></h2>
            </div>
            <button onclick="downloadDocxDirect()" id="download-docx-btn"
              class="px-5 py-3 rounded-xl bg-gradient-to-r from-emerald-500 to-teal-600 font-bold text-white text-xs sm:text-sm shadow-lg shrink-0">
              📥 Download Custom .docx Report
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

        <!-- Live Formatted Document Preview -->
        <div id="live-doc-preview" class="glass rounded-2xl p-6 space-y-4 border-2 border-blue-900/60">
          <h3 class="text-base font-bold text-sky-300">Live Document Preview</h3>
          <p id="rep-abstract" class="text-sm text-slate-200 leading-relaxed"></p>
          <p id="rep-working" class="text-sm text-slate-300 leading-relaxed border-t border-slate-800 pt-3"></p>
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
  </main>

  <!-- MOBILE BOTTOM NAVIGATION -->
  <nav class="md:hidden fixed bottom-0 inset-x-0 z-40 glass border-t border-slate-800 grid grid-cols-4 py-1.5 px-1">
    <button onclick="navigatePage('page-home')" id="nav-m-home" class="nav-m-btn flex flex-col items-center py-1 text-sky-400 font-bold">
      <span class="text-base">🏠</span><span class="text-[10px]">Create</span>
    </button>
    <button onclick="navigatePage('page-3d')" id="nav-m-3d" class="nav-m-btn flex flex-col items-center py-1 text-slate-400">
      <span class="text-base">🧊</span><span class="text-[10px]">3D Lab</span>
    </button>
    <button onclick="navigatePage('page-code')" id="nav-m-code" class="nav-m-btn flex flex-col items-center py-1 text-slate-400">
      <span class="text-base">💻</span><span class="text-[10px]">Code/Viva</span>
    </button>
    <button onclick="navigatePage('page-report')" id="nav-m-report" class="nav-m-btn flex flex-col items-center py-1 text-slate-400">
      <span class="text-base">📄</span><span class="text-[10px]">Report</span>
    </button>
  </nav>

  <script>
    let currentBlueprint = null;
    let studentProfile = JSON.parse(localStorage.getItem("pf_student") || "null");
    let activeCodeMode = "primary";
    let activeVivaLang = "en";

    window.addEventListener("DOMContentLoaded", () => {
      const urlParams = new URLSearchParams(window.location.search);
      const refParam = urlParams.get("ref");
      if (refParam) localStorage.setItem("pf_referred_by", refParam);

      setTimeout(() => {
        const s = document.getElementById("splash-screen");
        s.style.opacity = "0";
        s.style.visibility = "hidden";
        if (!studentProfile || !studentProfile.email) openLoginModal();
        else updateProfileUI();
      }, 2400);

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
      const m = document.getElementById("login-modal");
      m.classList.remove("hidden"); m.classList.add("flex");
    }
    function closeLoginModal() {
      const m = document.getElementById("login-modal");
      m.classList.add("hidden"); m.classList.remove("flex");
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

    // Refer
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
      window.open("https://wa.me/?text=" + encodeURIComponent(`🚀 Build Engineering Projects, 3D Circuits & VTU Reports on ProjForge AI:\n${link}`), "_blank");
    }

    function setSample(txt) { document.getElementById("topic-input").value = txt; }

    async function generateProject() {
      const topic = document.getElementById("topic-input").value.trim();
      const branch = document.getElementById("branch-input").value;
      const level = document.getElementById("level-input").value;
      if (!topic) return alert("Please enter your project topic first!");

      const btn = document.getElementById("generate-btn");
      btn.disabled = true; btn.innerText = "⏳ Building Custom 3D Circuit & Blueprint...";
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
        navigatePage("page-3d");
      } catch (err) {
        alert("Error: " + err.message);
      } finally {
        btn.disabled = false; btn.innerText = "⚡ Generate Project, 3D Circuit & Report";
      }
    }

    function populateAllPages(bp) {
      // 3D Lab
      document.getElementById("lab3d-title").innerText = `3D Circuit: ${bp.title}`;
      build3DSceneFromBlueprint(bp);
      document.getElementById("lab3d-pins").innerHTML = (bp.pin_connections || []).map(p => `
        <tr>
          <td class="p-2.5 font-mono text-sky-300">${p.from_pin}</td>
          <td class="p-2.5 font-mono text-indigo-300">${p.to_pin}</td>
          <td class="p-2.5">${p.signal}</td>
        </tr>
      `).join("");

      // Code & Viva
      document.getElementById("empty-code").classList.add("hidden");
      document.getElementById("content-code").classList.remove("hidden");
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
      }).join("") + `<div class="flex justify-between pt-1 font-bold text-white"><span>Total Cost</span><span>₹${total}</span></div>`;

      document.getElementById("rep-ppt").innerHTML = (bp.ppt_slides || []).map(s => `
        <div class="border-b border-slate-800 pb-1"><span class="text-sky-300 font-semibold">${s.slide}:</span> ${s.points}</div>
      `).join("");
    }

    // Dual Code Switcher
    function switchCodeMode(mode) {
      if (!currentBlueprint) return;
      activeCodeMode = mode;
      const pBtn = document.getElementById("btn-code-primary");
      const aBtn = document.getElementById("btn-code-arduino");
      if (mode === "primary") {
        pBtn.className = "px-3.5 py-2 rounded-xl bg-emerald-600 text-white text-xs font-bold";
        aBtn.className = "px-3.5 py-2 rounded-xl bg-slate-800 text-slate-300 text-xs font-semibold border border-slate-700";
        document.getElementById("active-code-subtitle").innerText = `Language: ${currentBlueprint.primary_code_language || "Python"}`;
        document.getElementById("code-block").innerText = currentBlueprint.primary_code || "";
      } else {
        aBtn.className = "px-3.5 py-2 rounded-xl bg-sky-600 text-white text-xs font-bold";
        pBtn.className = "px-3.5 py-2 rounded-xl bg-slate-800 text-slate-300 text-xs font-semibold border border-slate-700";
        document.getElementById("active-code-subtitle").innerText = "Optional Hardware Firmware (.ino) — Ready to Dump onto Arduino / ESP32";
        document.getElementById("code-block").innerText = currentBlueprint.arduino_code || "// No hardware code required";
      }
    }

    function copySourceCode() {
      navigator.clipboard.writeText(document.getElementById("code-block").innerText);
      alert("Code copied!");
    }

    // 5-Language Viva Selector
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

    // Report Formatting Live Preview & Download
    function updateLivePreviewStyle() {
      const box = document.getElementById("live-doc-preview");
      const fam = document.getElementById("fmt-font-family").value;
      const sz = document.getElementById("fmt-font-size").value;
      const lh = document.getElementById("fmt-line-spacing").value;
      box.style.fontFamily = fam;
      box.style.fontSize = sz + "px";
      box.style.lineHeight = lh;
    }

    async function downloadDocxDirect() {
      if (!currentBlueprint) return;
      const btn = document.getElementById("download-docx-btn");
      btn.disabled = true; btn.innerText = "⏳ Building DOCX...";
      try {
        const payload = {
          blueprint: currentBlueprint,
          student_name: (studentProfile && studentProfile.name) || "Student Name",
          usn: (studentProfile && studentProfile.usn) || "1VT23CS001",
          college: (studentProfile && studentProfile.college) || "VTU Affiliated Engineering College",
          branch: document.getElementById("branch-input").value,
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
        const a = document.createElement("a"); a.href = url; a.download = "ProjForge_Custom_VTU_Report.docx";
        document.body.appendChild(a); a.click(); a.remove();
      } catch (err) {
        alert("Download error: " + err.message);
      } finally {
        btn.disabled = false; btn.innerText = "📥 Download Custom .docx Report";
      }
    }

    // ==================== TINKERCAD-STYLE DYNAMIC 3D ENGINE ====================
    let scene, camera, renderer, controls, circuitGroup;
    let trackedLabels = [];
    let animParts = [];
    let simActive = true;

    function initThreeJS() {
      const container = document.getElementById("threejs-container");
      if (!container || typeof THREE === "undefined") return;

      scene = new THREE.Scene();
      scene.background = new THREE.Color(0x040814);

      camera = new THREE.PerspectiveCamera(42, container.clientWidth / container.clientHeight, 0.1, 1000);
      camera.position.set(0, 8.2, 11.2);

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

      scene.add(new THREE.AmbientLight(0xffffff, 0.72));
      const keyLight = new THREE.DirectionalLight(0xe0f2fe, 1.1);
      keyLight.position.set(8, 14, 10);
      scene.add(keyLight);

      const grid = new THREE.GridHelper(22, 22, 0x1e3a8a, 0x0f172a);
      grid.position.y = -0.25;
      scene.add(grid);

      circuitGroup = new THREE.Group();
      scene.add(circuitGroup);

      animate3D();
    }

    function createPartMesh(type) {
      const grp = new THREE.Group();
      if (type === "sensor_ultrasonic") {
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
      } else if (type === "oled_display") {
        const bezel = new THREE.Mesh(new THREE.BoxGeometry(1.25, 0.85, 0.14), new THREE.MeshStandardMaterial({ color: 0x1e3a8a }));
        const screen = new THREE.Mesh(new THREE.BoxGeometry(1.05, 0.62, 0.04), new THREE.MeshBasicMaterial({ color: 0x06b6d4 }));
        screen.position.z = 0.08;
        grp.add(bezel, screen);
      } else if (type === "camera_module") {
        const board = new THREE.Mesh(new THREE.BoxGeometry(1.0, 0.9, 0.16), new THREE.MeshStandardMaterial({ color: 0x0f172a }));
        const lens = new THREE.Mesh(new THREE.CylinderGeometry(0.26, 0.3, 0.35, 18), new THREE.MeshStandardMaterial({ color: 0x334155, metalness: 0.8 }));
        lens.rotation.x = Math.PI / 2; lens.position.z = 0.2;
        grp.add(board, lens);
      } else if (type === "buzzer_alarm") {
        const drum = new THREE.Mesh(new THREE.CylinderGeometry(0.38, 0.38, 0.45, 20), new THREE.MeshStandardMaterial({ color: 0x0f172a }));
        const led = new THREE.Mesh(new THREE.SphereGeometry(0.18, 16, 16), new THREE.MeshBasicMaterial({ color: 0xef4444 }));
        led.position.set(0.65, 0.1, 0);
        grp.add(drum, led);
        animParts.push({ mesh: led, mode: "blink" });
      } else {
        const mod = new THREE.Mesh(new THREE.BoxGeometry(1.15, 0.65, 0.7), new THREE.MeshStandardMaterial({ color: 0x10b981 }));
        grp.add(mod);
      }
      return grp;
    }

    function build3DSceneFromBlueprint(bp) {
      if (!circuitGroup) return;
      while (circuitGroup.children.length > 0) circuitGroup.remove(circuitGroup.children[0]);
      trackedLabels = [];
      animParts = [];
      document.getElementById("labels-overlay").innerHTML = "";

      function registerLabel(text, x, y, z) {
        const el = document.createElement("div");
        el.className = "label-3d";
        el.innerText = text;
        document.getElementById("labels-overlay").appendChild(el);
        trackedLabels.push({ el, pos: new THREE.Vector3(x, y, z) });
      }

      // 1. Detailed Arduino / ESP32 Board (Left)
      const mcuName = (bp && bp.controller_board) || "Arduino Uno R3";
      const mcu = new THREE.Mesh(
        new THREE.BoxGeometry(3.3, 0.22, 2.2),
        new THREE.MeshStandardMaterial({ color: 0x1d4ed8, roughness: 0.35 })
      );
      mcu.position.set(-3.2, 0, 0);
      circuitGroup.add(mcu);

      const usbPort = new THREE.Mesh(new THREE.BoxGeometry(0.65, 0.36, 0.55), new THREE.MeshStandardMaterial({ color: 0xcbd5e1, metalness: 0.75 }));
      usbPort.position.set(-4.65, 0.22, -0.45);
      circuitGroup.add(usbPort);

      const icChip = new THREE.Mesh(new THREE.BoxGeometry(1.35, 0.14, 0.42), new THREE.MeshStandardMaterial({ color: 0x090d16 }));
      icChip.position.set(-2.9, 0.16, 0.15);
      circuitGroup.add(icChip);

      // Black Pin Headers along MCU right edge
      const headerStrip = new THREE.Mesh(new THREE.BoxGeometry(0.22, 0.24, 1.9), new THREE.MeshStandardMaterial({ color: 0x0f172a }));
      headerStrip.position.set(-1.7, 0.2, 0);
      circuitGroup.add(headerStrip);
      registerLabel(mcuName, -3.2, 0.85, 0);

      // 2. Realistic 830-Point Breadboard with Tie-Point Holes (Right)
      const bb = new THREE.Mesh(
        new THREE.BoxGeometry(5.6, 0.26, 2.8),
        new THREE.MeshStandardMaterial({ color: 0xf8fafc, roughness: 0.45 })
      );
      bb.position.set(2.3, 0, 0);
      circuitGroup.add(bb);

      // Center DIP Channel Groove
      const groove = new THREE.Mesh(new THREE.BoxGeometry(5.4, 0.03, 0.24), new THREE.MeshStandardMaterial({ color: 0xcbd5e1 }));
      groove.position.set(2.3, 0.135, 0);
      circuitGroup.add(groove);

      // Power Rails (+5V Red & -GND Blue)
      const railRed = new THREE.Mesh(new THREE.BoxGeometry(5.2, 0.02, 0.07), new THREE.MeshBasicMaterial({ color: 0xef4444 }));
      railRed.position.set(2.3, 0.14, -1.2);
      const railBlue = new THREE.Mesh(new THREE.BoxGeometry(5.2, 0.02, 0.07), new THREE.MeshBasicMaterial({ color: 0x2563eb }));
      railBlue.position.set(2.3, 0.14, 1.2);
      circuitGroup.add(railRed, railBlue);

      // Tie-point pin holes grid
      const holeMat = new THREE.MeshBasicMaterial({ color: 0x334155 });
      const holeGeo = new THREE.BoxGeometry(0.06, 0.02, 0.06);
      for (let gx = -0.1; gx <= 4.7; gx += 0.32) {
        for (let gz of [-0.85, -0.55, 0.55, 0.85]) {
          const h = new THREE.Mesh(holeGeo, holeMat);
          h.position.set(gx, 0.14, gz);
          circuitGroup.add(h);
        }
      }
      registerLabel("830-Pt Breadboard", 2.3, 0.55, 1.35);

      // Helper for 3D Curved Jumper Wires
      function add3DWire(p1, p2, p3, colorHex) {
        const curve = new THREE.CatmullRomCurve3([
          new THREE.Vector3(...p1),
          new THREE.Vector3(...p2),
          new THREE.Vector3(...p3)
        ]);
        const geo = new THREE.TubeGeometry(curve, 24, 0.042, 8, false);
        const mat = new THREE.MeshStandardMaterial({ color: colorHex });
        circuitGroup.add(new THREE.Mesh(geo, mat));
      }

      // Standard +5V and GND Power Wires
      add3DWire([-1.7, 0.25, -0.75], [-0.4, 1.35, -1.0], [0.2, 0.15, -1.2], 0xef4444);
      add3DWire([-1.7, 0.25, 0.75], [-0.4, 1.25, 1.0], [0.2, 0.15, 1.2], 0x64748b);

      // 3. Dynamic Topic-Specific 3D Components Mounted on Breadboard
      const parts = (bp && bp.circuit_3d_parts) || [
        { type: "sensor_ir", label: "Input Sensor", wire_color: "#38bdf8" },
        { type: "motor_servo", label: "Output Actuator", wire_color: "#f59e0b" },
        { type: "oled_display", label: "I2C OLED Display", wire_color: "#a855f7" }
      ];

      const xSlots = [0.6, 2.3, 4.0];
      parts.slice(0, 3).forEach((part, i) => {
        const px = xSlots[i];
        const pz = i % 2 === 0 ? -0.45 : 0.45;
        const mesh = createPartMesh(part.type);
        mesh.position.set(px, 0.52, pz);
        circuitGroup.add(mesh);

        // Pin legs into breadboard
        const leg = new THREE.Mesh(new THREE.BoxGeometry(0.4, 0.35, 0.08), new THREE.MeshStandardMaterial({ color: 0x94a3b8 }));
        leg.position.set(px, 0.25, pz);
        circuitGroup.add(leg);

        // Signal Jumper Wire from MCU to Component
        const mcuZ = -0.35 + i * 0.35;
        add3DWire([-1.7, 0.25, mcuZ], [(px - 1.7) / 2, 1.7 + i * 0.25, pz], [px, 0.22, pz], part.wire_color || "#38bdf8");

        registerLabel(part.label || "Module", px, 1.25, pz);
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
            item.mesh.rotation.y = Math.sin(simTick) * 0.9;
          } else if (item.mode === "blink") {
            item.mesh.material.color.setHex(Math.sin(simTick * 2.5) > 0 ? 0xef4444 : 0x450a0a);
          }
        });
      }

      // Project 3D positions to 2D floating labels
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
      btn.innerText = simActive ? "⚡ Power Simulation: ON" : "⏸️ Simulation: PAUSED";
      btn.className = simActive
        ? "px-3.5 py-2 rounded-xl bg-emerald-600 font-bold text-white text-xs shadow"
        : "px-3.5 py-2 rounded-xl bg-amber-600 font-bold text-white text-xs shadow";
    }

    function reset3DCamera() {
      if (camera) camera.position.set(0, 8.2, 11.2);
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
    fmt = payload.get("formatting", {})
    docx_buffer = build_vtu_docx(
        blueprint,
        payload.get("student_name") or "Student Name",
        payload.get("usn") or "1VT23CS001",
        payload.get("college") or "VTU Affiliated Engineering College",
        payload.get("branch") or "Engineering",
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