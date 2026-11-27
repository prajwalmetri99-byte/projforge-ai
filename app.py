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
from pptx import Presentation
from pptx.util import Inches as PptInches, Pt as PptPt
from pptx.dml.color import RGBColor as PptRGB

app = Flask(__name__)

# ==================== CONFIGURATION ====================
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "").strip()


# ==================== 1. SMART HARDWARE / SOFTWARE DETECTOR ====================
def _detect_is_hardware(topic, description, branch, subject, explicit_flag=None):
    if isinstance(explicit_flag, bool):
        return explicit_flag
    text = f"{topic} {description} {branch} {subject}".lower()

    hw_keywords = [
        "circuit", "breadboard", "tinkercad", "schematic", "pcb", "solder",
        "ic 741", "ic741", "lm741", "741", "opamp", "op-amp", "555", "ne555", "lm324",
        "arduino", "esp32", "esp8266", "stm32", "raspberry pi", "pic", "8051",
        "sensor", "ultrasonic", "hc-sr04", "ir sensor", "pir", "dht11", "dht22",
        "gas sensor", "mq2", "mq-2", "mq135", "soil moisture", "water level",
        "relay", "servo", "motor", "dc motor", "stepper", "buzzer", "led",
        "oled", "lcd", "iot", "embedded", "vlsi", "verilog", "robotics",
        "smart irrigation", "fire alarm", "burglar", "obstacle", "line follower",
        "probe", "potentiometer", "resistor", "capacitor", "transistor", "bc547"
    ]
    sw_keywords = [
        "filter", "eligibility", "tkinter", "website", "web app", "portal",
        "database", "dbms", "sql", "management system", "erp", "crm",
        "chatbot", "nlp", "machine learning", "deep learning", "data science",
        "csv", "excel", "compiler", "operating system", "algorithm", "software",
        "blockchain", "cybersecurity", "app development", "calculator", "to-do"
    ]

    has_hw = any(k in text for k in hw_keywords)
    has_sw = any(k in text for k in sw_keywords)

    # If user explicitly describes hardware pins/sensors/LEDs/motors/ICs, treat as hardware
    if has_hw:
        return True
    if has_sw:
        return False
    return branch.lower() in {
        "electronics & communication (ece)",
        "electrical & electronics (eee)",
        "mechanical engineering"
    }


def _parse_budget_cap(budget_tier):
    nums = [int(x.replace(",", "")) for x in re.findall(r"\d[\d,]*", budget_tier or "")]
    if not nums:
        return 700
    if len(nums) >= 2:
        return nums[1]
    return nums[0]


# ==================== 2. DYNAMIC BLUEPRINT GENERATOR ====================
def generate_project_blueprint(topic, description, branch, subject, level, budget_tier):
    is_hw_default = _detect_is_hardware(topic, description, branch, subject)
    budget_cap = _parse_budget_cap(budget_tier)

    if GEMINI_API_KEY:
        prompt = f"""
You are a Senior Indian Engineering Professor & System Architect (VTU / Autonomous / AICTE).
Generate a completely customized, intelligent engineering project blueprint in strict JSON based on:
- Project Title: "{topic}"
- Student Custom Description & Requirements: "{description or 'None provided; infer best practical academic design.'}"
- Engineering Branch: "{branch}"
- Academic Subject: "{subject}"
- Academic Level: "{level}"
- Selected Budget Tier: "{budget_tier}" (Target max cost around ₹{budget_cap})

STRICT RULES:
1. "requires_circuit": MUST be true ONLY if the project involves physical electronics, circuits, ICs, sensors, relays, motors, buzzers, LEDs, or microcontrollers. MUST be false for pure software/Python/Tkinter/DBMS/Web/ML/Data/Filter projects.
2. Honor EVERY custom requirement mentioned in Student Custom Description ("{description}") and adapt the title, schematic, 3D parts, code, and viva to "{topic}". Do NOT default to water level or LM741 unless the user's project is actually about that!
3. Keep estimated_total_inr realistic and aligned with "{budget_tier}".
4. If "requires_circuit" is true:
   - Choose 3 hardware modules for "circuit_3d_parts" from: ["ic_741", "ic_555", "sensor_water_probes", "sensor_soil_moisture", "sensor_ultrasonic", "sensor_gas", "sensor_ir", "water_pump", "motor_servo", "relay_module", "oled_display", "piezo_buzzer", "led_indicator_bank"].
   - Provide accurate "schematic_spec" tailored to "{topic}" and "{description}" (supply_voltage, core_ic, input_stage, control_stage, output_stage, pins list).
5. "primary_code": Write 60-90 lines of clean, runnable code specifically for "{topic}".
6. "viva_qa": Provide 5 technical viva Q&As in "en", "kn", "hi", "te", "ta".

Return ONLY valid JSON (no markdown fences):
{{
  "title": "Formal Engineering Project Title",
  "tagline": "Concise summary tailored to topic & description",
  "requires_circuit": {str(is_hw_default).lower()},
  "budget_tier": "{budget_tier}",
  "estimated_total_inr": 200,
  "controller_or_supply": "Tailored controller or supply",
  "verification_summary": {{
    "syllabus_match": "98% Aligned ({subject})",
    "build_time": "1-2 Days",
    "expert_note": "Custom requirements verified and integrated into architecture."
  }},
  "schematic_spec": {{
    "supply_voltage": "+5V DC",
    "core_ic": "Controller / IC Name",
    "input_stage": "Input Sensor / Signal",
    "control_stage": "Logic / Comparator Stage",
    "output_stage": "Output Indicator / Actuator",
    "pins": [
      {{"pin_no": "Pin A", "pin_name": "Function", "connects_to": "Target Pin / Breadboard", "wire_color": "Red"}}
    ]
  }},
  "abstract": "Detailed 200-word abstract tailored to {topic} and {description}.",
  "problem_statement": "Detailed 130-word problem statement.",
  "objectives": ["Objective 1", "Objective 2", "Objective 3", "Objective 4"],
  "hardware_requirements": [
    {{"name": "Item", "spec": "Specification", "cost_inr": 100}}
  ],
  "software_requirements": [
    "OS: Windows 11 / Ubuntu 22.04 LTS",
    "Tools: Python 3.x / Arduino IDE / Simulator"
  ],
  "circuit_3d_parts": [
    {{"type": "sensor_ir", "label": "Module", "wire_color": "#2563eb"}}
  ],
  "software_architecture_nodes": ["Input Layer", "Validation", "Core Logic", "Storage", "Output / Export"],
  "ppt_slides": [
    {{"slide": "Title & Overview", "bullets": ["Project: {topic}", "Subject: {subject} ({branch})", "Budget: {budget_tier}"]}},
    {{"slide": "Problem & Custom Scope", "bullets": ["Addresses core problem statement", "Incorporates custom requirements", "Clear evaluation metrics"]}},
    {{"slide": "System / Circuit Design", "bullets": ["Modular architecture & connections", "Accurate rule/signal processing", "Low-cost reliable components"]}},
    {{"slide": "Implementation & Code", "bullets": ["Clean, documented implementation", "Handles edge cases safely", "Ready for live demonstration"]}},
    {{"slide": "Results & Future Scope", "bullets": ["Verified test performance", "Practical real-world applications", "Scalable future roadmap"]}}
  ],
  "working_principle": "Detailed 240-word working explanation tailored to {topic} and {description}.",
  "primary_code_language": "Python 3 / C++",
  "primary_code": "Complete runnable code.",
  "results_and_discussion": "Detailed 150-word results summary.",
  "advantages": ["Advantage 1", "Advantage 2", "Advantage 3"],
  "applications": ["Application 1", "Application 2", "Application 3"],
  "future_scope": "110-word future scope.",
  "references": [
    "[1] IEEE Standard Reference (2025).",
    "[2] Textbook on {subject}."
  ],
  "viva_qa": [
    {{
      "en": {{"q": "How does your system satisfy the project requirements?", "a": "It implements modular logic and verified components tailored to {topic}."}},
      "kn": {{"q": "ನಿಮ್ಮ ವ್ಯವಸ್ಥೆಯು ಪ್ರಾಜೆಕ್ಟ್ ಅಗತ್ಯಗಳನ್ನು ಹೇಗೆ ಪೂರೈಸುತ್ತದೆ?", "a": "ಇದು {topic} ಗೆ ತಕ್ಕಂತೆ ಮಾಡ್ಯುಲರ್ ವಿನ್ಯಾಸ ಮತ್ತು ನಿಖರ ನಿಯಮಗಳನ್ನು ಬಳಸುತ್ತದೆ."}},
      "hi": {{"q": "आपका सिस्टम प्रोजेक्ट की ज़रूरतों को कैसे पूरा करता है?", "a": "यह {topic} के अनुसार मॉड्यूलर लॉजिक और सत्यापित घटकों का उपयोग करता है।"}},
      "te": {{"q": "మీ సిస్టమ్ ప్రాజెక్ట్ అవసరాలను ఎలా తీరుస్తుంది?", "a": "ఇది {topic} కు అనుగుణంగా మాడ్యులర్ లాజిక్‌ను ఉపయోగిస్తుంది."}},
      "ta": {{"q": "உங்கள் அமைப்பு திட்டத் தேவைகளை எவ்வாறு நிறைவேற்றுகிறது?", "a": "இது {topic} க்கு ஏற்ப வடிவமைக்கப்பட்டு துல்லியமாகச் செயல்படுகிறது."}}
    }}
  ]
}}
"""
        try:
            url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key={GEMINI_API_KEY}"
            payload = {
                "contents": [{"parts": [{"text": prompt}]}],
                "tools": [{"google_search": {}}],
                "generationConfig": {"temperature": 0.35}
            }
            resp = requests.post(url, json=payload, timeout=20)
            if resp.status_code == 200:
                candidates = resp.json().get("candidates", [])
                if candidates:
                    parts_out = candidates[0].get("content", {}).get("parts", [])
                    raw_text = "".join(p.get("text", "") for p in parts_out).strip()
                    match = re.search(r"\{.*\}", raw_text, re.DOTALL)
                    if match:
                        parsed = json.loads(match.group(0))
                        parsed["requires_circuit"] = _detect_is_hardware(
                            topic, description, branch, subject, parsed.get("requires_circuit")
                        )
                        return parsed
        except Exception as e:
            print("Dynamic generator fallback used:", e)

    return _dynamic_fallback_blueprint(topic, description, branch, subject, level, budget_tier)


def _dynamic_fallback_blueprint(topic, description, branch, subject, level, budget_tier):
    clean_title = topic.strip().title() or f"Smart {subject} Project"
    desc_clean = description.strip()
    low = f"{clean_title} {desc_clean} {branch} {subject}".lower()
    requires_circuit = _detect_is_hardware(clean_title, desc_clean, branch, subject)
    budget_cap = _parse_budget_cap(budget_tier)

    custom_note_line = f" Custom Requirement: {desc_clean}." if desc_clean else ""

    if requires_circuit:
        # Dynamically pick hardware modules based on what the user typed
        if any(k in low for k in ["gas", "smoke", "lpg", "mq2", "mq-2", "mq135", "fire"]):
            core_ic = "MQ-2 / MQ-135 Gas Sensor + Comparator / MCU"
            supply_v = "+5V DC"
            est_cost = min(budget_cap, 380 if budget_cap >= 300 else 210)
            parts = [
                {"type": "sensor_gas", "label": "Gas / Smoke Sensor", "wire_color": "#d97706"},
                {"type": "piezo_buzzer", "label": "Alarm Buzzer", "wire_color": "#dc2626"},
                {"type": "led_indicator_bank", "label": "Alert LEDs", "wire_color": "#16a34a"}
            ]
            pins = [
                {"pin_no": "VCC", "pin_name": "Power Supply", "connects_to": "+5V Rail", "wire_color": "Red"},
                {"pin_no": "GND", "pin_name": "Ground", "connects_to": "0V GND Rail", "wire_color": "Black"},
                {"pin_no": "DOUT / AOUT", "pin_name": "Gas Signal", "connects_to": "Comparator / Pin D2/A0", "wire_color": "Green"},
                {"pin_no": "OUT", "pin_name": "Alarm Drive", "connects_to": "Buzzer & Red LED via Resistor", "wire_color": "Red"}
            ]
            hw_req = [
                {"name": "Gas / Smoke Sensor Module", "spec": "Detects combustible gas / smoke", "cost_inr": max(60, est_cost // 3)},
                {"name": "Buzzer, LEDs & Resistors", "spec": "Audio-visual alert stage", "cost_inr": max(40, est_cost // 4)},
                {"name": "Controller / Breadboard / Wires", "spec": "Power & interconnections", "cost_inr": max(50, est_cost - (est_cost // 3) - (est_cost // 4))}
            ]
        elif any(k in low for k in ["ultrasonic", "distance", "radar", "blind", "parking", "obstacle"]):
            core_ic = "HC-SR04 Ultrasonic + Controller"
            supply_v = "+5V DC"
            est_cost = min(budget_cap, 450 if budget_cap >= 300 else 230)
            parts = [
                {"type": "sensor_ultrasonic", "label": "HC-SR04 Ultrasonic", "wire_color": "#2563eb"},
                {"type": "piezo_buzzer", "label": "Proximity Buzzer", "wire_color": "#dc2626"},
                {"type": "led_indicator_bank", "label": "Distance LEDs", "wire_color": "#16a34a"}
            ]
            pins = [
                {"pin_no": "VCC & GND", "pin_name": "Power", "connects_to": "+5V & GND Rails", "wire_color": "Red / Black"},
                {"pin_no": "TRIG", "pin_name": "Trigger Pulse", "connects_to": "Digital Pin D9 (10us pulse)", "wire_color": "Blue"},
                {"pin_no": "ECHO", "pin_name": "Echo Return", "connects_to": "Digital Pin D10 (Pulse width)", "wire_color": "Green"},
                {"pin_no": "ALERT PIN", "pin_name": "Buzzer / LED", "connects_to": "Digital Pin D8 -> Buzzer/LED", "wire_color": "Orange"}
            ]
            hw_req = [
                {"name": "HC-SR04 Ultrasonic Sensor", "spec": "Non-contact distance measurement", "cost_inr": 95},
                {"name": "Buzzer, LEDs & Jumper Wires", "spec": "Alert & breadboard wiring", "cost_inr": max(60, est_cost - 95)}
            ]
        elif any(k in low for k in ["ir", "motion", "pir", "intruder", "counter", "visitor"]):
            core_ic = "IR / PIR Sensor + Logic / MCU"
            supply_v = "+5V DC"
            est_cost = min(budget_cap, 360 if budget_cap >= 300 else 190)
            parts = [
                {"type": "sensor_ir", "label": "IR / PIR Detector", "wire_color": "#2563eb"},
                {"type": "led_indicator_bank", "label": "Indicator LED", "wire_color": "#16a34a"},
                {"type": "piezo_buzzer", "label": "Buzzer / Relay", "wire_color": "#dc2626"}
            ]
            pins = [
                {"pin_no": "VCC", "pin_name": "Supply", "connects_to": "+5V Rail", "wire_color": "Red"},
                {"pin_no": "GND", "pin_name": "Ground", "connects_to": "0V GND Rail", "wire_color": "Black"},
                {"pin_no": "OUT", "pin_name": "Detection Signal", "connects_to": "Input Pin D2 / Logic Stage", "wire_color": "Green"},
                {"pin_no": "OUTPUT", "pin_name": "Indicator / Alarm", "connects_to": "LED & Buzzer", "wire_color": "Blue"}
            ]
            hw_req = [
                {"name": "IR / PIR Sensor Module", "spec": "Object / motion detection", "cost_inr": 80},
                {"name": "LEDs, Buzzer & Breadboard Kit", "spec": "Output & prototyping", "cost_inr": max(60, est_cost - 80)}
            ]
        elif any(k in low for k in ["soil", "irrigation", "plant", "moisture", "pump"]):
            core_ic = "Soil Moisture Sensor + Relay / Pump"
            supply_v = "+5V DC"
            est_cost = min(budget_cap, 520 if budget_cap >= 300 else 240)
            parts = [
                {"type": "sensor_soil_moisture", "label": "Soil Sensor", "wire_color": "#16a34a"},
                {"type": "relay_module", "label": "Relay Switch", "wire_color": "#d97706"},
                {"type": "water_pump", "label": "Mini Water Pump", "wire_color": "#2563eb"}
            ]
            pins = [
                {"pin_no": "VCC & GND", "pin_name": "Power Rail", "connects_to": "+5V & GND", "wire_color": "Red / Black"},
                {"pin_no": "AO / DO", "pin_name": "Moisture Level", "connects_to": "Analog A0 / Comparator", "wire_color": "Green"},
                {"pin_no": "RELAY IN", "pin_name": "Pump Trigger", "connects_to": "Control Pin D7", "wire_color": "Orange"}
            ]
            hw_req = [
                {"name": "Soil Moisture Sensor + Relay", "spec": "Dryness detection & switching", "cost_inr": 140},
                {"name": "Mini Pump & Tubing / Wires", "spec": "Water delivery & connections", "cost_inr": max(80, est_cost - 140)}
            ]
        elif any(k in low for k in ["555", "timer", "blink", "pulse", "pwm", "astable", "monostable"]):
            core_ic = "NE555 Timer IC"
            supply_v = "+9V DC"
            est_cost = min(budget_cap, 180)
            parts = [
                {"type": "ic_555", "label": "NE555 Timer", "wire_color": "#2563eb"},
                {"type": "led_indicator_bank", "label": "Pulse LEDs", "wire_color": "#16a34a"},
                {"type": "piezo_buzzer", "label": "Tone Buzzer", "wire_color": "#dc2626"}
            ]
            pins = [
                {"pin_no": "Pin 8 & Pin 4", "pin_name": "VCC & Reset", "connects_to": "+9V Supply Rail", "wire_color": "Red"},
                {"pin_no": "Pin 1", "pin_name": "GND", "connects_to": "0V GND Rail", "wire_color": "Black"},
                {"pin_no": "Pin 3", "pin_name": "OUT", "connects_to": "Resistor -> LED / Buzzer", "wire_color": "Blue"}
            ]
            hw_req = [
                {"name": "NE555 Timer IC + Resistors/Capacitor", "spec": "Timing & pulse generation", "cost_inr": 65},
                {"name": "LED, Buzzer, 9V Battery & Board", "spec": "Output & power supply", "cost_inr": max(65, est_cost - 65)}
            ]
        else:
            # General hardware circuit tailored to whatever hardware topic/description user typed
            core_ic = "Sensor + Controller / Comparator"
            supply_v = "+5V DC"
            est_cost = min(budget_cap, 420 if budget_cap >= 300 else 210)
            parts = [
                {"type": "sensor_ir", "label": "Input Stage", "wire_color": "#2563eb"},
                {"type": "led_indicator_bank", "label": "Status LEDs", "wire_color": "#16a34a"},
                {"type": "piezo_buzzer", "label": "Output / Alarm", "wire_color": "#dc2626"}
            ]
            pins = [
                {"pin_no": "VCC", "pin_name": "Positive Supply", "connects_to": supply_v + " Rail", "wire_color": "Red"},
                {"pin_no": "GND", "pin_name": "Ground", "connects_to": "0V Common Ground", "wire_color": "Black"},
                {"pin_no": "INPUT", "pin_name": "Signal Sense", "connects_to": "Input Module / Pin", "wire_color": "Green"},
                {"pin_no": "OUTPUT", "pin_name": "Actuator / LED", "connects_to": "Output Stage via Resistor", "wire_color": "Blue"}
            ]
            hw_req = [
                {"name": f"Core Hardware for {clean_title}", "spec": desc_clean or "Primary sensing & control", "cost_inr": max(90, est_cost // 2)},
                {"name": "Breadboard, Wires, LEDs & Buzzer", "spec": "Assembly & indicators", "cost_inr": max(60, est_cost - max(90, est_cost // 2))}
            ]

        schem = {
            "supply_voltage": supply_v,
            "core_ic": core_ic,
            "input_stage": f"Input for {clean_title}",
            "control_stage": core_ic,
            "output_stage": "LED / Buzzer / Load Output",
            "pins": pins
        }
        prim_lang = "Python 3 (Simulation & Logic Model) / Embedded C"
        prim_code = f"""# {clean_title} — Custom Logic & Simulation ({subject})
# Description: {desc_clean or 'Standard evaluation model'}

def evaluate_system(sensor_value, threshold=50):
    triggered = sensor_value >= threshold
    return {{
        "project": "{clean_title}",
        "sensor_value": sensor_value,
        "threshold": threshold,
        "output_state": "ACTIVE" if triggered else "NORMAL"
    }}

if __name__ == "__main__":
    for reading in [20, 48, 65, 90]:
        print(evaluate_system(reading))
"""
    else:
        # Pure Software Project (no hardware circuit)
        est_cost = 0
        parts = []
        schem = None
        hw_req = [{"name": "Standard PC / Laptop", "spec": "Python 3.x / Browser Runtime", "cost_inr": 0}]
        prim_lang = f"Python 3 ({subject})"
        prim_code = f"""# {clean_title} ({subject})
# Description: {desc_clean or 'Structured software workflow'}

class ProjectEngine:
    def __init__(self, rules=None):
        self.rules = rules or {{}}

    def execute(self, records):
        results = []
        for rec in records:
            if self._matches(rec):
                results.append(rec)
        return results

    def _matches(self, rec):
        min_val = self.rules.get("min_score", 0)
        return float(rec.get("score", 0)) >= min_val

if __name__ == "__main__":
    engine = ProjectEngine({{"min_score": 7.5}})
    sample = [{{"id": 1, "name": "Item A", "score": 8.4}}, {{"id": 2, "name": "Item B", "score": 6.2}}]
    print("Matched:", engine.execute(sample))
"""

    return {
        "title": clean_title,
        "tagline": f"{subject} • {budget_tier}" + (f" • {desc_clean[:70]}" if desc_clean else ""),
        "requires_circuit": requires_circuit,
        "budget_tier": budget_tier,
        "estimated_total_inr": est_cost,
        "controller_or_supply": core_ic if requires_circuit else "Software Architecture",
        "verification_summary": {
            "syllabus_match": f"98% Aligned ({subject})",
            "build_time": "1-2 Days",
            "expert_note": f"Dynamically tailored to '{clean_title}' within {budget_tier}.{custom_note_line}"
        },
        "schematic_spec": schem,
        "abstract": (
            f"This project presents '{clean_title}' developed for {subject} ({branch}, {level}) under budget tier "
            f"'{budget_tier}' (Estimated cost: ₹{est_cost}).{custom_note_line} The solution provides structured logic, "
            "clear visual documentation, runnable source code, and complete presentation and report deliverables."
        ),
        "problem_statement": (
            f"Existing manual approaches to '{clean_title}' lack automation, accuracy, or cost-effectiveness.{custom_note_line} "
            f"A reliable {subject} implementation is needed to streamline execution and verification."
        ),
        "objectives": [
            f"Implement '{clean_title}' tailored to {subject} requirements.",
            "Support custom project requirements and configurable thresholds/rules.",
            "Provide visual circuit/architecture diagrams, simulation, and clean source code.",
            "Generate ready-to-submit academic report and seminar slides."
        ],
        "hardware_requirements": hw_req,
        "software_requirements": [
            "OS: Windows 11 / Ubuntu 22.04 LTS",
            f"Tools: Python 3.x / Lab & Simulation Tools ({subject})"
        ],
        "circuit_3d_parts": parts,
        "software_architecture_nodes": [
            "User Input & Config",
            "Validation & Parsing",
            f"{subject} Core Engine",
            "Filtering & Evaluation",
            "Output & Export"
        ],
        "ppt_slides": [
            {"slide": "Title & Overview", "bullets": [f"Project: {clean_title}", f"Subject: {subject} ({branch})", f"Budget: {budget_tier}"]},
            {"slide": "Problem & Requirements", "bullets": ["Limitations of manual approach", desc_clean or "Custom project objectives", "Accuracy & reliability goals"]},
            {"slide": "Design & Architecture", "bullets": ["Modular input -> logic -> output flow", "Verified connections / rules", "Clean error handling"]},
            {"slide": "Implementation & Results", "bullets": ["Tested across multiple scenarios", "Clear, fast & verifiable output", "Ready for demo & viva"]},
            {"slide": "Applications & Future Scope", "bullets": ["Practical academic & real-world use", "Future automation & scaling", "Conclusion"]}
        ],
        "working_principle": (
            f"'{clean_title}' processes inputs according to {subject} principles, evaluates defined conditions or thresholds, "
            f"and generates accurate outputs.{custom_note_line}"
        ),
        "primary_code_language": prim_lang,
        "primary_code": prim_code,
        "results_and_discussion": f"Verified across test cases for '{clean_title}' within budget ₹{est_cost}.",
        "advantages": [
            f"Custom-tailored to '{clean_title}' and your description",
            f"Budget-aware design (Est. ₹{est_cost})",
            "Includes diagram/3D view, code, 5-language viva, report & PPT"
        ],
        "applications": [
            f"{subject} Academic & Lab Projects",
            "Department & Institutional Automation",
            "Prototyping & Demonstration"
        ],
        "future_scope": "Can be expanded with automated reporting, cloud sync, and additional custom rules.",
        "references": [
            f"[1] IEEE Standard Reference on {subject} (2025).",
            f"[2] Standard Textbook on {subject}."
        ],
        "viva_qa": [
            {
                "en": {"q": f"What makes your implementation of '{clean_title}' effective?", "a": f"It combines {subject} logic with custom requirements and clear verification steps."},
                "kn": {"q": f"'{clean_title}' ಪ್ರಾಜೆಕ್ಟ್ ಹೇಗೆ ಪರಿಣಾಮಕಾರಿಯಾಗಿ ಕೆಲಸ ಮಾಡುತ್ತದೆ?", "a": f"ಇದು {subject} ತತ್ವಗಳನ್ನು ಮತ್ತು ನಿಮ್ಮ ಅಗತ್ಯಗಳನ್ನು ಸರಿಯಾಗಿ ಒಗ್ಗೂಡಿಸಿ ನಿಖರ ಫಲಿತಾಂಶ ನೀಡುತ್ತದೆ."},
                "hi": {"q": f"'{clean_title}' का आपका समाधान कैसे प्रभावी है?", "a": f"यह {subject} सिद्धांतों और कस्टम आवश्यकताओं को मिलाकर सटीक परिणाम देता है।"},
                "te": {"q": f"'{clean_title}' ప్రాజెక్ట్ ఎలా సమర్థవంతంగా పనిచేస్తుంది?", "a": f"ఇది {subject} సూత్రాలు మరియు అవసరాలకు అనుగుణంగా కచ్చితమైన ఫలితం ఇస్తుంది."},
                "ta": {"q": f"'{clean_title}' திட்டம் எவ்வாறு சிறப்பாகச் செயல்படுகிறது?", "a": f"இது {subject} கோட்பாடுகள் மற்றும் தேவைகளுடன் துல்லியமான வெளியீட்டைத் தருகிறது."}
            }
        ]
    }


# ==================== 3. DOCX & PPTX BUILDERS ====================
def build_vtu_docx(blueprint, student_name, usn, college, branch, subject, guide_name, viva_lang="en"):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.1)
        section.right_margin = Inches(1.0)

    def add_h(text):
        p = doc.add_heading(text, level=1)
        for r in p.runs:
            r.font.name = "Times New Roman"
            r.font.size = Pt(14)
            r.bold = True
            r.font.color.rgb = RGBColor(20, 45, 110)
        return p

    def add_b(text, bold=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.line_spacing = 1.5
        p.paragraph_format.space_after = Pt(6)
        r = p.add_run(text)
        r.font.name = "Times New Roman"
        r.font.size = Pt(12)
        r.bold = bold
        return p

    title = blueprint.get("title", "Engineering Project")
    year_str = datetime.now().strftime("%Y")

    for txt, sz, bld, clr in [
        (f"{college.upper()}\nDepartment of {branch}\n\n", 14, True, RGBColor(15, 35, 95)),
        (f"PROJECT REPORT ON\n({subject.upper()})\n\n", 12, True, None),
        (f'"{title.upper()}"\n\n', 14, True, RGBColor(180, 30, 30)),
        (f"Submitted By:\n{student_name.upper()} (USN: {usn.upper()})\n\n", 12, True, None),
        (f"Under the Guidance of:\n{guide_name}\nAcademic Year: {year_str}", 12, False, None),
    ]:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(txt)
        r.font.name = "Times New Roman"
        r.font.size = Pt(sz)
        r.bold = bld
        if clr:
            r.font.color.rgb = clr
    doc.add_page_break()

    add_h("ABSTRACT").alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_b(blueprint.get("abstract", ""))
    add_h("1. INTRODUCTION & PROBLEM STATEMENT")
    add_b(blueprint.get("problem_statement", ""))
    for obj in blueprint.get("objectives", []):
        add_b("• " + obj, align=WD_ALIGN_PARAGRAPH.LEFT)

    add_h("2. WORKING PRINCIPLE & DESIGN")
    add_b(blueprint.get("working_principle", ""))

    add_h("3. BUDGET & REQUIREMENTS")
    for hw in blueprint.get("hardware_requirements", []):
        add_b(f"• {hw.get('name', '')} ({hw.get('spec', '')}) — ₹{hw.get('cost_inr', 0)}", align=WD_ALIGN_PARAGRAPH.LEFT)
    add_b(f"Estimated Total Cost: ₹{blueprint.get('estimated_total_inr', 0)}", bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)

    add_h("4. SOURCE CODE & VIVA Q&A")
    add_b(blueprint.get("primary_code", ""), align=WD_ALIGN_PARAGRAPH.LEFT)
    for idx, qa in enumerate(blueprint.get("viva_qa", []), start=1):
        entry = qa.get(viva_lang) or qa.get("en") or {}
        add_b(f"Q{idx}: {entry.get('q', '')}", bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)
        add_b(f"Ans: {entry.get('a', '')}", align=WD_ALIGN_PARAGRAPH.LEFT)

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
    p0.text = f"{college.upper()} • {branch.upper()}"
    p0.font.name = "Times New Roman"
    p0.font.size = PptPt(14)
    p0.font.color.rgb = PptRGB(56, 189, 248)
    p0.font.bold = True

    p1 = tf.add_paragraph()
    p1.text = "\n" + blueprint.get("title", "Project Presentation")
    p1.font.name = "Times New Roman"
    p1.font.size = PptPt(26)
    p1.font.color.rgb = PptRGB(255, 255, 255)
    p1.font.bold = True

    p2 = tf.add_paragraph()
    p2.text = f"\nSubject: {subject}  |  By: {student_name} ({usn})  |  Guide: {guide_name}"
    p2.font.name = "Times New Roman"
    p2.font.size = PptPt(14)
    p2.font.color.rgb = PptRGB(203, 213, 225)

    for idx, slide_data in enumerate(blueprint.get("ppt_slides", []), start=2):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        header = slide.shapes.add_shape(1, 0, 0, prs.slide_width, PptInches(1.1))
        header.fill.solid()
        header.fill.fore_color.rgb = PptRGB(30, 58, 138)
        header.line.fill.background()

        htb = slide.shapes.add_textbox(PptInches(0.8), PptInches(0.3), PptInches(11.5), PptInches(0.6))
        hp = htb.text_frame.paragraphs[0]
        hp.text = f"0{idx}. {slide_data.get('slide', 'Overview')}"
        hp.font.name = "Times New Roman"
        hp.font.size = PptPt(20)
        hp.font.bold = True
        hp.font.color.rgb = PptRGB(255, 255, 255)

        btb = slide.shapes.add_textbox(PptInches(1.0), PptInches(1.6), PptInches(11.0), PptInches(4.8))
        btf = btb.text_frame
        btf.word_wrap = True
        for b_i, bullet in enumerate(slide_data.get("bullets", [])):
            bp = btf.paragraphs[0] if b_i == 0 else btf.add_paragraph()
            bp.text = f"▸  {bullet}"
            bp.font.name = "Times New Roman"
            bp.font.size = PptPt(16)
            bp.font.color.rgb = PptRGB(30, 41, 59)
            bp.space_after = PptPt(14)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf


# ==================== 4. FRONTEND HTML ====================
HTML_PAGE = r"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
  <meta name="theme-color" content="#1e3a8a" />
  <title>ProjForge — Custom Project, 3D Circuit, Diagram &amp; Report Studio</title>
  <link rel="icon" type="image/svg+xml" href="/favicon.ico" />
  <script src="https://cdn.tailwindcss.com"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/three.js/r128/three.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/three@0.128.0/examples/js/controls/OrbitControls.js"></script>
  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@500;600;700;800&family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;600&display=swap" rel="stylesheet">
  <style>
    * { box-sizing: border-box; }
    html, body { max-width: 100%; overflow-x: hidden; font-family: 'Inter', sans-serif; background: #f8fafc; color: #0f172a; }
    h1, h2, h3, .brand-font { font-family: 'Plus Jakarta Sans', sans-serif; }
    code, pre { font-family: 'JetBrains Mono', monospace; }
    .card { background: #ffffff; border: 1px solid #e2e8f0; box-shadow: 0 4px 20px -4px rgba(15, 23, 42, 0.06); }
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

  <header class="sticky top-0 z-40 bg-white/95 backdrop-blur-md border-b border-slate-200">
    <div class="max-w-6xl mx-auto px-4 py-3 flex items-center justify-between gap-2">
      <div class="flex items-center gap-2.5 cursor-pointer" onclick="navigatePage('page-home')">
        <svg width="36" height="36" viewBox="0 0 88 88" fill="none" class="shrink-0 rounded-xl shadow-sm">
          <rect width="88" height="88" rx="22" fill="#1e3a8a"/>
          <path d="M24 44H36L42 29L49 59L55 44H65" stroke="#38bdf8" stroke-width="5.5" stroke-linecap="round" stroke-linejoin="round"/>
        </svg>
        <div>
          <span class="brand-font font-extrabold text-base sm:text-lg text-slate-900">ProjForge</span>
          <span class="hidden sm:inline-block ml-1.5 text-[10px] px-2 py-0.5 rounded-full bg-blue-50 text-blue-900 border border-blue-200 font-bold">v12.0 DYNAMIC</span>
        </div>
      </div>

      <nav class="hidden md:flex items-center gap-1 bg-slate-100 p-1 rounded-xl border border-slate-200">
        <button type="button" onclick="navigatePage('page-home')" id="nav-d-home" class="nav-d-btn px-3 py-1.5 rounded-lg text-xs font-bold bg-blue-900 text-white">🏠 Create &amp; Budget</button>
        <button type="button" onclick="navigatePage('page-3d')" id="nav-d-3d" class="nav-d-btn px-3 py-1.5 rounded-lg text-xs font-semibold text-slate-600 hover:text-slate-900">⚡ 3D &amp; Circuit Diagram</button>
        <button type="button" onclick="navigatePage('page-code')" id="nav-d-code" class="nav-d-btn px-3 py-1.5 rounded-lg text-xs font-semibold text-slate-600 hover:text-slate-900">💻 Code &amp; Viva</button>
        <button type="button" onclick="navigatePage('page-report')" id="nav-d-report" class="nav-d-btn px-3 py-1.5 rounded-lg text-xs font-semibold text-slate-600 hover:text-slate-900">📄 Report &amp; PPTX</button>
      </nav>
    </div>
  </header>

  <main class="max-w-6xl mx-auto px-4 pt-6">

    <!-- ==================== PAGE 1: CREATE ==================== -->
    <section id="page-home" class="page-section active space-y-5">
      <div class="text-center max-w-2xl mx-auto">
        <span class="text-[11px] uppercase tracking-wider px-3 py-1 rounded-full bg-blue-50 text-blue-900 border border-blue-200 font-bold">
          Any Topic • Custom Description • Budget-Aware • Verified Output
        </span>
        <h1 class="text-2xl sm:text-4xl font-extrabold mt-2.5 leading-tight text-slate-900">
          Enter Your Project &amp; How You Want It Built
        </h1>
        <p class="text-slate-600 text-xs sm:text-sm mt-2">
          Type any project title and optionally describe your exact requirements, budget, sensors/features, or logic—everything adapts to your project.
        </p>
      </div>

      <!-- QUICK SAMPLE TESTERS -->
      <div class="max-w-3xl mx-auto flex flex-wrap items-center justify-center gap-2 text-xs">
        <span class="text-slate-500 font-semibold">Quick Try:</span>
        <button type="button" onclick="fillPreset('Automatic Gas Leakage Detector & Buzzer Alert', 'Trigger loud buzzer and red LED when gas exceeds safe limit, low cost circuit', '₹100 - ₹250 (Ultra Low-Cost Analog / IC Circuit)')"
          class="px-2.5 py-1 rounded-lg bg-white border border-slate-300 hover:bg-slate-100 text-slate-700 font-medium">🔥 Gas Alarm (Hardware)</button>
        <button type="button" onclick="fillPreset('Smart Ultrasonic Distance Alert System', 'Measure obstacle distance and blink buzzer faster as object comes closer', '₹300 - ₹700 (Student Mini-Project)')"
          class="px-2.5 py-1 rounded-lg bg-white border border-slate-300 hover:bg-slate-100 text-slate-700 font-medium">📡 Distance Alert (Hardware)</button>
        <button type="button" onclick="fillPreset('Placement Eligibility Filter & Certificate App', 'Filter students by CGPA, backlogs, skills and certificates, export Excel & print certificate', '₹0 (Pure Software / Python / Web Project)')"
          class="px-2.5 py-1 rounded-lg bg-white border border-slate-300 hover:bg-slate-100 text-slate-700 font-medium">💻 Eligibility Filter (Software)</button>
      </div>

      <div class="card rounded-2xl p-5 sm:p-7 max-w-3xl mx-auto space-y-4">
        <div>
          <label class="text-xs font-bold text-slate-700">1. Project Title / Name *</label>
          <input id="topic-input" type="text"
            placeholder="Enter your project title (any hardware or software project)..."
            class="w-full mt-1.5 px-4 py-3 rounded-xl bg-slate-50 border border-slate-300 text-slate-900 text-sm font-medium focus:border-blue-800 focus:outline-none" />
        </div>

        <div>
          <label class="text-xs font-bold text-slate-700">2. Project Description / Custom Requirements (Optional — Write how you want it)</label>
          <textarea id="desc-input" rows="3"
            placeholder="Optional: Write specific features, sensors, outputs, budget notes, or how your project should work..."
            class="w-full mt-1.5 px-4 py-2.5 rounded-xl bg-slate-50 border border-slate-300 text-slate-900 text-xs sm:text-sm focus:border-blue-800 focus:outline-none"></textarea>
        </div>

        <div class="grid grid-cols-1 sm:grid-cols-2 gap-3.5">
          <div>
            <label class="text-xs font-bold text-slate-700">3. Budget Range (INR ₹)</label>
            <select id="budget-input"
              class="w-full mt-1 px-3.5 py-2.5 rounded-xl bg-emerald-50/70 border border-emerald-400 text-slate-900 text-xs sm:text-sm font-bold">
              <option value="₹100 - ₹250 (Ultra Low-Cost Analog / IC Circuit)">₹100 – ₹250 (Low-Cost Circuit)</option>
              <option value="₹300 - ₹700 (Student Mini-Project)">₹300 – ₹700 (Mini-Project)</option>
              <option value="₹700 - ₹1,500+ (Advanced / IoT)">₹700 – ₹1,500+ (Advanced / IoT)</option>
              <option value="₹0 (Pure Software / Python / Web Project)">₹0 (Pure Software)</option>
            </select>
          </div>
          <div>
            <label class="text-xs font-bold text-slate-700">4. Engineering Branch</label>
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
        </div>

        <div class="grid grid-cols-1 sm:grid-cols-2 gap-3.5">
          <div>
            <label class="text-xs font-bold text-slate-700">5. Academic Subject</label>
            <select id="subject-input" onchange="handleSubjectSelect()"
              class="w-full mt-1 px-3 py-2.5 rounded-xl bg-slate-50 border border-slate-300 text-xs sm:text-sm font-medium">
            </select>
          </div>
          <div>
            <label class="text-xs font-semibold text-slate-600">Guide / Faculty Name</label>
            <input id="guide-input" type="text" placeholder="Dr. Shivakumar G"
              class="w-full mt-1 px-3 py-2.5 rounded-xl bg-slate-50 border border-slate-300 text-xs sm:text-sm" />
          </div>
        </div>

        <div id="manual-subject-box" class="hidden">
          <label class="text-xs font-bold text-amber-800">✍️ Type Subject Manually *</label>
          <input id="manual-subject-input" type="text" placeholder="e.g., Python Programming (1BEC305)"
            class="w-full mt-1 px-3.5 py-2.5 rounded-xl bg-amber-50/50 border border-amber-400 text-sm" />
        </div>

        <button type="button" id="generate-btn" onclick="generateProject()"
          class="w-full py-3.5 rounded-xl bg-blue-900 hover:bg-blue-800 font-bold text-white text-sm sm:text-base shadow-md transition">
          Generate Tailored Project
        </button>
      </div>
    </section>

    <!-- ==================== PAGE 2: 3D + CIRCUIT DIAGRAM / ARCHITECTURE ==================== -->
    <section id="page-3d" class="page-section space-y-5">
      <div class="grid grid-cols-1 lg:grid-cols-12 gap-5">

        <aside class="lg:col-span-3 space-y-3">
          <div class="card rounded-2xl p-3 space-y-1.5">
            <p class="text-[11px] font-bold uppercase tracking-wider text-slate-400 px-2 pb-1">Navigation</p>
            <button type="button" onclick="navigatePage('page-home')" class="w-full text-left px-3 py-2.5 rounded-xl text-xs font-semibold text-slate-700 hover:bg-slate-100 flex justify-between">
              <span>🏠 Create &amp; Budget</span><span>→</span>
            </button>
            <button type="button" onclick="navigatePage('page-3d')" class="w-full text-left px-3 py-2.5 rounded-xl text-xs font-bold bg-blue-900 text-white flex justify-between">
              <span>⚡ Diagram &amp; 3D</span><span>●</span>
            </button>
            <button type="button" onclick="navigatePage('page-code')" class="w-full text-left px-3 py-2.5 rounded-xl text-xs font-semibold text-slate-700 hover:bg-slate-100 flex justify-between">
              <span>💻 Code &amp; Viva</span><span>→</span>
            </button>
            <button type="button" onclick="navigatePage('page-report')" class="w-full text-left px-3 py-2.5 rounded-xl text-xs font-semibold text-slate-700 hover:bg-slate-100 flex justify-between">
              <span>📄 Report &amp; PPTX</span><span>→</span>
            </button>
          </div>

          <div class="card rounded-2xl p-4 border-l-4 border-l-emerald-600 space-y-2">
            <div class="flex items-center justify-between">
              <span class="text-[10px] uppercase font-bold text-emerald-800 bg-emerald-50 px-2 py-0.5 rounded">Estimated Budget</span>
              <button type="button" onclick="copyBomList()" class="text-[10px] font-bold text-blue-800 underline">Copy BOM</button>
            </div>
            <p id="est-cost-display" class="text-xl font-extrabold text-slate-900">Est. ₹0</p>
            <div id="sidebar-bom-list" class="text-[11px] text-slate-600 space-y-1 border-t border-slate-100 pt-2"></div>
          </div>
        </aside>

        <div class="lg:col-span-9 space-y-5">
          <div id="circuit-lab-wrapper" class="space-y-5">
            <div class="card rounded-2xl p-5 space-y-3">
              <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                <div>
                  <span class="text-[11px] uppercase tracking-wider text-blue-800 font-bold">Dynamic Schematic Diagram</span>
                  <h3 id="schem-heading" class="text-base font-bold text-slate-900">📐 Circuit Diagram for Your Project</h3>
                </div>
                <button type="button" onclick="printCircuitDiagram()" class="px-3 py-1.5 rounded-lg bg-slate-100 hover:bg-slate-200 text-xs font-semibold border border-slate-300 self-start">
                  🖨️ Print Diagram
                </button>
              </div>
              <div id="schematic-svg-container" class="w-full overflow-x-auto bg-white border border-slate-300 rounded-xl p-2"></div>
            </div>

            <div class="card rounded-2xl p-4 space-y-3">
              <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                <div>
                  <span class="text-[11px] uppercase tracking-wider text-blue-800 font-bold">Interactive 3D Breadboard View</span>
                  <h2 id="lab3d-title" class="text-base sm:text-lg font-bold text-slate-900">3D Simulation</h2>
                </div>
                <div class="flex gap-2">
                  <button type="button" onclick="toggle3DSimulation()" id="sim-toggle-btn" class="px-3 py-1.5 rounded-xl bg-emerald-600 font-bold text-white text-xs">▶ Sim: ON</button>
                  <button type="button" onclick="reset3DCamera()" class="px-3 py-1.5 rounded-xl bg-slate-100 text-xs font-semibold border border-slate-300">🔄 Reset</button>
                </div>
              </div>
              <div id="threejs-wrapper" class="relative w-full h-[340px] sm:h-[400px] rounded-xl bg-[#dce5f0] overflow-hidden border border-slate-300">
                <div id="threejs-container" class="w-full h-full"></div>
                <div id="labels-overlay" class="absolute inset-0 pointer-events-none overflow-hidden"></div>
              </div>
              <div class="bg-slate-50 rounded-xl p-3 border border-slate-200 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                <span class="text-xs font-bold text-slate-800">🎛️ Input Level: <strong id="sensor-slider-val">65%</strong></span>
                <input id="sensor-stimulus-slider" type="range" min="0" max="100" value="65" oninput="updateSensorStimulus(this.value)" class="w-full sm:w-64 accent-blue-900" />
              </div>
            </div>

            <div class="card rounded-2xl p-5">
              <h3 class="text-sm font-bold text-slate-900 mb-3">🔌 Pin &amp; Wire Connection Guide</h3>
              <div class="overflow-x-auto">
                <table class="w-full text-xs text-left border border-slate-200">
                  <thead class="bg-slate-100 text-slate-800">
                    <tr>
                      <th class="p-2.5 border-b">Pin / Terminal</th>
                      <th class="p-2.5 border-b">Role</th>
                      <th class="p-2.5 border-b">Connects To</th>
                      <th class="p-2.5 border-b">Wire</th>
                    </tr>
                  </thead>
                  <tbody id="lab3d-pins" class="divide-y divide-slate-200 text-slate-700"></tbody>
                </table>
              </div>
            </div>
          </div>

          <div id="software-arch-wrapper" class="hidden card rounded-2xl p-6 space-y-4">
            <h2 id="sw-arch-title" class="text-lg font-bold text-slate-900">Software System Architecture</h2>
            <p class="text-xs text-slate-500">Software project detected — hardware circuit diagram is not shown.</p>
            <div id="sw-nodes-container" class="grid grid-cols-1 sm:grid-cols-5 gap-3"></div>
          </div>
        </div>

      </div>
    </section>

    <!-- ==================== PAGE 3: CODE & VIVA ==================== -->
    <section id="page-code" class="page-section space-y-5">
      <div class="card rounded-2xl p-5 space-y-3">
        <div class="flex justify-between items-center">
          <p id="active-code-subtitle" class="text-xs text-blue-900 font-bold">Source Code</p>
          <button type="button" onclick="copySourceCode()" class="px-3 py-1.5 rounded-lg bg-slate-100 text-xs font-semibold border border-slate-300">📋 Copy Code</button>
        </div>
        <pre id="code-block" class="bg-slate-900 p-4 rounded-xl text-xs text-emerald-300 overflow-x-auto leading-relaxed">Generate a project first to view code.</pre>
      </div>

      <div class="card rounded-2xl p-5 space-y-4">
        <div class="flex flex-wrap items-center justify-between gap-2">
          <h3 class="text-base font-bold text-slate-900">🎓 Viva Practice Q&amp;A</h3>
          <div class="flex flex-wrap gap-1.5">
            <button type="button" onclick="selectVivaLang('en', this)" class="viva-lang-btn px-3 py-1.5 rounded-lg bg-blue-900 text-white text-xs font-bold">English</button>
            <button type="button" onclick="selectVivaLang('kn', this)" class="viva-lang-btn px-3 py-1.5 rounded-lg bg-slate-100 text-slate-700 text-xs font-semibold border">ಕನ್ನಡ</button>
            <button type="button" onclick="selectVivaLang('hi', this)" class="viva-lang-btn px-3 py-1.5 rounded-lg bg-slate-100 text-slate-700 text-xs font-semibold border">हिन्दी</button>
            <button type="button" onclick="selectVivaLang('te', this)" class="viva-lang-btn px-3 py-1.5 rounded-lg bg-slate-100 text-slate-700 text-xs font-semibold border">తెలుగు</button>
            <button type="button" onclick="selectVivaLang('ta', this)" class="viva-lang-btn px-3 py-1.5 rounded-lg bg-slate-100 text-slate-700 text-xs font-semibold border">தமிழ்</button>
          </div>
        </div>
        <div id="viva-grid" class="grid grid-cols-1 md:grid-cols-2 gap-3.5"></div>
      </div>
    </section>

    <!-- ==================== PAGE 4: REPORT & PPTX ==================== -->
    <section id="page-report" class="page-section space-y-5">
      <div class="card rounded-2xl p-5 space-y-4">
        <h2 id="rep-title" class="text-lg font-bold text-slate-900">Download Project Deliverables</h2>
        <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <button type="button" onclick="downloadFile('/api/download-docx', 'ProjForge_Report.docx', 'dl-docx')" id="dl-docx" class="py-3 px-4 rounded-xl bg-blue-900 font-bold text-white text-xs">📄 Download Report (.docx)</button>
          <button type="button" onclick="downloadFile('/api/download-pptx', 'ProjForge_Seminar.pptx', 'dl-pptx')" id="dl-pptx" class="py-3 px-4 rounded-xl bg-emerald-700 font-bold text-white text-xs">📊 Download Seminar PPT (.pptx)</button>
        </div>
      </div>
      <div class="card rounded-2xl p-6 space-y-3">
        <h3 class="text-base font-bold text-blue-900">Abstract &amp; Working</h3>
        <p id="rep-abstract" class="text-sm text-slate-700 leading-relaxed"></p>
        <p id="rep-working" class="text-sm text-slate-600 leading-relaxed border-t pt-3"></p>
      </div>
    </section>

  </main>

  <nav class="md:hidden fixed bottom-0 inset-x-0 z-40 bg-white border-t border-slate-200 grid grid-cols-4 py-1.5 px-1 shadow-lg">
    <button type="button" onclick="navigatePage('page-home')" class="flex flex-col items-center py-1 text-blue-900 font-bold"><span class="text-base">🏠</span><span class="text-[10px]">Create</span></button>
    <button type="button" onclick="navigatePage('page-3d')" class="flex flex-col items-center py-1 text-slate-500"><span class="text-base">⚡</span><span class="text-[10px]">Diagram</span></button>
    <button type="button" onclick="navigatePage('page-code')" class="flex flex-col items-center py-1 text-slate-500"><span class="text-base">💻</span><span class="text-[10px]">Code</span></button>
    <button type="button" onclick="navigatePage('page-report')" class="flex flex-col items-center py-1 text-slate-500"><span class="text-base">📄</span><span class="text-[10px]">Report</span></button>
  </nav>

  <script>
    const BRANCH_SUBJECTS = {
      "Electronics & Communication (ECE)": ["Analog & Linear Circuits", "Embedded Systems & IoT", "Sensors & Instrumentation", "Digital System Design"],
      "Computer Science & Engineering (CSE)": ["Python Programming", "Database Management Systems", "Artificial Intelligence & ML", "Web & Software Engineering"],
      "Artificial Intelligence & ML (AIML)": ["Python & Machine Learning", "Deep Learning", "Data Analytics"],
      "Electrical & Electronics (EEE)": ["Analog & Digital Circuits", "Power Electronics", "Industrial Automation"],
      "Mechanical Engineering": ["Mechatronics & Automation", "Robotics"],
      "Civil Engineering": ["Smart Monitoring Systems"]
    };

    function updateSubjectsForBranch() {
      const branch = document.getElementById("branch-input").value;
      const subSelect = document.getElementById("subject-input");
      subSelect.innerHTML = (BRANCH_SUBJECTS[branch] || []).map(s => `<option value="${s}">${s}</option>`).join("")
        + `<option value="__MANUAL__">✍️ Other (Type Subject Manually...)</option>`;
      handleSubjectSelect();
    }
    function handleSubjectSelect() {
      const box = document.getElementById("manual-subject-box");
      if (document.getElementById("subject-input").value === "__MANUAL__") box.classList.remove("hidden");
      else box.classList.add("hidden");
    }
    function fillPreset(title, desc, budget) {
      document.getElementById("topic-input").value = title;
      document.getElementById("desc-input").value = desc;
      document.getElementById("budget-input").value = budget;
    }

    let currentBlueprint = null;
    let activeVivaLang = "en";
    let sensorStimulusLevel = 0.65;

    setInterval(() => fetch("/health").catch(() => {}), 240000);

    window.addEventListener("DOMContentLoaded", () => {
      updateSubjectsForBranch();
      try {
        initThreeJS();
        build3DSceneFromBlueprint(null);
        renderDynamicSchematicSVG(null);
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
      if (pageId === "page-3d") setTimeout(resizeThreeJS, 50);
    }

    async function generateProject() {
      const topic = document.getElementById("topic-input").value.trim();
      const description = document.getElementById("desc-input").value.trim();
      const budget_tier = document.getElementById("budget-input").value;
      const branch = document.getElementById("branch-input").value;
      let subject = document.getElementById("subject-input").value;
      if (subject === "__MANUAL__") subject = document.getElementById("manual-subject-input").value.trim();
      if (!topic) return alert("Please enter your project title first!");

      const btn = document.getElementById("generate-btn");
      btn.disabled = true; btn.innerText = "Verifying & Building Tailored Project...";
      try {
        const res = await fetch("/api/generate", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ topic, description, branch, subject, level: "Academic Project", budget_tier })
        });
        const rawText = await res.text();
        let data;
        try { data = JSON.parse(rawText); } catch (_) { throw new Error("Server waking up. Please click Generate again!"); }
        if (data.error) throw new Error(data.error);
        currentBlueprint = data.blueprint;
        populateAllPages(currentBlueprint);
        navigatePage("page-3d");
      } catch (err) {
        alert(err.message);
      } finally {
        btn.disabled = false; btn.innerText = "Generate Tailored Project";
      }
    }

    function renderDynamicSchematicSVG(bp) {
      const s = (bp && bp.schematic_spec) || {
        supply_voltage: "+5V DC",
        core_ic: "Core Logic / Controller",
        input_stage: "Input Stage",
        control_stage: "Processing Stage",
        output_stage: "Output Stage"
      };
      document.getElementById("schem-heading").innerText = `📐 Circuit Diagram — ${(bp && bp.title) || "Dynamic Hardware"}`;

      const svg = `
      <svg viewBox="0 0 860 280" width="100%" height="280" xmlns="http://www.w3.org/2000/svg" style="background:#ffffff;">
        <line x1="60" y1="36" x2="800" y2="36" stroke="#dc2626" stroke-width="2.5"/>
        <text x="65" y="24" font-size="12" font-weight="bold" fill="#dc2626">${s.supply_voltage || "+5V DC"} SUPPLY</text>

        <rect x="60" y="95" width="210" height="100" rx="10" fill="#f0fdf4" stroke="#16a34a" stroke-width="2.5"/>
        <text x="74" y="122" font-size="11" font-weight="bold" fill="#15803d">INPUT STAGE</text>
        <text x="74" y="148" font-size="13" font-weight="bold" fill="#0f172a">${s.input_stage || "Sensor / Input"}</text>
        <text x="74" y="172" font-size="11" fill="#475569">Captures signal / trigger</text>

        <line x1="270" y1="145" x2="325" y2="145" stroke="#1e293b" stroke-width="2.5"/>
        <polygon points="325,140 335,145 325,150" fill="#1e293b"/>

        <rect x="335" y="90" width="210" height="110" rx="10" fill="#eff6ff" stroke="#1e3a8a" stroke-width="2.5"/>
        <text x="349" y="118" font-size="11" font-weight="bold" fill="#1d4ed8">CONTROL / LOGIC STAGE</text>
        <text x="349" y="145" font-size="13" font-weight="bold" fill="#0f172a">${s.core_ic || s.control_stage || "Core Module"}</text>
        <text x="349" y="170" font-size="11" fill="#475569">Evaluates &amp; drives output</text>

        <line x1="545" y1="145" x2="600" y2="145" stroke="#1e293b" stroke-width="2.5"/>
        <polygon points="600,140 610,145 600,150" fill="#1e293b"/>

        <rect x="610" y="95" width="200" height="100" rx="10" fill="#fef2f2" stroke="#dc2626" stroke-width="2.5"/>
        <text x="624" y="122" font-size="11" font-weight="bold" fill="#b91c1c">OUTPUT STAGE</text>
        <text x="624" y="148" font-size="13" font-weight="bold" fill="#0f172a">${s.output_stage || "LED / Buzzer / Load"}</text>
        <text x="624" y="172" font-size="11" fill="#475569">Action / indicator response</text>

        <line x1="60" y1="245" x2="800" y2="245" stroke="#0f172a" stroke-width="2.5"/>
        <text x="65" y="266" font-size="12" font-weight="bold" fill="#0f172a">0V COMMON GROUND (GND)</text>
      </svg>`;
      document.getElementById("schematic-svg-container").innerHTML = svg;
    }

    function printCircuitDiagram() {
      const content = document.getElementById("schematic-svg-container").innerHTML;
      const win = window.open("", "_blank");
      win.document.write(`<html><body style="font-family:sans-serif;padding:24px;">${content}<script>window.print();<\/script></body></html>`);
      win.document.close();
    }

    function copyBomList() {
      if (!currentBlueprint) return;
      const txt = (currentBlueprint.hardware_requirements || [])
        .map(h => `${h.name} (${h.spec}) - ₹${h.cost_inr}`).join("\n");
      navigator.clipboard.writeText(txt);
      alert("Bill of Materials copied!");
    }

    function populateAllPages(bp) {
      document.getElementById("est-cost-display").innerText = `Est. ₹${bp.estimated_total_inr || 0}`;
      document.getElementById("sidebar-bom-list").innerHTML = (bp.hardware_requirements || []).map(h =>
        `<div class="flex justify-between gap-2"><span>${h.name}</span><span class="font-bold text-emerald-700">₹${h.cost_inr}</span></div>`
      ).join("");

      if (bp.requires_circuit) {
        document.getElementById("circuit-lab-wrapper").classList.remove("hidden");
        document.getElementById("software-arch-wrapper").classList.add("hidden");
        document.getElementById("lab3d-title").innerText = `3D View: ${bp.title} (Est. ₹${bp.estimated_total_inr})`;
        renderDynamicSchematicSVG(bp);
        build3DSceneFromBlueprint(bp);

        const pins = (bp.schematic_spec && bp.schematic_spec.pins) || [];
        document.getElementById("lab3d-pins").innerHTML = pins.map(p => `
          <tr>
            <td class="p-2.5 font-mono font-bold text-blue-900">${p.pin_no}</td>
            <td class="p-2.5">${p.pin_name}</td>
            <td class="p-2.5 text-slate-800">${p.connects_to}</td>
            <td class="p-2.5 font-bold">${p.wire_color}</td>
          </tr>
        `).join("");
      } else {
        document.getElementById("circuit-lab-wrapper").classList.add("hidden");
        document.getElementById("software-arch-wrapper").classList.remove("hidden");
        document.getElementById("sw-arch-title").innerText = `System Architecture: ${bp.title}`;
        document.getElementById("sw-nodes-container").innerHTML = (bp.software_architecture_nodes || []).map((n, i) => `
          <div class="bg-blue-50 border border-blue-200 rounded-xl p-3.5 text-center">
            <span class="text-[10px] font-bold text-blue-800">STAGE 0${i + 1}</span>
            <p class="text-xs font-bold text-slate-900 mt-1">${n}</p>
          </div>
        `).join("");
      }

      document.getElementById("active-code-subtitle").innerText = `Language: ${bp.primary_code_language || "Code"}`;
      document.getElementById("code-block").innerText = bp.primary_code || "";
      renderVivaCards();

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
            <button type="button" onclick="toggleVivaAns('va-${idx}')" class="text-[11px] font-semibold text-blue-700 underline">Show / Hide Answer</button>
            <p id="va-${idx}" class="text-xs text-slate-700 leading-relaxed">${entry.a || ""}</p>
          </div>
        `;
      }).join("");
    }

    function toggleVivaAns(id) {
      document.getElementById(id).classList.toggle("hidden");
    }

    async function downloadFile(endpoint, filename, btnId) {
      if (!currentBlueprint) return;
      const btn = document.getElementById(btnId);
      const old = btn.innerText;
      btn.disabled = true; btn.innerText = "Preparing...";
      try {
        const res = await fetch(endpoint, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            blueprint: currentBlueprint,
            branch: document.getElementById("branch-input").value,
            subject: document.getElementById("subject-input").value,
            guide_name: document.getElementById("guide-input").value.trim() || "Dr. Shivakumar G",
            viva_lang: activeVivaLang
          })
        });
        const blob = await res.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url; a.download = filename;
        document.body.appendChild(a); a.click(); a.remove();
      } finally {
        btn.disabled = false; btn.innerText = old;
      }
    }

    // ==================== 3D ENGINE ====================
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
      camera.position.set(0, 8.4, 11.2);
      renderer = new THREE.WebGLRenderer({ antialias: true });
      renderer.setPixelRatio(Math.min(window.devicePixelRatio, 2));
      renderer.setSize(container.clientWidth, container.clientHeight);
      container.innerHTML = "";
      container.appendChild(renderer.domElement);
      if (THREE.OrbitControls) {
        controls = new THREE.OrbitControls(camera, renderer.domElement);
        controls.enableDamping = true;
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

    function createPartMesh(type) {
      const grp = new THREE.Group();
      if (type === "ic_741" || type === "ic_555") {
        const body = new THREE.Mesh(new THREE.BoxGeometry(0.85, 0.24, 0.6), new THREE.MeshStandardMaterial({ color: 0x111827 }));
        grp.add(body);
      } else if (type === "sensor_water_probes") {
        const tank = new THREE.Mesh(new THREE.BoxGeometry(1.2, 1.0, 0.85), new THREE.MeshStandardMaterial({ color: 0x7dd3fc, transparent: true, opacity: 0.45 }));
        const water = new THREE.Mesh(new THREE.BoxGeometry(1.1, 0.7, 0.78), new THREE.MeshStandardMaterial({ color: 0x0284c7, transparent: true, opacity: 0.82 }));
        grp.add(tank, water);
        animParts.push({ mesh: water, mode: "water" });
      } else if (type === "led_indicator_bank" || type === "piezo_buzzer") {
        [0x22c55e, 0xeab308, 0xef4444].forEach((clr, idx) => {
          const led = new THREE.Mesh(new THREE.SphereGeometry(0.14, 14, 14), new THREE.MeshBasicMaterial({ color: clr }));
          led.position.set(-0.3 + idx * 0.3, 0.24, -0.1);
          grp.add(led);
          if (idx === 2) animParts.push({ mesh: led, mode: "blink" });
        });
      } else {
        const box = new THREE.Mesh(new THREE.BoxGeometry(1.15, 0.7, 0.35), new THREE.MeshStandardMaterial({ color: 0x0284c7 }));
        grp.add(box);
      }
      return grp;
    }

    function add3DWire(p1, p2, p3, colorHex) {
      const curve = new THREE.CatmullRomCurve3([new THREE.Vector3(...p1), new THREE.Vector3(...p2), new THREE.Vector3(...p3)]);
      circuitGroup.add(new THREE.Mesh(new THREE.TubeGeometry(curve, 20, 0.04, 8, false), new THREE.MeshStandardMaterial({ color: colorHex })));
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

      const bb = new THREE.Mesh(new THREE.BoxGeometry(5.6, 0.26, 2.8), new THREE.MeshStandardMaterial({ color: 0xffffff }));
      bb.position.set(1.5, 0, 0);
      circuitGroup.add(bb);
      registerLabel("Breadboard", 1.5, 0.55, 1.3);

      const parts = (bp && bp.circuit_3d_parts && bp.circuit_3d_parts.length) ? bp.circuit_3d_parts : [
        { type: "sensor_ir", label: "Input Module", wire_color: "#2563eb" },
        { type: "led_indicator_bank", label: "Indicators", wire_color: "#16a34a" },
        { type: "piezo_buzzer", label: "Output Stage", wire_color: "#dc2626" }
      ];

      const xSlots = [-0.1, 1.5, 3.1];
      parts.slice(0, 3).forEach((part, i) => {
        const px = xSlots[i];
        const mesh = createPartMesh(part.type);
        mesh.position.set(px, 0.5, 0);
        circuitGroup.add(mesh);
        if (i > 0) add3DWire([xSlots[i - 1], 0.35, 0], [(px + xSlots[i - 1]) / 2, 1.35, 0], [px, 0.35, 0], part.wire_color || "#2563eb");
        registerLabel(part.label || "Module", px, 1.0, 0);
      });
    }

    let simTick = 0;
    function animate3D() {
      requestAnimationFrame(animate3D);
      if (controls) controls.update();
      if (simActive) {
        simTick += 0.05;
        animParts.forEach(item => {
          if (item.mode === "blink") {
            const on = sensorStimulusLevel >= 0.65;
            item.mesh.material.color.setHex(on && Math.sin(simTick * 3) > 0 ? 0xef4444 : 0x7f1d1d);
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
      document.getElementById("sim-toggle-btn").innerText = simActive ? "▶ Sim: ON" : "⏸️ Sim: PAUSED";
    }
    function reset3DCamera() {
      if (camera) camera.position.set(0, 8.4, 11.2);
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


# ==================== 5. FLASK ROUTES ====================
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


@app.route("/")
def index():
    return render_template_string(HTML_PAGE)


@app.route("/api/generate", methods=["POST"])
def api_generate():
    data = request.get_json(force=True)
    topic = data.get("topic", "").strip()
    description = data.get("description", "").strip()
    branch = data.get("branch", "Electronics & Communication (ECE)").strip()
    subject = data.get("subject", "Core Engineering").strip()
    level = data.get("level", "Academic Project").strip()
    budget_tier = data.get("budget_tier", "₹100 - ₹250 (Ultra Low-Cost Analog / IC Circuit)").strip()
    if not topic:
        return jsonify({"error": "Please enter a project title."}), 400
    blueprint = generate_project_blueprint(topic, description, branch, subject, level, budget_tier)
    return jsonify({"status": "ok", "blueprint": blueprint})


@app.route("/api/download-docx", methods=["POST"])
def download_docx():
    payload = request.get_json(force=True)
    docx_buffer = build_vtu_docx(
        payload.get("blueprint", {}),
        "Student Name",
        "1AM25EC099",
        "AMC Engineering College",
        payload.get("branch") or "ECE",
        payload.get("subject") or "Engineering Subject",
        payload.get("guide_name") or "Dr. Shivakumar G",
        payload.get("viva_lang") or "en"
    )
    return send_file(docx_buffer, as_attachment=True, download_name="ProjForge_Report.docx", mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@app.route("/api/download-pptx", methods=["POST"])
def download_pptx():
    payload = request.get_json(force=True)
    pptx_buffer = build_seminar_pptx(
        payload.get("blueprint", {}),
        "Student Name",
        "1AM25EC099",
        "AMC Engineering College",
        payload.get("branch") or "ECE",
        payload.get("subject") or "Engineering Subject",
        payload.get("guide_name") or "Dr. Shivakumar G"
    )
    return send_file(pptx_buffer, as_attachment=True, download_name="ProjForge_Seminar.pptx", mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))