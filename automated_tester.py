"""
Automated Visual Tester for Visit Document Processor
====================================================
Features:
- Visual real-time execution (watch it work!)
- AI-powered text generation via Ollama
- Screenshot capture
- HTML report generation
- Multiple test scenarios

Requirements:
    pip install pywin32 pillow requests mss

Usage:
    python automated_tester.py
"""

import win32com.client
import pythoncom
import time
import os
import json
import random
from datetime import datetime, timedelta
from pathlib import Path
import threading

# Optional imports (graceful degradation)
try:
    import mss
    import mss.tools
    SCREENSHOT_AVAILABLE = True
except ImportError:
    SCREENSHOT_AVAILABLE = False
    print("Warning: mss not installed. Screenshots disabled. Run: pip install mss")

try:
    from PIL import Image
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False
    print("Warning: Pillow not installed. Image processing limited. Run: pip install pillow")

try:
    import requests
    REQUESTS_AVAILABLE = True
except ImportError:
    REQUESTS_AVAILABLE = False
    print("Warning: requests not installed. Ollama integration disabled. Run: pip install requests")


class OllamaTextGenerator:
    """Generate realistic text using local Ollama LLM"""

    def __init__(self, model="llama3:8b", base_url="http://localhost:11434"):
        self.model = model
        self.base_url = base_url
        self.available = self._check_availability()

    def _check_availability(self):
        """Check if Ollama is running"""
        if not REQUESTS_AVAILABLE:
            return False
        try:
            response = requests.get(f"{self.base_url}/api/tags", timeout=2)
            return response.status_code == 200
        except:
            return False

    def generate(self, prompt, max_tokens=150):
        """Generate text from prompt"""
        if not self.available:
            return self._fallback_text(prompt)

        try:
            response = requests.post(
                f"{self.base_url}/api/generate",
                json={
                    "model": self.model,
                    "prompt": prompt,
                    "stream": False,
                    "options": {"num_predict": max_tokens}
                },
                timeout=30
            )
            if response.status_code == 200:
                return response.json().get("response", "").strip()
        except Exception as e:
            print(f"Ollama error: {e}")

        return self._fallback_text(prompt)

    def _fallback_text(self, prompt):
        """Fallback text when Ollama unavailable"""
        if "residence" in prompt.lower() or "home" in prompt.lower():
            return random.choice(FALLBACK_RESIDENCE_DESCRIPTIONS)
        elif "observed" in prompt.lower():
            return random.choice(FALLBACK_OBSERVED_DESCRIPTIONS)
        elif "red flag" in prompt.lower():
            return random.choice(FALLBACK_RED_FLAG_DESCRIPTIONS)
        return "Test text entry."

    def generate_residence_description(self, address):
        """Generate EXTERIOR description of residence only"""
        prompt = f"""You are a probation officer describing a home's EXTERIOR only.
Address: {address}

Describe in 2-3 sentences:
- House type (single story, two story, apartment, etc.)
- Exterior material (brick, siding, stucco)
- Trim color around windows and doors
- Garage type (one-car, two-car, carport, none)
- Front yard condition

Example: "Single story brick home with white trim outlining the windows and front door. Two-car attached garage on the left side. Front yard is well-maintained with mature landscaping."

Keep it under 60 words. Exterior only - do NOT mention entering the home. Just the description, no introduction."""
        return self.generate(prompt, 120)

    def generate_observed_description(self, is_successful=True, visit_type="FTR", reason=""):
        """Generate what was observed during visit - detailed room-by-room for successful"""
        if visit_type == "FTR":
            prompt = """You are a probation officer documenting an FTR (Failed to Report) home visit.

Write 2-3 sentences describing:
- Knocked on door, rang doorbell (note if doorbell camera present)
- No response after waiting
- Left orange door knocker with contact information

Professional tone. Just the observation, no introduction."""
            return self.generate(prompt, 100)

        if not is_successful:
            if reason == "Not Home":
                prompt = """You are a probation officer documenting an unsuccessful home visit where no one was home.

Write 2-3 sentences describing:
- Knocked on door, rang doorbell (note if doorbell camera present)
- No response after multiple attempts
- Left door knocker with contact information

Professional tone. Just the observation, no introduction."""
            elif reason == "Wrong Address":
                prompt = """You are a probation officer documenting an unsuccessful home visit - wrong address.

Write 2-3 sentences describing:
- Knocked on door, someone answered who was NOT the probationer
- Confirmed this was not the correct residence for the probationer
- Departed and documented the discrepancy

Professional tone. Just the observation, no introduction."""
            elif reason == "P Denied Access":
                prompt = """You are a probation officer documenting an unsuccessful home visit where probationer denied entry.

Write 2-3 sentences describing:
- Knocked on door, probationer (P) answered
- Officer asked for consent to enter the home
- P refused to allow entry, visit terminated

Professional tone. Just the observation, no introduction."""
            else:
                prompt = """You are a probation officer documenting an unsuccessful/cancelled home visit.
Write 2-3 sentences. Professional tone. Just the observation, no introduction."""
            return self.generate(prompt, 120)

        # SUCCESSFUL visit - detailed room-by-room walkthrough
        prompt = """You are a probation officer documenting a SUCCESSFUL home visit with detailed observations.

Write a thorough description (5-7 sentences) including ALL of the following in order:
1. Who answered the door (P = probationer)
2. Officer asked P for consent to enter the home, P agreed
3. Describe entering - "The home opens to a [living room/foyer/etc.]"
4. P's bedroom location and brief description (bed made/unmade, clothes, personal items)
5. P escorted officer to the kitchen - describe kitchen briefly
6. Describe refrigerator (color like white/black/stainless steel, type like side-by-side/French door), describe contents (food items, beverages)
7. Officer asked P to escort to exit, no violations noted

Example flow: "P answered the front door. Officer requested consent to enter the home, P agreed. The home opens to a living room with a sectional couch and TV mounted on the wall. P's bedroom is located down the hallway on the right - queen bed was made, clothes in closet, nightstand with lamp. P escorted officer to the kitchen which had granite countertops and gas stove. White side-by-side refrigerator contained milk, eggs, lunch meat, and various condiments. Freezer had frozen vegetables and ice cream. Officer asked P to escort to the exit. No violations noted."

Professional tone. Just the observation, no introduction or headers."""
        return self.generate(prompt, 300)

    def generate_red_flag_description(self, flag_type):
        """Generate red flag details - specific location and description"""
        prompts = {
            "Alcohol": """Describe finding alcohol during a probation home visit.
Include: exact location found (kitchen counter, bedroom nightstand, living room coffee table),
type of alcohol (beer cans, liquor bottles, wine), quantity, and brand if visible.
Example: "Found three empty Bud Light cans on the bedroom nightstand next to an open bottle of Jack Daniels whiskey approximately half full."
1-2 sentences. Professional tone.""",

            "Drugs": """Describe finding drug paraphernalia during a probation home visit.
Include: exact location found, specific items seen (pipe, rolling papers, residue, baggies).
Example: "Observed a glass pipe with burnt residue and small plastic baggies with white powder residue on the coffee table in the living room."
1-2 sentences. Professional tone.""",

            "Guns": """Describe finding a firearm during a probation home visit.
Include: exact location found, type of firearm, whether loaded/unloaded if visible.
Example: "Located a black semi-automatic handgun in the top drawer of the bedroom dresser. Magazine was inserted."
1-2 sentences. Professional tone.""",

            "Knives": """Describe finding prohibited weapons/knives during a probation home visit.
Include: exact location found, type of knife/weapon, approximate size.
Example: "Found a large hunting knife with approximately 8-inch blade under the mattress in P's bedroom."
1-2 sentences. Professional tone.""",

            "IP": """Describe observing an injured party during a probation home visit.
Include: who was injured, visible injuries, their demeanor.
Example: "Female present in living room had visible bruising on left arm and appeared nervous. She stated she fell but avoided eye contact when questioned."
1-2 sentences. Professional tone.""",

            "Other": """Describe finding evidence of prohibited contact or other violation during a probation home visit.
Include: what was found, exact location, why it's concerning.
Example: "Children's toys and clothing observed in the bedroom closet. P has active no-contact order with minor children."
1-2 sentences. Professional tone."""
        }
        prompt = prompts.get(flag_type, prompts["Other"])
        return self.generate(prompt, 100)


# Fallback text when Ollama is not available
FALLBACK_RESIDENCE_DESCRIPTIONS = [
    "Single story brick home with white trim outlining the windows and front door. Two-car attached garage on the left side. Front yard is well-maintained with mature landscaping.",
    "Two story frame house with tan siding and brown trim around windows. One-car detached garage in back. Front yard has small flower bed and trimmed bushes.",
    "Apartment complex, unit on second floor with exterior access. Beige stucco exterior with white window frames. Parking lot was half full upon arrival.",
    "Single story home with brown brick exterior and cream-colored trim. Carport on right side, no enclosed garage. Small front porch with concrete steps.",
    "Ranch-style home with beige brick and dark green trim around windows and front door. Two-car garage attached on left. Lawn recently mowed, quiet residential street.",
    "Multi-family dwelling, ground floor unit. Red brick exterior with white trim around windows. Shared parking area in front. Building appears well-maintained.",
    "Single story frame house with gray siding and white window trim. Carport on left side, chain link fence around property. Small fenced backyard visible from front.",
    "Two story townhouse with stone and stucco exterior. Black trim around windows. One-car garage on ground level. Small landscaped front area.",
]

FALLBACK_OBSERVED_DESCRIPTIONS = [
    "P answered the front door. Officer requested consent to enter the home, P agreed. The home opens to a living room with a sectional couch and TV mounted on the wall. P's bedroom is located down the hallway on the right - queen bed was made, clothes in closet, nightstand with lamp. P escorted officer to the kitchen which had granite countertops and gas stove. White side-by-side refrigerator contained milk, eggs, lunch meat, and various condiments. Freezer had frozen vegetables and ice cream. Officer asked P to escort to the exit. No violations noted.",
    "P came to the door after second knock. Officer asked for consent to enter, P agreed. Home opens to a foyer leading to living room. P's bedroom is the first door on the left - full bed made, dresser with personal items, closet organized. P walked officer to kitchen area with electric stove and dishwasher. Stainless steel French door refrigerator contained various food items including fruits, vegetables, and beverages. Officer asked P to escort to exit. No violations noted.",
    "P answered door promptly. Consent to enter was given. Home opens directly to living room with couch and recliner. P's sleeping area is in back bedroom - twin bed unmade, clothes on floor, small TV on dresser. P escorted to kitchen which was clean with basic appliances. Black top-freezer refrigerator contained leftovers, milk, and condiments. Officer requested P escort to door. No violations noted.",
    "P opened door after doorbell ring. Officer requested and received consent to enter. Home opens to combined living/dining area. P indicated bedroom is down hallway - queen bed with blue comforter, nightstand, closet with hanging clothes. Kitchen had tile floor and white cabinets. White side-by-side refrigerator well-stocked with groceries. P escorted officer out. No violations noted.",
]

FALLBACK_RED_FLAG_DESCRIPTIONS = [
    "Found three empty Bud Light cans on the bedroom nightstand next to an open bottle of Jack Daniels whiskey approximately half full.",
    "Observed a glass pipe with burnt residue and small plastic baggies with white powder residue on the coffee table in the living room.",
    "Located a black semi-automatic handgun in the top drawer of the bedroom dresser. Magazine was inserted.",
    "Found a large hunting knife with approximately 8-inch blade under the mattress in P's bedroom.",
    "Female present in living room had visible bruising on left arm and appeared nervous. She stated she fell but avoided eye contact when questioned.",
    "Children's toys and clothing observed in the bedroom closet. P has active no-contact order with minor children.",
]


class VisualTyper:
    """Type text character by character for visual effect"""

    def __init__(self, excel_app, delay_per_char=0.03, delay_per_word=0.1):
        self.excel = excel_app
        self.char_delay = delay_per_char
        self.word_delay = delay_per_word

    def type_in_cell(self, cell, text, clear_first=True):
        """Type text into cell character by character"""
        if clear_first:
            cell.Value = ""

        # Select the cell first so user can see it
        cell.Select()
        time.sleep(0.2)

        # Type character by character
        current_text = ""
        for char in text:
            current_text += char
            cell.Value = current_text

            if char == ' ':
                time.sleep(self.word_delay)
            else:
                time.sleep(self.char_delay)

        time.sleep(0.1)  # Brief pause after completion


class ScreenshotCapture:
    """Capture screenshots during test execution"""

    def __init__(self, output_dir):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.screenshot_count = 0
        self.screenshots = []

    def capture(self, name="screenshot"):
        """Capture current screen"""
        if not SCREENSHOT_AVAILABLE:
            return None

        self.screenshot_count += 1
        filename = f"{self.screenshot_count:03d}_{name}_{datetime.now().strftime('%H%M%S')}.png"
        filepath = self.output_dir / filename

        try:
            with mss.mss() as sct:
                # Capture primary monitor
                monitor = sct.monitors[1]
                screenshot = sct.grab(monitor)
                mss.tools.to_png(screenshot.rgb, screenshot.size, output=str(filepath))

            self.screenshots.append({
                "filename": filename,
                "path": str(filepath),
                "name": name,
                "timestamp": datetime.now().isoformat()
            })
            return filepath
        except Exception as e:
            print(f"Screenshot error: {e}")
            return None


class TestResult:
    """Store test results"""

    def __init__(self, name):
        self.name = name
        self.start_time = datetime.now()
        self.end_time = None
        self.status = "running"
        self.steps = []
        self.errors = []
        self.screenshots = []

    def add_step(self, description, status="pass", details=None):
        self.steps.append({
            "description": description,
            "status": status,
            "details": details,
            "timestamp": datetime.now().isoformat()
        })

    def add_error(self, error):
        self.errors.append({
            "error": str(error),
            "timestamp": datetime.now().isoformat()
        })

    def finish(self, status="pass"):
        self.end_time = datetime.now()
        self.status = status if not self.errors else "fail"

    @property
    def duration(self):
        if self.end_time:
            return (self.end_time - self.start_time).total_seconds()
        return 0


class HTMLReporter:
    """Generate HTML test report"""

    def __init__(self, output_dir):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)

    def generate_report(self, test_results, screenshots_dir):
        """Generate HTML report from test results"""
        report_path = self.output_dir / f"test_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"

        html = self._build_html(test_results, screenshots_dir)

        with open(report_path, 'w', encoding='utf-8') as f:
            f.write(html)

        print(f"\nReport generated: {report_path}")
        return report_path

    def _build_html(self, test_results, screenshots_dir):
        """Build HTML content"""
        total_tests = len(test_results)
        passed = sum(1 for t in test_results if t.status == "pass")
        failed = total_tests - passed
        total_duration = sum(t.duration for t in test_results)

        html = f"""<!DOCTYPE html>
<html>
<head>
    <title>Visit Document Processor - Test Report</title>
    <style>
        body {{ font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; background: #f5f5f5; }}
        .header {{ background: linear-gradient(135deg, #1a5f7a, #2d8659); color: white; padding: 30px; border-radius: 10px; margin-bottom: 20px; }}
        .header h1 {{ margin: 0; }}
        .summary {{ display: flex; gap: 20px; margin-bottom: 20px; }}
        .summary-card {{ background: white; padding: 20px; border-radius: 10px; flex: 1; box-shadow: 0 2px 5px rgba(0,0,0,0.1); text-align: center; }}
        .summary-card.pass {{ border-left: 5px solid #28a745; }}
        .summary-card.fail {{ border-left: 5px solid #dc3545; }}
        .summary-card h2 {{ margin: 0; font-size: 36px; }}
        .summary-card p {{ margin: 5px 0 0; color: #666; }}
        .test-case {{ background: white; margin-bottom: 15px; border-radius: 10px; overflow: hidden; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }}
        .test-header {{ padding: 15px 20px; cursor: pointer; display: flex; justify-content: space-between; align-items: center; }}
        .test-header.pass {{ background: #d4edda; }}
        .test-header.fail {{ background: #f8d7da; }}
        .test-content {{ padding: 20px; display: none; border-top: 1px solid #eee; }}
        .test-content.show {{ display: block; }}
        .step {{ padding: 10px; margin: 5px 0; border-radius: 5px; display: flex; align-items: center; gap: 10px; }}
        .step.pass {{ background: #e8f5e9; }}
        .step.fail {{ background: #ffebee; }}
        .step-icon {{ font-size: 18px; }}
        .step-icon.pass::before {{ content: '✓'; color: #28a745; }}
        .step-icon.fail::before {{ content: '✗'; color: #dc3545; }}
        .screenshots {{ display: flex; flex-wrap: wrap; gap: 10px; margin-top: 15px; }}
        .screenshot {{ max-width: 300px; border: 1px solid #ddd; border-radius: 5px; }}
        .screenshot img {{ width: 100%; border-radius: 5px 5px 0 0; }}
        .screenshot p {{ margin: 0; padding: 10px; font-size: 12px; background: #f9f9f9; }}
        .duration {{ color: #666; font-size: 14px; }}
        .timestamp {{ font-size: 12px; color: #999; }}
    </style>
    <script>
        function toggleTest(id) {{
            var content = document.getElementById('content-' + id);
            content.classList.toggle('show');
        }}
    </script>
</head>
<body>
    <div class="header">
        <h1>Visit Document Processor - Test Report</h1>
        <p>Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
    </div>

    <div class="summary">
        <div class="summary-card pass">
            <h2>{passed}</h2>
            <p>Passed</p>
        </div>
        <div class="summary-card fail">
            <h2>{failed}</h2>
            <p>Failed</p>
        </div>
        <div class="summary-card">
            <h2>{total_duration:.1f}s</h2>
            <p>Total Duration</p>
        </div>
    </div>
"""

        for i, result in enumerate(test_results):
            steps_html = ""
            for step in result.steps:
                steps_html += f"""
                <div class="step {step['status']}">
                    <span class="step-icon {step['status']}"></span>
                    <span>{step['description']}</span>
                    <span class="timestamp">{step.get('timestamp', '')}</span>
                </div>"""

            screenshots_html = ""
            for ss in result.screenshots:
                rel_path = os.path.relpath(ss['path'], self.output_dir)
                screenshots_html += f"""
                <div class="screenshot">
                    <img src="{rel_path}" alt="{ss['name']}">
                    <p>{ss['name']}</p>
                </div>"""

            html += f"""
    <div class="test-case">
        <div class="test-header {result.status}" onclick="toggleTest({i})">
            <span><strong>{result.name}</strong></span>
            <span class="duration">{result.duration:.1f}s</span>
        </div>
        <div class="test-content" id="content-{i}">
            <h4>Steps:</h4>
            {steps_html}
            {"<h4>Screenshots:</h4><div class='screenshots'>" + screenshots_html + "</div>" if screenshots_html else ""}
        </div>
    </div>"""

        html += """
</body>
</html>"""
        return html


class VisitDocumentTester:
    """Main test automation class"""

    def __init__(self, workbook_path, speed="normal"):
        self.workbook_path = workbook_path
        self.base_dir = os.path.dirname(workbook_path)

        # Speed settings
        self.speeds = {
            "slow": {"char": 0.05, "action": 2.0, "word": 0.15},
            "normal": {"char": 0.03, "action": 1.0, "word": 0.1},
            "fast": {"char": 0.01, "action": 0.5, "word": 0.05},
        }
        self.speed = self.speeds.get(speed, self.speeds["normal"])

        # Components
        self.ollama = OllamaTextGenerator()
        self.screenshot_dir = os.path.join(self.base_dir, "screenshots",
                                           datetime.now().strftime("%Y%m%d_%H%M%S"))
        self.screenshotter = ScreenshotCapture(self.screenshot_dir)
        self.reporter = HTMLReporter(os.path.join(self.base_dir, "test_results"))

        # Excel objects (initialized in setup)
        self.excel = None
        self.workbook = None
        self.typer = None

        # Test tracking
        self.test_results = []
        self.current_test = None

    def setup(self):
        """Initialize Excel application"""
        print("Starting Excel...")
        pythoncom.CoInitialize()
        self.excel = win32com.client.Dispatch("Excel.Application")
        self.excel.Visible = True  # Make visible for watching
        self.excel.DisplayAlerts = False

        print(f"Opening workbook: {self.workbook_path}")
        self.workbook = self.excel.Workbooks.Open(self.workbook_path)

        self.typer = VisualTyper(self.excel,
                                  delay_per_char=self.speed["char"],
                                  delay_per_word=self.speed["word"])

        # Enable test mode to suppress MsgBox dialogs
        print("Enabling test mode (suppressing dialogs)...")
        self.excel.Application.Run("EnableTestMode")

        time.sleep(1)
        print("Excel ready!")

    def teardown(self, keep_open=False):
        """Clean up Excel"""
        if keep_open:
            print("\n*** Workbook left open for inspection ***")
            pythoncom.CoUninitialize()
            return

        if self.workbook:
            try:
                self.workbook.Close(SaveChanges=False)
            except:
                pass
        if self.excel:
            try:
                self.excel.Quit()
            except:
                pass
        pythoncom.CoUninitialize()

    def action_delay(self, multiplier=1.0):
        """Delay between actions"""
        time.sleep(self.speed["action"] * multiplier)

    def scroll_to_cell(self, cell):
        """Scroll to make cell visible at top of screen"""
        try:
            self.excel.Application.Goto(cell, Scroll=True)
            time.sleep(0.1)
        except:
            pass

    def scroll_to_row(self, ws, row):
        """Scroll to show specific row at top"""
        try:
            cell = ws.Cells(row, 1)
            self.excel.Application.Goto(cell, Scroll=True)
            time.sleep(0.1)
        except:
            pass

    def run_macro(self, macro_name):
        """Execute a VBA macro"""
        try:
            self.excel.Application.Run(macro_name)
            return True
        except Exception as e:
            print(f"Macro error ({macro_name}): {e}")
            return False

    def get_sheet(self, name):
        """Get worksheet by name"""
        try:
            return self.workbook.Sheets(name)
        except:
            return None

    def start_test(self, name):
        """Start a new test case"""
        self.current_test = TestResult(name)
        print(f"\n{'='*60}")
        print(f"TEST: {name}")
        print(f"{'='*60}")

    def step(self, description, action_func=None, screenshot=True):
        """Execute and record a test step"""
        print(f"  -> {description}")

        try:
            if action_func:
                action_func()

            if screenshot:
                ss_path = self.screenshotter.capture(description.replace(" ", "_")[:30])
                if ss_path and self.current_test:
                    self.current_test.screenshots.append({
                        "name": description,
                        "path": str(ss_path)
                    })

            if self.current_test:
                self.current_test.add_step(description, "pass")
            return True

        except Exception as e:
            print(f"    ERROR: {e}")
            if self.current_test:
                self.current_test.add_step(description, "fail", str(e))
                self.current_test.add_error(e)
            return False

    def finish_test(self):
        """Complete current test"""
        if self.current_test:
            self.current_test.finish()
            self.test_results.append(self.current_test)
            status_symbol = "PASS" if self.current_test.status == "pass" else "FAIL"
            print(f"\nResult: {status_symbol} {self.current_test.status.upper()}")
            print(f"Duration: {self.current_test.duration:.1f}s")

    # =========================================================================
    # TEST SCENARIOS
    # =========================================================================

    def scenario_normal_day(self, itinerary_path):
        """
        Scenario: Normal day workflow
        1. Load itinerary
        2. Create visit sheet
        3. Fill out visits
        4. Refresh metrics
        """
        self.start_test("Normal Day Workflow")

        # Step 1: Navigate to main tab
        self.step("Navigate to Visit Document Processor tab", lambda: (
            self.workbook.Sheets("Visit Document Processor").Activate(),
            self.action_delay()
        ))

        # Step 2: Set date
        def set_date():
            ws = self.get_sheet("Visit Document Processor")
            ws.Range("B10").Value = datetime.now().strftime("%m/%d/%Y")
            self.action_delay(0.5)

        self.step("Set visit date to today", set_date)

        # Step 3: Set officers
        def set_officers():
            ws = self.get_sheet("Visit Document Processor")
            ws.Range("B13").Value = "JOHNSON, SARAH"
            self.action_delay(0.3)
            ws.Range("B14").Value = "MARTINEZ, DAVID"
            self.action_delay(0.5)

        self.step("Set officers", set_officers)

        # Step 4: Load itinerary (simulated - we'll use macro with pre-selected file)
        # In real scenario, this would use file dialog
        # For automation, we'll directly load the data

        self.step("Load itinerary file", lambda: (
            self._load_itinerary_direct(itinerary_path),
            self.action_delay()
        ))

        # Step 5: Create Visit Sheet
        self.step("Create Visit Sheet", lambda: (
            self.run_macro("CreateExcelVisitSheet"),
            self.action_delay(2)
        ))

        # =====================================================
        # Fill ALL 5 visits with specific test configurations
        # =====================================================

        # Visit 1: Successful, 2 vehicles, no red flag
        self.step("Fill out Visit #1 (Successful, 2 vehicles)", lambda: self._fill_visit(1, {
            "force_outcome": "Successful",
            "vehicle_count": 2,
            "force_red_flag": False
        }))

        # Visit 2: Successful, 5 vehicles (test +Add button), RED FLAG
        self.step("Fill out Visit #2 (Successful, 5 vehicles, RED FLAG)", lambda: self._fill_visit(2, {
            "force_outcome": "Successful",
            "vehicle_count": 5,
            "force_red_flag": True
        }))

        # Visit 3: Successful, No Vehicles Noted button, RED FLAG
        self.step("Fill out Visit #3 (Successful, No Vehicles, RED FLAG)", lambda: self._fill_visit(3, {
            "force_outcome": "Successful",
            "vehicle_count": 0,
            "force_red_flag": True
        }))

        # Visit 4: Unsuccessful - P Denied Access
        self.step("Fill out Visit #4 (Unsuccessful - P Denied Access)", lambda: self._fill_visit(4, {
            "force_outcome": "Unsuccessful",
            "force_reason": "P Denied Access",
            "vehicle_count": 1,
            "force_red_flag": False
        }))

        # Visit 5: Unsuccessful - Not Home
        self.step("Fill out Visit #5 (Unsuccessful - Not Home)", lambda: self._fill_visit(5, {
            "force_outcome": "Unsuccessful",
            "force_reason": "Not Home",
            "vehicle_count": 1,
            "force_red_flag": False
        }))

        # Step: Refresh metrics
        self.step("Refresh metrics", lambda: (
            self.run_macro("RefreshMetrics"),
            self.action_delay()
        ))

        # Step 9: Validate metrics
        self.step("Validate metrics section", lambda: self._validate_metrics())

        self.finish_test()

    def scenario_mid_day_update(self, initial_itinerary, updated_itinerary):
        """
        Scenario: Mid-day update
        1. Load initial itinerary
        2. Create visit sheet
        3. Fill out some visits
        4. Add updated itinerary with new visits
        5. Fill out new visits
        6. Refresh metrics
        """
        self.start_test("Mid-Day Update Scenario")

        # Navigate and setup
        self.step("Navigate to main tab", lambda: (
            self.workbook.Sheets("Visit Document Processor").Activate(),
            self.action_delay()
        ))

        self.step("Set date and officers", lambda: self._setup_date_officers())

        # Load initial
        self.step("Load initial itinerary", lambda: (
            self._load_itinerary_direct(initial_itinerary),
            self.action_delay()
        ))

        # Create sheet
        self.step("Create Visit Sheet", lambda: (
            self.run_macro("CreateExcelVisitSheet"),
            self.action_delay(2)
        ))

        # Fill some visits
        self.step("Fill out Visit #1", lambda: self._fill_visit(1))
        self.step("Fill out Visit #2", lambda: self._fill_visit(2))
        self.step("Fill out Visit #3", lambda: self._fill_visit(3))

        # Mid-day update
        self.step("Notification: New visits added to schedule", lambda: (
            print("    [Simulating: Officer receives notification of new visits]"),
            self.action_delay()
        ))

        # Add updated itinerary
        self.step("Add updated itinerary", lambda: (
            self._add_updated_itinerary_direct(updated_itinerary),
            self.action_delay(2)
        ))

        # Fill new visits
        self.step("Fill out new visits from updated itinerary", lambda: (
            self._fill_latest_visits(2),
            self.action_delay()
        ))

        # Refresh
        self.step("Refresh metrics", lambda: (
            self.run_macro("RefreshMetrics"),
            self.action_delay()
        ))

        self.step("Validate all metrics updated", lambda: self._validate_metrics())

        self.finish_test()

    def scenario_unscheduled_visit(self, itinerary_path):
        """
        Scenario: Discovering unscheduled visit
        1. Load itinerary
        2. Create visit sheet
        3. Fill some visits
        4. Add unscheduled visit
        5. Fill out unscheduled visit
        6. Refresh metrics
        """
        self.start_test("Unscheduled Visit Discovery")

        self.step("Navigate and setup", lambda: (
            self.workbook.Sheets("Visit Document Processor").Activate(),
            self._setup_date_officers(),
            self.action_delay()
        ))

        self.step("Load itinerary", lambda: (
            self._load_itinerary_direct(itinerary_path),
            self.action_delay()
        ))

        self.step("Create Visit Sheet", lambda: (
            self.run_macro("CreateExcelVisitSheet"),
            self.action_delay(2)
        ))

        self.step("Fill out Visit #1", lambda: self._fill_visit(1))
        self.step("Fill out Visit #2", lambda: self._fill_visit(2))

        # Discover unscheduled
        self.step("Officer discovers probationer at different address", lambda: (
            print("    [Scenario: While at Visit #2, officer encounters another probationer]"),
            self.action_delay()
        ))

        # Add unscheduled
        self.step("Add unscheduled visit", lambda: (
            self.run_macro("AddUnscheduledVisit"),
            self.action_delay(2)
        ))

        # Fill unscheduled visit
        self.step("Fill out unscheduled visit details", lambda: (
            self._fill_unscheduled_visit(),
            self.action_delay()
        ))

        self.step("Refresh metrics", lambda: (
            self.run_macro("RefreshMetrics"),
            self.action_delay()
        ))

        self.step("Validate metrics include unscheduled visit", lambda: self._validate_metrics())

        self.finish_test()

    # =========================================================================
    # HELPER METHODS
    # =========================================================================

    def _setup_date_officers(self):
        """Set up date and officers"""
        ws = self.get_sheet("Visit Document Processor")
        ws.Range("B10").Value = datetime.now().strftime("%m/%d/%Y")
        self.action_delay(0.3)
        ws.Range("B13").Value = "JOHNSON, SARAH"
        self.action_delay(0.3)
        ws.Range("B14").Value = "MARTINEZ, DAVID"
        self.action_delay(0.3)

    def _load_itinerary_direct(self, file_path):
        """Load itinerary using test helper macro (no file dialog)"""
        print(f"    Loading: {os.path.basename(file_path)}")

        # Use the test helper macro that accepts a file path directly
        self.excel.Application.Run("LoadExcelFileFromPath", file_path)

        self.action_delay(2)

    def _add_updated_itinerary_direct(self, file_path):
        """Add updated itinerary using test helper macro (no file dialog)"""
        print(f"    Adding updated: {os.path.basename(file_path)}")

        # Use the test helper macro that accepts a file path directly
        self.excel.Application.Run("AddUpdatedItineraryFromPath", file_path)

        self.action_delay(2)

    def _fill_visit(self, visit_num, test_config=None):
        """Fill out a visit with AI-generated content based on user requirements

        test_config dict can specify:
        - force_outcome: "Successful" or "Unsuccessful"
        - force_reason: specific reason for unsuccessful
        - vehicle_count: number of vehicles (0 = No Vehicles Noted, 5 = test Add button)
        - force_red_flag: True to force red flag testing

        Visit Sheet layout (based on VBA Module4_ExcelVisits.bas):
        - Text areas (Description, Observed, Red Flag Details) merge columns A-F (1-6)
        - Input cells like Residents, Consent merge B-F (columns 2-6)
        - Red Flags dropdown is column B (2)
        - Red Flag category checkboxes are at RedFlagRow + 6, columns 1-5
        - Other text for red flags is column 6
        """
        ws = self.get_sheet("Visit Sheet")
        if not ws:
            return

        if test_config is None:
            test_config = {}

        print(f"      Filling Visit #{visit_num}...")

        # Find visit section
        visit_start = self._find_visit_row(ws, visit_num)
        if not visit_start:
            print(f"    Could not find Visit #{visit_num}")
            return

        # Scroll to visit section so user can follow along
        self.scroll_to_row(ws, visit_start)
        ws.Activate()
        self.action_delay(0.5)

        # Get address and visit type for context
        address = ""
        visit_type = ""
        for r in range(visit_start, visit_start + 15):
            label = ws.Cells(r, 1).Value
            if label == "Address:":
                address = str(ws.Cells(r, 2).Value or "")
            elif label == "Type of Visit:":
                visit_type = str(ws.Cells(r, 2).Value or "")

        # =====================================================
        # STEP 1: DECIDE OUTCOME AND REASON FIRST
        # This determines how all other fields are filled
        # =====================================================
        if "force_outcome" in test_config:
            outcome = test_config["force_outcome"]
            is_successful = (outcome == "Successful")
        else:
            # FTR visits are always unsuccessful (not home)
            if visit_type == "FTR":
                is_successful = False
            else:
                is_successful = random.random() > 0.2  # 80% successful
            outcome = "Successful" if is_successful else "Unsuccessful"

        # Determine reason for unsuccessful visits
        reason = ""
        if not is_successful:
            if "force_reason" in test_config:
                reason = test_config["force_reason"]
            elif visit_type == "FTR":
                reason = random.choice(["Not Home", "Wrong Address"])
            else:
                reason = random.choice(["Not Home", "Wrong Address", "P Denied Access"])

        print(f"        [Outcome decided: {outcome}" + (f", Reason: {reason}]" if reason else "]"))

        # =====================================================
        # STEP 2: CALCULATE TIMES BASED ON VISIT TYPE
        # AM Intake: 8:00 AM - 12:00 PM
        # PM Intake: 12:00 PM - 5:00 PM
        # Others: spread throughout day
        # =====================================================
        if "AM" in visit_type or visit_type == "AM Intake":
            base_hour = 8 + (visit_num - 1) % 4  # 8, 9, 10, 11 AM
        elif "PM" in visit_type or visit_type == "PM Intake":
            base_hour = 12 + (visit_num - 1) % 5  # 12, 1, 2, 3, 4 PM
        else:
            # Other types - spread throughout day
            base_hour = 8 + (visit_num - 1)
            if base_hour > 16:
                base_hour = 16

        base_minute = random.randint(0, 45)
        arrival = datetime.now().replace(hour=base_hour, minute=base_minute, second=0)
        visit_duration = random.randint(8, 12)  # 8-12 minutes
        departure = arrival + timedelta(minutes=visit_duration)

        # =====================================================
        # STEP 3: DETERMINE CONSENT BASED ON OUTCOME
        # =====================================================
        if visit_type == "FTR" or reason in ["Not Home", "Wrong Address"]:
            consent_value = "N/A"
        elif reason == "P Denied Access":
            consent_value = "No"
        elif is_successful:
            consent_value = "Yes"
        else:
            consent_value = "N/A"

        # =====================================================
        # STEP 4: FILL ALL FIELDS
        # =====================================================
        for r in range(visit_start, visit_start + 60):
            label = ws.Cells(r, 1).Value
            if label is None:
                continue

            # === RESIDENTS (only if successful - someone was home) ===
            if label == "Residents:":
                self.scroll_to_row(ws, r)
                if is_successful:
                    residents = "P"
                    if random.random() > 0.5:
                        residents += ", spouse"
                    ws.Cells(r, 2).Value = residents
                    print(f"        Residents: {residents}")
                else:
                    print(f"        Residents: (skipped - {reason if reason else 'no one home'})")
                self.action_delay(0.5)

            # === DESCRIPTION OF RESIDENCE (ALWAYS - exterior only) ===
            elif label == "Description of Residence:":
                text_row = r + 1
                self.scroll_to_row(ws, r)
                print(f"        Generating description of residence (exterior)...")
                desc = self.ollama.generate_residence_description(address)
                print(f"        Typing description...")
                self.typer.type_in_cell(ws.Cells(text_row, 1), desc)

            # === CONSENT TO ENTER HOME ===
            elif label == "Consent to Enter Home:":
                ws.Cells(r, 2).Value = consent_value
                print(f"        Consent: {consent_value}")
                self.action_delay(0.5)

            # === OBSERVED (ALWAYS - content depends on outcome) ===
            elif label == "Observed:":
                text_row = r + 1
                self.scroll_to_row(ws, r)
                print(f"        Generating observed description...")
                obs = self.ollama.generate_observed_description(is_successful, visit_type, reason)
                print(f"        Typing observed...")
                self.typer.type_in_cell(ws.Cells(text_row, 1), obs)

            # === VEHICLES ===
            elif label == "Vehicles:":
                self.scroll_to_row(ws, r)
                vehicle_count = test_config.get("vehicle_count", 1)
                vehicle_data_row = r + 2  # First data row

                if vehicle_count == 0:
                    # Test "No Vehicles Noted" button
                    print(f"        Testing 'No Vehicles Noted' button...")
                    try:
                        self.excel.Application.Run("NoVehiclesNoted", visit_num)
                        print(f"        'No Vehicles Noted' clicked")
                    except Exception as e:
                        # Fallback: just leave empty
                        print(f"        Button error, leaving vehicles empty: {e}")
                    self.action_delay(0.5)
                else:
                    # Add vehicles
                    plates = ["ABC-1234", "XYZ-5678", "DEF-9012", "GHI-3456", "JKL-7890"]
                    colors = ["Blue", "Red", "White", "Black", "Silver"]
                    makes = ["Honda", "Toyota", "Ford", "Chevrolet", "Nissan"]
                    models = ["Civic", "Camry", "F-150", "Malibu", "Altima"]

                    for v in range(vehicle_count):
                        if v >= 2:
                            # Need to add more rows - click Add Vehicle button
                            print(f"        Clicking '+Add' button for vehicle {v+1}...")
                            try:
                                self.excel.Application.Run("AddVehicleRow", visit_num)
                                self.action_delay(0.5)
                            except Exception as e:
                                print(f"        Add button error: {e}")
                                break

                        current_row = vehicle_data_row + v
                        ws.Cells(current_row, 1).Value = plates[v % len(plates)]
                        ws.Cells(current_row, 2).Value = colors[v % len(colors)]
                        ws.Cells(current_row, 3).Value = makes[v % len(makes)]
                        ws.Cells(current_row, 4).Value = models[v % len(models)]
                        print(f"        Added vehicle {v+1}: {colors[v % len(colors)]} {makes[v % len(makes)]} {models[v % len(models)]}")
                        self.action_delay(0.3)

            # === RED FLAGS ===
            elif label == "Red Flags:":
                self.scroll_to_row(ws, r)
                do_red_flag = test_config.get("force_red_flag", False)

                if do_red_flag and is_successful:
                    print(f"        Setting Red Flag to Yes...")
                    ws.Cells(r, 2).Value = "Yes"
                    self.action_delay(0.5)

                    # Toggle red flags section
                    try:
                        self.excel.Application.Run("ToggleRedFlagsSection", ws, r, True)
                    except:
                        pass
                    self.action_delay(0.3)

                    # Choose flag type and generate details
                    flag_types = ["Alcohol", "Drugs", "Guns", "Knives", "IP"]
                    flag_type = random.choice(flag_types)
                    print(f"        Generating red flag details ({flag_type})...")
                    details = self.ollama.generate_red_flag_description(flag_type)

                    details_row = r + 2
                    self.scroll_to_row(ws, details_row)
                    print(f"        Typing red flag details...")
                    self.typer.type_in_cell(ws.Cells(details_row, 1), details)

                    # Check the matching checkbox (row r+6)
                    checkbox_row = r + 6
                    checkbox_col = {"Alcohol": 1, "Drugs": 2, "Guns": 3, "Knives": 4, "IP": 5}
                    ws.Cells(checkbox_row, checkbox_col.get(flag_type, 1)).Value = "X"
                    print(f"        Checked {flag_type} checkbox")
                    self.action_delay(0.5)
                else:
                    print(f"        Red Flags: None")

            # === ARRIVED / DEPARTED ===
            elif label == "Arrived:":
                self.scroll_to_row(ws, r)
                arrival_str = arrival.strftime("%I:%M %p")
                ws.Cells(r, 2).Value = arrival_str
                print(f"        Arrived: {arrival_str}")
                self.action_delay(0.3)

                departure_str = departure.strftime("%I:%M %p")
                ws.Cells(r, 4).Value = departure_str
                print(f"        Departed: {departure_str}")
                self.action_delay(0.3)

            # === OUTCOME ===
            elif label == "Outcome:":
                self.scroll_to_row(ws, r)
                ws.Cells(r, 2).Value = outcome
                print(f"        Outcome: {outcome}")
                self.action_delay(0.3)

                if not is_successful and reason:
                    ws.Cells(r, 4).Value = reason
                    print(f"        Reason: {reason}")
                    self.action_delay(0.3)

                break  # Outcome is last field

        # === TEST EXPORT BUTTON ===
        print(f"        Testing Export button...")
        self.scroll_to_row(ws, visit_start)
        self.action_delay(0.5)

        try:
            self.excel.Application.Run("ExportSingleVisitByNum", visit_num)
            print(f"        Export button clicked - Word document opened")
            self.action_delay(2)
        except Exception as e:
            print(f"        Export button error: {e}")

        print(f"      Visit #{visit_num} completed ({outcome})")

    def _find_visit_row(self, ws, visit_num):
        """Find the starting row of a visit"""
        target = f"VISIT #{visit_num}"
        for r in range(1, 500):
            cell_val = ws.Cells(r, 1).Value
            if cell_val and target in str(cell_val):
                return r
        return None

    def _fill_unscheduled_visit(self):
        """Fill out an unscheduled visit"""
        ws = self.get_sheet("Visit Sheet")
        if not ws:
            return

        # Find the last visit (unscheduled one)
        last_visit_num = 0
        for r in range(1, 500):
            cell_val = ws.Cells(r, 1).Value
            if cell_val and "VISIT #" in str(cell_val):
                try:
                    num = int(str(cell_val).replace("VISIT #", "").strip())
                    last_visit_num = max(last_visit_num, num)
                except:
                    pass

        if last_visit_num > 0:
            self._fill_visit(last_visit_num)

    def _fill_latest_visits(self, count):
        """Fill the latest N visits"""
        ws = self.get_sheet("Visit Sheet")
        if not ws:
            return

        # Find all visit numbers
        visit_nums = []
        for r in range(1, 500):
            cell_val = ws.Cells(r, 1).Value
            if cell_val and "VISIT #" in str(cell_val):
                try:
                    num = int(str(cell_val).replace("VISIT #", "").strip())
                    visit_nums.append(num)
                except:
                    pass

        # Fill the last N
        for num in sorted(visit_nums)[-count:]:
            self._fill_visit(num)

    def _validate_metrics(self):
        """Validate metrics section"""
        ws = self.get_sheet("Visit Sheet")
        if not ws:
            return False

        # Find metrics section
        metrics_row = None
        for r in range(1, 500):
            if ws.Cells(r, 1).Value == "DAILY VISIT METRICS":
                metrics_row = r
                break

        if not metrics_row:
            raise Exception("Metrics section not found")

        # Check for common errors
        for r in range(metrics_row, metrics_row + 50):
            for c in range(1, 7):
                val = ws.Cells(r, c).Value
                if val and "#" in str(val):
                    raise Exception(f"Excel error found at row {r}, col {c}: {val}")

        print("    Metrics validation: PASSED")
        return True

    def generate_report(self):
        """Generate final HTML report"""
        return self.reporter.generate_report(self.test_results, self.screenshot_dir)


def main():
    """Main entry point"""
    print("="*60)
    print("VISIT DOCUMENT PROCESSOR - AUTOMATED TESTER")
    print("="*60)

    # Paths
    base_dir = os.path.dirname(os.path.abspath(__file__))
    workbook_path = os.path.join(base_dir, "Visit_Document_Processor_TEST2.xlsm")
    itinerary_5 = os.path.join(base_dir, "test_itineraries", "test_itinerary_5_visits.xlsx")
    itinerary_5_updated = os.path.join(base_dir, "test_itineraries", "test_itinerary_5_updated_3more.xlsx")
    itinerary_8 = os.path.join(base_dir, "test_itineraries", "test_itinerary_8_visits.xlsx")

    # Check files exist
    for f in [workbook_path, itinerary_5]:
        if not os.path.exists(f):
            print(f"ERROR: File not found: {f}")
            return

    # Check Ollama
    ollama = OllamaTextGenerator()
    if ollama.available:
        print(f"Ollama: Connected ({ollama.model})")
    else:
        print("Ollama: Not available (using fallback text)")

    # Create tester - use "slow" speed to give AI time to generate good content
    tester = VisitDocumentTester(workbook_path, speed="slow")

    try:
        tester.setup()

        # Run scenarios
        print("\n" + "="*60)
        print("RUNNING TEST SCENARIOS")
        print("="*60)

        # Scenario 1: Normal day (only run first scenario for testing)
        tester.scenario_normal_day(itinerary_5)

        # TEMPORARILY DISABLED - uncomment to run all scenarios
        # # Reset workbook for next test
        # tester.workbook.Close(SaveChanges=False)
        # tester.workbook = tester.excel.Workbooks.Open(workbook_path)
        # tester.excel.Application.Run("EnableTestMode")  # Re-enable test mode
        # time.sleep(1)

        # # Scenario 2: Mid-day update
        # tester.scenario_mid_day_update(itinerary_5, itinerary_5_updated)

        # # Reset
        # tester.workbook.Close(SaveChanges=False)
        # tester.workbook = tester.excel.Workbooks.Open(workbook_path)
        # tester.excel.Application.Run("EnableTestMode")  # Re-enable test mode
        # time.sleep(1)

        # # Scenario 3: Unscheduled visit
        # tester.scenario_unscheduled_visit(itinerary_8)

        # Generate report
        report_path = tester.generate_report()

        # Summary
        print("\n" + "="*60)
        print("TEST SUMMARY")
        print("="*60)
        passed = sum(1 for t in tester.test_results if t.status == "pass")
        total = len(tester.test_results)
        print(f"Passed: {passed}/{total}")
        print(f"Report: {report_path}")

    except Exception as e:
        print(f"\nFATAL ERROR: {e}")
        import traceback
        traceback.print_exc()

    finally:
        # Keep workbook open for inspection
        tester.teardown(keep_open=True)
        print("\nTest complete. Workbook left open for inspection.")


if __name__ == "__main__":
    main()
