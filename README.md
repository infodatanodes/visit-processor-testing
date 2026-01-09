# Visit Document Processor - Automated Testing Framework

An intelligent, visual test automation suite for the Visit Document Processor Excel application.

## Features

- **Visual Execution**: Watch tests run in real-time with visible typing
- **AI-Powered Text**: Uses Ollama LLM to generate realistic visit descriptions
- **Multiple Scenarios**: Normal day, mid-day updates, unscheduled visits
- **Screenshot Capture**: Automatic screenshots at each test step
- **HTML Reports**: Beautiful test reports with pass/fail status
- **Real DFW Addresses**: Test data uses actual Dallas-Fort Worth area addresses

## Requirements

### Python Packages
```bash
pip install pywin32 openpyxl pillow mss requests
```

### Ollama (Optional but Recommended)
For AI-generated text descriptions:
1. Install Ollama: https://ollama.ai
2. Pull a model: `ollama pull llama3:8b`
3. Ollama runs automatically on `http://localhost:11434`

If Ollama is not available, the tester uses pre-written fallback text.

## Folder Structure

```
Testing/
├── Visit_Document_Processor_TEST.xlsm  # Test copy of workbook
├── automated_tester.py                  # Main test runner
├── generate_test_itinerary.py           # Creates test data
├── README.md                            # This file
├── test_itineraries/                    # Generated test files
│   ├── test_itinerary_5_visits.xlsx
│   ├── test_itinerary_5_updated_3more.xlsx
│   ├── test_itinerary_8_visits.xlsx
│   └── ...
├── screenshots/                         # Test screenshots
│   └── YYYYMMDD_HHMMSS/
└── test_results/                        # HTML reports
    └── test_report_YYYYMMDD_HHMMSS.html
```

## Quick Start

### 1. Generate Test Itineraries
```bash
cd Testing
python generate_test_itinerary.py
```

This creates test itineraries with:
- 5, 8, 12, and 18 visit variants
- Updated versions with additional visits
- Real DFW addresses
- Randomized visit types (AM Intake, PM Intake, FTR, HR, CM)

### 2. Run Automated Tests
```bash
python automated_tester.py
```

The tester will:
1. Open Excel with the test workbook
2. Run through test scenarios visually
3. Fill in fields with AI-generated or fallback text
4. Take screenshots at each step
5. Generate an HTML report

### 3. View Results
Open the HTML report in `test_results/` to see:
- Pass/fail status for each scenario
- Detailed steps with timestamps
- Screenshots from execution

## Test Scenarios

### 1. Normal Day Workflow
- Load itinerary
- Create visit sheet
- Fill out visits with times and outcomes
- Validate metrics

### 2. Mid-Day Update
- Start with initial itinerary
- Fill out some visits
- Add updated itinerary with new visits
- Fill out new visits
- Validate all metrics update correctly

### 3. Unscheduled Visit Discovery
- Load itinerary
- Fill out scheduled visits
- Add unscheduled visit (discovered in field)
- Fill out unscheduled visit
- Validate metrics include new visit

## Configuration

### Speed Settings
In `automated_tester.py`, change the speed parameter:

```python
tester = VisitDocumentTester(workbook_path, speed="normal")
```

Options:
- `"slow"` - 2 second delays, 0.05s per character (best for demos)
- `"normal"` - 1 second delays, 0.03s per character (default)
- `"fast"` - 0.5 second delays, 0.01s per character (quick runs)

### Ollama Model
Change the model in `OllamaTextGenerator`:

```python
self.ollama = OllamaTextGenerator(model="mistral:7b")  # or llama3:8b
```

## Extending the Framework

### Add New Scenarios
Create a new method in `VisitDocumentTester`:

```python
def scenario_corrections(self, itinerary_path):
    """Test correcting mistakes"""
    self.start_test("Correction Scenario")

    # Your test steps here
    self.step("Description", lambda: your_action())

    self.finish_test()
```

### Add New Test Data
Edit `generate_test_itinerary.py` to add:
- More addresses in `DFW_ADDRESSES`
- More names in `FIRST_NAMES` / `LAST_NAMES`
- Different visit type distributions

## Troubleshooting

### Excel not opening
- Ensure no other Excel instances are running
- Run as Administrator if needed
- Check that macro settings allow VBA execution

### Ollama not connecting
- Verify Ollama is running: `ollama list`
- Check it's on default port: `http://localhost:11434`
- Test manually: `curl http://localhost:11434/api/tags`

### Screenshots not working
- Install mss: `pip install mss`
- May need to run as Administrator for screen capture

## Changelog

### v1.1.0 (2026-01-08)
**Major improvements to AI-generated content and test coverage:**

- **Timing Logic**: AM Intake visits now use 8:00 AM - 12:00 PM times, PM Intake uses 12:00 PM - 5:00 PM
- **Description of Residence**: Now generates EXTERIOR-only descriptions including:
  - House type (single/two story, apartment)
  - Exterior material and trim colors
  - Garage type (one-car, two-car, carport)
  - Front yard condition
- **Observed Field**: Detailed room-by-room walkthrough for successful visits:
  - Who answered the door
  - Consent request and response
  - Home entry description ("home opens to a...")
  - Bedroom location and description
  - Kitchen walkthrough
  - Refrigerator description (color, type, contents)
  - Exit and violations status
- **Consent Logic**: Now properly matches outcome:
  - "Yes" only for successful visits
  - "No" for P Denied Access
  - "N/A" for Not Home, Wrong Address, FTR
- **Vehicle Testing**:
  - Test +Add button with 5 vehicles
  - Test 2 vehicles (default)
  - Test "No Vehicles Noted" button
- **Red Flags Testing**: Forced testing on 2 visits with detailed descriptions including location and specifics
- **All 5 Visits**: Now fills all visits in the test scenario
- **Workbook Left Open**: Test completes but leaves workbook open for inspection

### v1.0.0 (2026-01-08)
- Initial release
- Visual test execution with character-by-character typing
- Ollama LLM integration for AI-generated text
- Screenshot capture and HTML reporting
- Test mode in VBA to suppress dialogs
- Three test scenarios: Normal Day, Mid-Day Update, Unscheduled Visit

## Future Enhancements

- [ ] Video recording of test execution
- [ ] Parallel test execution
- [ ] Test data variations (error cases)
- [ ] Integration with CI/CD
- [ ] Performance benchmarking
