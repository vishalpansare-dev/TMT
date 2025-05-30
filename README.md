# Manual Test Execution Tool

This project is a simple web-based tool for manual test execution. It allows users to import Qtest-exported Excel files, view and execute test cases, add actual results, paste screenshots, and generate a detailed HTML report. Built with only HTML, CSS, and JavaScript (no frameworks).

## Features
- Import Qtest Excel file and display test cases in a table
- Add Actual Result, Screenshot (paste from clipboard), and Test Step Result (passed/failed/skipped)
- Auto-calculate overall test result based on step results
- Generate a single-file HTML report with inline CSS and JS

## Usage
1. Open `index.html` in your browser.
2. Import your Qtest Excel file.
3. Execute tests, add results, screenshots, and step status.
4. Click 'Generate Report' to export a detailed HTML report.

## Project Structure
- `index.html` - Main UI
- `style.css` - Styles
- `script.js` - Functionality

## Requirements
- No dependencies or build tools required. Just open in a browser.
