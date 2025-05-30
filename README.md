# TestEase Manual Execution Tool

A web-based tool for manual test execution using only HTML, CSS, and JavaScript (no frameworks or build tools).

## Features
- **Import Qtest or any Excel file**: Supports multi-sheet selection and flexible step parsing.
- **Dynamic Column Mapping**: Map any Excel columns to required fields after import. Mapping is persistent via localStorage.
- **Admin/Settings Panel**: Add/remove columns, edit mapping JSON, and customize table view (step numbers, compact mode).
- **Test Case & Step Table**: Displays all mapped columns and test steps. Supports editing actual results, setting step results, and pasting screenshots directly into steps.
- **Default Status**: All test cases and steps are set to "Not Executed" by default until updated.
- **Screenshot Pasting**: Paste images from clipboard directly into step screenshot cells.
- **HTML Report Generation**: Generates a single-file HTML report reflecting all mapped columns, step results, and screenshots.
- **Reset**: Clear all imported data and UI.

## Usage
1. **Open `index.html` in your browser.**
2. **Import Excel File**: Click the file input, select your Excel file, and choose the sheet if prompted.
3. **Map Columns**: Use the mapping UI or Admin/Settings panel to map Excel columns to required fields. Mapping is saved for future imports.
4. **Review/Edit Test Cases**: The table displays all test cases and steps. Edit actual results, set step results (Passed/Failed/Skipped/Not Executed), and paste screenshots as needed.
5. **Admin/Settings**: Click the Admin/Settings button (top-right) to add/remove columns, edit mapping JSON, or change table view options.
6. **Generate Report**: Click "Generate Report" to download a single-file HTML report with all data and screenshots.
7. **Reset**: Click "Reset" to clear all data and start over.

## Notes
- All data and mapping are stored in your browser (localStorage). No data is sent to a server.
- Supports flexible step parsing for multi-line, numbered steps with inline expected results.
- The tool is fully client-side and requires no installation or backend.

## Customization
- Use the Admin/Settings panel to:
  - Add/remove any number of columns for mapping.
  - Edit the mapping JSON directly.
  - Toggle step numbers and compact table view.

---

For any issues or feature requests, please open an issue or contact the maintainer.
