// Handles Excel import, table rendering, test execution, screenshot paste, and report generation
let testCases = [];
let columnMapping = null;
const defaultHeaders = [
    'Test Case ID', 'Name', 'Description', 'Preconditions', 'Priority', 'Type', 'Status'
];
const stepHeaders = [
    'Step', 'Expected Result', 'Actual Result', 'Screenshot', 'Step Result'
];

// Default mapping for common Qtest export
const defaultMapping = {
    'Test Case ID': 'Test Case ID',
    'Name': 'Name',
    'Description': 'Description',
    'Preconditions': 'Preconditions',
    'Priority': 'Priority',
    'Type': 'Type',
    'Status': 'Status',
    'Test Steps': 'Test Steps',
    'Expected Results': 'Expected Results'
};

const excelFile = document.getElementById('excelFile');
const importBtn = document.getElementById('importBtn');
const tableContainer = document.getElementById('tableContainer');
const generateReportBtn = document.getElementById('generateReportBtn');
const resetBtn = document.getElementById('resetBtn');

/**
 * Utility: Get mapping from localStorage or default
 */
function getCurrentMapping(excelHeaders) {
    const savedMapping = localStorage.getItem('testToolColumnMapping');
    if (savedMapping) {
        try {
            const parsed = JSON.parse(savedMapping);
            const valid = Object.entries(parsed).every(([key, col]) => col && excelHeaders.includes(col));
            if (valid) return parsed;
        } catch (e) { /* ignore */ }
    }
    const validDefault = Object.entries(defaultMapping).every(([key, col]) => excelHeaders.includes(col));
    if (validDefault) return defaultMapping;
    return null;
}

/**
 * Show the column mapping UI for the user to map Excel columns to required fields.
 * @param {string[]} headersFromExcel - Array of column headers from the Excel file.
 */
function showMappingUI(headersFromExcel) {
    const requiredFields = [
        { label: 'Test Case ID', key: 'Test Case ID' },
        { label: 'Name', key: 'Name' },
        { label: 'Description', key: 'Description' },
        { label: 'Preconditions', key: 'Preconditions' },
        { label: 'Priority', key: 'Priority' },
        { label: 'Type', key: 'Type' },
        { label: 'Status', key: 'Status' },
        { label: 'Test Steps', key: 'Test Steps' },
        { label: 'Expected Results', key: 'Expected Results' }
    ];
    let html = '<div id="mappingUI"><h3>Map Excel Columns</h3><table>';
    requiredFields.forEach(field => {
        html += `<tr><td>${field.label}</td><td><select id="map_${field.key.replace(/\s/g,'_')}">`;
        html += '<option value="">-- None --</option>';
        headersFromExcel.forEach(h => {
            html += `<option value="${h}">${h}</option>`;
        });
        html += '</select></td></tr>';
    });
    html += '</table><button id="applyMappingBtn">Apply Mapping</button></div>';
    tableContainer.innerHTML = html;
    document.getElementById('applyMappingBtn').onclick = () => {
        columnMapping = {};
        requiredFields.forEach(field => {
            const val = document.getElementById('map_' + field.key.replace(/\s/g,'_')).value;
            columnMapping[field.key] = val;
        });
        localStorage.setItem('testToolColumnMapping', JSON.stringify(columnMapping));
        parseExcelWithMapping();
    };
}

/**
 * Show the Admin/Settings panel for column mapping and customization.
 * @param {string[]} headersFromExcel - Array of headers from Excel file.
 */
function showAdminPanel(headersFromExcel) {
    // Remove any existing admin panel
    const oldPanel = document.getElementById('adminPanel');
    if (oldPanel) oldPanel.remove();
    let mapping = columnMapping ? { ...columnMapping } : {};
    let customColumns = Object.keys(mapping);
    let html = '<div id="adminPanel"><h3>Admin Settings</h3>';
    html += '<h4>Column Mapping (add/remove any columns)</h4>';
    html += '<table id="mappingTable">';
    customColumns.forEach((col, idx) => {
        html += `<tr><td><input type="text" value="${col}" class="colKeyInput" data-idx="${idx}"></td><td><select class="colValSelect" data-idx="${idx}">`;
        html += '<option value="">-- None --</option>';
        headersFromExcel.forEach(h => {
            html += `<option value="${h}"${mapping[col]===h?' selected':''}>${h}</option>`;
        });
        html += `</select></td><td><button class="removeColBtn" data-idx="${idx}">Remove</button></td></tr>`;
    });
    html += '</table>';
    html += '<button id="addColBtn">Add Column</button>';
    html += '<button id="saveMappingBtn">Save Mapping</button>';
    html += '<button id="editRawMappingBtn">Edit Mapping JSON</button>';
    html += '<button id="closeAdminBtn">Close</button>';
    html += '<hr><h4>Customization</h4>';
    html += '<label><input type="checkbox" id="showStepNumbers"> Show Step Numbers</label><br>';
    html += '<label><input type="checkbox" id="compactView"> Compact Table View</label><br>';
    html += '</div>';
    const adminDiv = document.createElement('div');
    adminDiv.innerHTML = html;
    document.body.appendChild(adminDiv);

    // Add column
    document.getElementById('addColBtn').onclick = () => {
        const keys = Array.from(adminDiv.querySelectorAll('.colKeyInput')).map(i=>i.value.trim());
        const vals = Array.from(adminDiv.querySelectorAll('.colValSelect')).map(s=>s.value);
        let newMapping = {};
        keys.forEach((k,i)=>{if(k)newMapping[k]=vals[i];});
        newMapping[''] = '';
        columnMapping = newMapping;
        adminDiv.remove();
        showAdminPanel(headersFromExcel);
    };
    // Remove column
    adminDiv.querySelectorAll('.removeColBtn').forEach(btn => {
        btn.onclick = () => {
            const idx = btn.getAttribute('data-idx');
            const keys = Array.from(adminDiv.querySelectorAll('.colKeyInput')).map(i=>i.value.trim());
            const vals = Array.from(adminDiv.querySelectorAll('.colValSelect')).map(s=>s.value);
            let newMapping = {};
            keys.forEach((k,i)=>{if(k)newMapping[k]=vals[i];});
            const key = keys[idx];
            delete newMapping[key];
            columnMapping = newMapping;
            adminDiv.remove();
            showAdminPanel(headersFromExcel);
        };
    });
    // Save mapping
    document.getElementById('saveMappingBtn').onclick = () => {
        const keys = Array.from(adminDiv.querySelectorAll('.colKeyInput')).map(i=>i.value.trim());
        const vals = Array.from(adminDiv.querySelectorAll('.colValSelect')).map(s=>s.value);
        let newMapping = {};
        keys.forEach((k,i)=>{if(k)newMapping[k]=vals[i];});
        columnMapping = newMapping;
        localStorage.setItem('testToolColumnMapping', JSON.stringify(columnMapping));
        alert('Mapping saved!');
        adminDiv.remove();
    };
    // Edit raw mapping JSON
    document.getElementById('editRawMappingBtn').onclick = () => {
        let raw = prompt('Edit mapping JSON:', JSON.stringify(columnMapping, null, 2));
        if (raw) {
            try {
                columnMapping = JSON.parse(raw);
                localStorage.setItem('testToolColumnMapping', JSON.stringify(columnMapping));
                alert('Mapping updated!');
                adminDiv.remove();
            } catch(e) { alert('Invalid JSON!'); }
        }
    };
    // Customization options
    document.getElementById('showStepNumbers').checked = localStorage.getItem('showStepNumbers') === 'true';
    document.getElementById('showStepNumbers').onchange = function() {
        localStorage.setItem('showStepNumbers', this.checked);
        renderTable();
    };
    document.getElementById('compactView').checked = localStorage.getItem('compactView') === 'true';
    document.getElementById('compactView').onchange = function() {
        localStorage.setItem('compactView', this.checked);
        renderTable();
    };
    document.getElementById('closeAdminBtn').onclick = () => {
        adminDiv.remove();
    };
}

/**
 * Render the main test case table with dynamic columns and steps.
 */
function renderTable() {
    if (!columnMapping) return;
    const showStepNumbers = localStorage.getItem('showStepNumbers') === 'true';
    const compactView = localStorage.getItem('compactView') === 'true';
    const dynamicHeaders = Object.keys(columnMapping || {}).filter(k => !/step/i.test(k));
    let html = `<table${compactView ? ' class="compactView"' : ''}><thead><tr>`;
    dynamicHeaders.forEach(h => html += `<th>${h}</th>`);
    html += '<th>Test Steps</th></tr></thead><tbody>';
    testCases.forEach((tc, i) => {
        let overall = 'passed';
        if (tc.steps.some(s => s.result === 'failed')) overall = 'failed';
        else if (tc.steps.some(s => s.result === 'skipped')) overall = 'skipped';
        html += `<tr data-idx="${i}" class="${overall}">`;
        dynamicHeaders.forEach(h => {
            html += `<td>${tc[h]||''}</td>`;
        });
        // Steps as sub-table
        html += `<td><table class='steps-table${compactView ? ' compactView' : ''}'><thead><tr>`;
        if (showStepNumbers) html += '<th>#</th>';
        stepHeaders.forEach(sh => html += `<th>${sh}</th>`);
        html += '</tr></thead><tbody>';
        tc.steps.forEach((step, si) => {
            let stepClass = step.result ? step.result : '';
            html += `<tr data-stepidx="${si}" class="${stepClass}">`;
            if (showStepNumbers) html += `<td>${si+1}</td>`;
            html += `<td>${step.step}</td>`;
            html += `<td>${step.expected}</td>`;
            html += `<td><input type="text" value="${step.actual}" onchange="updateStepActual(${i},${si}, this.value)"></td>`;
            html += `<td><div class="screenshot-cell" data-idx="${i}" data-stepidx="${si}" tabindex="0" style="min-width:110px;min-height:30px;background:#fafafa;border:1px dashed #ccc;text-align:center;cursor:pointer;" onpaste="handleStepPaste(event,${i},${si})">${step.screenshot ? `<img src="${step.screenshot}" class="screenshot-img">` : '<span style="color:#bbb;">Paste Image</span>'}</div></td>`;
            html += `<td><select onchange="updateStepResult(${i},${si}, this.value)">
                <option value="">Select</option>
                <option value="passed" ${step.result==='passed'?'selected':''}>Passed</option>
                <option value="failed" ${step.result==='failed'?'selected':''}>Failed</option>
                <option value="skipped" ${step.result==='skipped'?'selected':''}>Skipped</option>
            </select></td>`;
            html += '</tr>';
        });
        html += '</tbody></table></td>';
        html += '</tr>';
    });
    html += '</tbody></table>';
    tableContainer.innerHTML = html;
    updateRowColors();
}

/**
 * Import Excel file and handle sheet selection if multiple sheets exist.
 */
importBtn.onclick = () => {
    if (!excelFile.files[0]) return alert('Please select an Excel file.');
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetNames = workbook.SheetNames;
            if (sheetNames.length > 1) {
                let sheetHtml = '<div id="sheetSelectPanel"><h3>Select Sheet</h3>';
                sheetNames.forEach((name) => {
                    sheetHtml += `<button class="sheetBtn" data-sheet="${name}">${name}</button> `;
                });
                sheetHtml += '</div>';
                tableContainer.innerHTML = sheetHtml;
                document.querySelectorAll('.sheetBtn').forEach(btn => {
                    btn.onclick = () => {
                        document.getElementById('sheetSelectPanel').remove();
                        importExcelWithSheet(workbook, btn.getAttribute('data-sheet'));
                    };
                });
            } else {
                importExcelWithSheet(workbook, sheetNames[0]);
            }
        } catch (err) {
            alert('Failed to read Excel file. Please check the file format.');
        }
    };
    reader.readAsArrayBuffer(excelFile.files[0]);
};

function importExcelWithSheet(workbook, sheetName) {
    const sheet = workbook.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    if (json.length === 0) return alert('No data found in Excel.');
    const excelHeaders = Object.keys(json[0]);
    columnMapping = getCurrentMapping(excelHeaders);
    // If mapping is missing or incomplete, always show Admin panel for mapping
    if (!columnMapping || Object.keys(columnMapping).length === 0 || Object.values(columnMapping).some(v => !v || !excelHeaders.includes(v))) {
        showAdminPanel(excelHeaders);
        resetBtn.style.display = 'inline-block';
        return;
    }
    parseExcelWithMappingFromJson(json);
    resetBtn.style.display = 'inline-block';
}

/**
 * Parse Excel rows using current mapping
 * @param {object[]} json - Array of row objects from Excel.
 */
function parseExcelWithMappingFromJson(json) {
    testCases = json.map(row => {
        let stepObjs = [];
        let stepCol = Object.keys(columnMapping).find(k => /step/i.test(k));
        let stepsRaw = stepCol ? row[columnMapping[stepCol]] || '' : '';
        if (stepsRaw) {
            // Enhanced step parsing for numbered steps and expected results
            const stepBlocks = stepsRaw.split(/(?=\n?\d+\))/g).map(s => s.trim()).filter(Boolean);
            for (const block of stepBlocks) {
                const lines = block.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
                let stepText = '';
                let expected = '';
                for (const line of lines) {
                    if (/^Expected result[:：]/i.test(line)) {
                        expected += (expected ? ' ' : '') + line.replace(/^Expected result[:：]/i, '').trim();
                    } else {
                        stepText += (stepText ? ' ' : '') + line;
                    }
                }
                if (stepText) {
                    stepObjs.push({
                        step: stepText,
                        expected: expected,
                        actual: '',
                        screenshot: '',
                        result: ''
                    });
                }
            }
        }
        let tc = {};
        Object.keys(columnMapping).forEach(col => {
            if (!/step/i.test(col)) tc[col] = row[columnMapping[col]] || '';
        });
        tc.steps = stepObjs;
        return tc;
    });
    renderTable();
    generateReportBtn.style.display = 'inline-block';
}

// Ensure Admin/Settings button is always visible and functional
function ensureAdminButton() {
    if (!document.getElementById('adminBtn')) {
        const adminBtn = document.createElement('button');
        adminBtn.id = 'adminBtn';
        adminBtn.textContent = 'Admin/Settings';
        adminBtn.style = 'position:fixed;top:10px;right:10px;z-index:1000;';
        document.body.appendChild(adminBtn);
        adminBtn.onclick = () => {
            let headersFromExcel = [];
            if (testCases.length > 0) {
                headersFromExcel = Object.keys(testCases[0]).filter(k => k !== 'steps');
            } else if (columnMapping) {
                headersFromExcel = Object.values(columnMapping);
            }
            if (!headersFromExcel.length) headersFromExcel = [''];
            showAdminPanel(headersFromExcel);
        };
    }
}
document.addEventListener('DOMContentLoaded', ensureAdminButton);

/**
 * Update the result of a test step and re-render the table.
 */
function updateStepResult(tcIdx, stepIdx, value) {
    testCases[tcIdx].steps[stepIdx].result = value;
    renderTable();
}

/**
 * Update the actual result of a test step and re-render the table.
 */
function updateStepActual(tcIdx, stepIdx, value) {
    testCases[tcIdx].steps[stepIdx].actual = value;
    renderTable();
}

/**
 * Update row colors based on test case result (passed/failed/skipped).
 */
function updateRowColors() {
    const rows = tableContainer.querySelectorAll('tbody > tr');
    rows.forEach((row, i) => {
        if (!testCases[i]) return;
        let overall = 'passed';
        if (testCases[i].steps.some(s => s.result === 'failed')) overall = 'failed';
        else if (testCases[i].steps.some(s => s.result === 'skipped')) overall = 'skipped';
        row.className = overall;
    });
}

/**
 * Generate a single-file HTML report of the current test execution.
 */
generateReportBtn.onclick = function() {
    const html = generateReportHTML();
    const blob = new Blob([html], {type: 'text/html'});
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'TestExecutionReport.html';
    a.click();
};

resetBtn.onclick = () => {
    testCases = [];
    tableContainer.innerHTML = '';
    excelFile.value = '';
    generateReportBtn.style.display = 'none';
    resetBtn.style.display = 'none';
};
function generateReportHTML() {
    // Get headers from current columnMapping (localStorage JSON)
    let mapping = columnMapping || {};
    let headers = Object.keys(mapping).filter(h => mapping[h]);
    let style = `<style>body{font-family:Arial,sans-serif;}table{border-collapse:collapse;width:100%;}th,td{border:1px solid #ccc;padding:8px;}th{background:#f0f0f0;}input,select{width:100%;}img{max-width:100px;max-height:60px;display:block;}.passed{background:#d4edda;}.failed{background:#f8d7da;}.skipped{background:#fff3cd;} .steps-table{margin:8px 0;} .steps-table th, .steps-table td{font-size:13px;}</style>`;
    let script = `<script>document.querySelectorAll('img').forEach(img=>{img.onclick=()=>{let w=window.open();w.document.write('<img src="'+img.src+'" style="max-width:100%">');};});<\/script>`;
    let html = `<html><head><meta charset='UTF-8'><title>Test Execution Report</title>${style}</head><body><h2>Test Execution Report</h2><table><thead><tr>`;
    headers.forEach(h=>{html+=`<th>${h}</th>`;});
    html+='<th>Test Steps</th></tr></thead><tbody>';
    testCases.forEach(tc=>{
        let overall = 'passed';
        if (tc.steps.some(s => s.result === 'failed')) overall = 'failed';
        else if (tc.steps.some(s => s.result === 'skipped')) overall = 'skipped';
        html+=`<tr class="${overall}">`;
        headers.forEach(h=>{html+=`<td>${tc[h]||''}</td>`;});
        html += `<td><table class='steps-table'><thead><tr>`;
        stepHeaders.forEach(sh => html += `<th>${sh}</th>`);
        html += '</tr></thead><tbody>';
        tc.steps.forEach(step => {
            html += '<tr>';
            html += `<td>${step.step}</td>`;
            html += `<td>${step.expected}</td>`;
            html += `<td>${step.actual||''}</td>`;
            html += `<td>${step.screenshot?`<img src="${step.screenshot}" style="max-width:100px;max-height:60px;">`:''}</td>`;
            html += `<td>${step.result||''}</td>`;
            html += '</tr>';
        });
        html += '</tbody></table></td>';
        html += '</tr>';
    });
    html+='</tbody></table>'+script+'</body></html>';
    return html;
}
/**
 * Handle image paste event for a step's screenshot cell.
 */
function handleStepPaste(event, tcIdx, stepIdx) {
    event.preventDefault();
    const items = (event.clipboardData || event.originalEvent.clipboardData).items;
    for (let i = 0; i < items.length; i++) {
        if (items[i].type.indexOf('image') !== -1) {
            const file = items[i].getAsFile();
            const reader = new FileReader();
            reader.onload = function(e) {
                testCases[tcIdx].steps[stepIdx].screenshot = e.target.result;
                renderTable();
            };
            reader.readAsDataURL(file);
            break;
        }
    }
}
