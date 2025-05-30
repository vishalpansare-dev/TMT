// Handles Excel import, table rendering, test execution, screenshot paste, and report generation
let testCases = [];
let columnMapping = null;
const headers = [
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

function showMappingUI(headersFromExcel) {
    // Required fields for mapping
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
        // Save mapping to localStorage
        localStorage.setItem('testToolColumnMapping', JSON.stringify(columnMapping));
        parseExcelWithMapping();
    };
}

function getCurrentMapping(excelHeaders) {
    // Try localStorage, else use default if all fields exist
    const savedMapping = localStorage.getItem('testToolColumnMapping');
    if (savedMapping) {
        const parsed = JSON.parse(savedMapping);
        const valid = Object.entries(parsed).every(([key, col]) => col && excelHeaders.includes(col));
        if (valid) return parsed;
    }
    // Use default if all fields exist
    const validDefault = Object.entries(defaultMapping).every(([key, col]) => excelHeaders.includes(col));
    if (validDefault) return defaultMapping;
    return null;
}

// Admin/settings panel
function showAdminPanel(headersFromExcel) {
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
    let html = '<div id="adminPanel"><h3>Admin Settings</h3>';
    html += '<h4>Column Mapping</h4><table>';
    requiredFields.forEach(field => {
        html += `<tr><td>${field.label}</td><td><select id="map_${field.key.replace(/\s/g,'_')}">`;
        html += '<option value="">-- None --</option>';
        headersFromExcel.forEach(h => {
            html += `<option value="${h}">${h}</option>`;
        });
        // Preselect current mapping
        const current = columnMapping && columnMapping[field.key] ? columnMapping[field.key] : (defaultMapping[field.key] || '');
        html = html.replace(`<option value="${current}">`, `<option value="${current}" selected>`);
        html += '</select></td></tr>';
    });
    html += '</table>';
    html += '<button id="saveMappingBtn">Save Mapping</button>';
    html += '<button id="closeAdminBtn">Close</button>';
    html += '<hr><h4>Customization</h4>';
    html += '<label><input type="checkbox" id="showStepNumbers"> Show Step Numbers</label><br>';
    html += '<label><input type="checkbox" id="compactView"> Compact Table View</label><br>';
    html += '</div>';
    const adminDiv = document.createElement('div');
    adminDiv.innerHTML = html;
    document.body.appendChild(adminDiv);
    document.getElementById('saveMappingBtn').onclick = () => {
        columnMapping = {};
        requiredFields.forEach(field => {
            const val = document.getElementById('map_' + field.key.replace(/\s/g,'_')).value;
            columnMapping[field.key] = val;
        });
        localStorage.setItem('testToolColumnMapping', JSON.stringify(columnMapping));
        alert('Mapping saved!');
        adminDiv.remove(); // Hide admin panel after saving
    };
    document.getElementById('closeAdminBtn').onclick = () => {
        adminDiv.remove(); // Hide admin panel on cancel/close
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
}

// Add Admin button to UI
if (!document.getElementById('adminBtn')) {
    const adminBtn = document.createElement('button');
    adminBtn.id = 'adminBtn';
    adminBtn.textContent = 'Admin/Settings';
    adminBtn.style = 'position:fixed;top:10px;right:10px;z-index:1000;';
    document.body.appendChild(adminBtn);
    adminBtn.onclick = () => {
        // Check if testCases are loaded, if not use default mapping
        if (testCases.length === 0 && !columnMapping) {
            columnMapping = getCurrentMapping(Object.values(defaultMapping));
        }
        // If Already visible, close it
        const existingPanel = document.getElementById('adminPanel');
        if (existingPanel) {
            existingPanel.remove();
            return;
        }
        // If not, show admin panel
        if (!columnMapping) {
            columnMapping = getCurrentMapping(Object.values(defaultMapping));
        }
        if (!columnMapping) {
            showMappingUI(Object.values(defaultMapping));
            return;
        }
        // If mapping exists, show admin panel with current mapping
        if (Object.keys(columnMapping).length === 0) {
            columnMapping = getCurrentMapping(Object.values(defaultMapping));
        }
        if (!columnMapping) {
            columnMapping = defaultMapping; // Fallback to default mapping
        }
        // Use last imported headers or default
        let headersFromExcel = [];
        if (testCases.length > 0) {
            headersFromExcel = Object.keys(testCases[0]).filter(k => k !== 'steps');
        } else {
            headersFromExcel = Object.values(defaultMapping);
        }
        showAdminPanel(headersFromExcel);
    };
}

importBtn.onclick = () => {
    if (!excelFile.files[0]) return alert('Please select an Excel file.');
    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, { defval: '' });
        if (json.length === 0) return alert('No data found in Excel.');
        const excelHeaders = Object.keys(json[0]);
        columnMapping = getCurrentMapping(excelHeaders);
        if (columnMapping) {
            parseExcelWithMapping();
            // generateReportBtn.style.display = 'none';
            resetBtn.style.display = 'inline-block';
            return;
        }
        showAdminPanel(excelHeaders);
        // generateReportBtn.style.display = 'none';
        resetBtn.style.display = 'inline-block';
    };
    reader.readAsArrayBuffer(excelFile.files[0]);
};

function parseExcelWithMapping() {
    // Use columnMapping to parse testCases
    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, { defval: '' });
        testCases = json.map(row => {
            // Enhanced step parsing for numbered steps and expected results
            const stepsRaw = row[columnMapping['Test Steps']] || '';
            // Split by regex: new step starts with number and parenthesis, e.g., 1) or 2)
            const stepBlocks = stepsRaw.split(/(?=\n?\d+\))/g).map(s => s.trim()).filter(Boolean);
            let stepObjs = [];
            stepBlocks.forEach(block => {
                // Split block into lines
                const lines = block.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
                let stepText = '';
                let expected = '';
                lines.forEach(line => {
                    if (/^Expected result[:：]/i.test(line)) {
                        expected += (expected ? ' ' : '') + line.replace(/^Expected result[:：]/i, '').trim();
                    } else {
                        stepText += (stepText ? ' ' : '') + line;
                    }
                });
                if (stepText) {
                    stepObjs.push({
                        step: stepText,
                        expected: expected,
                        actual: '',
                        screenshot: '',
                        result: ''
                    });
                }
            });
            return {
                'Test Case ID': row[columnMapping['Test Case ID']] || '',
                'Name': row[columnMapping['Name']] || '',
                'Description': row[columnMapping['Description']] || '',
                'Preconditions': row[columnMapping['Preconditions']] || '',
                'Priority': row[columnMapping['Priority']] || '',
                'Type': row[columnMapping['Type']] || '',
                'Status': row[columnMapping['Status']] || '',
                steps: stepObjs
            };
        });
        renderTable();
        generateReportBtn.style.display = 'inline-block';
    };
    reader.readAsArrayBuffer(excelFile.files[0]);
};

function renderTable() {
    const showStepNumbers = localStorage.getItem('showStepNumbers') === 'true';
    const compactView = localStorage.getItem('compactView') === 'true';
    let html = `<table${compactView ? ' class="compactView"' : ''}><thead><tr>`;
    headers.forEach(h => html += `<th>${h}</th>`);
    html += '<th>Test Steps</th></tr></thead><tbody>';
    testCases.forEach((tc, i) => {
        let overall = 'passed';
        if (tc.steps.some(s => s.result === 'failed')) overall = 'failed';
        else if (tc.steps.some(s => s.result === 'skipped')) overall = 'skipped';
        html += `<tr data-idx="${i}" class="${overall}">`;
        headers.forEach(h => {
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

window.updateStepActual = function(tcIdx, stepIdx, val) {
    testCases[tcIdx].steps[stepIdx].actual = val;
};

window.updateStepResult = function(tcIdx, stepIdx, val) {
    testCases[tcIdx].steps[stepIdx].result = val;
    updateRowColors();
};

window.handleStepPaste = function(e, tcIdx, stepIdx) {
    const items = (e.clipboardData || window.clipboardData).items;
    for (let item of items) {
        if (item.type.indexOf('image') !== -1) {
            const file = item.getAsFile();
            const reader = new FileReader();
            reader.onload = function(evt) {
                testCases[tcIdx].steps[stepIdx].screenshot = evt.target.result;
                renderTable();
            };
            reader.readAsDataURL(file);
            e.preventDefault();
            return;
        }
    }
};

function updateRowColors() {
    const rows = tableContainer.querySelectorAll('tbody > tr');
    testCases.forEach((tc, i) => {
        let overall = 'passed';
        if (tc.steps.some(s => s.result === 'failed')) overall = 'failed';
        else if (tc.steps.some(s => s.result === 'skipped')) overall = 'skipped';
        rows[i].classList.remove('passed','failed','skipped');
        if (overall) rows[i].classList.add(overall);
        // Update step row colors
        const stepRows = rows[i].querySelectorAll('.steps-table tbody tr');
        tc.steps.forEach((step, si) => {
            stepRows[si].classList.remove('passed','failed','skipped');
            if (step.result) stepRows[si].classList.add(step.result);
        });
    });
}

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
