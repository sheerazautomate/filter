document.addEventListener('DOMContentLoaded', () => {

    // ================================================================
    // SCREEN NAVIGATION
    // ================================================================
    const screen1 = document.getElementById('screen-1');
    const screen2 = document.getElementById('screen-2');
    const toast = document.getElementById('toast');

    function safeScrollIntoView(el) {
        if (el && typeof el.scrollIntoView === 'function') {
            el.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }
    }

    function goToScreen(fromEl, toEl) {
        fromEl.classList.remove('active');
        toEl.classList.add('active');
        safeScrollIntoView(toEl);
    }

    function showToast(message, isSuccess) {
        document.getElementById('toast-message').textContent = message;
        toast.classList.toggle('success', !!isSuccess);
        toast.classList.remove('hidden');

        // restart animation cleanly
        toast.classList.remove('show');
        void toast.offsetWidth; // force reflow

        setTimeout(() => toast.classList.add('show'), 10);
        setTimeout(() => {
            toast.classList.remove('show');
            setTimeout(() => toast.classList.add('hidden'), 400);
        }, 4000);
    }

    function escHtml(s) {
        return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    }

    function fmt(n) { return n.toLocaleString(); }

    // ================================================================
    // SCREEN 1: DATA MERGER
    // ================================================================
    const btnMerge = document.getElementById('btn-merge');

    // Custom Multi-select Elements
    const multiSelectBox = document.getElementById('multi-select-box');
    const multiSelectDropdown = document.getElementById('multi-select-dropdown');
    const multiSelectText = document.getElementById('multi-select-text');
    const checkboxes = document.querySelectorAll('.append-cb');

    // Data Engine Elements
    const excelUpload = document.getElementById('excel-upload');
    const headerRowInput = document.getElementById('header-row');
    const excelColumnSelect = document.getElementById('excel-column');
    const baseColumnSelect = document.getElementById('base-column');

    // State
    let excelData2D = [];   // raw sheet as 2D array
    let baseData = [];      // base.json records
    let uploadedFileName = '';

    // --- Fetch Base JSON on Load ---
    fetch('./base.json')
        .then(response => {
            if (!response.ok) throw new Error('Network response was not ok');
            return response.json();
        })
        .then(data => {
            baseData = Array.isArray(data) ? data : [data];
            console.log('Base data loaded into memory:', baseData.length, 'records.');
        })
        .catch(error => {
            console.error('Error loading base.json:', error);
            showToast('Failed to load base.json. Check console.');
        });

    // --- Excel File Upload ---
    excelUpload.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (!file) return;

        uploadedFileName = file.name.replace(/\.[^.]+$/, '');

        const reader = new FileReader();
        reader.onload = (event) => {
            const data = new Uint8Array(event.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];

            excelData2D = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
            populateHeaders();
        };
        reader.onerror = () => showToast('Could not read that file. Please try again.');
        reader.readAsArrayBuffer(file);
    });

    headerRowInput.addEventListener('change', () => {
        if (excelData2D.length > 0) populateHeaders();
    });

    function populateHeaders() {
        const rowNumber = parseInt(headerRowInput.value, 10) || 1;
        const rowIndex = rowNumber - 1;

        excelColumnSelect.innerHTML = '<option value="" disabled selected>Select column from sheet...</option>';

        if (rowIndex >= 0 && rowIndex < excelData2D.length) {
            const headerRow = excelData2D[rowIndex];
            headerRow.forEach((header) => {
                if (header !== undefined && header !== null && header.toString().trim() !== '') {
                    const headerText = header.toString().trim();
                    const option = document.createElement('option');
                    option.value = headerText;
                    option.textContent = headerText;
                    excelColumnSelect.appendChild(option);
                }
            });
        }
    }

    // --- The Core Merging Engine ---
    btnMerge.addEventListener('click', () => {

        if (excelData2D.length === 0) {
            showToast('Please upload an Excel file first.');
            return;
        }
        if (!excelColumnSelect.value) {
            showToast('Please select the Excel column to match.');
            return;
        }
        if (baseData.length === 0) {
            showToast('Base data is missing or still loading.');
            return;
        }

        const colsToAppend = Array.from(checkboxes).filter(cb => cb.checked).map(cb => cb.value);
        if (colsToAppend.length === 0) {
            showToast('Please select at least one column to append.');
            return;
        }

        const excelColName = excelColumnSelect.value;
        const baseColName = baseColumnSelect.value;
        const headerRowIndex = (parseInt(headerRowInput.value, 10) || 1) - 1;
        const originalHeaders = excelData2D[headerRowIndex];

        if (!originalHeaders) {
            showToast('That header row number is beyond the end of the sheet.');
            return;
        }

        const excelColIndex = originalHeaders.findIndex(header =>
            header !== undefined &&
            header !== null &&
            header.toString().trim() === excelColName
        );
        if (excelColIndex === -1) {
            showToast('Selected column not found in the header row.');
            return;
        }

        // Fast lookup map from base.json (strict string keys)
        const baseMap = new Map();
        baseData.forEach(record => {
            if (record[baseColName] !== undefined && record[baseColName] !== null) {
                baseMap.set(String(record[baseColName]).trim(), record);
            }
        });

        let failedMatchCount = 0;
        let successMatchCount = 0;

        let merged2D = JSON.parse(JSON.stringify(excelData2D));
        merged2D[headerRowIndex] = [...originalHeaders, ...colsToAppend];

        for (let i = headerRowIndex + 1; i < merged2D.length; i++) {
            let row = merged2D[i];

            if (!row || row.length === 0 || row.every(cell => cell === '' || cell === null)) {
                continue;
            }

            let cellValue = row[excelColIndex];
            let matchFound = false;

            if (cellValue !== undefined && cellValue !== null && cellValue.toString().trim() !== '') {
                const searchKey = String(cellValue).trim();
                const matchedRecord = baseMap.get(searchKey);

                if (matchedRecord) {
                    colsToAppend.forEach(col => {
                        row.push(matchedRecord[col] !== undefined ? matchedRecord[col] : '');
                    });
                    matchFound = true;
                    successMatchCount++;
                }
            }

            if (!matchFound) {
                colsToAppend.forEach(() => row.push(''));
                failedMatchCount++;
            }
        }

        const finalSheet = XLSX.utils.aoa_to_sheet(merged2D.slice(headerRowIndex));
        const finalMergedData = XLSX.utils.sheet_to_json(finalSheet, { defval: '' });

        if (finalMergedData.length === 0) {
            showToast('No data rows found below the header row.');
            return;
        }

        loadMergedData(finalMergedData, { successMatchCount, failedMatchCount });
    });

    // Multi-select dropdown
    multiSelectBox.addEventListener('click', (e) => {
        e.stopPropagation();
        multiSelectDropdown.classList.toggle('hidden');
    });

    document.addEventListener('click', (e) => {
        if (!multiSelectBox.contains(e.target) && !multiSelectDropdown.contains(e.target)) {
            multiSelectDropdown.classList.add('hidden');
        }
    });

    checkboxes.forEach(cb => cb.addEventListener('change', updateMultiSelectText));

    function updateMultiSelectText() {
        const checked = Array.from(checkboxes).filter(cb => cb.checked).map(cb => cb.value);
        if (checked.length === 0) {
            multiSelectText.textContent = 'Select columns...';
        } else if (checked.length === 1) {
            multiSelectText.textContent = '1 column selected';
        } else {
            multiSelectText.textContent = `${checked.length} columns selected`;
        }
    }

    // ================================================================
    // BRIDGE: merged data -> filter tool
    // ================================================================
    function loadMergedData(data, stats) {
        allRows = data;
        headers = Object.keys(data[0]);
        filteredRows = [];
        filterLogic = 'and';
        document.getElementById('logic-and').classList.add('active');
        document.getElementById('logic-or').classList.remove('active');

        document.getElementById('file-meta').textContent =
            `${fmt(allRows.length)} rows · ${headers.length} columns · (merged data)`;

        // Reset filter rows and results/download panels
        document.getElementById('filter-rows').innerHTML = '';
        filterCount = 0;
        clearErr('err-filter');
        addFilterRow();

        document.getElementById('results-section').classList.add('hidden');
        document.getElementById('download-section').classList.add('hidden');

        goToScreen(screen1, screen2);

        if (stats) {
            if (stats.failedMatchCount > 0) {
                showToast(`Merged with ${stats.failedMatchCount} unmatched row(s) — ${stats.successMatchCount} matched.`, false);
            } else {
                showToast(`Success: all ${stats.successMatchCount} row(s) matched!`, true);
            }
        }
    }

    // ================================================================
    // SCREEN 2: FILTER / RESULTS / DOWNLOAD
    // (adapted from the original standalone SheetFilter tool,
    //  driven entirely by data handed off from the merge step)
    // ================================================================
    let allRows = [];
    let headers = [];
    let filteredRows = [];
    let filterLogic = 'and';
    let filterCount = 0;
    let customFileName = '';

    const OPS = [
        { val: 'contains', label: 'contains' },
        { val: 'not_contains', label: 'does not contain' },
        { val: 'equals', label: 'equals' },
        { val: 'not_equals', label: 'not equals' },
        { val: 'starts', label: 'starts with' },
        { val: 'ends', label: 'ends with' },
        { val: 'gt', label: '> greater than' },
        { val: 'lt', label: '< less than' },
        { val: 'gte', label: '>= at least' },
        { val: 'lte', label: '<= at most' },
        { val: 'empty', label: 'is empty' },
        { val: 'not_empty', label: 'is not empty' },
    ];

    function showErr(id, msg) {
        const el = document.getElementById(id);
        el.textContent = msg;
        el.className = 'msg-err show';
    }

    function clearErr(id) {
        document.getElementById(id).className = 'msg-err';
    }

    function makeHeaderOptions(selectedVal) {
        return headers.map(h =>
            `<option value="${escHtml(h)}" ${h === selectedVal ? 'selected' : ''}>${escHtml(h)}</option>`
        ).join('');
    }

    function makeOpOptions(selectedVal) {
        return OPS.map(op =>
            `<option value="${op.val}" ${op.val === selectedVal ? 'selected' : ''}>${op.label}</option>`
        ).join('');
    }

    function addFilterRow(col, op, val) {
        filterCount++;
        const id = 'fr-' + filterCount;
        const row = document.createElement('div');
        row.className = 'filter-row';
        row.id = id;
        row.innerHTML = `
            <select class="f-col">${makeHeaderOptions(col || headers[0])}</select>
            <select class="f-op">${makeOpOptions(op || 'contains')}</select>
            <input class="f-val" type="text" placeholder="value…" value="${escHtml(val || '')}">
            <button class="f-remove" title="Remove filter" type="button">✕</button>
        `;
        const opSel = row.querySelector('.f-op');
        const valIn = row.querySelector('.f-val');
        opSel.addEventListener('change', () => {
            valIn.style.display = ['empty', 'not_empty'].includes(opSel.value) ? 'none' : '';
        });
        if (['empty', 'not_empty'].includes(op)) valIn.style.display = 'none';
        row.querySelector('.f-remove').addEventListener('click', () => row.remove());
        document.getElementById('filter-rows').appendChild(row);
    }

    document.getElementById('btn-add-filter').addEventListener('click', () => addFilterRow());

    document.querySelectorAll('.logic-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            filterLogic = btn.dataset.logic;
            document.querySelectorAll('.logic-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
        });
    });

    // --- Apply Filters ---
    document.getElementById('btn-apply').addEventListener('click', applyFilters);

    function applyFilters() {
        clearErr('err-filter');
        const rows = document.querySelectorAll('#filter-rows .filter-row');
        const rules = [];
        for (const row of rows) {
            const col = row.querySelector('.f-col').value;
            const op = row.querySelector('.f-op').value;
            const val = row.querySelector('.f-val').value.trim();
            if (!['empty', 'not_empty'].includes(op) && val === '') {
                showErr('err-filter', `Value is empty for column "${col}". Enter a value or choose "is empty".`);
                return;
            }
            rules.push({ col, op, val });
        }
        if (rules.length === 0) {
            showErr('err-filter', 'Add at least one filter rule.');
            return;
        }

        filteredRows = allRows.filter(row => {
            const results = rules.map(r => matchRule(row, r));
            return filterLogic === 'and' ? results.every(Boolean) : results.some(Boolean);
        });

        renderResults(rules);
    }

    function matchRule(row, rule) {
        const cellRaw = row[rule.col];
        const cell = String(cellRaw ?? '').toLowerCase();
        const val = rule.val.toLowerCase();
        const num = parseFloat(cellRaw);
        const rNum = parseFloat(rule.val);
        switch (rule.op) {
            case 'contains': return cell.includes(val);
            case 'not_contains': return !cell.includes(val);
            case 'equals': return cell === val;
            case 'not_equals': return cell !== val;
            case 'starts': return cell.startsWith(val);
            case 'ends': return cell.endsWith(val);
            case 'gt': return !isNaN(num) && !isNaN(rNum) && num > rNum;
            case 'lt': return !isNaN(num) && !isNaN(rNum) && num < rNum;
            case 'gte': return !isNaN(num) && !isNaN(rNum) && num >= rNum;
            case 'lte': return !isNaN(num) && !isNaN(rNum) && num <= rNum;
            case 'empty': return cell === '' || cellRaw === null || cellRaw === undefined;
            case 'not_empty': return cell !== '' && cellRaw !== null && cellRaw !== undefined;
            default: return true;
        }
    }

    function renderResults(rules) {
        const total = allRows.length;
        const match = filteredRows.length;
        const pct = total > 0 ? Math.round((match / total) * 100) : 0;

        document.getElementById('stat-total').textContent = fmt(total);
        document.getElementById('stat-match').textContent = fmt(match);
        document.getElementById('stat-pct').textContent = pct + '%';

        const tagRow = document.getElementById('active-filters-tag');
        tagRow.innerHTML = rules.map(r => {
            const opLabel = (OPS.find(o => o.val === r.op) || {}).label || r.op;
            const txt = ['empty', 'not_empty'].includes(r.op)
                ? `${r.col} ${opLabel}`
                : `${r.col} ${opLabel} "${r.val}"`;
            return `<span class="tag">${escHtml(txt)}</span>`;
        }).join('') + `<span class="tag" style="background:transparent;border-color:var(--glass-border);color:var(--text-muted);">${filterLogic.toUpperCase()}</span>`;

        const container = document.getElementById('table-container');
        if (match === 0) {
            container.innerHTML = `<div class="empty-results"><div class="big">🔍</div>No rows matched your filters.</div>`;
        } else {
            const preview = filteredRows.slice(0, 50);
            let html = `<div class="table-wrap"><table><thead><tr>`;
            headers.forEach(h => { html += `<th>${escHtml(h)}</th>`; });
            html += `</tr></thead><tbody>`;
            preview.forEach(row => {
                html += '<tr>';
                headers.forEach(h => { html += `<td>${escHtml(String(row[h] ?? ''))}</td>`; });
                html += '</tr>';
            });
            html += `</tbody></table>`;
            if (match > 50) {
                html += `<div class="more-rows">Showing 50 of ${fmt(match)} rows — full data in download</div>`;
            }
            html += `</div>`;
            container.innerHTML = html;
        }

        document.getElementById('results-section').classList.remove('hidden');
        const downloadSection = document.getElementById('download-section');
        if (match > 0) {
            downloadSection.classList.remove('hidden');
            updateFilenameDisplay(rules);
        } else {
            downloadSection.classList.add('hidden');
        }

        safeScrollIntoView(document.getElementById('results-section'));
    }

    // --- Filename ---
    function buildSmartName(rules) {
        const base = uploadedFileName || 'filtered_data';
        const parts = rules.map(r => {
            const val = ['empty', 'not_empty'].includes(r.op) ? r.op.replace('_', ' ') : r.val;
            return (r.col + '_' + val).replace(/[^a-zA-Z0-9_-]/g, '_').replace(/_+/g, '_').slice(0, 30);
        });
        const suffix = parts.slice(0, 3).join('-');
        return base + '_' + suffix;
    }

    function getDownloadName() {
        return (customFileName || document.getElementById('fn-display').textContent).replace(/\.[^.]+$/, '');
    }

    function updateFilenameDisplay(rules) {
        const smart = buildSmartName(rules);
        customFileName = '';
        document.getElementById('fn-display').textContent = smart;
        document.getElementById('fn-input').value = smart;
    }

    document.getElementById('fn-edit-btn').addEventListener('click', () => {
        const row = document.getElementById('fn-edit-row');
        row.classList.toggle('show');
        if (row.classList.contains('show')) {
            const input = document.getElementById('fn-input');
            input.value = document.getElementById('fn-display').textContent;
            input.focus();
            input.select();
        }
    });

    document.getElementById('fn-save-btn').addEventListener('click', () => {
        const val = document.getElementById('fn-input').value.trim()
            .replace(/\.[^.]+$/, '')
            .replace(/[\\/:*?"<>|]/g, '_');
        if (val) {
            customFileName = val;
            document.getElementById('fn-display').textContent = val;
        }
        document.getElementById('fn-edit-row').classList.remove('show');
    });

    document.getElementById('fn-input').addEventListener('keydown', (e) => {
        if (e.key === 'Enter') document.getElementById('fn-save-btn').click();
        if (e.key === 'Escape') document.getElementById('fn-edit-row').classList.remove('show');
    });

    document.getElementById('btn-dl-xlsx').addEventListener('click', () => {
        const ws = XLSX.utils.json_to_sheet(filteredRows, { header: headers });
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Filtered');
        XLSX.writeFile(wb, getDownloadName() + '.xlsx');
    });

    document.getElementById('btn-dl-csv').addEventListener('click', () => {
        const ws = XLSX.utils.json_to_sheet(filteredRows, { header: headers });
        const csv = XLSX.utils.sheet_to_csv(ws);
        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = getDownloadName() + '.csv';
        a.click();
        URL.revokeObjectURL(a.href);
    });

    // --- Navigation: back to merger ---
    document.getElementById('btn-back').addEventListener('click', () => {
        goToScreen(screen2, screen1);
    });

});
