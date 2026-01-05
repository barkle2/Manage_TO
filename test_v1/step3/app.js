
// CRUD Logic for TO Editor

const CSV_PATH = './TO_list.csv';
const SERVER_URL = 'http://localhost:8000/save';

// State
let rawData = [];
let filteredData = [];
// We need to keep track of filtered indices to update rawData correctly if we edit filtered view?
// Actually simpler: 
// - Rendering: Draw filteredData. 
// - Editing: Update object in memory (which is same reference as in rawData if we don't clone).
// - Deleting: Need to remove from rawData.

// DOM Elements
const els = {
    slicers: {
        d3: document.getElementById('slicer-d3'),
        d4: document.getElementById('slicer-d4'),
        d5: document.getElementById('slicer-d5'),
        d6: document.getElementById('slicer-d6'),
    },
    gridBody: document.getElementById('grid-body'),
    totalCountBadge: document.getElementById('total-count-badge'),
    statsContainer: document.getElementById('rank-stats-container'),
    btnReset: document.getElementById('btn-reset'),
    btnAdd: document.getElementById('btn-add'),
    btnSave: document.getElementById('btn-save'),
};

// Initialize
document.addEventListener('DOMContentLoaded', () => {
    loadData();
    setupEventListeners();
});

function loadData() {
    Papa.parse(CSV_PATH, {
        download: true,
        header: true,
        skipEmptyLines: true,
        complete: (results) => {
            console.log('Data Loaded:', results.data.length);
            rawData = results.data.filter(r => r['직급'] || r['부서6']); // Clean empty rows

            initSlicers();
            applyFilters();
        },
        error: (err) => {
            console.error('Error loading CSV:', err);
            alert('Error loading CSV data.');
        }
    });
}

function setupEventListeners() {
    // Slicers
    Object.values(els.slicers).forEach(select => {
        select.addEventListener('change', () => {
            applyFilters();
            // Optional: update dependent slicers
        });
    });

    els.btnReset.addEventListener('click', () => {
        Object.values(els.slicers).forEach(s => s.value = '');
        applyFilters();
    });

    // Load Button
    els.btnLoad = document.getElementById('btn-load');
    els.fileInput = document.getElementById('file-input');

    els.btnLoad.addEventListener('click', () => {
        els.fileInput.click();
    });

    els.fileInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) {
            Papa.parse(file, {
                header: true,
                skipEmptyLines: true,
                complete: (results) => {
                    console.log('File Loaded:', file.name, results.data.length);
                    // Update rawData
                    rawData = results.data.filter(r => r['직급'] || r['부서6']);

                    // Reset Slicers and Re-render
                    initSlicers();
                    Object.values(els.slicers).forEach(s => s.value = '');
                    applyFilters();

                    // Optional: Update title or notify user
                    alert(`Loaded ${file.name} (${rawData.length} rows)`);
                },
                error: (err) => {
                    console.error('Error parsing file:', err);
                    alert('Error parsing CSV file.');
                }
            });
        }
    });

    // CRUD Buttons
    els.btnAdd.addEventListener('click', () => {
        // Add empty row at top of rawData
        // Default structure
        const newRow = {
            '부서1': '', '부서2': '', '부서3': '', '부서4': '', '부서5': '', '부서6': '',
            '직급': '', '정원수': '1'
        };
        rawData.unshift(newRow);
        applyFilters(); // Re-render
        // Scroll to top?
        document.querySelector('main').scrollTop = 0;
    });

    els.btnSave.addEventListener('click', saveData);
}

// ---------------------------------------------------------
// Slicer Logic (simplified)
// ---------------------------------------------------------
function getUniqueValues(data, col) {
    const set = new Set(data.map(r => r[col]).filter(v => v && v.trim() !== ''));
    return Array.from(set).sort();
}

function populateSelect(select, values) {
    const current = select.value;
    select.innerHTML = '<option value="">All</option>';
    values.forEach(v => {
        const option = document.createElement('option');
        option.value = v;
        option.textContent = v;
        select.appendChild(option);
    });
    // Restore selection if valid
    if (values.includes(current)) select.value = current;
}

function initSlicers() {
    // Initial populate
    populateSelect(els.slicers.d3, getUniqueValues(rawData, '부서3'));
    populateSelect(els.slicers.d4, getUniqueValues(rawData, '부서4'));
    populateSelect(els.slicers.d5, getUniqueValues(rawData, '부서5'));
    populateSelect(els.slicers.d6, getUniqueValues(rawData, '부서6'));
}

// ---------------------------------------------------------
// Filtering & Rendering
// ---------------------------------------------------------
function applyFilters() {
    const sel = {
        d3: els.slicers.d3.value,
        d4: els.slicers.d4.value,
        d5: els.slicers.d5.value,
        d6: els.slicers.d6.value,
    };

    filteredData = rawData.filter(r => {
        if (sel.d3 && r['부서3'] !== sel.d3) return false;
        if (sel.d4 && r['부서4'] !== sel.d4) return false;
        if (sel.d5 && r['부서5'] !== sel.d5) return false;
        if (sel.d6 && r['부서6'] !== sel.d6) return false;
        return true;
    });

    renderGrid();
    updateStats();
}

function renderGrid() {
    els.gridBody.innerHTML = '';

    // Virtual rendering would be better, but 650 rows is manageable in modern structure
    filteredData.forEach((row, index) => {
        // Find actual index in rawData for deletion? 
        // We can pass the object reference to delete function?
        // Or store rawIndex.

        const tr = document.createElement('tr');
        tr.className = 'hover:bg-indigo-50/50 group';

        // Helper to bind data
        const createCell = (key, widthClass, isNum = false) => {
            const td = document.createElement('td');
            td.className = `border border-gray-200 px-2 py-1 outline-none focus:bg-white focus:ring-2 focus:ring-indigo-500 focus:z-10 truncate ${widthClass}`;
            td.contentEditable = true;
            td.textContent = row[key] || '';
            if (isNum) td.className += ' text-right';

            // Edit Listener
            td.addEventListener('input', (e) => {
                row[key] = e.target.textContent; // Direct update by reference
                // Recalculate stats debounced?
                // For now, update stats on blur or simple timeout could be added.
            });
            td.addEventListener('blur', updateStats); // Update sums on blur

            return td;
        };

        // Index Cell (Read Only)
        const tdIdx = document.createElement('td');
        tdIdx.className = 'border border-gray-200 px-2 py-1 text-center bg-gray-50 text-gray-400 select-none';
        tdIdx.textContent = index + 1;
        tr.appendChild(tdIdx);

        // Data Cells
        tr.appendChild(createCell('부서1', ''));
        tr.appendChild(createCell('부서2', ''));
        tr.appendChild(createCell('부서3', 'font-medium'));
        tr.appendChild(createCell('부서4', ''));
        tr.appendChild(createCell('부서5', ''));
        tr.appendChild(createCell('부서6', 'text-indigo-900 font-medium'));
        tr.appendChild(createCell('직급', 'text-indigo-700'));
        tr.appendChild(createCell('정원수', 'font-bold text-indigo-700', true));

        // Actions Cell
        const tdAct = document.createElement('td');
        tdAct.className = 'border border-gray-200 px-2 py-1 text-center';

        const btnDel = document.createElement('button');
        btnDel.className = 'text-gray-400 hover:text-red-600 transition-colors p-1 opacity-0 group-hover:opacity-100';
        btnDel.innerHTML = `<svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"></path></svg>`;
        btnDel.title = 'Delete Row';
        btnDel.onclick = () => deleteRow(row);

        tdAct.appendChild(btnDel);
        tr.appendChild(tdAct);

        els.gridBody.appendChild(tr);
    });
}

function deleteRow(rowObj) {
    if (!confirm('Delete this row?')) return;

    // Remove from rawData
    const idx = rawData.indexOf(rowObj);
    if (idx > -1) {
        rawData.splice(idx, 1);
        applyFilters(); // Re-render
    }
}

function updateStats() {
    // Calc Total
    const total = filteredData.reduce((sum, r) => sum + (parseFloat(r['정원수']) || 0), 0);
    els.totalCountBadge.textContent = `${total} Total (${filteredData.length} Rows)`;

    // Calc Rank Breakdown
    const rankMap = {};
    filteredData.forEach(r => {
        const rank = r['직급'] || '(Empty)';
        const count = parseFloat(r['정원수']) || 0;
        rankMap[rank] = (rankMap[rank] || 0) + count;
    });

    // Render Transposed Stats
    els.statsContainer.innerHTML = '';

    // Sort Descending
    const sorted = Object.entries(rankMap).sort((a, b) => b[1] - a[1]);

    sorted.forEach(([rank, count]) => {
        const el = document.createElement('div');
        el.className = 'flex flex-col min-w-[100px] border-r border-gray-100 px-3 last:border-0';
        const ratio = total > 0 ? ((count / total) * 100).toFixed(1) : '0';

        el.innerHTML = `
            <span class="text-[10px] text-gray-500 truncate w-full" title="${rank}">${rank}</span>
            <div class="flex items-baseline gap-1">
                <span class="text-lg font-bold text-indigo-700 leading-none">${count}</span>
                <span class="text-[10px] text-gray-400">${ratio}%</span>
            </div>
        `;
        els.statsContainer.appendChild(el);
    });
}

// ---------------------------------------------------------
// Saving
// ---------------------------------------------------------
function saveData() {
    const btn = els.btnSave;
    const origText = btn.innerHTML;
    btn.textContent = 'Saving...';
    btn.disabled = true;

    // Convert rawData back to CSV string
    const csv = Papa.unparse(rawData);

    // Send to Server
    fetch(SERVER_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ csvContent: csv })
    })
        .then(res => res.json())
        .then(data => {
            if (data.status === 'success') {
                alert(data.message);
            } else {
                alert('Error saving: ' + data.message);
            }
        })
        .catch(err => {
            console.error(err);
            alert('Network error saving file.');
        })
        .finally(() => {
            btn.innerHTML = origText;
            btn.disabled = false;
        });
}
