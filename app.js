// ================================
// State
// ================================
const state = {
    data: null,
    longData: null,
    currentFile: null,
    selectedProduct: null,
    selectedSubtype: null,
    factors: {},
    results: [],
    charts: {}
};

// ================================
// Data Parsing
// ================================
function parseExcelData(jsonData) {
    if (!jsonData || jsonData.length < 3) {
        alert("Excel muss mindestens 3 Spalten haben");
        return null;
    }

    const headers = Object.keys(jsonData[0]);
    
    // Find product and subtype columns
    let productCol = headers.find(h => /produkt|product/i.test(h)) || headers[0];
    let subtypeCol = headers.find(h => /untertyp|subtype|typ/i.test(h)) || headers[1];
    
    // Find date columns (MM/YYYY format)
    const dateCols = headers.filter(h => {
        if (h === productCol || h === subtypeCol) return false;
        const match = h.match(/^(\d{1,2})\/(\d{4})$/);
        return match && parseInt(match[1]) >= 1 && parseInt(match[1]) <= 12;
    });

    if (dateCols.length === 0) {
        alert("Keine gültigen Monatsspalten gefunden (erwartet: MM/YYYY)");
        return null;
    }

    // Sort date columns
    dateCols.sort((a, b) => {
        const [ma, ya] = a.split('/').map(Number);
        const [mb, yb] = b.split('/').map(Number);
        return ya !== yb ? ya - yb : ma - mb;
    });

    // Convert to long format
    const records = [];
    jsonData.forEach(row => {
        const produkt = row[productCol];
        const untertyp = row[subtypeCol];
        
        if (!produkt) return;

        dateCols.forEach(col => {
            const [m, y] = col.split('/').map(Number);
            const val = row[col];
            if (val !== undefined && val !== null && val !== '') {
                records.push({
                    Produkt: String(produkt),
                    Untertyp: untertyp ? String(untertyp) : '',
                    Datum: new Date(y, m - 1, 1),
                    Verbrauch: parseFloat(val),
                    MonthStr: `${String(m).padStart(2, '0')}/${y}`
                });
            }
        });
    });

    if (records.length === 0) {
        alert("Keine gültigen Daten gefunden");
        return null;
    }

    return records;
}

// ================================
// Forecasting - Holt-Winters Triple Exponential Smoothing
// ================================
function holtWintersForecast(data, periods = 60, alpha = 0.3, beta = 0.1, gamma = 0.1, seasonalLength = 12) {
    const n = data.length;
    if (n < seasonalLength * 2) {
        return simpleForecast(data, periods);
    }

    // Initialize
    let level = data.slice(0, seasonalLength).reduce((a, b) => a + b, 0) / seasonalLength;
    let trend = (data[seasonalLength] - data[0]) / seasonalLength;
    
    const seasonal = [];
    for (let i = 0; i < seasonalLength; i++) {
        seasonal[i] = data[i] / level;
    }

    // Fit model
    for (let i = 0; i < n; i++) {
        const val = data[i];
        const sIdx = i % seasonalLength;
        
        const prevLevel = level;
        level = alpha * (val / seasonal[sIdx]) + (1 - alpha) * (level + trend);
        trend = beta * (level - prevLevel) + (1 - beta) * trend;
        seasonal[sIdx] = gamma * (val / level) + (1 - gamma) * seasonal[sIdx];
    }

    // Forecast
    const forecast = [];
    for (let h = 1; h <= periods; h++) {
        const sIdx = (n + h - 1) % seasonalLength;
        const pred = (level + h * trend) * seasonal[sIdx];
        forecast.push(Math.max(0, pred));
    }

    return forecast;
}

function simpleForecast(data, periods) {
    const n = data.length;
    let sumX = 0, sumY = 0, sumXY = 0, sumXX = 0;
    
    for (let i = 0; i < n; i++) {
        sumX += i;
        sumY += data[i];
        sumXY += i * data[i];
        sumXX += i * i;
    }
    
    const slope = (n * sumXY - sumX * sumY) / (n * sumXX - sumX * sumX);
    const intercept = (sumY - slope * sumX) / n;
    
    const forecast = [];
    for (let h = 1; h <= periods; h++) {
        forecast.push(Math.max(0, intercept + slope * (n + h - 1)));
    }
    
    return forecast;
}

function makeForecast(subtypeData, factor = 1.0) {
    subtypeData.sort((a, b) => a.Datum - b.Datum);
    const values = subtypeData.map(d => d.Verbrauch);
    
    let forecast;
    try {
        forecast = holtWintersForecast(values, 60);
    } catch (e) {
        forecast = simpleForecast(values, 60);
    }
    
    if (factor !== 1.0) {
        forecast = forecast.map(v => v * factor);
    }
    
    const lastDate = subtypeData[subtypeData.length - 1].Datum;
    const forecastDates = [];
    for (let i = 0; i < 60; i++) {
        const d = new Date(lastDate);
        d.setMonth(d.getMonth() + i + 1);
        forecastDates.push(d);
    }
    
    return { dates: forecastDates, values: forecast };
}

// ================================
// UI Functions
// ================================
function showApp() {
    document.getElementById('upload-section').classList.add('hidden');
    document.getElementById('app-section').classList.remove('hidden');
}

function hideApp() {
    document.getElementById('upload-section').classList.remove('hidden');
    document.getElementById('app-section').classList.add('hidden');
    document.getElementById('export-section').classList.add('hidden');
}

function showUpload() {
    document.getElementById('app-section').classList.add('hidden');
    document.getElementById('upload-section').classList.remove('hidden');
}

function updateInfoBar() {
    const minDate = new Date(Math.min(...state.longData.map(d => d.Datum)));
    const maxDate = new Date(Math.max(...state.longData.map(d => d.Datum)));
    const minStr = `${String(minDate.getMonth() + 1).padStart(2, '0')}/${minDate.getFullYear()}`;
    const maxStr = `${String(maxDate.getMonth() + 1).padStart(2, '0')}/${maxDate.getFullYear()}`;
    
    document.getElementById('current-file').textContent = 
        `📄 ${state.currentFile} | ${state.longData.length} Datensätze | Zeitraum: ${minStr} - ${maxStr}`;
}

function populateProductSelect() {
    const dropdown = document.getElementById('product-dropdown');
    const products = [...new Set(state.longData.map(d => d.Produkt))].sort();
    
    if (products.length === 0) return;
    
    dropdown.innerHTML = products.map(p => `<option value="${p}">${p}</option>`).join('');
    
    dropdown.onchange = () => {
        state.selectedProduct = dropdown.value;
        state.selectedSubtype = null;
        renderSubtypeTabs(products);
    };
    
    state.selectedProduct = products[0];
    renderSubtypeTabs(products);
}

function renderSubtypeTabs(products) {
    const container = document.getElementById('tabs-container');
    const productData = state.longData.filter(d => d.Produkt === state.selectedProduct);
    const subtypes = [...new Set(productData.map(d => d.Untertyp))].sort();
    
    state.view = 'subtypes';
    
    // Show only subtypes for selected product
    let html = subtypes.map(subtype => `
        <button class="tab-btn subtype-tab ${state.selectedSubtype === subtype ? 'active' : ''}" data-subtype="${subtype}">${subtype}</button>
    `).join('');
    
    container.innerHTML = html;
    
    // Click subtype to show content
    container.querySelectorAll('.subtype-tab').forEach(btn => {
        btn.onclick = () => {
            container.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            state.selectedSubtype = btn.dataset.subtype;
            renderTabContent();
        };
    });
    
    // Auto-select first subtype
    if (subtypes.length > 0 && !state.selectedSubtype) {
        state.selectedSubtype = subtypes[0];
        const subtypeBtn = container.querySelector(`[data-subtype="${subtypes[0]}"]`);
        if (subtypeBtn) subtypeBtn.classList.add('active');
        renderTabContent();
    } else if (subtypes.length === 0) {
        document.getElementById('tab-content').innerHTML = '<p>Keine Unterprodukte gefunden</p>';
    }
}

function renderTabContent() {
    const container = document.getElementById('tab-content');
    const subtypeData = state.longData
        .filter(d => d.Produkt === state.selectedProduct && d.Untertyp === state.selectedSubtype)
        .sort((a, b) => a.Datum - b.Datum);
    
    if (subtypeData.length === 0) {
        container.innerHTML = '<p>Keine Daten vorhanden</p>';
        return;
    }
    
    const first6 = subtypeData.slice(0, 6);
    const last6 = subtypeData.slice(-6);
    const key = `${state.selectedProduct}_${state.selectedSubtype}`;
    const factor = state.factors[key] || 1.0;
    
    container.innerHTML = `
        <h3>${state.selectedProduct} - ${state.selectedSubtype}</h3>
        
        <div class="tables-row">
            <div class="table-card">
                <h4>Erste 6 Monate</h4>
                <table class="prognose-table">
                    <tr><th>Monat</th><th>Verbrauch</th></tr>
                    ${first6.map(d => `<tr><td>${d.MonthStr}</td><td>${d.Verbrauch.toFixed(1)}</td></tr>`).join('')}
                </table>
            </div>
            <div class="table-card">
                <h4>Letzte 6 Monate</h4>
                <table class="prognose-table">
                    <tr><th>Monat</th><th>Verbrauch</th></tr>
                    ${last6.map(d => `<tr><td>${d.MonthStr}</td><td>${d.Verbrauch.toFixed(1)}</td></tr>`).join('')}
                </table>
            </div>
        </div>
        
        <div class="chart-container">
            <canvas id="history-chart"></canvas>
        </div>
        
        <div class="controls-row">
            <div class="factor-control">
                <label>Faktor für ${state.selectedSubtype}:</label>
                <input type="range" id="factor-slider" min="0" max="5" step="0.1" value="${factor}">
                <span class="factor-value" id="factor-value">${factor.toFixed(1)}</span>
                ${factor !== 1.0 ? '<span class="factor-warning">⚠️ ' + factor + 'x</span>' : ''}
            </div>
            <button id="forecast-btn" class="btn primary">🔮 Prognose berechnen</button>
        </div>
        
        <div id="forecast-results" class="hidden"></div>
    `;
    
    renderHistoryChart(subtypeData);
    
    const slider = document.getElementById('factor-slider');
    const factorValue = document.getElementById('factor-value');
    slider.oninput = () => {
        const val = parseFloat(slider.value);
        factorValue.textContent = val.toFixed(1);
        state.factors[key] = val;
        
        const warning = document.querySelector('.factor-warning');
        if (val !== 1.0) {
            if (warning) {
                warning.textContent = `⚠️ ${val}x`;
            } else {
                factorValue.insertAdjacentHTML('afterend', `<span class="factor-warning">⚠️ ${val}x</span>`);
            }
        } else if (warning) {
            warning.remove();
        }
    };
    
    document.getElementById('forecast-btn').onclick = () => calculateForecast(subtypeData, key);
}

function renderHistoryChart(data) {
    const ctx = document.getElementById('history-chart').getContext('2d');
    
    if (state.charts.history) {
        state.charts.history.destroy();
    }
    
    const isDark = document.body.dataset.theme === 'dark';
    const gridColor = isDark ? 'rgba(255,255,255,0.1)' : 'rgba(0,0,0,0.1)';
    const textColor = isDark ? '#e0e0e0' : '#262730';
    
    state.charts.history = new Chart(ctx, {
        type: 'line',
        data: {
            labels: data.map(d => d.MonthStr),
            datasets: [{
                label: 'Verbrauch',
                data: data.map(d => d.Verbrauch),
                borderColor: '#0078d4',
                backgroundColor: 'rgba(0, 120, 212, 0.1)',
                fill: true,
                tension: 0.3
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    labels: { color: textColor }
                }
            },
            scales: {
                x: {
                    ticks: { color: textColor },
                    grid: { color: gridColor }
                },
                y: {
                    ticks: { color: textColor },
                    grid: { color: gridColor }
                }
            }
        }
    });
}

function calculateForecast(subtypeData, key) {
    const factor = state.factors[key] || 1.0;
    const { dates, values } = makeForecast(subtypeData, factor);
    
    // Year 1: monthly
    const y1 = values.slice(0, 12);
    const y1Sum = y1.reduce((a, b) => a + b, 0);
    const y1Data = y1.map((v, i) => ({ Monat: `${i + 1}/${dates[0].getFullYear()}`, Prognose: v }));
    
    // Year 2-3: quarterly
    const y2Data = [];
    for (let q = 0; q < 4; q++) {
        const sum = values.slice(12 + q * 3, 12 + (q + 1) * 3).reduce((a, b) => a + b, 0);
        y2Data.push({ Quartal: `Q${q + 1} ${dates[12].getFullYear()}`, Prognose: sum });
    }
    const y2Sum = y2Data.reduce((a, b) => a + b.Prognose, 0);
    
    const y3Data = [];
    for (let q = 0; q < 4; q++) {
        const sum = values.slice(24 + q * 3, 24 + (q + 1) * 3).reduce((a, b) => a + b, 0);
        y3Data.push({ Quartal: `Q${q + 1} ${dates[24].getFullYear()}`, Prognose: sum });
    }
    const y3Sum = y3Data.reduce((a, b) => a + b.Prognose, 0);
    
    // Year 4-5: yearly
    const y4 = values.slice(36, 48).reduce((a, b) => a + b, 0);
    const y5 = values.slice(48, 60).reduce((a, b) => a + b, 0);
    const y4Data = [
        { Jahr: dates[36].getFullYear(), Prognose: y4 },
        { Jahr: dates[48].getFullYear(), Prognose: y5 }
    ];
    
    // Last year actual
    const lastYear = subtypeData.filter(d => d.Datum.getFullYear() === subtypeData[subtypeData.length - 1].Datum.getFullYear());
    const lastYearSum = lastYear.reduce((a, b) => a + b.Verbrauch, 0);
    
    // Percentage changes
    const pctY1 = ((y1Sum - lastYearSum) / lastYearSum) * 100;
    const pctY2 = ((y2Sum - y1Sum) / y1Sum) * 100;
    const pctY3 = ((y3Sum - y2Sum) / y2Sum) * 100;
    const pctY4 = ((y4 - y3Sum) / y3Sum) * 100;
    const pctY5 = ((y5 - y4) / y4) * 100;
    
    const pctData = [
        { Jahr: 'Jahr 1', 'Veränderung %': pctY1 },
        { Jahr: 'Jahr 2', 'Veränderung %': pctY2 },
        { Jahr: 'Jahr 3', 'Veränderung %': pctY3 },
        { Jahr: 'Jahr 4', 'Veränderung %': pctY4 },
        { Jahr: 'Jahr 5', 'Veränderung %': pctY5 }
    ];
    
    // Render results
    const resultsContainer = document.getElementById('forecast-results');
    resultsContainer.classList.remove('hidden');
    resultsContainer.innerHTML = `
        <div class="results-section">
            <h3>🔮 Prognose</h3>
            
            <h4>Jahr 1 (monatlich)</h4>
            <table class="prognose-table">
                <tr><th>Monat</th><th>Prognose</th></tr>
                ${y1Data.map(d => `<tr><td>${d.Monat}</td><td>${d.Prognose.toFixed(1)}</td></tr>`).join('')}
            </table>
            
            <h4>Jahr 2 (quartalsweise)</h4>
            <table class="prognose-table">
                <tr><th>Quartal</th><th>Prognose</th></tr>
                ${y2Data.map(d => `<tr><td>${d.Quartal}</td><td>${d.Prognose.toFixed(1)}</td></tr>`).join('')}
            </table>
            
            <h4>Jahr 3 (quartalsweise)</h4>
            <table class="prognose-table">
                <tr><th>Quartal</th><th>Prognose</th></tr>
                ${y3Data.map(d => `<tr><td>${d.Quartal}</td><td>${d.Prognose.toFixed(1)}</td></tr>`).join('')}
            </table>
            
            <h4>Jahr 4-5 (jährlich)</h4>
            <table class="prognose-table">
                <tr><th>Jahr</th><th>Prognose</th></tr>
                ${y4Data.map(d => `<tr><td>${d.Jahr}</td><td>${d.Prognose.toFixed(1)}</td></tr>`).join('')}
            </table>
            
            <h3>📈 Veränderung zum Vorjahr</h3>
            <table class="prognose-table">
                <tr><th>Jahr</th><th>Veränderung %</th></tr>
                ${pctData.map(d => `
                    <tr>
                        <td>${d.Jahr}</td>
                        <td class="${d['Veränderung %'] >= 0 ? 'positive' : 'negative'}">${d['Veränderung %'] >= 0 ? '+' : ''}${d['Veränderung %'].toFixed(1)}%</td>
                    </tr>
                `).join('')}
            </table>
        </div>
    `;
    
    // Add to results
    const existingIdx = state.results.findIndex(r => r.Produkt === state.selectedProduct && r.Untertyp === state.selectedSubtype);
    const result = {
        Produkt: state.selectedProduct,
        Untertyp: state.selectedSubtype,
        Faktor: factor,
        'Jahr 1 (monatl.)': Math.round(y1Sum * 10) / 10,
        'Jahr 2 (Quartal)': Math.round(y2Sum * 10) / 10,
        'Jahr 3 (Quartal)': Math.round(y3Sum * 10) / 10,
        'Jahr 4': Math.round(y4 * 10) / 10,
        'Jahr 5': Math.round(y5 * 10) / 10,
        'Veränd. Jahr 1 %': Math.round(pctY1 * 10) / 10,
        'Veränd. Jahr 2 %': Math.round(pctY2 * 10) / 10,
        'Veränd. Jahr 3 %': Math.round(pctY3 * 10) / 10,
        'Veränd. Jahr 4 %': Math.round(pctY4 * 10) / 10,
        'Veränd. Jahr 5 %': Math.round(pctY5 * 10) / 10
    };
    
    if (existingIdx >= 0) {
        state.results[existingIdx] = result;
    } else {
        state.results.push(result);
    }
    
    showExport();
}

function showExport() {
    const section = document.getElementById('export-section');
    const container = document.getElementById('results-table-container');
    
    section.classList.remove('hidden');
    
    if (state.results.length === 0) {
        container.innerHTML = '<p>Noch keine Ergebnisse berechnet</p>';
        return;
    }
    
    const headers = Object.keys(state.results[0]);
    container.innerHTML = `
        <table class="prognose-table">
            <tr>${headers.map(h => `<th>${h}</th>`).join('')}</tr>
            ${state.results.map(r => `
                <tr>${headers.map(h => `<td>${r[h]}</td>`).join('')}</tr>
            `).join('')}
        </table>
    `;
}

function exportToExcel() {
    if (state.results.length === 0) return;
    
    const ws = XLSX.utils.json_to_sheet(state.results);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Prognose");
    XLSX.writeFile(wb, "prognose.xlsx");
}

// ================================
// Theme Toggle
// ================================
function initTheme() {
    const toggle = document.getElementById('theme-toggle');
    const icon = document.getElementById('theme-icon');
    
    const saved = localStorage.getItem('theme');
    if (saved === 'dark') {
        document.body.dataset.theme = 'dark';
        icon.textContent = '☀️';
    }
    
    toggle.onclick = () => {
        if (document.body.dataset.theme === 'dark') {
            delete document.body.dataset.theme;
            localStorage.setItem('theme', 'light');
            icon.textContent = '🌙';
        } else {
            document.body.dataset.theme = 'dark';
            localStorage.setItem('theme', 'dark');
            icon.textContent = '☀️';
        }
        
        if (state.selectedProduct && state.selectedSubtype) {
            const subtypeData = state.longData
                .filter(d => d.Produkt === state.selectedProduct && d.Untertyp === state.selectedSubtype)
                .sort((a, b) => a.Datum - b.Datum);
            renderHistoryChart(subtypeData);
        }
    };
}

// ================================
// Event Listeners
// ================================
document.addEventListener('DOMContentLoaded', () => {
    initTheme();
    
    const fileInput = document.getElementById('file-input');
    const uploadBtn = document.getElementById('upload-btn');
    const filename = document.getElementById('filename');
    
    uploadBtn.onclick = () => fileInput.click();
    
    fileInput.onchange = () => {
        if (fileInput.files.length > 0) {
            filename.textContent = fileInput.files[0].name;
        }
    };
    
    // Load Excel data on file selection (auto-load)
    fileInput.onchange = () => {
        const file = fileInput.files[0];
        if (!file) return;
        
        filename.textContent = file.name;
        
        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(firstSheet);
            
            state.longData = parseExcelData(jsonData);
            state.currentFile = file.name;
            
            if (state.longData) {
                state.factors = {};
                state.results = [];
                state.selectedProduct = null;
                state.selectedSubtype = null;
                updateInfoBar();
                populateProductSelect();
                showApp();
            }
        };
        reader.readAsArrayBuffer(file);
    };
    
    // Change file button
    document.getElementById('change-file-btn').onclick = () => {
        document.getElementById('file-input').click();
    };
    
    // Export
    document.getElementById('export-btn').onclick = exportToExcel;
});
