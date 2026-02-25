// ================================
// Translations
// ================================
const translations = {
    de: {
        upload: {
            description: "Excel-Datei mit Produkten, Untertypen und monatlichen Verbrauchsdaten hochladen",
            selectBtn: "📁 Excel auswählen",
            noFile: "Keine Datei gewählt"
        },
        app: {
            changeFile: "📁 Datei wechseln",
            product: "Produkt:",
            first6Months: "Erste 6 Monate",
            last6Months: "Letzte 6 Monate",
            month: "Monat",
            consumption: "Verbrauch",
            factor: "Faktor für",
            calculateForecast: "Prognose berechnen",
            noSubproducts: "Keine Unterprodukte gefunden",
            noData: "Keine Daten vorhanden"
        },
        forecast: {
            title: "Prognose",
            year1Monthly: "Jahr 1 (monatlich)",
            year2Quarterly: "Jahr 2 (quartalsweise)",
            year3Quarterly: "Jahr 3 (quartalsweise)",
            year4Yearly: "Jahr 4-5 (jährlich)",
            changeTitle: "📈 Veränderung zum Vorjahr",
            year: "Jahr",
            change: "Veränderung %"
        },
        export: {
            title: "💾 Export",
            downloadBtn: "📥 Excel herunterladen",
            noResults: "Noch keine Ergebnisse berechnet"
        },
        errors: {
            minColumns: "Excel muss mindestens 3 Spalten haben",
            noDateCols: "Keine gültigen Monatsspalten gefunden (erwartet: MM/YYYY)",
            noData: "Keine gültigen Daten gefunden"
        },
        infoBar: {
            records: "Datensätze",
            period: "Zeitraum"
        }
    },
    en: {
        upload: {
            description: "Upload an Excel file with products, subtypes and monthly consumption data",
            selectBtn: "📁 Select Excel",
            noFile: "No file selected"
        },
        app: {
            changeFile: "📁 Change File",
            product: "Product:",
            first6Months: "First 6 Months",
            last6Months: "Last 6 Months",
            month: "Month",
            consumption: "Consumption",
            factor: "Factor for",
            calculateForecast: "Calculate Forecast",
            noSubproducts: "No subproducts found",
            noData: "No data available"
        },
        forecast: {
            title: "Forecast",
            year1Monthly: "Year 1 (monthly)",
            year2Quarterly: "Year 2 (quarterly)",
            year3Quarterly: "Year 3 (quarterly)",
            year4Yearly: "Year 4-5 (yearly)",
            changeTitle: "📈 Change from Previous Year",
            year: "Year",
            change: "Change %"
        },
        export: {
            title: "💾 Export",
            downloadBtn: "📥 Download Excel",
            noResults: "No results calculated yet"
        },
        errors: {
            minColumns: "Excel must have at least 3 columns",
            noDateCols: "No valid month columns found (expected: MM/YYYY)",
            noData: "No valid data found"
        },
        infoBar: {
            records: "records",
            period: "Period"
        }
    }
};

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
    charts: {},
    lang: 'de'
};

// ================================
// Translation Helper
// ================================
function t(key) {
    const keys = key.split('.');
    let value = translations[state.lang];
    for (const k of keys) {
        value = value?.[k];
    }
    return value || key;
}

function updateLanguage(lang) {
    state.lang = lang;
    localStorage.setItem('language', lang);
    
    // Update static texts
    document.querySelectorAll('[data-i18n]').forEach(el => {
        const key = el.getAttribute('data-i18n');
        const text = t(key);
        if (text) el.textContent = text;
    });
    
    // Re-render if in app view
    if (state.longData) {
        updateInfoBar();
        if (state.selectedProduct && state.selectedSubtype) {
            renderTabContent();
        }
        if (state.results.length > 0) {
            showExport();
        }
    }
}

// ================================
// Data Parsing
// ================================
function parseExcelData(jsonData) {
    if (!jsonData || jsonData.length < 3) {
        alert(t('errors.minColumns'));
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
        alert(t('errors.noDateCols'));
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
                    Product: String(produkt),
                    Subtype: untertyp ? String(untertyp) : '',
                    Date: new Date(y, m - 1, 1),
                    Consumption: parseFloat(val),
                    MonthStr: `${String(m).padStart(2, '0')}/${y}`
                });
            }
        });
    });

    if (records.length === 0) {
        alert(t('errors.noData'));
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
    subtypeData.sort((a, b) => a.Date - b.Date);
    const values = subtypeData.map(d => d.Consumption);
    
    let forecast;
    try {
        forecast = holtWintersForecast(values, 60);
    } catch (e) {
        forecast = simpleForecast(values, 60);
    }
    
    if (factor !== 1.0) {
        forecast = forecast.map(v => v * factor);
    }
    
    const lastDate = subtypeData[subtypeData.length - 1].Date;
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
    const minDate = new Date(Math.min(...state.longData.map(d => d.Date)));
    const maxDate = new Date(Math.max(...state.longData.map(d => d.Date)));
    const minStr = `${String(minDate.getMonth() + 1).padStart(2, '0')}/${minDate.getFullYear()}`;
    const maxStr = `${String(maxDate.getMonth() + 1).padStart(2, '0')}/${maxDate.getFullYear()}`;
    
    document.getElementById('current-file').textContent = 
        `📄 ${state.currentFile} | ${state.longData.length} ${t('infoBar.records')} | ${t('infoBar.period')}: ${minStr} - ${maxStr}`;
}

function populateProductSelect() {
    const dropdown = document.getElementById('product-dropdown');
    const products = [...new Set(state.longData.map(d => d.Product))].sort();
    
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
    const productData = state.longData.filter(d => d.Product === state.selectedProduct);
    const subtypes = [...new Set(productData.map(d => d.Subtype))].sort();
    
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
        document.getElementById('tab-content').innerHTML = `<p>${t('app.noSubproducts')}</p>`;
    }
}

function renderTabContent() {
    const container = document.getElementById('tab-content');
    const subtypeData = state.longData
        .filter(d => d.Product === state.selectedProduct && d.Subtype === state.selectedSubtype)
        .sort((a, b) => a.Date - b.Date);
    
    if (subtypeData.length === 0) {
        container.innerHTML = `<p>${t('app.noData')}</p>`;
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
                <h4>${t('app.first6Months')}</h4>
                <table class="prognose-table">
                    <tr><th>${t('app.month')}</th><th>${t('app.consumption')}</th></tr>
                    ${first6.map(d => `<tr><td>${d.MonthStr}</td><td>${d.Consumption.toFixed(1)}</td></tr>`).join('')}
                </table>
            </div>
            <div class="table-card">
                <h4>${t('app.last6Months')}</h4>
                <table class="prognose-table">
                    <tr><th>${t('app.month')}</th><th>${t('app.consumption')}</th></tr>
                    ${last6.map(d => `<tr><td>${d.MonthStr}</td><td>${d.Consumption.toFixed(1)}</td></tr>`).join('')}
                </table>
            </div>
        </div>
        
        <div class="chart-container">
            <canvas id="history-chart"></canvas>
        </div>
        
        <div class="controls-row">
            <div class="factor-control">
                <label>${t('app.factor')} ${state.selectedSubtype}:</label>
                <input type="range" id="factor-slider" min="0" max="5" step="0.1" value="${factor}">
                <span class="factor-value" id="factor-value">${factor.toFixed(1)}</span>
                ${factor !== 1.0 ? '<span class="factor-warning">⚠️ ' + factor + 'x</span>' : ''}
            </div>
            <button id="forecast-btn" class="btn primary">${t('app.calculateForecast')}</button>
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
                label: t('app.consumption'),
                data: data.map(d => d.Consumption),
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
    const y1Data = y1.map((v, i) => ({ Month: `${i + 1}/${dates[0].getFullYear()}`, Forecast: v }));
    
    // Year 2-3: quarterly
    const y2Data = [];
    for (let q = 0; q < 4; q++) {
        const sum = values.slice(12 + q * 3, 12 + (q + 1) * 3).reduce((a, b) => a + b, 0);
        y2Data.push({ Quarter: `Q${q + 1} ${dates[12].getFullYear()}`, Forecast: sum });
    }
    const y2Sum = y2Data.reduce((a, b) => a + b.Forecast, 0);
    
    const y3Data = [];
    for (let q = 0; q < 4; q++) {
        const sum = values.slice(24 + q * 3, 24 + (q + 1) * 3).reduce((a, b) => a + b, 0);
        y3Data.push({ Quarter: `Q${q + 1} ${dates[24].getFullYear()}`, Forecast: sum });
    }
    const y3Sum = y3Data.reduce((a, b) => a + b.Forecast, 0);
    
    // Year 4-5: yearly
    const y4 = values.slice(36, 48).reduce((a, b) => a + b, 0);
    const y5 = values.slice(48, 60).reduce((a, b) => a + b, 0);
    const y4Data = [
        { Year: dates[36].getFullYear(), Forecast: y4 },
        { Year: dates[48].getFullYear(), Forecast: y5 }
    ];
    
    // Last year actual
    const lastYear = subtypeData.filter(d => d.Date.getFullYear() === subtypeData[subtypeData.length - 1].Date.getFullYear());
    const lastYearSum = lastYear.reduce((a, b) => a + b.Consumption, 0);
    
    // Percentage changes
    const pctY1 = ((y1Sum - lastYearSum) / lastYearSum) * 100;
    const pctY2 = ((y2Sum - y1Sum) / y1Sum) * 100;
    const pctY3 = ((y3Sum - y2Sum) / y2Sum) * 100;
    const pctY4 = ((y4 - y3Sum) / y3Sum) * 100;
    const pctY5 = ((y5 - y4) / y4) * 100;
    
    const pctData = [
        { Year: 'Year 1', 'Change %': pctY1 },
        { Year: 'Year 2', 'Change %': pctY2 },
        { Year: 'Year 3', 'Change %': pctY3 },
        { Year: 'Year 4', 'Change %': pctY4 },
        { Year: 'Year 5', 'Change %': pctY5 }
    ];
    
    // Render results
    const resultsContainer = document.getElementById('forecast-results');
    resultsContainer.classList.remove('hidden');
    resultsContainer.innerHTML = `
        <div class="results-section">
            <h3>${t('forecast.title')}</h3>
            
            <h4>${t('forecast.year1Monthly')}</h4>
            <table class="prognose-table">
                <tr><th>${t('app.month')}</th><th>Forecast</th></tr>
                ${y1Data.map(d => `<tr><td>${d.Month}</td><td>${d.Forecast.toFixed(1)}</td></tr>`).join('')}
            </table>
            
            <h4>${t('forecast.year2Quarterly')}</h4>
            <table class="prognose-table">
                <tr><th>Quarter</th><th>Forecast</th></tr>
                ${y2Data.map(d => `<tr><td>${d.Quarter}</td><td>${d.Forecast.toFixed(1)}</td></tr>`).join('')}
            </table>
            
            <h4>${t('forecast.year3Quarterly')}</h4>
            <table class="prognose-table">
                <tr><th>Quarter</th><th>Forecast</th></tr>
                ${y3Data.map(d => `<tr><td>${d.Quarter}</td><td>${d.Forecast.toFixed(1)}</td></tr>`).join('')}
            </table>
            
            <h4>${t('forecast.year4Yearly')}</h4>
            <table class="prognose-table">
                <tr><th>${t('forecast.year')}</th><th>Forecast</th></tr>
                ${y4Data.map(d => `<tr><td>${d.Year}</td><td>${d.Forecast.toFixed(1)}</td></tr>`).join('')}
            </table>
            
            <h3>${t('forecast.changeTitle')}</h3>
            <table class="prognose-table">
                <tr><th>${t('forecast.year')}</th><th>${t('forecast.change')}</th></tr>
                ${pctData.map(d => `
                    <tr>
                        <td>${d.Year}</td>
                        <td class="${d['Change %'] >= 0 ? 'positive' : 'negative'}">${d['Change %'] >= 0 ? '+' : ''}${d['Change %'].toFixed(1)}%</td>
                    </tr>
                `).join('')}
            </table>
        </div>
    `;
    
    // Add to results
    const existingIdx = state.results.findIndex(r => r.Product === state.selectedProduct && r.Subtype === state.selectedSubtype);
    const result = {
        Product: state.selectedProduct,
        Subtype: state.selectedSubtype,
        Factor: factor,
        'Year 1 (monthly)': Math.round(y1Sum * 10) / 10,
        'Year 2 (Quarter)': Math.round(y2Sum * 10) / 10,
        'Year 3 (Quarter)': Math.round(y3Sum * 10) / 10,
        'Year 4': Math.round(y4 * 10) / 10,
        'Year 5': Math.round(y5 * 10) / 10,
        'Change Year 1 %': Math.round(pctY1 * 10) / 10,
        'Change Year 2 %': Math.round(pctY2 * 10) / 10,
        'Change Year 3 %': Math.round(pctY3 * 10) / 10,
        'Change Year 4 %': Math.round(pctY4 * 10) / 10,
        'Change Year 5 %': Math.round(pctY5 * 10) / 10
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
        container.innerHTML = `<p>${t('export.noResults')}</p>`;
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
    XLSX.utils.book_append_sheet(wb, ws, "Forecast");
    XLSX.writeFile(wb, "forecast.xlsx");
}

// ================================
// Theme & Language Toggle
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
                .filter(d => d.Product === state.selectedProduct && d.Subtype === state.selectedSubtype)
                .sort((a, b) => a.Date - b.Date);
            renderHistoryChart(subtypeData);
        }
    };
}

function initLanguage() {
    const select = document.getElementById('language-select');
    const saved = localStorage.getItem('language');
    if (saved && (saved === 'de' || saved === 'en')) {
        select.value = saved;
        state.lang = saved;
    }
    
    select.onchange = () => {
        updateLanguage(select.value);
    };
}

// ================================
// Event Listeners
// ================================
document.addEventListener('DOMContentLoaded', () => {
    initTheme();
    initLanguage();
    
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
