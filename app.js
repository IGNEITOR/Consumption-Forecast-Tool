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
    lang: 'en',
    collapse: {
        history: true,
        forecast: true
    },
    view: null
};

// ================================
// Constants & Helpers
// ================================
const DEFAULT_LANG = 'en';
const STORAGE_KEYS = {
    language: 'language',
    theme: 'theme'
};
const FORECAST_MONTHS = 60;
const MONTHS_IN_YEAR = 12;
const QUARTER_MONTHS = 3;
const SEASONAL_LENGTH = 12;

const $ = (selector) => document.querySelector(selector);
const $$ = (selector) => Array.from(document.querySelectorAll(selector));
const formatMonthString = (date) => `${String(date.getMonth() + 1).padStart(2, '0')}/${date.getFullYear()}`;
const formatNumber = (value, decimals = 1) => Number.isFinite(value) ? value.toFixed(decimals) : '–';
const sum = (arr) => arr.reduce((acc, val) => acc + val, 0);
const chunkSum = (arr, start, length) => sum(arr.slice(start, start + length));
const round1 = (value) => Math.round(value * 10) / 10;
const percentageChange = (current, previous) => {
    if (!Number.isFinite(previous) || previous === 0) return null;
    return ((current - previous) / previous) * 100;
};

const exportHeaders = () => ([
    { key: 'Product', label: t('export.headers.product') },
    { key: 'Subtype', label: t('export.headers.subtype') },
    { key: 'Factor', label: t('export.headers.factor') },
    { key: 'Year 1 (monthly)', label: t('export.headers.year1Monthly') },
    { key: 'Year 2 (Quarter)', label: t('export.headers.year2Quarter') },
    { key: 'Year 3 (Quarter)', label: t('export.headers.year3Quarter') },
    { key: 'Year 4', label: t('export.headers.year4') },
    { key: 'Year 5', label: t('export.headers.year5') },
    { key: 'Change Year 1 %', label: t('export.headers.change1') },
    { key: 'Change Year 2 %', label: t('export.headers.change2') },
    { key: 'Change Year 3 %', label: t('export.headers.change3') },
    { key: 'Change Year 4 %', label: t('export.headers.change4') },
    { key: 'Change Year 5 %', label: t('export.headers.change5') }
]);

function moveResult(index, direction) {
    const newIndex = index + direction;
    if (newIndex < 0 || newIndex >= state.results.length) return;
    const copy = [...state.results];
    const [item] = copy.splice(index, 1);
    copy.splice(newIndex, 0, item);
    state.results = copy;
    showExport();
}

function deleteResult(index) {
    state.results.splice(index, 1);
    showExport();
}

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
    localStorage.setItem(STORAGE_KEYS.language, lang);
    
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
    if (!Array.isArray(jsonData) || jsonData.length < 3) {
        alert(t('errors.minColumns'));
        return null;
    }

    const headers = Object.keys(jsonData[0]);
    const productCol = headers.find((h) => /produkt|product/i.test(h)) || headers[0];
    const subtypeCol = headers.find((h) => /untertyp|subtype|typ/i.test(h)) || headers[1];

    const dateCols = headers
        .filter((h) => h !== productCol && h !== subtypeCol)
        .filter((h) => /^([0-1]?\d)\/\d{4}$/.test(h))
        .map((h) => {
            const [month, year] = h.split('/').map(Number);
            return month >= 1 && month <= 12 ? { header: h, month, year } : null;
        })
        .filter(Boolean)
        .sort((a, b) => (a.year !== b.year ? a.year - b.year : a.month - b.month));

    if (dateCols.length === 0) {
        alert(t('errors.noDateCols'));
        return null;
    }

    const records = jsonData.reduce((acc, row) => {
        const product = row[productCol];
        if (!product) return acc;

        const subtype = row[subtypeCol] ? String(row[subtypeCol]) : '';

        dateCols.forEach(({ header, month, year }) => {
            const rawValue = row[header];
            if (rawValue === undefined || rawValue === null || rawValue === '') return;

            const numeric = Number(rawValue);
            if (!Number.isFinite(numeric)) return;

            acc.push({
                Product: String(product),
                Subtype: subtype,
                Date: new Date(year, month - 1, 1),
                Consumption: numeric,
                MonthStr: `${String(month).padStart(2, '0')}/${year}`
            });
        });

        return acc;
    }, []);

    if (records.length === 0) {
        alert(t('errors.noData'));
        return null;
    }

    return records;
}

// ================================
// Forecasting - Holt-Winters Triple Exponential Smoothing
// ================================
function holtWintersForecast(data, periods = FORECAST_MONTHS, alpha = 0.3, beta = 0.1, gamma = 0.1, seasonalLength = SEASONAL_LENGTH) {
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

    const denominator = n * sumXX - sumX * sumX;
    const slope = denominator !== 0 ? (n * sumXY - sumX * sumY) / denominator : 0;
    const intercept = n !== 0 ? (sumY - slope * sumX) / n : 0;

    return Array.from({ length: periods }, (_, idx) => {
        const value = intercept + slope * (n + idx);
        return Math.max(0, value);
    });
}

function makeForecast(subtypeData, factor = 1.0) {
    subtypeData.sort((a, b) => a.Date - b.Date);
    const values = subtypeData.map(d => d.Consumption);

    let forecast;
    try {
        forecast = holtWintersForecast(values, FORECAST_MONTHS);
    } catch (e) {
        forecast = simpleForecast(values, FORECAST_MONTHS);
    }

    const adjusted = factor !== 1.0 ? forecast.map(v => v * factor) : forecast;
    const lastDate = subtypeData[subtypeData.length - 1].Date;
    const forecastDates = Array.from({ length: FORECAST_MONTHS }, (_, idx) => {
        const date = new Date(lastDate);
        date.setMonth(date.getMonth() + idx + 1);
        return date;
    });

    return { dates: forecastDates, values: adjusted };
}

// ================================
// UI Functions
// ================================
function showApp() {
    $('#upload-section').classList.add('hidden');
    $('#app-section').classList.remove('hidden');
}

function hideApp() {
    $('#upload-section').classList.remove('hidden');
    $('#app-section').classList.add('hidden');
    $('#export-section').classList.add('hidden');
}

function showUpload() {
    $('#app-section').classList.add('hidden');
    $('#upload-section').classList.remove('hidden');
}

function updateInfoBar() {
    if (!state.longData || state.longData.length === 0) return;

    const minDate = new Date(Math.min(...state.longData.map(d => d.Date)));
    const maxDate = new Date(Math.max(...state.longData.map(d => d.Date)));
    const minStr = formatMonthString(minDate);
    const maxStr = formatMonthString(maxDate);

    $('#current-file').textContent =
        `📄 ${state.currentFile} | ${state.longData.length} ${t('infoBar.records')} | ${t('infoBar.period')}: ${minStr} - ${maxStr}`;
}

function populateProductSelect() {
    const dropdown = $('#product-dropdown');
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
    const container = $('#tabs-container');
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
        $('#tab-content').innerHTML = `<p>${t('app.noSubproducts')}</p>`;
    }
}

function renderTabContent() {
    const container = $('#tab-content');
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
    const factor = state.factors[key] ?? 1.0;
    
    container.innerHTML = `
        <h3>${state.selectedProduct} - ${state.selectedSubtype}</h3>

        <details id="history-details" class="collapsible" ${state.collapse.history ? 'open' : ''}>
            <summary>${t('app.first6Months')} / ${t('app.last6Months')} & ${t('app.consumption')}</summary>
            <div class="tables-row">
                <div class="table-card">
                    <h4>${t('app.first6Months')}</h4>
                    <table class="prognose-table">
                        <tr><th>${t('app.month')}</th><th>${t('app.consumption')}</th></tr>
                        ${first6.map(d => `<tr><td>${d.MonthStr}</td><td>${formatNumber(d.Consumption)}</td></tr>`).join('')}
                    </table>
                </div>
                <div class="table-card">
                    <h4>${t('app.last6Months')}</h4>
                    <table class="prognose-table">
                        <tr><th>${t('app.month')}</th><th>${t('app.consumption')}</th></tr>
                        ${last6.map(d => `<tr><td>${d.MonthStr}</td><td>${formatNumber(d.Consumption)}</td></tr>`).join('')}
                    </table>
                </div>
            </div>

            <div class="chart-container">
                <canvas id="history-chart"></canvas>
            </div>
        </details>

        <div class="controls-row">
            <div class="factor-control">
                <label>${t('app.factor')} ${state.selectedSubtype}:</label>
                <input type="range" id="factor-slider" min="0" max="5" step="0.1" value="${factor}">
                <span class="factor-value" id="factor-value">${factor.toFixed(1)}</span>
                ${factor !== 1.0 ? '<span class="factor-warning">⚠️ ' + factor + 'x</span>' : ''}
            </div>
            <button id="forecast-btn" class="btn primary">${t('app.calculateForecast')}</button>
        </div>
        <details id="forecast-details" class="collapsible" ${state.collapse.forecast ? 'open' : ''}>
            <summary>${t('forecast.tablesTitle')}</summary>
            <div id="forecast-results" class="hidden"></div>
        </details>
    `;
    
    renderHistoryChart(subtypeData);

    const historyDetails = $('#history-details');
    if (historyDetails) {
        historyDetails.ontoggle = () => {
            state.collapse.history = historyDetails.open;
        };
    }
    
    const forecastDetails = $('#forecast-details');
    if (forecastDetails) {
        forecastDetails.ontoggle = () => {
            state.collapse.forecast = forecastDetails.open;
        };
    }
    
    const slider = $('#factor-slider');
    const factorValue = $('#factor-value');
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
    
    $('#forecast-btn').onclick = () => calculateForecast(subtypeData, key);
}

function renderHistoryChart(data) {
    const canvas = $('#history-chart');
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    
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

function buildQuarterlyData(values, offset, dates) {
    const baseYear = dates[offset]?.getFullYear() ?? dates[0]?.getFullYear() + Math.floor(offset / MONTHS_IN_YEAR);
    return Array.from({ length: 4 }, (_, q) => {
        const start = offset + q * QUARTER_MONTHS;
        const total = chunkSum(values, start, QUARTER_MONTHS);
        return { label: `Q${q + 1} ${baseYear}`, total };
    });
}

function renderForecastTables({ y1Data, y2Data, y3Data, y4Data, pctData }) {
    return `
        <div class="results-section">
            <h4>${t('forecast.year1Monthly')}</h4>
            <table class="prognose-table">
                <tr><th>${t('app.month')}</th><th>${t('forecast.value')}</th></tr>
                ${y1Data.map(d => `<tr><td>${d.Month}</td><td>${formatNumber(d.Forecast)}</td></tr>`).join('')}
            </table>

            <h4>${t('forecast.year2Quarterly')}</h4>
            <table class="prognose-table">
                <tr><th>Quarter</th><th>${t('forecast.value')}</th></tr>
                ${y2Data.map(d => `<tr><td>${d.Quarter}</td><td>${formatNumber(d.Forecast)}</td></tr>`).join('')}
            </table>

            <h4>${t('forecast.year3Quarterly')}</h4>
            <table class="prognose-table">
                <tr><th>Quarter</th><th>${t('forecast.value')}</th></tr>
                ${y3Data.map(d => `<tr><td>${d.Quarter}</td><td>${formatNumber(d.Forecast)}</td></tr>`).join('')}
            </table>

            <h4>${t('forecast.year4Yearly')}</h4>
            <table class="prognose-table">
                <tr><th>${t('forecast.year')}</th><th>${t('forecast.value')}</th></tr>
                ${y4Data.map(d => `<tr><td>${d.Year}</td><td>${formatNumber(d.Forecast)}</td></tr>`).join('')}
            </table>

            <h4>${t('forecast.changeTitle')}</h4>
            <table class="prognose-table">
                <tr><th>${t('forecast.year')}</th><th>${t('forecast.change')}</th></tr>
                ${pctData.map(d => {
                    const isPositive = typeof d.change === 'number' && d.change >= 0;
                    const isNegative = typeof d.change === 'number' && d.change < 0;
                    const className = isPositive ? 'positive' : isNegative ? 'negative' : '';
                    const value = typeof d.change === 'number'
                        ? `${isPositive ? '+' : ''}${formatNumber(d.change)}%`
                        : '–';

                    return `
                        <tr>
                            <td>${d.Year}</td>
                            <td class="${className}">${value}</td>
                        </tr>
                    `;
                }).join('')}
            </table>
        </div>
    `;
}

function calculateForecast(subtypeData, key) {
    const factor = state.factors[key] ?? 1.0;
    const { dates, values } = makeForecast(subtypeData, factor);

    // Year 1: monthly
    const y1Values = values.slice(0, MONTHS_IN_YEAR);
    const y1Sum = sum(y1Values);
    const y1Data = y1Values.map((v, i) => ({ Month: `${i + 1}/${dates[0].getFullYear()}`, Forecast: v }));

    // Year 2-3: quarterly
    const y2Quarters = buildQuarterlyData(values, MONTHS_IN_YEAR, dates);
    const y2Sum = sum(y2Quarters.map(q => q.total));
    const y2Data = y2Quarters.map(q => ({ Quarter: q.label, Forecast: q.total }));

    const y3Quarters = buildQuarterlyData(values, MONTHS_IN_YEAR * 2, dates);
    const y3Sum = sum(y3Quarters.map(q => q.total));
    const y3Data = y3Quarters.map(q => ({ Quarter: q.label, Forecast: q.total }));

    // Year 4-5: yearly
    const y4 = chunkSum(values, MONTHS_IN_YEAR * 3, MONTHS_IN_YEAR);
    const y5 = chunkSum(values, MONTHS_IN_YEAR * 4, MONTHS_IN_YEAR);
    const y4Data = [
        { Year: dates[MONTHS_IN_YEAR * 3].getFullYear(), Forecast: y4 },
        { Year: dates[MONTHS_IN_YEAR * 4].getFullYear(), Forecast: y5 }
    ];

    // Last year actual
    const latestYear = subtypeData[subtypeData.length - 1].Date.getFullYear();
    const lastYearSum = sum(subtypeData.filter(d => d.Date.getFullYear() === latestYear).map(d => d.Consumption));

    // Percentage changes
    const pctData = [
        { Year: 'Year 1', change: percentageChange(y1Sum, lastYearSum) },
        { Year: 'Year 2', change: percentageChange(y2Sum, y1Sum) },
        { Year: 'Year 3', change: percentageChange(y3Sum, y2Sum) },
        { Year: 'Year 4', change: percentageChange(y4, y3Sum) },
        { Year: 'Year 5', change: percentageChange(y5, y4) }
    ];

    const resultsContainer = $('#forecast-results');
    resultsContainer.classList.remove('hidden');
    resultsContainer.innerHTML = renderForecastTables({ y1Data, y2Data, y3Data, y4Data, pctData });
    const forecastDetails = $('#forecast-details');
    if (forecastDetails && !forecastDetails.open && state.collapse.forecast) {
        forecastDetails.open = true;
    }


    const existingIdx = state.results.findIndex(r => r.Product === state.selectedProduct && r.Subtype === state.selectedSubtype);
    const toRounded = (value) => value == null ? null : round1(value);

    const result = {
        Product: state.selectedProduct,
        Subtype: state.selectedSubtype,
        Factor: factor,
        'Year 1 (monthly)': round1(y1Sum),
        'Year 2 (Quarter)': round1(y2Sum),
        'Year 3 (Quarter)': round1(y3Sum),
        'Year 4': round1(y4),
        'Year 5': round1(y5),
        'Change Year 1 %': toRounded(pctData[0].change),
        'Change Year 2 %': toRounded(pctData[1].change),
        'Change Year 3 %': toRounded(pctData[2].change),
        'Change Year 4 %': toRounded(pctData[3].change),
        'Change Year 5 %': toRounded(pctData[4].change)
    };

        state.results.push(result);

    showExport();
}

function showExport() {
    const section = $('#export-section');
    const container = $('#results-table-container');

    section.classList.remove('hidden');

    if (state.results.length === 0) {
        container.innerHTML = `<p>${t('export.noResults')}</p>`;
        return;
    }

    const headers = exportHeaders();
    const formatCell = (value) => typeof value === 'number' ? formatNumber(value) : (value ?? '–');
    container.innerHTML = `
        <details class="collapsible" open>
            <summary>${t('export.downloadBtn')}</summary>
            <div class="table-wrapper">
                <table class="prognose-table">
                    <tr>
                        <th></th>
                        ${headers.map(h => `<th>${h.label}</th>`).join('')}
                        <th></th>
                    </tr>
                    ${state.results.map((r, idx) => `
                        <tr>
                            <td class="row-actions">
                                <button class="icon-btn" data-action="up" data-idx="${idx}" ${idx === 0 ? 'disabled' : ''}>↑</button>
                                <button class="icon-btn" data-action="down" data-idx="${idx}" ${idx === state.results.length - 1 ? 'disabled' : ''}>↓</button>
                            </td>
                            ${headers.map(h => `<td>${formatCell(r[h.key])}</td>`).join('')}
                            <td class="row-actions">
                                <button class="icon-btn danger" data-action="delete" data-idx="${idx}">✕</button>
                            </td>
                        </tr>
                    `).join('')}
                </table>
            </div>
        </details>
    `;

    // Wire row action buttons
    $$('.icon-btn').forEach(btn => {
        const action = btn.dataset.action;
        const idx = Number(btn.dataset.idx);
        if (action === 'up') btn.onclick = () => moveResult(idx, -1);
        if (action === 'down') btn.onclick = () => moveResult(idx, 1);
        if (action === 'delete') btn.onclick = () => deleteResult(idx);
    });
}

function exportToExcel() {
    if (state.results.length === 0) return;

    const headers = exportHeaders();
    const localizedRows = state.results.map(result => {
        const row = {};
        headers.forEach(h => {
            row[h.label] = result[h.key];
        });
        return row;
    });

    const ws = XLSX.utils.json_to_sheet(localizedRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, t('forecast.title'));
    XLSX.writeFile(wb, `${t('forecast.title').toLowerCase().replace(/\s+/g, '_')}.xlsx`);
}

function processFile(file) {
    if (!file) return;

    $('#filename').textContent = file.name;

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
            state.collapse = { history: true, forecast: true };
            updateInfoBar();
            populateProductSelect();
            showApp();
        }
    };

    reader.readAsArrayBuffer(file);
}

// ================================
// Theme & Language Toggle
// ================================
function initTheme() {
    const toggle = $('#theme-toggle');
    const icon = $('#theme-icon');

    const saved = localStorage.getItem(STORAGE_KEYS.theme);
    if (saved === 'dark') {
        document.body.dataset.theme = 'dark';
        icon.textContent = '☀️';
    }

    toggle.onclick = () => {
        if (document.body.dataset.theme === 'dark') {
            delete document.body.dataset.theme;
            localStorage.setItem(STORAGE_KEYS.theme, 'light');
            icon.textContent = '🌙';
        } else {
            document.body.dataset.theme = 'dark';
            localStorage.setItem(STORAGE_KEYS.theme, 'dark');
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
    const select = $('#language-select');
    const saved = localStorage.getItem(STORAGE_KEYS.language);
    const initialLang = saved === 'de' || saved === 'en' ? saved : DEFAULT_LANG;

    select.value = initialLang;
    state.lang = initialLang;
    localStorage.setItem(STORAGE_KEYS.language, initialLang);

    updateLanguage(state.lang);

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

    const fileInput = $('#file-input');
    $('#upload-btn').onclick = () => fileInput.click();
    fileInput.onchange = () => processFile(fileInput.files[0]);

    $('#change-file-btn').onclick = () => fileInput.click();
    $('#export-btn').onclick = exportToExcel;
});
