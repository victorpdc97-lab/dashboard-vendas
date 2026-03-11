// ===================== CONFIGURAÇÃO =====================
const DEFAULT_VENDORS = [
    { name: 'Leonardo', id: '1GTBw9T4Y3W8of3faSXgxnAZzZAgaZPmBpYDg8tmmAEM' },
    { name: 'Diogo',    id: '1lrWS7Jhn_70RoVpKLeDLgRhVTyLYVXrQm9bsY-iRMZ8' }
];
const MONTHS_ALT = ['JANEIRO','FEVEREIRO','MARÇO','ABRIL','MAIO','JUNHO','JULHO','AGOSTO','SETEMBRO','OUTUBRO','NOVEMBRO','DEZEMBRO'];
const WEEKDAYS = ['Dom', 'Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb'];
const REFRESH_MS = 60000;

// ===================== STATE =====================
let vendors = loadVendors();
let charts = {};
let refreshTimer = null;
let compareMode = false;
let channelFilter = 'all'; // 'all' | 'loja' | 'site'
let lastResults = null;
let lastPrevResults = null;
let activeTab = 'visao-geral';
let fullscreenChart = null;
let yearlyFetched = false;

// ===================== SVG ICONS =====================
const ICON_SUN = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="5"/><line x1="12" y1="1" x2="12" y2="3"/><line x1="12" y1="21" x2="12" y2="23"/><line x1="4.22" y1="4.22" x2="5.64" y2="5.64"/><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"/><line x1="1" y1="12" x2="3" y2="12"/><line x1="21" y1="12" x2="23" y2="12"/><line x1="4.22" y1="19.78" x2="5.64" y2="18.36"/><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"/></svg>';
const ICON_MOON = '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 12.79A9 9 0 1111.21 3 7 7 0 0021 12.79z"/></svg>';
const ICON_DOLLAR = '<svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="12" y1="1" x2="12" y2="23"/><path d="M17 5H9.5a3.5 3.5 0 000 7h5a3.5 3.5 0 010 7H6"/></svg>';
const ICON_TICKET = '<svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M20 12V8H6a2 2 0 01-2-2c0-1.1.9-2 2-2h12v4"/><path d="M4 6v12c0 1.1.9 2 2 2h12v-4"/><path d="M18 12a2 2 0 000 4h4v-4z"/></svg>';
const ICON_USERS = '<svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 00-3-3.87"/><path d="M16 3.13a4 4 0 010 7.75"/></svg>';
const ICON_USER = '<svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M20 21v-2a4 4 0 00-4-4H8a4 4 0 00-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>';
const ICON_ALERT = '<svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>';
const ICON_FORECAST = '<svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/></svg>';
const ICON_EXPAND = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="15 3 21 3 21 9"/><polyline points="9 21 3 21 3 15"/><line x1="21" y1="3" x2="14" y2="10"/><line x1="3" y1="21" x2="10" y2="14"/></svg>';

// ===================== UTILIDADES =====================
function parseBR(s) {
    if (s == null) return null;
    if (typeof s === 'number') return s;
    s = String(s).trim();
    if (!s || s === '#DIV/0!' || s === '#REF!' || s === '-') return null;
    const n = parseFloat(s.replace(/\./g, '').replace(',', '.'));
    return isNaN(n) ? null : n;
}

function cellNum(cell) {
    if (!cell || cell.v == null) return null;
    if (typeof cell.v === 'number') return cell.v;
    return parseBR(String(cell.v));
}

function isDateStr(s) { return /^\d{2}\/\d{2}/.test(String(s || '')); }

function fmtBRL(v) {
    return 'R$ ' + v.toFixed(2).replace('.', ',').replace(/\B(?=(\d{3})+(?!\d))/g, '.');
}
function fmtInt(v) { return Math.round(v).toLocaleString('pt-BR'); }

function showToast(message, type = 'info') {
    const container = document.getElementById('toastContainer');
    if (!container) return;
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    toast.textContent = message;
    container.appendChild(toast);
    setTimeout(() => { toast.classList.add('toast-exit'); }, 2500);
    setTimeout(() => { if (toast.parentNode) toast.remove(); }, 2800);
}

function animateChildren(parent, selector) {
    if (!parent) return;
    parent.querySelectorAll(selector).forEach((el, i) => {
        el.style.animationDelay = (i * 0.06) + 's';
        el.classList.add('animate-in');
    });
}

function trendHTML(current, previous) {
    if (previous == null || previous === 0) return '';
    const pct = ((current - previous) / previous) * 100;
    if (Math.abs(pct) < 0.5) return '<span class="trend trend-neutral">= 0%</span>';
    const cls = pct > 0 ? 'trend-up' : 'trend-down';
    const arrow = pct > 0 ? '&#9650;' : '&#9660;';
    return `<span class="trend ${cls}">${arrow} ${Math.abs(pct).toFixed(1)}%</span>`;
}

// ===================== CHANNEL FILTER =====================
function getFiltered(data) {
    if (channelFilter === 'loja') return { total: data.lojaTotal, clients: data.lojaClients, ticket: data.lojaTicket };
    if (channelFilter === 'site') return { total: data.siteTotal, clients: data.siteClients, ticket: data.siteTicket };
    return { total: data.lojaTotal + data.siteTotal, clients: data.lojaClients + data.siteClients, ticket: (data.lojaClients + data.siteClients) > 0 ? (data.lojaTotal + data.siteTotal) / (data.lojaClients + data.siteClients) : 0 };
}

function setChannelFilter(filter) {
    channelFilter = filter;
    document.querySelectorAll('.filter-btn').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.filter === filter);
    });
    if (lastResults) reRenderAll();
    updateURLParams();
}

function reRenderAll() {
    renderGlobalKpis(lastResults, lastPrevResults);
    renderPodium(lastResults);
    renderParticipation(lastResults);
    if (compareMode) return;
    const oldPos = captureVendorPositions();
    renderVendorCards(lastResults, lastPrevResults);
    animateVendorReorder(oldPos);
    if (activeTab === 'vendedores') {
        setTimeout(() => renderVendorInlineCharts(lastResults), 0);
        renderVendorComparator(lastResults);
    }
    if (activeTab === 'graficos') {
        renderCharts(lastResults);
        renderWeekdayChart(lastResults);
        renderHeatmap(lastResults);
        addExpandButtons();
    }
    animateCountUp();
}

// ===================== TABS =====================
function switchTab(tabId) {
    activeTab = tabId;
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.tab === tabId);
    });
    document.querySelectorAll('.tab-content').forEach(tc => {
        tc.classList.toggle('active', tc.id === 'tab-' + tabId);
    });
    const tabEl = document.getElementById('tab-' + tabId);
    if (tabEl) {
        tabEl.style.animation = 'none';
        void tabEl.offsetWidth;
        tabEl.style.animation = 'fadeSlideIn 0.3s ease';
    }
    if (lastResults) {
        if (tabId === 'graficos') {
            renderCharts(lastResults);
            renderWeekdayChart(lastResults);
            renderHeatmap(lastResults);
            if (!yearlyFetched) fetchAndRenderYearly();
            setTimeout(() => addExpandButtons(), 50);
        }
        if (tabId === 'vendedores') {
            setTimeout(() => renderVendorInlineCharts(lastResults), 0);
            renderVendorComparator(lastResults);
        }
    }
    updateURLParams();
}

// ===================== URL PARAMS / SHARE =====================
function updateURLParams() {
    const params = new URLSearchParams();
    const m = document.getElementById('month');
    const y = document.getElementById('year');
    if (m && y) {
        params.set('month', m.value);
        params.set('year', y.value);
    }
    params.set('tab', activeTab);
    params.set('filter', channelFilter);
    const newUrl = window.location.pathname + '?' + params.toString();
    window.history.replaceState({}, '', newUrl);
}

function readURLParams() {
    const params = new URLSearchParams(window.location.search);
    const month = params.get('month');
    const year = params.get('year');
    const tab = params.get('tab');
    const filter = params.get('filter');

    if (year) {
        const yel = document.getElementById('year');
        yel.value = year;
        yel.dispatchEvent(new Event('change'));
    }
    if (month) document.getElementById('month').value = month;
    if (filter && ['all', 'loja', 'site'].includes(filter)) {
        channelFilter = filter;
        document.querySelectorAll('.filter-btn').forEach(btn => {
            btn.classList.toggle('active', btn.dataset.filter === filter);
        });
    }
    if (tab && ['visao-geral', 'vendedores', 'graficos'].includes(tab)) {
        switchTab(tab);
    }
}

function shareLink() {
    updateURLParams();
    navigator.clipboard.writeText(window.location.href).then(() => {
        showToast('Link copiado para a área de transferência!', 'success');
    }).catch(() => {
        prompt('Copie o link:', window.location.href);
    });
}

// ===================== THEME =====================
function getThemeColors() {
    const s = getComputedStyle(document.documentElement);
    return { text2: s.getPropertyValue('--chart-text').trim(), grid: s.getPropertyValue('--chart-grid').trim() };
}

function applyTheme(theme) {
    document.documentElement.setAttribute('data-theme', theme);
    localStorage.setItem('dashboard-theme', theme);
    document.getElementById('themeToggle').innerHTML = theme === 'dark' ? ICON_SUN : ICON_MOON;
    if (lastResults) {
        if (activeTab === 'graficos') {
            renderCharts(lastResults);
            renderWeekdayChart(lastResults);
        }
        if (activeTab === 'vendedores') {
            renderVendorInlineCharts(lastResults);
        }
    }
}

function toggleTheme() {
    const current = document.documentElement.getAttribute('data-theme') || 'dark';
    applyTheme(current === 'dark' ? 'light' : 'dark');
}

// ===================== VENDOR MANAGEMENT =====================
function loadVendors() {
    const saved = localStorage.getItem('dashboard-vendors');
    if (saved) { try { return JSON.parse(saved); } catch(e) {} }
    return DEFAULT_VENDORS.map(v => ({...v}));
}

function saveVendors() {
    localStorage.setItem('dashboard-vendors', JSON.stringify(vendors));
}

function openVendorModal() {
    document.getElementById('vendorName').value = '';
    document.getElementById('vendorSheetId').value = '';
    document.getElementById('vendorModal').style.display = 'flex';
}

function closeVendorModal() {
    document.getElementById('vendorModal').style.display = 'none';
}

function addVendor() {
    const name = document.getElementById('vendorName').value.trim();
    const id = document.getElementById('vendorSheetId').value.trim();
    if (!name || !id) return alert('Preencha todos os campos.');
    vendors.push({ name, id });
    saveVendors();
    closeVendorModal();
    refreshData();
    showToast(`Vendedor "${name}" adicionado!`, 'success');
}

function removeVendor(index) {
    const name = vendors[index].name;
    if (!confirm('Remover vendedor "' + name + '"?')) return;
    vendors.splice(index, 1);
    saveVendors();
    refreshData();
    showToast(`Vendedor "${name}" removido`, 'info');
}

// ===================== FETCH VIA JSONP =====================
function fetchSheet(sheetId, sheetName) {
    return new Promise((resolve, reject) => {
        const cb = '_cb_' + Math.random().toString(36).substr(2, 8);
        const timeout = setTimeout(() => { cleanup(); reject(new Error('Tempo esgotado')); }, 15000);
        const script = document.createElement('script');
        function cleanup() { clearTimeout(timeout); delete window[cb]; if (script.parentNode) script.remove(); }
        window[cb] = function(resp) {
            cleanup();
            if (resp.status === 'error') { reject(new Error(resp.errors?.[0]?.detailed_message || 'Erro')); return; }
            resolve(resp.table);
        };
        let url = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:json;responseHandler:${cb}`;
        if (sheetName) url += `&sheet=${encodeURIComponent(sheetName)}`;
        script.src = url;
        script.onerror = () => { cleanup(); reject(new Error('Falha ao carregar planilha')); };
        document.body.appendChild(script);
    });
}

function getSheetNameFor(monthIdx, year) {
    const yy = String(year).slice(-2);
    return MONTHS_ALT[monthIdx] + ' ' + yy;
}

function getSheetName() {
    return getSheetNameFor(
        parseInt(document.getElementById('month').value),
        parseInt(document.getElementById('year').value)
    );
}

async function fetchForVendorWithSheet(vendorId, sheetName) {
    const variants = [sheetName];
    const noAccent = sheetName.replace('Ç', 'C').replace('ç', 'c');
    if (noAccent !== sheetName) variants.push(noAccent);
    let lastErr;
    for (const name of variants) {
        try { return await fetchSheet(vendorId, name); } catch (e) { lastErr = e; }
    }
    throw lastErr || new Error('Aba não encontrada');
}

function getPreviousMonthInfo() {
    const m = parseInt(document.getElementById('month').value);
    const y = parseInt(document.getElementById('year').value);
    return m === 0 ? { monthIdx: 11, year: y - 1 } : { monthIdx: m - 1, year: y };
}

async function fetchMonthData(sheetName) {
    return Promise.all(vendors.map(async (vendor, idx) => {
        try {
            const table = await fetchForVendorWithSheet(vendor.id, sheetName);
            const data = processTable(table);
            return { name: vendor.name, data, originalIdx: idx };
        } catch {
            return { name: vendor.name, error: 'Aba "' + sheetName + '" não encontrada nesta planilha.', originalIdx: idx };
        }
    }));
}

// ===================== DETECÇÃO DE COLUNAS =====================
function detectColumns(table) {
    const cols = table.cols || [];
    const rows = table.rows || [];
    const labels = cols.map(c => (c.label || '').toUpperCase().trim());
    const firstRowTexts = [];
    for (let r = 0; r < Math.min(3, rows.length); r++) {
        const cells = rows[r]?.c;
        if (cells) firstRowTexts.push(cells.map(c => String(c?.v ?? c?.f ?? '').toUpperCase().trim()));
    }
    const allHeaders = [labels, ...firstRowTexts];

    let lojaSections = [], siteSections = [];
    for (const headers of allHeaders) {
        for (let i = 0; i < headers.length; i++) {
            if (headers[i].includes('LOJA') && !lojaSections.includes(i)) lojaSections.push(i);
            if (headers[i].includes('SITE') && !siteSections.includes(i)) siteSections.push(i);
        }
        if (lojaSections.length > 0 && siteSections.length > 0) break;
    }

    function colHasNumbers(colIdx) {
        let numCount = 0;
        for (let r = 0; r < Math.min(40, rows.length); r++) {
            const cell = rows[r]?.c?.[colIdx];
            if (!cell || cell.v == null) continue;
            if (typeof cell.v === 'number') { numCount++; continue; }
            const s = String(cell.v).trim();
            if (isDateStr(s)) continue;
            if (parseBR(s) !== null) numCount++;
        }
        return numCount >= 2;
    }

    let lojaCols = [], siteCols = [];
    for (const headers of allHeaders) {
        for (let i = 0; i < headers.length; i++) {
            if ((/^LOJA\s*F[IÍ]SICA$/.test(headers[i]) || /^META\s*MASTER$/.test(headers[i])) && colHasNumbers(i) && !lojaCols.includes(i)) lojaCols.push(i);
            if (/^SITE$/.test(headers[i]) && colHasNumbers(i) && !siteCols.includes(i)) siteCols.push(i);
        }
        if (lojaCols.length > 0 && siteCols.length > 0) break;
    }

    if (lojaCols.length > 0 || siteCols.length > 0) {
        if (siteCols.length === 0) {
            for (const headers of allHeaders) {
                for (let i = 0; i < headers.length; i++) {
                    if (/^SITE$/.test(headers[i]) && !siteCols.includes(i)) siteCols.push(i);
                }
            }
        }
        return { lojaCols, siteCols };
    }

    lojaSections.sort((a, b) => a - b);
    siteSections.sort((a, b) => a - b);
    const lojaStart = lojaSections[0] ?? -1;
    const siteStart = siteSections[0] ?? -1;
    if (lojaStart < 0 || siteStart < 0) return { lojaCols: [], siteCols: [] };

    let afterSite = cols.length;
    for (const headers of allHeaders) {
        for (let i = siteStart + 1; i < headers.length; i++) {
            const h = headers[i];
            if (h.includes('ESTORNO') || h.includes('PARCERIA') || h.includes('BARBE') ||
                (h === 'DATA' && i > siteStart + 1)) {
                afterSite = Math.min(afterSite, i);
            }
        }
    }

    lojaCols = [];
    for (let i = lojaStart + 1; i < siteStart; i++) { if (colHasNumbers(i)) lojaCols.push(i); }
    siteCols = [];
    for (let i = siteStart + 1; i < afterSite; i++) { if (colHasNumbers(i)) siteCols.push(i); }
    return { lojaCols, siteCols };
}

// ===================== PROCESSAMENTO =====================
function processTable(table) {
    const rows = table.rows || [];
    const { lojaCols, siteCols } = detectColumns(table);

    if (lojaCols.length === 0 && siteCols.length === 0) {
        return { lojaTotal: 0, lojaClients: 0, lojaTicket: 0, siteTotal: 0, siteClients: 0, siteTicket: 0, dailyData: [] };
    }

    let lojaTotal = 0, lojaClients = 0, siteTotal = 0, siteClients = 0;
    let reachedSummary = false;
    const dailyData = [];

    for (let r = 1; r < rows.length; r++) {
        if (reachedSummary) break;
        const cells = rows[r]?.c;
        if (!cells) continue;

        const rowText = cells.map(c => String(c?.v ?? c?.f ?? '')).join('|').toUpperCase();
        if (rowText.includes('%') && !rowText.match(/\d{2}\/\d{2}/)) continue;
        if (rowText.includes('LOJA F') || rowText.includes('META MASTER') ||
            rowText.includes('BARBEARIA') || rowText.includes('PARCERIA') ||
            rowText.includes('ESTORNO')) continue;
        if (rowText.includes('COMISS') || rowText.includes('TOTAL')) { reachedSummary = true; continue; }

        let hasDate = false, dateStr = '';
        for (let i = 0; i < Math.min(cells.length, 20); i++) {
            const v = cells[i]?.f || String(cells[i]?.v || '');
            if (isDateStr(v)) { hasDate = true; dateStr = v.substring(0, 5); break; }
        }

        if (!hasDate) {
            let filledTotal = 0;
            for (let i = 0; i < cells.length; i++) {
                if (cells[i]?.v != null && String(cells[i].v).trim() !== '') filledTotal++;
            }
            if (filledTotal >= 6) { reachedSummary = true; continue; }
        }

        let rowLoja = 0, rowSite = 0;
        for (const ci of lojaCols) {
            const val = cellNum(cells[ci]);
            if (val !== null && val > 0) { lojaTotal += val; lojaClients++; rowLoja += val; }
        }
        for (const ci of siteCols) {
            const val = cellNum(cells[ci]);
            if (val !== null && val > 0) { siteTotal += val; siteClients++; rowSite += val; }
        }

        if (hasDate && (rowLoja > 0 || rowSite > 0)) {
            dailyData.push({ date: dateStr, loja: rowLoja, site: rowSite });
        }
    }

    return {
        lojaTotal, lojaClients,
        lojaTicket: lojaClients > 0 ? lojaTotal / lojaClients : 0,
        siteTotal, siteClients,
        siteTicket: siteClients > 0 ? siteTotal / siteClients : 0,
        dailyData
    };
}

// ===================== COUNT-UP ANIMATION =====================
function animateCountUp() {
    document.querySelectorAll('[data-count]').forEach(el => {
        const end = parseFloat(el.dataset.count);
        if (isNaN(end) || end === 0) return;
        const fmt = el.dataset.format === 'brl' ? fmtBRL : fmtInt;
        const duration = 800;
        const start = performance.now();
        function tick(now) {
            const p = Math.min((now - start) / duration, 1);
            const eased = p === 1 ? 1 : 1 - Math.pow(2, -10 * p);
            el.textContent = fmt(end * eased);
            if (p < 1) requestAnimationFrame(tick);
        }
        el.textContent = fmt(0);
        requestAnimationFrame(tick);
    });
}

// ===================== FORECAST =====================
function calcForecast(results) {
    const valid = results.filter(v => !v.error);
    if (valid.length === 0) return null;

    const monthIdx = parseInt(document.getElementById('month').value);
    const year = parseInt(document.getElementById('year').value);
    const now = new Date();
    const isCurrentMonth = (year === now.getFullYear() && monthIdx === now.getMonth());
    const daysInMonth = new Date(year, monthIdx + 1, 0).getDate();

    let maxDays = 0;
    valid.forEach(v => { if (v.data.dailyData.length > maxDays) maxDays = v.data.dailyData.length; });

    if (maxDays === 0) return null;

    const totalVendas = valid.reduce((s, v) => {
        const f = getFiltered(v.data);
        return s + f.total;
    }, 0);

    const mediaDiaria = totalVendas / maxDays;
    const diasRestantes = daysInMonth - maxDays;
    const projecao = totalVendas + (mediaDiaria * Math.max(0, diasRestantes));
    const percentual = Math.min((maxDays / daysInMonth) * 100, 100);

    return {
        projecao,
        totalAtual: totalVendas,
        mediaDiaria,
        diasVendidos: maxDays,
        diasNoMes: daysInMonth,
        percentual,
        mesEncerrado: !isCurrentMonth || maxDays >= daysInMonth
    };
}

// ===================== RENDER: GLOBAL KPIs (with tooltips) =====================
function renderGlobalKpis(results, prevResults) {
    const container = document.getElementById('globalKpis');
    const valid = results.filter(v => !v.error);

    if (valid.length === 0) { container.innerHTML = ''; return; }

    const t = {
        vendas: valid.reduce((s, v) => s + getFiltered(v.data).total, 0),
        clientes: valid.reduce((s, v) => s + getFiltered(v.data).clients, 0),
        vendedores: valid.length
    };
    t.ticket = t.clientes > 0 ? t.vendas / t.clientes : 0;

    let pt = null;
    if (prevResults) {
        const pv = prevResults.filter(v => !v.error);
        if (pv.length > 0) {
            pt = {
                vendas: pv.reduce((s, v) => s + getFiltered(v.data).total, 0),
                clientes: pv.reduce((s, v) => s + getFiltered(v.data).clients, 0),
            };
            pt.ticket = pt.clientes > 0 ? pt.vendas / pt.clientes : 0;
        }
    }

    // Build tooltip breakdowns
    const bd = valid.map(v => ({ name: v.name, ...getFiltered(v.data) }));
    const ttVendas = bd.map(v => `${v.name}: ${fmtBRL(v.total)}`).join('&#10;');
    const ttTicket = bd.map(v => `${v.name}: ${fmtBRL(v.ticket)}`).join('&#10;');
    const ttClientes = bd.map(v => `${v.name}: ${fmtInt(v.clients)}`).join('&#10;');

    // Forecast
    const fc = calcForecast(results);
    let forecastHTML = '';
    if (fc) {
        if (fc.mesEncerrado) {
            forecastHTML = `
                <div class="global-kpi-card">
                    <div class="global-kpi-icon i-teal">${ICON_FORECAST}</div>
                    <div style="flex:1">
                        <div class="card-label">Fechamento do Mês</div>
                        <div class="card-value teal" data-count="${fc.totalAtual}" data-format="brl">${fmtBRL(fc.totalAtual)}</div>
                        <div class="forecast-label">Mês encerrado (${fc.diasVendidos} dias)</div>
                    </div>
                </div>`;
        } else {
            forecastHTML = `
                <div class="global-kpi-card">
                    <div class="global-kpi-icon i-teal">${ICON_FORECAST}</div>
                    <div style="flex:1">
                        <div class="card-label">Projeção do Mês</div>
                        <div class="card-value teal" data-count="${fc.projecao}" data-format="brl">${fmtBRL(fc.projecao)}</div>
                        <div class="forecast-bar"><div class="forecast-bar-fill" style="width:${fc.percentual}%"></div></div>
                        <div class="forecast-label">${fc.diasVendidos}/${fc.diasNoMes} dias | Média: ${fmtBRL(fc.mediaDiaria)}/dia</div>
                    </div>
                </div>`;
        }
    }

    container.innerHTML = `
        <div class="global-kpi-card" data-tooltip="${ttVendas}">
            <div class="global-kpi-icon i-blue">${ICON_DOLLAR}</div>
            <div>
                <div class="card-label">Vendas Totais</div>
                <div class="card-value blue" data-count="${t.vendas}" data-format="brl">${fmtBRL(t.vendas)}</div>
                ${trendHTML(t.vendas, pt?.vendas)}
            </div>
        </div>
        <div class="global-kpi-card" data-tooltip="${ttTicket}">
            <div class="global-kpi-icon i-green">${ICON_TICKET}</div>
            <div>
                <div class="card-label">Ticket Médio</div>
                <div class="card-value green" data-count="${t.ticket}" data-format="brl">${fmtBRL(t.ticket)}</div>
                ${trendHTML(t.ticket, pt?.ticket)}
            </div>
        </div>
        <div class="global-kpi-card" data-tooltip="${ttClientes}">
            <div class="global-kpi-icon i-purple">${ICON_USERS}</div>
            <div>
                <div class="card-label">Total Clientes</div>
                <div class="card-value purple" data-count="${t.clientes}" data-format="int">${fmtInt(t.clientes)}</div>
                ${trendHTML(t.clientes, pt?.clientes)}
            </div>
        </div>
        <div class="global-kpi-card">
            <div class="global-kpi-icon i-orange">${ICON_USER}</div>
            <div>
                <div class="card-label">Vendedores Ativos</div>
                <div class="card-value orange">${t.vendedores}</div>
            </div>
        </div>
        ${forecastHTML}`;
    animateChildren(container, '.global-kpi-card');
}

// ===================== RENDER: PODIUM =====================
function renderPodium(results) {
    const container = document.getElementById('podium');
    if (!container) return;
    const valid = results.filter(v => !v.error);
    if (valid.length < 2) { container.innerHTML = ''; return; }

    const ranked = valid.map(v => ({
        name: v.name,
        total: getFiltered(v.data).total
    })).sort((a, b) => b.total - a.total).slice(0, 3);

    const medals = ['\u{1F947}', '\u{1F948}', '\u{1F949}'];
    const classes = ['podium-1', 'podium-2', 'podium-3'];

    // Display order: 2nd, 1st, 3rd (classic podium)
    const order = ranked.length >= 3 ? [1, 0, 2] : [1, 0];
    const placesHTML = order.map(i => {
        const v = ranked[i];
        if (!v) return '';
        return `
            <div class="podium-place ${classes[i]}">
                <span class="podium-medal">${medals[i]}</span>
                <span class="podium-name">${v.name}</span>
                <span class="podium-value">${fmtBRL(v.total)}</span>
                <div class="podium-bar">
                    <span class="podium-rank">#${i + 1}</span>
                </div>
            </div>`;
    }).join('');

    container.innerHTML = `<div class="podium">${placesHTML}</div>`;
}

// ===================== RENDER: PARTICIPATION DONUT =====================
function renderParticipation(results) {
    const container = document.getElementById('participationArea');
    if (!container) return;
    const valid = results.filter(v => !v.error);
    if (valid.length === 0) { container.innerHTML = ''; return; }

    if (!document.getElementById('chartParticipation')) {
        container.innerHTML = `
            <div class="chart-box" style="max-width:500px; margin: 2rem auto;">
                <h3>Participação por Vendedor (%)</h3>
                <canvas id="chartParticipation"></canvas>
            </div>`;
    }

    const tc = getThemeColors();
    const vendorColors = [
        'rgba(59,130,246,0.8)', 'rgba(139,92,246,0.8)', 'rgba(16,185,129,0.8)',
        'rgba(245,158,11,0.8)', 'rgba(239,68,68,0.8)', 'rgba(236,72,153,0.8)',
    ];

    const data = valid.map(v => getFiltered(v.data).total);
    const labels = valid.map(v => v.name);
    const total = data.reduce((s, v) => s + v, 0);

    const el = document.getElementById('chartParticipation');
    if (!el) return;
    if (charts['chartParticipation']) charts['chartParticipation'].destroy();

    charts['chartParticipation'] = new Chart(el.getContext('2d'), {
        type: 'doughnut',
        data: {
            labels,
            datasets: [{ data, backgroundColor: vendorColors.slice(0, valid.length), borderWidth: 0 }]
        },
        options: {
            responsive: true,
            plugins: {
                legend: { position: 'bottom', labels: { color: tc.text2, padding: 16 } },
                tooltip: {
                    callbacks: {
                        label: ctx => `${ctx.label}: ${fmtBRL(ctx.raw)} (${total > 0 ? ((ctx.raw / total) * 100).toFixed(1) : 0}%)`
                    }
                }
            }
        }
    });
}

// ===================== FLIP ANIMATION =====================
function captureVendorPositions() {
    const pos = {};
    document.querySelectorAll('.vendor-section[data-vendor]').forEach(el => {
        pos[el.dataset.vendor] = el.getBoundingClientRect();
    });
    return pos;
}

function animateVendorReorder(oldPos) {
    document.querySelectorAll('.vendor-section[data-vendor]').forEach(el => {
        const old = oldPos[el.dataset.vendor];
        if (!old) return;
        const cur = el.getBoundingClientRect();
        const dy = old.top - cur.top;
        if (Math.abs(dy) < 5) return;
        el.style.transform = `translateY(${dy}px)`;
        el.style.transition = 'none';
        requestAnimationFrame(() => requestAnimationFrame(() => {
            el.style.transition = 'transform 0.6s cubic-bezier(0.4,0,0.2,1)';
            el.style.transform = '';
        }));
    });
}

// ===================== RENDER: VENDOR CARDS =====================
function renderVendorCards(results, prevResults) {
    const container = document.getElementById('vendorSections');
    container.innerHTML = '';

    if (vendors.length === 0) {
        container.innerHTML = '<div class="empty-state"><p>Nenhum vendedor configurado.</p><button class="btn btn-primary" onclick="openVendorModal()">+ Adicionar Vendedor</button></div>';
        return;
    }

    const ranked = results.map((v, i) => ({ ...v, oi: v.originalIdx != null ? v.originalIdx : i }));
    ranked.sort((a, b) => {
        if (a.error && b.error) return 0;
        if (a.error) return 1;
        if (b.error) return -1;
        return getFiltered(b.data).total - getFiltered(a.data).total;
    });

    const medals = ['\u{1F947}', '\u{1F948}', '\u{1F949}'];
    const showLoja = channelFilter === 'all' || channelFilter === 'loja';
    const showSite = channelFilter === 'all' || channelFilter === 'site';

    ranked.forEach((vd, rank) => {
        const section = document.createElement('div');
        section.className = 'vendor-section' + (rank === 0 && !vd.error ? ' rank-1' : '');
        section.dataset.vendor = vd.name;

        const prevVd = prevResults?.find(p => p.name === vd.name);
        function vTrend(cur, prev) {
            if (!prevVd || prevVd.error) return '';
            return trendHTML(cur, prev);
        }

        if (vd.error) {
            const vi = vendors[vd.oi];
            section.innerHTML = `
                <div class="section-title">
                    ${vd.name}
                    <button class="btn-remove" onclick="removeVendor(${vd.oi})" title="Remover">&#10005;</button>
                </div>
                <div class="error-card">
                    <div class="error-card-icon">${ICON_ALERT}</div>
                    <div class="error-card-content">
                        <h4>Dados não encontrados</h4>
                        <p>${vd.error}</p>
                        <p class="error-hint">Verifique se a aba do mês existe na planilha ou se o nome está correto.</p>
                        <div class="error-actions">
                            <button class="btn btn-primary btn-sm" onclick="retryVendor(${vd.oi})">Tentar Novamente</button>
                            ${vi ? `<a href="https://docs.google.com/spreadsheets/d/${vi.id}" target="_blank" rel="noopener">Abrir Planilha &#8599;</a>` : ''}
                        </div>
                    </div>
                </div>`;
        } else {
            const d = vd.data;
            const medal = rank < medals.length ? `<span class="rank-badge">${medals[rank]}</span>` : '';
            const cDaily = 'daily_' + vd.oi;
            const cPie = 'pie_' + vd.oi;

            let cardsHTML = '';
            if (showLoja) {
                cardsHTML += `
                <div class="subsection-label">Loja Física</div>
                <div class="cards">
                    <div class="card"><div class="card-label">Vendas Total</div><div class="card-value blue" data-count="${d.lojaTotal}" data-format="brl">${fmtBRL(d.lojaTotal)}</div>${vTrend(d.lojaTotal, prevVd?.data?.lojaTotal)}</div>
                    <div class="card"><div class="card-label">Ticket Médio</div><div class="card-value green" data-count="${d.lojaTicket}" data-format="brl">${fmtBRL(d.lojaTicket)}</div>${vTrend(d.lojaTicket, prevVd?.data?.lojaTicket)}</div>
                    <div class="card"><div class="card-label">Clientes Atendidos</div><div class="card-value purple" data-count="${d.lojaClients}" data-format="int">${fmtInt(d.lojaClients)}</div>${vTrend(d.lojaClients, prevVd?.data?.lojaClients)}</div>
                </div>`;
            }
            if (showSite) {
                cardsHTML += `
                <div class="subsection-label">Site</div>
                <div class="cards">
                    <div class="card"><div class="card-label">Vendas Total</div><div class="card-value blue" data-count="${d.siteTotal}" data-format="brl">${fmtBRL(d.siteTotal)}</div>${vTrend(d.siteTotal, prevVd?.data?.siteTotal)}</div>
                    <div class="card"><div class="card-label">Ticket Médio</div><div class="card-value green" data-count="${d.siteTicket}" data-format="brl">${fmtBRL(d.siteTicket)}</div>${vTrend(d.siteTicket, prevVd?.data?.siteTicket)}</div>
                    <div class="card"><div class="card-label">Clientes Atendidos</div><div class="card-value purple" data-count="${d.siteClients}" data-format="int">${fmtInt(d.siteClients)}</div>${vTrend(d.siteClients, prevVd?.data?.siteClients)}</div>
                </div>`;
            }

            section.innerHTML = `
                <div class="section-title">
                    ${medal}${vd.name}
                    <span class="badge badge-blue">#${rank + 1}</span>
                    <button class="btn-remove" onclick="removeVendor(${vd.oi})" title="Remover">&#10005;</button>
                </div>
                ${cardsHTML}
                <div class="vendor-charts">
                    <div class="vendor-chart-box"><h4>Evolução Diária</h4><canvas id="${cDaily}"></canvas></div>
                    <div class="vendor-chart-box"><h4>Loja vs Site</h4><canvas id="${cPie}"></canvas></div>
                </div>`;
        }
        container.appendChild(section);
    });
}

// ===================== RENDER: COMPARE VIEW =====================
function renderCompareView(resultsA, resultsB, nameA, nameB) {
    const container = document.getElementById('vendorSections');
    container.innerHTML = '';

    for (let i = 0; i < vendors.length; i++) {
        const a = resultsA[i];
        const b = resultsB[i];
        const section = document.createElement('div');
        section.className = 'vendor-section';

        if (a.error && b.error) {
            section.innerHTML = `<div class="section-title">${a.name}<button class="btn-remove" onclick="removeVendor(${i})">&#10005;</button></div><div class="error-card"><div class="error-card-icon">${ICON_ALERT}</div><div class="error-card-content"><h4>Sem dados em ambos os meses</h4></div></div>`;
        } else {
            const da = a.error ? null : a.data;
            const db = b.error ? null : b.data;

            function cmpCard(label, valA, valB, fmt, cls) {
                const va = valA != null ? fmt(valA) : '\u2014';
                const vb = valB != null ? fmt(valB) : '\u2014';
                let delta = '';
                if (valA != null && valB != null && valB !== 0) {
                    const pct = ((valA - valB) / valB) * 100;
                    const c = pct > 0 ? 'trend-up' : pct < 0 ? 'trend-down' : 'trend-neutral';
                    const arr = pct > 0 ? '&#9650;' : pct < 0 ? '&#9660;' : '=';
                    delta = `<span class="trend ${c}">${arr} ${Math.abs(pct).toFixed(1)}%</span>`;
                }
                return `<div class="card"><div class="card-label">${label}</div><div class="compare-values"><div><span class="compare-month-label">${nameA}</span><div class="card-value ${cls}">${va}</div></div><div><span class="compare-month-label">${nameB}</span><div class="card-value ${cls}" style="opacity:0.6">${vb}</div></div></div>${delta}</div>`;
            }

            section.innerHTML = `
                <div class="section-title">${a.name}<button class="btn-remove" onclick="removeVendor(${i})">&#10005;</button></div>
                <div class="subsection-label">Loja Física</div>
                <div class="cards">
                    ${cmpCard('Vendas Total', da?.lojaTotal, db?.lojaTotal, fmtBRL, 'blue')}
                    ${cmpCard('Ticket Médio', da?.lojaTicket, db?.lojaTicket, fmtBRL, 'green')}
                    ${cmpCard('Clientes', da?.lojaClients, db?.lojaClients, fmtInt, 'purple')}
                </div>
                <div class="subsection-label">Site</div>
                <div class="cards">
                    ${cmpCard('Vendas Total', da?.siteTotal, db?.siteTotal, fmtBRL, 'blue')}
                    ${cmpCard('Ticket Médio', da?.siteTicket, db?.siteTicket, fmtBRL, 'green')}
                    ${cmpCard('Clientes', da?.siteClients, db?.siteClients, fmtInt, 'purple')}
                </div>`;
        }
        container.appendChild(section);
    }
}

// ===================== RENDER: INLINE CHARTS =====================
function renderVendorInlineCharts(results) {
    const tc = getThemeColors();
    const valid = results.filter(v => !v.error);

    valid.forEach(vd => {
        const idx = vd.oi != null ? vd.oi : vd.originalIdx;
        const cDaily = 'daily_' + idx;
        const cPie = 'pie_' + idx;
        const d = vd.data;

        const canvasD = document.getElementById(cDaily);
        if (canvasD && d.dailyData && d.dailyData.length > 0) {
            if (charts[cDaily]) charts[cDaily].destroy();
            const datasets = [];
            if (channelFilter !== 'site') datasets.push({ label: 'Loja', data: d.dailyData.map(x => x.loja), borderColor: 'rgba(59,130,246,0.8)', backgroundColor: 'rgba(59,130,246,0.1)', fill: true, tension: 0.3, pointRadius: 2 });
            if (channelFilter !== 'loja') datasets.push({ label: 'Site', data: d.dailyData.map(x => x.site), borderColor: 'rgba(139,92,246,0.8)', backgroundColor: 'rgba(139,92,246,0.1)', fill: true, tension: 0.3, pointRadius: 2 });
            charts[cDaily] = new Chart(canvasD.getContext('2d'), {
                type: 'line',
                data: { labels: d.dailyData.map(x => x.date), datasets },
                options: {
                    responsive: true,
                    plugins: { legend: { labels: { color: tc.text2, font: { size: 11 } } } },
                    scales: {
                        x: { ticks: { color: tc.text2, font: { size: 10 }, maxRotation: 45 }, grid: { color: tc.grid } },
                        y: { ticks: { color: tc.text2, font: { size: 10 }, callback: v => 'R$ ' + v.toLocaleString('pt-BR') }, grid: { color: tc.grid } }
                    }
                }
            });
        }

        const canvasP = document.getElementById(cPie);
        if (canvasP && (d.lojaTotal > 0 || d.siteTotal > 0)) {
            if (charts[cPie]) charts[cPie].destroy();
            const total = d.lojaTotal + d.siteTotal;
            charts[cPie] = new Chart(canvasP.getContext('2d'), {
                type: 'doughnut',
                data: {
                    labels: ['Loja Física', 'Site'],
                    datasets: [{ data: [d.lojaTotal, d.siteTotal], backgroundColor: ['rgba(59,130,246,0.8)', 'rgba(139,92,246,0.8)'], borderWidth: 0 }]
                },
                options: {
                    responsive: true,
                    plugins: {
                        legend: { labels: { color: tc.text2, font: { size: 11 } } },
                        tooltip: { callbacks: { label: ctx => `${ctx.label}: ${fmtBRL(ctx.raw)} (${((ctx.raw/total)*100).toFixed(1)}%)` } }
                    }
                }
            });
        }
    });
}

// ===================== RENDER: MAIN CHARTS =====================
function renderCharts(vendorData) {
    const valid = vendorData.filter(v => !v.error);
    if (valid.length === 0) { document.getElementById('chartsArea').style.display = 'none'; return; }
    document.getElementById('chartsArea').style.display = '';

    const tc = getThemeColors();
    const names = valid.map(v => v.name);
    const colors = { loja: 'rgba(59,130,246,0.8)', site: 'rgba(139,92,246,0.8)' };

    function makeChart(canvasId, label1, data1, label2, data2, prefix) {
        if (charts[canvasId]) charts[canvasId].destroy();
        const el = document.getElementById(canvasId);
        if (!el) return;
        const datasets = [];
        if (channelFilter !== 'site') datasets.push({ label: label1, data: data1, backgroundColor: colors.loja, borderRadius: 6 });
        if (channelFilter !== 'loja') datasets.push({ label: label2, data: data2, backgroundColor: colors.site, borderRadius: 6 });
        charts[canvasId] = new Chart(el.getContext('2d'), {
            type: 'bar',
            data: { labels: names, datasets },
            options: {
                responsive: true,
                plugins: { legend: { labels: { color: tc.text2 } } },
                scales: {
                    x: { ticks: { color: tc.text2 }, grid: { color: tc.grid } },
                    y: { ticks: { color: tc.text2, callback: v => prefix ? ('R$ ' + v.toLocaleString('pt-BR')) : v }, grid: { color: tc.grid } }
                }
            }
        });
    }

    makeChart('chartTotals', 'Loja', valid.map(v => v.data.lojaTotal), 'Site', valid.map(v => v.data.siteTotal), true);
    makeChart('chartTicket', 'Loja', valid.map(v => v.data.lojaTicket), 'Site', valid.map(v => v.data.siteTicket), true);
    makeChart('chartClients', 'Loja', valid.map(v => v.data.lojaClients), 'Site', valid.map(v => v.data.siteClients), false);

    const totalLoja = valid.reduce((s, v) => s + v.data.lojaTotal, 0);
    const totalSite = valid.reduce((s, v) => s + v.data.siteTotal, 0);
    const pieEl = document.getElementById('chartPieGlobal');
    if (pieEl && (totalLoja > 0 || totalSite > 0)) {
        if (charts['chartPieGlobal']) charts['chartPieGlobal'].destroy();
        const total = totalLoja + totalSite;
        charts['chartPieGlobal'] = new Chart(pieEl.getContext('2d'), {
            type: 'doughnut',
            data: {
                labels: ['Loja Física', 'Site'],
                datasets: [{ data: [totalLoja, totalSite], backgroundColor: [colors.loja, colors.site], borderWidth: 0 }]
            },
            options: {
                responsive: true,
                plugins: {
                    legend: { labels: { color: tc.text2 } },
                    tooltip: { callbacks: { label: ctx => `${ctx.label}: ${fmtBRL(ctx.raw)} (${((ctx.raw/total)*100).toFixed(1)}%)` } }
                }
            }
        });
    }
}

// ===================== WEEKDAY AVERAGE CHART =====================
function renderWeekdayChart(results) {
    const el = document.getElementById('chartWeekday');
    if (!el) return;
    const valid = results.filter(v => !v.error);
    if (valid.length === 0) return;

    const year = parseInt(document.getElementById('year').value);
    const monthIdx = parseInt(document.getElementById('month').value);
    const tc = getThemeColors();

    const weekdaySums = Array(7).fill(0);
    const weekdayCounts = Array(7).fill(0);

    valid.forEach(v => {
        v.data.dailyData.forEach(dd => {
            const parts = dd.date.split('/');
            const day = parseInt(parts[0]);
            const month = parseInt(parts[1]) - 1;
            const d = new Date(year, month >= 0 ? month : monthIdx, day);
            if (isNaN(d.getTime())) return;
            const dow = d.getDay();
            const val = channelFilter === 'loja' ? dd.loja : channelFilter === 'site' ? dd.site : dd.loja + dd.site;
            weekdaySums[dow] += val;
            weekdayCounts[dow] += 1;
        });
    });

    const avgData = weekdaySums.map((sum, i) => weekdayCounts[i] > 0 ? sum / weekdayCounts[i] : 0);
    const ordered = [1, 2, 3, 4, 5, 6, 0];
    const labels = ordered.map(i => WEEKDAYS[i]);
    const data = ordered.map(i => avgData[i]);
    const bgColors = ordered.map(i => {
        const max = Math.max(...data);
        const ratio = max > 0 ? data[ordered.indexOf(i)] / max : 0;
        return `rgba(59,130,246,${0.3 + ratio * 0.5})`;
    });

    if (charts['chartWeekday']) charts['chartWeekday'].destroy();
    charts['chartWeekday'] = new Chart(el.getContext('2d'), {
        type: 'bar',
        data: {
            labels,
            datasets: [{
                label: 'Média de Vendas',
                data,
                backgroundColor: bgColors,
                borderRadius: 6,
            }]
        },
        options: {
            responsive: true,
            plugins: {
                legend: { display: false },
                tooltip: { callbacks: { label: ctx => fmtBRL(ctx.raw) } }
            },
            scales: {
                x: { ticks: { color: tc.text2 }, grid: { color: tc.grid } },
                y: { ticks: { color: tc.text2, callback: v => 'R$ ' + v.toLocaleString('pt-BR') }, grid: { color: tc.grid } }
            }
        }
    });
}

// ===================== HEATMAP =====================
function renderHeatmap(results) {
    const container = document.getElementById('heatmapArea');
    if (!container) return;
    const valid = results.filter(v => !v.error);
    if (valid.length === 0) { container.innerHTML = ''; return; }

    const monthIdx = parseInt(document.getElementById('month').value);
    const year = parseInt(document.getElementById('year').value);
    const daysInMonth = new Date(year, monthIdx + 1, 0).getDate();
    const firstDayOfWeek = new Date(year, monthIdx, 1).getDay(); // 0=Sun

    const dayTotals = {};
    valid.forEach(v => {
        v.data.dailyData.forEach(dd => {
            const day = parseInt(dd.date.split('/')[0]);
            const val = channelFilter === 'loja' ? dd.loja : channelFilter === 'site' ? dd.site : dd.loja + dd.site;
            dayTotals[day] = (dayTotals[day] || 0) + val;
        });
    });

    const values = Object.values(dayTotals);
    const maxVal = Math.max(...values, 1);

    function getHeatColor(value) {
        if (value === 0 || value == null) return 'var(--border)';
        const ratio = value / maxVal;
        if (ratio < 0.2) return 'rgba(59,130,246,0.15)';
        if (ratio < 0.4) return 'rgba(59,130,246,0.3)';
        if (ratio < 0.6) return 'rgba(59,130,246,0.5)';
        if (ratio < 0.8) return 'rgba(59,130,246,0.7)';
        return 'rgba(59,130,246,0.9)';
    }

    const weekdayHeaders = ['Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb', 'Dom'];
    const startOffset = (firstDayOfWeek + 6) % 7; // Mon-based

    let cellsHTML = weekdayHeaders.map(d => `<div class="heatmap-header">${d}</div>`).join('');

    for (let i = 0; i < startOffset; i++) {
        cellsHTML += '<div class="heatmap-cell empty"></div>';
    }

    for (let day = 1; day <= daysInMonth; day++) {
        const val = dayTotals[day] || 0;
        const color = getHeatColor(val);
        const tooltip = val > 0 ? fmtBRL(val) : 'Sem vendas';
        cellsHTML += `<div class="heatmap-cell" style="background:${color}" title="Dia ${day}: ${tooltip}">${day}</div>`;
    }

    const legendColors = ['rgba(59,130,246,0.15)', 'rgba(59,130,246,0.3)', 'rgba(59,130,246,0.5)', 'rgba(59,130,246,0.7)', 'rgba(59,130,246,0.9)'];

    container.innerHTML = `
        <div class="heatmap-box">
            <h3>Heatmap de Vendas - ${MONTHS_ALT[monthIdx]} ${year}</h3>
            <div class="heatmap-grid">${cellsHTML}</div>
            <div class="heatmap-legend">
                <span>Menos</span>
                <div class="heatmap-legend-bar">
                    ${legendColors.map(c => `<span style="background:${c}"></span>`).join('')}
                </div>
                <span>Mais</span>
            </div>
        </div>`;
}

// ===================== VENDOR COMPARATOR =====================
function renderVendorComparator(results) {
    const container = document.getElementById('vendorCompare');
    if (!container) return;
    const valid = results.filter(v => !v.error);
    if (valid.length < 2) { container.innerHTML = ''; return; }

    const options = valid.map((v, i) => `<option value="${i}">${v.name}</option>`).join('');

    container.innerHTML = `
        <div class="comparator-box">
            <h3>Comparativo Vendedor vs Vendedor</h3>
            <div class="comparator-selectors">
                <select id="cmpVendorA" onchange="updateComparator()">${options}</select>
                <span class="comparator-vs">VS</span>
                <select id="cmpVendorB" onchange="updateComparator()">
                    ${valid.map((v, i) => `<option value="${i}" ${i === 1 ? 'selected' : ''}>${v.name}</option>`).join('')}
                </select>
            </div>
            <div id="comparatorResults"></div>
            <div class="comparator-chart-area">
                <canvas id="chartComparator"></canvas>
            </div>
        </div>`;

    updateComparator();
}

function updateComparator() {
    if (!lastResults) return;
    const valid = lastResults.filter(v => !v.error);
    const selA = document.getElementById('cmpVendorA');
    const selB = document.getElementById('cmpVendorB');
    if (!selA || !selB) return;
    const idxA = parseInt(selA.value);
    const idxB = parseInt(selB.value);
    if (isNaN(idxA) || isNaN(idxB) || idxA === idxB) return;

    const a = valid[idxA], b = valid[idxB];
    if (!a || !b) return;

    const fa = getFiltered(a.data), fb = getFiltered(b.data);
    const metrics = [
        { label: 'Vendas Totais', a: fa.total, b: fb.total, fmt: fmtBRL },
        { label: 'Ticket Médio', a: fa.ticket, b: fb.ticket, fmt: fmtBRL },
        { label: 'Clientes', a: fa.clients, b: fb.clients, fmt: fmtInt },
    ];

    const resultsDiv = document.getElementById('comparatorResults');
    resultsDiv.innerHTML = metrics.map(m => {
        const winA = m.a > m.b, winB = m.b > m.a;
        return `
            <div class="comparator-row">
                <div class="comparator-metric ${winA ? 'winner' : ''}">
                    <span style="font-weight:700">${m.fmt(m.a)}</span>
                    ${winA ? ' <span style="color:var(--green)">&#10003;</span>' : ''}
                </div>
                <div class="comparator-label">${m.label}</div>
                <div class="comparator-metric ${winB ? 'winner' : ''}">
                    ${winB ? '<span style="color:var(--green)">&#10003;</span> ' : ''}
                    <span style="font-weight:700">${m.fmt(m.b)}</span>
                </div>
            </div>`;
    }).join('');

    const tc = getThemeColors();
    const el = document.getElementById('chartComparator');
    if (!el) return;
    if (charts['chartComparator']) charts['chartComparator'].destroy();

    charts['chartComparator'] = new Chart(el.getContext('2d'), {
        type: 'bar',
        data: {
            labels: ['Vendas', 'Ticket Médio', 'Clientes'],
            datasets: [
                { label: a.name, data: [fa.total, fa.ticket, fa.clients], backgroundColor: 'rgba(59,130,246,0.8)', borderRadius: 6 },
                { label: b.name, data: [fb.total, fb.ticket, fb.clients], backgroundColor: 'rgba(139,92,246,0.8)', borderRadius: 6 }
            ]
        },
        options: {
            responsive: true,
            plugins: { legend: { labels: { color: tc.text2 } } },
            scales: {
                x: { ticks: { color: tc.text2 }, grid: { color: tc.grid } },
                y: { ticks: { color: tc.text2 }, grid: { color: tc.grid } }
            }
        }
    });
}

// ===================== CHART FULLSCREEN =====================
function addExpandButtons() {
    document.querySelectorAll('.chart-box').forEach(box => {
        if (box.querySelector('.chart-expand-btn')) return;
        const btn = document.createElement('button');
        btn.className = 'chart-expand-btn';
        btn.innerHTML = ICON_EXPAND;
        btn.onclick = () => openChartFullscreen(box);
        box.appendChild(btn);
    });
}

function openChartFullscreen(chartBox) {
    const title = chartBox.querySelector('h3')?.textContent || 'Gráfico';
    const originalCanvas = chartBox.querySelector('canvas');
    if (!originalCanvas) return;

    const chartId = originalCanvas.id;
    const chartInstance = charts[chartId];
    if (!chartInstance) return;

    document.getElementById('chartFullscreenTitle').textContent = title;
    const overlay = document.getElementById('chartFullscreen');
    overlay.style.display = 'flex';

    const fsCanvas = document.getElementById('chartFullscreenCanvas');
    if (fullscreenChart) fullscreenChart.destroy();

    // Clone chart data and create new instance at full size
    const clonedData = JSON.parse(JSON.stringify(chartInstance.data));
    const origOpts = chartInstance.options;
    const tc = getThemeColors();

    fullscreenChart = new Chart(fsCanvas.getContext('2d'), {
        type: chartInstance.config.type,
        data: clonedData,
        options: {
            ...origOpts,
            maintainAspectRatio: false,
            plugins: {
                ...origOpts.plugins,
                legend: { labels: { color: tc.text2, font: { size: 14 } } },
                tooltip: {
                    callbacks: {
                        label: ctx => {
                            const v = ctx.raw;
                            if (typeof v === 'number' && v > 100) return `${ctx.dataset.label || ''}: ${fmtBRL(v)}`;
                            return `${ctx.dataset.label || ''}: ${v}`;
                        }
                    }
                }
            },
            scales: origOpts.scales ? {
                x: { ...origOpts.scales.x, ticks: { color: tc.text2, font: { size: 13 } }, grid: { color: tc.grid } },
                y: { ...origOpts.scales.y, ticks: { color: tc.text2, font: { size: 13 } }, grid: { color: tc.grid } }
            } : undefined
        }
    });

    document.addEventListener('keydown', handleEscFullscreen);
}

function closeChartFullscreen() {
    document.getElementById('chartFullscreen').style.display = 'none';
    if (fullscreenChart) { fullscreenChart.destroy(); fullscreenChart = null; }
    document.removeEventListener('keydown', handleEscFullscreen);
}

function handleEscFullscreen(e) {
    if (e.key === 'Escape') closeChartFullscreen();
}

// ===================== COMPARE MODE =====================
function toggleCompare() {
    compareMode = !compareMode;
    document.getElementById('compareControls').classList.toggle('active', compareMode);
    document.getElementById('compareToggle').classList.toggle('btn-primary', compareMode);
    if (compareMode) refreshData({ toast: true });
    else if (lastResults) reRenderAll();
}

// ===================== RETRY VENDOR =====================
async function retryVendor(index) {
    const vendor = vendors[index];
    const sheetName = getSheetName();
    try {
        const table = await fetchForVendorWithSheet(vendor.id, sheetName);
        const data = processTable(table);
        if (lastResults) {
            lastResults[index] = { name: vendor.name, data, originalIdx: index };
            reRenderAll();
        }
    } catch {
        alert('Ainda não foi possível carregar dados de ' + vendor.name + '.');
    }
}

// ===================== YEARLY EVOLUTION CHART =====================
async function fetchAndRenderYearly() {
    const year = parseInt(document.getElementById('year').value);
    const now = new Date();
    const maxMonth = (year === now.getFullYear()) ? now.getMonth() : 11;
    const tc = getThemeColors();

    document.getElementById('yearlyLabel').textContent = '(' + year + ')';

    const lineColors = [
        { border: '#3b82f6', bg: 'rgba(59,130,246,0.1)' },
        { border: '#8b5cf6', bg: 'rgba(139,92,246,0.1)' },
        { border: '#10b981', bg: 'rgba(16,185,129,0.1)' },
        { border: '#f59e0b', bg: 'rgba(245,158,11,0.1)' },
        { border: '#ef4444', bg: 'rgba(239,68,68,0.1)' },
        { border: '#ec4899', bg: 'rgba(236,72,153,0.1)' },
    ];

    const monthLabels = [];
    const vendorMonthlyTotals = vendors.map(() => []);

    const monthPromises = [];
    for (let m = 0; m <= maxMonth; m++) {
        const sheetName = getSheetNameFor(m, year);
        monthLabels.push(MONTHS_ALT[m].substring(0, 3));
        monthPromises.push(
            Promise.all(vendors.map(async (vendor, vi) => {
                try {
                    const table = await fetchForVendorWithSheet(vendor.id, sheetName);
                    const data = processTable(table);
                    return { vi, total: data.lojaTotal + data.siteTotal };
                } catch {
                    return { vi, total: 0 };
                }
            }))
        );
    }

    const monthResults = await Promise.all(monthPromises);
    monthResults.forEach(monthVendors => {
        monthVendors.forEach(r => { vendorMonthlyTotals[r.vi].push(r.total); });
    });

    const datasets = vendors.map((vendor, vi) => {
        const color = lineColors[vi % lineColors.length];
        return { label: vendor.name, data: vendorMonthlyTotals[vi], borderColor: color.border, backgroundColor: color.bg, fill: true, tension: 0.3, pointRadius: 4, pointHoverRadius: 6 };
    });

    const totalData = [];
    for (let m = 0; m <= maxMonth; m++) {
        let sum = 0;
        vendors.forEach((_, vi) => { sum += vendorMonthlyTotals[vi][m] || 0; });
        totalData.push(sum);
    }
    datasets.push({ label: 'Total', data: totalData, borderColor: '#fbbf24', backgroundColor: 'rgba(251,191,36,0.08)', fill: false, tension: 0.3, pointRadius: 4, pointHoverRadius: 6, borderWidth: 3, borderDash: [6, 3] });

    const elY = document.getElementById('chartYearly');
    if (!elY) return;
    if (charts['chartYearly']) charts['chartYearly'].destroy();

    charts['chartYearly'] = new Chart(elY.getContext('2d'), {
        type: 'line',
        data: { labels: monthLabels, datasets },
        options: {
            responsive: true,
            interaction: { mode: 'index', intersect: false },
            plugins: {
                legend: { labels: { color: tc.text2 } },
                tooltip: { callbacks: { label: ctx => `${ctx.dataset.label}: ${fmtBRL(ctx.raw)}` } }
            },
            scales: {
                x: { ticks: { color: tc.text2 }, grid: { color: tc.grid } },
                y: { ticks: { color: tc.text2, callback: v => 'R$ ' + v.toLocaleString('pt-BR') }, grid: { color: tc.grid } }
            }
        }
    });

    yearlyFetched = true;
}

// ===================== INICIALIZAÇÃO =====================
function initSelectors() {
    const now = new Date();
    const currentMonth = now.getMonth();
    const currentYear = now.getFullYear();

    function populateSelector(selMonth, selYear, linkRefresh) {
        selYear.innerHTML = '';
        for (let y = 2025; y <= currentYear; y++) {
            const opt = document.createElement('option');
            opt.value = y; opt.textContent = y;
            if (y === currentYear) opt.selected = true;
            selYear.appendChild(opt);
        }

        function fillMonths() {
            const sy = parseInt(selYear.value);
            const maxM = (sy === currentYear) ? currentMonth : 11;
            selMonth.innerHTML = '';
            for (let i = 0; i <= maxM; i++) {
                const opt = document.createElement('option');
                opt.value = i; opt.textContent = MONTHS_ALT[i];
                if (sy === currentYear && i === currentMonth) opt.selected = true;
                else if (sy < currentYear && i === maxM) opt.selected = true;
                selMonth.appendChild(opt);
            }
        }

        fillMonths();
        selYear.addEventListener('change', () => { fillMonths(); if (linkRefresh) refreshData({ toast: true }); });
        if (linkRefresh) selMonth.addEventListener('change', () => refreshData({ toast: true }));
    }

    populateSelector(document.getElementById('month'), document.getElementById('year'), true);
    populateSelector(document.getElementById('month2'), document.getElementById('year2'), false);

    const prev = currentMonth === 0 ? { m: 11, y: currentYear - 1 } : { m: currentMonth - 1, y: currentYear };
    document.getElementById('year2').value = prev.y;
    document.getElementById('year2').dispatchEvent(new Event('change'));
    document.getElementById('month2').value = prev.m;
}

async function refreshData(opts) {
    const loading = document.getElementById('loading');
    const content = document.getElementById('content');
    loading.style.display = '';
    content.style.display = 'none';
    yearlyFetched = false;

    const sheetName = getSheetName();
    const prev = getPreviousMonthInfo();
    const prevSheetName = getSheetNameFor(prev.monthIdx, prev.year);

    const [results, prevResults] = await Promise.all([
        fetchMonthData(sheetName),
        fetchMonthData(prevSheetName)
    ]);

    lastResults = results;
    lastPrevResults = prevResults;

    // Always render Visão Geral content
    renderGlobalKpis(results, prevResults);
    renderPodium(results);
    renderParticipation(results);

    if (compareMode) {
        const m2 = parseInt(document.getElementById('month2').value);
        const y2 = parseInt(document.getElementById('year2').value);
        const sheetName2 = getSheetNameFor(m2, y2);
        const compareResults = await fetchMonthData(sheetName2);
        renderCompareView(results, compareResults, sheetName, sheetName2);
    } else {
        renderVendorCards(results, prevResults);
        animateChildren(document.getElementById('vendorSections'), '.vendor-section');
    }

    // Render charts only if their tab is active
    if (activeTab === 'vendedores') {
        setTimeout(() => renderVendorInlineCharts(results), 0);
        renderVendorComparator(results);
    }
    if (activeTab === 'graficos') {
        renderCharts(results);
        renderWeekdayChart(results);
        renderHeatmap(results);
        fetchAndRenderYearly();
        setTimeout(() => addExpandButtons(), 50);
    }

    animateCountUp();

    loading.style.display = 'none';
    content.style.display = '';
    document.getElementById('lastUpdate').textContent = 'Atualizado: ' + new Date().toLocaleTimeString('pt-BR') + ' | Aba: ' + sheetName;
    if (opts && opts.toast) showToast('Dados atualizados!', 'success');

    updateURLParams();

    if (refreshTimer) clearInterval(refreshTimer);
    refreshTimer = setInterval(refreshData, REFRESH_MS);
}

// ===================== KEYBOARD SHORTCUTS =====================
function initKeyboardShortcuts() {
    document.addEventListener('keydown', e => {
        if (['INPUT', 'SELECT', 'TEXTAREA'].includes(e.target.tagName)) return;
        switch (e.key) {
            case '1': switchTab('visao-geral'); break;
            case '2': switchTab('vendedores'); break;
            case '3': switchTab('graficos'); break;
            case 'r': case 'R':
                if (!e.ctrlKey && !e.metaKey) { e.preventDefault(); refreshData({ toast: true }); }
                break;
            case 't': case 'T':
                if (!e.ctrlKey && !e.metaKey) toggleTheme();
                break;
            case 'Escape':
                closeVendorModal(); closeChartFullscreen();
                break;
        }
    });
}

// ===================== PULL TO REFRESH =====================
function initPullToRefresh() {
    const indicator = document.getElementById('ptrIndicator');
    if (!indicator || !('ontouchstart' in window)) return;
    let startY = 0, pulling = false;

    document.addEventListener('touchstart', e => {
        if (window.scrollY <= 0) { startY = e.touches[0].clientY; pulling = true; }
    }, { passive: true });

    document.addEventListener('touchmove', e => {
        if (!pulling) return;
        const diff = e.touches[0].clientY - startY;
        indicator.classList.toggle('visible', diff > 70 && window.scrollY <= 0);
    }, { passive: true });

    document.addEventListener('touchend', () => {
        if (!pulling) return;
        pulling = false;
        if (indicator.classList.contains('visible')) {
            indicator.querySelector('.ptr-text').textContent = 'Atualizando...';
            refreshData({ toast: true }).then(() => {
                indicator.classList.remove('visible');
                indicator.querySelector('.ptr-text').textContent = 'Puxe para atualizar';
            });
        }
    });
}

// ===================== BOOT =====================
(function init() {
    const savedTheme = localStorage.getItem('dashboard-theme') || 'dark';
    applyTheme(savedTheme);
    initSelectors();
    readURLParams();
    refreshData();
    initKeyboardShortcuts();
    initPullToRefresh();

    // Register Service Worker
    if ('serviceWorker' in navigator) {
        navigator.serviceWorker.register('./sw.js').catch(() => {});
    }
})();
