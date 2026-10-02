'use strict';

const CONFIG = {
  spreadsheetId: '1MkKh_8l21494mUAo3xKdzVVdDzmX29Ha3K8GGXCax6U',
  sheetName: '03_LiveData',
  range: 'A3:M',
  refreshMs: 30_000,
  highTempC: 80,
  // Google Visualization endpoint. The Sheet must be readable without an interactive sign-in.
  // If your Sheet is private, replace dataUrl with an Apps Script / API proxy endpoint returning JSON.
  get dataUrl() {
    const base = `https://docs.google.com/spreadsheets/d/${this.spreadsheetId}/gviz/tq`;
    const params = new URLSearchParams({ tqx: 'out:json', sheet: this.sheetName, range: this.range, _: Date.now().toString() });
    return `${base}?${params.toString()}`;
  }
};

const SNAPSHOT_FALLBACK = [
  ['A01','L1','CNC 加工中心 01','2026-10-02 09:00',66,2.8,96.8,99.4,492,'—','Normal','2026-09-22','持續監控'],
  ['A02','L1','CNC 加工中心 02','2026-10-02 09:00',73,3.6,95.1,98.9,486,'—','Normal','2026-09-20','持續監控'],
  ['A03','L1','CNC 加工中心 03','2026-10-02 09:00',89,7.2,82.5,94.1,421,'E204','Critical','2026-08-28','優先檢查冷卻系統與主軸軸承；人工確認是否停機'],
  ['A04','L1','CNC 加工中心 04','2026-10-02 09:00',81,4.1,91.2,97.6,470,'W102','Warning','2026-09-08','確認散熱與切削液狀態'],
  ['A05','L2','射出成型機 05','2026-10-02 09:00',69,3.0,97.4,99.6,498,'—','Normal','2026-09-25','持續監控'],
  ['A06','L2','射出成型機 06','2026-10-02 09:00',77,5.4,88.9,96.9,455,'E118','Warning','2026-09-01','檢查液壓系統與模具固定'],
  ['A07','L2','射出成型機 07','2026-10-02 09:00',71,2.5,96.1,99.1,487,'—','Normal','2026-09-19','持續監控'],
  ['A08','L2','射出成型機 08','2026-10-02 09:00',84,6.1,86.3,95.8,438,'E121','Critical','2026-08-26','優先檢查油溫、壓力與模具阻力'],
  ['B01','L3','組裝機器人 01','2026-10-02 09:00',58,2.1,98.2,99.7,512,'—','Normal','2026-09-27','持續監控'],
  ['B02','L3','組裝機器人 02','2026-10-02 09:00',63,3.8,94.6,98.7,501,'W205','Warning','2026-09-06','檢查減速機與末端治具'],
  ['B03','L3','組裝機器人 03','2026-10-02 09:00',61,2.6,97.8,99.3,508,'—','Normal','2026-09-18','持續監控'],
  ['B04','L3','組裝機器人 04','2026-10-02 09:00',67,4.9,90.1,97.2,476,'W210','Warning','2026-08-30','檢查關節 4 振動與線材拖鏈'],
  ['C01','L4','AOI 檢測機 01','2026-10-02 09:00',49,1.2,99,99,620,'—','Normal','2026-09-21','持續監控'],
  ['C02','L4','AOI 檢測機 02','2026-10-02 09:00',52,1.5,98.6,97.8,608,'Q301','Warning','2026-09-12','清潔鏡頭並重新確認光源穩定度'],
  ['D03','L4','冷卻水泵 01','2026-10-02 09:00',82,6.8,84,97,0,'E502','Critical','2026-08-22','優先檢查泵浦軸承、流量與冷卻水溫']
];

const headers = ['設備編號','產線','設備名稱','更新時間','溫度°C','振動 mm/s','稼動率%','良率%','當日產量','異常碼','狀態','最後保養日','建議處置'];
let machines = [];
let isFetching = false;
let fallbackWasUsed = false;

const $ = (sel) => document.querySelector(sel);
const fmt = new Intl.NumberFormat('zh-TW', { maximumFractionDigits: 1 });

function rowsToObjects(rows) {
  return rows
    .filter(row => row && row[0])
    .map(row => Object.fromEntries(headers.map((h, i) => [h, row[i] ?? ''])));
}

function parseGviz(text) {
  const match = text.match(/google\.visualization\.Query\.setResponse\((.*)\);?\s*$/s);
  if (!match) throw new Error('無法解析 Google Sheet 回應格式');
  const payload = JSON.parse(match[1]);
  if (payload.status === 'error') throw new Error(payload.errors?.[0]?.detailed_message || 'Google Sheet 回傳錯誤');
  const table = payload.table;
  const labels = table.cols.map(c => c.label || '');
  const rows = table.rows.map(r => r.c.map(c => c ? (c.f ?? c.v ?? '') : ''));

  // The range begins at header row A3:M, so gviz normally returns the header as column labels.
  // When labels are present, return only data rows. If not, tolerate a literal header row.
  if (labels.some(Boolean)) return rows;
  if (rows.length && String(rows[0][0]).includes('設備')) return rows.slice(1);
  return rows;
}

function num(v) {
  const n = Number(String(v).replace(/[% ,]/g, ''));
  return Number.isFinite(n) ? n : 0;
}

function avg(list, key) { return list.length ? list.reduce((s, x) => s + num(x[key]), 0) / list.length : 0; }
function esc(v) { return String(v ?? '').replace(/[&<>'"]/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[ch])); }

async function loadData({ manual = false } = {}) {
  if (isFetching) return;
  isFetching = true;
  setSyncState('loading', '同步中');
  $('#refreshBtn .refresh-icon').classList.add('spin');

  try {
    const response = await fetch(CONFIG.dataUrl, { cache: 'no-store', mode: 'cors' });
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const text = await response.text();
    const rows = parseGviz(text);
    if (!rows.length) throw new Error('Google Sheet 沒有可用資料');
    machines = rowsToObjects(rows);
    fallbackWasUsed = false;
    setSyncState('ok', 'Live 已連線');
    if (manual) showToast('Live Data 已更新');
  } catch (err) {
    if (!machines.length) machines = rowsToObjects(SNAPSHOT_FALLBACK);
    fallbackWasUsed = true;
    setSyncState('error', '離線快照');
    console.warn('Live Data fetch failed:', err);
    if (manual) showToast('無法連線 Google Sheet，已顯示最近快照');
  } finally {
    renderAll();
    isFetching = false;
    $('#refreshBtn .refresh-icon').classList.remove('spin');
  }
}

function setSyncState(state, text) {
  const pill = $('#syncPill');
  pill.dataset.state = state;
  $('#syncText').textContent = text;
}

function renderAll() {
  renderKpis();
  renderStatusDonut();
  renderLinePerformance();
  renderAlerts();
  renderRankBars();
  renderOutputs();
  hydrateFilters();
  renderTable();
  const latest = machines.map(m => m['更新時間']).filter(Boolean).sort().at(-1) || '--';
  $('#lastUpdated').textContent = `資料時間：${latest}${fallbackWasUsed ? '（快照）' : ''}`;
}

function renderKpis() {
  const normal = machines.filter(m => m['狀態'] === 'Normal').length;
  const warning = machines.filter(m => m['狀態'] === 'Warning').length;
  const critical = machines.filter(m => m['狀態'] === 'Critical').length;
  const lines = new Set(machines.map(m => m['產線']).filter(Boolean)).size;
  $('#kpiTotal').textContent = machines.length;
  $('#kpiLines').textContent = lines;
  $('#kpiNormal').textContent = normal;
  $('#kpiNormalRate').textContent = machines.length ? fmt.format(normal / machines.length * 100) : '0';
  $('#kpiWarning').textContent = warning;
  $('#kpiCritical').textContent = critical;
  $('#kpiUtil').textContent = fmt.format(avg(machines, '稼動率%'));
  $('#kpiYield').textContent = fmt.format(avg(machines, '良率%'));
}

function renderStatusDonut() {
  const counts = {
    Normal: machines.filter(m => m['狀態'] === 'Normal').length,
    Warning: machines.filter(m => m['狀態'] === 'Warning').length,
    Critical: machines.filter(m => m['狀態'] === 'Critical').length
  };
  const total = Math.max(machines.length, 1);
  const normalDeg = counts.Normal / total * 360;
  const warningDeg = normalDeg + counts.Warning / total * 360;
  const donut = $('#statusDonut');
  donut.style.setProperty('--normal', `${normalDeg}deg`);
  donut.style.setProperty('--warning', `${warningDeg}deg`);
  $('#donutTotal').textContent = machines.length;
  $('#legendNormal').textContent = counts.Normal;
  $('#legendWarning').textContent = counts.Warning;
  $('#legendCritical').textContent = counts.Critical;
}

function groupByLine() {
  const map = new Map();
  machines.forEach(m => {
    const line = m['產線'] || '未分類';
    if (!map.has(line)) map.set(line, []);
    map.get(line).push(m);
  });
  return [...map.entries()].sort((a,b) => a[0].localeCompare(b[0], 'zh-Hant'));
}

function renderLinePerformance() {
  $('#linePerformance').innerHTML = groupByLine().map(([line, list]) => {
    const util = avg(list, '稼動率%');
    const yieldRate = avg(list, '良率%');
    return `<div class="line-row">
      <span class="line-code">${esc(line)}</span>
      <div class="dual-track">
        <div class="track" title="稼動率 ${fmt.format(util)}%"><span class="util-fill" style="width:${Math.min(100,util)}%"></span></div>
        <div class="track" title="良率 ${fmt.format(yieldRate)}%"><span class="yield-fill" style="width:${Math.min(100,yieldRate)}%"></span></div>
      </div>
      <div class="line-values"><div>稼 ${fmt.format(util)}%</div><div>良 ${fmt.format(yieldRate)}%</div></div>
    </div>`;
  }).join('') || '<div class="empty-state">暫無資料</div>';
}

function renderAlerts() {
  const alerts = machines
    .filter(m => m['狀態'] !== 'Normal')
    .sort((a,b) => severity(b['狀態']) - severity(a['狀態']) || num(b['溫度°C']) - num(a['溫度°C']));
  $('#alertCount').textContent = `${alerts.length} 筆`;
  $('#priorityAlerts').innerHTML = alerts.slice(0,8).map(m => `
    <div class="alert-item ${String(m['狀態']).toLowerCase()}">
      <span class="alert-state"></span>
      <div class="alert-main"><strong>${esc(m['設備編號'])} · ${esc(m['設備名稱'])}</strong><span>${esc(m['異常碼'])}｜${esc(m['建議處置'])}</span></div>
      <div class="alert-metric">${fmt.format(num(m['溫度°C']))}°C</div>
    </div>`).join('') || '<div class="empty-state">目前無警告或嚴重異常</div>';
}

function severity(status) { return status === 'Critical' ? 2 : status === 'Warning' ? 1 : 0; }

function renderRankBars() {
  const temp = [...machines].sort((a,b) => num(b['溫度°C']) - num(a['溫度°C'])).slice(0,6);
  const vib = [...machines].sort((a,b) => num(b['振動 mm/s']) - num(a['振動 mm/s'])).slice(0,6);
  const tempMax = Math.max(100, ...temp.map(m => num(m['溫度°C'])));
  const vibMax = Math.max(1, ...vib.map(m => num(m['振動 mm/s'])));
  $('#temperatureBars').innerHTML = temp.map(m => rankRow(m['設備編號'], num(m['溫度°C']), tempMax, '°C', num(m['溫度°C']) >= CONFIG.highTempC)).join('');
  $('#vibrationBars').innerHTML = vib.map(m => rankRow(m['設備編號'], num(m['振動 mm/s']), vibMax, ' mm/s', num(m['振動 mm/s']) >= vibMax * .8)).join('');
}

function rankRow(code, value, max, unit, hot) {
  const width = Math.max(2, Math.min(100, value / max * 100));
  return `<div class="rank-row ${hot ? 'hot':''}"><span class="rank-code">${esc(code)}</span><div class="rank-track"><span style="width:${width}%"></span></div><strong>${fmt.format(value)}${unit}</strong></div>`;
}

function renderOutputs() {
  const total = machines.reduce((s,m) => s + num(m['當日產量']), 0);
  $('#totalOutput').textContent = new Intl.NumberFormat('zh-TW').format(total);
  $('#lineOutput').innerHTML = groupByLine().map(([line,list]) => {
    const sum = list.reduce((s,m) => s + num(m['當日產量']), 0);
    return `<div class="output-chip"><span>${esc(line)}</span><strong>${new Intl.NumberFormat('zh-TW').format(sum)}</strong></div>`;
  }).join('');
}

function hydrateFilters() {
  const select = $('#lineFilter');
  const current = select.value;
  const lines = [...new Set(machines.map(m => m['產線']).filter(Boolean))].sort();
  select.innerHTML = '<option value="all">全部</option>' + lines.map(line => `<option value="${esc(line)}">${esc(line)}</option>`).join('');
  if (lines.includes(current)) select.value = current;
}

function renderTable() {
  const line = $('#lineFilter').value;
  const status = $('#statusFilter').value;
  const q = $('#searchInput').value.trim().toLowerCase();
  const filtered = machines.filter(m => {
    const matchesLine = line === 'all' || m['產線'] === line;
    const matchesStatus = status === 'all' || m['狀態'] === status;
    const haystack = `${m['設備編號']} ${m['設備名稱']} ${m['異常碼']} ${m['建議處置']}`.toLowerCase();
    return matchesLine && matchesStatus && (!q || haystack.includes(q));
  });

  $('#machineTableBody').innerHTML = filtered.map(m => {
    const temp = num(m['溫度°C']);
    return `<tr>
      <td class="machine-id">${esc(m['設備編號'])}</td>
      <td>${esc(m['產線'])}</td>
      <td>${esc(m['設備名稱'])}</td>
      <td>${esc(m['更新時間'])}</td>
      <td class="${temp >= CONFIG.highTempC ? 'metric-hot':''}">${fmt.format(temp)}°C</td>
      <td>${fmt.format(num(m['振動 mm/s']))}</td>
      <td>${fmt.format(num(m['稼動率%']))}%</td>
      <td>${fmt.format(num(m['良率%']))}%</td>
      <td>${new Intl.NumberFormat('zh-TW').format(num(m['當日產量']))}</td>
      <td>${esc(m['異常碼'])}</td>
      <td><span class="status-badge status-${esc(m['狀態'])}">${esc(m['狀態'])}</span></td>
      <td>${esc(m['建議處置'])}</td>
    </tr>`;
  }).join('') || '<tr><td colspan="12"><div class="empty-state">沒有符合條件的設備</div></td></tr>';

  $('#tableCount').textContent = `顯示 ${filtered.length} / ${machines.length} 筆`;
}

function setupTheme() {
  const saved = localStorage.getItem('factory-theme');
  if (saved === 'light' || saved === 'dark') document.body.dataset.theme = saved;
  syncThemeLabel();
  $('#themeToggle').addEventListener('click', () => {
    document.body.dataset.theme = document.body.dataset.theme === 'dark' ? 'light' : 'dark';
    localStorage.setItem('factory-theme', document.body.dataset.theme);
    syncThemeLabel();
  });
}

function syncThemeLabel() {
  const dark = document.body.dataset.theme === 'dark';
  $('#themeLabel').textContent = dark ? '深色' : '淺色';
  $('.theme-icon').textContent = dark ? '☾' : '☀';
}

function setupClock() {
  const tick = () => {
    const now = new Date();
    $('#clock').textContent = now.toLocaleTimeString('zh-TW', { hour12:false });
    $('#today').textContent = now.toLocaleDateString('zh-TW', { year:'numeric', month:'2-digit', day:'2-digit', weekday:'short' });
  };
  tick(); setInterval(tick, 1000);
}

function showToast(msg) {
  const toast = $('#toast');
  toast.textContent = msg;
  toast.classList.add('show');
  clearTimeout(showToast.timer);
  showToast.timer = setTimeout(() => toast.classList.remove('show'), 2200);
}

function bindEvents() {
  $('#refreshBtn').addEventListener('click', () => loadData({ manual:true }));
  $('#lineFilter').addEventListener('change', renderTable);
  $('#statusFilter').addEventListener('change', renderTable);
  $('#searchInput').addEventListener('input', renderTable);
}

function init() {
  $('#pollingLabel').textContent = `${CONFIG.refreshMs / 1000} 秒`;
  setupTheme();
  setupClock();
  bindEvents();
  loadData();
  setInterval(() => loadData(), CONFIG.refreshMs);
}

document.addEventListener('DOMContentLoaded', init);
