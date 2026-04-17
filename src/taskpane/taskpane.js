/* global Office, Excel */

/**
 * Cognos Analytics Excel Add-in — Task Pane Controller
 *
 * All Cognos API calls go through the local add-in server (same origin).
 * Cognos session tokens never reach this file — they are held server-side only.
 * Presets are persisted to localStorage on the client machine only.
 */

'use strict';

// ── Constants ─────────────────────────────────────────────────────────────────

const PRESETS_KEY = 'cognos_addin_presets_v1';
const REPORT_META_PREFIX = 'CognosReport_';

// ── State ─────────────────────────────────────────────────────────────────────

let selectedReport = null;        // { id, name }
let departmentList = [];           // [{ id, name }]
let breadcrumbStack = [];          // [{ id, name }]

// ── Office JS init ────────────────────────────────────────────────────────────

Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    initUI();
    checkAuthStatus();
  }
});

// ── UI Init ───────────────────────────────────────────────────────────────────

function initUI() {
  // Login
  document.getElementById('btn-login').addEventListener('click', handleLogin);
  document.getElementById('input-password').addEventListener('keydown', function (e) {
    if (e.key === 'Enter') handleLogin();
  });

  // Logout
  document.getElementById('btn-logout').addEventListener('click', handleLogout);

  // Tabs
  document.querySelectorAll('.tab').forEach(function (tab) {
    tab.addEventListener('click', function () {
      switchTab(this.dataset.tab);
    });
  });

  // Report actions
  document.getElementById('btn-run').addEventListener('click', handleRunReport);
  document.getElementById('btn-refresh').addEventListener('click', handleRefresh);
  document.getElementById('btn-save-preset').addEventListener('click', handleSavePreset);

  // Populate fiscal year options
  populateFiscalYears();

  // Populate multi-dept list placeholder (populated after login)
  populateMultiDeptList([]);

  // Render presets
  renderPresets();

  // Set default dates (current fiscal year Jul-Jun)
  setDefaultDates();
}

function switchTab(tab) {
  document.querySelectorAll('.tab').forEach(function (t) {
    t.classList.toggle('active', t.dataset.tab === tab);
  });
  document.querySelectorAll('.tab-content').forEach(function (c) {
    c.classList.toggle('active', c.id === 'tab-' + tab);
  });
  if (tab === 'presets') renderPresets();
}

function setDefaultDates() {
  const now = new Date();
  const fiscalStart = now.getMonth() >= 6
    ? new Date(now.getFullYear(), 6, 1)
    : new Date(now.getFullYear() - 1, 6, 1);
  const today = now.toISOString().slice(0, 10);
  const fyStart = fiscalStart.toISOString().slice(0, 10);
  document.getElementById('param-start').value = fyStart;
  document.getElementById('param-end').value = today;
}

function populateFiscalYears() {
  const sel = document.getElementById('param-fiscal-year');
  const now = new Date();
  const currentFY = now.getMonth() >= 6 ? now.getFullYear() : now.getFullYear() - 1;
  for (let y = currentFY; y >= currentFY - 4; y--) {
    const opt = document.createElement('option');
    opt.value = 'FY' + y;
    opt.textContent = 'FY ' + y + '–' + (y + 1).toString().slice(2);
    sel.appendChild(opt);
  }
}

function populateDepartmentSelect(depts) {
  departmentList = depts;
  const sel = document.getElementById('param-department');
  sel.innerHTML = '<option value="">(All Departments)</option>';
  depts.forEach(function (d) {
    const opt = document.createElement('option');
    opt.value = d.id;
    opt.textContent = d.name;
    sel.appendChild(opt);
  });
  populateMultiDeptList(depts);
}

function populateMultiDeptList(depts) {
  const container = document.getElementById('multi-dept-list');
  container.innerHTML = '';
  if (!depts.length) {
    container.innerHTML = '<div style="padding:6px 8px;color:#aaa;font-size:12px">No departments loaded</div>';
    return;
  }
  depts.forEach(function (d) {
    const row = document.createElement('label');
    row.className = 'multi-select-item';
    row.innerHTML =
      '<input type="checkbox" value="' + escHtml(d.id) + '" /> ' + escHtml(d.name);
    container.appendChild(row);
  });
}

// ── Auth ──────────────────────────────────────────────────────────────────────

async function checkAuthStatus() {
  try {
    const res = await apiFetch('/api/auth/status');
    const data = await res.json();
    if (data.authenticated) {
      showMainPanel(data.username);
      loadContentRoot();
    } else {
      showLoginPanel();
    }
  } catch (_) {
    showLoginPanel();
  }
}

async function handleLogin() {
  const username = document.getElementById('input-username').value.trim();
  const password = document.getElementById('input-password').value;
  const errEl = document.getElementById('login-error');

  errEl.classList.add('hidden');
  errEl.textContent = '';

  if (!username || !password) {
    showLoginError('Please enter your username and password.');
    return;
  }

  const btn = document.getElementById('btn-login');
  btn.disabled = true;
  btn.textContent = 'Signing in…';

  try {
    const res = await apiFetch('/api/auth/login', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ username, password }),
    });
    const data = await res.json();

    if (!res.ok) {
      throw new Error(data.detail || data.error || 'Login failed');
    }

    document.getElementById('input-password').value = '';
    showMainPanel(data.username);
    loadContentRoot();
  } catch (err) {
    showLoginError(err.message);
  } finally {
    btn.disabled = false;
    btn.textContent = 'Sign In';
  }
}

async function handleLogout() {
  try {
    await apiFetch('/api/auth/logout', { method: 'POST' });
  } catch (_) {
    // ignore
  }
  selectedReport = null;
  breadcrumbStack = [];
  showLoginPanel();
}

function showLoginPanel() {
  document.getElementById('panel-login').classList.remove('hidden');
  document.getElementById('panel-main').classList.add('hidden');
}

function showMainPanel(username) {
  document.getElementById('panel-login').classList.add('hidden');
  document.getElementById('panel-main').classList.remove('hidden');
  document.getElementById('header-username').textContent = username || '';
}

function showLoginError(msg) {
  const el = document.getElementById('login-error');
  el.textContent = msg;
  el.classList.remove('hidden');
}

// ── Content Browser ───────────────────────────────────────────────────────────

async function loadContentRoot() {
  breadcrumbStack = [];
  renderBreadcrumb([]);
  await loadContent('root');
}

async function loadContent(folderId) {
  const list = document.getElementById('content-list');
  list.innerHTML = '<div class="loading-msg">Loading…</div>';

  try {
    const url = folderId === 'root'
      ? '/api/content'
      : '/api/content?id=' + encodeURIComponent(folderId);
    const res = await apiFetch(url);
    if (res.status === 401) { handleLogout(); return; }
    const data = await res.json();
    renderContentList(data.items || []);

    // After loading root, fetch departments for param selectors
    if (folderId === 'root') {
      fetchDepartments();
    }
  } catch (err) {
    list.innerHTML = '<div class="status-error">Failed to load: ' + escHtml(err.message) + '</div>';
  }
}

function renderContentList(items) {
  const list = document.getElementById('content-list');
  list.innerHTML = '';

  if (!items.length) {
    list.innerHTML = '<div class="loading-msg">No items found.</div>';
    return;
  }

  items.forEach(function (item) {
    const row = document.createElement('div');
    row.className = 'content-item item-type-' + item.type;
    row.dataset.id = item.id;
    row.dataset.type = item.type;
    row.dataset.name = item.name;

    row.innerHTML =
      '<span class="item-icon"></span>' +
      '<span class="item-name">' + escHtml(item.name) + '</span>';

    row.addEventListener('click', function () {
      onContentItemClick(item);
    });

    list.appendChild(row);
  });
}

function onContentItemClick(item) {
  if (item.type === 'folder') {
    breadcrumbStack.push({ id: item.id, name: item.name });
    renderBreadcrumb(breadcrumbStack);
    loadContent(item.id);
    hideParamPanel();
  } else if (item.type === 'report') {
    // Highlight selection
    document.querySelectorAll('.content-item').forEach(function (el) {
      el.classList.toggle('selected', el.dataset.id === item.id);
    });
    selectReport(item);
  }
}

function renderBreadcrumb(stack) {
  const bc = document.getElementById('breadcrumb');
  bc.innerHTML = '';

  const rootCrumb = document.createElement('span');
  rootCrumb.className = 'crumb crumb-root';
  rootCrumb.dataset.id = 'root';
  rootCrumb.textContent = 'Home';
  rootCrumb.addEventListener('click', function () {
    breadcrumbStack = [];
    renderBreadcrumb([]);
    loadContent('root');
    hideParamPanel();
  });
  bc.appendChild(rootCrumb);

  stack.forEach(function (crumb, idx) {
    const el = document.createElement('span');
    el.className = 'crumb';
    el.dataset.id = crumb.id;
    el.textContent = crumb.name;
    el.addEventListener('click', function () {
      breadcrumbStack = breadcrumbStack.slice(0, idx + 1);
      renderBreadcrumb(breadcrumbStack);
      loadContent(crumb.id);
      hideParamPanel();
    });
    bc.appendChild(el);
  });
}

async function fetchDepartments() {
  // Departments come from a known report's parameter definitions.
  // Use a well-known report ID to seed the department list.
  try {
    const res = await apiFetch('/api/report/parameters?reportId=rpt-budget-summary');
    if (!res.ok) return;
    const data = await res.json();
    const deptParam = (data.parameters || []).find(function (p) {
      return p.name === 'p_department';
    });
    if (deptParam && deptParam.values) {
      populateDepartmentSelect(
        deptParam.values.map(function (v) {
          return { id: v.use, name: v.display };
        })
      );
    }
  } catch (_) {
    // Not fatal — user can still type values manually
  }
}

// ── Parameter Panel ───────────────────────────────────────────────────────────

function selectReport(item) {
  selectedReport = { id: item.id, name: item.name };
  document.getElementById('selected-report-name').textContent = item.name;
  document.getElementById('panel-params').classList.remove('hidden');
  document.getElementById('report-status').textContent = '';
}

function hideParamPanel() {
  document.getElementById('panel-params').classList.add('hidden');
  selectedReport = null;
  document.querySelectorAll('.content-item').forEach(function (el) {
    el.classList.remove('selected');
  });
}

function collectParameters() {
  const multiDepts = [];
  document.querySelectorAll('#multi-dept-list input[type="checkbox"]:checked').forEach(function (cb) {
    multiDepts.push(cb.value);
  });

  return {
    department: document.getElementById('param-department').value || null,
    startDate: document.getElementById('param-start').value || null,
    endDate: document.getElementById('param-end').value || null,
    fiscalYear: document.getElementById('param-fiscal-year').value || null,
    fundCode: document.getElementById('param-fund').value.trim() || null,
    costCenter: document.getElementById('param-cost-center').value.trim() || null,
    budgetView: document.getElementById('param-budget-view').value || null,
    glAccountType: document.getElementById('param-gl-type').value || null,
    grantCode: document.getElementById('param-grant').value.trim() || null,
    reportingPeriod: document.getElementById('param-reporting-period').value || null,
    comparisonPeriod: document.getElementById('param-comparison').value || null,
    departments: multiDepts.length ? multiDepts : null,
  };
}

// ── Run Report ────────────────────────────────────────────────────────────────

async function handleRunReport() {
  if (!selectedReport) return;

  const params = collectParameters();
  const rawMode = document.getElementById('param-raw-mode').checked;

  clearAllStatus();
  setReportStatus('working', 'Running report…');
  document.getElementById('btn-run').disabled = true;

  try {
    const res = await apiFetch('/api/report/run', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        reportId: selectedReport.id,
        parameters: params,
        rawMode: rawMode,
      }),
    });

    if (res.status === 401) { handleLogout(); return; }
    if (!res.ok) {
      const err = await res.json();
      throw new Error(err.detail || err.error || 'Report execution failed');
    }

    const data = await res.json();
    await placeReportInSheet(data, selectedReport, params, rawMode);
    setReportStatus('ok', 'Report placed in sheet.');
    updateRefreshBar(selectedReport, params, rawMode);
  } catch (err) {
    setReportStatus('error', err.message);
  } finally {
    document.getElementById('btn-run').disabled = false;
  }
}

// ── Place data in Excel ───────────────────────────────────────────────────────

async function placeReportInSheet(reportData, report, params, rawMode) {
  const crosstab = reportData.crosstab;
  if (!crosstab) throw new Error('No crosstab data returned from server.');

  const numCols = 1 + crosstab.columnHeaders.length;
  const numRows = 1 + 1 + crosstab.rows.length + (crosstab.totalsRow ? 1 : 0);

  // ── Phase 1: read selection + check for existing data ────────────────────
  // Must be a separate Excel.run so window.confirm (blocked in Office) is
  // never called inside a batch. We resolve startRow/startCol here too so the
  // write phase targets the same cell even if the user moves the selection
  // while the confirm bar is visible.
  let startRow, startCol, hasData;
  try {
    await Excel.run(async function (context) {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const startCell = context.workbook.getSelectedRange();
      startCell.load('rowIndex,columnIndex');
      await context.sync();

      startRow = startCell.rowIndex;
      startCol = startCell.columnIndex;

      const checkRange = sheet.getRangeByIndexes(startRow, startCol, numRows, numCols);
      checkRange.load('values');
      await context.sync();

      hasData = checkRange.values.some(function (row) {
        return row.some(function (cell) { return cell !== '' && cell !== null; });
      });
    });
  } catch (err) {
    const code = err.code ? ' [' + err.code + ']' : '';
    console.error('placeReportInSheet (check) failed' + code + ':', err.message, err.debugInfo || '');
    throw new Error((err.message || 'Excel read error') + code);
  }

  // ── Phase 2: overwrite confirmation (outside any Excel.run) ──────────────
  if (hasData) {
    const ok = await showConfirmBar('Target area has data. Overwrite it with the report output?');
    if (!ok) return;
  }

  // ── Phase 3: write the report ────────────────────────────────────────────
  try {
    await Excel.run(async function (context) {
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      let r = startRow;

      // Title row — only the first cell (no merge: merge can cause ItemNotFound
      // on some sheet states when paired with pending writes in the same batch)
      const titleRange = sheet.getRangeByIndexes(r, startCol, 1, 1);
      titleRange.values = [[reportData.reportTitle || report.name]];
      titleRange.format.font.bold = true;
      titleRange.format.font.size = 13;
      r++;

      if (rawMode) {
        writeFlatTable(sheet, crosstab, r, startCol);
      } else {
        writeCrosstab(sheet, crosstab, r, startCol);
      }

      // storeReportMetadata does its own intermediate syncs for safe named-range upsert
      await storeReportMetadata(context, sheet, report, params, rawMode);
    });
  } catch (err) {
    const code = err.code ? ' [' + err.code + ']' : '';
    console.error('placeReportInSheet (write) failed' + code + ':', err.message, err.debugInfo || '');
    throw new Error((err.message || 'Excel write error') + code);
  }
}

function writeCrosstab(sheet, crosstab, startRow, startCol) {
  let r = startRow;

  // Column header row
  const colHeaders = [''].concat(crosstab.columnHeaders);
  const headerRange = sheet.getRangeByIndexes(r, startCol, 1, colHeaders.length);
  headerRange.values = [colHeaders];
  headerRange.format.font.bold = true;
  headerRange.format.fill.color = '#003E7E';
  headerRange.format.font.color = '#FFFFFF';
  r++;

  // Data rows
  crosstab.rows.forEach(function (row) {
    const rowData = [row.header].concat(row.values);
    const dataRange = sheet.getRangeByIndexes(r, startCol, 1, rowData.length);
    dataRange.values = [rowData];

    // Alternate row shading
    if (r % 2 === 0) {
      dataRange.format.fill.color = '#F0F4FA';
    }

    // Format number cells
    const numRange = sheet.getRangeByIndexes(r, startCol + 1, 1, row.values.length);
    numRange.numberFormat = [new Array(row.values.length).fill('#,##0')];

    // Bold the row header cell
    sheet.getRangeByIndexes(r, startCol, 1, 1).format.font.bold = true;

    r++;
  });

  // Totals row
  if (crosstab.totalsRow) {
    const totalsData = [crosstab.totalsRow.header].concat(crosstab.totalsRow.values);
    const totalsRange = sheet.getRangeByIndexes(r, startCol, 1, totalsData.length);
    totalsRange.values = [totalsData];
    totalsRange.format.font.bold = true;
    totalsRange.format.fill.color = '#003E7E';
    totalsRange.format.font.color = '#FFFFFF';

    const numRange = sheet.getRangeByIndexes(r, startCol + 1, 1, crosstab.totalsRow.values.length);
    numRange.numberFormat = [new Array(crosstab.totalsRow.values.length).fill('#,##0')];
  }

  // Auto-fit the row header column
  sheet.getRangeByIndexes(startRow, startCol, r - startRow + 2, 1).format.columnWidth = 200;
}

function writeFlatTable(sheet, crosstab, startRow, startCol) {
  if (!crosstab) return;

  const headers = ['Row Category', 'Column Period', 'Value'];
  const headerRange = sheet.getRangeByIndexes(startRow, startCol, 1, headers.length);
  headerRange.values = [headers];
  headerRange.format.font.bold = true;
  headerRange.format.fill.color = '#003E7E';
  headerRange.format.font.color = '#FFFFFF';

  let r = startRow + 1;
  crosstab.rows.forEach(function (row) {
    row.values.forEach(function (val, ci) {
      const colLabel = crosstab.columnHeaders[ci] || 'Total';
      sheet.getRangeByIndexes(r, startCol, 1, 3).values = [[row.header, colLabel, val]];
      r++;
    });
  });
}

// ── Report Metadata (hidden named ranges) ──────────────────────────────────────

async function storeReportMetadata(context, sheet, report, params, rawMode) {
  const metadata = JSON.stringify({
    reportId: report.id,
    reportName: report.name,
    parameters: params,
    rawMode: rawMode,
    lastRefresh: new Date().toISOString(),
  });

  const metaName = REPORT_META_PREFIX + report.id.replace(/[^a-zA-Z0-9_]/g, '_');

  // Write metadata string to a hidden cell at the far right of row 1.
  // Range proxy remains valid across syncs within the same Excel.run context.
  const hiddenCell = sheet.getRangeByIndexes(0, 16383, 1, 1); // XFD1
  hiddenCell.values = [[metadata]];

  // Office JS operations are deferred — synchronous try/catch cannot catch
  // errors from queued ops. Load existing names first so we can safely upsert.
  context.workbook.names.load('items/name');
  await context.sync(); // flushes the hiddenCell write and loads names list

  const nameExists = context.workbook.names.items.some(function (n) {
    return n.name === metaName;
  });

  if (nameExists) {
    context.workbook.names.getItem(metaName).delete();
    await context.sync(); // must flush the delete before re-adding
  }

  context.workbook.names.add(metaName, hiddenCell, 'Cognos report metadata');
  // Caller does not need an additional sync — Excel.run flushes on exit
}

async function loadReportMetadata(context) {
  // Find any named range starting with our prefix in the active sheet
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.load('name');
  await context.sync();

  const names = context.workbook.names;
  names.load('items/name,items/value');
  await context.sync();

  for (let i = 0; i < names.items.length; i++) {
    const n = names.items[i];
    if (n.name.startsWith(REPORT_META_PREFIX)) {
      try {
        const range = names.getItem(n.name).getRange();
        range.load('values');
        await context.sync();
        const raw = range.values[0][0];
        if (raw) return JSON.parse(raw);
      } catch (_) {}
    }
  }
  return null;
}

// ── Refresh ───────────────────────────────────────────────────────────────────

async function handleRefresh() {
  clearAllStatus();
  setGlobalStatus('working', 'Refreshing…');
  document.getElementById('btn-refresh').disabled = true;

  try {
    let meta = null;
    await Excel.run(async function (context) {
      meta = await loadReportMetadata(context);
    });

    if (!meta) {
      setGlobalStatus('error', 'No Cognos report metadata found in this sheet.');
      return;
    }

    const res = await apiFetch('/api/report/run', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        reportId: meta.reportId,
        parameters: meta.parameters,
        rawMode: meta.rawMode,
      }),
    });

    if (res.status === 401) { handleLogout(); return; }
    if (!res.ok) throw new Error('Refresh failed: ' + res.status);

    const data = await res.json();
    const report = { id: meta.reportId, name: meta.reportName };
    await placeReportInSheet(data, report, meta.parameters, meta.rawMode);
    setGlobalStatus('ok', 'Refreshed at ' + new Date().toLocaleTimeString());
    updateRefreshBar(report, meta.parameters, meta.rawMode);
  } catch (err) {
    setGlobalStatus('error', err.message);
  } finally {
    document.getElementById('btn-refresh').disabled = false;
  }
}

function updateRefreshBar(report, params, rawMode) {
  const bar = document.getElementById('refresh-bar');
  const info = document.getElementById('refresh-info');
  bar.classList.remove('hidden');
  const deptLabel = params.department
    ? (departmentList.find(function (d) { return d.id === params.department; }) || {}).name || params.department
    : 'All Depts';
  info.textContent =
    report.name + ' · ' + deptLabel + ' · ' + new Date().toLocaleTimeString();
}

// ── Presets ───────────────────────────────────────────────────────────────────

async function handleSavePreset() {
  if (!selectedReport) return;
  clearAllStatus();

  const name = await showPresetNameInput();
  if (!name) return;

  const preset = {
    id: 'preset-' + Date.now(),
    name: name,
    report: selectedReport,
    parameters: collectParameters(),
    rawMode: document.getElementById('param-raw-mode').checked,
    savedAt: new Date().toISOString(),
  };

  const presets = loadPresets();
  presets.push(preset);
  savePresets(presets);
  renderPresets();
  setReportStatus('ok', 'Preset "' + escHtml(preset.name) + '" saved.');
}

function loadPresets() {
  try {
    return JSON.parse(localStorage.getItem(PRESETS_KEY) || '[]');
  } catch (_) {
    return [];
  }
}

function savePresets(presets) {
  try {
    localStorage.setItem(PRESETS_KEY, JSON.stringify(presets));
  } catch (err) {
    setReportStatus('error', 'Could not save preset: ' + err.message);
  }
}

function renderPresets() {
  const list = document.getElementById('presets-list');
  const presets = loadPresets();

  if (!presets.length) {
    list.innerHTML = '<div class="empty-msg">No presets saved yet.</div>';
    return;
  }

  list.innerHTML = '';
  presets.forEach(function (preset) {
    const item = document.createElement('div');
    item.className = 'preset-item';

    const deptLabel = preset.parameters.department
      ? (departmentList.find(function (d) { return d.id === preset.parameters.department; }) || {}).name || preset.parameters.department
      : 'All Depts';

    item.innerHTML =
      '<div style="flex:1;overflow:hidden;">' +
        '<div class="preset-name">' + escHtml(preset.name) + '</div>' +
        '<div class="preset-sub">' + escHtml(preset.report.name) + ' · ' + escHtml(deptLabel) + '</div>' +
      '</div>' +
      '<div class="preset-actions">' +
        '<button class="btn btn-secondary btn-sm" data-id="' + escAttr(preset.id) + '" data-action="run">Run</button>' +
        '<button class="btn btn-link btn-sm" data-id="' + escAttr(preset.id) + '" data-action="delete" style="color:#c0392b">✕</button>' +
      '</div>';

    list.appendChild(item);
  });

  list.querySelectorAll('button[data-action]').forEach(function (btn) {
    btn.addEventListener('click', function () {
      const id = this.dataset.id;
      const action = this.dataset.action;
      if (action === 'run') runPreset(id);
      else if (action === 'delete') deletePreset(id);
    });
  });
}

async function runPreset(presetId) {
  const presets = loadPresets();
  const preset = presets.find(function (p) { return p.id === presetId; });
  if (!preset) return;

  clearAllStatus();
  setGlobalStatus('working', 'Running preset "' + preset.name + '"…');

  try {
    const res = await apiFetch('/api/report/run', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        reportId: preset.report.id,
        parameters: preset.parameters,
        rawMode: preset.rawMode,
      }),
    });

    if (res.status === 401) { handleLogout(); return; }
    if (!res.ok) throw new Error('Report run failed: ' + res.status);

    const data = await res.json();
    await placeReportInSheet(data, preset.report, preset.parameters, preset.rawMode);
    setGlobalStatus('ok', 'Preset "' + preset.name + '" placed in sheet.');
    updateRefreshBar(preset.report, preset.parameters, preset.rawMode);
  } catch (err) {
    setGlobalStatus('error', err.message);
  }
}

function deletePreset(presetId) {
  const itemEl = document.querySelector('[data-action="delete"][data-id="' + CSS.escape(presetId) + '"]')
    && document.querySelector('[data-action="delete"][data-id="' + CSS.escape(presetId) + '"]').closest('.preset-item');
  if (itemEl) {
    showDeleteConfirmInItem(itemEl, presetId);
  } else {
    // Fallback if item can't be located (shouldn't happen in normal flow)
    const presets = loadPresets().filter(function (p) { return p.id !== presetId; });
    savePresets(presets);
    renderPresets();
  }
}

// ── In-pane dialog helpers (replaces window.confirm / alert / prompt) ─────────

/**
 * Shows the shared #confirm-bar and resolves true (Yes) or false (Cancel).
 * The bar sits below the tab content and above the refresh bar, always visible.
 */
function showConfirmBar(message) {
  return new Promise(function (resolve) {
    const bar = document.getElementById('confirm-bar');
    const msg = document.getElementById('confirm-bar-msg');
    const btnYes = document.getElementById('confirm-bar-yes');
    const btnNo = document.getElementById('confirm-bar-no');

    msg.textContent = message;
    bar.classList.remove('hidden');

    function finish(result) {
      bar.classList.add('hidden');
      // Clone to drop any previous listeners without leaking memory
      btnYes.replaceWith(btnYes.cloneNode(true));
      btnNo.replaceWith(btnNo.cloneNode(true));
      resolve(result);
    }

    // Re-query after replaceWith on subsequent calls
    document.getElementById('confirm-bar-yes').addEventListener('click', function () { finish(true); });
    document.getElementById('confirm-bar-no').addEventListener('click', function () { finish(false); });
  });
}

/**
 * Shows the inline #preset-name-area input and resolves with the trimmed name
 * string, or null if the user cancels.
 */
function showPresetNameInput() {
  return new Promise(function (resolve) {
    const area = document.getElementById('preset-name-area');
    const field = document.getElementById('preset-name-field');
    const btnSave = document.getElementById('preset-name-save');
    const btnCancel = document.getElementById('preset-name-cancel');

    field.value = '';
    area.classList.remove('hidden');
    field.focus();

    function finish(value) {
      area.classList.add('hidden');
      btnSave.replaceWith(btnSave.cloneNode(true));
      btnCancel.replaceWith(btnCancel.cloneNode(true));
      field.removeEventListener('keydown', onKey);
      resolve(value);
    }

    function onKey(e) {
      if (e.key === 'Enter') finish(field.value.trim() || null);
      if (e.key === 'Escape') finish(null);
    }

    document.getElementById('preset-name-save').addEventListener('click', function () {
      finish(field.value.trim() || null);
    });
    document.getElementById('preset-name-cancel').addEventListener('click', function () {
      finish(null);
    });
    field.addEventListener('keydown', onKey);
  });
}

/**
 * Replaces the ✕ button in a preset item with inline Yes/No delete confirmation.
 * On cancel, restores the original buttons and re-wires their listeners.
 */
function showDeleteConfirmInItem(itemEl, presetId) {
  const actionsEl = itemEl.querySelector('.preset-actions');

  actionsEl.innerHTML =
    '<span style="font-size:12px;color:#555;margin-right:4px">Delete?</span>' +
    '<button class="btn btn-sm confirm-delete-yes" ' +
      'style="background:#c0392b;color:#fff;border-color:#c0392b">' +
      'Yes' +
    '</button>' +
    '<button class="btn btn-secondary btn-sm confirm-delete-no">No</button>';

  actionsEl.querySelector('.confirm-delete-yes').addEventListener('click', function () {
    const presets = loadPresets().filter(function (p) { return p.id !== presetId; });
    savePresets(presets);
    renderPresets();
  });

  actionsEl.querySelector('.confirm-delete-no').addEventListener('click', function () {
    // Restore original buttons by re-rendering the whole list
    renderPresets();
  });
}

// ── Utilities ─────────────────────────────────────────────────────────────────

async function apiFetch(url, opts) {
  return fetch(url, opts || {});
}

function setReportStatus(type, msg) {
  const el = document.getElementById('report-status');
  el.className = 'status-area' + (type === 'error' ? ' status-error' : type === 'ok' ? ' status-ok' : ' status-working');
  el.textContent = msg;
}

function setGlobalStatus(type, msg) {
  const el = document.getElementById('global-status');
  el.className = 'status-area' + (type === 'error' ? ' status-error' : type === 'ok' ? ' status-ok' : ' status-working');
  el.textContent = msg;
}

function clearAllStatus() {
  const r = document.getElementById('report-status');
  r.className = 'status-area';
  r.textContent = '';
  const g = document.getElementById('global-status');
  g.className = 'status-area';
  g.textContent = '';
  // Also hide the confirm bar if it was left visible by a prior cancelled action
  document.getElementById('confirm-bar').classList.add('hidden');
}

function escHtml(str) {
  if (str == null) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function escAttr(str) {
  return escHtml(str);
}
