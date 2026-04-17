/**
 * Mock Cognos 11 REST API Server
 * Simulates AD auth and report execution for local development and testing.
 * All data is fabricated. No real PII or student records.
 */

require('dotenv').config({ path: '../.env' });

const https = require('https');
const fs = require('fs');
const path = require('path');
const express = require('express');

const app = express();
const PORT = 3001;

app.use(express.json());

// ── Fake data ─────────────────────────────────────────────────────────────────

const FAKE_SESSIONS = new Map();

const FAKE_DEPARTMENTS = [
  { id: 'dept-001', name: 'Academic Affairs' },
  { id: 'dept-002', name: 'Admissions & Records' },
  { id: 'dept-003', name: 'Business Office' },
  { id: 'dept-004', name: 'Campus Safety' },
  { id: 'dept-005', name: 'Financial Aid' },
  { id: 'dept-006', name: 'Human Resources' },
  { id: 'dept-007', name: 'Information Technology' },
  { id: 'dept-008', name: 'Instruction - Business' },
  { id: 'dept-009', name: 'Instruction - STEM' },
  { id: 'dept-010', name: 'Instruction - Health Sciences' },
  { id: 'dept-011', name: 'Instruction - Liberal Arts' },
  { id: 'dept-012', name: 'Library Services' },
  { id: 'dept-013', name: 'Maintenance & Operations' },
  { id: 'dept-014', name: 'Student Services' },
  { id: 'dept-015', name: 'Vice President of Instruction' },
];

const CONTENT_STORE = [
  {
    id: 'folder-budget',
    name: 'Budget Reports',
    type: 'folder',
    children: [
      { id: 'rpt-budget-summary', name: 'Department Budget Summary', type: 'report' },
      { id: 'rpt-budget-vs-actual', name: 'Budget vs. Actual by Department', type: 'report' },
      { id: 'rpt-ytd-expenditure', name: 'YTD Expenditure Report', type: 'report' },
      { id: 'rpt-encumbrances', name: 'Encumbrance Detail', type: 'report' },
    ],
  },
  {
    id: 'folder-gl',
    name: 'General Ledger',
    type: 'folder',
    children: [
      { id: 'rpt-gl-detail', name: 'GL Transaction Detail', type: 'report' },
      { id: 'rpt-gl-summary', name: 'GL Account Summary', type: 'report' },
      { id: 'rpt-trial-balance', name: 'Trial Balance', type: 'report' },
    ],
  },
  {
    id: 'folder-grants',
    name: 'Grants & Projects',
    type: 'folder',
    children: [
      { id: 'rpt-grant-status', name: 'Grant Status Report', type: 'report' },
      { id: 'rpt-grant-expenditure', name: 'Grant Expenditure by Period', type: 'report' },
    ],
  },
  {
    id: 'folder-enrollment',
    name: 'Enrollment Finance',
    type: 'folder',
    children: [
      { id: 'rpt-revenue-by-term', name: 'Revenue by Term', type: 'report' },
      { id: 'rpt-ftes-funding', name: 'FTES & State Funding Summary', type: 'report' },
    ],
  },
];

// ── Auth ──────────────────────────────────────────────────────────────────────

app.put('/api/v1/session', (req, res) => {
  const params = req.body?.parameters || [];
  const usernameParam = params.find((p) => p.name === 'CAMUsername');
  const passwordParam = params.find((p) => p.name === 'CAMPassword');

  const username = usernameParam?.value;
  const password = passwordParam?.value;

  // Accept any non-empty credentials in mock mode
  if (!username || !password) {
    return res.status(401).json({ error: 'Invalid credentials' });
  }

  const token = `MOCK-TOKEN-${Date.now()}-${Math.random().toString(36).slice(2)}`;
  FAKE_SESSIONS.set(token, { username, createdAt: Date.now() });

  res.setHeader('ibm-ba-authorization', token);
  res.json({ session: { token }, user: { defaultName: username } });
});

app.delete('/api/v1/session', (req, res) => {
  const token = req.headers['ibm-ba-authorization'];
  if (token) FAKE_SESSIONS.delete(token);
  res.status(204).send();
});

function requireMockAuth(req, res, next) {
  const token = req.headers['ibm-ba-authorization'];
  if (!token || !FAKE_SESSIONS.has(token)) {
    return res.status(401).json({ error: 'Session expired or invalid' });
  }
  next();
}

// ── Content Store ─────────────────────────────────────────────────────────────

app.get('/api/v1/content', requireMockAuth, (req, res) => {
  const parentId = req.query.parentId;

  if (!parentId) {
    // Return top-level folders
    return res.json({
      items: CONTENT_STORE.map((f) => ({ id: f.id, name: f.name, type: f.type })),
    });
  }

  const folder = CONTENT_STORE.find((f) => f.id === parentId);
  if (!folder) return res.status(404).json({ error: 'Folder not found' });

  res.json({ items: folder.children });
});

// ── Report Parameters ─────────────────────────────────────────────────────────

app.get('/api/v1/reports/:reportId/parameters', requireMockAuth, (req, res) => {
  res.json({
    parameters: [
      {
        name: 'p_department',
        label: 'Department',
        type: 'string',
        required: false,
        allowMultiple: false,
        values: FAKE_DEPARTMENTS.map((d) => ({ use: d.id, display: d.name })),
      },
      {
        name: 'p_startDate',
        label: 'Start Date',
        type: 'date',
        required: false,
      },
      {
        name: 'p_endDate',
        label: 'End Date',
        type: 'date',
        required: false,
      },
    ],
  });
});

// ── Report Execution ──────────────────────────────────────────────────────────

app.post('/api/v1/reports/:reportId/reportData', requireMockAuth, (req, res) => {
  const { reportId } = req.params;
  const { parameterValues = [], outputFormat } = req.body;

  const getParam = (name) =>
    (parameterValues.find((p) => p.name === name)?.value || [])[0] || null;

  const department = getParam('p_department');
  const startDate = getParam('p_startDate') || '2024-07-01';
  const endDate = getParam('p_endDate') || '2025-06-30';

  const deptLabel = department
    ? FAKE_DEPARTMENTS.find((d) => d.id === department)?.name || department
    : 'All Departments';

  const data = generateFakeCrosstab(reportId, deptLabel, startDate, endDate);

  res.json(data);
});

// ── Fake Crosstab Generator ───────────────────────────────────────────────────

function generateFakeCrosstab(reportId, deptLabel, startDate, endDate) {
  const months = getMonthsBetween(startDate, endDate).slice(0, 6);
  const isGrant = reportId.startsWith('rpt-grant');
  const isGL = reportId.startsWith('rpt-gl') || reportId === 'rpt-trial-balance';

  let rowHeaders, budgetBase, actualVariance;

  if (isGrant) {
    rowHeaders = ['Personnel', 'Equipment', 'Supplies', 'Travel', 'Indirect Costs'];
    budgetBase = [48000, 12000, 6500, 3200, 8500];
    actualVariance = [0.92, 0.75, 1.05, 0.60, 0.95];
  } else if (isGL) {
    rowHeaders = [
      'Salaries - Full Time',
      'Salaries - Part Time',
      'Benefits',
      'Operating Supplies',
      'Contracted Services',
      'Equipment',
      'Other Expenses',
    ];
    budgetBase = [210000, 85000, 92000, 18500, 34000, 15000, 7200];
    actualVariance = [0.99, 0.88, 1.01, 0.72, 0.85, 0.40, 1.12];
  } else {
    rowHeaders = [
      'Salaries - Full Time',
      'Salaries - Part Time',
      'Benefits',
      'Operating Supplies',
      'Contracted Services',
      'Equipment',
      'Travel & Conference',
      'Other Operating',
    ];
    budgetBase = [245000, 92000, 105000, 22000, 41000, 18000, 5500, 8200];
    actualVariance = [0.98, 0.91, 1.00, 0.68, 0.82, 0.35, 0.55, 0.90];
  }

  const colHeaders = months;
  const rows = rowHeaders.map((label, ri) => {
    const monthBudget = Math.round(budgetBase[ri] / 12);
    const cells = months.map((_, mi) => {
      const actual = Math.round(monthBudget * actualVariance[ri] * (0.9 + Math.random() * 0.2));
      return actual;
    });
    const rowTotal = cells.reduce((a, b) => a + b, 0);
    return { header: label, cells, total: rowTotal };
  });

  const colTotals = months.map((_, mi) =>
    rows.reduce((sum, r) => sum + r.cells[mi], 0)
  );
  const grandTotal = rows.reduce((sum, r) => sum + r.total, 0);

  return {
    reportId,
    reportTitle: reportTitleFromId(reportId),
    department: deptLabel,
    startDate,
    endDate,
    generatedAt: new Date().toISOString(),
    crosstab: {
      columnHeaders: [...colHeaders, 'Total'],
      rows: rows.map((r) => ({
        header: r.header,
        values: [...r.cells, r.total],
      })),
      totalsRow: {
        header: 'Total',
        values: [...colTotals, grandTotal],
      },
    },
  };
}

function reportTitleFromId(id) {
  const map = {
    'rpt-budget-summary': 'Department Budget Summary',
    'rpt-budget-vs-actual': 'Budget vs. Actual by Department',
    'rpt-ytd-expenditure': 'YTD Expenditure Report',
    'rpt-encumbrances': 'Encumbrance Detail',
    'rpt-gl-detail': 'GL Transaction Detail',
    'rpt-gl-summary': 'GL Account Summary',
    'rpt-trial-balance': 'Trial Balance',
    'rpt-grant-status': 'Grant Status Report',
    'rpt-grant-expenditure': 'Grant Expenditure by Period',
    'rpt-revenue-by-term': 'Revenue by Term',
    'rpt-ftes-funding': 'FTES & State Funding Summary',
  };
  return map[id] || id;
}

function getMonthsBetween(startDate, endDate) {
  const months = [];
  const start = new Date(startDate);
  const end = new Date(endDate);
  const cur = new Date(start.getFullYear(), start.getMonth(), 1);
  while (cur <= end) {
    months.push(
      cur.toLocaleString('default', { month: 'short', year: '2-digit' })
    );
    cur.setMonth(cur.getMonth() + 1);
  }
  return months;
}

// ── Start ─────────────────────────────────────────────────────────────────────

const certBase = path.join(__dirname, '..', 'certs');
let sslOptions;
try {
  sslOptions = {
    cert: fs.readFileSync(path.join(certBase, 'server.crt')),
    key: fs.readFileSync(path.join(certBase, 'server.key')),
  };
} catch (_) {
  console.warn('SSL certs not found — mock server starting in HTTP mode (dev only)');
  sslOptions = null;
}

if (sslOptions) {
  https.createServer(sslOptions, app).listen(PORT, () => {
    console.log(`Mock Cognos server running on https://localhost:${PORT}`);
  });
} else {
  const http = require('http');
  http.createServer(app).listen(PORT, () => {
    console.log(`Mock Cognos server running on http://localhost:${PORT} (no TLS)`);
  });
}
