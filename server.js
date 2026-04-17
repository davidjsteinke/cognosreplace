/**
 * Cognos Add-in Host Server
 * Brokers authentication between the Excel task pane and Cognos REST API.
 * Auth tokens are held server-side in session memory only — never sent to the client.
 */

require('dotenv').config();

const https = require('https');
const fs = require('fs');
const path = require('path');
const express = require('express');
const session = require('express-session');
const cookieParser = require('cookie-parser');
const fetch = require('node-fetch');

const app = express();
const PORT = process.env.SERVER_PORT || 3000;

const COGNOS_BASE =
  process.env.COGNOS_MODE === 'mock'
    ? process.env.COGNOS_MOCK_URL || 'https://localhost:3001'
    : process.env.COGNOS_BASE_URL;

const COGNOS_NAMESPACE = process.env.COGNOS_NAMESPACE || 'CognosEx';

// ── Middleware ──────────────────────────────────────────────────────────────

app.use(cookieParser());
app.use(express.json());
app.use(express.urlencoded({ extended: false }));

app.use(
  session({
    secret: process.env.SESSION_SECRET || 'replace-with-random-secret',
    resave: false,
    saveUninitialized: false,
    cookie: {
      httpOnly: true,
      secure: true,
      sameSite: 'strict',
      maxAge: 8 * 60 * 60 * 1000, // 8 hours
    },
  })
);

// Serve static task pane files
app.use(express.static(path.join(__dirname, 'src')));

// ── Helpers ─────────────────────────────────────────────────────────────────

/**
 * Fetch options that trust the campus internal CA.
 * node-fetch does not read system trust store by default, so we load the
 * CA cert from disk and pass it via the https agent.
 */
function makeFetchAgent() {
  const certPath = process.env.SSL_CA_CERT_PATH;
  if (certPath && fs.existsSync(certPath)) {
    const ca = fs.readFileSync(certPath);
    return new https.Agent({ ca, rejectUnauthorized: true });
  }
  // Dev/mock: skip verification only when explicitly opted in
  if (process.env.COGNOS_MODE === 'mock') {
    return new https.Agent({ rejectUnauthorized: false });
  }
  return undefined; // use default Node trust store
}

async function cognosGet(session_, urlPath) {
  const token = session_.cognosToken;
  if (!token) throw new Error('NOT_AUTHENTICATED');

  const res = await fetch(`${COGNOS_BASE}${urlPath}`, {
    headers: {
      'IBM-BA-Authorization': token,
      Accept: 'application/json',
    },
    agent: makeFetchAgent(),
  });

  if (res.status === 401) {
    session_.cognosToken = null;
    throw new Error('SESSION_EXPIRED');
  }
  if (!res.ok) throw new Error(`COGNOS_ERROR:${res.status}`);
  return res.json();
}

async function cognosPost(session_, urlPath, body) {
  const token = session_.cognosToken;
  if (!token) throw new Error('NOT_AUTHENTICATED');

  const res = await fetch(`${COGNOS_BASE}${urlPath}`, {
    method: 'POST',
    headers: {
      'IBM-BA-Authorization': token,
      'Content-Type': 'application/json',
      Accept: 'application/json',
    },
    body: JSON.stringify(body),
    agent: makeFetchAgent(),
  });

  if (res.status === 401) {
    session_.cognosToken = null;
    throw new Error('SESSION_EXPIRED');
  }
  if (!res.ok) throw new Error(`COGNOS_ERROR:${res.status}`);
  return res.json();
}

function requireAuth(req, res, next) {
  if (!req.session.cognosToken) {
    return res.status(401).json({ error: 'NOT_AUTHENTICATED' });
  }
  next();
}

// ── Auth Routes ─────────────────────────────────────────────────────────────

/**
 * POST /api/auth/login
 * Accepts AD credentials from the task pane (over HTTPS, campus-only network).
 * Exchanges them for a Cognos session token which is stored server-side only.
 */
app.post('/api/auth/login', async (req, res) => {
  const { username, password } = req.body;

  if (!username || !password) {
    return res.status(400).json({ error: 'MISSING_CREDENTIALS' });
  }

  try {
    const loginRes = await fetch(
      `${COGNOS_BASE}/api/v1/session`,
      {
        method: 'PUT',
        headers: {
          'Content-Type': 'application/json',
          Accept: 'application/json',
        },
        body: JSON.stringify({
          parameters: [
            {
              name: 'CAMNamespace',
              value: COGNOS_NAMESPACE,
            },
            {
              name: 'CAMUsername',
              value: username,
            },
            {
              name: 'CAMPassword',
              value: password,
            },
          ],
        }),
        agent: makeFetchAgent(),
      }
    );

    if (!loginRes.ok) {
      const errBody = await loginRes.text();
      console.error('Cognos login failed:', loginRes.status, errBody);
      return res
        .status(401)
        .json({ error: 'AUTH_FAILED', detail: 'Invalid credentials or Cognos auth error' });
    }

    const data = await loginRes.json();
    // Cognos returns the token in the ibm-ba-authorization header on successful auth
    const token =
      loginRes.headers.get('ibm-ba-authorization') ||
      data?.session?.token ||
      data?.token;

    if (!token) {
      return res.status(502).json({ error: 'TOKEN_MISSING', detail: 'Cognos did not return a session token' });
    }

    // Store token server-side only
    req.session.cognosToken = token;
    req.session.username = username;

    return res.json({ ok: true, username });
  } catch (err) {
    console.error('Login error:', err.message);
    return res.status(502).json({ error: 'CONNECT_ERROR', detail: err.message });
  }
});

/**
 * POST /api/auth/logout
 */
app.post('/api/auth/logout', async (req, res) => {
  if (req.session.cognosToken) {
    try {
      await fetch(`${COGNOS_BASE}/api/v1/session`, {
        method: 'DELETE',
        headers: { 'IBM-BA-Authorization': req.session.cognosToken },
        agent: makeFetchAgent(),
      });
    } catch (_) {
      // best-effort logout from Cognos
    }
  }
  req.session.destroy();
  res.json({ ok: true });
});

/**
 * GET /api/auth/status
 */
app.get('/api/auth/status', (req, res) => {
  if (req.session.cognosToken) {
    return res.json({ authenticated: true, username: req.session.username });
  }
  res.json({ authenticated: false });
});

// ── Content Store Browser ────────────────────────────────────────────────────

/**
 * GET /api/content?id=<storeId>
 * Returns child items of a content store folder.
 * Default root is the public folders root.
 */
app.get('/api/content', requireAuth, async (req, res) => {
  const id = req.query.id || 'root';
  const urlPath =
    id === 'root'
      ? '/api/v1/content?fields=id,name,type&sort=name'
      : `/api/v1/content?parentId=${encodeURIComponent(id)}&fields=id,name,type&sort=name`;

  try {
    const data = await cognosGet(req.session, urlPath);
    res.json(data);
  } catch (err) {
    if (err.message === 'NOT_AUTHENTICATED' || err.message === 'SESSION_EXPIRED') {
      return res.status(401).json({ error: err.message });
    }
    res.status(502).json({ error: err.message });
  }
});

// ── Report Execution ─────────────────────────────────────────────────────────

/**
 * POST /api/report/run
 * Body: { reportId, parameters: { department, startDate, endDate, ... }, rawMode: bool }
 * Returns crosstab or flat data depending on rawMode.
 */
app.post('/api/report/run', requireAuth, async (req, res) => {
  const { reportId, parameters = {}, rawMode = false } = req.body;

  if (!reportId) {
    return res.status(400).json({ error: 'MISSING_REPORT_ID' });
  }

  try {
    // Build Cognos report request
    const reportRequest = buildReportRequest(reportId, parameters, rawMode);

    // Submit report to Cognos
    const submitData = await cognosPost(
      req.session,
      `/api/v1/reports/${encodeURIComponent(reportId)}/reportData`,
      reportRequest
    );

    res.json(submitData);
  } catch (err) {
    if (err.message === 'NOT_AUTHENTICATED' || err.message === 'SESSION_EXPIRED') {
      return res.status(401).json({ error: err.message });
    }
    console.error('Report run error:', err.message);
    res.status(502).json({ error: err.message });
  }
});

/**
 * GET /api/report/parameters?reportId=<id>
 * Returns the parameter definitions for a report.
 */
app.get('/api/report/parameters', requireAuth, async (req, res) => {
  const { reportId } = req.query;
  if (!reportId) return res.status(400).json({ error: 'MISSING_REPORT_ID' });

  try {
    const data = await cognosGet(
      req.session,
      `/api/v1/reports/${encodeURIComponent(reportId)}/parameters`
    );
    res.json(data);
  } catch (err) {
    if (err.message === 'NOT_AUTHENTICATED' || err.message === 'SESSION_EXPIRED') {
      return res.status(401).json({ error: err.message });
    }
    res.status(502).json({ error: err.message });
  }
});

// ── Helpers ──────────────────────────────────────────────────────────────────

function buildReportRequest(reportId, parameters, rawMode) {
  const paramList = [];

  if (parameters.department) {
    paramList.push({ name: 'p_department', value: [parameters.department] });
  }
  if (parameters.startDate) {
    paramList.push({ name: 'p_startDate', value: [parameters.startDate] });
  }
  if (parameters.endDate) {
    paramList.push({ name: 'p_endDate', value: [parameters.endDate] });
  }
  if (parameters.fiscalYear) {
    paramList.push({ name: 'p_fiscalYear', value: [parameters.fiscalYear] });
  }
  if (parameters.fundCode) {
    paramList.push({ name: 'p_fundCode', value: [parameters.fundCode] });
  }
  if (parameters.costCenter) {
    paramList.push({ name: 'p_costCenter', value: [parameters.costCenter] });
  }
  if (parameters.budgetView) {
    paramList.push({ name: 'p_budgetView', value: [parameters.budgetView] });
  }
  if (parameters.glAccountType) {
    paramList.push({ name: 'p_glAccountType', value: [parameters.glAccountType] });
  }
  if (parameters.grantCode) {
    paramList.push({ name: 'p_grantCode', value: [parameters.grantCode] });
  }
  if (parameters.reportingPeriod) {
    paramList.push({ name: 'p_reportingPeriod', value: [parameters.reportingPeriod] });
  }
  if (parameters.comparisonPeriod) {
    paramList.push({ name: 'p_comparisonPeriod', value: [parameters.comparisonPeriod] });
  }
  if (parameters.departments && Array.isArray(parameters.departments)) {
    paramList.push({ name: 'p_departments', value: parameters.departments });
  }

  return {
    outputFormat: rawMode ? 'CSV' : 'spreadsheetML',
    parameterValues: paramList,
  };
}

// ── Start Server ─────────────────────────────────────────────────────────────

const sslOptions = {
  cert: fs.readFileSync(process.env.SSL_CERT_PATH || './certs/server.crt'),
  key: fs.readFileSync(process.env.SSL_KEY_PATH || './certs/server.key'),
};

https.createServer(sslOptions, app).listen(PORT, () => {
  console.log(`Cognos Add-in server running on https://localhost:${PORT}`);
  console.log(`Cognos mode: ${process.env.COGNOS_MODE || 'mock'} → ${COGNOS_BASE}`);
});
