require('dotenv').config();
const express = require('express');
const { google } = require('googleapis');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3001;
const SPREADSHEET_ID = process.env.SPREADSHEET_ID || '10qALiiSGlZjmzQYv2TihB9ybzcb2iDoOuqRPtkcnO24';

const SHEETS = {
  uebersicht: process.env.SHEET_UEBERSICHT || 'Übersicht',
  marketing:  process.env.SHEET_MARKETING  || 'Marketing',
  stage2:     process.env.SHEET_STAGE2     || 'Stage 2',
  stage3:     process.env.SHEET_STAGE3     || 'Stage 3',
  rebuy:      process.env.SHEET_REBUY      || 'Rebuy Funnel',
};

const HUBSPOT_TOKEN = process.env.HUBSPOT_ACCESS_TOKEN;

// ─── Middleware ───────────────────────────────────────────────────────────────
app.use((req, res, next) => {
  res.set('Cache-Control', 'no-cache, no-store, must-revalidate');
  next();
});
app.use(express.static(path.join(__dirname, 'public')));

// ─── Auth ─────────────────────────────────────────────────────────────────────
async function getAuthClient() {
  let credentials;
  if (process.env.GOOGLE_CREDENTIALS) {
    credentials = JSON.parse(process.env.GOOGLE_CREDENTIALS);
  } else {
    const credFile = path.join(__dirname, process.env.CREDENTIALS_FILE || 'credentials.json');
    if (!fs.existsSync(credFile)) throw new Error(`No credentials file found: ${credFile}`);
    credentials = JSON.parse(fs.readFileSync(credFile, 'utf8'));
  }
  return new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
  });
}

// ─── Helpers ──────────────────────────────────────────────────────────────────
function toNum(val) {
  if (val === null || val === undefined) return null;
  if (typeof val === 'number') return val;
  let s = String(val).trim().replace(/[€$%\s]/g, '');
  if (!s || s === '-' || s === 'N/A' || s === '#DIV/0!' || s === '#REF!') return null;
  // German number format: dots as thousands separators, comma as decimal
  // "4.000.000" or "1.234,56" → normalize to standard float string
  if (s.includes(',') && s.includes('.')) {
    s = s.replace(/\./g, '').replace(',', '.'); // "1.234,56" → "1234.56"
  } else if (s.includes(',')) {
    s = s.replace(',', '.'); // "3,14" or "1,000" – treat comma as decimal
  } else if (/\.\d{3}(\.|$)/.test(s)) {
    s = s.replace(/\./g, ''); // "4.000.000" → "4000000" (thousands separators)
  }
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}

// Parse the "Jahresziele" section at the top of Übersicht (rows 3-11)
// Returns { totalRevenue, stage2Revenue, stage34Revenue, rebuyRevenue, deals }
// Each value: { plan, stretch }
function parseJahresziele(rows) {
  const targets = {};
  for (const row of rows) {
    if (!row || !row[0]) continue;
    const label = String(row[0]).toLowerCase();
    const plan    = toNum(row[1]);
    const stretch = toNum(row[2]);
    if (!plan) continue;
    if (label.includes('gesamt') || label.includes('total'))           targets.totalRevenue   = { plan, stretch };
    else if (label.includes('poc') || label.includes('stage 2'))       targets.stage2Revenue  = { plan, stretch };
    else if (label.includes('folge') || label.includes('stage 3'))     targets.stage34Revenue = { plan, stretch };
    else if (label.includes('rebuy'))                                  targets.rebuyRevenue   = { plan, stretch };
    else if (label.includes('unterschrift') || label.includes('vp'))   targets.deals          = { plan, stretch };
  }
  return targets;
}

function findLastUpdated(rows) {
  for (const row of rows.slice(0, 6)) {
    for (const cell of (row || [])) {
      const s = String(cell || '');
      const mDMY = s.match(/(\d{1,2})[.\-\/](\d{1,2})[.\-\/](\d{4})/);
      if (mDMY) return `${mDMY[3]}-${mDMY[2].padStart(2,'0')}-${mDMY[1].padStart(2,'0')}`;
      const mISO = s.match(/(\d{4})[.\-\/](\d{2})[.\-\/](\d{2})/);
      if (mISO) return `${mISO[1]}-${mISO[2]}-${mISO[3]}`;
    }
  }
  return null;
}

const MONTH_LABELS_DE = ['Jan','Feb','Mär','Apr','Mai','Jun','Jul','Aug','Sep','Okt','Nov','Dez'];
const MONTH_TOKENS = ['jan','feb','mär','mar','apr','mai','may','jun','jul','aug','sep','okt','oct','nov','dez','dec'];
const TYPE_WORDS   = ['plan','soll','ist','actual','budget','ziel','target','forecast'];

function isMonthCell(val) {
  const s = String(val || '').toLowerCase().trim();
  return MONTH_TOKENS.some(m => s === m || s.startsWith(m + ' ') || s.startsWith(m + '.'));
}

function isTypeWord(val) {
  const s = String(val || '').toLowerCase().trim();
  // Match exact type words OR compound labels like "PLAN TOTAL", "IST TOTAL", "Ist Inbound"
  return TYPE_WORDS.some(t => s === t || s.startsWith(t + ' '));
}

function findMonthHeaderRow(rows) {
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i] || [];
    let count = 0, first = -1;
    for (let j = 0; j < row.length; j++) {
      if (isMonthCell(row[j])) { count++; if (first === -1) first = j; }
    }
    if (count >= 10) return { rowIdx: i, firstCol: first };
  }
  return null;
}

// Map raw metric names to stable keys for the Übersicht
function metricKey(name) {
  const l = name.toLowerCase();
  const isRevenue = l.includes('€') || l.includes('revenue') || l.includes('umsatz');
  if (isRevenue) {
    if (l.includes('gesamt') || l.includes('total')) return 'totalRevenue';
    if (l.includes('stage 2') || l.includes('stage2') || l.includes('poc')) return 'stage2Revenue';
    if (l.includes('stage 3') || l.includes('stage 4') || l.includes('follow')) return 'stage34Revenue';
    if (l.includes('rebuy') || l.includes('renewal')) return 'rebuyRevenue';
    return 'totalRevenue';
  }
  // Deal / VP signature counts
  if (l.includes('deals closed') || l.includes('closed-won') || l.includes('closed won') ||
      l.includes('vp-unterschrift') || l.includes('abschluss') || l.includes('deals gewonnen')) return 'deals';
  if ((l.includes('vp') || l.includes('unterschrift')) && l.includes('stage')) return 'deals';
  return null;
}

function slugify(s) {
  return s.toLowerCase().replace(/[^a-z0-9]+/g, '_').replace(/^_|_$/g, '');
}

// ─── Sheet parser ─────────────────────────────────────────────────────────────
// Sheet layout: row = [MetricName | TypeWord(Plan/Ist) | Jan | Feb | ... | Dec | Q1..Q4 | Annual]
// For sub-rows (Ist/Soll), col 0 is empty and col 1 = "Ist"/"Soll"
function parseSheet(rows, sheetLabel) {
  const result = {
    sheetLabel,
    lastUpdated: findLastUpdated(rows),
    months: MONTH_LABELS_DE,
    metrics: {},
  };

  const header = findMonthHeaderRow(rows);
  if (!header) return result;

  const { rowIdx: hRow, firstCol } = header;

  // Build month column indices from header
  const hRowData = rows[hRow] || [];
  const monthCols = [];
  for (let j = firstCol; j < hRowData.length && monthCols.length < 12; j++) {
    if (isMonthCell(hRowData[j])) monthCols.push(j);
  }
  while (monthCols.length < 12) monthCols.push(firstCol + monthCols.length);

  // Find JAHR and STRETCH columns from header
  const jahrCol    = hRowData.findIndex(c => String(c || '').toUpperCase().trim() === 'JAHR');
  const stretchCol = hRowData.findIndex(c => String(c || '').toUpperCase().trim() === 'STRETCH');

  let currentMetricName = '';

  for (let i = hRow + 1; i < rows.length; i++) {
    const row = rows[i] || [];
    if (row.every(c => c === null || c === undefined || c === '')) continue;

    const col0 = String(row[0] || '').trim();
    const col1 = String(row[1] || '').trim();

    let metricName, typeLabel;

    if (col0 && !isTypeWord(col0)) {
      // col 0 = metric name, col 1 = type (Plan/Ist/...)
      metricName = col0;
      typeLabel  = col1.toLowerCase();
      currentMetricName = col0;
    } else if (!col0 && isTypeWord(col1)) {
      // col 0 empty, col 1 = "Ist" / "Soll" / "Plan" → sub-row of currentMetricName
      metricName = currentMetricName;
      typeLabel  = col1.toLowerCase();
    } else if (!col0 && col1) {
      // col 0 empty, col 1 = new metric name, col 2 = type
      metricName = col1;
      typeLabel  = String(row[2] || '').toLowerCase();
      currentMetricName = col1;
    } else {
      continue;
    }

    if (!metricName) continue;

    // Skip cumulative "Soll" rows
    if (typeLabel === 'soll') continue;

    const monthly = monthCols.map(col => toNum(row[col]));
    if (monthly.every(v => v === null)) continue;

    const isPlan   = typeLabel.includes('plan') || typeLabel.includes('budget') || typeLabel.includes('ziel') || typeLabel.includes('target');
    const isActual = typeLabel.includes('ist') || typeLabel.includes('actual') || typeLabel.includes('aktuell');

    // Sub-category rows (e.g. "Plan Inbound", "Ist Messe") → store as separate metric, don't overwrite parent
    const SUB_QUALIFIERS = ['inbound','messe','andere','direct','organic','paid','trade','fair','referral','event'];
    const isSubCategory = SUB_QUALIFIERS.some(q => typeLabel.includes(q));

    // For sub-categories, key = parent_metric_type (e.g. "websessions_held_plan_inbound")
    // For TOTAL rows ("plan total", "ist total") or exact type words → use parent key
    const isTotal = typeLabel.includes('total');
    let key;
    if (isSubCategory && !isTotal) {
      key = slugify(metricName + '_' + typeLabel);
    } else {
      key = metricKey(metricName) || slugify(metricName);
    }

    if (!result.metrics[key]) {
      result.metrics[key] = {
        name: isSubCategory && !isTotal ? `${metricName} (${typeLabel})` : metricName,
        plan:          new Array(12).fill(null),
        actual:        new Array(12).fill(null),
        annualPlan:    null,
        annualStretch: null,
      };
    }
    const m = result.metrics[key];

    // Annual value: use JAHR column if present, otherwise fall back to last non-zero after monthly cols
    let annualVal = null;
    if (jahrCol >= 0 && typeof row[jahrCol] === 'number' && row[jahrCol] > 0) {
      annualVal = row[jahrCol];
    } else {
      const after = row.slice(monthCols[11] + 1);
      const annNums = after.filter(v => typeof v === 'number' && v > 0);
      annualVal = annNums.length ? annNums[0] : null; // take first (Q1 or nearest), not last (STRETCH)
    }
    const stretchVal = (stretchCol >= 0 && typeof row[stretchCol] === 'number' && row[stretchCol] > 0)
      ? row[stretchCol] : null;

    if (isPlan) {
      // "PLAN TOTAL" overwrites; sub-plan rows already went to their own key above
      m.plan = monthly;
      if (annualVal && annualVal > (m.annualPlan || 0)) m.annualPlan = annualVal;
      if (stretchVal) m.annualStretch = stretchVal;
    } else if (isActual) {
      m.actual = monthly;
    } else if (m.plan.every(v => v === null)) {
      m.plan = monthly;
      if (annualVal) m.annualPlan = annualVal;
    } else {
      m.actual = monthly;
    }
  }

  // Derive annualPlan from plan sum if missing
  for (const m of Object.values(result.metrics)) {
    if (!m.annualPlan) {
      const s = m.plan.reduce((a, v) => a + (v || 0), 0);
      if (s > 0) m.annualPlan = s;
    }
  }

  return result;
}

// ─── Sheet fetch ──────────────────────────────────────────────────────────────
async function fetchRange(sheets, sheetName, range, renderOption = 'UNFORMATTED_VALUE') {
  // Google Sheets API: wrap names with spaces/special chars in single quotes
  const safe = (sheetName.includes(' ') || /[^\x00-\x7E]/.test(sheetName))
    ? `'${sheetName.replace(/'/g, "\\'")}'`
    : sheetName;
  try {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${safe}!${range}`,
      valueRenderOption: renderOption,
    });
    return res.data.values || [];
  } catch (err) {
    console.warn(`Could not fetch "${sheetName}" (${range}): ${err.message}`);
    return null; // null = sheet not found
  }
}

// List all sheet titles in the spreadsheet (for debugging)
async function listSheets(sheets) {
  try {
    const res = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID, fields: 'sheets.properties' });
    return (res.data.sheets || []).map(s => s.properties.title);
  } catch { return []; }
}

// Try multiple possible names for a sheet
async function fetchSheetAny(sheets, candidates, range) {
  for (const name of candidates) {
    const rows = await fetchRange(sheets, name, range);
    if (rows !== null) return { name, rows };
  }
  return { name: null, rows: [] };
}

// ─── API: /api/data ───────────────────────────────────────────────────────────
app.get('/api/data', async (req, res) => {
  try {
    const auth   = await getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth });
    const RANGE  = 'A1:V80';

    const allSheetNames = await listSheets(sheets);
    console.log('Available sheets:', allSheetNames);

    // Flexible name resolution with fallbacks
    const candidates = {
      uebersicht: [SHEETS.uebersicht, 'Übersicht',              'Uebersicht', 'Overview'],
      marketing:  [SHEETS.marketing,  'Marketing',              'Marketing Funnel', 'Mkt'],
      stage2:     [SHEETS.stage2,     'STAGE 2 - PoC Funnel',   'Stage 2', 'Stage2', 'PoC Funnel', 'Stufe 2'],
      stage3:     [SHEETS.stage3,     'STAGE 3 - Folgeprojekte','Stage 3', 'Stage3', 'Folgeprojekte', 'Follow-Up'],
      rebuy:      [SHEETS.rebuy,      'Rebuy Funnel',           'Rebuy',   'Wiederkauf'],
    };

    // Fetch all in parallel (include Jahresziele section of Übersicht)
    const [uRes, mRes, s2Res, s3Res, rbRes, zielRes, rbAvgDealRes] = await Promise.all([
      fetchSheetAny(sheets, candidates.uebersicht, RANGE),
      fetchSheetAny(sheets, candidates.marketing,  RANGE),
      fetchSheetAny(sheets, candidates.stage2,     RANGE),
      fetchSheetAny(sheets, candidates.stage3,     RANGE),
      fetchSheetAny(sheets, candidates.rebuy,      RANGE),
      fetchRange(sheets, candidates.uebersicht[0], 'A3:C11', 'FORMATTED_VALUE'),
      fetchRange(sheets, candidates.rebuy[0],      'P20',    'UNFORMATTED_VALUE'),
    ]);

    const uebersicht = parseSheet(uRes.rows,  uRes.name  || 'Übersicht');
    const marketing  = parseSheet(mRes.rows,  mRes.name  || 'Marketing');
    const stage2     = parseSheet(s2Res.rows, s2Res.name || 'Stage 2');
    const stage3     = parseSheet(s3Res.rows, s3Res.name || 'Stage 3');
    const rebuy      = parseSheet(rbRes.rows, rbRes.name || 'Rebuy Funnel');
    const avgDealVolume = rbAvgDealRes && rbAvgDealRes[0] && rbAvgDealRes[0][0];
    if (avgDealVolume !== null && avgDealVolume !== undefined && avgDealVolume !== '') {
      rebuy.avgDealVolume = Number(avgDealVolume) || null;
    }

    // Override annualPlan with official Scenario-1 targets from Jahresziele section
    const jahresziele = parseJahresziele(zielRes || []);
    for (const [key, val] of Object.entries(jahresziele)) {
      if (uebersicht.metrics[key]) {
        uebersicht.metrics[key].annualPlan    = val.plan;
        uebersicht.metrics[key].annualStretch = val.stretch;
      }
    }
    uebersicht.jahresziele = jahresziele;

    const lastUpdated = uebersicht.lastUpdated || marketing.lastUpdated || null;

    res.json({
      lastUpdated,
      availableSheets: allSheetNames,
      uebersicht,
      marketing,
      stage2,
      stage3,
      rebuy,
    });
  } catch (err) {
    console.error('API error:', err.message);
    res.status(500).json({ error: err.message });
  }
});

// ─── API: /api/sheets ─────────────────────────────────────────────────────────
app.get('/api/sheets', async (req, res) => {
  try {
    const auth   = await getAuthClient();
    const sheets = google.sheets({ version: 'v4', auth });
    const names  = await listSheets(sheets);
    res.json({ sheets: names });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ─── API: /api/config ─────────────────────────────────────────────────────────
app.get('/api/config', (_req, res) => {
  res.json({ hubspotToken: HUBSPOT_TOKEN || null });
});

// ─── HubSpot deals endpoint ───────────────────────────────────────────────────
let portalIdCache = null;
async function getPortalId() {
  if (portalIdCache) return portalIdCache;
  if (!HUBSPOT_TOKEN) return null;
  try {
    const r = await fetch('https://api.hubapi.com/account-info/v3/details', {
      headers: { Authorization: `Bearer ${HUBSPOT_TOKEN}` },
    });
    const d = await r.json();
    portalIdCache = d.portalId;
    return portalIdCache;
  } catch { return null; }
}

// ─── API: /api/hubspot/deals ──────────────────────────────────────────────────
// type=ws     → Stage 2 websession deals  (datum_websession filter)
// type=offers → Stage 2 offers sent       (angebot_verschickt_am filter)
// Accepts ?start=<unix ms>&end=<unix ms> OR ?month=YYYY-MM
app.get('/api/hubspot/deals', async (req, res) => {
  if (!HUBSPOT_TOKEN) return res.status(500).json({ error: 'HUBSPOT_ACCESS_TOKEN not configured' });
  const { month, type, start, end } = req.query;

  let startMs, endMs;
  if (start && end) {
    startMs = start;
    endMs   = end;
  } else if (month) {
    const [yr, mo] = month.split('-');
    startMs = String(Date.UTC(+yr, +mo - 1, 1));
    endMs   = String(Date.UTC(+yr, +mo, 1) - 1);
  } else {
    return res.status(400).json({ error: 'start+end or month param required' });
  }

  const isWs       = type === 'ws';
  const isOffers   = type === 'offers';
  const isWon      = type === 'won';
  const isS3Direct = type === 's3-direct';
  const isS3Phase1 = type === 's3-phase1';

  const S3_PROPERTIES = ['dealname', 'datum_websession', 'amount', 'createdate', 'hs_analytics_source', 'dealstage', 'ai_see_lifecycle_time', 'hs_projected_amount', 'hs_object_id'];

  const filters = isWs
    ? [
        { propertyName: 'datum_websession', operator: 'BETWEEN', value: startMs, highValue: endMs },
        { propertyName: 'pipeline',         operator: 'EQ',      value: '775171' },
      ]
    : isOffers
    ? [
        { propertyName: 'angebot_verschickt_am', operator: 'BETWEEN', value: startMs, highValue: endMs },
        { propertyName: 'pipeline',              operator: 'EQ',      value: '775171' },
      ]
    : isWon
    ? [
        { propertyName: 'hs_v2_date_entered_2660879', operator: 'BETWEEN', value: startMs, highValue: endMs },
        { propertyName: 'pipeline',                   operator: 'EQ',      value: '775171' },
      ]
    : isS3Direct
    ? [
        { propertyName: 'hs_v2_date_entered_25756819', operator: 'BETWEEN', value: startMs, highValue: endMs },
        { propertyName: 'pipeline',                    operator: 'EQ',      value: '9038883' },
        { propertyName: 'hubspot_owner_id',            operator: 'EQ',      value: '616536452' },
      ]
    : isS3Phase1
    ? [
        { propertyName: 'hs_v2_date_entered_25756819', operator: 'BETWEEN', value: startMs, highValue: endMs },
        { propertyName: 'pipeline',                    operator: 'EQ',      value: '9038883' },
        { propertyName: 'hubspot_owner_id',            operator: 'EQ',      value: '398377013' },
      ]
    : [
        { propertyName: 'closedate', operator: 'BETWEEN', value: startMs, highValue: endMs },
        { propertyName: 'pipeline',  operator: 'EQ',      value: '775171' },
      ];

  const properties = isWs
    ? ['dealname', 'datum_websession', 'amount', 'createdate', 'hs_analytics_source', 'dealstage', 'hs_projected_amount', 'hs_object_id']
    : isOffers
    ? ['dealname', 'datum_websession', 'amount', 'createdate', 'hs_analytics_source', 'dealstage', 'time_websession_to_offer', 'hs_projected_amount', 'hs_object_id']
    : isWon
    ? ['dealname', 'datum_websession', 'hs_v2_date_entered_2660879', 'amount', 'createdate', 'hs_analytics_source', 'dealstage', 'in_deal_phase_seit', 'hs_projected_amount', 'hs_object_id']
    : (isS3Direct || isS3Phase1)
    ? S3_PROPERTIES
    : ['dealname', 'amount', 'closedate', 'dealstage', 'hs_analytics_source', 'createdate', 'hs_object_id'];

  const body = {
    filterGroups: [{ filters }],
    limit: 100,
    properties,
    sorts: [{ propertyName: 'createdate', direction: 'ASCENDING' }],
  };

  console.log(`[HubSpot /deals] type=${type || 'default'} startMs=${startMs} (${new Date(+startMs).toISOString()}) endMs=${endMs} (${new Date(+endMs).toISOString()})`);
  console.log('[HubSpot /deals] body:', JSON.stringify(body, null, 2));

  try {
    const [hubRes, portalId] = await Promise.all([
      fetch('https://api.hubapi.com/crm/v3/objects/deals/search', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${HUBSPOT_TOKEN}` },
        body: JSON.stringify(body),
      }),
      getPortalId(),
    ]);
    const data = await hubRes.json();
    if (!hubRes.ok) return res.status(hubRes.status).json({ error: data.message || 'HubSpot error' });

    console.log(`[HubSpot /deals] total returned: ${data.total}, results count: ${(data.results || []).length}`);

    res.json({
      results:  data.results || [],
      total:    data.total ?? (data.results || []).length,
      portalId: portalId || null,
    });
  } catch (err) {
    console.error('HubSpot deals error:', err.message);
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`Sales KPI Dashboard running at http://localhost:${PORT}`);
});
