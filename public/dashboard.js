/* ═══════════════════════════════════════════════════════════
   AI.SEE Sales KPI Dashboard – dashboard.js
═══════════════════════════════════════════════════════════ */

// ─── Colours ──────────────────────────────────────────────
const C = {
  navyDark:  '#0D2137',
  navy:      '#1B3A5C',
  blue:      '#2B7CE9',
  skyBlue:   '#5BB8F5',
  blueFade:  '#C9DFF7',
  green:     '#22C55E',
  greenFade: '#DCFCE7',
  orange:    '#F59E0B',
  orangeFade:'#FEF9C3',
  purple:    '#8B5CF6',
  purpleFade:'#EDE9FE',
  red:       '#EF4444',
  redFade:   '#FEE2E2',
  gray:      '#9CA3AF',
};

const MONTHS = ['Jan','Feb','Mär','Apr','Mai','Jun','Jul','Aug','Sep','Okt','Nov','Dez'];

// ─── State ────────────────────────────────────────────────
let activeTab = 'uebersicht';
let apiData   = null;
const chartRegistry = {};

// ─── Init ─────────────────────────────────────────────────
window.addEventListener('DOMContentLoaded', loadData);

// ─── Tab switching ────────────────────────────────────────
function switchTab(tab) {
  activeTab = tab;
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  document.getElementById(`tab-${tab}`)?.classList.add('active');

  document.querySelectorAll('.view').forEach(v => v.classList.add('hidden'));
  document.getElementById(`view-${tab}`)?.classList.remove('hidden');

  const titles = {
    uebersicht: ['AI.SEE Funnel Reporting', 'Übersicht aller Vertriebsstufen'],
    marketing:  ['Marketing Funnel', 'Leads & Conversion in den Vertrieb'],
    stage2:     ['Stage 2 – PoC Funnel', 'Proof-of-Concept Abschlüsse & Umsatz'],
    stage3:     ['Stage 3 – Follow-Up Orders', 'Folgeaufträge & Umsatzentwicklung'],
    rebuy:      ['Rebuy Funnel', 'Wiederkehrende Kunden & Erneuerungen'],
  };
  const [title, sub] = titles[tab] || ['AI.SEE Funnel Reporting', ''];
  document.getElementById('pageTitle').textContent    = title;
  document.getElementById('pageSubtitle').textContent = sub;

  if (apiData) renderTab(tab, apiData);
}

// ─── Data load ────────────────────────────────────────────
async function loadData() {
  setUiState('loading');
  try {
    const res = await fetch('/api/data');
    if (!res.ok) throw new Error(`Server error ${res.status}: ${await res.text()}`);
    apiData = await res.json();
    if (apiData.error) throw new Error(apiData.error);

    // Last updated badge
    if (apiData.lastUpdated) {
      const d = new Date(apiData.lastUpdated);
      const label = d.toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' });
      document.getElementById('dateBadge').textContent = `Aktualisiert: ${label}`;
    }

    setUiState('data');
    renderTab(activeTab, apiData);
  } catch (err) {
    console.error(err);
    setUiState('error', err.message);
  }
}

function setUiState(state, msg = '') {
  document.getElementById('loadingState').classList.toggle('hidden', state !== 'loading');
  document.getElementById('errorState').classList.toggle('hidden',   state !== 'error');
  document.querySelectorAll('.view').forEach(v => v.classList.toggle('hidden', state !== 'data'));
  if (state === 'error')  document.getElementById('errorMsg').textContent = msg;
  if (state === 'data') {
    document.querySelectorAll('.view').forEach(v => v.classList.add('hidden'));
    document.getElementById(`view-${activeTab}`)?.classList.remove('hidden');
  }
}

// ─── Tab dispatch ─────────────────────────────────────────
function renderTab(tab, data) {
  switch (tab) {
    case 'uebersicht': renderUebersicht(data.uebersicht); break;
    case 'marketing':  renderFunnelTab('mkt', data.marketing,  'Marketing');   break;
    case 'stage2':     renderFunnelTab('s2',  data.stage2,     'Stage 2');     break;
    case 'stage3':     renderFunnelTab('s3',  data.stage3,     'Stage 3');     break;
    case 'rebuy':      renderFunnelTab('rb',  data.rebuy,      'Rebuy Funnel');break;
  }
}

// ══════════════════════════════════════════════════════════
// ÜBERSICHT RENDERER
// ══════════════════════════════════════════════════════════
function renderUebersicht(d) {
  if (!d) return;
  const m = d.metrics || {};

  // ─── KPI Cards ──
  const CARD_DEFS = [
    { id: 'totalRevenue',  label: 'Gesamtumsatz',   fmt: fmtEur,  prefix: 'ue-totalRevenue'  },
    { id: 'stage2Revenue', label: 'Stage 2 Umsatz', fmt: fmtEur,  prefix: 'ue-stage2Revenue' },
    { id: 'stage34Revenue',label: 'Stage 3/4',      fmt: fmtEur,  prefix: 'ue-stage34Revenue'},
    { id: 'rebuyRevenue',  label: 'Rebuy Umsatz',   fmt: fmtEur,  prefix: 'ue-rebuyRevenue'  },
    { id: 'deals',         label: 'Deals',          fmt: fmtInt,  prefix: 'ue-deals'         },
  ];

  for (const def of CARD_DEFS) {
    const metric = m[def.id];
    const ytdActual = metric ? ytdSum(metric.actual) : null;
    const ytdPlan   = metric ? ytdSum(metric.plan)   : null;
    const annPlan   = metric?.annualPlan ?? null;

    setText(`${def.prefix}`, ytdActual !== null ? def.fmt(ytdActual) : '–');
    if (ytdPlan !== null && ytdPlan > 0) {
      const pct = (ytdActual / ytdPlan * 100).toFixed(0);
      setText(`${def.prefix}Sub`, `${pct}% des Plan-YTD` + (annPlan ? ` | Jahresziel: ${def.fmt(annPlan)}` : ''));
    } else if (annPlan) {
      const pct = ytdActual !== null ? (ytdActual / annPlan * 100).toFixed(0) : null;
      setText(`${def.prefix}Sub`, pct ? `${pct}% des Jahresziels (${def.fmt(annPlan)})` : `Jahresziel: ${def.fmt(annPlan)}`);
    }
  }

  // ─── Revenue Trend Chart ──
  const totalRev = m.totalRevenue;
  if (totalRev) {
    createBarChart('ue-barRevenue', {
      labels: MONTHS,
      datasets: [
        barDataset('Plan', totalRev.plan,   C.blueFade),
        barDataset('Ist',  totalRev.actual, C.blue),
      ],
      yLabel: '€',
      yFmt: fmtEurShort,
    });

    // Cumulative line
    createLineChart('ue-lineRevenueCum', {
      labels: MONTHS,
      datasets: [
        lineDataset('Plan (kum.)', cumulative(totalRev.plan),   C.blueFade, C.blue),
        lineDataset('Ist (kum.)',  cumulative(totalRev.actual), C.blue,     C.blue),
      ],
      yFmt: fmtEurShort,
    });
  } else {
    showNoData('ue-barRevenue');
    showNoData('ue-lineRevenueCum');
  }

  // ─── Annual Target Progress Bars ──
  const progressDefs = [
    { id: 'totalRevenue',   label: 'Gesamtumsatz',     fmt: fmtEurShort },
    { id: 'stage2Revenue',  label: 'Stage 2 Umsatz',   fmt: fmtEurShort },
    { id: 'stage34Revenue', label: 'Stage 3/4 Umsatz', fmt: fmtEurShort },
    { id: 'rebuyRevenue',   label: 'Rebuy Umsatz',     fmt: fmtEurShort },
    { id: 'deals',          label: 'Deals gewonnen',   fmt: fmtInt       },
  ];
  renderProgressBars('ue-progressBars', m, progressDefs);

  // ─── Donut: Revenue split ──
  const stage2Ytd  = m.stage2Revenue  ? ytdSum(m.stage2Revenue.actual)  : 0;
  const stage34Ytd = m.stage34Revenue ? ytdSum(m.stage34Revenue.actual) : 0;
  const rebuyYtd   = m.rebuyRevenue   ? ytdSum(m.rebuyRevenue.actual)   : 0;
  const total      = stage2Ytd + stage34Ytd + rebuyYtd;

  if (total > 0) {
    createDoughnutChart('ue-donut', {
      labels: ['Stage 2 (PoC)', 'Stage 3/4 (Follow-Up)', 'Rebuy'],
      data:   [stage2Ytd, stage34Ytd, rebuyYtd],
      colors: [C.purple, C.blue, C.green],
      fmt:    fmtEurShort,
    });
  } else {
    showNoData('ue-donut');
  }

  // ─── Bar: Split comparison ──
  const splitDatasets = [];
  if (m.stage2Revenue)  splitDatasets.push(barDataset('Stage 2',    m.stage2Revenue.actual,  C.purple));
  if (m.stage34Revenue) splitDatasets.push(barDataset('Stage 3/4',  m.stage34Revenue.actual, C.blue));
  if (m.rebuyRevenue)   splitDatasets.push(barDataset('Rebuy',      m.rebuyRevenue.actual,   C.green));

  if (splitDatasets.length) {
    createBarChart('ue-barSplit', {
      labels: MONTHS,
      datasets: splitDatasets,
      stacked: true,
      yFmt: fmtEurShort,
    });
  } else {
    showNoData('ue-barSplit');
  }

  // ─── Insights ──
  renderInsights('ue-insights', buildUebersichtInsights(m));
}

function buildUebersichtInsights(m) {
  const insights = [];
  const METRIC_LABELS = {
    totalRevenue:   'Gesamtumsatz',
    stage2Revenue:  'Stage 2 Umsatz',
    stage34Revenue: 'Stage 3/4 Umsatz',
    rebuyRevenue:   'Rebuy Umsatz',
    deals:          'Deals',
  };

  for (const [key, metric] of Object.entries(m)) {
    const ytdA = ytdSum(metric.actual);
    const ytdP = ytdSum(metric.plan);
    if (ytdA === null || ytdP === null || ytdP === 0) continue;
    const pct  = ytdA / ytdP * 100;
    const label = METRIC_LABELS[key] || metric.name;
    const fmt  = key === 'deals' ? fmtInt : fmtEurShort;

    if (pct >= 100) {
      insights.push({ text: `<span class="tag-good">✓ ${label}:</span> Plan-YTD übertroffen (${fmt(ytdA)} / ${fmt(ytdP)} = ${pct.toFixed(0)}%)`, type: 'good' });
    } else if (pct >= 75) {
      insights.push({ text: `<span class="tag-warn">~ ${label}:</span> ${pct.toFixed(0)}% des Plan-YTD erreicht (${fmt(ytdA)} von ${fmt(ytdP)})`, type: 'warn' });
    } else {
      insights.push({ text: `<span class="tag-bad">✗ ${label}:</span> Nur ${pct.toFixed(0)}% des Plan-YTD (${fmt(ytdA)} von ${fmt(ytdP)}) – Handlungsbedarf`, type: 'bad' });
    }
  }

  if (!insights.length) {
    insights.push({ text: 'Daten werden geladen oder sind noch nicht verfügbar.' });
  }
  return insights;
}

// ══════════════════════════════════════════════════════════
// GENERIC FUNNEL TAB RENDERER
// ══════════════════════════════════════════════════════════
function renderFunnelTab(prefix, sheetData, tabLabel) {
  if (!sheetData) {
    showTabNoData(prefix, tabLabel);
    return;
  }

  const metrics = sheetData.metrics || {};
  const keys    = Object.keys(metrics);

  // ─── KPI Cards ──
  const cardsEl = document.getElementById(`${prefix}-cards`);
  if (cardsEl) {
    cardsEl.innerHTML = '';
    if (!keys.length) {
      cardsEl.innerHTML = `<div class="no-data"><strong>Keine Daten gefunden</strong>Überprüfe den Sheet-Namen in der .env Datei.</div>`;
    } else {
      const colors = [C.blue, C.green, C.orange, C.purple, C.red, C.skyBlue];
      const borderClasses = ['border-blue','border-green','border-orange','border-purple','border-red','border-blue'];
      const cardKeys = prefix === 's2'
        ? ['deals_created_stage_2', 'websessions_held', 'offers_sent', 'deals_won_stage_2', 'totalRevenue'].filter(k => metrics[k])
        : prefix === 'rb'
          ? keys.filter(k => k !== 'deal_volume' && k !== 'totalRevenue').slice(0, 5)
          : keys.slice(0, 6);
      cardKeys.forEach((key, i) => {
        const metric = metrics[key];
        const ytdA = ytdSum(metric.actual);
        const ytdP = ytdSum(metric.plan);
        const annP = metric.annualPlan;
        const isRevenue = (metric.name || '').toLowerCase().includes('umsatz') ||
                          (metric.name || '').toLowerCase().includes('revenue') ||
                          (metric.name || '').toLowerCase().includes('€') ||
                          (ytdA && ytdA > 1000);
        const fmt = isRevenue ? fmtEur : fmtInt;
        const pct = (ytdA && ytdP && ytdP > 0) ? (ytdA / ytdP * 100) : null;
        const chipClass = pct ? (pct >= 100 ? 'good' : pct >= 75 ? 'warn' : 'bad') : '';

        cardsEl.innerHTML += `
          <div class="kpi-card ${borderClasses[i]}">
            <span class="kpi-label">${metric.name || key.replace(/_/g,' ').replace(/\b\w/g,c=>c.toUpperCase())}</span>
            <span class="kpi-value">${ytdA !== null ? fmt(ytdA) : '–'}</span>
            ${pct !== null ? `<span class="kpi-chip ${chipClass}">${pct.toFixed(0)}% YTD</span>` : ''}
            <span class="kpi-sub">${annP ? `Jahresziel: ${fmt(annP)}` : 'YTD Ist'}</span>
          </div>`;
      });

      // Rebuy: append Avg Deal Volume card from P20
      if (prefix === 'rb' && sheetData.avgDealVolume != null) {
        const idx = cardKeys.length;
        const bc  = borderClasses[idx % borderClasses.length];
        cardsEl.innerHTML += `
          <div class="kpi-card ${bc}">
            <span class="kpi-label">Ø Deal Volume</span>
            <span class="kpi-value">${fmtEur(sheetData.avgDealVolume)}</span>
            <span class="kpi-sub">YTD Ist</span>
          </div>`;
      }
    }
  }

  if (!keys.length) {
    showTabNoData(prefix, tabLabel);
    return;
  }

  // ─── Main bar chart ──
  let mainMetric;
  if (prefix === 's2') {
    mainMetric = metrics['websessions_held'] || findBestMetric(metrics);
  } else if (prefix === 's3') {
    mainMetric = metrics['deals_won_stage_3_direct_nl'] || findBestMetric(metrics);
  } else {
    mainMetric = findBestMetric(metrics);
  }

  if (mainMetric) {
    const mainFmt = prefix === 's3' ? fmtInt : guessFormat(mainMetric);
    createBarChart(`${prefix}-barMain`, {
      labels: MONTHS,
      datasets: [
        barDataset('Plan', mainMetric.plan,   C.blueFade),
        barDataset('Ist',  mainMetric.actual, C.blue),
      ],
      yFmt: mainFmt,
      onClickIst: prefix === 's2' ? (monthLabel) => openWsDealsModal(monthLabel)
                : prefix === 's3' ? (monthLabel) => openS3DirectModal(monthLabel)
                : null,
    });

    createLineChart(`${prefix}-lineCum`, {
      labels: MONTHS,
      datasets: [
        lineDataset('Plan (kum.)', cumulative(mainMetric.plan),   C.blueFade, C.blue),
        lineDataset('Ist (kum.)',  cumulative(mainMetric.actual), C.blue,     C.blue),
      ],
      yFmt: mainFmt,
    });
  } else {
    showNoData(`${prefix}-barMain`);
    showNoData(`${prefix}-lineCum`);
  }

  // ─── Stage 3 extra chart: Phase 1 (JH) ──
  if (prefix === 's3') {
    const phase1 = metrics['deals_won_stage_3_phase_1_jh'];
    if (phase1) {
      createBarChart('s3-barPhase1', {
        labels: MONTHS,
        datasets: [
          barDataset('Plan', phase1.plan,   C.greenFade),
          barDataset('Ist',  phase1.actual, C.green),
        ],
        yFmt: fmtInt,
        onClickIst: (monthLabel) => openS3Phase1Modal(monthLabel),
      });
      createLineChart('s3-lineCumPhase1', {
        labels: MONTHS,
        datasets: [
          lineDataset('Plan (kum.)', cumulative(phase1.plan),   C.greenFade, C.green),
          lineDataset('Ist (kum.)',  cumulative(phase1.actual), C.green,     C.green),
        ],
        yFmt: fmtInt,
      });
    } else {
      showNoData('s3-barPhase1');
      showNoData('s3-lineCumPhase1');
    }
  }

  // ─── Stage 2 extra charts: Offers Sent & Deals Won ──
  if (prefix === 's2') {
    const offersMetric = metrics['offers_sent'];
    if (offersMetric) {
      createBarChart('s2-barOffers', {
        labels: MONTHS,
        datasets: [
          barDataset('Plan', offersMetric.plan,   C.greenFade),
          barDataset('Ist',  offersMetric.actual, C.green),
        ],
        yFmt: fmtInt,
        onClickIst: (monthLabel) => openOffersDealsModal(monthLabel),
      });
      createLineChart('s2-lineCumOffers', {
        labels: MONTHS,
        datasets: [
          lineDataset('Plan (kum.)', cumulative(offersMetric.plan),   C.greenFade, C.green),
          lineDataset('Ist (kum.)',  cumulative(offersMetric.actual), C.green,                   C.green),
        ],
        yFmt: fmtInt,
      });
    } else {
      showNoData('s2-barOffers');
      showNoData('s2-lineCumOffers');
    }

    const dealsMetric = metrics['deals_won_stage_2'];
    if (dealsMetric) {
      createBarChart('s2-barDeals', {
        labels: MONTHS,
        datasets: [
          barDataset('Plan', dealsMetric.plan,   C.orangeFade),
          barDataset('Ist',  dealsMetric.actual, C.orange),
        ],
        yFmt: fmtInt,
        onClickIst: (monthLabel) => openDealsWonModal(monthLabel),
      });
      createLineChart('s2-lineCumDeals', {
        labels: MONTHS,
        datasets: [
          lineDataset('Plan (kum.)', cumulative(dealsMetric.plan),   C.orangeFade, C.orange),
          lineDataset('Ist (kum.)',  cumulative(dealsMetric.actual), C.orange,                   C.orange),
        ],
        yFmt: fmtInt,
      });
    } else {
      showNoData('s2-barDeals');
      showNoData('s2-lineCumDeals');
    }
  }

  // ─── Progress bars ──
  if (prefix === 's2') {
    renderStage2ProgressBars(`${prefix}-progressBars`, metrics);
  } else {
    const progressDefs = keys.map(key => ({
      id:    key,
      label: metrics[key].name || key,
      fmt:   guessFormat(metrics[key]),
    }));
    renderProgressBars(`${prefix}-progressBars`, metrics, progressDefs);
  }

  // ─── Insights ──
  renderInsights(`${prefix}-insights`, buildGenericInsights(metrics, tabLabel));
}

function findBestMetric(metrics) {
  // Prefer metric with both plan and actual data
  for (const m of Object.values(metrics)) {
    const hasP = m.plan.some(v => v !== null && v > 0);
    const hasA = m.actual.some(v => v !== null && v > 0);
    if (hasP && hasA) return m;
  }
  // Fallback: any metric with data
  for (const m of Object.values(metrics)) {
    if (m.plan.some(v => v !== null && v > 0) || m.actual.some(v => v !== null && v > 0)) return m;
  }
  return null;
}

function guessFormat(metric) {
  if (!metric) return fmtInt;
  const name = (metric.name || '').toLowerCase();
  if (name.includes('umsatz') || name.includes('revenue') || name.includes('€')) return fmtEurShort;
  const maxVal = Math.max(...(metric.plan || []).concat(metric.actual || []).filter(v => v !== null));
  if (maxVal > 1000) return fmtEurShort;
  if (maxVal <= 1 && maxVal > 0) return fmtPct;
  return fmtInt;
}

function buildGenericInsights(metrics, tabLabel) {
  const insights = [];
  for (const metric of Object.values(metrics)) {
    const ytdA = ytdSum(metric.actual);
    const ytdP = ytdSum(metric.plan);
    if (ytdA === null || ytdP === null || ytdP === 0) continue;
    const pct  = ytdA / ytdP * 100;
    const fmt  = guessFormat(metric);

    if (pct >= 100) {
      insights.push({ text: `<span class="tag-good">✓ ${metric.name}:</span> Plan übertroffen (${fmt(ytdA)} / ${fmt(ytdP)} = ${pct.toFixed(0)}%)` });
    } else if (pct >= 75) {
      insights.push({ text: `<span class="tag-warn">~ ${metric.name}:</span> ${pct.toFixed(0)}% des Plans erreicht` });
    } else {
      insights.push({ text: `<span class="tag-bad">✗ ${metric.name}:</span> Nur ${pct.toFixed(0)}% – ${fmt(ytdA)} von ${fmt(ytdP)}` });
    }
  }
  if (!insights.length) {
    insights.push({ text: `Daten für ${tabLabel} noch nicht verfügbar. Bitte Sheet-Namen in .env prüfen.` });
  }
  return insights;
}

function showTabNoData(prefix, tabLabel) {
  const ids = [`${prefix}-barMain`, `${prefix}-lineCum`];
  ids.forEach(id => showNoData(id));
  document.getElementById(`${prefix}-progressBars`)?.replaceWith(
    Object.assign(document.createElement('div'), {
      className: 'no-data',
      innerHTML: `<strong>Keine Daten geladen</strong>Sheet <em>${tabLabel}</em> nicht gefunden oder leer.<br>Überprüfe den Sheet-Namen in der <code>.env</code> Datei.`
    })
  );
}

// ══════════════════════════════════════════════════════════
// PROGRESS BARS
// ══════════════════════════════════════════════════════════
function renderProgressBars(containerId, metrics, defs) {
  const el = document.getElementById(containerId);
  if (!el) return;
  el.innerHTML = '';

  let any = false;
  for (const def of defs) {
    const metric = metrics[def.id];
    if (!metric) continue;

    const ytdA  = ytdSum(metric.actual) || 0;
    const annP  = metric.annualPlan;
    if (!annP || annP <= 0) continue;

    any = true;
    const pct       = Math.min((ytdA / annP) * 100, 100);
    const rawPct    = (ytdA / annP) * 100;
    const fillClass = rawPct >= 100 ? 'good' : rawPct >= 60 ? 'warn' : 'bad';
    const fmt       = def.fmt || fmtInt;

    el.innerHTML += `
      <div class="progress-item">
        <div class="progress-header">
          <span class="progress-label">${def.label}</span>
          <span class="progress-values">
            <strong>${fmt(ytdA)}</strong> / ${fmt(annP)} Jahresplan
          </span>
        </div>
        <div class="progress-track">
          <div class="progress-fill ${fillClass}" style="width:${pct.toFixed(1)}%"></div>
        </div>
        <div class="progress-pct">${rawPct.toFixed(1)}% des Jahresziels erreicht</div>
      </div>`;
  }

  if (!any) {
    el.innerHTML = '<p class="no-data">Jahresziele nicht verfügbar.</p>';
  }
}

// ══════════════════════════════════════════════════════════
// STAGE 2 SPECIFIC: Quarter-based progress bars
// Shows exactly 4 metrics vs current-quarter plan
// ══════════════════════════════════════════════════════════
function renderStage2ProgressBars(containerId, metrics) {
  const el = document.getElementById(containerId);
  if (!el) return;

  // Current quarter: 0=Q1, 1=Q2, 2=Q3, 3=Q4
  const now     = new Date();
  const qIdx    = Math.floor(now.getMonth() / 3); // 0-indexed quarter
  const qStart  = qIdx * 3;                        // first month index
  const qEnd    = qStart + 3;                      // exclusive
  const qLabel  = `Q${qIdx + 1} ${now.getFullYear()}`;

  function qSum(arr) {
    if (!arr) return null;
    let sum = 0, hasVal = false;
    for (let i = qStart; i < qEnd && i < arr.length; i++) {
      if (arr[i] !== null && arr[i] !== undefined) { sum += arr[i]; hasVal = true; }
    }
    return hasVal ? sum : null;
  }

  // The 4 required metrics and how to find them in the data
  // For Websessions, we want the "PLAN TOTAL" and "IST TOTAL" sub-rows
  // which land on the same key as "websessions_held" after our parser fix
  const DEFS = [
    {
      label:   'Deals Created',
      planKey: 'deals_created_stage_2',
      actKey:  'deals_created_stage_2',
      usePlan: true,
    },
    {
      label:   'Websessions held',
      planKey: 'websessions_held',
      actKey:  'websessions_held',
      usePlan: true,
    },
    {
      label:   'Offers Sent',
      planKey: 'offers_sent',
      actKey:  'offers_sent',
      usePlan: true,
    },
    {
      label:   'Deals Won',
      planKey: 'deals_won_stage_2',
      actKey:  'deals_won_stage_2',
      usePlan: true,
    },
  ];

  el.innerHTML = '';
  let any = false;

  for (const def of DEFS) {
    const planMetric = metrics[def.planKey];
    const actMetric  = metrics[def.actKey];

    const qIst  = actMetric  ? qSum(actMetric.actual)  : null;
    const qPlan = planMetric ? qSum(planMetric.plan)    : null;

    // If only actuals (no plan), show raw count
    const displayIst  = qIst  !== null ? qIst  : 0;
    const displayPlan = qPlan !== null && qPlan > 0 ? qPlan : null;

    // Skip entirely if no data at all
    if (qIst === null && qPlan === null) continue;
    any = true;

    const pct       = displayPlan ? Math.min((displayIst / displayPlan) * 100, 120) : null;
    const rawPct    = displayPlan ? (displayIst / displayPlan) * 100 : null;
    const fillClass = rawPct === null ? 'warn' : rawPct >= 100 ? 'good' : rawPct >= 60 ? 'warn' : 'bad';
    const fillWidth = pct !== null ? pct.toFixed(1) : '50';

    const planStr  = displayPlan !== null ? ` / ${fmtInt(displayPlan)} Plan` : '';
    const pctStr   = rawPct !== null ? `${rawPct.toFixed(1)}% des Quartalsziels (${qLabel})` : qLabel;

    el.innerHTML += `
      <div class="progress-item">
        <div class="progress-header">
          <span class="progress-label">${def.label}</span>
          <span class="progress-values">
            <strong>${fmtInt(displayIst)}</strong>${planStr}
          </span>
        </div>
        <div class="progress-track">
          <div class="progress-fill ${fillClass}" style="width:${fillWidth}%"></div>
        </div>
        <div class="progress-pct">${pctStr}</div>
      </div>`;
  }

  if (!any) {
    el.innerHTML = '<p class="no-data">Stage 2 Quartalsdaten nicht verfügbar.</p>';
  }
}

// ══════════════════════════════════════════════════════════
// INSIGHTS
// ══════════════════════════════════════════════════════════
function renderInsights(elId, insights) {
  const el = document.getElementById(elId);
  if (!el) return;
  el.innerHTML = insights.map(ins => `<li>${ins.text}</li>`).join('');
}

// ══════════════════════════════════════════════════════════
// CHART HELPERS
// ══════════════════════════════════════════════════════════
function destroyChart(id) {
  if (chartRegistry[id]) { chartRegistry[id].destroy(); delete chartRegistry[id]; }
}

function createBarChart(id, { labels, datasets, stacked = false, yLabel = '', yFmt = fmtInt, onClickIst = null }) {
  destroyChart(id);
  const canvas = document.getElementById(id);
  if (!canvas) return;
  const isInt = yFmt === fmtInt;
  chartRegistry[id] = new Chart(canvas.getContext('2d'), {
    type: 'bar',
    data: { labels, datasets },
    options: {
      responsive: true,
      plugins: {
        legend: { position: 'top', labels: { font: { size: 11 }, boxWidth: 12 } },
        tooltip: {
          callbacks: {
            label: ctx => ` ${ctx.dataset.label}: ${yFmt(ctx.parsed.y)}`,
          },
        },
      },
      scales: {
        x: { stacked, grid: { display: false }, ticks: { font: { size: 11 } } },
        y: {
          stacked,
          grid: { color: '#EBF3FC' },
          ticks: { font: { size: 11 }, precision: isInt ? 0 : undefined, callback: v => yFmt(v) },
        },
      },
      ...(onClickIst ? {
        onClick: (_evt, elements) => {
          if (!elements.length) return;
          const { datasetIndex, index } = elements[0];
          if (datasetIndex === 1) onClickIst(labels[index]);
        },
        onHover: (_evt, elements, chart) => {
          const el = elements[0];
          chart.canvas.style.cursor = (el && el.datasetIndex === 1) ? 'pointer' : 'default';
        },
      } : {}),
    },
  });
}

function createLineChart(id, { labels, datasets, yFmt = fmtInt }) {
  destroyChart(id);
  const canvas = document.getElementById(id);
  if (!canvas) return;
  const isInt = yFmt === fmtInt;
  chartRegistry[id] = new Chart(canvas.getContext('2d'), {
    type: 'line',
    data: { labels, datasets },
    options: {
      responsive: true,
      plugins: {
        legend: { position: 'top', labels: { font: { size: 11 }, boxWidth: 12 } },
        tooltip: {
          callbacks: { label: ctx => ` ${ctx.dataset.label}: ${yFmt(ctx.parsed.y)}` },
        },
      },
      scales: {
        x: { grid: { display: false }, ticks: { font: { size: 11 } } },
        y: {
          grid: { color: '#EBF3FC' },
          ticks: { font: { size: 11 }, precision: isInt ? 0 : undefined, callback: v => yFmt(v) },
        },
      },
    },
  });
}

function createDoughnutChart(id, { labels, data, colors, fmt = fmtInt }) {
  destroyChart(id);
  const canvas = document.getElementById(id);
  if (!canvas) return;
  chartRegistry[id] = new Chart(canvas.getContext('2d'), {
    type: 'doughnut',
    data: {
      labels,
      datasets: [{ data, backgroundColor: colors, borderWidth: 2, borderColor: '#fff' }],
    },
    options: {
      responsive: true,
      cutout: '65%',
      plugins: {
        legend: { position: 'bottom', labels: { font: { size: 11 }, boxWidth: 12, padding: 16 } },
        tooltip: {
          callbacks: {
            label: ctx => {
              const total = ctx.dataset.data.reduce((a, b) => a + (b || 0), 0);
              const pct   = total > 0 ? ((ctx.parsed / total) * 100).toFixed(1) : 0;
              return ` ${ctx.label}: ${fmt(ctx.parsed)} (${pct}%)`;
            },
          },
        },
      },
    },
  });
}

// Dataset builders
function barDataset(label, data, color, alpha = 0.85) {
  return {
    label,
    data: (data || []).map(v => v ?? null),
    backgroundColor: color + (alpha < 1 ? Math.round(alpha * 255).toString(16).padStart(2,'0') : ''),
    borderRadius: 4,
    borderSkipped: false,
  };
}

function lineDataset(label, data, borderColor, pointColor) {
  return {
    label,
    data: (data || []).map(v => v ?? null),
    borderColor,
    pointBackgroundColor: pointColor,
    backgroundColor: borderColor + '22',
    borderWidth: 2,
    pointRadius: 3,
    pointHoverRadius: 5,
    tension: 0.3,
    fill: false,
    spanGaps: true,
  };
}

// ══════════════════════════════════════════════════════════
// UTILITIES
// ══════════════════════════════════════════════════════════

// Sum actuals up to the current month (YTD)
function ytdSum(arr) {
  if (!arr || !arr.length) return null;
  const now = new Date();
  const currentMonth = now.getMonth(); // 0-indexed
  let sum = 0;
  let hasValue = false;
  for (let i = 0; i <= currentMonth && i < arr.length; i++) {
    if (arr[i] !== null && arr[i] !== undefined) {
      sum += arr[i];
      hasValue = true;
    }
  }
  return hasValue ? sum : null;
}

// Cumulative sum (null-safe)
function cumulative(arr) {
  if (!arr) return [];
  let running = 0;
  return arr.map(v => {
    if (v !== null && v !== undefined) running += v;
    return running || null;
  });
}

function setText(id, text) {
  const el = document.getElementById(id);
  if (el) el.textContent = text;
}

function showNoData(canvasId) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;
  const parent = canvas.parentElement;
  if (parent && !parent.querySelector('.no-data')) {
    canvas.style.display = 'none';
    const d = document.createElement('div');
    d.className = 'no-data';
    d.innerHTML = '<strong>Keine Daten</strong>Daten für dieses Segment noch nicht verfügbar.';
    parent.appendChild(d);
  }
}

// ─── Formatters ───────────────────────────────────────────
function fmtEur(val) {
  if (val === null || val === undefined) return '–';
  return new Intl.NumberFormat('de-DE', { style: 'currency', currency: 'EUR', maximumFractionDigits: 0 }).format(val);
}

function fmtEurShort(val) {
  if (val === null || val === undefined) return '–';
  const abs = Math.abs(val);
  if (abs >= 1_000_000) return `€${(val / 1_000_000).toFixed(2)}M`;
  if (abs >= 1_000)     return `€${(val / 1_000).toFixed(0)}K`;
  return `€${val.toFixed(0)}`;
}

function fmtInt(val) {
  if (val === null || val === undefined) return '–';
  return new Intl.NumberFormat('de-DE', { maximumFractionDigits: 0 }).format(val);
}

function fmtPct(val) {
  if (val === null || val === undefined) return '–';
  const v = Math.abs(val) <= 1 ? val * 100 : val;
  return `${v.toFixed(1)}%`;
}

// ─── Deal Modal ────────────────────────────────────────────────────────────────
function closeModal() {
  document.getElementById('dealModal').classList.add('hidden');
}
function handleModalOverlayClick(e) {
  if (e.target.id === 'dealModal') closeModal();
}
document.addEventListener('keydown', e => { if (e.key === 'Escape') closeModal(); });

// ─── HubSpot config (token fetched once from server) ──────────────────────────
let _hsToken = null;
async function getHubspotToken() {
  if (_hsToken) return _hsToken;
  const r = await fetch('/api/config');
  const d = await r.json();
  _hsToken = d.hubspotToken;
  return _hsToken;
}

// Convert MONTHS label + current year to Unix millisecond range
function monthLabelToUnixRange(label) {
  const idx = MONTHS.indexOf(label);
  if (idx < 0) return null;
  const year  = new Date().getFullYear();
  const start = Date.UTC(year, idx, 1);
  const end   = Date.UTC(year, idx + 1, 1) - 1;
  return { start: String(start), end: String(end) };
}

function fmtHubDate(dateStr) {
  if (!dateStr) return '–';
  const d = new Date(isNaN(dateStr) ? dateStr : Number(dateStr));
  return isNaN(d) ? dateStr : d.toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' });
}

function renderDealsTable(container, results, total, portalId) {
  if (!results || !results.length) {
    container.innerHTML = '<div class="modal-empty">Keine Deals für diesen Monat gefunden.</div>';
    return;
  }
  const deals = results.map(d => ({
    name:           d.properties.dealname || '(kein Name)',
    websessionDate: d.properties.datum_websession,
    amount:         d.properties.amount,
    stage:          d.properties.dealstage || '–',
    link:           portalId ? `https://app.hubspot.com/contacts/${portalId}/deal/${d.id}` : null,
  }));
  container.innerHTML = `
    <p class="modal-count">${total} Deal${total !== 1 ? 's' : ''} gefunden</p>
    <table class="deals-table">
      <thead><tr>
        <th>Deal-Name</th>
        <th>Websession</th>
        <th>Betrag</th>
        <th>Deal-Phase</th>
        <th></th>
      </tr></thead>
      <tbody>
        ${deals.map(d => `<tr>
          <td class="deal-name">${d.name}</td>
          <td>${fmtHubDate(d.websessionDate)}</td>
          <td class="deal-amount">${d.amount ? fmtEur(Number(d.amount)) : '–'}</td>
          <td><span class="deal-stage">${d.stage}</span></td>
          <td>${d.link ? `<a href="${d.link}" target="_blank" rel="noopener" class="deal-link">↗</a>` : ''}</td>
        </tr>`).join('')}
      </tbody>
    </table>`;
}

async function openWsDealsModal(monthLabel) {
  const modal = document.getElementById('dealModal');
  const body  = document.getElementById('modalBody');
  document.getElementById('modalTitle').textContent    = 'Stage 2 – Websession Deals';
  document.getElementById('modalSubtitle').textContent = monthLabel;
  body.innerHTML = '<div class="modal-loading"><div class="spinner"></div><p>Lade Deals…</p></div>';
  modal.classList.remove('hidden');

  try {
    const range = monthLabelToUnixRange(monthLabel);
    if (!range) throw new Error('Ungültiger Monat');
    console.log(`[WS Modal] ${monthLabel} → startMs: ${range.start}, endMs: ${range.end}`);

    const res  = await fetch(`/api/hubspot/deals?start=${range.start}&end=${range.end}&type=ws`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || `HTTP ${res.status}`);

    renderDealsTable(body, data.results || [], data.total ?? (data.results || []).length, data.portalId);
  } catch (e) {
    body.innerHTML = `<div class="modal-error">Fehler: ${e.message}</div>`;
  }
}

async function openOffersDealsModal(monthLabel) {
  const modal = document.getElementById('dealModal');
  const body  = document.getElementById('modalBody');
  document.getElementById('modalTitle').textContent    = 'Stage 2 – Angebote versandt';
  document.getElementById('modalSubtitle').textContent = monthLabel;
  body.innerHTML = '<div class="modal-loading"><div class="spinner"></div><p>Lade Deals…</p></div>';
  modal.classList.remove('hidden');

  try {
    const range = monthLabelToUnixRange(monthLabel);
    if (!range) throw new Error('Ungültiger Monat');
    console.log(`[Offers Modal] ${monthLabel} → startMs: ${range.start}, endMs: ${range.end}`);

    const res  = await fetch(`/api/hubspot/deals?start=${range.start}&end=${range.end}&type=offers`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || `HTTP ${res.status}`);

    renderDealsTable(body, data.results || [], data.total ?? (data.results || []).length, data.portalId);
  } catch (e) {
    body.innerHTML = `<div class="modal-error">Fehler: ${e.message}</div>`;
  }
}

async function openDealsWonModal(monthLabel) {
  const modal = document.getElementById('dealModal');
  const body  = document.getElementById('modalBody');
  document.getElementById('modalTitle').textContent    = 'Stage 2 – Deals Won';
  document.getElementById('modalSubtitle').textContent = monthLabel;
  body.innerHTML = '<div class="modal-loading"><div class="spinner"></div><p>Lade Deals…</p></div>';
  modal.classList.remove('hidden');

  try {
    const range = monthLabelToUnixRange(monthLabel);
    if (!range) throw new Error('Ungültiger Monat');
    console.log(`[Deals Won Modal] ${monthLabel} → startMs: ${range.start}, endMs: ${range.end}`);

    const res  = await fetch(`/api/hubspot/deals?start=${range.start}&end=${range.end}&type=won`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || `HTTP ${res.status}`);

    renderDealsTable(body, data.results || [], data.total ?? (data.results || []).length, data.portalId);
  } catch (e) {
    body.innerHTML = `<div class="modal-error">Fehler: ${e.message}</div>`;
  }
}

async function openS3DirectModal(monthLabel) {
  const modal = document.getElementById('dealModal');
  const body  = document.getElementById('modalBody');
  document.getElementById('modalTitle').textContent    = 'Stage 3 – Deals Won Direct (NL)';
  document.getElementById('modalSubtitle').textContent = monthLabel;
  body.innerHTML = '<div class="modal-loading"><div class="spinner"></div><p>Lade Deals…</p></div>';
  modal.classList.remove('hidden');

  try {
    const range = monthLabelToUnixRange(monthLabel);
    if (!range) throw new Error('Ungültiger Monat');
    console.log(`[S3 Direct Modal] ${monthLabel} → startMs: ${range.start}, endMs: ${range.end}`);

    const res  = await fetch(`/api/hubspot/deals?start=${range.start}&end=${range.end}&type=s3-direct`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || `HTTP ${res.status}`);

    renderDealsTable(body, data.results || [], data.total ?? (data.results || []).length, data.portalId);
  } catch (e) {
    body.innerHTML = `<div class="modal-error">Fehler: ${e.message}</div>`;
  }
}

async function openS3Phase1Modal(monthLabel) {
  const modal = document.getElementById('dealModal');
  const body  = document.getElementById('modalBody');
  document.getElementById('modalTitle').textContent    = 'Stage 3 – Deals Won Phase 1 (JH)';
  document.getElementById('modalSubtitle').textContent = monthLabel;
  body.innerHTML = '<div class="modal-loading"><div class="spinner"></div><p>Lade Deals…</p></div>';
  modal.classList.remove('hidden');

  try {
    const range = monthLabelToUnixRange(monthLabel);
    if (!range) throw new Error('Ungültiger Monat');
    console.log(`[S3 Phase1 Modal] ${monthLabel} → startMs: ${range.start}, endMs: ${range.end}`);

    const res  = await fetch(`/api/hubspot/deals?start=${range.start}&end=${range.end}&type=s3-phase1`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || `HTTP ${res.status}`);

    renderDealsTable(body, data.results || [], data.total ?? (data.results || []).length, data.portalId);
  } catch (e) {
    body.innerHTML = `<div class="modal-error">Fehler: ${e.message}</div>`;
  }
}
