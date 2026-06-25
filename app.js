/**
 * PLAN DE DESARROLLO GACHANCIPÁ 2024-2027
 * Dashboard Ejecutivo — app.js
 * Arquitectura: filtrado reactivo + Plotly.js + DataTables
 */

'use strict';

// ══════════════════════════════════════════════════════════════
// CONSTANTES Y CONFIGURACIÓN
// ══════════════════════════════════════════════════════════════
const COLORS = {
  blueDark : '#0B3C5D',
  blueMid  : '#1A6FB5',
  blueLight: '#E8F1FA',
  green    : '#2E8B57',
  gold     : '#D4AF37',
  red      : '#C0392B',
  orange   : '#E67E22',
  purple   : '#7D3C98',
  teal     : '#1B8A6B',
  grayMuted: '#8A93A6',
};

const SEC_ABBR = {
  'Secretaría de Gobierno':                                   'Gobierno',
  'Secretaría de Planeación y Servicios Públicos':            'Planeación',
  'Secretaría de Educación, cultura, recreación y deportes.': 'Educación',
  'Secretaría de Desarrollo Social':                          'Des. Social',
  'Secretaría de Desarrollo Económico, Agropecuario y Ambiente': 'Des. Económico',
  'Secretaria de Obras Públicas':                             'Obras Púb.',
  'Secretaría de Hacienda':                                   'Hacienda',
  'Secretaría de Desarrollo Institucional':                   'Des. Inst.',
  'Secretaría General':                                       'Sec. General',
};

const SEC_COLORS = [
  '#0B3C5D','#1A6FB5','#2E8B57','#D4AF37','#C0392B',
  '#7D3C98','#1B8A6B','#E67E22','#2980B9',
];

// ══════════════════════════════════════════════════════════════
// ESTADO GLOBAL
// ══════════════════════════════════════════════════════════════
let dataTable = null;
let filteredData = [];
let activeAnio = '2024';

// ══════════════════════════════════════════════════════════════
// UTILIDADES
// ══════════════════════════════════════════════════════════════
const $ = id => document.getElementById(id);
const fmt1 = n => n.toFixed(1);
const fmtPct = n => n.toFixed(1) + '%';
const isDark = () => document.documentElement.dataset.theme === 'dark';

function plotlyLayout(extra = {}) {
  const dark = isDark();
  return {
    paper_bgcolor: 'transparent',
    plot_bgcolor : dark ? '#182234' : '#F5F7FA',
    font: { family: 'Segoe UI, system-ui, sans-serif', color: dark ? '#C9D6E8' : '#4A5568', size: 11 },
    margin: { l: 10, r: 10, t: 10, b: 10 },
    ...extra,
  };
}

function getAvanceForYear(rec, anio) {
  const map = { '2024': rec.avance_2024, '2025': rec.avance_2025, '2026': rec.avance_2026, '2027': rec.avance_2027 };
  return map[anio] ?? 0;
}
function getProgForYear(rec, anio) {
  return anio === '2024' ? rec.programado_2024 : rec.programado_2025;
}
function getMetaForYear(rec, anio) {
  return anio === '2024' ? rec.meta_2024 : rec.meta_2025;
}
function getPlaneadoForYear(rec, anio) {
  return anio === '2024' ? rec.total_planeado_2024 : rec.total_planeado_2025;
}
function getEjecutadoForYear(rec, anio) {
  return anio === '2024' ? rec.total_ejecutado_2024 : rec.total_ejecutado_2025;
}

function getQuarterData(rec, anio, trimestre) {
  const tMap = {
    '2024': { T1: [rec.t1_plan_2024, rec.t1_ejec_2024], T2: [rec.t2_plan_2024, rec.t2_ejec_2024], T3: [rec.t3_plan_2024, rec.t3_ejec_2024], T4: [rec.t4_plan_2024, rec.t4_ejec_2024] },
    '2025': { T1: [rec.t1_plan_2025, rec.t1_ejec_2025], T2: [rec.t2_plan_2025, rec.t2_ejec_2025], T3: [rec.t3_plan_2025, rec.t3_ejec_2025], T4: [0,0] },
  };
  return tMap[anio]?.[trimestre] ?? [0, 0];
}

function statusLabel(avance) {
  if (avance >= 90) return { label: 'Cumplida', cls: 'badge-cumplida' };
  if (avance >  0)  return { label: 'En Ejecución', cls: 'badge-ejecucion' };
  return { label: 'Sin Iniciar', cls: 'badge-sininicio' };
}

// ══════════════════════════════════════════════════════════════
// INICIALIZACIÓN DE FILTROS
// ══════════════════════════════════════════════════════════════
function initFilters() {
  const secSel = $('fSecretaria');
  const ejeSel = $('fEje');

  const secs = [...new Set(PLAN_DATA.map(r => r.secretaria))].sort();
  secs.forEach(s => {
    const o = document.createElement('option');
    o.value = s; o.textContent = s;
    secSel.appendChild(o);
  });

  const ejeMap = {};
  PLAN_DATA.forEach(r => { if (r.eje > 0) ejeMap[r.eje] = r.eje_estrategico.replace(/\n/g,' ').substring(0, 55); });
  Object.keys(ejeMap).sort((a,b)=>a-b).forEach(k => {
    const o = document.createElement('option');
    o.value = k; o.textContent = `Eje ${k} – ${ejeMap[k]}`;
    ejeSel.appendChild(o);
  });

  ['fAnio','fTrimestre','fSecretaria','fEstado','fEje'].forEach(id => {
    $(id).addEventListener('change', () => { activeAnio = $('fAnio').value === 'todos' ? '2024' : $('fAnio').value; applyFilters(); });
  });

  $('resetFilters').addEventListener('click', resetFilters);
  $('refreshBtn').addEventListener('click', () => applyFilters());
  $('darkToggle').addEventListener('click', toggleTheme);
  $('exportPdfBtn').addEventListener('click', exportPDF);
  $('exportExcelBtn').addEventListener('click', exportExcelFull);
  $('tblExportCsv').addEventListener('click', exportTableCSV);
  $('tblExportExcel').addEventListener('click', exportTableExcel);
  $('tblExportPdf').addEventListener('click', exportTablePDF);
}

function resetFilters() {
  $('fAnio').value = '2024'; $('fTrimestre').value = 'todos';
  $('fSecretaria').value = 'todas'; $('fEstado').value = 'todos';
  $('fEje').value = 'todos';
  activeAnio = '2024';
  applyFilters();
}

// ══════════════════════════════════════════════════════════════
// FILTRADO DE DATOS
// ══════════════════════════════════════════════════════════════
function applyFilters() {
  const anio      = $('fAnio').value;
  const trimestre = $('fTrimestre').value;
  const secretaria= $('fSecretaria').value;
  const estado    = $('fEstado').value;
  const eje       = $('fEje').value;

  // Determine active year for avance display
  activeAnio = (anio === 'todos') ? '2024' : anio;

  filteredData = PLAN_DATA.filter(r => {
    if (secretaria !== 'todas' && r.secretaria !== secretaria) return false;
    if (eje !== 'todos' && String(r.eje) !== String(eje)) return false;

    // For year filter, check if we have data for that year
    if (anio !== 'todos') {
      const progKey = anio === '2024' ? r.programado_2024 : r.programado_2025;
      if (estado !== 'todos' && progKey !== estado) return false;
    } else {
      // Check any year's estado
      if (estado !== 'todos') {
        const matchAny = [r.programado_2024, r.programado_2025].includes(estado);
        if (!matchAny) return false;
      }
    }

    // Trimestre filter: only meaningful for concrete year
    if (trimestre !== 'todos' && anio !== 'todos') {
      const [plan, ejec] = getQuarterData(r, anio, trimestre);
      if (plan === 0 && ejec === 0) return false;
    }

    return true;
  });

  updateKPIs();
  renderGauge();
  renderDonut();
  renderBarrasH();
  renderLinea();
  renderApilado();
  renderTreemap();
  renderHeatmap();
  updateTable();
  updateBadges();
}

// ══════════════════════════════════════════════════════════════
// KPIs
// ══════════════════════════════════════════════════════════════
function updateKPIs() {
  const data = filteredData;
  const total = data.length;
  const prog = data.filter(r => getProgForYear(r, activeAnio) === 'P').length;
  const noProg = data.filter(r => getProgForYear(r, activeAnio) === 'NP').length;
  const avances = data.map(r => getAvanceForYear(r, activeAnio));
  const cumplidas = avances.filter(a => a >= 90).length;
  const sinInicio = avances.filter(a => a === 0).length;
  const enEjecucion = avances.filter(a => a > 0 && a < 90).length;
  const avgAvance = avances.length ? avances.reduce((a,b) => a+b, 0) / avances.length : 0;
  const secs = new Set(data.map(r => r.secretaria)).size;

  setKPI('kTotal', total, null);
  setKPI('kProg', prog, `${((prog/total)*100).toFixed(0)}% del total`);
  setKPI('kNoProg', noProg, `${((noProg/total)*100).toFixed(0)}% sin programar`);
  setKPI('kCumplidas', cumplidas, `${((cumplidas/total)*100).toFixed(0)}% cumplimiento`);
  setKPI('kEjecucion', enEjecucion, `${((enEjecucion/total)*100).toFixed(0)}% en proceso`);
  setKPI('kAvance', fmtPct(avgAvance), activeAnio);
  setKPI('kSinInicio', sinInicio, `${((sinInicio/total)*100).toFixed(0)}% no iniciadas`);
  setKPI('kSecretarias', secs, 'dependencias');
}

function setKPI(id, val, delta) {
  const el = $(id);
  if (el) { el.textContent = val; el.style.opacity=0; requestAnimationFrame(()=>{ el.style.transition='opacity .4s'; el.style.opacity=1; }); }
  const dEl = $(id+'Delta');
  if (dEl && delta !== null) dEl.textContent = delta;
}

// ══════════════════════════════════════════════════════════════
// GAUGE CHART
// ══════════════════════════════════════════════════════════════
function renderGauge() {
  const avances = filteredData.map(r => getAvanceForYear(r, activeAnio));
  const avg = avances.length ? avances.reduce((a,b)=>a+b,0)/avances.length : 0;

  const trace = {
    type: 'indicator', mode: 'gauge+number+delta',
    value: parseFloat(avg.toFixed(1)),
    delta: { reference: 50, valueformat: '.1f', suffix:'%' },
    number: { suffix: '%', font: { size: 32, color: COLORS.blueDark } },
    gauge: {
      axis: { range: [0,100], tickcolor: '#8A93A6', tickwidth: 1, tickfont:{size:10} },
      bar: { color: avg >= 70 ? COLORS.green : avg >= 40 ? COLORS.gold : COLORS.red, thickness: .6 },
      bgcolor: isDark() ? '#182234' : '#F5F7FA',
      borderwidth: 0,
      steps: [
        { range:[0,30],  color: isDark() ? 'rgba(192,57,43,.15)' : '#FDE8E8' },
        { range:[30,70], color: isDark() ? 'rgba(212,175,55,.1)' : '#FBF5DC' },
        { range:[70,100],color: isDark() ? 'rgba(46,139,87,.15)' : '#E6F4EC' },
      ],
      threshold: { line:{ color:COLORS.gold, width:3 }, thickness:.75, value:70 },
    },
  };
  Plotly.newPlot('chartGauge', [trace], plotlyLayout({ height:220, margin:{l:20,r:20,t:20,b:10} }), { responsive:true, displayModeBar:false });
}

// ══════════════════════════════════════════════════════════════
// DONUT CHART
// ══════════════════════════════════════════════════════════════
function renderDonut() {
  const avances = filteredData.map(r => getAvanceForYear(r, activeAnio));
  const cumplidas = avances.filter(a => a >= 90).length;
  const enEjecucion = avances.filter(a => a > 0 && a < 90).length;
  const sinInicio = avances.filter(a => a === 0).length;
  const total = filteredData.length;
  const retrasadas = filteredData.filter(r => getProgForYear(r, activeAnio) === 'P' && getAvanceForYear(r, activeAnio) < 30).length;

  const trace = {
    type: 'pie', hole: .55,
    values: [cumplidas, enEjecucion, sinInicio],
    labels: ['Cumplidas ≥90%', 'En Ejecución', 'Sin Iniciar'],
    marker: { colors: [COLORS.green, COLORS.blueMid, COLORS.red], line:{ color:isDark()?'#182234':'#fff', width:2 } },
    textinfo: 'percent',
    hovertemplate: '<b>%{label}</b><br>%{value} metas (%{percent})<extra></extra>',
    textfont: { size: 11 },
  };
  Plotly.newPlot('chartDonut', [trace], plotlyLayout({
    height:220, showlegend:true,
    legend: { orientation:'h', y:-0.2, font:{size:10} },
    margin:{l:10,r:10,t:10,b:40},
    annotations: [{ text: `<b>${total}</b><br>Metas`, x:.5, y:.5, xref:'paper', yref:'paper', showarrow:false, font:{size:13, color: isDark()?'#C9D6E8':COLORS.blueDark} }],
  }), { responsive:true, displayModeBar:false });
}

// ══════════════════════════════════════════════════════════════
// BARRAS HORIZONTALES — AVANCE POR SECRETARÍA
// ══════════════════════════════════════════════════════════════
function renderBarrasH() {
  const secMap = {};
  filteredData.forEach(r => {
    const s = SEC_ABBR[r.secretaria] || r.secretaria;
    if (!secMap[s]) secMap[s] = [];
    secMap[s].push(getAvanceForYear(r, activeAnio));
  });

  const labels = [], values = [], colors = [];
  Object.entries(secMap).sort((a,b)=>{
    const avgA = a[1].reduce((x,y)=>x+y,0)/a[1].length;
    const avgB = b[1].reduce((x,y)=>x+y,0)/b[1].length;
    return avgA - avgB;
  }).forEach(([sec, avs], i) => {
    const avg = avs.reduce((a,b)=>a+b,0)/avs.length;
    labels.push(sec);
    values.push(parseFloat(avg.toFixed(1)));
    colors.push(avg>=70 ? COLORS.green : avg>=40 ? COLORS.gold : COLORS.red);
  });

  const trace = {
    type: 'bar', orientation:'h',
    x: values, y: labels,
    marker: { color: colors, line:{ color:'transparent' } },
    text: values.map(v => v.toFixed(1)+'%'),
    textposition: 'outside', cliponaxis: false,
    hovertemplate: '<b>%{y}</b><br>Avance: %{x:.1f}%<extra></extra>',
  };
  Plotly.newPlot('chartBarrasH', [trace], plotlyLayout({
    height: Math.max(220, labels.length * 40),
    xaxis: { range:[0,110], ticksuffix:'%', gridcolor: isDark()?'#253550':'#DDE3EC', showgrid:true },
    yaxis: { automargin:true, tickfont:{size:11} },
    margin: { l:10, r:60, t:10, b:30 },
    shapes: [{ type:'line', x0:70, x1:70, y0:-0.5, y1:labels.length-0.5, line:{color:COLORS.gold,width:2,dash:'dot'} }],
  }), { responsive:true, displayModeBar:false });
}

// ══════════════════════════════════════════════════════════════
// LÍNEA TEMPORAL — EVOLUCIÓN TRIMESTRAL
// ══════════════════════════════════════════════════════════════
function renderLinea() {
  const years = ['2024','2025'];
  const trimestres = ['T1','T2','T3','T4'];
  const xLabels = ['T1-24','T2-24','T3-24','T4-24','T1-25','T2-25','T3-25','T4-25'];

  // Compute average avance per trimestre per año from all data
  const getAvgTrimestre = (anio, trim) => {
    const data = PLAN_DATA.filter(r => {
      const secF = $('fSecretaria').value;
      const ejeF = $('fEje').value;
      if (secF !== 'todas' && r.secretaria !== secF) return false;
      if (ejeF !== 'todos' && String(r.eje) !== String(ejeF)) return false;
      return true;
    });
    const vals = data.map(r => {
      const [plan, ejec] = getQuarterData(r, anio, trim);
      if (plan === 0) return null;
      return Math.min((ejec/plan)*100, 100);
    }).filter(v => v !== null);
    return vals.length ? vals.reduce((a,b)=>a+b,0)/vals.length : 0;
  };

  const yValues = years.flatMap(a => trimestres.map(t => parseFloat(getAvgTrimestre(a,t).toFixed(1))));

  const trace = {
    type:'scatter', mode:'lines+markers',
    x: xLabels, y: yValues,
    line: { color: COLORS.blueMid, width:3, shape:'spline' },
    marker: { size:8, color: yValues.map(v => v>=70?COLORS.green:v>0?COLORS.gold:COLORS.red), line:{color:'#fff',width:2} },
    fill:'tozeroy', fillcolor: isDark() ? 'rgba(26,111,181,.12)' : 'rgba(26,111,181,.08)',
    hovertemplate: '<b>%{x}</b><br>Avance promedio: %{y:.1f}%<extra></extra>',
  };

  const refLine = { type:'scatter', mode:'lines', x:xLabels, y:xLabels.map(()=>70),
    line:{color:COLORS.gold,width:2,dash:'dot'}, name:'Meta 70%', hoverinfo:'none', showlegend:true };

  Plotly.newPlot('chartLinea', [trace, refLine], plotlyLayout({
    height:260,
    xaxis: { gridcolor: isDark()?'#253550':'#DDE3EC', showgrid:true },
    yaxis: { range:[0,105], ticksuffix:'%', gridcolor: isDark()?'#253550':'#DDE3EC', showgrid:true },
    margin:{l:40,r:20,t:10,b:40},
    showlegend:true,
    legend:{ orientation:'h', y:-0.25, font:{size:10} },
  }), { responsive:true, displayModeBar:false });
}

// ══════════════════════════════════════════════════════════════
// BARRAS APILADAS — PROGRAMADAS VS NO PROGRAMADAS
// ══════════════════════════════════════════════════════════════
function renderApilado() {
  const secMap = {};
  filteredData.forEach(r => {
    const s = SEC_ABBR[r.secretaria] || r.secretaria;
    if (!secMap[s]) secMap[s] = { P:0, NP:0 };
    const prog = getProgForYear(r, activeAnio);
    if (prog === 'P') secMap[s].P++;
    else secMap[s].NP++;
  });

  const labels = Object.keys(secMap);
  const trP = {
    type:'bar', name:'Programadas',
    x: labels, y: labels.map(s => secMap[s].P),
    marker:{ color: COLORS.green },
    hovertemplate:'<b>%{x}</b><br>Programadas: %{y}<extra></extra>',
  };
  const trNP = {
    type:'bar', name:'No Programadas',
    x: labels, y: labels.map(s => secMap[s].NP),
    marker:{ color: COLORS.gold },
    hovertemplate:'<b>%{x}</b><br>No Programadas: %{y}<extra></extra>',
  };

  Plotly.newPlot('chartApilado', [trP, trNP], plotlyLayout({
    barmode:'stack', height:260,
    xaxis: { tickangle:-25, automargin:true, tickfont:{size:10} },
    yaxis: { gridcolor: isDark()?'#253550':'#DDE3EC' },
    margin:{ l:40, r:10, t:10, b:80 },
    showlegend:true,
    legend:{ orientation:'h', y:-0.45, font:{size:10} },
  }), { responsive:true, displayModeBar:false });
}

// ══════════════════════════════════════════════════════════════
// TREEMAP — POR EJE ESTRATÉGICO Y SECRETARÍA
// ══════════════════════════════════════════════════════════════
function renderTreemap() {
  const ejeLabel = n => {
    const labels = {1:'Seguridad',2:'Des. Social',3:'Deporte',4:'Ambiente',5:'Prosperidad',6:'Planificación',7:'Agropecuario',8:'Gestión Riesgo',9:'Transparencia'};
    return labels[n] || `Eje ${n}`;
  };

  const ids=[], labels=[], parents=[], values=[], colors=[];
  ids.push('root'); labels.push('Gachancipá'); parents.push(''); values.push(0); colors.push(COLORS.blueDark);

  const ejeMap = {};
  filteredData.forEach(r => {
    if (!ejeMap[r.eje]) ejeMap[r.eje] = {};
    const s = SEC_ABBR[r.secretaria] || r.secretaria;
    ejeMap[r.eje][s] = (ejeMap[r.eje][s] || 0) + 1;
  });

  let ei = 0;
  Object.entries(ejeMap).sort((a,b)=>a[0]-b[0]).forEach(([eje, secs]) => {
    const ejeId = `eje_${eje}`;
    const total = Object.values(secs).reduce((a,b)=>a+b,0);
    ids.push(ejeId); labels.push(`Eje ${eje}: ${ejeLabel(+eje)}`);
    parents.push('root'); values.push(total); colors.push(SEC_COLORS[ei % SEC_COLORS.length]);
    Object.entries(secs).forEach(([sec, cnt]) => {
      ids.push(`${ejeId}_${sec}`); labels.push(sec);
      parents.push(ejeId); values.push(cnt); colors.push(SEC_COLORS[ei % SEC_COLORS.length]+'CC');
    });
    ei++;
  });

  const trace = {
    type:'treemap', ids, labels, parents, values,
    marker:{ colors, colorscale:null, line:{width:.5, color:'#fff'} },
    textinfo:'label+value',
    hovertemplate:'<b>%{label}</b><br>%{value} metas<extra></extra>',
    textfont:{ size:11 },
  };
  Plotly.newPlot('chartTreemap',[trace],plotlyLayout({height:300,margin:{l:5,r:5,t:5,b:5}}),{responsive:true,displayModeBar:false});
}

// ══════════════════════════════════════════════════════════════
// HEATMAP — CUMPLIMIENTO POR SECRETARÍA Y TRIMESTRE
// ══════════════════════════════════════════════════════════════
function renderHeatmap() {
  const secs = [...new Set(filteredData.map(r => SEC_ABBR[r.secretaria] || r.secretaria))].sort();
  const xLabels = ['T1-24','T2-24','T3-24','T4-24','T1-25','T2-25','T3-25','T4-25'];
  const combos = [['2024','T1'],['2024','T2'],['2024','T3'],['2024','T4'],['2025','T1'],['2025','T2'],['2025','T3'],['2025','T4']];

  const z = secs.map(sec => {
    return combos.map(([anio, trim]) => {
      const recs = filteredData.filter(r => (SEC_ABBR[r.secretaria]||r.secretaria) === sec);
      const vals = recs.map(r => {
        const [plan, ejec] = getQuarterData(r, anio, trim);
        if (plan === 0) return null;
        return Math.min((ejec/plan)*100, 100);
      }).filter(v => v !== null);
      return vals.length ? parseFloat((vals.reduce((a,b)=>a+b,0)/vals.length).toFixed(1)) : null;
    });
  });

  const trace = {
    type:'heatmap', z, x:xLabels, y:secs,
    colorscale:[[0,'#C0392B'],[0.4,'#D4AF37'],[0.7,'#F4D03F'],[1,'#2E8B57']],
    zmin:0, zmax:100,
    text: z.map(row => row.map(v => v !== null ? v.toFixed(1)+'%' : 'N/D')),
    hovertemplate:'<b>%{y} — %{x}</b><br>Avance: %{text}<extra></extra>',
    showscale:true,
    colorbar:{ thickness:12, len:.8, ticksuffix:'%', outlinecolor:'transparent', tickfont:{size:10} },
  };
  Plotly.newPlot('chartHeatmap',[trace],plotlyLayout({
    height:300,
    xaxis:{ side:'bottom', automargin:true, tickfont:{size:10} },
    yaxis:{ automargin:true, tickfont:{size:10} },
    margin:{l:10,r:60,t:10,b:40},
  }),{responsive:true,displayModeBar:false});
}

// ══════════════════════════════════════════════════════════════
// TABLA EJECUTIVA — DATATABLES
// ══════════════════════════════════════════════════════════════
function buildTableRows(data) {
  return data.map(r => {
    const av = getAvanceForYear(r, activeAnio);
    const prog = getProgForYear(r, activeAnio);
    const meta = getMetaForYear(r, activeAnio);
    const plan = getPlaneadoForYear(r, activeAnio);
    const ejec = getEjecutadoForYear(r, activeAnio);
    const st = statusLabel(av);
    const pctColor = av>=70?'#2E8B57':av>=30?'#D4AF37':'#C0392B';
    const progBadge = prog==='P'
      ? `<span class="badge-prog">Programada</span>`
      : `<span class="badge-noprog">No Programada</span>`;
    const stBadge = `<span class="${st.cls}">${st.label}</span>`;
    const bar = `<div>${av.toFixed(1)}%<div class="progress-mini"><div class="progress-mini-fill" style="width:${av}%;background:${pctColor}"></div></div></div>`;

    return [
      `<span class="fw-bold text-muted">${r.eje}</span>`,
      `<small>${r.secretaria.replace('Secretaría de ','Sec. ')}</small>`,
      `<small>${(r.programa||'').replace(/\n/g,' ').substring(0,60)}</small>`,
      `<small>${(r.producto||'').substring(0,55)}</small>`,
      `<small>${(meta||'').replace(/\n/g,' ').substring(0,65)}</small>`,
      activeAnio,
      stBadge,
      fmt1(plan),
      fmt1(ejec),
      bar,
      progBadge,
      `<small style="color:var(--gray-muted)">${(r.logros||'').substring(0,80)}</small>`,
    ];
  });
}

function updateTable() {
  if (dataTable) {
    dataTable.clear();
    dataTable.rows.add(buildTableRows(filteredData));
    dataTable.draw();
    return;
  }

  const tbody = $('tablaBody');
  buildTableRows(filteredData).forEach(row => {
    const tr = document.createElement('tr');
    row.forEach(cell => {
      const td = document.createElement('td');
      td.innerHTML = cell;
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  dataTable = $('#tablaEjecutiva').DataTable({
    language: {
      url: null,
      search:'Buscar:', lengthMenu:'Mostrar _MENU_ filas',
      info:'Mostrando _START_–_END_ de _TOTAL_ metas',
      paginate:{ previous:'‹', next:'›' },
      zeroRecords:'Sin resultados', infoEmpty:'', infoFiltered:'(filtrado de _MAX_)',
    },
    pageLength: 15,
    order:[[9,'desc']],
    scrollX: false,
    columnDefs:[
      { targets:[4,5,11], orderable:false },
      { targets:[7,8], className:'text-end' },
    ],
  });
}

// ══════════════════════════════════════════════════════════════
// ACTUALIZAR BADGES DE VIGENCIA
// ══════════════════════════════════════════════════════════════
function updateBadges() {
  ['gaugeBadge','donutBadge','barrasHBadge'].forEach(id => {
    const el = $(id);
    if (el) el.textContent = activeAnio;
  });
}

// ══════════════════════════════════════════════════════════════
// TEMA OSCURO
// ══════════════════════════════════════════════════════════════
function toggleTheme() {
  const html = document.documentElement;
  const isDarkNow = html.dataset.theme === 'dark';
  html.dataset.theme = isDarkNow ? 'light' : 'dark';
  const btn = $('darkToggle');
  btn.innerHTML = isDarkNow
    ? '<i class="bi bi-moon-stars-fill"></i>'
    : '<i class="bi bi-sun-fill"></i>';
  // Re-render all charts
  renderGauge(); renderDonut(); renderBarrasH();
  renderLinea(); renderApilado(); renderTreemap(); renderHeatmap();
}

// ══════════════════════════════════════════════════════════════
// EXPORTACIONES
// ══════════════════════════════════════════════════════════════
async function exportPDF() {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation:'landscape', unit:'mm', format:'a4' });
  const W = doc.internal.pageSize.getWidth();

  // Portada
  doc.setFillColor(11,60,93);
  doc.rect(0,0,W,40,'F');
  doc.setTextColor(212,175,55);
  doc.setFontSize(16);
  doc.text('MUNICIPIO DE GACHANCIPÁ — CUNDINAMARCA', W/2, 15, {align:'center'});
  doc.setTextColor(255,255,255);
  doc.setFontSize(12);
  doc.text('Plan de Desarrollo: MI COMPROMISO ES GACHANCIPÁ 2024–2027', W/2, 25, {align:'center'});
  doc.setFontSize(9);
  doc.text(`Informe de Seguimiento al Plan Indicativo · Vigencia ${activeAnio} · Generado: ${new Date().toLocaleDateString('es-CO')}`, W/2, 33, {align:'center'});

  // KPIs
  doc.setTextColor(11,60,93);
  doc.setFontSize(11);
  doc.text('INDICADORES CLAVE DE DESEMPEÑO', 14, 52);
  doc.setDrawColor(212,175,55); doc.setLineWidth(.5); doc.line(14,54,W-14,54);

  const kpis = [
    ['Total Metas', $('kTotal').textContent],
    ['Programadas', $('kProg').textContent],
    ['No Programadas', $('kNoProg').textContent],
    ['Cumplidas ≥90%', $('kCumplidas').textContent],
    ['En Ejecución', $('kEjecucion').textContent],
    ['Avance General', $('kAvance').textContent],
    ['Sin Iniciar', $('kSinInicio').textContent],
    ['Secretarías', $('kSecretarias').textContent],
  ];
  doc.setFontSize(9);
  kpis.forEach(([label,val],i) => {
    const x = 14 + (i%4)*68, y = 65 + Math.floor(i/4)*16;
    doc.setFillColor(245,247,250);
    doc.roundedRect(x,y-6,64,14,2,2,'F');
    doc.setTextColor(138,147,166); doc.text(label, x+3, y);
    doc.setTextColor(11,60,93); doc.setFontSize(13);
    doc.text(String(val), x+3, y+7);
    doc.setFontSize(9);
  });

  // Capturas gráficos
  const charts = ['chartGauge','chartDonut','chartBarrasH','chartLinea'];
  doc.addPage();
  doc.setFillColor(11,60,93); doc.rect(0,0,W,12,'F');
  doc.setTextColor(212,175,55); doc.setFontSize(11);
  doc.text('VISUALIZACIONES ESTRATÉGICAS', W/2, 8, {align:'center'});

  let yPos = 20;
  for (const id of charts) {
    const el = document.getElementById(id);
    if (!el) continue;
    try {
      const canvas = await html2canvas(el, {backgroundColor:null,scale:1.5});
      const imgData = canvas.toDataURL('image/png');
      const maxW = (W-28)/2, maxH = 60;
      const xPos = charts.indexOf(id) % 2 === 0 ? 14 : 14+maxW+4;
      if (charts.indexOf(id) % 2 === 0 && charts.indexOf(id) > 0) yPos += maxH+6;
      doc.addImage(imgData,'PNG',xPos,yPos,maxW,maxH);
    } catch(e){}
  }

  // Footer
  const pageCount = doc.getNumberOfPages();
  for (let i=1;i<=pageCount;i++) {
    doc.setPage(i);
    doc.setFillColor(11,60,93);
    doc.rect(0,doc.internal.pageSize.getHeight()-10,W,10,'F');
    doc.setTextColor(255,255,255); doc.setFontSize(7);
    doc.text(`Plan de Desarrollo Gachancipá 2024-2027 · PSP-130-2026 · Pág. ${i}/${pageCount}`, W/2, doc.internal.pageSize.getHeight()-3, {align:'center'});
  }

  doc.save(`Informe_PDM_Gachancipa_${activeAnio}.pdf`);
}

function exportExcelFull() {
  const wb = XLSX.utils.book_new();
  const headers = ['Eje','Secretaría','Programa','Producto','Meta','Vigencia','Planeado','Ejecutado','% Avance','Clasificación','Logros'];
  const rows = filteredData.map(r => [
    r.eje, r.secretaria,
    (r.programa||'').replace(/\n/g,' '),
    r.producto,
    getMetaForYear(r,activeAnio),
    activeAnio,
    getPlaneadoForYear(r,activeAnio),
    getEjecutadoForYear(r,activeAnio),
    getAvanceForYear(r,activeAnio),
    getProgForYear(r,activeAnio)==='P'?'Programada':'No Programada',
    r.logros,
  ]);
  const ws = XLSX.utils.aoa_to_sheet([headers,...rows]);
  ws['!cols'] = [4,25,30,25,40,8,10,10,10,15,40].map(w=>({wch:w}));
  XLSX.utils.book_append_sheet(wb, ws, `PDM ${activeAnio}`);
  XLSX.writeFile(wb, `PDM_Gachancipa_${activeAnio}.xlsx`);
}

function exportTableCSV() {
  const headers = ['Eje','Secretaría','Programa','Producto','Meta','Vigencia','Planeado','Ejecutado','% Avance','Clasificación'];
  const rows = filteredData.map(r => [
    r.eje, `"${r.secretaria}"`,
    `"${(r.programa||'').replace(/\n/g,' ').replace(/"/g,'')}"`,
    `"${(r.producto||'').replace(/"/g,'')}"`,
    `"${(getMetaForYear(r,activeAnio)||'').replace(/\n/g,' ').replace(/"/g,'')}"`,
    activeAnio,
    getPlaneadoForYear(r,activeAnio),
    getEjecutadoForYear(r,activeAnio),
    getAvanceForYear(r,activeAnio),
    getProgForYear(r,activeAnio)==='P'?'Programada':'No Programada',
  ].join(','));
  const csv = [headers.join(','), ...rows].join('\n');
  const blob = new Blob(['\uFEFF'+csv],{type:'text/csv;charset=utf-8'});
  const a = document.createElement('a'); a.href=URL.createObjectURL(blob);
  a.download=`PDM_Gachancipa_${activeAnio}.csv`; a.click();
}

function exportTableExcel() { exportExcelFull(); }

async function exportTablePDF() {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation:'landscape', unit:'mm', format:'a4' });
  const W = doc.internal.pageSize.getWidth();

  doc.setFillColor(11,60,93); doc.rect(0,0,W,14,'F');
  doc.setTextColor(212,175,55); doc.setFontSize(11);
  doc.text('Plan de Desarrollo Gachancipá · Tabla Ejecutiva de Metas · '+activeAnio, W/2, 9, {align:'center'});

  const headers = ['Eje','Secretaría','Meta','Vigencia','Plan.','Ejec.','% Av.','Clasif.'];
  const rows = filteredData.map(r => [
    String(r.eje),
    (SEC_ABBR[r.secretaria]||r.secretaria),
    (getMetaForYear(r,activeAnio)||'').replace(/\n/g,' ').substring(0,55),
    activeAnio,
    String(fmt1(getPlaneadoForYear(r,activeAnio))),
    String(fmt1(getEjecutadoForYear(r,activeAnio))),
    fmtPct(getAvanceForYear(r,activeAnio)),
    getProgForYear(r,activeAnio)==='P'?'Prog':'No Prog',
  ]);

  // Manual table rendering
  let y=22, rowH=8;
  const colWidths=[10,28,90,14,14,14,16,16];
  const totalW=colWidths.reduce((a,b)=>a+b,0);
  let x=14;

  doc.setFillColor(26,111,181);
  doc.rect(x,y-5,totalW,rowH,'F');
  doc.setTextColor(255,255,255); doc.setFontSize(7.5);
  let cx=x;
  headers.forEach((h,i)=>{ doc.text(h,cx+1,y); cx+=colWidths[i]; });
  y+=rowH;

  rows.forEach((row,ri)=>{
    if(y>190){ doc.addPage(); y=14; }
    doc.setFillColor(ri%2===0?245:255,ri%2===0?247:255,ri%2===0?250:255);
    doc.rect(x,y-5,totalW,rowH,'F');
    doc.setTextColor(74,85,104); doc.setFontSize(7);
    cx=x;
    row.forEach((cell,i)=>{
      const txt=doc.splitTextToSize(String(cell),colWidths[i]-2);
      doc.text(txt[0],cx+1,y);
      cx+=colWidths[i];
    });
    y+=rowH;
  });

  doc.setFillColor(11,60,93); doc.rect(0,198,W,8,'F');
  doc.setTextColor(255,255,255); doc.setFontSize(7);
  doc.text(`Municipio de Gachancipá · Plan de Desarrollo 2024-2027 · Generado: ${new Date().toLocaleDateString('es-CO')}`, W/2, 203, {align:'center'});

  doc.save(`Tabla_PDM_Gachancipa_${activeAnio}.pdf`);
}

// ══════════════════════════════════════════════════════════════
// BOOTSTRAP: INICIALIZACIÓN
// ══════════════════════════════════════════════════════════════
window.addEventListener('DOMContentLoaded', () => {
  initFilters();
  activeAnio = '2024';
  filteredData = PLAN_DATA.slice();
  applyFilters();

  // Ocultar loader
  setTimeout(() => {
    const loader = $('loader');
    loader.classList.add('hidden');
    setTimeout(() => loader.style.display='none', 500);
  }, 1800);
});
