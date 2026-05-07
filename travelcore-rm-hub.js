var _realAgGrid = window.agGrid || null; // capture real AG Grid before shim
(function(){
  function makeGrid(el, opts){
    if(!el) return {};
    var cols = opts.columnDefs || [];
    var rows = opts.rowData || [];
    function render(rowData){
      var s = '<div style="overflow:auto;width:100%"><table style="width:100%;border-collapse:collapse;font-size:13px">';
      s += '<thead><tr>';
      cols.forEach(function(c){
        s += '<th style="text-align:left;padding:8px 12px;background:#f1f5f9;border-bottom:2px solid #e2e8f0;font-size:11px;font-weight:700;color:#64748b;text-transform:uppercase;letter-spacing:.4px;white-space:nowrap">'+(c.headerName||c.field||'')+'</th>';
      });
      s += '</tr></thead><tbody>';
      rowData.forEach(function(row,i){
        s += '<tr style="background:'+(i%2?'#f8fafc':'#fff')+';border-bottom:1px solid #f1f5f9">';
        cols.forEach(function(c){
          var v = row[c.field];
          if(v==null) v='';
          if(c.cellRenderer){try{v=c.cellRenderer({value:v,data:row});}catch(e){}}
          s += '<td style="padding:8px 12px;color:#374151">'+v+'</td>';
        });
        s += '</tr>';
      });
      s += '</tbody></table></div>';
      el.style.height='';
      el.innerHTML=s;
    }
    render(rows);
    return {
      setGridOption:function(k,v){if(k==='rowData')render(v);},
      setColumnsVisible:function(){},
      setQuickFilter:function(q){
        var filtered=rows.filter(function(r){
          return !q||Object.values(r).some(function(v){return String(v).toLowerCase().includes(q.toLowerCase());});
        });
        render(filtered);
      },
      destroy:function(){}
    };
  }
  window.agGrid = {
    createGrid: makeGrid,
    themeQuartz: { withParams: function(p){ return p||{}; } }
  };
})();

window.html2canvas=function(){return Promise.resolve(document.createElement("canvas"))};

'use strict';

/* ─── SHARED AG-GRID THEME ─── */
var sharedTheme = agGrid.themeQuartz.withParams({
  accentColor: '#0E7B80',
  backgroundColor: '#ffffff',
  foregroundColor: '#181D1F',
  headerBackgroundColor: '#F8F9FA',
  headerTextColor: '#374151',
  borderColor: '#E9ECEF',
  rowBorder: true,
  columnBorder: false,
  rowHoverColor: 'rgba(0,0,0,0.03)',
  selectedRowBackgroundColor: 'rgba(14,123,128,0.08)',
  oddRowBackgroundColor: '#ffffff',
  fontFamily: 'Lato, sans-serif',
  fontSize: 14,
  headerFontSize: 11,
  headerFontWeight: 700,
  wrapperBorderRadius: 8,
  wrapperBorder: true,
  spacing: 8,
  cellHorizontalPadding: 16,
});

/* ─── TARGETS PERFORMANCE DATA ─── */
const TARGETS_DATA = [
  { category: 'Occupancy Rate',    sub: 'Overall hotel occupancy',              forecast: '85%',             current: '87%',                          progress: 102 },
  { category: 'Contract Bookings', sub: 'Travel Distribution Hub bookings',               forecast: '500 | 40% total', current: '523 | 43% total',              progress: 105 },
  { category: 'Promotions',        sub: 'Booking share & conversions',          forecast: '',                current: '28% share · 265 conversions',  progress: 94  },
  { category: 'Monthly Revenue',   sub: 'Total revenue forecast',               forecast: '$750,000',        current: '$698,500',                     progress: 93  },
  { category: 'Average Daily Rate',sub: 'ADR performance forecast',             forecast: '$185H | $165 TO', current: '$172H | $158 TO',              progress: 93  },
];

var _targetsGridApi = null;

function initTargetsGrid() {
  const el = document.getElementById('targetsGrid');
  if (!el || typeof agGrid === 'undefined') return;

  const progressBarHtml = pct => {
    const w = Math.min(pct, 100);
    const cls = pct >= 100 ? 'navy' : 'blue';
    return `<div class="progress-cell"><div class="prog-track"><div class="prog-fill ${cls}" style="width:${w}%"></div></div><span class="prog-label">${pct}%</span></div>`;
  };

  const colDefs = [
    { field: 'category', headerName: 'Target category', flex: 2,
      cellRenderer: p => `<div class="tg-cat-cell"><div class="cell-primary">${p.data.category}</div><div class="cell-sub">${p.data.sub}</div></div>` },
    { field: 'forecast', headerName: 'Forecast', flex: 1.2, headerClass: 'ag-header-center', cellStyle: { textAlign:'center' },
      cellRenderer: p => `<span class="cell-primary">${p.value}</span>` },
    { field: 'current',  headerName: 'Current',  flex: 1.2, headerClass: 'ag-header-center', cellStyle: { textAlign:'center' },
      cellRenderer: p => `<strong>${p.value}</strong>` },
    { field: 'progress', headerName: 'Progress', flex: 2,
      cellRenderer: p => progressBarHtml(p.value) },
  ];

  _targetsGridApi = agGrid.createGrid(el, {
    theme: sharedTheme,
    columnDefs: colDefs,
    rowData: TARGETS_DATA,
    rowHeight: 48,
    headerHeight: 42,
    suppressMovableColumns: true,
    suppressCellFocus: true,
    defaultColDef: { sortable: true, resizable: true, floatingFilter: false, minWidth: 200 },
  });
}

/* ─── ROOM TYPE DATA ─── */
const RT_DATA = [
  { room: 'Standard Room',    to: { b:145, r:'$21.6k', p:'$6.9k', adr:'$150' }, db: { b:93,  r:'$10.6k', p:'$8.2k',  adr:'$190' }, winner:'to-leads',     label:'T LEADS'     },
  { room: 'Deluxe Room',      to: { b:152, r:'$22.8k', p:'$7.3k', adr:'$193' }, db: { b:156, r:'$15.4k', p:'$9.6k',  adr:'$231' }, winner:'direct-leads', label:'DIRECT LEADS' },
  { room: 'Premium Suite',    to: { b:79,  r:'$10.7k', p:'$5.9k', adr:'$240' }, db: { b:92,  r:'$25.8k', p:'$11.6k', adr:'$280' }, winner:'direct-leads', label:'DIRECT LEADS' },
  { room: 'Executive Suite',  to: { b:45,  r:'$10.1k', p:'$4.0k', adr:'$203' }, db: { b:67,  r:'$23.9k', p:'$10.6k', adr:'$350' }, winner:'direct-leads', label:'DIRECT LEADS' },
  { room: 'Presidential Suite',to:{ b:21,  r:'$11.5k', p:'$3.1k', adr:'$930' }, db: { b:34,  r:'$10.6k', p:'$9.2k',  adr:'$690' }, winner:'to-leads',     label:'T LEADS'     },
  { room: 'Family Room',      to: { b:94,  r:'$17.6k', p:'$3.3k', adr:'$380' }, db: { b:74,  r:'$16.7k', p:'$7.5k',  adr:'$270' }, winner:'to-leads',     label:'T LEADS'     },
];

const TOTAL_ROW = {
  room: 'TOTAL', isTotal: true,
  to:  { b:521, r:'$106.9k', p:'$22.1k', adr:'$205' },
  db:  { b:523, r:'$139.3k', p:'$62.5k', adr:'$266' },
  winner: 'direct-wins', label: 'DIRECT WINS'
};

function buildRoomTypeTable() {
  const el = document.getElementById('roomTypeGrid');
  if (!el || typeof agGrid === 'undefined') return;

  const TOTAL_ROW_formatted = {
    room: 'TOTAL', isTotal: true,
    toB: TOTAL_ROW.to.b, toR: TOTAL_ROW.to.r, toP: TOTAL_ROW.to.p, toAdr: TOTAL_ROW.to.adr,
    dbB: TOTAL_ROW.db.b, dbR: TOTAL_ROW.db.r, dbP: TOTAL_ROW.db.p, dbAdr: TOTAL_ROW.db.adr,
    winner: TOTAL_ROW.winner, label: TOTAL_ROW.label,
  };

  const rowData = RT_DATA.map(r => ({
    room: r.room,
    toB: r.to.b, toR: r.to.r, toP: r.to.p, toAdr: r.to.adr,
    dbB: r.db.b, dbR: r.db.r, dbP: r.db.p, dbAdr: r.db.adr,
    winner: r.winner, label: r.label,
  }));

  const colDefs = [
    { field: 'room',  headerName: 'Room type',    flex: 1.4,
      cellRenderer: p => `<span class="cell-primary">${p.value}</span>` },
    { field: 'toB',   headerName: 'T.O bookings', width: 110, type: 'numericColumn' },
    { field: 'toR',   headerName: 'T.O revenue',  width: 110, cellStyle: { textAlign: 'right' } },
    { field: 'toP',   headerName: 'T.O profit',   width: 100, cellStyle: { textAlign: 'right' } },
    { field: 'toAdr', headerName: 'TO ADR',        width: 90,  cellStyle: { textAlign: 'right' } },
    { field: 'dbR',   headerName: 'DB revenue',   width: 110, cellStyle: { textAlign: 'right' } },
    { field: 'dbP',   headerName: 'DB profit',    width: 100, cellStyle: { textAlign: 'right' } },
    { field: 'dbAdr', headerName: 'DB ADR',       width: 90,  cellStyle: { textAlign: 'right' } },
    { field: 'label', headerName: 'Performance',  width: 130, cellStyle: { textAlign: 'center' },
      cellRenderer: p => `<span class="status-badge ${p.data.winner}">${p.value}</span>` },
  ];

  agGrid.createGrid(el, {
    theme: sharedTheme,
    columnDefs: colDefs,
    rowData: rowData,
    pinnedBottomRowData: [TOTAL_ROW_formatted],
    rowHeight: 42,
    headerHeight: 42,
    suppressMovableColumns: true,
    suppressCellFocus: true,
    defaultColDef: { sortable: true, resizable: true, floatingFilter: false, minWidth: 200 },
  });
}

/* ─── FILTER-DRIVEN DATA ─── */

// Targets data per month (1=Jan … 12=Dec)
const TARGETS_BY_MONTH = {
  1:  [ {cat:'Occupancy Rate',    sub:'Overall hotel occupancy',     fc:'81%',            cur:'79%',              prog:97 },
        {cat:'Contract Bookings', sub:'Travel Distribution Hub bookings',      fc:'460 | 37%',      cur:'441 | 35%',        prog:96 },
        {cat:'Promotions',        sub:'Booking share & conversions', fc:'',               cur:'19% share · 270 conversions', prog:87 },
        {cat:'Monthly Revenue',   sub:'Total revenue forecast',      fc:'$680,000',       cur:'$644,000',         prog:95 },
        {cat:'Average Daily Rate',sub:'ADR performance forecast',    fc:'$180H | $162TO', cur:'$175H | $159TO',   prog:97 } ],
  2:  [ {cat:'Occupancy Rate',    sub:'Overall hotel occupancy',     fc:'85%',            cur:'87%',              prog:102 },
        {cat:'Contract Bookings', sub:'Travel Distribution Hub bookings',      fc:'500 | 40%',      cur:'523 | 43%',        prog:105 },
        {cat:'Promotions',        sub:'Booking share & conversions', fc:'',               cur:'28% share · 265 conversions', prog:94 },
        {cat:'Monthly Revenue',   sub:'Total revenue forecast',      fc:'$750,000',       cur:'$698,500',         prog:93 },
        {cat:'Average Daily Rate',sub:'ADR performance forecast',    fc:'$185H | $165TO', cur:'$172H | $158TO',   prog:93 } ],
  3:  [ {cat:'Occupancy Rate',    sub:'Overall hotel occupancy',     fc:'83%',            cur:'86%',              prog:104 },
        {cat:'Contract Bookings', sub:'Travel Distribution Hub bookings',      fc:'510 | 41%',      cur:'530 | 43%',        prog:104 },
        {cat:'Promotions',        sub:'Booking share & conversions', fc:'',               cur:'24% share · 382 conversions', prog:98 },
        {cat:'Monthly Revenue',   sub:'Total revenue forecast',      fc:'$790,000',       cur:'$810,000',         prog:103 },
        {cat:'Average Daily Rate',sub:'ADR performance forecast',    fc:'$188H | $168TO', cur:'$191H | $172TO',   prog:102 } ],
  4:  [ {cat:'Occupancy Rate',    sub:'Overall hotel occupancy',     fc:'80%',            cur:'78%',              prog:98 },
        {cat:'Contract Bookings', sub:'Travel Distribution Hub bookings',      fc:'480 | 38%',      cur:'462 | 37%',        prog:96 },
        {cat:'Promotions',        sub:'Booking share & conversions', fc:'',               cur:'22% share · 298 conversions', prog:90 },
        {cat:'Monthly Revenue',   sub:'Total revenue forecast',      fc:'$720,000',       cur:'$695,000',         prog:97 },
        {cat:'Average Daily Rate',sub:'ADR performance forecast',    fc:'$183H | $163TO', cur:'$178H | $160TO',   prog:97 } ],
  5:  [ {cat:'Occupancy Rate',    sub:'Overall hotel occupancy',     fc:'88%',            cur:'91%',              prog:103 },
        {cat:'Contract Bookings', sub:'Travel Distribution Hub bookings',      fc:'540 | 43%',      cur:'561 | 45%',        prog:104 },
        {cat:'Promotions',        sub:'Booking share & conversions', fc:'',               cur:'31% share · 420 conversions', prog:108 },
        {cat:'Monthly Revenue',   sub:'Total revenue forecast',      fc:'$840,000',       cur:'$875,000',         prog:104 },
        {cat:'Average Daily Rate',sub:'ADR performance forecast',    fc:'$192H | $172TO', cur:'$196H | $176TO',   prog:102 } ],
  6:  [ {cat:'Occupancy Rate',    sub:'Overall hotel occupancy',     fc:'90%',            cur:'93%',              prog:103 },
        {cat:'Contract Bookings', sub:'Travel Distribution Hub bookings',      fc:'560 | 45%',      cur:'581 | 47%',        prog:104 },
        {cat:'Promotions',        sub:'Booking share & conversions', fc:'',               cur:'32% share · 445 conversions', prog:105 },
        {cat:'Monthly Revenue',   sub:'Total revenue forecast',      fc:'$920,000',       cur:'$955,000',         prog:104 },
        {cat:'Average Daily Rate',sub:'ADR performance forecast',    fc:'$198H | $178TO', cur:'$204H | $183TO',   prog:103 } ],
};
// Pad missing months with Feb data
for (let m = 7; m <= 12; m++) TARGETS_BY_MONTH[m] = TARGETS_BY_MONTH[2];

// Seasonality multipliers applied on top of month data
const SEASON_MULT = { all:1, high:1.12, shoulder:0.93, low:0.78 };

function updateTargetsTable() {
  if (!_targetsGridApi) return;
  const month  = parseInt(document.getElementById('tgtMonth')?.value || 2);
  const season = document.getElementById('tgtSeasonality')?.value || 'all';
  const mult   = SEASON_MULT[season];
  const rows   = TARGETS_BY_MONTH[month] || TARGETS_BY_MONTH[2];

  const newData = rows.map(r => {
    const p = Math.round(r.prog * (mult > 1 ? 1 + (mult-1)*0.5 : mult));
    return { category: r.cat, sub: r.sub, forecast: r.fc, current: r.cur, progress: p };
  });

  _targetsGridApi.setGridOption('rowData', newData);
}

// Chart data per metric
const CHART_CONFIGS = {
  revenue:          { label:'Monthly Revenue – Segment Comparison', ticks:['$350K','$262K','$175K','$87K','$0K'],
    base:[95,88,82,78,74,80] },
  adr:              { label:'ADR – Segment Comparison', ticks:['$250','$188','$125','$63','$0'],
    base:[96,90,85,80,76,82] },
  nights:           { label:'Room Nights – Segment Comparison', ticks:['1,000','750','500','250','0'],
    base:[94,87,81,76,72,78] },
  occ:              { label:'Occupancy – Segment Comparison', ticks:['100%','75%','50%','25%','0%'],
    base:[92,86,80,75,70,76] },
  bookings:         { label:'Contract Bookings – Segment Comparison', ticks:['600','450','300','150','0'],
    base:[98,91,85,78,74,80] },
  avg_guests:       { label:'Avg Guests per Room – Segment Comparison', ticks:['4.0','3.0','2.0','1.0','0'],
    base:[100,94,88,82,78,84] },
  total_guests:     { label:'Total Guests – Segment Comparison', ticks:['8,000','6,000','4,000','2,000','0'],
    base:[98,92,86,80,76,82] },
  contracted_rates: { label:'Contracted Rates – Operator', ticks:['$300','$225','$150','$75','$0'],
    base:[96,90,84,78,74,80], toOnly:true },
  allotments:       { label:'Allotments / Guarantees – Operator', ticks:['200','150','100','50','0'],
    base:[100,94,88,82,78,84] },
};

// Segment offset (Y-shift relative to base TO data — higher Y = lower on chart)
const SEGMENT_META = {
  all:        { label:'All Segments',          color:'#006461', yOff: 0  },
  to:         { label:'Static FIT Rates',      color:'#006461', yOff: 0  },
  to_dynamic: { label:'Operator Dynamic', color:'#0891b2', yOff: 14 },
  direct:     { label:'Direct Bookings',       color:'#2563EB', yOff: 28 },
  gds:        { label:'GDS',                   color:'#D97706', yOff: 54 },
  corporate:  { label:'Corporate',             color:'#7C3AED', yOff: 68 },
};

// Comparison line meta
const CMP_META = {
  projection: { label:'Projection by Contract', colorFn: seg => SEGMENT_META[seg]?.color || '#006461', dash: true,  yOff: -3  },
  forecast:   { label:'Forecast',               color:'#f59e0b', dash: true,  yOff: -6  },
  ly:         { label:'Last Year',              color:'#2563EB', dash: false, yOff: 28  },
  stly:       { label:'Same Time Last Year',    color:'#15803d', dash: false, yOff: 22  },
};

// Board type offset + labels + colors
const BOARD_MULT   = { all:0, ai:-8, fb:-4, hb:0, bb:6, ro:10 };
const BOARD_LABELS = { ai:'All Inclusive', fb:'Full Board', hb:'Half Board', bb:'Bed & Breakfast', ro:'Room Only' };
const BOARD_COLORS = { ai:'#065f46', fb:'#1d4ed8', hb:'#b45309', bb:'#7c3aed', ro:'#be185d' };

// Room type offset + labels + colors
const ROOM_OFFSET = { all:0, standard:4, deluxe:0, premium:-6, executive:-10, presidential:-16, family:2 };
const ROOM_LABELS = { standard:'Standard Room', deluxe:'Deluxe Room', premium:'Premium Suite', executive:'Executive Suite', presidential:'Presidential Suite', family:'Family Room' };
const ROOM_COLORS = { standard:'#059669', deluxe:'#0ea5e9', premium:'#967EF3', executive:'#f59e0b', presidential:'#ec4899', family:'#14b8a6' };

const MONTHS_SHORT = ['Aug','Sep','Oct','Nov','Dec','Jan'];

/* active comparison toggles */
let revActiveComps = new Set(['projection', 'direct', 'compset']);

// Visual styles cycling per series index
const SERIES_STYLES = [
  { sw: 2.8, dash: '',            marker: 'circle'   },
  { sw: 2,   dash: '9,5',        marker: 'square'   },
  { sw: 2,   dash: '3,4',        marker: 'diamond'  },
  { sw: 2.2, dash: '10,4,2,4',   marker: 'triangle' },
  { sw: 1.8, dash: '',            marker: 'hollow'   },
  { sw: 2.5, dash: '5,3',        marker: 'square'   },
  { sw: 2,   dash: '2,5',        marker: 'diamond'  },
  { sw: 3,   dash: '8,4',        marker: 'circle'   },
];
// Per-series Y-perturbation (varies shape, not just offset)
const SERIES_VARIATION = [
  [  0,   0,   0,   0,   0,   0],
  [ -8,   6,  -4,  10,  -6,   8],
  [  6,  -8,  12,  -6,   4, -10],
  [-12,   8,   2, -10,  16,  -4],
  [  4,   4, -10,   4,   6,  -6],
  [ -2,  14,  -8,   2, -12,  10],
  [ 10,  -6,   4, -14,   2,   8],
  [ -6,  10,  -2,   8, -10,   4],
];
// Smooth catmull-rom SVG path through points
function smoothPath(pts, xs) {
  const P = xs.map((x, i) => [x, pts[i]]);
  let d = `M${P[0][0]},${P[0][1]}`;
  for (let i = 0; i < P.length - 1; i++) {
    const p0 = P[Math.max(0, i - 1)];
    const p1 = P[i], p2 = P[i + 1];
    const p3 = P[Math.min(P.length - 1, i + 2)];
    const cp1x = p1[0] + (p2[0] - p0[0]) / 5;
    const cp1y = p1[1] + (p2[1] - p0[1]) / 5;
    const cp2x = p2[0] - (p3[0] - p1[0]) / 5;
    const cp2y = p2[1] - (p3[1] - p1[1]) / 5;
    d += ` C${cp1x},${cp1y} ${cp2x},${cp2y} ${p2[0]},${p2[1]}`;
  }
  return d;
}
// Marker SVG per shape type
function markerSvg(type, x, y, color, r) {
  switch (type) {
    case 'square':   return `<rect x="${x-r}" y="${y-r}" width="${r*2}" height="${r*2}" fill="${color}" rx="1"/>`;
    case 'diamond':  return `<polygon points="${x},${y-r*1.4} ${x+r*1.2},${y} ${x},${y+r*1.4} ${x-r*1.2},${y}" fill="${color}"/>`;
    case 'triangle': return `<polygon points="${x},${y-r*1.4} ${x+r*1.3},${y+r} ${x-r*1.3},${y+r}" fill="${color}"/>`;
    case 'hollow':   return `<circle cx="${x}" cy="${y}" r="${r}" fill="white" stroke="${color}" stroke-width="2"/>`;
    default:         return `<circle cx="${x}" cy="${y}" r="${r}" fill="${color}"/>`;
  }
}

// Helper to get checked values from a dropdown group
function getCheckedVals(dropdownId) {
  return Array.from(document.querySelectorAll('#' + dropdownId + ' input[type=checkbox]:checked')).map(cb => cb.value);
}
function getRadioVal(name) {
  const el = document.querySelector('input[name="' + name + '"]:checked');
  return el ? el.value : 'combined';
}

// Color palettes for TO/Source Geo/Source individual lines
const TO_LINE_COLORS   = { tui:'#0ea5e9', 'thomas-cook':'#f59e0b', jet2:'#10b981', fti:'#967EF3', sunwing:'#ec4899', 'club-med':'#f97316' };
const TO_LINE_LABELS   = { tui:'TUI Group', 'thomas-cook':'Thomas Cook', jet2:'Jet2holidays', fti:'FTI Group', sunwing:'Sunwing', 'club-med':'Club Med' };
const ORIGIN_COLORS    = { UK:'#3b82f6', SP:'#f97316', US:'#10b981', MX:'#ec4899' };
const SOURCE_COLORS    = { contract:'#006461', promo:'#967EF3', direct_web:'#0ea5e9', gds_source:'#f59e0b', walk_in:'#6b7280' };
const SOURCE_LABELS    = { contract:'Contract', promo:'Promotion', direct_web:'Direct Web', gds_source:'GDS', walk_in:'Walk-in' };

// ── Revenue Trend: date range state (start/end as Date objects) ──────────────
var revDRFrom = new Date(2025, 7, 1);   // Aug 1 2025
var revDRTo   = new Date(2026, 0, 31);  // Jan 31 2026
var revGranularity = 'month';
var revChartMode   = 'histogram'; // 'lines' | 'histogram'

// Distinct colour palette for unique combinations (primary + compare)
const REV_PALETTE = [
  '#006461','#0891b2','#6366f1','#f59e0b','#ec4899',
  '#10b981','#967EF3','#f97316','#3b82f6','#dc2626',
  '#14b8a6','#a855f7','#eab308','#06b6d4','#84cc16',
];

window.revSetChartMode = function(mode) {
  revChartMode = mode;
  document.getElementById('revChartLines').classList.toggle('active', mode === 'lines');
  document.getElementById('revChartHisto').classList.toggle('active', mode === 'histogram');
  updateChart();
};


function revDRFmt(d) {
  return (d.getMonth()+1) + '/' + d.getDate() + '/' + d.getFullYear();
}
function revDRShortLabel(d) {
  var mn=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return mn[d.getMonth()]+' '+d.getFullYear();
}
function revDRDayDiff(a, b) {
  return Math.round(Math.abs(b - a) / 86400000);
}

window.revSetGran = function(g) {
  revGranularity = g;
  ['month','week','day'].forEach(function(k){
    var btn = document.getElementById('revGran'+k.charAt(0).toUpperCase()+k.slice(1));
    if(btn) btn.classList.toggle('active', k===g);
  });
  updateChart();
};

function revDRUpdateDailyBtn() {
  var days = revDRDayDiff(revDRFrom, revDRTo);
  var btn = document.getElementById('revGranDay');
  if(!btn) return;
  var ok = days <= 31;
  btn.disabled = !ok;
  btn.title = ok ? '' : 'Select 31 days or fewer to enable Daily view';
  if(!ok && revGranularity === 'day') revSetGran('month');
}

function fmtStatK(v) {
  if (v >= 1e6) return '$' + (v / 1e6).toFixed(1) + 'M';
  if (v >= 1e3) return '$' + (v / 1e3).toFixed(1) + 'k';
  return '$' + Math.round(v);
}

function updateRevStats() {
  var BASE_DAYS  = 184; // Aug 1 2025 → Jan 31 2026
  var BASE_T_BOOK = 521,  BASE_T_REV  = 106900;
  var BASE_D_BOOK = 523,  BASE_D_REV  = 139300;
  var days = Math.max(1, Math.round((revDRTo - revDRFrom) / 864e5) + 1);
  var f    = days / BASE_DAYS;
  var tBook = Math.max(1, Math.round(BASE_T_BOOK * f));
  var tRev  = Math.max(0, Math.round(BASE_T_REV  * f));
  var dBook = Math.max(1, Math.round(BASE_D_BOOK * f));
  var dRev  = Math.max(0, Math.round(BASE_D_REV  * f));
  var revGap = Math.abs(dRev - tRev);
  var volDiff = Math.max(1, Math.round(2 * f));
  function set(id, val) { var el = document.getElementById(id); if (el) el.textContent = val; }
  set('rtStatTotalBook',  tBook.toLocaleString());
  set('rtStatTotalRev',   fmtStatK(tRev));
  set('rtStatDirectBook', dBook.toLocaleString());
  set('rtStatDirectRev',  fmtStatK(dRev));
  var fnAdr = document.getElementById('rtFnAdr');
  if (fnAdr) fnAdr.innerHTML = '<span class="fn-dot aug"></span> Avg TO ADR $205 | Direct Bookings ADR $266';
  var fnGap = document.getElementById('rtFnRevGap');
  if (fnGap) fnGap.innerHTML = '<span class="fn-dot rev"></span> Revenue Gap ' + fmtStatK(revGap);
  var fnVol = document.getElementById('rtFnVol');
  if (fnVol) fnVol.innerHTML = '<span class="fn-dot vol"></span> Booking Volume Direct Bookings higher by ' + volDiff + ' room' + (volDiff === 1 ? '' : 's');
}

function revDRApplyRange() {
  var lbl = document.getElementById('revDRLabel');
  if(lbl) lbl.textContent = revDRFmt(revDRFrom) + ' – ' + revDRFmt(revDRTo);
  revDRUpdateDailyBtn();
  updateChart();
  updateRevStats();
}

/* ── IIFE: Revenue Trend Date Range Picker ─── */
(function() {
  var TODAY   = new Date(2026, 2, 9);
  var drFrom  = revDRFrom;
  var drTo    = revDRTo;
  var drPickingTo = false;
  var drHover = null;
  var drViewYear  = 2025;
  var drViewMonth = 7; // 0-indexed = Aug

  var MONTHS = ['January','February','March','April','May','June',
                'July','August','September','October','November','December'];
  var DOWS   = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];

  function fmt(d) { return (d.getMonth()+1) + '/' + d.getDate() + '/' + d.getFullYear(); }
  function sameDay(a,b) { return a&&b&&a.getFullYear()===b.getFullYear()&&a.getMonth()===b.getMonth()&&a.getDate()===b.getDate(); }

  function buildMonthHTML(y, m) {
    var first = new Date(y, m, 1).getDay();
    var days  = new Date(y, m+1, 0).getDate();
    var html  = '<div>';
    html += '<div style="display:grid;grid-template-columns:repeat(7,36px);gap:2px 0;margin-bottom:2px">';
    DOWS.forEach(function(d){ html += '<div class="caldr-dow">'+d+'</div>'; });
    html += '</div><div style="display:grid;grid-template-columns:repeat(7,36px);gap:2px 0">';
    for(var i=0;i<first;i++) html += '<div class="caldr-day caldr-empty"></div>';
    for(var d=1;d<=days;d++) {
      var ts = new Date(y,m,d).getTime();
      html += '<div class="caldr-day" data-ts="'+ts+'" onclick="revDRDayClick('+ts+')">'+d+'</div>';
    }
    html += '</div></div>';
    return html;
  }

  function refreshClasses() {
    var grids = document.getElementById('revDRGrids');
    if(!grids) return;
    var rangeEnd = drPickingTo ? (drHover||null) : drTo;
    grids.querySelectorAll('.caldr-day[data-ts]').forEach(function(el) {
      var dt = new Date(parseInt(el.dataset.ts));
      el.className = 'caldr-day';
      if(sameDay(dt,TODAY))              el.classList.add('caldr-today');
      if(drFrom&&sameDay(dt,drFrom))     el.classList.add('caldr-start');
      if(rangeEnd&&sameDay(dt,rangeEnd)) el.classList.add('caldr-end');
      if(drFrom&&rangeEnd&&!sameDay(drFrom,rangeEnd)){
        var lo=drFrom<rangeEnd?drFrom:rangeEnd, hi=drFrom<rangeEnd?rangeEnd:drFrom;
        if(dt>=lo&&dt<=hi) el.classList.add('caldr-in-range');
      }
    });
    var banner = document.getElementById('revDRPickingBanner');
    if(banner) banner.style.display = drPickingTo ? '' : 'none';
  }

  function updateFooter() {
    var foot = document.getElementById('revDRFooterLabel');
    if(!foot) return;
    var rangeEnd = drPickingTo ? (drHover||null) : drTo;
    if(drFrom&&(drTo||rangeEnd)){
      foot.textContent = fmt(drFrom)+' – '+fmt(rangeEnd||drTo);
      if(drPickingTo&&!drHover) foot.textContent = fmt(drFrom)+' – ... (click end date)';
    } else if(drFrom) {
      foot.textContent = fmt(drFrom)+' – ... (click end date)';
    } else { foot.textContent = 'Select start date'; }
  }

  function renderPicker() {
    var lbl1=document.getElementById('revDRLeft'), lbl2=document.getElementById('revDRRight');
    var grids=document.getElementById('revDRGrids');
    if(!grids) return;
    var m2m=drViewMonth+1, m2y=drViewYear;
    if(m2m>11){m2m=0;m2y++;}
    if(lbl1) lbl1.textContent=MONTHS[drViewMonth]+' '+drViewYear;
    if(lbl2) lbl2.textContent=MONTHS[m2m]+' '+m2y;
    grids.innerHTML = buildMonthHTML(drViewYear,drViewMonth)+buildMonthHTML(m2y,m2m);
    refreshClasses(); updateFooter();
    grids.querySelectorAll('.caldr-day[data-ts]').forEach(function(el){
      el.addEventListener('mouseenter', function(){
        if(drPickingTo){ drHover=new Date(parseInt(this.dataset.ts)); refreshClasses(); updateFooter(); }
      });
    });
    grids.addEventListener('mouseleave', function(){
      if(drPickingTo){ drHover=null; refreshClasses(); updateFooter(); }
    });
  }

  window.revDRDayClick = function(ts) {
    var dt = new Date(ts);
    if(!drPickingTo){
      drFrom=dt; drTo=null; drHover=null; drPickingTo=true;
    } else {
      if(sameDay(dt,drFrom)){ drTo=dt; }
      else if(dt<drFrom){ drTo=drFrom; drFrom=dt; }
      else { drTo=dt; }
      drPickingTo=false; drHover=null;
    }
    refreshClasses(); updateFooter();
  };

  window.revDRNav = function(delta) {
    drViewMonth+=delta;
    while(drViewMonth>11){drViewMonth-=12;drViewYear++;}
    while(drViewMonth<0) {drViewMonth+=12;drViewYear--;}
    renderPicker();
  };

  window.revDRToggle = function() {
    var panel  = document.getElementById('revDRPanel');
    var trigger= document.getElementById('revDRTrigger');
    if(!panel) return;
    if(panel.style.display!=='none'){panel.style.display='none';return;}
    var rect=trigger.getBoundingClientRect();
    var pw=Math.min(720,window.innerWidth*0.95);
    var left=rect.left;
    if(left+pw>window.innerWidth-8) left=Math.max(8,window.innerWidth-pw-8);
    panel.style.left=left+'px'; panel.style.top=(rect.bottom+6)+'px';
    panel.style.width=pw+'px'; panel.style.display='block';
    drPickingTo=false; renderPicker();
  };

  window.revDRCancel = function() {
    document.getElementById('revDRPanel').style.display='none';
    drPickingTo=false; drHover=null;
  };

  window.revDRApply = function() {
    if(!drFrom||!drTo) return;
    document.getElementById('revDRPanel').style.display='none';
    revDRFrom=drFrom; revDRTo=drTo;
    revDRApplyRange();
  };

  window.revDRPreset = function(key) {
    var from=new Date(TODAY), to=new Date(TODAY);
    from.setHours(0,0,0,0); to.setHours(23,59,59,999);
    if     (key==='7d')  { from=new Date(TODAY); from.setDate(from.getDate()-6); }
    else if(key==='30d') { from=new Date(TODAY); from.setDate(from.getDate()-29); }
    else if(key==='90d') { from=new Date(TODAY); from.setDate(from.getDate()-89); }
    else if(key==='3m')  { from=new Date(TODAY); from.setMonth(from.getMonth()-3); from.setDate(1); to=new Date(TODAY.getFullYear(),TODAY.getMonth()+1,0); }
    else if(key==='6m')  { from=new Date(TODAY); from.setMonth(from.getMonth()-6); from.setDate(1); to=new Date(TODAY.getFullYear(),TODAY.getMonth()+1,0); }
    else if(key==='12m') { from=new Date(TODAY); from.setMonth(from.getMonth()-12); from.setDate(1); to=new Date(TODAY.getFullYear(),TODAY.getMonth()+1,0); }
    else if(key==='ytd') { from=new Date(TODAY.getFullYear(),0,1); to=new Date(TODAY); }
    drFrom=from; drTo=to; drPickingTo=false; drHover=null;
    drViewYear=from.getFullYear(); drViewMonth=from.getMonth();
    renderPicker();
    document.getElementById('revDRPanel').style.display='none';
    revDRFrom=drFrom; revDRTo=drTo;
    revDRApplyRange();
  };

  document.addEventListener('click', function(e) {
    var panel=document.getElementById('revDRPanel');
    var wrap=document.getElementById('revDRWrap');
    if(!panel||panel.style.display==='none') return;
    if(wrap&&wrap.contains(e.target)) return;
    panel.style.display='none'; drPickingTo=false;
  }, true);
})();

function revGetXLabels() {
  var mn=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  var labels = [];
  var fromY=revDRFrom.getFullYear(), fromM=revDRFrom.getMonth();
  var toY=revDRTo.getFullYear(),     toM=revDRTo.getMonth();
  var multiYear = fromY !== toY;
  if(revGranularity === 'month') {
    var cy=fromY, cm=fromM;
    while(cy<toY||(cy===toY&&cm<=toM)){
      labels.push(mn[cm]+(multiYear?' '+cy:''));
      cm++; if(cm>11){cm=0;cy++;}
    }
  } else if(revGranularity === 'week') {
    var cur=new Date(revDRFrom); cur.setHours(0,0,0,0); var wk=1;
    while(cur<=revDRTo){ labels.push('W'+wk); cur.setDate(cur.getDate()+7); wk++; }
  } else {
    var cur2=new Date(revDRFrom); cur2.setHours(0,0,0,0);
    while(cur2<=revDRTo){
      labels.push((cur2.getMonth()+1)+'/'+(cur2.getDate()));
      cur2.setDate(cur2.getDate()+1);
    }
  }
  return labels.length ? labels : [mn[fromM], mn[toM]];
}

// Forecast is only available when no Room, Board, Origin, TO, Source are selected
function revForecastEligible() {
  return getCheckedVals('revRoomDropdown').length   === 0
      && getCheckedVals('revBoardDropdown').length  === 0
      && getCheckedVals('revOriginDropdown').length === 0
      && getCheckedVals('revTODropdown').length     === 0
      && getCheckedVals('revSourceDropdown').length === 0;
}

function revSyncForecastChip() {
  var chip = document.querySelector('.rev-cmp-chip[data-cmp="forecast"]');
  if (!chip) return;
  var eligible = revForecastEligible();
  chip.disabled = !eligible;
  chip.style.opacity   = eligible ? '' : '0.35';
  chip.style.cursor    = eligible ? '' : 'not-allowed';
  chip.title = eligible ? '' : 'Forecast is only available when no Room Type, Meal Plan, Origin, Operator or Source Code filters are selected';
  // If currently active but now ineligible, deactivate
  if (!eligible && revActiveComps.has('forecast')) {
    revActiveComps.delete('forecast');
    chip.classList.remove('rev-cmp-active');
  }
}

function updateChart() {
  revSyncForecastChip();
  const svg0 = document.getElementById('revChart');
  if (svg0) svg0.classList.add('updating');
  setTimeout(function() { if (svg0) svg0.classList.remove('updating'); }, 220);

  const metricKey  = document.getElementById('revMetric')?.value  || 'revenue';
  const displayKey = document.getElementById('revDisplay')?.value || 'values';
  const cfg = CHART_CONFIGS[metricKey] || CHART_CONFIGS.revenue;

  // ── Contracted Rates: only Operators have contracted rates ──────────
  // Show one line per selected/all TO; all other filters show an info notice
  const isContractedRates = metricKey === 'contracted_rates';

  // Dynamic x-positions based on date range
  const xLabels2 = revGetXLabels();
  const nPts2    = Math.max(2, xLabels2.length);
  const xs       = Array.from({length: nPts2}, function(_,i){ return Math.round(i * 1305 / (nPts2-1)); });
  const fillPts  = pts => xs.map((x, i) => `${x},${(pts[i]||170)}`).join(' ') + ' 1305,340 0,340';

  // Generate varied base points matching nPts2
  const nudgeVaried = (base, yOff, variation) => {
    return xs.map(function(_, i) {
      const t    = i / Math.max(1, nPts2 - 1);
      const bi   = Math.floor(t * (base.length - 1));
      const bf   = t * (base.length - 1) - bi;
      const bv   = base[bi] + bf * ((base[bi+1]||base[bi]) - base[bi]);
      const varV = variation ? (variation[i % variation.length] || 0) : 0;
      return Math.max(5, Math.min(335, bv + (yOff||0) + varV));
    });
  };

  // Collect selections
  const checkedSegs    = getCheckedVals('revSegmentDropdown');
  const checkedRooms   = getCheckedVals('revRoomDropdown');
  const checkedBoards  = getCheckedVals('revBoardDropdown');
  const checkedTOs     = getCheckedVals('revTODropdown');
  const checkedOrigins = getCheckedVals('revOriginDropdown');
  const checkedSources = getCheckedVals('revSourceDropdown');

  const toMode      = getRadioVal('revTOMode');
  const originMode  = getRadioVal('revOriginMode');
  const sourceMode  = getRadioVal('revSourceMode');
  const segMode     = getRadioVal('revSegmentMode');
  const roomMode    = getRadioVal('revRoomMode');
  const boardMode   = getRadioVal('revBoardMode');

  // Determine if cross-dimension labelling is needed (>1 active dimension)
  const activeDims = [
    checkedSegs.length > 0    ? checkedSegs    : null,
    checkedTOs.length > 0     ? checkedTOs     : null,
    checkedOrigins.length > 0 ? checkedOrigins : null,
    checkedSources.length > 0 ? checkedSources : null,
    checkedRooms.length > 0   ? checkedRooms   : null,
    checkedBoards.length > 0  ? checkedBoards  : null,
  ].filter(Boolean);
  const multiDim = activeDims.length > 1;

  // ── Contracted Rates override ────────────────────────────────────────
  if (isContractedRates) {
    const svg = document.getElementById('revChart');
    const yAxis = document.getElementById('revYAxis');
    const legend = document.getElementById('revLegend');
    const sub = document.getElementById('revSubtitle');

    // Determine which TOs to show
    const toKeys = checkedTOs.length > 0 ? checkedTOs
      : Object.keys(TO_LINE_COLORS); // all TOs if none selected
    const gridLines = [0,60,120,180,240].map(y=>`<line x1="0" y1="${y}" x2="1305" y2="${y}" stroke="#f0f0f0" stroke-width="1"/>`).join('');
    const toSeries = toKeys.map(function(toKey, ti) {
      const color = TO_LINE_COLORS[toKey] || '#888';
      const label = TO_LINE_LABELS[toKey] || toKey;
      const yOff  = ti * 10;
      const variation = SERIES_VARIATION[ti % SERIES_VARIATION.length];
      const pts = nudgeVaried(cfg.base, yOff, variation);
      const st = SERIES_STYLES[ti % SERIES_STYLES.length];
      return { label, color, pts, style:st, isFirst: ti===0 };
    });
    const fillArea = toSeries.length ? `<path d="${smoothPath(toSeries[0].pts, xs)} L1305,340 L0,340 Z" fill="${toSeries[0].color}15"/>` : '';
    const seriesSvg = toSeries.map(s => {
      const markers = xs.slice(1).map((x,i) => markerSvg(s.style.marker, x, s.pts[i+1], s.color, s.isFirst?4.5:4)).join('');
      return `<path d="${smoothPath(s.pts, xs)}" fill="none" stroke="${s.color}" stroke-width="${s.style.sw}"/>${markers}`;
    }).join('');
    if(svg) svg.innerHTML = gridLines + fillArea + seriesSvg;
    if(yAxis) yAxis.innerHTML = cfg.ticks.map(t=>`<span>${t}</span>`).join('');
    const granLbl2 = revGranularity==='week'?'Weekly':revGranularity==='day'?'Daily':'Monthly';
    if(sub) sub.textContent = granLbl2 + ' ' + cfg.label + ' · ' + revDRShortLabel(revDRFrom) + ' – ' + revDRShortLabel(revDRTo);
    if(legend) legend.innerHTML = toSeries.map(s => `<div class="legend-item"><svg class="legend-line-svg" viewBox="0 0 28 10" width="28" height="10"><line x1="0" y1="5" x2="28" y2="5" stroke="${s.color}" stroke-width="2"/>${markerSvg(s.style.marker,14,5,s.color,3.5)}</svg>${s.label}</div>`).join('')
      + '<div class="legend-item" style="color:#9ca3af;font-size:11px;margin-left:8px">⚠ Other segments: no contracted rates apply</div>';
    // Update x-axis
    const xAxisEl2 = document.querySelector('.chart-x-axis');
    if(xAxisEl2){ const step2=Math.ceil(nPts2/10); xAxisEl2.innerHTML=xLabels2.filter((_,i)=>i%step2===0||i===nPts2-1).map(l=>`<span>${l}</span>`).join(''); }
    return;
  }

  // Build primary series
  const primarySeries = [];

  // Helper: build a series entry
  function makeSeries(label, color, yOff, dimTag) {
    const lbl = multiDim && dimTag ? label + ' · ' + dimTag : label;
    return { label: lbl, color, yOff: yOff || 0 };
  }

  if (checkedTOs.length > 0 && toMode === 'individual') {
    // One line per TO
    checkedTOs.forEach((to, ti) => {
      const color = TO_LINE_COLORS[to] || '#888';
      const toLabel = TO_LINE_LABELS[to] || to;
      if (checkedOrigins.length > 0 && originMode === 'individual') {
        checkedOrigins.forEach((org, oi) => {
          primarySeries.push(makeSeries(toLabel, color, ti * 8 + oi * 4, org));
        });
      } else {
        primarySeries.push(makeSeries(toLabel, color, ti * 8));
      }
    });
  } else if (checkedOrigins.length > 0 && originMode === 'individual') {
    // One line per origin
    checkedOrigins.forEach((org, oi) => {
      const color = ORIGIN_COLORS[org] || '#888';
      if (checkedSources.length > 0 && sourceMode === 'individual') {
        checkedSources.forEach((src, si) => {
          primarySeries.push(makeSeries(org, color, oi * 8 + si * 4, SOURCE_LABELS[src] || src));
        });
      } else {
        primarySeries.push(makeSeries(org, color, oi * 8));
      }
    });
  } else if (checkedSources.length > 0 && sourceMode === 'individual') {
    checkedSources.forEach((src, si) => {
      primarySeries.push(makeSeries(SOURCE_LABELS[src] || src, SOURCE_COLORS[src] || '#888', si * 8));
    });
  } else {
    // ── Cartesian product: Seg × Room, Seg × Board, or Seg × Room × Board ────
    const segsInd   = checkedSegs.length > 0 && segMode === 'individual';
    const roomsInd  = checkedRooms.length > 0 && roomMode === 'individual';
    const boardsInd = checkedBoards.length > 0 && boardMode === 'individual';

    // Resolve each dimension to a flat list of {label, color, yOff} entries
    // '_combined_*' sentinel = dimension selected but in combined mode
    const segList = segsInd
      ? checkedSegs.map(s => { const m = SEGMENT_META[s]||SEGMENT_META.all; return { key:s, label:m.label, color:m.color, yOff:m.yOff }; })
      : (checkedSegs.length > 0 ? [{ key:'_seg', label:'Segs (combined)', color:'#006461', yOff:0 }] : []);

    const roomList = roomsInd
      ? checkedRooms.map(r => ({ key:r, label:ROOM_LABELS[r]||r, color:ROOM_COLORS[r]||'#059669', yOff:ROOM_OFFSET[r]||0 }))
      : (checkedRooms.length > 0 ? [{ key:'_room', label:'Rooms (combined)', color:'#059669', yOff:4 }] : []);

    const boardList = boardsInd
      ? checkedBoards.map(b => ({ key:b, label:BOARD_LABELS[b]||b, color:BOARD_COLORS[b]||'#967EF3', yOff:BOARD_MULT[b]||0 }))
      : (checkedBoards.length > 0 ? [{ key:'_board', label:'Meal Plans (combined)', color:'#967EF3', yOff:6 }] : []);

    const hasSeg   = segList.length   > 0;
    const hasRoom  = roomList.length  > 0;
    const hasBoard = boardList.length > 0;

    // Build cartesian product across whichever dimensions are active with segs
    // Priority: Seg × Room × Board > Seg × Room > Seg × Board > standalone each
    let lineIdx = 0;
    if (hasSeg && hasRoom && hasBoard) {
      // 3-way product
      segList.forEach(seg => {
        roomList.forEach(room => {
          boardList.forEach(board => {
            const parts = [seg.label, room.label, board.label];
            primarySeries.push({ label: parts.join(' · '), color: seg.color, yOff: seg.yOff + room.yOff + board.yOff + lineIdx * 3 });
            lineIdx++;
          });
        });
      });
    } else if (hasSeg && hasRoom) {
      // Seg × Room
      segList.forEach(seg => {
        roomList.forEach(room => {
          primarySeries.push({ label: seg.label + ' · ' + room.label, color: seg.color, yOff: seg.yOff + room.yOff + lineIdx * 3 });
          lineIdx++;
        });
      });
    } else if (hasSeg && hasBoard) {
      // Seg × Board (Meal Plan)
      segList.forEach(seg => {
        boardList.forEach(board => {
          primarySeries.push({ label: seg.label + ' · ' + board.label, color: seg.color, yOff: seg.yOff + board.yOff + lineIdx * 3 });
          lineIdx++;
        });
      });
    } else if (hasSeg) {
      segList.forEach(seg => primarySeries.push({ label: seg.label, color: seg.color, yOff: seg.yOff }));
    } else if (hasRoom && hasBoard) {
      // Room × Board (no seg)
      roomList.forEach(room => {
        boardList.forEach(board => {
          primarySeries.push({ label: room.label + ' · ' + board.label, color: room.color, yOff: room.yOff + board.yOff + lineIdx * 3 });
          lineIdx++;
        });
      });
    } else if (hasRoom) {
      roomList.forEach(room => primarySeries.push({ label: room.label, color: room.color, yOff: room.yOff }));
    } else if (hasBoard) {
      boardList.forEach(board => primarySeries.push({ label: board.label, color: board.color, yOff: board.yOff }));
    }

    if (primarySeries.length === 0) {
      primarySeries.push({ label: 'All Segments', color: '#006461', yOff: 0 });
    }
  }

  // Build compare series
  const compareSeries = [];

  if (revActiveComps.has('forecast') && revForecastEligible()) {
    // One forecast line per primary series (segment-only context)
    primarySeries.forEach(function(ps) {
      compareSeries.push({
        label: ps.label + ' (Forecast)',
        color: ps.color,
        yOff:  ps.yOff - 6,
        dash:  '4 4',
        forceOpacity: 0.7
      });
    });
  }

  // LY and STLY: one mirrored line per primary series combination
  // Each compare line gets the same color as its primary but lighter + dashed
  if (revActiveComps.has('ly')) {
    primarySeries.forEach(function(ps) {
      compareSeries.push({
        label: ps.label + ' (LY)',
        color: ps.color,
        yOff:  ps.yOff + 18,
        dash:  '8 4',
        forceOpacity: 0.65
      });
    });
  }
  if (revActiveComps.has('stly')) {
    primarySeries.forEach(function(ps) {
      compareSeries.push({
        label: ps.label + ' (STLY)',
        color: ps.color,
        yOff:  ps.yOff + 14,
        dash:  '5 5',
        forceOpacity: 0.5
      });
    });
  }

  // ── Assign unique colours per series from palette ────────────────────────
  const totalSeriesCount = primarySeries.length + compareSeries.length;
  let paletteIdx = 0;
  function nextColor(hint) {
    // Use palette index; hint (original color) used as fallback only when palette exhausted
    return REV_PALETTE[paletteIdx++ % REV_PALETTE.length];
  }

  let styleIdx = 0;
  const colSeries = primarySeries.map((s, i) => {
    const si = styleIdx++;
    const variation = SERIES_VARIATION[si % SERIES_VARIATION.length];
    const color = nextColor(s.color);
    return { ...s, color, pts: nudgeVaried(cfg.base, s.yOff, variation), style: SERIES_STYLES[si % SERIES_STYLES.length], isCol: true };
  });
  const lineSeries = compareSeries.map(s => {
    const si = styleIdx++;
    const color = nextColor(s.color);
    return { ...s, color, pts: nudgeVaried(cfg.base, s.yOff, SERIES_VARIATION[si % SERIES_VARIATION.length]), style: SERIES_STYLES[si % SERIES_STYLES.length], isCol: false };
  });
  const allSeries = [...colSeries, ...lineSeries];

  let ticks = cfg.ticks;
  if (displayKey === 'pct') ticks = ['+20%', '+10%', '0%', '-10%', '-20%'];

  const svg = document.getElementById('revChart');
  if (!svg) return;

  const svgW = 1305, svgH = 340;
  const refY  = 170;   // mid-point parity line (SVG y-coord)
  const slotW = svgW / nPts2;

  // ── 1. Weekend shading (daily granularity only) ─────────────────────────
  let weekendBands = '';
  if (revGranularity === 'day') {
    var cur3 = new Date(revDRFrom); cur3.setHours(0,0,0,0);
    xs.forEach(function(x, i) {
      var dow = cur3.getDay(); // 0=Sun,6=Sat
      if (dow === 0 || dow === 6) {
        var bw = Math.max(slotW * 0.92, 6);
        weekendBands += `<rect x="${(x - bw/2).toFixed(1)}" y="0" width="${bw.toFixed(1)}" height="${svgH}" fill="#f0f0f0" opacity="0.7"/>`;
      }
      cur3.setDate(cur3.getDate() + 1);
    });
  }

  // ── 2. Horizontal grid lines ─────────────────────────────────────────────
  const gridLines = [0, 68, 170, 272, 340].map(y =>
    `<line x1="0" y1="${y}" x2="${svgW}" y2="${y}" stroke="#d4d4d4" stroke-width="0.8"/>`
  ).join('');

  // ── 3. Neutral zone — removed ────────────────────────────────────────────
  const neutralZone = '';

  // ── 4. Benchmark / goal line (dashed) ───────────────────────────────────
  const goalY = refY - 70;
  const benchmarkLine = `<line x1="0" y1="${goalY}" x2="${svgW}" y2="${goalY}" stroke="#004948" stroke-width="1.5" stroke-dasharray="6 4" opacity="0.5"/>`;

  // ── 5. Parity reference line — removed ──────────────────────────────────
  const parityLine = '';

  // ── 6. Above / Below labels — removed ───────────────────────────────────
  const aboveLabel = '';
  const belowLabel = '';

  let mainSvg = '';

  if (revChartMode === 'histogram') {
    // ── Figma-style above/below bars split at parity line ────────────────
    const barW = Math.max(4, Math.min(slotW * 0.6, 36));

    // Draw bars for first (primary) series only — split at refY
    if (colSeries.length > 0) {
      const primary = colSeries[0];
      for (let xi = 0; xi < nPts2; xi++) {
        const cx   = xs[xi];
        const left = (cx - barW / 2).toFixed(1);
        const py   = primary.pts[xi] || refY;
        if (py < refY) {
          // Above parity → bar grows upward from refY, deep teal
          const h = (refY - py).toFixed(1);
          mainSvg += `<rect x="${left}" y="${py.toFixed(1)}" width="${barW.toFixed(1)}" height="${h}" fill="#004948" opacity="0.85" rx="2"/>`;
        } else if (py > refY) {
          // Below parity → bar grows downward from refY, coral/red
          const h = (py - refY).toFixed(1);
          mainSvg += `<rect x="${left}" y="${refY}" width="${barW.toFixed(1)}" height="${h}" fill="#e85d4a" opacity="0.75" rx="2"/>`;
        }
      }
      // Additional series as narrower offset bars
      colSeries.slice(1).forEach(function(s, si) {
        const bw2  = Math.max(3, barW * 0.5);
        const off  = (si % 2 === 0 ? 1 : -1) * (barW * 0.35);
        for (let xi = 0; xi < nPts2; xi++) {
          const cx  = xs[xi] + off;
          const lft = (cx - bw2 / 2).toFixed(1);
          const py  = s.pts[xi] || refY;
          const clr = py < refY ? '#0891b2' : '#f59e0b';
          const yR  = py < refY ? py : refY;
          const hR  = Math.abs(py - refY).toFixed(1);
          mainSvg += `<rect x="${lft}" y="${yR.toFixed(1)}" width="${bw2.toFixed(1)}" height="${hR}" fill="${clr}" opacity="0.55" rx="1"/>`;
        }
      });
    }

    // Trend line through primary series — bold, dark
    if (colSeries.length > 0) {
      mainSvg += `<path d="${smoothPath(colSeries[0].pts, xs)}" fill="none" stroke="#1e293b" stroke-width="2.5" stroke-linejoin="round"/>`;
    }

    // Scatter dots for each series at each data point
    const DOT_COLORS = ['#004948','#0891b2','#6366f1','#dc2626','#f59e0b','#ec4899','#10b981','#967EF3'];
    allSeries.forEach(function(s, si) {
      const dotClr = DOT_COLORS[si % DOT_COLORS.length];
      const r = si === 0 ? 5.5 : 4.5;
      xs.forEach(function(x, xi) {
        const py = s.pts[xi] || refY;
        mainSvg += `<circle cx="${x}" cy="${py.toFixed(1)}" r="${r}" fill="${dotClr}" stroke="#fff" stroke-width="2"/>`;
      });
    });

    // Compare lines overlaid
    lineSeries.forEach(function(s) {
      const dashStr  = s.dash || '';
      const dashAttr = dashStr ? ` stroke-dasharray="${dashStr}"` : '';
      const op       = s.forceOpacity || 1;
      const opAttr   = op < 1 ? ` opacity="${op}"` : '';
      mainSvg += `<path d="${smoothPath(s.pts, xs)}" fill="none" stroke="${s.color}" stroke-width="1.5"${dashAttr}${opAttr}/>`;
    });

  } else {
    // ── Lines mode: fill area, smooth lines, dots ────────────────────────
    if (colSeries.length > 0) {
      const fc = colSeries[0];
      // Positive yield: teal fill above refY
      mainSvg += `<clipPath id="aboveClip"><rect x="0" y="0" width="${svgW}" height="${refY}"/></clipPath>`;
      mainSvg += `<path d="${smoothPath(fc.pts, xs)} L${svgW},${svgH} L0,${svgH} Z" fill="#004948" opacity="0.12" clip-path="url(#aboveClip)"/>`;
      // Net leak: coral fill below refY
      mainSvg += `<clipPath id="belowClip"><rect x="0" y="${refY}" width="${svgW}" height="${svgH - refY}"/></clipPath>`;
      mainSvg += `<path d="${smoothPath(fc.pts, xs)} L${svgW},${svgH} L0,${svgH} Z" fill="#e85d4a" opacity="0.10" clip-path="url(#belowClip)"/>`;
    }
    allSeries.forEach(function(s, ai) {
      const st       = s.style || SERIES_STYLES[ai % SERIES_STYLES.length];
      const isComp   = !s.isCol;
      const dashStr  = isComp ? (s.dash || '6 4') : '';
      const dashAttr = dashStr ? ` stroke-dasharray="${dashStr}"` : '';
      const op       = s.forceOpacity || 1;
      const opAttr   = op < 1 ? ` opacity="${op}"` : '';
      const sw       = s.isCol ? 2.8 : 2;
      const r        = s.isCol ? 5 : 4;
      const markers  = xs.map(function(x, i) {
        if (isComp) {
          return `<circle cx="${x}" cy="${(s.pts[i]||refY).toFixed(1)}" r="${r}" fill="#fff" stroke="${s.color}" stroke-width="2.5"/>`;
        }
        return `<circle cx="${x}" cy="${(s.pts[i]||refY).toFixed(1)}" r="${r}" fill="${s.color}" stroke="#fff" stroke-width="2"/>`;
      }).join('');
      mainSvg += `<path d="${smoothPath(s.pts, xs)}" fill="none" stroke="${s.color}" stroke-width="${sw}"${dashAttr}${opAttr}/>${markers}`;
    });
  }

  // Compose SVG layers: weekend → grid → neutral → bars/lines → refs → labels
  svg.innerHTML = weekendBands + gridLines + neutralZone + mainSvg + benchmarkLine + parityLine + aboveLabel + belowLabel;

  const yAxis = document.getElementById('revYAxis');
  if (yAxis) yAxis.innerHTML = ticks.map(t => `<span>${t}</span>`).join('');

  // Update chart x-axis label bar
  var xAxisEl = document.querySelector('.chart-x-axis');
  if(xAxisEl) {
    var step = Math.ceil(nPts2 / 10);
    xAxisEl.innerHTML = xLabels2.filter(function(_,i){ return i % step === 0 || i === nPts2-1; }).map(function(l){ return '<span>'+l+'</span>'; }).join('');
  }

  // Granularity label
  var granLbl = revGranularity === 'week' ? 'Weekly' : revGranularity === 'day' ? 'Daily' : 'Monthly';
  var sub = document.getElementById('revSubtitle');
  if (sub) sub.textContent = granLbl + ' ' + cfg.label + ' · ' + revDRShortLabel(revDRFrom) + ' – ' + revDRShortLabel(revDRTo);

  const legend = document.getElementById('revLegend');
  if (legend) {
    // Fixed legend: Above / Below parity + trend line + series dots
    const DOT_COLORS2 = ['#004948','#0891b2','#6366f1','#dc2626','#f59e0b','#ec4899','#10b981','#967EF3'];
    let legendHtml = '';
    if (revChartMode === 'histogram') {
      legendHtml += `<div class="legend-item"><svg viewBox="0 0 28 10" width="28" height="10"><line x1="0" y1="5" x2="28" y2="5" stroke="#1e293b" stroke-width="2.5"/></svg>Trend</div>`;
      legendHtml += `<div class="legend-item"><svg viewBox="0 0 28 10" width="28" height="10"><line x1="0" y1="5" x2="28" y2="5" stroke="#004948" stroke-width="1.5" stroke-dasharray="6 4" opacity="0.6"/></svg>Benchmark</div>`;
    }
    allSeries.forEach(function(s, si) {
      const dotClr = DOT_COLORS2[si % DOT_COLORS2.length];
      const dashStr  = s.isCol ? '' : (s.dash || '');
      const dashAttr = dashStr ? ` stroke-dasharray="${dashStr}"` : '';
      if (revChartMode === 'histogram') {
        legendHtml += `<div class="legend-item"><svg viewBox="0 0 10 10" width="10" height="10"><circle cx="5" cy="5" r="4.5" fill="${dotClr}" stroke="#fff" stroke-width="1.5"/></svg>${s.label}</div>`;
      } else {
        legendHtml += `<div class="legend-item"><svg class="legend-line-svg" viewBox="0 0 28 10" width="28" height="10"><line x1="0" y1="5" x2="28" y2="5" stroke="${s.color}" stroke-width="2.5"${dashAttr}/><circle cx="14" cy="5" r="4" fill="${s.color}" stroke="#fff" stroke-width="1.5"/></svg>${s.label}</div>`;
      }
    });
    legend.innerHTML = legendHtml;
  }

}

// ── Close all dropdowns (call before opening any one dropdown) ────────────────
function _closeAllDropdowns(exceptId) {
  var _allDdIds = [
    'revBoardDropdown','revSegmentDropdown','revRoomDropdown',
    'revTODropdown','revOriginDropdown','revSourceDropdown',
    'revCmpDropdown','revDRPanel',
    'calFiltersDropdown','wvFiltersDropdown'
  ];
  _allDdIds.forEach(function(id) {
    if (id === exceptId) return;
    var el = document.getElementById(id);
    if (el) el.style.display = 'none';
  });
  // Monthly filter panels
  ['calFiltTOPanel','calFiltRTPanel','calFiltMPPanel','calFiltOriginPanel'].forEach(function(pid) {
    if (pid === exceptId) return;
    var p = document.getElementById(pid);
    if (p) p.style.display = 'none';
  });
  // Remove active states from all filter / dropdown buttons
  ['calFiltersBtn','wvFiltersBtn','revCmpBtn',
   'revBoardBtn','revSegmentBtn','revRoomBtn','revTOBtn','revOriginBtn','revSourceBtn',
   'calFiltTOBtn','calFiltRTBtn','calFiltMPBtn','calFiltOriginBtn'].forEach(function(bid) {
    var b = document.getElementById(bid);
    if (b) b.classList.remove('active');
  });
}

// Generic multi-select dropdown builder
function makeRevMultiSelect(wrapId, btnId, ddId, allLabel) {
  const wrap = document.getElementById(wrapId);
  const btn  = document.getElementById(btnId);
  const dd   = document.getElementById(ddId);
  if (!btn || !dd) return;
  btn.addEventListener('click', function(e) {
    e.stopPropagation();
    const isOpen = dd.style.display !== 'none';
    if (!isOpen) _closeAllDropdowns(ddId);
    dd.style.display = isOpen ? 'none' : '';
  });
  document.addEventListener('click', function(e) {
    if (wrap && !wrap.contains(e.target)) dd.style.display = 'none';
  });
  /* Keep dropdown open when clicking inside it */
  dd.addEventListener('click', function(e) { e.stopPropagation(); });
  dd.querySelectorAll('input').forEach(cb => {
    cb.addEventListener('change', function() {
      const checked = Array.from(dd.querySelectorAll('input:checked'));
      btn.firstChild.textContent = checked.length === 0 ? allLabel
        : checked.length === 1 ? checked[0].closest('label').textContent.trim()
        : checked.length + ' selected';
      updateChart();
    });
  });
}
makeRevMultiSelect('revBoardWrap',   'revBoardBtn',   'revBoardDropdown',   'All Types');
makeRevMultiSelect('revSegmentWrap', 'revSegmentBtn', 'revSegmentDropdown', 'All Segments');
makeRevMultiSelect('revRoomWrap',    'revRoomBtn',    'revRoomDropdown',    'All Rooms');
makeRevMultiSelect('revTOWrap',      'revTOBtn',      'revTODropdown',      'All TOs');
makeRevMultiSelect('revOriginWrap',  'revOriginBtn',  'revOriginDropdown',  'All Origins');
makeRevMultiSelect('revSourceWrap',  'revSourceBtn',  'revSourceDropdown',  'All Sources');

// Sync button labels with pre-checked defaults
(function syncDefaultLabels() {
  var pairs = [
    ['revSegmentBtn','revSegmentDropdown','All Segments'],
    ['revRoomBtn','revRoomDropdown','All Rooms'],
    ['revBoardBtn','revBoardDropdown','All Types'],
    ['revTOBtn','revTODropdown','All TOs'],
    ['revOriginBtn','revOriginDropdown','All Origins'],
    ['revSourceBtn','revSourceDropdown','All Sources']
  ];
  pairs.forEach(function(p) {
    var btn = document.getElementById(p[0]);
    var dd  = document.getElementById(p[1]);
    if (!btn || !dd) return;
    var checked = Array.from(dd.querySelectorAll('input[type=checkbox]:checked'));
    if (checked.length > 0) {
      btn.firstChild.textContent = checked.length === 1
        ? checked[0].closest('label').textContent.trim()
        : checked.length + ' selected';
    }
  });
})();

// Listen for mode radio changes on all dropdowns
['revTOMode','revOriginMode','revSourceMode','revSegmentMode','revRoomMode','revBoardMode'].forEach(function(name) {
  document.querySelectorAll('input[name="'+name+'"]').forEach(function(radio) {
    radio.addEventListener('change', updateChart);
  });
});

// Compare against dropdown
(function() {
  const wrap = document.getElementById('revCmpWrap');
  const btn  = document.getElementById('revCmpBtn');
  const dd   = document.getElementById('revCmpDropdown');
  if (!btn || !dd) return;

  function syncCmpBtn() {
    const count = revActiveComps.size;
    const textNode = btn.childNodes[1];
    if (textNode) textNode.textContent = ' Compare' + (count ? ' (' + count + ')' : '') + ' ';
  }

  btn.addEventListener('click', function(e) {
    e.stopPropagation();
    const isOpen = dd.style.display !== 'none';
    if (!isOpen) _closeAllDropdowns('revCmpDropdown');
    dd.style.display = isOpen ? 'none' : '';
  });
  document.addEventListener('click', function(e) {
    if (wrap && !wrap.contains(e.target)) dd.style.display = 'none';
  });
  dd.querySelectorAll('.rev-cmp-chip').forEach(function(chipBtn) {
    chipBtn.addEventListener('click', function() {
      const cmp = this.dataset.cmp;
      // Block forecast if ineligible
      if (cmp === 'forecast' && !revForecastEligible()) return;
      if (revActiveComps.has(cmp)) {
        revActiveComps.delete(cmp);
        this.classList.remove('rev-cmp-active');
      } else {
        revActiveComps.add(cmp);
        this.classList.add('rev-cmp-active');
      }
      syncCmpBtn();
      updateChart();
    });
  });
})();

// Listen on remaining revenue filters
['revMetric','revDisplay'].forEach(id => {
  document.getElementById(id)?.addEventListener('change', updateChart);
});

// Listen on targets filters
document.getElementById('tgtMonth')?.addEventListener('change', updateTargetsTable);
document.getElementById('tgtSeasonality')?.addEventListener('change', updateTargetsTable);

/* ─── DEMAND CALENDAR ─── */
const ALL_MONTHS = [
  { name: 'January 2026',  year: 2026, month: 1, days: 31, firstDay: 4,
    stats: { occ: '62%', occDelta: '+2.10', adr: '$165', adrDelta: '+1.80', rev: '$421k', revDelta: '+3.90' }, lockedCount: 1 },
  { name: 'February 2026', year: 2026, month: 2, days: 28, firstDay: 0,
    stats: { occ: '65%', occDelta: '+4.37', adr: '$178', adrDelta: '+3.90', rev: '$483k', revDelta: '+6.47' }, lockedCount: 2 },
  { name: 'March 2026',    year: 2026, month: 3, days: 31, firstDay: 0,
    stats: { occ: '66%', occDelta: '+3.42', adr: '$183', adrDelta: '+1 vs LY', rev: '$562k', revDelta: '+6 vs LY' }, lockedCount: 3 },
  { name: 'April 2026',    year: 2026, month: 4, days: 30, firstDay: 3,
    stats: { occ: '66%', occDelta: '+2.12', adr: '$188', adrDelta: '+1.18', rev: '$570k', revDelta: '+6.18' }, lockedCount: 2 },
  { name: 'May 2026',      year: 2026, month: 5, days: 31, firstDay: 5,
    stats: { occ: '71%', occDelta: '+5.20', adr: '$196', adrDelta: '+4.10', rev: '$641k', revDelta: '+8.30' }, lockedCount: 1 },
  { name: 'June 2026',     year: 2026, month: 6, days: 30, firstDay: 1,
    stats: { occ: '74%', occDelta: '+3.80', adr: '$210', adrDelta: '+5.60', rev: '$712k', revDelta: '+9.10' }, lockedCount: 0 },
  { name:'July 2026',      year:2026, month:7,  days:31, firstDay:3, lockedCount:2, stats:{occ:'78%',occDelta:'+2.1',adr:'$172',adrDelta:'+$8',rev:'$562k',revDelta:'+6.1%'} },
  { name:'August 2026',    year:2026, month:8,  days:31, firstDay:6, lockedCount:4, stats:{occ:'91%',occDelta:'+5.3',adr:'$198',adrDelta:'+$14',rev:'$710k',revDelta:'+9.2%'} },
  { name:'September 2026', year:2026, month:9,  days:30, firstDay:2, lockedCount:1, stats:{occ:'74%',occDelta:'+1.8',adr:'$162',adrDelta:'+$5',rev:'$490k',revDelta:'+3.5%'} },
  { name:'October 2026',   year:2026, month:10, days:31, firstDay:4, lockedCount:3, stats:{occ:'69%',occDelta:'-1.2',adr:'$148',adrDelta:'-$3',rev:'$434k',revDelta:'-2.1%'} },
  { name:'November 2026',  year:2026, month:11, days:30, firstDay:0, lockedCount:2, stats:{occ:'62%',occDelta:'-3.4',adr:'$138',adrDelta:'-$7',rev:'$375k',revDelta:'-5.8%'} },
  { name:'December 2026',  year:2026, month:12, days:31, firstDay:2, lockedCount:5, stats:{occ:'85%',occDelta:'+4.1',adr:'$205',adrDelta:'+$18',rev:'$692k',revDelta:'+8.4%'} },
];

let calStartIdx = 0; // start at January
let calView = 2;        // default 2 months on load
let calDisplayView = 2; // default 2 months on load
let calRangeFrom   = new Date(2026, 0, 1);  // active date-range start (global)
let calRangeTo     = new Date(2026, 11, 31); // active date-range end   (global)
let calDateRangeStart = null; // start of selected date range (for navigation)
let calSelStart  = null;  // { month, day } — range start
let calSelEnd    = null;  // { month, day } — range end
let calSelPicking = false; // true after first click, waiting for end

// Filter bar state
const TO_FILTER_MULT = { all:1.0, sunwing:0.82, tui:1.18, 'thomas-cook':0.71, 'club-med':1.08 };
let calFiltTO = 'all';
let calCompareMode = 'none'; // 'ly', 'stly', 'fcst', 'budget', 'none'
function calSetCompare(val) {
  calCompareMode = val || 'ly';
  renderCalendar();
}
// Disable compare dropdown when viewport is too narrow for the current view
function _calUpdateCompareState() {
  var sel = document.getElementById('calCompare');
  var wrap = document.getElementById('calCompareWrap');
  if (!sel || !wrap) return;
  var w = window.innerWidth;
  var v = calDisplayView;
  // 3-month under 2100px or 2-month under 1537px → disable
  var hide = (v === 3 && w < 2100) || (v === 2 && w < 1537);
  sel.disabled = hide;
  wrap.classList.toggle('cmp-disabled', hide);
  if (hide && calCompareMode !== 'none') {
    calCompareMode = 'none';
    sel.value = 'none';
  }
}
window.addEventListener('resize', _calUpdateCompareState);
const ALLOTMENTS = {
  sunwing:       { total: 42, pct: 0.88 },
  tui:           { total: 55, pct: 0.72 },
  'thomas-cook': { total: 30, pct: 0.58 },
  'club-med':    { total: 25, pct: 0.94 },
};

// Closed-out days
const LOCKED_DAYS = new Set(['2-1', '2-23', '3-3', '3-17', '4-8']);

// Metadata for fully-locked days (who applied the closeout and when)
const LOCKED_DAYS_META = {
  '2-1':  { appliedBy: 'Sarah M.',  appliedAt: '2026-01-14T08:47:00' },
  '2-23': { appliedBy: 'James K.',  appliedAt: '2026-02-06T15:22:00' },
  '3-3':  { appliedBy: 'Ana L.',    appliedAt: '2026-02-12T09:05:00' },
  '3-17': { appliedBy: 'Priya T.',  appliedAt: '2026-02-28T11:34:00' },
  '4-8':  { appliedBy: 'Carlos R.', appliedAt: '2026-03-21T16:58:00' },
};

// Partial closures: specific TOs, room types, board types closed out per date
// PARTIAL_CLOSURES: each day has an array of strategies.
// Each rule: { tos:[], roomTypes:[], boards:[], appliedBy:'', appliedAt:'' }
// Empty array = applies to ALL of that dimension.
// Multiple strategies on same day = independent close-out entries.
const PARTIAL_CLOSURES = {
  '2-5':  [
    { tos:['Sunshine Tours'], roomTypes:['Suite','Jr. Suite'], boards:[],         appliedBy:'Sarah M.',  appliedAt:'2026-01-22T10:14:00' },
    { tos:[],                 roomTypes:['Standard'],           boards:['ro'],    appliedBy:'James K.',  appliedAt:'2026-01-22T10:31:00' },
  ],
  '2-12': [
    { tos:[],                 roomTypes:['Standard','Deluxe'],  boards:['ro','hb'], appliedBy:'Priya T.',  appliedAt:'2026-01-30T14:09:00' },
  ],
  '2-18': [
    { tos:['Global Adv.','City Breaks'], roomTypes:[], boards:[],                appliedBy:'Carlos R.', appliedAt:'2026-02-03T09:52:00' },
  ],
  '2-25': [
    { tos:['Adventure'],      roomTypes:['Family'],             boards:['ro'],    appliedBy:'Ana L.',    appliedAt:'2026-02-10T16:41:00' },
    { tos:[],                 roomTypes:[],                     boards:['hb'],    appliedBy:'Ana L.',    appliedAt:'2026-02-10T16:55:00' },
  ],
  '3-4':  [
    { tos:['Sunshine Tours','Beach Hols'], roomTypes:[], boards:['ro'],           appliedBy:'Sarah M.',  appliedAt:'2026-02-11T08:30:00' },
  ],
  '3-7':  [
    { tos:[],                 roomTypes:['Suite','Jr. Suite','Family'], boards:['ai'], appliedBy:'James K.',  appliedAt:'2026-02-18T11:07:00' },
    { tos:['Global Adv.'],    roomTypes:['Standard'],           boards:[],        appliedBy:'James K.',  appliedAt:'2026-02-18T11:22:00' },
  ],
  '3-9':  [
    { tos:['Global Adv.','City Breaks'], roomTypes:['Standard'], boards:[],       appliedBy:'Priya T.',  appliedAt:'2026-02-20T13:45:00' },
    { tos:[],                 roomTypes:[],                     boards:['ro'],    appliedBy:'Carlos R.', appliedAt:'2026-02-20T14:03:00' },
  ],
  '3-11': [
    { tos:[],                 roomTypes:['Deluxe','Suite'],     boards:['ro','hb'], appliedBy:'Sarah M.',  appliedAt:'2026-02-24T09:18:00' },
  ],
  '3-13': [
    { tos:['Sunshine Tours'], roomTypes:[],                     boards:['ai','ro'], appliedBy:'Ana L.',    appliedAt:'2026-02-26T10:44:00' },
    { tos:[],                 roomTypes:['Jr. Suite'],           boards:[],        appliedBy:'Ana L.',    appliedAt:'2026-02-26T10:57:00' },
  ],
  '3-15': [
    { tos:['Beach Hols','City Breaks'], roomTypes:['Standard','Superior'], boards:['bb'], appliedBy:'James K.',  appliedAt:'2026-02-28T08:12:00' },
    { tos:['Adventure'],      roomTypes:[],                     boards:['ai'],    appliedBy:'Priya T.',  appliedAt:'2026-02-28T08:29:00' },
  ],
  '3-18': [
    { tos:['Global Adv.'],    roomTypes:['Jr. Suite'],           boards:[],       appliedBy:'Carlos R.', appliedAt:'2026-03-04T15:33:00' },
  ],
  '3-20': [
    { tos:[],                 roomTypes:['Family'],             boards:['ro','hb'], appliedBy:'Sarah M.',  appliedAt:'2026-03-06T11:20:00' },
  ],
  '3-22': [
    { tos:['Adventure'],      roomTypes:['Suite'],              boards:[],        appliedBy:'Priya T.',  appliedAt:'2026-03-09T09:05:00' },
    { tos:[],                 roomTypes:[],                     boards:['bb'],    appliedBy:'Priya T.',  appliedAt:'2026-03-09T09:17:00' },
  ],
  '3-25': [
    { tos:['Sunshine Tours'], roomTypes:['Standard'],           boards:['ro'],    appliedBy:'James K.',  appliedAt:'2026-03-10T14:48:00' },
    { tos:['Global Adv.'],    roomTypes:['Deluxe'],             boards:['ro'],    appliedBy:'James K.',  appliedAt:'2026-03-10T15:02:00' },
    { tos:[],                 roomTypes:[],                     boards:['hb'],    appliedBy:'Ana L.',    appliedAt:'2026-03-11T08:31:00' },
  ],
  '3-28': [
    { tos:['City Breaks'],    roomTypes:[],                     boards:['ai','fb'], appliedBy:'Carlos R.', appliedAt:'2026-03-13T10:19:00' },
  ],
  '4-5':  [
    { tos:['Sunshine Tours'], roomTypes:[],                     boards:['ai','ro'], appliedBy:'Sarah M.',  appliedAt:'2026-03-19T09:40:00' },
  ],
  '4-12': [
    { tos:['Beach Hols'],     roomTypes:['Suite'],              boards:[],        appliedBy:'Ana L.',    appliedAt:'2026-03-25T11:55:00' },
    { tos:['Adventure'],      roomTypes:['Family'],             boards:[],        appliedBy:'Ana L.',    appliedAt:'2026-03-25T12:08:00' },
    { tos:[],                 roomTypes:[],                     boards:['ro'],    appliedBy:'James K.',  appliedAt:'2026-03-26T08:44:00' },
  ],
  '4-17': [
    { tos:[],                 roomTypes:['Standard','Superior','Deluxe'], boards:['ro'], appliedBy:'Priya T.',  appliedAt:'2026-04-01T13:27:00' },
  ],
  '4-20': [
    { tos:['Global Adv.','City Breaks'], roomTypes:['Jr. Suite'], boards:['hb','bb'], appliedBy:'Carlos R.', appliedAt:'2026-04-03T10:06:00' },
  ],
  '4-25': [
    { tos:['Adventure'],      roomTypes:['Family'],             boards:[],        appliedBy:'Sarah M.',  appliedAt:'2026-04-08T09:33:00' },
  ],
};

// Blue online/offline dot days
const SPECIAL_DOTS = { '3-8': '#2563EB', '4-3': '#2563EB', '2-14': '#2563EB' };

// ── Calendar Events data
const CAL_EVENTS = {
  "1-1":  [{ name: "New Year Rate Launch",     type: "One-time",  date: "1/1/2026"  }],
  "1-3":  [{ name: "Weekend Opener",           type: "Recurring", date: "1/3/2026"  }],
  "1-5":  [{ name: "January Flash Sale",       type: "One-time",  date: "1/5/2026"  }],
  "1-7":  [{ name: "Midweek Special",          type: "Recurring", date: "1/7/2026"  }],
  "1-10": [{ name: "Weekend Rate Boost",       type: "Recurring", date: "1/10/2026" }],
  "1-12": [{ name: "Winter Warmup Package",    type: "One-time",  date: "1/12/2026" }],
  "1-15": [{ name: "Mid-Month Review",         type: "One-time",  date: "1/15/2026" },
            { name: "Weekend Flash Sale",      type: "Recurring", date: "1/15/2026" }],
  "1-17": [{ name: "Midweek Offer",            type: "Recurring", date: "1/17/2026" }],
  "1-19": [{ name: "Martin Luther King Rate",  type: "One-time",  date: "1/19/2026" }],
  "1-21": [{ name: "Winter Rate Push",         type: "Recurring", date: "1/21/2026" }],
  "1-24": [{ name: "Weekend Package",          type: "Recurring", date: "1/24/2026" }],
  "1-26": [{ name: "Late Jan Flash Sale",      type: "One-time",  date: "1/26/2026" }],
  "1-28": [{ name: "Midweek Promo",            type: "Recurring", date: "1/28/2026" }],
  "1-31": [{ name: "Month-End Rate Review",    type: "One-time",  date: "1/31/2026" }],
  "2-14": [{ name: "Valentine Day Promo",      type: "One-time",  date: "2/14/2026" }],
  "2-28": [{ name: "Q1 Rate Review",           type: "One-time",  date: "2/28/2026" }],
  "3-1":  [{ name: "Spring Season Launch",     type: "One-time",  date: "3/1/2026"  }],
  "3-3":  [{ name: "Weekend Flash Sale",       type: "Recurring", date: "3/3/2026"  }],
  "3-5":  [{ name: "Midweek Offer",            type: "Recurring", date: "3/5/2026"  }],
  "3-8":  [{ name: "International Womens Day", type: "One-time",  date: "3/8/2026"  },
            { name: "Weekend Rate Boost",      type: "Recurring", date: "3/8/2026"  }],
  "3-10": [{ name: "Spring Promotion",         type: "Recurring", date: "3/10/2026" }],
  "3-12": [{ name: "Midweek Special",          type: "Recurring", date: "3/12/2026" }],
  "3-15": [{ name: "Mid-Month Rate Review",    type: "One-time",  date: "3/15/2026" }],
  "3-17": [{ name: "St Patricks Day Event",    type: "One-time",  date: "3/17/2026" }],
  "3-19": [{ name: "Weekend Package",          type: "Recurring", date: "3/19/2026" }],
  "3-20": [{ name: "Spring Equinox Special",   type: "One-time",  date: "3/20/2026" }],
  "3-24": [{ name: "Midweek Offer",            type: "Recurring", date: "3/24/2026" }],
  "3-29": [{ name: "Easter Long Weekend",      type: "One-time",  date: "3/29/2026" }],
  "3-31": [{ name: "Month-End Flash Sale",     type: "One-time",  date: "3/31/2026" }],
  "4-1":  [{ name: "April Fools Flash Sale",   type: "One-time",  date: "4/1/2026"  }],
  "4-3":  [{ name: "Weekend Rate Boost",       type: "Recurring", date: "4/3/2026"  }],
  "4-5":  [{ name: "Spring Bank Holiday Rate", type: "One-time",  date: "4/5/2026"  }],
  "4-7":  [{ name: "Midweek Promo",            type: "Recurring", date: "4/7/2026"  }],
  "4-10": [{ name: "Good Friday Close",        type: "One-time",  date: "4/10/2026" }],
  "4-12": [{ name: "Easter Sunday Package",    type: "One-time",  date: "4/12/2026" }],
  "4-14": [{ name: "Post-Easter Offer",        type: "One-time",  date: "4/14/2026" }],
  "4-17": [{ name: "Weekend Package",          type: "Recurring", date: "4/17/2026" }],
  "4-20": [{ name: "Earth Day Eco Rate",       type: "One-time",  date: "4/20/2026" }],
  "4-22": [{ name: "Midweek Special",          type: "Recurring", date: "4/22/2026" }],
  "4-25": [{ name: "Weekend Flash Sale",       type: "Recurring", date: "4/25/2026" }],
  "4-26": [{ name: "Sunday Funday Package",    type: "One-time",  date: "4/26/2026" }],
  "4-28": [{ name: "Midweek Offer",            type: "Recurring", date: "4/28/2026" }],
  "4-29": [{ name: "Late April Promo",         type: "One-time",  date: "4/29/2026" },
            { name: "Weekend Rate Boost",      type: "Recurring", date: "4/29/2026" }],
  "5-1":  [{ name: "May Day Special",          type: "One-time",  date: "5/1/2026"  }],
  "5-3":  [{ name: "Weekend Package",          type: "Recurring", date: "5/3/2026"  }],
  "5-5":  [{ name: "Cinco de Mayo Promo",      type: "One-time",  date: "5/5/2026"  }],
  "5-7":  [{ name: "Midweek Flash Sale",       type: "Recurring", date: "5/7/2026"  }],
  "5-10": [{ name: "Mothers Day Package",      type: "One-time",  date: "5/10/2026" }],
  "5-12": [{ name: "Spring Offer",             type: "Recurring", date: "5/12/2026" }],
};


// Total hotel capacity (rooms)
const HOTEL_CAPACITY = 210;

// Demo days: hotel is high-demand but TO rooms sold are much lower (contrast examples)
const LOW_TO_DAYS = {
  '2-7':  { hotel: 87, to: 18 },
  '2-15': { hotel: 82, to: 22 },
  '2-22': { hotel: 91, to: 14 },
  '3-5':  { hotel: 85, to: 19 },
  '3-11': { hotel: 79, to: 25 },
  '3-20': { hotel: 88, to: 12 },
  '4-2':  { hotel: 84, to: 17 },
  '4-14': { hotel: 90, to: 21 },
  '4-21': { hotel: 78, to: 16 },
};

function getOccupancy(month, day) {
  const key = `${month}-${day}`;
  if (LOW_TO_DAYS[key]) return LOW_TO_DAYS[key];
  const s = month * 31 + day;
  const hotel = 20 + Math.abs((s * 47 + 31 + s * s * 3) % 72); // 20-92%
  // TO occupancy is a subset of hotel — must not exceed it
  const to    = Math.max(5, Math.min(hotel, hotel + Math.floor((s * 17 + 7) % 21) - 10));
  return { hotel, to };
}

// Convert occupancy % to room count
function toRooms(pct) { return Math.round(HOTEL_CAPACITY * pct / 100); }

/* Map occupancy % to 10-step Heatmap Blue scale (hotel) — Figma 2026 Design System */
function getHotelClass(pct) { return ''; }

/* Heatmap removed */
function getSegClass(pct) { return ''; }

// Legacy alias
function getCellClass(hotel) { return getHotelClass(hotel); }

function getPickupPct(month, day) {
  const s = month * 31 + day;
  return Math.abs((s * 37 + 19 + s * s * 5) % 28);
}
function getGuaranteeFill(month, day) {
  const s = month * 31 + day;
  return 40 + Math.abs((s * 29 + 11) % 50);
}

let calMetric = 'occupancy';

const CAL_METRIC_DEFS = {
  hotelOcc:    { label: 'H-Occ',   color: '#5883ed', maxVal: 100,   fmt: function(v){ return v + '%'; },                        name: 'Hotel Occ',       group: 'Occupancy'  },
  toOcc:       { label: 'TO-Occ',   color: '#006461', maxVal: 100,   fmt: function(v){ return v + '%'; },                        name: 'TO Occ',           group: 'Occupancy'  },
  lyOcc:       { label: 'LY-Occ',  color: '#93c5fd', maxVal: 100,   fmt: function(v){ return v + '%'; },                        name: 'LY Occ',          group: 'Occupancy'  },
  fcstOcc:     { label: 'Fc-Occ',  color: '#fbbf24', maxVal: 100,   fmt: function(v){ return v + '%'; },                        name: 'Fcst Occ',        group: 'Occupancy'  },
  hotelAdr:    { label: 'H-ADR',   color: '#7c3aed', maxVal: 300,   fmt: function(v){ return '$' + v; },                        name: 'Hotel ADR',       group: 'ADR'        },
  toAdr:       { label: 'TO-ADR',   color: '#4f46e5', maxVal: 300,   fmt: function(v){ return '$' + v; },                        name: 'TO ADR',           group: 'ADR'        },
  lyAdr:       { label: 'LY-ADR',  color: '#c4b5fd', maxVal: 300,   fmt: function(v){ return '$' + v; },                        name: 'LY ADR',          group: 'ADR'        },
  fcstAdr:     { label: 'Fc-ADR',  color: '#fde68a', maxVal: 300,   fmt: function(v){ return '$' + v; },                        name: 'Fcst ADR',        group: 'ADR'        },
  hotelRev:    { label: 'H-Rev',   color: '#ea580c', maxVal: 50000, fmt: function(v){ return '$' + Math.round(v/1000) + 'k'; }, name: 'Hotel Revenue',   group: 'Revenue'    },
  toRev:       { label: 'TO-Rev',   color: '#b45309', maxVal: 50000, fmt: function(v){ return '$' + Math.round(v/1000) + 'k'; }, name: 'TO Revenue',       group: 'Revenue'    },
  lyRev:       { label: 'LY-Rev',  color: '#fdba74', maxVal: 50000, fmt: function(v){ return '$' + Math.round(v/1000) + 'k'; }, name: 'LY Revenue',      group: 'Revenue'    },
  fcstRev:     { label: 'Fc-Rev',  color: '#fcd34d', maxVal: 50000, fmt: function(v){ return '$' + Math.round(v/1000) + 'k'; }, name: 'Fcst Revenue',    group: 'Revenue'    },
  hotelPickup: { label: 'H-Pkp',   color: '#16a34a', maxVal: 30,    fmt: function(v){ return (v>=0?'+':'') + v; },              name: 'Hotel Pickup',    group: 'Pickup'     },
  toPickup:    { label: 'TO-Pkp',   color: '#0d9488', maxVal: 30,    fmt: function(v){ return (v>=0?'+':'') + v; },              name: 'TO Pickup',        group: 'Pickup'     },
  hotelRn:     { label: 'H-RN',    color: '#2e65e8', maxVal: 210,   fmt: function(v){ return String(v); },                      name: 'Hotel RN Sold',   group: 'RN Sold'    },
  toRn:        { label: 'TO-RN',    color: '#0284c7', maxVal: 210,   fmt: function(v){ return String(v); },                      name: "TO RN Sold",       group: 'RN Sold'    },
  hotelTrev:   { label: 'H-TRV',   color: '#9333ea', maxVal: 500,   fmt: function(v){ return '$' + v; },                        name: 'Hotel REVPAR',    group: 'REVPAR'     },
  toTrev:      { label: 'TO-TRV',   color: '#7c3aed', maxVal: 500,   fmt: function(v){ return '$' + v; },                        name: 'TO REVPAR',        group: 'REVPAR'     },
  lyRevpar:    { label: 'LY-RVP',  color: '#d8b4fe', maxVal: 500,   fmt: function(v){ return '$' + v; },                        name: 'LY REVPAR',       group: 'REVPAR'     },
  fcstRevpar:  { label: 'Fc-RVP',  color: '#fef08a', maxVal: 500,   fmt: function(v){ return '$' + v; },                        name: 'Fcst REVPAR',     group: 'REVPAR'     },
  remainRooms:  { label: 'Rem',    color: '#16a34a', maxVal: 210,   fmt: function(v){ return String(v); },                        name: 'Remaining Rooms',    group: 'Other'         },
  avgAdults:    { label: 'AdA',    color: '#2e65e8', maxVal: 4,     fmt: function(v){ return v.toFixed(1); },                     name: 'Avg Adults',         group: 'Other'         },
  avgChildren:  { label: 'AdC',   color: '#d33030', maxVal: 2,     fmt: function(v){ return v.toFixed(1); },                     name: 'Avg Children',       group: 'Other'         },
  availRooms:   { label: 'AvR',   color: '#16a34a', maxVal: 210,   fmt: function(v){ return String(v); },                        name: 'Avail Rooms',        group: 'Other'         },
  availGuar:    { label: 'AvG',   color: '#ea580c', maxVal: 30,    fmt: function(v){ return String(v); },                        name: 'Avail Guar.',        group: 'Other'         },
  avgLos:       { label: 'LOS',   color: '#0891b2', maxVal: 14,    fmt: function(v){ return v.toFixed(1) + 'n'; },               name: 'Avg LOS',            group: 'Stay Behaviour'},
  avgLeadTime:  { label: 'Lead',  color: '#6366f1', maxVal: 365,   fmt: function(v){ return v + 'd'; },                          name: 'Avg Lead Time',      group: 'Stay Behaviour'},
  bizMixTO:     { label: 'TO%',   color: '#006461', maxVal: 100,   fmt: function(v){ return v + '%'; },                          name: 'TO Mix %',           group: 'Business Mix'  },
  bizMixDirect: { label: 'Dir%',  color: '#0284c7', maxVal: 100,   fmt: function(v){ return v + '%'; },                          name: 'Direct Mix %',       group: 'Business Mix'  },
  bizMixOTA:    { label: 'OTA%',  color: '#D97706', maxVal: 100,   fmt: function(v){ return v + '%'; },                          name: 'OTA Mix %',          group: 'Business Mix'  },
  rateTO:       { label: 'TO-R',  color: '#0f766e', maxVal: 500,   fmt: function(v){ return '$' + v; },                          name: 'TO Contract Rate',   group: 'Selling Rates' },
  ratePromo:    { label: 'Prmo%', color: '#d97706', maxVal: 50,    fmt: function(v){ return v + '%'; },                          name: 'Promotion %',        group: 'Selling Rates' },
  rateBase:     { label: 'Base',  color: '#9333ea', maxVal: 500,   fmt: function(v){ return '$' + v; },                          name: 'Base Segment Rate',  group: 'Selling Rates' },
};
let calCellMetrics = ['hotelOcc', 'toOcc'];

let showForecast = false;
function getForecast(month, day) {
  var v = Math.floor(Math.abs((month * 67 + day * 43 + month * day * 3) % 22)) - 8;
  var base = getOccupancy(month, day);
  return {
    hotel: Math.max(5, Math.min(100, base.hotel + v)),
    to:    Math.max(5, Math.min(100, base.to + Math.floor(v * 0.6))),
    adr:   150 + Math.abs((month * 47 + day * 31) % 130) + v * 3,
    rev:   (base.hotel + v) * (150 + Math.abs((month * 47 + day * 31) % 130)) * HOTEL_CAPACITY / 100 * 1.1,
  };
}

let bulkSelectMode = false;
let bulkSelected = new Set();

function renderCalendar() {
  const container = document.getElementById('calMonths');
  if (!container) return;

  const visible = ALL_MONTHS.slice(calStartIdx, calStartIdx + calView);

  // Update nav range label — target the dedicated date-nav row element
  const rangeLabel = calView <= 2
    ? visible[0].name
    : `${visible[0].name.split(' ')[0]} – ${visible[visible.length-1].name}`;
  const rangeEl = document.getElementById('calRange') || document.querySelector('.cal-range');
  if (rangeEl) rangeEl.textContent = rangeLabel;
  var moRangeEl = document.getElementById('moShufRange');
  if (moRangeEl) moRangeEl.textContent = rangeLabel;

  // Grid columns
  // 12M: show 6 per row (wraps to 2 rows of 6); 6M: single row of 6
  var gridCols = calView;
  if (calDisplayView === 12) gridCols = 6;
  else if (calDisplayView === 6) gridCols = 6;
  container.style.gridTemplateColumns = 'repeat(' + gridCols + ', 1fr)';

  const DOW = ['S','M','T','W','T','F','S'];
  container.innerHTML = visible.map(m => {
    let cells = '';
    for (let i = 0; i < m.firstDay; i++) {
      cells += `<div class="cal-day empty"></div>`;
    }
    for (let d = 1; d <= m.days; d++) {
      const key     = `${m.month}-${d}`;
      const isLocked = LOCKED_DAYS.has(key);
      const isToday  = (m.month === 3 && d === 9);
      const dot      = SPECIAL_DOTS[key];
      const { hotel, to: toRaw } = getOccupancy(m.month, d);
      const toMult = (typeof TO_FILTER_MULT !== 'undefined' && TO_FILTER_MULT[calFiltTO]) ? TO_FILTER_MULT[calFiltTO] : 1;
      const to = Math.min(95, Math.round(toRaw * toMult));

      let metricVal = hotel;
      if (calMetric === 'pickup') metricVal = getPickupPct(m.month, d);
      else if (calMetric === 'guarantees') metricVal = getGuaranteeFill(m.month, d);
      const isActionNeeded = hotel >= 65 && to < 40 && !isLocked;
      const cellClass = isActionNeeded ? getSegClass(to) : getHotelClass(metricVal);

      const cellAdr = 150 + Math.abs((m.month * 47 + d * 31) % 130);
      const cellRev = Math.floor(hotel * cellAdr * HOTEL_CAPACITY / 100 * 1.1);
      const cellRnSold = Math.floor(hotel * HOTEL_CAPACITY / 100);
      const cellPickup = getPickupPct(m.month, d) - 10;
      const _pdv = (typeof pickupDayValues !== 'undefined' && pickupDayValues) || window.pickupDayValues || [1, 3, 7];
      const cellRemainRooms = HOTEL_CAPACITY - cellRnSold;
      const cellAvgAdults = parseFloat((1.8 + Math.abs((m.month * 11 + d * 7) % 3) * 0.1).toFixed(1));
      const cellAvgChildren = parseFloat((0.3 + Math.abs((m.month * 7 + d * 13) % 5) * 0.1).toFixed(1));
      const cellTrevpar = 180 + Math.abs((m.month * 53 + d * 29) % 200);
      const cellAvailGuar = Math.max(0, 5 + Math.floor(Math.abs((m.month * 7 + d * 11) % 20)));
      const toRnSold   = Math.round(HOTEL_CAPACITY * to / 100);
      const toAdrVal   = Math.max(80, cellAdr - 20 - Math.abs((m.month * 3 + d * 7) % 15));
      const toRevVal   = Math.floor(toRnSold * toAdrVal);
      const toPickupV  = Math.max(0, Math.floor(cellPickup * to / Math.max(1, hotel)));
      const toTrevVal  = Math.max(50, cellTrevpar - 30 - Math.abs((m.month * 5 + d * 3) % 20));
      const lyF   = 0.88 + Math.abs((m.month * 3 + d * 7) % 8) * 0.005;
      const fcF   = 1.04 + Math.abs((m.month * 5 + d * 11) % 6) * 0.005;
      const cellMetricVals = {
        hotelOcc: hotel, toOcc: to,
        lyOcc: Math.round(hotel * lyF), fcstOcc: Math.min(100, Math.round(hotel * fcF)),
        hotelAdr: cellAdr, toAdr: toAdrVal,
        lyAdr: Math.round(cellAdr * lyF), fcstAdr: Math.round(cellAdr * fcF),
        hotelRev: cellRev, toRev: toRevVal,
        lyRev: Math.round(cellRev * lyF), fcstRev: Math.round(cellRev * fcF),
        hotelPickup: cellPickup, toPickup: toPickupV,
        hotelPickup_0: Math.max(0, Math.round(cellPickup * (_pdv[0]<=1?0.3:_pdv[0]<=3?0.6:_pdv[0]<=7?1:Math.min(2,_pdv[0]/7)))),
        hotelPickup_1: Math.max(0, Math.round(cellPickup * (_pdv[1]<=1?0.3:_pdv[1]<=3?0.6:_pdv[1]<=7?1:Math.min(2,_pdv[1]/7)))),
        hotelPickup_2: Math.max(0, Math.round(cellPickup * (_pdv[2]<=1?0.3:_pdv[2]<=3?0.6:_pdv[2]<=7?1:Math.min(2,_pdv[2]/7)))),
        toPickup_0: Math.max(0, Math.round(toPickupV * (_pdv[0]<=1?0.3:_pdv[0]<=3?0.6:_pdv[0]<=7?1:Math.min(2,_pdv[0]/7)))),
        toPickup_1: Math.max(0, Math.round(toPickupV * (_pdv[1]<=1?0.3:_pdv[1]<=3?0.6:_pdv[1]<=7?1:Math.min(2,_pdv[1]/7)))),
        toPickup_2: Math.max(0, Math.round(toPickupV * (_pdv[2]<=1?0.3:_pdv[2]<=3?0.6:_pdv[2]<=7?1:Math.min(2,_pdv[2]/7)))),
        hotelRn: cellRnSold, toRn: toRnSold,
        hotelTrev: cellTrevpar, toTrev: toTrevVal,
        lyRevpar: Math.round(cellTrevpar * lyF), fcstRevpar: Math.round(cellTrevpar * fcF),
        remainRooms: cellRemainRooms,
        avgAdults: cellAvgAdults, avgChildren: cellAvgChildren,
        availRooms: cellRemainRooms, availGuar: cellAvailGuar,
        avgLos:       parseFloat((2.8 + Math.abs((m.month*11+d*7)%5)*0.3).toFixed(1)),
        avgLeadTime:  18 + Math.abs((m.month*13+d*11)%60),
        bizMixTO:     28 + Math.abs((m.month*7+d*5)%25),
        bizMixDirect: 30 + Math.abs((m.month*5+d*9)%20),
        bizMixOTA:    20 + Math.abs((m.month*9+d*3)%18),
        rateTO:       Math.round(cellAdr * 0.82),
        ratePromo:    5 + Math.abs((m.month*3+d*7)%18),
        rateBase:     Math.round(cellAdr * 1.08),
        totalGuests:  Math.round(hotel * HOTEL_CAPACITY / 100 * (cellAvgAdults + cellAvgChildren)),
      };
      // ── Icons (Material: apartment = hotel, confirmation_number = TO) ──
      const icoHotel = `<span class="material-icons cell-m-ico" style="font-size:10px;color:#b0b5ba">apartment</span>`;
      const icoTO    = `<span class="material-icons cell-m-ico" style="font-size:10px;color:#b0b5ba">confirmation_number</span>`;
      const lockIcoRed    = `<span class="material-icons cell-lock-ico" style="font-size:14px;color:#dc2626">lock</span>`;
      const lockIcoOrange = `<span class="material-icons cell-lock-ico" style="font-size:14px;color:#FF9800">lock</span>`;
      const eyeSvg  = `<button class="cell-eye" aria-label="Quick view" data-month="${m.month}" data-day="${d}"><span class="material-icons" style="font-size:14px">visibility</span></button>`;

      // ── Build metric rows from Cell Metrics selection ──
      const isCompact = (calDisplayView >= 3);
      const metricRows = (function() {
        if (isCompact) return '';
        // Use cmBuildRows to get what user selected in Cell Metrics panel
        var rows = (typeof window.cmBuildRows === 'function')
          ? window.cmBuildRows(cellMetricVals)
          : [
              { label: 'H-Occ', value: hotel + '%', raw: hotel, color: '#5883ed' },
              { label: 'TO-Occ', value: to + '%',    raw: to,    color: '#006461' },
            ];

        // Compare multipliers
        var _stlyF = 0.85 + Math.abs((m.month * 5 + d * 3) % 10) * 0.006;
        var _fcstF = 1.04 + Math.abs((m.month * 5 + d * 11) % 6) * 0.005;
        var _budgF = 0.95 + Math.abs((m.month * 9 + d * 4) % 8) * 0.004;
        var _cmpMult = calCompareMode === 'ly' ? lyF
          : calCompareMode === 'stly' ? _stlyF
          : calCompareMode === 'fcst' ? _fcstF
          : calCompareMode === 'budget' ? _budgF : 0;

        return rows.map(function(r) {
          // Determine if this is a Hotel (H-) or TO (TO-) metric by label prefix
          var lbl = r.label || '';
          var isTO = lbl.substring(0, 3) === 'TO-';
          var isH  = lbl.charAt(0) === 'H' && lbl.charAt(1) === '-';
          var metricColorClass = isTO ? 'cell-m-to' : 'cell-m-hotel';
          // Short label: strip H-/TO- prefix, then strip LY-/STLY-/Fcst- for cleanliness
          var shortLabel = isTO ? lbl.substring(3) : (isH ? lbl.substring(2) : lbl);
          shortLabel = shortLabel.replace(/^LY-|^STLY-|^Fcst-/, '');

          // Comparison value
          var compHtml = '';
          if (_cmpMult > 0 && r.raw !== undefined) {
            var compRaw = Math.round(r.raw * _cmpMult * 10) / 10;
            var diff = r.raw - compRaw;
            var upOrDown = diff > 0 ? 'cell-cmp-up' : diff < 0 ? 'cell-cmp-dn' : '';
            // Format comparison value same way as actual
            var compStr = r.value; // fallback
            if (typeof r.raw === 'number') {
              // Detect format from actual value string
              var v = r.value || '';
              if (v.indexOf('$') >= 0 && v.indexOf('k') >= 0) compStr = '$' + Math.round(compRaw / 1000) + 'k';
              else if (v.indexOf('$') >= 0) compStr = '$' + Math.round(compRaw);
              else if (v.indexOf('%') >= 0) compStr = Math.round(compRaw) + '%';
              else if (v.indexOf('n') >= 0) compStr = compRaw.toFixed ? compRaw.toFixed(1) + 'n' : Math.round(compRaw) + 'n';
              else if (v.indexOf('d') >= 0) compStr = Math.round(compRaw) + 'd';
              else if (v.indexOf('rn') >= 0) compStr = Math.round(compRaw) + ' rn';
              else compStr = String(Math.round(compRaw));
            }
            compHtml = '<span class="cell-cmp"> / <span class="' + upOrDown + '">' + compStr + '</span></span>';
          }

          if (r._html) return r._html;
          return '<div class="cell-m-row ' + metricColorClass + '">'
            + '<span class="cell-m-label">' + shortLabel + '</span>'
            + '<span class="cell-m-val">' + r.value + compHtml + '</span>'
            + '</div>';
        }).join('');
      })();

      const calCl = PARTIAL_CLOSURES[m.month + '-' + d];
      const hasCalCl = !isLocked && calCl && Array.isArray(calCl) && calCl.length > 0;
      const hasCalEvents = (typeof CAL_EVENTS !== 'undefined') && CAL_EVENTS[m.month + '-' + d] && CAL_EVENTS[m.month + '-' + d].length > 0;
      const _isStopSalesActive = typeof window.hmIsStopSales === 'function' && window.hmIsStopSales();
      const isBulkSel = bulkSelectMode && isLocked && bulkSelected.has(key);
      const isInRange = calSelStart && calSelEnd && (function(){
        var s = calSelStart, e = calSelEnd;
        if (s.month === e.month && s.day === e.day) return false;
        var before = s.month < m.month || (s.month === m.month && s.day <= d);
        var after = e.month > m.month || (e.month === m.month && e.day >= d);
        return before && after;
      })();
      const hmDayData = {
        hotel: hotel, to: to,
        remainRooms: cellMetricVals ? (cellMetricVals.remainRooms || 0) : 0,
        totalGuests: cellMetricVals ? (cellMetricVals.totalGuests || 0) : 0,
        isFullClose: isLocked,
        hasPartialClose: isActionNeeded,
        closureRules: calCl || [],
        toOtb: to * 1.8,
        toFcst: to * 1.6 + Math.abs((m.month * 7 + d * 11) % 15)
      };
      const hmClass = (typeof window.hmGetCellClass === 'function') ? window.hmGetCellClass(hmDayData) : '';
      const classes = ['cal-day', cellClass, hmClass, isLocked ? 'locked' : '', isToday ? 'today' : '', isActionNeeded ? 'action-needed' : '', bulkSelectMode && isLocked ? 'bulk-selectable' : '', isBulkSel ? 'bulk-sel' : '', isInRange ? 'in-range' : ''].filter(Boolean).join(' ');

      const hotelRooms = toRooms(hotel);
      const toRoomsSold = toRooms(to);
      const capTipAttr = isLocked ? '' : ` onmouseenter="calShowCapTip(event,${hotel},${hotelRooms},${to},${toRoomsSold},${210-hotelRooms-toRoomsSold},${m.month},${d})" onmouseleave="calHideCapTip()"`;
      const moIso = `${m.year}-${String(m.month).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
      const moChk = _moSelectedDays.has(moIso) ? ' checked' : '';
      const _showEye = !isCompact;
      cells += `<div class="${classes}" data-month="${m.month}" data-day="${d}"${capTipAttr}>
        <div class="cell-day-hdr">
          <input type="checkbox" class="wv-day-chk mo-day-chk"${moChk} onclick="event.stopPropagation();moDayCheck('${moIso}',this)" title="Select for close-out">
          ${(isLocked || hasCalCl) ? '<span class="mo-lock-ico" style="cursor:pointer" onmouseenter="calShowEventTip(event,\'' + m.month + '-' + d + '\')" onmouseleave="calHideEventTip()">' + (isLocked ? lockIcoRed : lockIcoOrange) + '</span>' : ''}
          <span class="cell-hdr-left"><span class="day-num">${d}</span></span>
          ${_showEye ? eyeSvg : ''}
        </div>
        ${!isCompact ? `<div class="cell-content">${metricRows}</div>` : ''}
        ${!isCompact && hasCalEvents ? '<span class="cell-event-ico" onmouseenter="calShowEventTip(event,\''+m.month+'-'+d+'\')" onmouseleave="calHideEventTip()"><span class="material-icons" style="font-size:16px;color:#006461">today</span></span>' : ''}
      </div>`;
    }

    const s = m.stats;

    // Monthly summary data (seeded per month)
    const mSeed = m.month * 17;
    const mOcc   = 58 + mSeed % 28;
    const mAdr   = 142 + mSeed % 60;
    const mRev   = Math.round(mOcc * mAdr * 210 / 100 * m.days / 1000);
    const mRn    = Math.round(mOcc * 210 / 100 * m.days);
    const mRevpar= Math.round(mAdr * mOcc / 100);
    const mPickup= 8 + mSeed % 18;
    // Avg Adults / Avg Children
    const mAvgA  = (1.8 + mSeed%3*0.1).toFixed(1);
    const mAvgC  = (0.3 + mSeed%2*0.1).toFixed(1);
    // Avail rooms/guar
    const mAvailR= 210 - Math.round(mOcc*210/100);
    const mAvailG= 8 + mSeed%12;
    // TO RN sold
    const mToOcc = 28 + mSeed % 20;
    const mToRn  = Math.round(mToOcc * 210 / 100 * m.days);
    const mToAdr = Math.round(mAdr * 0.88);
    const mToRev = Math.round(mToRn * mToAdr / 1000);
    // Meal plan breakdown
    const mpAI   = Math.round(mRn * 0.52);
    const mpHB   = Math.round(mRn * 0.24);
    const mpBB   = Math.round(mRn * 0.14);
    const mpRO   = mRn - mpAI - mpHB - mpBB;
    // Room avail by type
    const avStd  = Math.round(mAvailR * 0.40);
    const avDel  = Math.round(mAvailR * 0.22);
    const avSte  = Math.round(mAvailR * 0.18);
    const avFam  = mAvailR - avStd - avDel - avSte;

    // Compute locked day counts from live data
    const _mPrefix = m.month + '-';
    const _actualFullLocked = Array.from(LOCKED_DAYS).filter(k => k.startsWith(_mPrefix)).length;
    const _actualPartLocked = Object.keys(PARTIAL_CLOSURES).filter(k => k.startsWith(_mPrefix) && !LOCKED_DAYS.has(k)).length;
    const _totalLocked = _actualFullLocked + _actualPartLocked;

    return `
      <div class="cal-month">
        <div class="cal-month-hdr">
          <span class="cal-month-name">${m.name}</span>
          ${!(typeof window.hmIsStopSales === 'function' && window.hmIsStopSales()) && _totalLocked > 0 ? `<span class="cal-lock-badge"><span class="material-icons" style="font-size:12px">lock</span>${_totalLocked}</span>` : ''}
        </div>
        <div class="cal-dow">${DOW.map(d => `<span>${d}</span>`).join('')}</div>
        <div class="cal-days">${cells}</div>
        
        <div class="cal-month-summary">
        </div>
      </div>`;
  }).join('');

  // Re-attach popup listeners after re-render
  const calMonths = document.getElementById('calMonths');
  if (calMonths) calMonths._popupBound = false; // reset so popup IIFE re-binds below

  // Re-apply any active range selection
  applyCalSelection();
  // Monthly summary (1M/2M/3M)
  renderCalMonthlySummary();
}

// Keep old name for legacy call at bottom
function buildCalendar() { renderCalendar(); }

// ── Monthly summary accordion state ──────────────────────────────────────────
var _calAccState = { daily: false, more: false, meals: false, biz: false, tc: false, overview: true };

window.calAccClick = function(hdr) {
  var sect = hdr.closest('.wv-acc-sect');
  var body = hdr.nextElementSibling;
  if (!sect || !body) return;
  var isOpen = sect.classList.contains('wv-acc-open');
  // Toggle all sections with same data-cal-section (multi-month view has duplicates)
  var key = hdr.dataset.calSection;
  _calAccState[key] = isOpen; // isOpen means it was open and is now closing
  document.querySelectorAll('.cal-summary-wrap .wv-acc-hdr[data-cal-section="' + key + '"]').forEach(function(h) {
    var s = h.closest('.wv-acc-sect');
    var b = h.nextElementSibling;
    if (s) s.classList.toggle('wv-acc-open', !isOpen);
    if (b) b.classList.toggle('wv-body-hidden', isOpen);
    // Rotate chevron
    var chev = h.querySelector('.wv-acc-chev svg');
    if (chev) chev.style.transform = isOpen ? '' : 'rotate(180deg)';
  });
};

// ── Monthly Summary Metrics (1M / 2M / 3M) ─────────────────────────────────
function renderCalMonthlySummary() {
  var el = document.getElementById('calMonthlySummary');
  if (!el) return;
  if (calDisplayView > 3) { el.style.display = 'none'; return; }
  el.style.display = 'block';

  var visible = ALL_MONTHS.slice(calStartIdx, calStartIdx + calView);
  var isSingle = calView === 1;
  var WV = 250;

  // ── Helpers ────────────────────────────────────────────────────────────
  function fR(v){ return v>=1000000?'$'+(v/1000000).toFixed(1)+'M':'$'+Math.round(v/1000)+'k'; }

  function dualBar(tPct, hPct, clr) {
    return '<div style="height:3px;border-radius:2px;margin-top:3px;background:#e5e7eb;position:relative">'
      +(hPct!=null?'<div style="position:absolute;top:0;left:0;height:100%;width:'+Math.min(92,hPct)+'%;background:#d1d5db;border-radius:2px"></div>':'')
      +'<div style="position:absolute;top:0;left:0;height:100%;width:'+Math.min(92,tPct)+'%;background:'+clr+';border-radius:2px"></div>'
      +'</div>';
  }
  function sBar(segs) {
    return '<div style="height:5px;background:#e5e7eb;border-radius:3px;display:flex;overflow:hidden;margin:3px 0">'
      +segs.map(function(s){return '<div style="width:'+s.p+'%;background:'+s.c+'"></div>';}).join('')
      +'</div>';
  }
  function refChips(pairs) {
    var CSS={stly:'background:#e0e7ff;color:#4338ca',ly:'background:#dcfce7;color:#15803d',fcst:'background:#fef9c3;color:#a16207'};
    return '<div style="display:flex;gap:2px;flex-wrap:wrap;margin-top:2px">'
      +pairs.map(function(p){return '<span style="font-size:7px;font-weight:700;padding:1px 4px;border-radius:3px;'+CSS[p.k]+'">'+p.l+' '+p.v+'</span>';}).join('')
      +'</div>';
  }
  function mRow(lbl, tVal, hVal, tPct, hPct, clr, refs) {
    return '<div style="display:flex;align-items:center;gap:3px;margin-bottom:4px">'
      +'<span style="font-size:8px;color:#6b7280;flex:1;min-width:0">'+lbl+'</span>'
      +(hVal?'<span style="display:flex;flex-direction:column;align-items:flex-end">'
        +'<span style="font-size:6px;font-weight:700;color:#9ca3af;text-transform:uppercase">Hotel</span>'
        +'<span style="font-size:8px;color:#6b7280">'+hVal+'</span></span>':'')
      +'<span style="display:flex;flex-direction:column;align-items:flex-end;margin-left:6px">'
      +'<span style="font-size:6px;font-weight:700;color:'+clr+';text-transform:uppercase">TO</span>'
      +'<span style="font-size:9px;font-weight:800;color:'+clr+'">'+tVal+'</span></span>'
      +'</div>'
      +(tPct!=null?dualBar(tPct,hPct,clr):'')
      +(refs?refChips(refs):'');
  }
  function colHdr(clr) {
    return '<div style="display:flex;justify-content:flex-end;gap:12px;margin-bottom:5px;padding-bottom:3px;border-bottom:1px solid #f3f4f6">'
      +'<span style="font-size:6.5px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.3px">Hotel</span>'
      +'<span style="font-size:6.5px;font-weight:700;color:'+(clr||'#006461')+';text-transform:uppercase;letter-spacing:.3px;min-width:24px;text-align:right">TO</span>'
      +'</div>';
  }
  var chevUp   = '<span class="material-icons" style="font-size:16px">expand_less</span>';
  var chevDown = '<span class="material-icons" style="font-size:16px">expand_more</span>';
  function sec(title, key, content) {
    var collapsed = _calAccState[key] !== false ? !!_calAccState[key] : false;
    return '<div class="wv-acc-sect' + (collapsed ? '' : ' wv-acc-open') + '">'
      + '<div class="wv-acc-hdr" data-cal-section="' + key + '" onclick="calAccClick(this)">'
      + '<span class="wv-acc-chev" style="color:#006461">' + (collapsed ? chevDown : chevUp) + '</span>'
      + '<span class="wv-acc-title">' + title + '</span>'
      + '</div>'
      + '<div class="wv-acc-body' + (collapsed ? ' wv-body-hidden' : '') + '" style="padding:8px 10px 6px">'
      + content
      + '</div></div>';
  }

  // ── Aggregate per-month ─────────────────────────────────────────────────
  var months = visible.map(function(m) {
    var nd = m.days, nm = m.month;
    var sumHotel=0,sumTo=0,sumAdr=0,sumToAdr=0,sumRev=0,sumHRev=0;
    var sumPickup=0,sumHPickup=0,sumRevpar=0,sumHRevpar=0;
    var sumRn=0,sumHRn=0;
    var sumLos=0,sumHLos=0,sumLead=0,sumHLead=0;
    var sumTotA=0,sumTotC=0,sumHTotA=0,sumHTotC=0;
    var sumAvailRooms=0,sumAvailGuar=0;
    var sumAi=0,sumBb=0,sumHb=0;
    var sumToMix=0,sumDirMix=0,sumOtaMix=0;
    var sumFit=0,sumDyn=0,sumSer=0,sumOther=0,sumFree=0;
    var sumOnline=0;
    var tcRateSum=[0,0,0,0,0];
    for (var d = 1; d <= nd; d++) {
      var hh = getOccupancy(nm, d);
      var hotel=hh.hotel, to=hh.to;
      var adr=150+Math.abs((nm*47+d*31)%130);
      var v=Math.abs((nm*127+d*53+nm*d*7+d*d*3))%100;
      var toAdr=Math.max(80,adr-20-Math.abs((nm*3+d*7)%15));
      var toRn=Math.round(WV*to/100), hnRn=Math.round(WV*hotel/100);
      var avgA=(1.8+v%3*0.1), avgC=(0.3+v%2*0.1);
      sumHotel+=hotel; sumTo+=to;
      sumAdr+=adr; sumToAdr+=toAdr;
      sumRev+=Math.floor(toRn*toAdr); sumHRev+=Math.floor(hnRn*adr);
      sumRn+=toRn; sumHRn+=hnRn;
      sumPickup+=Math.max(0,Math.floor((v%25+5)*to/Math.max(1,hotel)));
      sumHPickup+=Math.floor(v%25+5);
      sumRevpar+=Math.max(50,(adr+80)-30-Math.abs((nm*5+d*3)%20));
      sumHRevpar+=adr+80;
      sumLos+=2.8+v%5*0.3; sumHLos+=2.8+v%5*0.3+0.4;
      sumLead+=18+v%60; sumHLead+=18+v%60+12;
      sumTotA+=Math.round(toRn*avgA); sumTotC+=Math.round(toRn*avgC);
      sumHTotA+=Math.round(hnRn*(avgA+0.3)); sumHTotC+=Math.round(hnRn*(avgC+0.1));
      sumAvailRooms+=Math.max(0,102-Math.floor(hotel*1.02));
      sumAvailGuar+=Math.floor(8+v%5);
      sumAi+=Math.max(45,Math.min(68,55+(nm*7+d*3)%14));
      sumBb+=Math.max(14,Math.min(28,20+(nm*11+d*5)%11));
      sumHb+=Math.max(6,Math.min(16,10+(nm*5+d*7)%9));
      sumToMix+=28+Math.abs((nm*7+d*5)%25);
      sumDirMix+=30+Math.abs((nm*5+d*9)%20);
      sumOtaMix+=20+Math.abs((nm*9+d*3)%18);
      var fitP=Math.round(to*0.45),dynP=Math.round(to*0.35),serP=to-fitP-dynP;
      sumFit+=fitP; sumDyn+=dynP; sumSer+=serP;
      sumOther+=Math.max(0,hotel-to); sumFree+=Math.max(0,100-hotel);
      sumOnline+=Math.max(30,Math.min(80,45+Math.abs((nm*13+d*7)%35)));
      for(var ii=0;ii<5;ii++) tcRateSum[ii]+=adr-15+Math.abs((nm*(ii+3)+d*(ii+5))%50);
    }
    var n=nd;
    var avgH=Math.round(sumHotel/n), avgT=Math.round(sumTo/n);
    var avgAdr=Math.round(sumAdr/n), avgToAdr2=Math.round(sumToAdr/n);
    var avgRev=fR(Math.round(sumRev/n)), avgHRev=fR(Math.round(sumHRev/n));
    var totalRev=fR(sumRev), totalHRev=fR(sumHRev);
    var avgRevpar=Math.round(sumRevpar/n), avgHRevpar=Math.round(sumHRevpar/n);
    var avgPickup=Math.round(sumPickup/n), avgHPickup=Math.round(sumHPickup/n);
    var avgRn=Math.round(sumRn/n), avgHRn=Math.round(sumHRn/n);
    var avgLos=(sumLos/n).toFixed(1)+'n', avgHLos=(sumHLos/n).toFixed(1)+'n';
    var avgLead=Math.round(sumLead/n)+'d', avgHLead=Math.round(sumHLead/n)+'d';
    var avgAvailRooms=Math.round(sumAvailRooms/n), avgAvailGuar=Math.round(sumAvailGuar/n);
    var totA=Math.round(sumTotA/n), totC=Math.round(sumTotC/n);
    var hTotA=Math.round(sumHTotA/n), hTotC=Math.round(sumHTotC/n);
    var avgA=(sumTotA/Math.max(1,sumRn)).toFixed(1), avgC=(sumTotC/Math.max(1,sumRn)).toFixed(1);
    var hAvgA=(sumHTotA/Math.max(1,sumHRn)).toFixed(1), hAvgC=(sumHTotC/Math.max(1,sumHRn)).toFixed(1);
    var totG=Math.round((sumTotA+sumTotC)/n), hTotG=Math.round((sumHTotA+sumHTotC)/n);
    var avgAi=Math.round(sumAi/n), avgBb=Math.round(sumBb/n), avgHb=Math.round(sumHb/n);
    var avgRo=100-avgAi-avgBb-avgHb;
    var toPct2=avgT/Math.max(1,avgH);
    var aiTo=Math.max(0,Math.round(avgAi*toPct2*0.9)), bbTo=Math.max(0,Math.round(avgBb*toPct2*0.85));
    var hbTo=Math.max(0,Math.round(avgHb*toPct2*0.8)), roTo=Math.max(0,Math.round(avgRo*toPct2*0.95));
    var avgToMix2=Math.round(sumToMix/n), avgDirMix2=Math.round(sumDirMix/n), avgOtaMix2=Math.round(sumOtaMix/n);
    var avgOtherMix=Math.max(0,100-avgToMix2-avgDirMix2-avgOtaMix2);
    var avgFit=Math.round(sumFit/n), avgDyn=Math.round(sumDyn/n), avgSer=Math.round(sumSer/n);
    var avgOtherSeg=Math.round(sumOther/n), avgFree=Math.round(sumFree/n);
    var fitRm=Math.round(WV*avgFit/100), dynRm=Math.round(WV*avgDyn/100), serRm=Math.round(WV*avgSer/100);
    var othRm=Math.round(WV*avgOtherSeg/100), freeRm=Math.round(WV*avgFree/100);
    var avgOnline=Math.round(sumOnline/n);
    var tcRates=tcRateSum.map(function(s){return Math.round(s/n);});
    var baseRate=avgAdr+8;
    // STLY/LY/Fcst
    var sdlyT=Math.max(5,avgT-9), lyT=Math.max(5,avgT-6), fcstT=Math.min(100,avgT+4);
    var sdlyAdr=avgToAdr2-8, lyAdr=avgToAdr2-4, fcstAdr=avgToAdr2+6;
    var sdlyRev=fR(Math.floor(sumRev*0.9/n)), lyRev=fR(Math.floor(sumRev*0.95/n)), fcstRev=fR(Math.floor(sumRev*1.06/n));
    var sdlyRevpar=Math.max(40,avgRevpar-8), lyRevpar=Math.max(40,avgRevpar-4);
    var sdlyRn=Math.round(avgRn*0.88), lyRn=Math.round(avgRn*0.93), fcstRn=Math.round(avgRn*1.06);
    var isEbb=(new Date(2026,nm-1,1)).getDay()<3;
    // ── Close-out counts for this month ──
    var fullCoCount=0, partCoCount=0;
    var _coRtSet=new Set(), _coBdSet=new Set(), _coToSet=new Set();
    for (var _cd=1; _cd<=nd; _cd++) {
      var _cKey=nm+'-'+_cd;
      if (LOCKED_DAYS.has(_cKey)) fullCoCount++;
      var _pRules=PARTIAL_CLOSURES[_cKey];
      if (_pRules && _pRules.length>0) {
        partCoCount++;
        _pRules.forEach(function(r){
          (r.roomTypes||[]).forEach(function(rt){ _coRtSet.add(rt); });
          (r.boards||[]).forEach(function(b){ _coBdSet.add(b.toUpperCase()); });
          (r.tos||[]).forEach(function(t){ _coToSet.add(t); });
        });
      }
    }
    var coRoomTypes=Array.from(_coRtSet), coBoards=Array.from(_coBdSet), coTOs=Array.from(_coToSet);
    return {name:m.name,avgH,avgT,avgAdr,avgToAdr:avgToAdr2,avgRev,avgHRev,totalRev,totalHRev,
      avgRevpar,avgHRevpar,avgPickup,avgHPickup,avgRn,avgHRn,
      avgLos,avgHLos,avgLead,avgHLead,totA,totC,hTotA,hTotC,avgA,avgC,hAvgA,hAvgC,totG,hTotG,
      avgAvailRooms,avgAvailGuar,
      avgAi,avgBb,avgHb,avgRo,aiTo,bbTo,hbTo,roTo,
      avgToMix:avgToMix2,avgDirMix:avgDirMix2,avgOtaMix:avgOtaMix2,avgOtherMix,
      avgFit,avgDyn,avgSer,avgOtherSeg,avgFree,fitRm,dynRm,serRm,othRm,freeRm,avgOnline,
      tcRates,baseRate,isEbb,
      sdlyT,lyT,fcstT,sdlyAdr,lyAdr,fcstAdr,sdlyRev,lyRev,fcstRev,
      sdlyRevpar,lyRevpar,sdlyRn,lyRn,fcstRn,
      fullCoCount,partCoCount,nd,coRoomTypes,coBoards,coTOs};
  });

  var tcOps=[['Sunshine Tours','#3b82f6'],['Global Adv.','#967EF3'],['Beach Hols','#0ea5e9'],['City Breaks','#10b981'],['Adventure','#f59e0b']];

  // ── Build Daily-B style grid ──────────────────────────────────────────
  // Reuse wb- classes from the weekly view but with month columns

  function moGrad(clr) {
    if (clr==='#004948') return 'linear-gradient(to right,#004948,#007a75)';
    if (clr==='#52d9ce') return 'linear-gradient(to right,#52d9ce,#8aeee8)';
    if (clr==='#006461') return 'linear-gradient(to right,#006461,#009c96)';
    if (clr==='#0891b2') return 'linear-gradient(to right,#0891b2,#22d3ee)';
    if (clr==='#6366f1') return 'linear-gradient(to right,#6366f1,#818cf8)';
    if (clr==='#5883ed') return 'linear-gradient(to right,#5883ed,#93b4f6)';
    if (clr==='#D97706') return 'linear-gradient(to right,#D97706,#F59E0B)';
    if (clr==='#967EF3') return 'linear-gradient(to right,#967EF3,#a78bfa)';
    if (clr==='#3b82f6') return 'linear-gradient(to right,#3b82f6,#60a5fa)';
    if (clr==='#f59e0b') return 'linear-gradient(to right,#f59e0b,#fbbf24)';
    if (clr==='#0284c7') return 'linear-gradient(to right,#0284c7,#38bdf8)';
    if (clr==='#16a34a') return 'linear-gradient(to right,#16a34a,#22c55e)';
    if (clr==='#9333ea') return 'linear-gradient(to right,#9333ea,#a855f7)';
    if (clr==='#10b981') return 'linear-gradient(to right,#10b981,#34d399)';
    if (clr==='#0ea5e9') return 'linear-gradient(to right,#0ea5e9,#38bdf8)';
    if (clr==='#d33030') return 'linear-gradient(to right,#d33030,#ef4444)';
    return clr;
  }
  function moBar(pct, clr) {
    return '<div class="wv-occ-bar-track"><div style="width:'+pct+'%;background:'+moGrad(clr)+';height:6px"></div></div>';
  }
  function moStackBar(segs) {
    return '<div class="wv-occ-bar-track">'
      + segs.map(function(s){ return '<div style="width:'+s.p+'%;background:'+moGrad(s.c)+';height:6px"></div>'; }).join('')
      + '</div>';
  }

  // Collapse state — reuse _calAccState with 'mo_' prefix for groups
  if (!_calAccState._moInit) {
    _calAccState._moInit = true;
    _calAccState.mo_daily = false;
    _calAccState.mo_more = false;
    _calAccState.mo_meals = false;
    _calAccState.mo_biz = false;
    _calAccState.mo_tc = false;
    _calAccState.mo_seg = false;
    _calAccState.mo_closeouts = false;
  }

  // Row definitions (same pattern as Daily B)
  var moRows = [];
  // ── Close Outs group (top)
  moRows.push({type:'top', id:'mo_closeouts', label:'Close Outs'});
  moRows.push({type:'sect', id:'mos_co_full', label:'Full Close Out', parent:'mo_closeouts'});
  moRows.push({type:'sect', id:'mos_co_part', label:'Partial Lock', parent:'mo_closeouts'});
  moRows.push({type:'sub', id:'mos_co_rooms', label:'Room Types', dot:'#fca5a5', parent:'mos_co_part', gp:'mo_closeouts'});
  moRows.push({type:'sub', id:'mos_co_boards', label:'Board Types', dot:'#fde68a', parent:'mos_co_part', gp:'mo_closeouts'});
  moRows.push({type:'sub', id:'mos_co_tos', label:'Tour Operators', dot:'#d8b4fe', parent:'mos_co_part', gp:'mo_closeouts'});

  // ── Daily Metrics group
  moRows.push({type:'top', id:'mo_daily', label:'Daily Metrics'});
  moRows.push({type:'sect', id:'mos_occ', label:'Occupancy', parent:'mo_daily'});
  moRows.push({type:'sub', id:'mos_occ_to', label:'TO', dot:'#004948', parent:'mos_occ', gp:'mo_daily'});
  moRows.push({type:'sub', id:'mos_occ_htl', label:'Hotel', dot:'#52d9ce', parent:'mos_occ', gp:'mo_daily'});
  moRows.push({type:'sub', id:'mos_occ_stly', label:'STLY', dot:'#818cf8', parent:'mos_occ', gp:'mo_daily'});
  moRows.push({type:'sect', id:'mos_adr', label:'ADR', parent:'mo_daily'});
  moRows.push({type:'sub', id:'mos_adr_to', label:'TO ADR', dot:'#004948', parent:'mos_adr', gp:'mo_daily'});
  moRows.push({type:'sub', id:'mos_adr_htl', label:'Hotel ADR', dot:'#52d9ce', parent:'mos_adr', gp:'mo_daily'});
  moRows.push({type:'sect', id:'mos_rev', label:'Revenue /day', parent:'mo_daily'});
  moRows.push({type:'sub', id:'mos_rev_to', label:'TO Revenue', dot:'#004948', parent:'mos_rev', gp:'mo_daily'});
  moRows.push({type:'sub', id:'mos_rev_htl', label:'Hotel Revenue', dot:'#52d9ce', parent:'mos_rev', gp:'mo_daily'});
  moRows.push({type:'sect', id:'mos_revpar', label:'REVPAR', parent:'mo_daily'});
  moRows.push({type:'sub', id:'mos_revpar_stly', label:'STLY', dot:'#818cf8', parent:'mos_revpar', gp:'mo_daily'});
  moRows.push({type:'sect', id:'mos_pickup', label:'Pickup /day', parent:'mo_daily'});
  moRows.push({type:'sect', id:'mos_onoff', label:'Online / Offline', parent:'mo_daily'});
  moRows.push({type:'sub', id:'mos_onoff_on', label:'Online', dot:'#3b82f6', parent:'mos_onoff', gp:'mo_daily'});
  moRows.push({type:'sub', id:'mos_onoff_off', label:'Offline', dot:'#f97316', parent:'mos_onoff', gp:'mo_daily'});

  // ── Segments group
  moRows.push({type:'top', id:'mo_seg', label:'Segment Mix'});
  moRows.push({type:'sect', id:'mos_segbar', label:'Summary', parent:'mo_seg'});
  moRows.push({type:'sub', id:'mos_seg_fit', label:'Static FIT', dot:'#006461', parent:'mos_segbar', gp:'mo_seg'});
  moRows.push({type:'sub', id:'mos_seg_dyn', label:'TO Dynamic', dot:'#0891b2', parent:'mos_segbar', gp:'mo_seg'});
  moRows.push({type:'sub', id:'mos_seg_ser', label:'Tour Series', dot:'#6366f1', parent:'mos_segbar', gp:'mo_seg'});
  moRows.push({type:'sub', id:'mos_seg_oth', label:'Other Segs', dot:'#5883ed', parent:'mos_segbar', gp:'mo_seg'});
  moRows.push({type:'sub', id:'mos_seg_rem', label:'Remaining', dot:'#9ca3af', parent:'mos_segbar', gp:'mo_seg', isRem:true});

  // ── More Metrics group
  moRows.push({type:'top', id:'mo_more', label:'More Metrics'});
  moRows.push({type:'sect', id:'mos_rn', label:'RN Sold /day', parent:'mo_more'});
  moRows.push({type:'sub', id:'mos_rn_stly', label:'STLY', dot:'#818cf8', parent:'mos_rn', gp:'mo_more'});
  moRows.push({type:'sect', id:'mos_avga', label:'Avg Adults', parent:'mo_more'});
  moRows.push({type:'sect', id:'mos_avgc', label:'Avg Children', parent:'mo_more'});
  moRows.push({type:'sect', id:'mos_tota', label:'Total Adults', parent:'mo_more'});
  moRows.push({type:'sect', id:'mos_totc', label:'Total Children', parent:'mo_more'});
  moRows.push({type:'sect', id:'mos_totg', label:'Total Guests', parent:'mo_more'});
  moRows.push({type:'sect', id:'mos_los', label:'Avg LOS', parent:'mo_more'});
  moRows.push({type:'sect', id:'mos_lead', label:'Lead Time', parent:'mo_more'});
  moRows.push({type:'sect', id:'mos_avail', label:'Avail Rooms', parent:'mo_more'});
  moRows.push({type:'sect', id:'mos_availg', label:'Avail Guar.', parent:'mo_more'});

  // ── Meal Plans group
  moRows.push({type:'top', id:'mo_meals', label:'Meal Plans'});
  moRows.push({type:'sect', id:'mos_mpsum', label:'Summary', parent:'mo_meals'});
  moRows.push({type:'sub', id:'mos_mp_ai', label:'All Inclusive', dot:'#006461', parent:'mos_mpsum', gp:'mo_meals'});
  moRows.push({type:'sub', id:'mos_mp_bb', label:'Bed & Breakfast', dot:'#3b82f6', parent:'mos_mpsum', gp:'mo_meals'});
  moRows.push({type:'sub', id:'mos_mp_hb', label:'Half Board', dot:'#967EF3', parent:'mos_mpsum', gp:'mo_meals'});
  moRows.push({type:'sub', id:'mos_mp_ro', label:'Room Only', dot:'#f59e0b', parent:'mos_mpsum', gp:'mo_meals'});

  // ── Business Mix group
  moRows.push({type:'top', id:'mo_biz', label:'Business Mix'});
  moRows.push({type:'sect', id:'mos_bizbar', label:'Summary', parent:'mo_biz'});
  moRows.push({type:'sub', id:'mos_biz_to', label:'TO', dot:'#006461', parent:'mos_bizbar', gp:'mo_biz'});
  moRows.push({type:'sub', id:'mos_biz_dir', label:'Direct', dot:'#0284c7', parent:'mos_bizbar', gp:'mo_biz'});
  moRows.push({type:'sub', id:'mos_biz_ota', label:'OTA', dot:'#D97706', parent:'mos_bizbar', gp:'mo_biz'});
  moRows.push({type:'sub', id:'mos_biz_oth', label:'Other', dot:'#9ca3af', parent:'mos_bizbar', gp:'mo_biz'});

  // ── Travel Co. Rates group
  moRows.push({type:'top', id:'mo_tc', label:'Travel Co. Rates'});
  tcOps.forEach(function(op,i){
    moRows.push({type:'sect', id:'mos_tc'+i, label:op[0], parent:'mo_tc', toIdx:i, toClr:op[1]});
  });
  moRows.push({type:'sect', id:'mos_tcbase', label:'Base Seg. Rate', parent:'mo_tc', toBase:true});

  // Helper: check if row is hidden by collapsed parent
  function moIsHidden(row) {
    if (row.type === 'top') return false;
    // Group collapsed?
    if (row.type === 'sect' && _calAccState[row.parent]) return true;
    if (row.type === 'sub') {
      if (_calAccState[row.gp]) return true;
      if (_calAccState[row.parent]) return true;
    }
    return false;
  }

  var html = '<div class="wb-layout">';

  // ── Header row with month names ─────────────────────────────────────────
  html += '<div class="wb-row wb-hdr-row">';
  html += '<div class="wb-label-cell wb-hdr-label-cell"></div>';
  months.forEach(function(mo) {
    html += '<div class="wb-data-cell wb-hdr-cell"><span class="wb-hdr-dow">'+mo.name+'</span></div>';
  });
  html += '</div>';

  // ── Data rows ───────────────────────────────────────────────────────────
  moRows.forEach(function(row) {
    var collapsed = !!_calAccState[row.id];
    var hidden = moIsHidden(row);
    var rowCls = 'wb-row wb-row-' + row.type + (hidden ? ' wb-row-hidden' : '');

    html += '<div class="' + rowCls + '" data-mo-id="' + row.id + '">';

    // ── Label cell
    if (row.type === 'top') {
      html += '<div class="wb-label-cell wb-grp-hdr" onclick="moToggle(\'' + row.id + '\')">'
            + '<span class="wb-chev">' + (collapsed ? chevDown : chevUp) + '</span>'
            + '<span class="wb-grp-label">' + row.label + '</span></div>';
    } else if (row.type === 'sect') {
      html += '<div class="wb-label-cell wb-sect-lbl" onclick="moToggle(\'' + row.id + '\')">'
            + '<span class="wb-chev">' + (collapsed ? chevDown : chevUp) + '</span>'
            + '<span class="wb-sect-label">' + row.label + '</span></div>';
    } else {
      var dotHtml = row.dot ? '<span class="wb-sub-dot" style="background:' + row.dot + '"></span>' : '';
      html += '<div class="wb-label-cell wb-sub-lbl-cell">'
            + dotHtml
            + '<span class="wb-sub-label' + (row.isRem ? ' wb-sub-lbl-rem' : '') + '">' + row.label + '</span></div>';
    }

    // ── Data cells (one per month)
    months.forEach(function(mo) {
      var cc = '';

      if (row.type === 'top') {
        cc = '';  // group header — empty data cells

      } else if (row.type === 'sect') {
        switch (row.id) {
          case 'mos_occ':
            cc = '<div class="wb-sect-val"><span class="wv-occ-total">'+mo.avgH+'%</span></div>'
              + '<div class="wv-occ-bar-track">'
              + '<div style="width:'+mo.avgT+'%;background:'+moGrad('#004948')+';height:6px"></div>'
              + '<div style="width:'+Math.max(0,mo.avgH-mo.avgT)+'%;background:'+moGrad('#52d9ce')+';height:6px"></div>'
              + '</div>';
            break;
          case 'mos_adr':
            cc = '<div class="wb-sect-val"><span class="wv-occ-total">$'+mo.avgToAdr+'</span></div>'
              + moBar(Math.round(mo.avgToAdr/3.5), '#004948');
            break;
          case 'mos_rev':
            cc = '<div class="wb-sect-val"><span class="wv-occ-total">'+mo.avgRev+'</span></div>'
              + '<div style="font-size:10px;color:#9ca3af;margin-top:1px">Total: '+mo.totalRev+'</div>';
            break;
          case 'mos_revpar':
            cc = '<div class="wb-sect-val"><span class="wv-occ-total">$'+mo.avgRevpar+'</span></div>'
              + moBar(Math.round(mo.avgRevpar/4), '#004948');
            break;
          case 'mos_pickup':
            cc = '<div class="wb-sect-val"><span class="wv-occ-total">+'+mo.avgPickup+'</span><span style="font-size:11px;color:#9ca3af;margin-left:6px">H: +'+mo.avgHPickup+'</span></div>'
              + moBar(Math.min(90,mo.avgPickup*3), '#004948');
            break;
          case 'mos_onoff':
            cc = '<div class="wb-sect-val"><span class="wv-occ-total">'+mo.avgOnline+'% online</span></div>'
              + '<div class="wv-occ-bar-track">'
              + '<div style="width:'+mo.avgOnline+'%;background:'+moGrad('#004948')+';height:6px"></div>'
              + '<div style="width:'+(100-mo.avgOnline)+'%;background:'+moGrad('#52d9ce')+';height:6px"></div>'
              + '</div>';
            break;
          case 'mos_segbar':
            cc = moStackBar([{p:mo.avgFit,c:'#006461'},{p:mo.avgDyn,c:'#0891b2'},{p:mo.avgSer,c:'#6366f1'},{p:mo.avgOtherSeg,c:'#5883ed'},{p:mo.avgFree,c:'#e5e7eb'}])
              + '<div style="display:flex;gap:4px;flex-wrap:wrap;margin-top:3px">'
              + '<span style="font-size:12px;font-family:Lato,sans-serif;color:#006461">FIT '+mo.avgFit+'%</span>'
              + '<span style="font-size:12px;font-family:Lato,sans-serif;color:#0891b2">Dyn '+mo.avgDyn+'%</span>'
              + '<span style="font-size:12px;font-family:Lato,sans-serif;color:#6366f1">Ser '+mo.avgSer+'%</span>'
              + '</div>';
            break;
          case 'mos_rn':
            cc = '<div class="wb-sect-val"><span class="wv-occ-total">'+mo.avgRn+' rn</span><span style="font-size:11px;color:#9ca3af;margin-left:6px">H: '+mo.avgHRn+'</span></div>'
              + moBar(Math.round(mo.avgRn/WV*100), '#004948') + moBar(Math.round(mo.avgHRn/WV*100), '#52d9ce');
            break;
          case 'mos_avga':
            cc = '<div class="wb-sect-val"><span class="wv-occ-total">'+mo.avgA+'</span><span style="font-size:11px;color:#9ca3af;margin-left:6px">H: '+mo.hAvgA+'</span></div>'
              + moBar(Math.min(90,parseFloat(mo.avgA)/3*100), '#004948');
            break;
          case 'mos_avgc':
            cc = '<div class="wb-sect-val"><span class="wv-occ-total">'+mo.avgC+'</span><span style="font-size:11px;color:#9ca3af;margin-left:6px">H: '+mo.hAvgC+'</span></div>'
              + moBar(Math.min(90,parseFloat(mo.avgC)/2*100), '#d33030');
            break;
          case 'mos_tota':
            cc = '<div class="wb-sect-val"><span class="wv-occ-total">'+mo.totA+'</span><span style="font-size:11px;color:#9ca3af;margin-left:6px">H: '+mo.hTotA+'</span></div>';
            break;
          case 'mos_totc':
            cc = '<div class="wb-sect-val"><span class="wv-occ-total">'+mo.totC+'</span><span style="font-size:11px;color:#9ca3af;margin-left:6px">H: '+mo.hTotC+'</span></div>';
            break;
          case 'mos_totg':
            cc = '<div class="wb-sect-val"><span class="wv-occ-total">'+mo.totG+'</span><span style="font-size:11px;color:#9ca3af;margin-left:6px">H: '+mo.hTotG+'</span></div>';
            break;
          case 'mos_los':
            cc = '<div class="wb-sect-val"><span class="wv-occ-total">'+mo.avgLos+'</span><span style="font-size:11px;color:#9ca3af;margin-left:6px">H: '+mo.avgHLos+'</span></div>'
              + moBar(Math.min(90,parseFloat(mo.avgLos)/10*100), '#004948');
            break;
          case 'mos_lead':
            cc = '<div class="wb-sect-val"><span class="wv-occ-total">'+mo.avgLead+'</span><span style="font-size:11px;color:#9ca3af;margin-left:6px">H: '+mo.avgHLead+'</span></div>'
              + moBar(Math.min(90,parseInt(mo.avgLead)/90*100), '#004948');
            break;
          case 'mos_avail':
            cc = '<div class="wb-sect-val"><span class="wv-occ-total">'+mo.avgAvailRooms+' rm</span></div>'
              + moBar(Math.min(90,Math.round(mo.avgAvailRooms/WV*100)), '#16a34a');
            break;
          case 'mos_availg':
            cc = '<div class="wb-sect-val"><span class="wv-occ-total">'+mo.avgAvailGuar+' rm</span></div>'
              + moBar(Math.min(90,Math.round(mo.avgAvailGuar/20*100)), '#004948');
            break;
          case 'mos_mpsum':
            { var _moGPR=parseFloat(mo.hAvgA)+parseFloat(mo.hAvgC);
              var _msAiR=Math.round(mo.avgHRn*mo.avgAi/100),_msAiSt=Math.round(_msAiR*_moGPR);
              var _msBbR=Math.round(mo.avgHRn*mo.avgBb/100),_msBbSt=Math.round(_msBbR*_moGPR);
              var _msHbR=Math.round(mo.avgHRn*mo.avgHb/100),_msHbSt=Math.round(_msHbR*_moGPR);
              var _msRoR=Math.round(mo.avgHRn*mo.avgRo/100),_msRoSt=Math.round(_msRoR*_moGPR);
            cc = moStackBar([{p:mo.avgAi,c:'#006461'},{p:mo.avgBb,c:'#3b82f6'},{p:mo.avgHb,c:'#967EF3'},{p:mo.avgRo,c:'#f59e0b'}])
              + '<div style="display:flex;gap:4px;flex-wrap:wrap;margin-top:3px">'
              + '<span style="font-size:12px;font-family:Lato,sans-serif;color:#006461">AI '+mo.avgAi+'% · '+_msAiSt+' seats</span>'
              + '<span style="font-size:12px;font-family:Lato,sans-serif;color:#3b82f6">BB '+mo.avgBb+'% · '+_msBbSt+' seats</span>'
              + '<span style="font-size:12px;font-family:Lato,sans-serif;color:#967EF3">HB '+mo.avgHb+'% · '+_msHbSt+' seats</span>'
              + '<span style="font-size:12px;font-family:Lato,sans-serif;color:#f59e0b">RO '+mo.avgRo+'% · '+_msRoSt+' seats</span>'
              + '</div>'; }
            break;
          case 'mos_bizbar':
            cc = moStackBar([{p:mo.avgToMix,c:'#006461'},{p:mo.avgDirMix,c:'#0284c7'},{p:mo.avgOtaMix,c:'#D97706'},{p:mo.avgOtherMix,c:'#9ca3af'}])
              + '<div style="display:flex;gap:4px;flex-wrap:wrap;margin-top:3px">'
              + '<span style="font-size:12px;font-family:Lato,sans-serif;color:#006461">TO '+mo.avgToMix+'%</span>'
              + '<span style="font-size:12px;font-family:Lato,sans-serif;color:#0284c7">D '+mo.avgDirMix+'%</span>'
              + '<span style="font-size:12px;font-family:Lato,sans-serif;color:#D97706">OTA '+mo.avgOtaMix+'%</span>'
              + '</div>';
            break;
          case 'mos_co_full':
            if (mo.fullCoCount > 0) {
              cc = '<div class="wb-sect-val">'
                + '<span class="material-icons" style="font-size:13px;color:#fca5a5;vertical-align:middle;margin-right:3px">lock</span>'
                + '<span class="wv-occ-total" style="color:#ef4444">' + mo.fullCoCount + ' day' + (mo.fullCoCount!==1?'s':'') + '</span>'
                + '<span style="font-size:10px;color:#9ca3af;margin-left:6px">/ ' + mo.nd + '</span>'
                + '</div>'
                + moBar(Math.min(90, Math.round(mo.fullCoCount/mo.nd*100)), '#ef4444');
            } else {
              cc = '<div class="wb-sect-val" style="color:#9ca3af;font-size:12px">None</div>';
            }
            break;
          case 'mos_co_part':
            if (mo.partCoCount > 0) {
              cc = '<div class="wb-sect-val">'
                + '<span class="material-icons" style="font-size:13px;color:#fde68a;vertical-align:middle;margin-right:3px">lock_open</span>'
                + '<span class="wv-occ-total" style="color:#d97706">' + mo.partCoCount + ' day' + (mo.partCoCount!==1?'s':'') + '</span>'
                + '<span style="font-size:10px;color:#9ca3af;margin-left:6px">/ ' + mo.nd + '</span>'
                + '</div>'
                + moBar(Math.min(90, Math.round(mo.partCoCount/mo.nd*100)), '#f59e0b');
            } else {
              cc = '<div class="wb-sect-val" style="color:#9ca3af;font-size:12px">None</div>';
            }
            break;
          case 'mos_tcbase': {
            cc = '<div class="wb-sect-val"><span class="wv-occ-total" style="font-weight:700;color:#1C1C1C">$'+mo.baseRate+'</span></div>'
              + moBar(Math.min(90,Math.round(mo.baseRate/280*100)), '#004948');
            break;
          }
          default:
            // Travel Co. rates (dynamic toIdx)
            if (row.toIdx !== undefined) {
              var isEbb = mo.isEbb;
              var promoTxt = isEbb ? 'EBB 10%' : 'Contract';
              var promoClr = isEbb ? '#16a34a' : '#2563eb';
              cc = '<div class="wb-sect-val" style="justify-content:space-between">'
                + '<span class="wv-occ-total" style="color:#1C1C1C">$'+mo.tcRates[row.toIdx]+'</span>'
                + '<span style="font-size:11px;font-weight:700;padding:1px 5px;border-radius:3px;background:'+promoClr+'22;color:'+promoClr+';border:1px solid '+promoClr+'44">'+promoTxt+'</span>'
                + '</div>'
                + moBar(Math.min(90,Math.round(mo.tcRates[row.toIdx]/280*100)), '#004948');
            }
            break;
        }

      } else {
        // Sub rows
        var v1 = '';
        switch (row.id) {
          case 'mos_occ_to':    v1 = mo.avgT+'%'; break;
          case 'mos_occ_htl':   v1 = mo.avgH+'%'; break;
          case 'mos_occ_stly':  v1 = mo.sdlyT+'%'; break;
          case 'mos_adr_to':    v1 = '$'+mo.avgToAdr; break;
          case 'mos_adr_htl':   v1 = '$'+mo.avgAdr; break;
          case 'mos_rev_to':    v1 = mo.avgRev; break;
          case 'mos_rev_htl':   v1 = mo.avgHRev; break;
          case 'mos_revpar_stly': v1 = '$'+mo.sdlyRevpar; break;
          case 'mos_rn_stly':   v1 = mo.sdlyRn+' rn'; break;
          case 'mos_onoff_on':  v1 = mo.avgOnline+'%'; break;
          case 'mos_onoff_off': v1 = (100-mo.avgOnline)+'%'; break;
          case 'mos_seg_fit':   v1 = mo.avgFit+'% · '+mo.fitRm+' rm'; break;
          case 'mos_seg_dyn':   v1 = mo.avgDyn+'% · '+mo.dynRm+' rm'; break;
          case 'mos_seg_ser':   v1 = mo.avgSer+'% · '+mo.serRm+' rm'; break;
          case 'mos_seg_oth':   v1 = mo.avgOtherSeg+'% · '+mo.othRm+' rm'; break;
          case 'mos_seg_rem':   v1 = mo.avgFree+'% · '+mo.freeRm+' rm'; break;
          case 'mos_mp_ai':     { var _moAiRm=Math.round(mo.avgHRn*mo.avgAi/100),_moAiSt=Math.round(_moAiRm*(parseFloat(mo.hAvgA)+parseFloat(mo.hAvgC))); v1=mo.avgAi+'% · '+_moAiRm+'r · '+_moAiSt+' seats'; } break;
          case 'mos_mp_bb':     { var _moBbRm=Math.round(mo.avgHRn*mo.avgBb/100),_moBbSt=Math.round(_moBbRm*(parseFloat(mo.hAvgA)+parseFloat(mo.hAvgC))); v1=mo.avgBb+'% · '+_moBbRm+'r · '+_moBbSt+' seats'; } break;
          case 'mos_mp_hb':     { var _moHbRm=Math.round(mo.avgHRn*mo.avgHb/100),_moHbSt=Math.round(_moHbRm*(parseFloat(mo.hAvgA)+parseFloat(mo.hAvgC))); v1=mo.avgHb+'% · '+_moHbRm+'r · '+_moHbSt+' seats'; } break;
          case 'mos_mp_ro':     { var _moRoRm=Math.round(mo.avgHRn*mo.avgRo/100),_moRoSt=Math.round(_moRoRm*(parseFloat(mo.hAvgA)+parseFloat(mo.hAvgC))); v1=mo.avgRo+'% · '+_moRoRm+'r · '+_moRoSt+' seats'; } break;
          case 'mos_biz_to':    v1 = mo.avgToMix+'%'; break;
          case 'mos_biz_dir':   v1 = mo.avgDirMix+'%'; break;
          case 'mos_biz_ota':   v1 = mo.avgOtaMix+'%'; break;
          case 'mos_biz_oth':   v1 = mo.avgOtherMix+'%'; break;
          case 'mos_co_rooms':  v1 = mo.coRoomTypes.length>0 ? mo.coRoomTypes.join(', ') : '—'; break;
          case 'mos_co_boards': v1 = mo.coBoards.length>0    ? mo.coBoards.join(', ')    : '—'; break;
          case 'mos_co_tos':    v1 = mo.coTOs.length>0       ? mo.coTOs.join(', ')       : '—'; break;
        }
        cc = '<span class="wb-sub-val">' + v1 + '</span>';
      }

      html += '<div class="wb-data-cell">' + cc + '</div>';
    });

    html += '</div>';
  });

  html += '</div>'; // close wb-layout

  // Wrap in overview accordion
  var ovLabel = months.length===1 ? months[0].name+' Overview' : months[0].name+' – '+months[months.length-1].name+' Overview';
  var ovCollapsed = _calAccState['overview'] === false ? false : true;
  var ovChev = ovCollapsed
    ? '<span class="material-icons" style="font-size:16px">expand_more</span>'
    : '<span class="material-icons" style="font-size:16px">expand_less</span>';

  el.innerHTML = '<div class="cal-summary-wrap" style="background:#fff">'
    +'<div class="wv-acc-sect' + (ovCollapsed ? '' : ' wv-acc-open') + '" style="border:1px solid #dde1e2;border-radius:0;overflow:hidden">'
    +'<div class="wv-acc-hdr" data-cal-section="overview" onclick="calAccClick(this)" style="background:#fff;border-bottom:none;border-radius:0">'
    +'<span class="wv-acc-chev" style="color:#006461">'+ovChev+'</span>'
    +'<span class="wv-acc-title" style="font-weight:700">'+ovLabel+'</span>'
    +'</div>'
    +'<div class="wv-acc-body' + (ovCollapsed ? ' wv-body-hidden' : '') + '" style="padding:0;background:#fff">'
    +html
    +'</div></div>'
    +'</div>';
}

// ── Monthly overview toggle handler ──────────────────────────────────────
window.moToggle = function(id) {
  _calAccState[id] = !_calAccState[id];
  renderCalMonthlySummary();
};

// ── Monthly tab bar interactions ──────────────────────────────────────────
// Tab clicks: Monthly stays in month view; Daily/Close Outs/Close Out Report switch to week view
document.querySelectorAll('.mo-grp-btn').forEach(function(btn) {
  btn.addEventListener('click', function() {
    var tab = this.dataset.mogroupby;
    if (tab === 'monthly') return; // already on monthly
    // Map monthly tab names to weekly groupby values
    var groupbyMap = { daily: 'dailyB', closeouts: 'roomType', coReport: 'coReport' };
    var wvGrp = groupbyMap[tab] || 'dailyB';
    // Switch to week view with this groupby active
    wvGroupBy = wvGrp;
    openWeekView(wvMonth, wvWeekStart);
    // Highlight correct tab in weekly bar
    document.querySelectorAll('#weekView .wv-groupby-btn').forEach(function(b) {
      b.classList.toggle('active', b.dataset.groupby === wvGrp);
    });
    _updateAccBtnState();
  });
});

window.moAccCloseAll = function() {
  ['mo_daily','mo_seg','mo_more','mo_meals','mo_biz','mo_tc','overview'].forEach(function(k) { _calAccState[k] = true; });
  renderCalMonthlySummary();
};
window.moAccOpenAll = function() {
  ['mo_daily','mo_seg','mo_more','mo_meals','mo_biz','mo_tc','overview'].forEach(function(k) { _calAccState[k] = false; });
  renderCalMonthlySummary();
};

// ── View toggle: 1 / 3 / 6 / 12 months ───────────────────────────────────
window.calSetDisplayView = function(n) {
  calDisplayView = n;
  calView = n;

  // Update select value
  var selEl = document.getElementById('calViewSelect');
  if (selEl) selEl.value = String(n);

  // Disable Cell Metrics control in 6/12 mode (heatmap only — no cell content)
  var metricsWrap = document.getElementById('calMetricsWrap');
  var metricsBtn  = document.getElementById('calMetricsBtn');
  if (metricsWrap && metricsBtn) {
    var isCompact = (n >= 3);
    metricsBtn.disabled = isCompact;
    metricsWrap.classList.toggle('cal-metrics-disabled', isCompact);
    // Close dropdown if open
    if (isCompact) {
      var dd = document.getElementById('calMetricsDropdown');
      if (dd) dd.style.display = 'none';
    }
  }

  // Show/hide monthly tab bar (only for 1/2/3 month views)
  var moBar = document.getElementById('moGroupbyBar');
  if (moBar) moBar.style.display = (n <= 3) ? '' : 'none';

  // Apply compact CSS class + view class
  var grid = document.getElementById('calMonths');
  if (grid) {
    if (n === 6 || n === 12) { grid.classList.add('cal-compact'); }
    else { grid.classList.remove('cal-compact'); }
    if (n === 12) grid.classList.add('cal-12m');
    else grid.classList.remove('cal-12m');
    // View-specific class for responsive breakpoints
    grid.className = grid.className.replace(/\bcal-view-\d+\b/g, '');
    grid.classList.add('cal-view-' + n);
  }
  // Update compare dropdown state based on breakpoint
  _calUpdateCompareState();

  // Clamp start index
  calStartIdx = Math.min(calStartIdx, Math.max(0, ALL_MONTHS.length - calView));
  renderCalendar();

  // Re-apply compact class + out-of-range after render
  setTimeout(function() {
    var g = document.getElementById('calMonths');
    if (g) {
      if (n === 6 || n === 12) g.classList.add('cal-compact');
      else g.classList.remove('cal-compact');
      if (n === 12) g.classList.add('cal-12m');
      else g.classList.remove('cal-12m');
      g.className = g.className.replace(/\bcal-view-\d+\b/g, '');
      g.classList.add('cal-view-' + n);
    }
    if (typeof applyOutOfRange === 'function') applyOutOfRange();
  }, 80);
};

// Monthly view filter dropdown toggle
window.calToggleMFilt = function(panelId, btn) {
  const panels = ['calFiltTOPanel','calFiltRTPanel','calFiltMPPanel','calFiltOriginPanel'];
  const btns   = ['calFiltTOBtn','calFiltRTBtn','calFiltMPBtn','calFiltOriginBtn'];
  const targetPanel = document.getElementById(panelId);
  const isCurrentlyOpen = targetPanel && targetPanel.style.display !== 'none';
  // Close all other dropdowns (including rev chart, cal/wv filters) before opening
  if (!isCurrentlyOpen) _closeAllDropdowns(panelId);
  panels.forEach(function(pid, i) {
    const p = document.getElementById(pid);
    if (!p) return;
    if (pid === panelId) {
      p.style.display = isCurrentlyOpen ? 'none' : 'block';
      const b = document.getElementById(btns[i]);
      if (b) b.classList.toggle('active', !isCurrentlyOpen);
    } else {
      p.style.display = 'none';
      const b = document.getElementById(btns[i]);
      if (b) b.classList.remove('active');
    }
  });
};

// Close monthly filter dropdowns when clicking outside
document.addEventListener('click', function(e) {
  if (!e.target.closest('.cal-mfilt-wrap')) {
    ['calFiltTOPanel','calFiltRTPanel','calFiltMPPanel','calFiltOriginPanel'].forEach(function(pid) {
      const p = document.getElementById(pid);
      if (p) p.style.display = 'none';
    });
    ['calFiltTOBtn','calFiltRTBtn','calFiltMPBtn','calFiltOriginBtn'].forEach(function(bid) {
      const b = document.getElementById(bid);
      if (b) b.classList.remove('active');
    });
  }
});

/* ─── CALENDAR RANGE SELECTION ─── */
function calDv(m, d) { return m * 100 + d; }

function applyCalSelection(hoverMonth, hoverDay) {
  const cells = document.querySelectorAll('#calMonths .cal-day:not(.empty)');
  const sv = calSelStart ? calDv(calSelStart.month, calSelStart.day) : null;
  let ev = calSelEnd ? calDv(calSelEnd.month, calSelEnd.day)
         : (calSelPicking && hoverMonth) ? calDv(hoverMonth, hoverDay)
         : null;

  const lo = (sv !== null && ev !== null) ? Math.min(sv, ev) : sv;
  const hi = (sv !== null && ev !== null) ? Math.max(sv, ev) : null;

  cells.forEach(cell => {
    cell.classList.remove('cal-sel-lo', 'cal-sel-hi', 'cal-sel-mid');
    if (sv === null) return;
    const v = calDv(+cell.dataset.month, +cell.dataset.day);
    if (v === lo)                        cell.classList.add('cal-sel-lo');
    else if (hi && v === hi)             cell.classList.add('cal-sel-hi');
    else if (hi && v > lo && v < hi)     cell.classList.add('cal-sel-mid');
  });

  // Update header-right: show range-mode (hide all but Select Range + Close Out)
  const hdrRight = document.querySelector('.cal-header-right');
  if (hdrRight) hdrRight.classList.toggle('range-mode', !!(calSelStart || calSelPicking));

  // Update Select Range button label
  const selBtn = document.getElementById('calSelBtn');
  if (selBtn) {
    if (calSelStart && calSelEnd) {
      const MNAMES = ['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
      const startV = calDv(calSelStart.month, calSelStart.day);
      const endV   = calDv(calSelEnd.month,   calSelEnd.day);
      const loD = startV <= endV ? calSelStart : calSelEnd;
      const hiD = startV <= endV ? calSelEnd   : calSelStart;
      selBtn.innerHTML = svgCal + ` ${MNAMES[loD.month]} ${loD.day} – ${MNAMES[hiD.month]} ${hiD.day}` + svgChevronDown;
    } else if (calSelStart && !calSelEnd) {
      selBtn.innerHTML = svgCal + ' Pick end…';
    }
  }

  // Update grid picking cursor
  const grid = document.getElementById('calMonths');
  if (grid) grid.classList.toggle('range-picking', calSelPicking);
}

const svgCal = `<svg viewBox="0 0 14 14" fill="none" stroke="currentColor" stroke-width="1.5" width="12" height="12" style="flex-shrink:0"><rect x="1" y="2" width="12" height="11" rx="1"/><line x1="4" y1="1" x2="4" y2="3"/><line x1="10" y1="1" x2="10" y2="3"/><line x1="1" y1="6" x2="13" y2="6"/></svg>`;
const svgChevronDown = `<span class="material-icons" style="font-size:14px;flex-shrink:0;margin-left:2px">expand_more</span>`;

function clearCalSelection() {
  calSelStart = null; calSelEnd = null; calSelPicking = false;
  // Reset Close button appearance
  const closeBtn = document.getElementById('calCloseOutBtn');
  if (closeBtn) { closeBtn.style.background = ''; closeBtn.style.color = ''; }
  const hdrRight = document.querySelector('.cal-header-right');
  if (hdrRight) hdrRight.classList.remove('range-mode');
  applyCalSelection();
}

/* ─── CALENDAR NAV & VIEW SELECTOR ─── */
(function () {
  function clamp() {
    calStartIdx = Math.max(0, Math.min(calStartIdx, ALL_MONTHS.length - calView));
  }

  function renderAndRestoreCompact() {
    renderCalendar();
    var g = document.getElementById('calMonths');
    if (g && (calDisplayView === 6 || calDisplayView === 12)) g.classList.add('cal-compact');
    if (calDisplayView === 12) g && g.classList.add('cal-12m');
    if (typeof applyOutOfRange === 'function') applyOutOfRange();
  }

  // New date-nav row buttons — move by calDisplayView months in 6/12 mode
  document.getElementById('calPrev')
    ?.addEventListener('click', () => {
      calStartIdx -= (calDisplayView >= 6 ? calDisplayView : 1);
      clamp(); renderAndRestoreCompact();
    });
  document.getElementById('calNext')
    ?.addEventListener('click', () => {
      calStartIdx += (calDisplayView >= 6 ? calDisplayView : 1);
      clamp(); renderAndRestoreCompact();
    });

  // Monthly tab-bar shuffler (mirrors calPrev/calNext)
  document.getElementById('moShufPrev')
    ?.addEventListener('click', () => {
      calStartIdx -= (calDisplayView >= 6 ? calDisplayView : 1);
      clamp(); renderAndRestoreCompact();
    });
  document.getElementById('moShufNext')
    ?.addEventListener('click', () => {
      calStartIdx += (calDisplayView >= 6 ? calDisplayView : 1);
      clamp(); renderAndRestoreCompact();
    });

  // Legacy selectors (kept for safety)
  document.querySelector('.cal-nav-left .cal-nav-btn:first-child')
    ?.addEventListener('click', () => { calStartIdx--; clamp(); renderCalendar(); });
  document.querySelector('.cal-nav-left .cal-nav-btn:last-child')
    ?.addEventListener('click', () => { calStartIdx++; clamp(); renderCalendar(); });

  document.querySelector('.cal-nav-right .cal-selector select')
    ?.addEventListener('change', e => {
      calView = parseInt(e.target.value, 10);
      clamp();
      renderCalendar();
    });
})();

/* ─── CLOSE DROPDOWN (replaces standalone Select Range button) ─── */
(function () {
  // Toggle the Close dropdown
  window.calCloseDropdownToggle = function(e) {
    e.stopPropagation();
    var dd = document.getElementById('calCloseDropdown');
    if (!dd) return;
    var open = dd.style.display !== 'none';
    dd.style.display = open ? 'none' : 'block';
  };

  // "Select range" option — enter range-picking mode (same as old calSelBtn)
  window.calCloseSelectRange = function() {
    document.getElementById('calCloseDropdown').style.display = 'none';
    if (calSelStart) { clearCalSelection(); return; }
    calSelPicking = true;
    const hdrRight = document.getElementById('calMonths')?.closest('.section-card')
                       ?.querySelector('.cal-header-right');
    if (hdrRight) hdrRight.classList.add('range-mode');
    const grid = document.getElementById('calMonths');
    if (grid) grid.classList.add('range-picking');
    // Update Close button label to show picking state
    const closeBtn = document.getElementById('calCloseOutBtn');
    if (closeBtn) {
      closeBtn.style.background = '#006461';
      closeBtn.style.color = '#fff';
    }
  };

  // "Custom" option — open modal with no pre-populated dates
  window.calCloseCustom = function() {
    var dd = document.getElementById('calCloseDropdown');
    if (dd) dd.style.display = 'none';
    if (typeof window._coOpenModal === 'function') window._coOpenModal('', '', 'cal');
  };

  // Monthly Close Out button (opens modal directly)
  window.moOpenCloseOut = function() {
    if (typeof window._coOpenModal === 'function') window._coOpenModal('', '', 'cal');
  };

  // Close dropdown on outside click
  document.addEventListener('click', function(e) {
    var wrap = document.getElementById('calCloseWrap');
    var dd   = document.getElementById('calCloseDropdown');
    if (dd && wrap && !wrap.contains(e.target)) dd.style.display = 'none';
  });

  // Escape cancels range picking
  document.addEventListener('keydown', e => { if (e.key === 'Escape') clearCalSelection(); });
})();

/* ─── THEME TOGGLE ─── */
(function () {
  const btn = document.getElementById('themeToggle');
  const saved = localStorage.getItem('theme');
  if (saved === 'dark') document.body.classList.add('dark');

  btn.addEventListener('click', () => {
    const isDark = document.body.classList.toggle('dark');
    localStorage.setItem('theme', isDark ? 'dark' : 'light');
  });
})();

/* ─── DAY POPUP ─── */
(function () {
  const popup  = document.getElementById('dayPopup');
  const close  = document.getElementById('popupClose');
  const dateEl = document.getElementById('popupDate');
  if (!popup) return;

  const DAY_NAMES   = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
  const MONTH_NAMES = ['','January','February','March','April','May','June',
                        'July','August','September','October','November','December'];

  function openPopup(cell, month, day) {
    const dm = month, dd = day;
    const d = new Date(2026, dm - 1, dd);
    const TODAY_POPUP = new Date(2026, 2, 9);
    const popupDba = Math.round((d - TODAY_POPUP) / 86400000);
    const popupDbaHtml = popupDba === 0
      ? ' <span style="font-size:9px;font-weight:700;color:#006461;background:#ccfbf1;padding:1px 6px;border-radius:4px;margin-left:4px;vertical-align:middle">Today</span>'
      : popupDba > 0
      ? ' <span style="font-size:9px;font-weight:700;color:#006461;background:#ccfbf1;padding:1px 6px;border-radius:4px;margin-left:4px;vertical-align:middle">' + popupDba + ' DBA</span>'
      : '';
    dateEl.innerHTML = DAY_NAMES[d.getDay()] + ', ' + MONTH_NAMES[dm] + ' ' + dd + ', 2026' + popupDbaHtml;

    // ── Read active filters ──
    const _fCal = typeof filterState !== 'undefined' ? filterState.cal : {};
    const toKey  = _fCal.calFiltTO || calFiltTO || 'all';
    const _rtFilt = _fCal.calFiltRoom || 'all';
    const _bdFilt = _fCal.calFiltBoard || 'all';
    const _mkFilt = _fCal.calFiltMarket || 'all';
    const _hasAnyFilter = toKey !== 'all' || _rtFilt !== 'all' || _bdFilt !== 'all' || _mkFilt !== 'all';

    // ── Compute base values ──
    const { hotel: hotelBase, to: toRaw } = getOccupancy(dm, dd);
    const toMult = TO_FILTER_MULT[toKey] || 1.0;
    const toBase = Math.min(95, Math.round(toRaw * toMult));
    const toLabel = toKey !== 'all'
      ? toKey.charAt(0).toUpperCase() + toKey.slice(1).replace(/-/g,' ')
      : 'All Operators';

    const adrBase = 150 + Math.abs((dm * 47 + dd * 31) % 130);
    const v       = Math.abs((dm * 127 + dd * 53 + dm * dd * 7 + dd * dd * 3)) % 100;

    // ── Apply filter multipliers to metrics ──
    // Room type filter: each room type represents a share of total inventory
    var _rtMult = 1.0;
    if (_rtFilt !== 'all') {
      var _rtShares = {standard:0.34,superior:0.24,deluxe:0.18,suite:0.08,'jr. suite':0.10,family:0.06};
      var _rtParts = _rtFilt.split(',');
      _rtMult = _rtParts.reduce(function(a,b){ return a + (_rtShares[b.trim().toLowerCase()] || 0.15); }, 0);
      _rtMult = Math.min(1, _rtMult);
    }
    // Board type filter
    var _bdMult = 1.0;
    if (_bdFilt !== 'all') {
      var _bdShares = {ai:0.55,bb:0.20,hb:0.15,ro:0.10,fb:0.05};
      var _bdParts = _bdFilt.split(',');
      _bdMult = _bdParts.reduce(function(a,b){ return a + (_bdShares[b.trim()] || 0.15); }, 0);
      _bdMult = Math.min(1, _bdMult);
    }
    // Market filter
    var _mkMult = _mkFilt !== 'all' ? 0.6 : 1.0;

    var _filterMult = _rtMult * _bdMult * _mkMult;
    // Filtered occupancy: scale rooms sold by filter scope
    var hotel = _hasAnyFilter ? Math.max(5, Math.round(hotelBase * _filterMult)) : hotelBase;
    var to    = _hasAnyFilter ? Math.max(2, Math.round(toBase * _filterMult))    : toBase;
    // Filtered ADR: slight variation by room type (suites higher, standard lower)
    var _adrAdj = 0;
    if (_rtFilt !== 'all') {
      var _rtAdrMap = {standard:-15,superior:0,deluxe:20,suite:80,'jr. suite':50,family:10};
      var _rtFirst = _rtFilt.split(',')[0].trim().toLowerCase();
      _adrAdj = _rtAdrMap[_rtFirst] || 0;
    }
    var adr = adrBase + _adrAdj;
    var rev = Math.floor(hotel * adr * HOTEL_CAPACITY / 100 * 1.1);
    var onlinePct = Math.max(30, Math.min(80, 45 + Math.abs((dm * 13 + dd * 7) % 35)));
    var offlinePct = 100 - onlinePct;
    var adrBar = Math.min(95, 40 + Math.abs((dm * 11 + dd * 19) % 55));
    var revBar = Math.min(95, Math.round((35 + Math.abs((dm * 17 + dd * 13) % 60)) * _filterMult));
    var hotelSDLY = Math.max(5, hotel - 3 - (v % 5));
    var toSDLY    = Math.max(5, to    - 2 - (v % 4));
    var adrSDLY   = adr - 8;
    var revSDLY   = Math.floor(rev * 0.9);
    const pad = function(n){ return String(n).padStart(2,'0'); };

    // update operator label
    const opEl = document.getElementById('popupOperator');
    if (opEl) opEl.textContent = toLabel;

    // ── Computed values ──
    var sign       = function(n){ return n >= 0 ? '+' + n : String(n); };
    var otherPct   = Math.max(0, hotel - to);
    var freePct    = 100 - hotel;
    var filteredCap = _hasAnyFilter ? Math.round(HOTEL_CAPACITY * _rtMult) : HOTEL_CAPACITY;
    var toRms      = Math.round(filteredCap * to       / 100);
    var otherRms   = Math.round(filteredCap * otherPct / 100);
    var freeRms    = Math.max(0, filteredCap - toRms - otherRms);
    var rnSold     = Math.floor(hotel * filteredCap / 100);
    var availRooms = Math.max(0, filteredCap - rnSold);
    var availGuar  = Math.max(0, Math.floor((8 + v % 5) * _filterMult));
    var sdlyR      = Math.floor(rev * 0.9);

    // Filter label for sections
    var _filtLabel = '';
    if (_hasAnyFilter) {
      var _parts = [];
      if (toKey !== 'all') _parts.push(toLabel);
      if (_rtFilt !== 'all') _parts.push(_rtFilt.split(',').map(function(s){return s.trim();}).join(', '));
      if (_bdFilt !== 'all') { var _bm={ai:'AI',bb:'B&B',hb:'HB',ro:'RO',fb:'FB'}; _parts.push(_bdFilt.split(',').map(function(b){return _bm[b.trim()]||b;}).join(', ')); }
      if (_mkFilt !== 'all') _parts.push('Market: '+_mkFilt);
      _filtLabel = '<div style="font-size:8px;font-weight:600;color:#006461;background:#d7f7ed;padding:2px 8px;border-radius:4px;margin-bottom:6px;text-align:center">Filtered: '+_parts.join(' · ')+'</div>';
    }

    var _basePickup = Math.max(1, Math.floor((v%25+5)*_filterMult));
    var detRows = [
      ['RN Sold',      rnSold,                            Math.floor(rnSold*0.88),          '+' + Math.floor(v%30+5)],
      ['ADR',          '$' + adr,                         '$' + adrSDLY,                    '+' + (3+v%12)+'%'],
      ['Revenue',      '$' + Math.floor(rev/1000) + 'k',  '$' + Math.floor(sdlyR/1000)+'k','+' + (5+v%15)+'%'],
      ...pickupDayValues.map(function(dv, i) {
        if (!wvMetricState['dm_pickup_' + i]) return null;
        var sc  = dv<=1?0.3:dv<=3?0.6:dv<=7?1:Math.min(2,dv/7);
        var val = Math.max(0, Math.round(_basePickup * sc));
        return ['Pickup ' + dv, '+' + val, '+0', '+' + Math.floor((v%15+5)*sc)];
      }).filter(Boolean),
      ['Avg Adults',   (1.8+v%3*.1).toFixed(1),           '1.9',                            '-0.1'],
      ['Avg Children', (0.3+v%2*.1).toFixed(1),           '0.4',                            '-0.1'],
      ['REVPAR',       '$' + Math.round(adr*hotel/100),   '$' + Math.floor(adr*0.92),       '+' + (10+v%20)+'%'],
      ['Avail Rooms',  availRooms,                         availRooms+3,                     '-' + (Math.floor(v%8)+1)],
      ['Avail Guar.',  availGuar,                          availGuar+2,                      '-' + (Math.floor(v%4)+1)],
    ];

    // Meal plans — filtered by board type
    var aiPct = Math.max(45, Math.min(68, 55 + (dm*7+dd*3)%14));
    var bbPct = Math.max(14, Math.min(28, 20 + (dm*11+dd*5)%11));
    var hbPct = Math.max(6,  Math.min(16, 10 + (dm*5+dd*7)%9));
    var roPct = Math.max(2,  100 - aiPct - bbPct - hbPct);
    var mealPlans = [
      { short:'AI', pct:aiPct, color:'#004948', key:'ai' },
      { short:'BB', pct:bbPct, color:'#52d9ce', key:'bb' },
      { short:'HB', pct:hbPct, color:'#C4FF45', key:'hb' },
      { short:'RO', pct:roPct, color:'#D97706', key:'ro' },
    ];
    // When board filter active, filter meal plans to only show selected
    if (_bdFilt !== 'all') {
      var _bdSel = _bdFilt.split(',').map(function(s){ return s.trim(); });
      mealPlans = mealPlans.filter(function(mp){ return _bdSel.indexOf(mp.key) >= 0; });
      if (mealPlans.length === 0) mealPlans = [{ short:'AI', pct:aiPct, color:'#004948', key:'ai' }]; // fallback
      // Normalize percentages to 100%
      var _mpTotal = mealPlans.reduce(function(a,p){ return a+p.pct; },0);
      if (_mpTotal > 0) mealPlans.forEach(function(p){ p.pct = Math.round(p.pct / _mpTotal * 100); });
    }
    var mealBarHtml = '<div class="wv-meals-bar" style="margin:4px 0 6px">'
      + mealPlans.map(function(p){ return '<div style="width:'+p.pct+'%;background:'+p.color+';height:100%"></div>'; }).join('')
      + '</div>';
    var _popAvgGPR = (1.8+v%3*0.1+0.3) + (0.3+v%2*0.1+0.1);
    var mealRowsHtml = mealPlans.map(function(p) {
      var rooms = Math.round(rnSold * p.pct / 100);
      var seats = Math.round(rooms * _popAvgGPR);
      var lyPct = Math.max(1, p.pct - 3 + (dm+dd)%5 - 2);
      var diff  = p.pct - lyPct;
      return '<div class="wv-meal-row">'
        +'<span class="wv-meal-dot" style="background:'+p.color+'"></span>'
        +'<span class="wv-meal-name">'+p.short+'</span>'
        +'<div class="wv-meal-bar-track"><div class="wv-meal-bar-fill" style="width:'+p.pct+'%;background:'+p.color+'"></div></div>'
        +'<span class="wv-meal-pct">'+p.pct+'%</span>'
        +'<span class="wv-meal-rooms">'+rooms+' rm</span>'
        +'<span class="wv-meal-rooms" style="color:#374151">'+seats+' seats</span>'
        +'<span class="wv-meal-delta '+(diff>=0?'pos':'neg')+'">'+(diff>=0?'+':'')+diff+'pp</span>'
        +'</div>';
    }).join('');

    // Room availability — filtered by active calendar filters
    var rtRowsAll = [['Standard',51],['Superior',36],['Deluxe',27],['Suite',12],['Jr. Suite',15],['Family',9]];
    // Filter room types when a specific room filter is active
    var rtRows = _rtFilt !== 'all' ? rtRowsAll.filter(function(r) {
      var selected = _rtFilt.split(',');
      return selected.some(function(s) { return s.trim().toLowerCase() === r[0].toLowerCase(); });
    }) : rtRowsAll;
    if (rtRows.length === 0) rtRows = rtRowsAll; // fallback if no match

    // Per-room-type sold: compute bookings specific to active filters
    // TO share per room type (seed-based variation)
    var _toShareMap = {'sunshine-tours':[.22,.20,.18,.25,.20,.22],'global-adv':[.20,.22,.24,.20,.22,.18],
      'beach-hols':[.18,.16,.20,.15,.18,.24],'city-breaks':[.24,.22,.18,.20,.20,.18],'adventure':[.16,.20,.20,.20,.20,.18]};
    // Board share per room type
    var _bdShareMap = {ai:[.55,.50,.45,.35,.40,.60],bb:[.20,.22,.24,.25,.25,.18],
      hb:[.15,.16,.18,.22,.20,.14],ro:[.10,.12,.13,.18,.15,.08]};

    var rtHTML = rtRows.map(function(row) {
      var name = row[0], inv = row[1];
      var origIdx = rtRowsAll.findIndex(function(r){ return r[0] === name; });
      var colorIdx = origIdx >= 0 ? origIdx : 0;
      // Base sold for this room type from overall hotel occupancy
      var baseSold = Math.min(inv, Math.floor(inv * hotelBase / 110));
      // Apply per-room-type filter multipliers
      var rtFiltMult = 1.0;
      if (toKey !== 'all') {
        var _ts = _toShareMap[toKey];
        rtFiltMult *= _ts ? _ts[origIdx] || 0.20 : 0.20;
      }
      if (_bdFilt !== 'all') {
        var _bdParts = _bdFilt.split(',').map(function(s){ return s.trim(); });
        var _bdSum = _bdParts.reduce(function(a,b){
          var _bs = _bdShareMap[b];
          return a + (_bs ? _bs[origIdx] || 0.15 : 0.15);
        }, 0);
        rtFiltMult *= Math.min(1, _bdSum);
      }
      if (_mkFilt !== 'all') rtFiltMult *= 0.6;
      var sold = _hasAnyFilter
        ? Math.min(inv, Math.max(0, Math.round(baseSold * rtFiltMult)))
        : baseSold;
      var avail = Math.max(0, inv - sold);
      var pct   = Math.round(sold / inv * 100);
      var barClr  = avail === 0 ? '#dc2626' : pct >= 85 ? '#ea580c' : pct >= 60 ? '#f59e0b' : '#16a34a';
      var availClr = avail === 0 ? '#dc2626' : avail <= 3 ? '#d97706' : '#16a34a';
      var rowBg = avail === 0 ? 'background:#fff1f2;border-radius:4px;padding:2px 4px;margin:1px -4px;' : 'padding:2px 0;';
      return '<div class="popup-rt-row" style="flex-direction:column;gap:2px;align-items:stretch;'+rowBg+'">'
        +'<div style="display:flex;align-items:center;gap:4px">'
        +'<span class="popup-rt-sw" style="background:'+RT_COLORS[colorIdx]+'"></span>'
        +'<span class="popup-rt-nm" style="flex:1">'+name+'</span>'
        +(_hasAnyFilter
          ? '<span style="font-size:12px;font-weight:600;color:#004948;margin-right:2px">'+sold+'<span style="color:#94a3b8;font-weight:400"> booked</span></span>'
          : '')
        +(avail === 0
          ? '<span style="font-size:12px;font-weight:700;color:#16a34a">0 available</span>'
          : '<span style="font-size:12px;font-weight:700;color:'+availClr+'">'+avail+' avail</span>')
        +'</div>'
        +'<div style="height:6px;border-radius:2px;background:#e5e7eb;overflow:hidden;margin-left:11px">'
        +'<div style="height:100%;width:'+pct+'%;background:'+barClr+';border-radius:2px"></div>'
        +'</div>'
        +'</div>';
    }).join('');

    // TO Rates — filtered by active TO filter
    var toNamesAll  = ['Sunshine Tours','Global Adv.','Beach Hols','City Breaks','Adventure'];
    var toColorsAll = ['#3b82f6','#967EF3','#0ea5e9','#10b981','#f59e0b'];
    var _toFilterMap = {'sunshine-tours':0,'global-adv':1,'beach-hols':2,'city-breaks':3,'adventure':4};
    var toNames = toNamesAll, toColors = toColorsAll;
    if (toKey !== 'all' && _toFilterMap[toKey] !== undefined) {
      var _tfi = _toFilterMap[toKey];
      toNames = [toNamesAll[_tfi]];
      toColors = [toColorsAll[_tfi]];
    }
    var toRatesHTML = toNames.map(function(name, i) {
      var origIdx = toNamesAll.indexOf(name);
      if (origIdx < 0) origIdx = i;
      var toRate  = adr - 15 + Math.abs((dm*(origIdx+3) + dd*(origIdx+5)) % 50);
      var toAllot = 5  + Math.abs((dm*(origIdx+2) + dd*(origIdx+3)) % 20);
      var toUsed  = Math.max(0, toAllot - Math.floor(hotel / 20));
      var barPct  = Math.round((toUsed / toAllot) * 100);
      var barCls  = barPct >= 90 ? 'wv-to-bar-high' : barPct >= 60 ? 'wv-to-bar-mid' : 'wv-to-bar-low';
      return '<div class="wv-to-rate-row">'
        +'<span class="wv-to-dot" style="background:'+toColors[i]+'"></span>'
        +'<span class="wv-to-name">'+name+'</span>'
        +'<div class="wv-to-bar-wrap"><div class="wv-to-bar '+barCls+'" style="width:'+barPct+'%"></div></div>'
        +'<span class="wv-to-rate">$'+toRate+'</span>'
        +'<span class="wv-to-allot">'+(toAllot-toUsed)+' rooms</span>'
        +'</div>';
    }).join('');

    // Restrictions for eye popup
    const popupCl = PARTIAL_CLOSURES[dm + '-' + dd];
    const hasPopupCl = popupCl && Array.isArray(popupCl) && popupCl.length > 0;
    function popupClChip(label, color) {
      return '<span class="wv-cl-chip" style="color:'+color+';background:'+color+'18;border-color:'+color+'3a">'+label+'</span>';
    }
    const BMAP_P={ai:'All Inclusive',bb:'Bed & Breakfast',hb:'Half Board',ro:'Room Only',fb:'Full Board'};
    const popupClHtml = !hasPopupCl ? '' : (function(){
      const ruleCards = popupCl.map(function(rule, ri){
        const toPart = rule.tos.length ? rule.tos.map(function(n){ return popupClChip(n, TO_COLORS_MAP[n]||'#dc2626'); }).join('') : popupClChip('All Operators','#9ca3af');
        const rtPart = rule.roomTypes.length ? rule.roomTypes.map(function(n){ return popupClChip(n, RT_NAME_COLORS[n]||'#b45309'); }).join('') : popupClChip('All Room Types','#9ca3af');
        const bdPart = rule.boards.length ? rule.boards.map(function(b){ return popupClChip(BMAP_P[b]||b,'#7c3aed'); }).join('') : popupClChip('All Meal Plans','#9ca3af');
        return '<div style="margin-bottom:4px;padding:4px 0;border-bottom:1px solid #f3f4f6">'
          +'<span style="font-size:12px;font-weight:700;color:#dc2626">Strategy</span>'
          +'<div style="display:flex;flex-wrap:wrap;gap:3px;margin-top:3px">'+toPart+rtPart+bdPart+'</div>'
          +'</div>';
      }).join('');
      return '<div class="popup-metrics-section popup-closures-section">'
        +'<div class="popup-metrics-title" style="color:#dc2626">CLOSED OUT</div>'
        +ruleCards+'</div>';
    })();

    // ── Daily B–style group/section/sub builders ──
    var _C1='#004948',_C2='#52d9ce',_C3='#D97706',_CSTLY='#C4FF45',_CREM='#445e0d';
    function _pGrad(c){if(c==='#004948')return'linear-gradient(to right,#004948,#007a75)';if(c==='#52d9ce')return'linear-gradient(to right,#52d9ce,#8aeee8)';if(c==='#445e0d')return'linear-gradient(to right,#445e0d,#6a9014)';if(c==='#D97706')return'linear-gradient(to right,#D97706,#F59E0B)';if(c==='#16a34a')return'linear-gradient(to right,#16a34a,#22c55e)';if(c==='#C4FF45')return'linear-gradient(to right,#C4FF45,#D4FF73)';return c;}
    function _pBar(pct,c){return'<div class="wv-occ-bar-track" style="margin:2px 0 0;height:6px;border-radius:2px"><div style="width:'+pct+'%;background:'+_pGrad(c)+';height:6px"></div></div>';}
    function _pSbar(segs){return'<div class="wv-occ-bar-track" style="margin:2px 0 0;height:6px;border-radius:2px">'+segs.map(function(s){return'<div style="width:'+s.p+'%;background:'+_pGrad(s.c)+';height:6px"></div>';}).join('')+'</div>';}
    function _pGrp(label,clr){return'<div class="pb-grp" style="background:'+clr+';color:#fff;font-size:12px;font-weight:700;padding:4px 8px;margin:0 -10px;letter-spacing:.5px">'+label+'</div>';}
    function _pGrpStart(label,clr,uid){return'<div class="pb-grp pb-grp-toggle" data-grpid="'+uid+'" style="background:'+clr+';color:#fff;font-size:12px;font-weight:700;padding:4px 8px;margin:0 -10px;letter-spacing:.5px;display:flex;align-items:center;justify-content:space-between;cursor:pointer;user-select:none">'+label+'<span class="pb-grp-chevron" style="font-size:11px;opacity:.85;transition:transform .18s">▾</span></div><div class="pb-grp-body" data-grpid="'+uid+'">';}
    function _pGrpEnd(){return'</div>';}
    function _pSect(label,val,barHtml,dot){return'<div class="pb-sect" style="padding:5px 0 3px;border-bottom:1px solid #f0f0f0"><div style="display:flex;align-items:center;gap:4px;margin-bottom:2px">'+(dot?'<span style="width:6px;height:6px;border-radius:50%;background:'+dot+';flex-shrink:0"></span>':'')+'<span style="font-size:12px;font-weight:600;color:#111827;flex:1">'+label+'</span><span style="font-size:12px;font-weight:700;color:#111827">'+val+'</span></div>'+barHtml+'</div>';}
    function _pSub(label,val,dot,isRem){var c=isRem?'#388c3f':'#6b7280';return'<div style="display:flex;align-items:center;gap:4px;padding:2px 0 2px 10px">'+(dot?'<span style="width:5px;height:5px;border-radius:50%;background:'+dot+';flex-shrink:0"></span>':'')+'<span style="font-size:12px;color:'+c+';flex:1">'+label+'</span><span style="font-size:12px;font-weight:600;color:'+(isRem?'#388c3f':'#111827')+'">'+val+'</span></div>';}
    function _pRef(stlyVal,delta){return'<div style="display:flex;gap:6px;padding:1px 0 0 10px"><span class="wv-ref-tag wv-ref-sdly" style="font-size:10px">STLY '+stlyVal+'</span><span class="wv-ref-tag '+(String(delta).startsWith('+')?'wv-ref-fcst':'wv-ref-sdly')+'" style="font-size:10px">'+delta+'</span></div>';}

    var _pb = '';
    _pb += _filtLabel;

    // ── Close Outs ──
    if (hasPopupCl || LOCKED_DAYS.has(dm+'-'+dd)) {
      _pb += _pGrpStart('Close Outs', '#dc2626', 'co');
      _pb += popupClHtml;
      _pb += _pGrpEnd();
    }

    // ── Daily Metrics ──
    _pb += _pGrpStart('Daily Metrics', _C1, 'dm');
    _pb += _pSect('Occupancy', hotel+'%', _pSbar([{p:to,c:_C1},{p:otherPct,c:_C2}]));
    _pb += _pSub('Travel Distribution Hubs', toRms+' rms / '+to+'%', _C1);
    _pb += _pSub('Other Segments', otherRms+' rms / '+otherPct+'%', _C2);
    _pb += _pSub('STLY', hotelSDLY+'%', _CSTLY);
    _pb += _pSub('Remaining', freeRms+' rms / '+freePct+'%', _CREM, true);
    _pb += _pSect('Online / Offline', onlinePct+'%', _pSbar([{p:onlinePct,c:_C1},{p:offlinePct,c:_C2}]));
    _pb += _pSub('Online', onlinePct+'%', _C1);
    _pb += _pSub('Offline', offlinePct+'%', _C2);
    _pb += _pSect('ADR', '$'+adr, _pBar(adrBar, _C1));
    _pb += _pSub('Hotel ADR', '$'+adr, _C2);
    _pb += _pRef('$'+adrSDLY, '+'+(adr-adrSDLY));
    _pb += _pSect('Revenue', '$'+Math.floor(rev/1000)+'k', _pBar(revBar, _C1));
    _pb += _pSub('Hotel Revenue', '$'+Math.floor(rev/1000)+'k', _C2);
    _pb += _pRef('$'+Math.floor(sdlyR/1000)+'k', '+'+Math.round((rev-sdlyR)/sdlyR*100)+'%');
    _pb += _pGrpEnd();

    // ── More Metrics ──
    _pb += _pGrpStart('More Metrics', _C1, 'mm');
    detRows.forEach(function(r, ri) {
      var bp = Math.min(92, 30 + Math.abs((dm*(ri+3)+dd*7)%55));
      _pb += _pSect(r[0], String(r[1]), _pBar(bp, _C1));
      _pb += _pRef(String(r[2]), String(r[3]));
    });
    _pb += _pGrpEnd();

    // ── Meal Plans ──
    _pb += _pGrpStart('Meal Plans', _C1, 'mp');
    _pb += '<div style="padding:4px 0">'+mealBarHtml+'</div>';
    mealPlans.forEach(function(mp){
      var rn = Math.round(rnSold * mp.pct / 100);
      var seats = Math.round(rn * _popAvgGPR);
      _pb += _pSect(mp.short, mp.pct+'% · '+rn+' rms · '+seats+' seats', _pBar(mp.pct, _C1), mp.color);
    });
    _pb += _pGrpEnd();

    // ── Room Availability ──
    _pb += _pGrpStart('Room Availability' + (_hasAnyFilter ? ' (Filtered)' : ''), _C1, 'ra');
    _pb += rtHTML;
    _pb += _pGrpEnd();

    // ── TO Rates ──
    _pb += _pGrpStart('Tour Operator Rates', _C1, 'to');
    _pb += toRatesHTML;
    _pb += _pGrpEnd();

    var _popupBodyEl = document.getElementById('popupBody');
    _popupBodyEl.innerHTML = _pb;

    // Wire up collapsible group headers
    _popupBodyEl.querySelectorAll('.pb-grp-toggle').forEach(function(hdr) {
      hdr.addEventListener('click', function(e) {
        e.stopPropagation();
        var uid  = this.dataset.grpid;
        var body = _popupBodyEl.querySelector('.pb-grp-body[data-grpid="'+uid+'"]');
        if (!body) return;
        var collapsed = body.style.display === 'none';
        body.style.display = collapsed ? '' : 'none';
        var chev = this.querySelector('.pb-grp-chevron');
        if (chev) chev.style.transform = collapsed ? '' : 'rotate(-90deg)';
      });
    });

    // ── Position popup ──
    const rect = cell.getBoundingClientRect();
    const popW = 320;
    let left = rect.right + 10;
    let top  = rect.top;
    if (left + popW > window.innerWidth - 12) left = rect.left - popW - 10;
    popup.style.left = left + 'px';
    popup.style.top  = top  + 'px';
    popup.style.maxHeight = (window.innerHeight - top - 20) + 'px';
    popup.classList.add('visible');
  }

  function closePopup() { popup.classList.remove('visible'); }

  document.getElementById('calMonths').addEventListener('click', e => {
    const cell = e.target.closest('.cal-day:not(.empty)');
    if (!cell) return;
    const m = +cell.dataset.month, d = +cell.dataset.day;
    const isLocked = cell.classList.contains('locked');

    if (bulkSelectMode && isLocked) {
      var bKey = cell.dataset.month + '-' + cell.dataset.day;
      if (bulkSelected.has(bKey)) { bulkSelected.delete(bKey); cell.classList.remove('bulk-sel'); }
      else { bulkSelected.add(bKey); cell.classList.add('bulk-sel'); }
      var cnt = bulkSelected.size;
      document.getElementById('bulkCount').textContent = cnt === 0 ? 'Click closed dates to select' : cnt + ' date' + (cnt > 1 ? 's' : '') + ' selected';
      document.getElementById('bulkReopenBtn').disabled = cnt === 0;
      return;
    }

    // In range-selection mode: locked days skip range selection but fall through to week view
    if (calSelPicking || (calSelStart && !calSelEnd)) {
      if (!isLocked) {
        if (!calSelStart) {
          calSelStart  = { month: m, day: d };
          calSelPicking = true;
        } else {
          calSelEnd    = { month: m, day: d };
          calSelPicking = false;
          // Open Close Out modal pre-populated with the selected range
          (function() {
            var s = calSelStart, en = calSelEnd;
            var startV = s.month * 100 + s.day, endV = en.month * 100 + en.day;
            var lo = startV <= endV ? s : en, hi = startV <= endV ? en : s;
            var pad = function(n){ return String(n).padStart(2,'0'); };
            var fromStr = '2026-' + pad(lo.month) + '-' + pad(lo.day);
            var toStr   = '2026-' + pad(hi.month) + '-' + pad(hi.day);
            if (typeof window._coOpenModal === 'function') window._coOpenModal(fromStr, toStr, 'cal');
          })();
        }
        applyCalSelection();
        e.stopPropagation();
        return;
      }
      // Locked day clicked in range mode → still navigate to week view (fall through)
    }

    // Eye icon → quick-view popup (all days including closed/partial)
    const eye = e.target.closest('.cell-eye');
    if (eye) {
      openPopup(cell, +eye.dataset.month, +eye.dataset.day);
      e.stopPropagation();
      return;
    }

    // Cell click behaviour depends on view mode
    if (calDisplayView >= 6) {
      // 6M / 12M: always open quick-view popup, never week view
      openPopup(cell, m, d);
    } else {
      // 1M / 2M / 3M: open week view (existing behaviour)
      openWeekView(m, d);
    }
    e.stopPropagation();
  });

  // Hover preview during picking
  document.getElementById('calMonths').addEventListener('mouseover', e => {
    if (!calSelPicking) return;
    const cell = e.target.closest('.cal-day:not(.empty)');
    if (!cell) return;
    applyCalSelection(+cell.dataset.month, +cell.dataset.day);
  });

  close.addEventListener('click', closePopup);
  document.addEventListener('click', e => { if (!popup.contains(e.target)) closePopup(); });
  popup.addEventListener('click', e => {
    // "View Week" button — navigate to week view for this day
    if (e.target.closest('.popup-btn-week')) {
      const dateText = document.getElementById('popupDate')?.textContent || '';
      const match = dateText.match(/(\w+), (\w+) (\d+)/);
      if (match) {
        const MMAP = {January:1,February:2,March:3,April:4,May:5,June:6,July:7,August:8,September:9,October:10,November:11,December:12};
        const goMonth = MMAP[match[2]] || 3;
        const goDay   = +match[3];
        closePopup();
        openWeekView(goMonth, goDay);
        setTimeout(function() {
          const wv = document.getElementById('weekView');
          if (!wv) return;
          const rect = wv.getBoundingClientRect();
          const scrollTop = window.pageYOffset + rect.top - 12;
          window.scrollTo({ top: Math.max(0, scrollTop), behavior: 'smooth' });
        }, 60);
      }
      return; // don't stopPropagation so the document click outside handler also fires (harmless)
    }
    e.stopPropagation();
  });
})();

/* ─── WEEK VIEW ─── */
const DOW_FULL = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
const RT_COLORS = ['#ff5900','#547733','#604f35','#3f3e78','#967ef3','#248b86'];
const RT_NAMES  = ['Standard','Superior','Deluxe','Suite','Jr. Suite','Family'];
const SEG_COLORS = ['#3b82f6','#967EF3','#0ea5e9','#10b981','#f59e0b','#ec4899'];

// Color maps for partial closure chips
const TO_COLORS_MAP = {
  'Sunshine Tours':'#3b82f6', 'Global Adv.':'#967EF3',
  'Beach Hols':'#0ea5e9',     'City Breaks':'#10b981', 'Adventure':'#f59e0b',
};
const RT_NAME_COLORS = {
  'Standard':'#ff5900', 'Superior':'#547733', 'Deluxe':'#604f35',
  'Suite':'#3f3e78',    'Jr. Suite':'#967ef3', 'Family':'#248b86',
};

// Build closures section HTML for weekly column
function buildClosuresHtml(dm, dd) {
  const cl = PARTIAL_CLOSURES[dm + '-' + dd];
  if (!cl || !Array.isArray(cl) || cl.length === 0) return '';
  const lockIcon = '<svg viewBox="0 0 10 12" fill="none" stroke="#dc2626" stroke-width="1.5" width="8" height="10" style="flex-shrink:0"><rect x="1" y="5" width="8" height="7" rx="1"/><path d="M3 5V3.5a2 2 0 0 1 4 0V5"/></svg>';
  const BMAP={ai:'All Inclusive',bb:'Bed & Breakfast',hb:'Half Board',ro:'Room Only',fb:'Full Board'};
  function chip(label, color) {
    return '<span class="wv-cl-chip" style="color:'+color+';background:'+color+'18;border-color:'+color+'3a">'+label+'</span>';
  }
  const rows = cl.map(function(rule, ri) {
    const toChips = rule.tos.length ? rule.tos.map(function(n){ return chip(n, TO_COLORS_MAP[n]||'#dc2626'); }).join('') : chip('All TO','#dc2626');
    const rtChips = rule.roomTypes.length ? rule.roomTypes.map(function(n){ return chip(n, RT_NAME_COLORS[n]||'#b45309'); }).join('') : chip('All Rooms','#b45309');
    const bdChips = rule.boards.length ? rule.boards.map(function(b){ return chip(BMAP[b]||b, BOARD_COLORS[b]||'#7c3aed'); }).join('') : chip('All Plans','#7c3aed');
    return '<div class="wv-cl-row" style="padding:4px 0;border-bottom:1px solid #f3f4f6">'
      +'<span class="wv-cl-cat" style="font-size:9px;color:#9ca3af;min-width:16px">'+(ri+1)+'.</span>'
      +'<div style="display:flex;flex-wrap:wrap;gap:3px">'
      +toChips+rtChips+bdChips
      +'</div></div>';
  });
  return '<div class="wv-closures-wrap">'
    + '<div class="wv-closures-title">'+lockIcon+' Closed Out</div>'
    + rows.join('')
    + '</div>';
}

let wvYear = 2026; let wvMonth = 3; let wvWeekStart = 1; // day of month for week start

// Weekly view range selection
let wvSelStart   = null;  // { month, day }
let wvSelEnd     = null;  // { month, day }
let wvSelPicking = false; // true after first click, awaiting end

// Weekly section collapse state (persists across rebuilds)
const wvCollapsed = { daily: false, detailed: false, meals: false, avail: false, availAlloc: false, toRates: false, promos: false, mealsSummary: true };
// LY comparison mode: 'sdly' | 'final-ly' | 'forecast'
let wvCompMode = 'sdly';
// Active weekly content tab
let wvActiveTab = 'occupancy';
// Weekly group-by: 'combined' | 'roomType' | 'boardType'
let wvGroupBy = 'dailyB';
let wvSegMode = 'combined'; // 'combined' | 'individual'
let wvCompare = new Set();  // multi-select Set of active compares: 'stly' | 'ly' | 'fcst'

function wvCmpDdToggle(e) {
  if (e) e.stopPropagation();
  var menu = document.getElementById('wvCmpDdMenu');
  var btn  = document.getElementById('wvCmpDdBtn');
  if (!menu || !btn) return;
  var opening = !menu.classList.contains('open');
  menu.classList.toggle('open', opening);
  btn.classList.toggle('open', opening);
}
function wvCmpDdToggle2(e) {
  if (e) e.stopPropagation();
  var menu = document.getElementById('wvCmpDdMenu2');
  var btn  = document.getElementById('wvCmpDdBtn2');
  if (!menu || !btn) return;
  var opening = !menu.classList.contains('open');
  menu.classList.toggle('open', opening);
  btn.classList.toggle('open', opening);
}
// Close both dropdowns when clicking outside
document.addEventListener('click', function(e) {
  ['wvCmpDd','wvCmpDd2'].forEach(function(id) {
    var dd = document.getElementById(id);
    if (dd && !dd.contains(e.target)) {
      var sfx = id === 'wvCmpDd' ? '' : '2';
      var m = document.getElementById('wvCmpDdMenu'+sfx), b = document.getElementById('wvCmpDdBtn'+sfx);
      if (m) m.classList.remove('open');
      if (b) b.classList.remove('open');
    }
  });
});
function wvSyncCmpDd() {
  var _names = {stly:'STLY', ly:'LY', fcst:'Fcst'};
  var labelTxt = wvCompare.size === 0
    ? 'Compare'
    : ['stly','ly','fcst'].filter(function(k){ return wvCompare.has(k); }).map(function(k){ return _names[k]; }).join(', ');
  // Sync both dropdown instances
  ['wvCmpDdMenu', 'wvCmpDdMenu2'].forEach(function(menuId) {
    document.querySelectorAll('#'+menuId+' .wv-cmp-dd-item').forEach(function(item) {
      var k = item.dataset.cmp;
      var active = (k === 'none') ? wvCompare.size === 0 : wvCompare.has(k);
      item.classList.toggle('active', active);
      var chk = item.querySelector('.wv-cmp-chk');
      if (chk) chk.checked = active;
    });
  });
  var lbl1 = document.getElementById('wvCmpDdLabel');
  if (lbl1) lbl1.textContent = labelTxt;
  var lbl2 = document.getElementById('wvCmpDdLabel2');
  if (lbl2) lbl2.textContent = labelTxt;
}
function wvSetCompare(val) {
  if (val === 'none') {
    wvCompare.clear();
  } else {
    if (wvCompare.has(val)) wvCompare.delete(val);
    else wvCompare.add(val);
  }
  wvSyncCmpDd();
  buildWeekGrid(wvMonth, wvWeekStart, wvWeekStart);
}

// "/ compareVal" inline suffix appended to main value text
function wvCmpValSuffix(stlyStr, lyStr, fcstStr) {
  if (wvCompare.size === 0) return '';
  var parts = [];
  if (wvCompare.has('stly') && stlyStr && stlyStr !== 'null') parts.push(stlyStr);
  if (wvCompare.has('ly')   && lyStr   && lyStr   !== 'null') parts.push(lyStr);
  if (wvCompare.has('fcst') && fcstStr && fcstStr !== 'null') parts.push(fcstStr);
  return parts.map(function(s){ return '<span class="wv-cmp-sep"> / </span><span class="wv-cmp-val-txt">' + s + '</span>'; }).join('');
}

// Multi-compare inline suffix — shows "/ STLY:val / LY:val" etc. for all active
function _wvMultiCmpSfx(curr, stlyVal, lyVal, fcstVal, fmtFn) {
  if (wvCompare.size === 0) return '';
  var _ORDER = [['stly','STLY',stlyVal],['ly','LY',lyVal],['fcst','Fc',fcstVal]];
  return _ORDER.filter(function(t){ return wvCompare.has(t[0]) && t[2] != null; }).map(function(t) {
    var s = fmtFn(t[2]);
    var clr = '#9ca3af';
    var c = parseFloat(curr), p = parseFloat(t[2]);
    if (!isNaN(c) && !isNaN(p)) { if (c > p) clr = '#16a34a'; else if (c < p) clr = '#dc2626'; }
    return '<span class="wv-cmp-sep"> / </span><span class="wv-cmp-val-txt" style="color:'+clr+'">'+t[1]+':'+s+'</span>';
  }).join('');
}

// Multi-compare trend badges
function _wvMultiTrendBadge(curr, stlyVal, lyVal, fcstVal) {
  if (wvCompare.size === 0) return '';
  var _ORDER = [['stly','STLY',stlyVal],['ly','LY',lyVal],['fcst','Fc',fcstVal]];
  var base = 'font-size:11px;font-weight:600;border-radius:4px;padding:2px 5px;flex-shrink:0;white-space:nowrap;margin-left:3px;display:inline-flex;align-items:center;gap:2px;line-height:1.2;';
  return _ORDER.filter(function(t){ return wvCompare.has(t[0]) && t[2]!=null && !isNaN(curr) && !isNaN(t[2]) && t[2]!==0; }).map(function(t) {
    var diff = curr - t[2], pct = Math.round(Math.abs(diff)/Math.abs(t[2])*100);
    var clr = diff>0?'#15803d':diff<0?'#b91c1c':'#6b7280', bg = diff>0?'#dcfce7':diff<0?'#fee2e2':'#f3f4f6';
    var arrow = diff>0?'▲':diff<0?'▼':'';
    return '<span style="'+base+'color:'+clr+';background:'+bg+'">'+arrow+pct+'%<span style="opacity:.65;font-weight:500"> vs '+t[1]+'</span></span>';
  }).join('');
}

// Right-side header block: "mainVal / cmpVal" (no chip)
function wvHdrRight(mainVal, stlyStr, lyStr, fcstStr) {
  const valSuffix = wvCmpValSuffix(stlyStr, lyStr, fcstStr);
  return '<div class="wv-hdr-right">'
    + '<span class="wv-occ-total">' + mainVal + valSuffix + '</span>'
    + '</div>';
}
// Tracks which TO detail panels are open (key: 'tos_ri_bi' for BT, 'rtos_ri' for RT)
const wvTosOpen = {};

// Re-open a range of days (removes from LOCKED_DAYS)
function reopenRange(start, end) {
  const startV = start.month * 100 + start.day;
  const endV   = end.month * 100 + end.day;
  const loD = startV <= endV ? start : end;
  const hiD = startV <= endV ? end   : start;
  let cur = new Date(2026, loD.month - 1, loD.day);
  const endDate = new Date(2026, hiD.month - 1, hiD.day);
  while (cur <= endDate) {
    LOCKED_DAYS.delete(`${cur.getMonth()+1}-${cur.getDate()}`);
    cur.setDate(cur.getDate() + 1);
  }
  renderCalendar();
  buildWeekGrid(wvMonth, wvWeekStart, wvWeekStart);
}

function clearWvSelection() {
  wvSelStart = null; wvSelEnd = null; wvSelPicking = false;
  // Reset Close button appearance
  const closeBtn = document.getElementById('wvCloseOutBtn');
  if (closeBtn) { closeBtn.style.background = ''; closeBtn.style.color = ''; closeBtn.style.boxShadow = ''; }
  const right = document.querySelector('.wv-topbar-right');
  if (right) right.classList.remove('range-mode');
  buildWeekGrid(wvMonth, wvWeekStart, wvWeekStart);
}

function getWeekDays(year, month, startDay) {
  // Returns array of 7 {month, day} objects starting from startDay
  const days = [];
  const dim = [0,31,28,31,30,31,30,31,31,30,31,30,31];
  let m = month, d = startDay;
  for (let i = 0; i < 7; i++) {
    days.push({ year, month: m, day: d });
    d++;
    if (d > dim[m]) { d = 1; m++; if (m > 12) m = 1; }
  }
  return days;
}

function renderWeekView(month, day) {
  const weekStartDay = day;
  wvMonth = month;
  wvWeekStart = weekStartDay;

  // Sync calendar filters → weekly so selections persist across views
  syncFiltersCalToWv();
  applyFilterUI('wvFiltersDropdown');
  _syncPickupBtnUI('wv');

  const calSection = document.getElementById('demand-calendar');
  const wvSection  = document.getElementById('weekView');
  if (calSection) calSection.style.display = 'none';
  if (wvSection)  wvSection.classList.add('visible');
  var backArrow = document.getElementById('wvBack');
  if (backArrow) backArrow.style.display = 'inline-flex';
  var hdrCtr = document.getElementById('wvHeaderCenter');
  if (hdrCtr) hdrCtr.style.display = 'flex';
  var moBar = document.getElementById('moGroupbyBar');
  if (moBar) moBar.style.display = 'none';

  buildWeekGrid(month, weekStartDay, day);
}

function openWeekView(month, day) { renderWeekView(month, day); _updateAccBtnState(); }

/* ── Metrics selector state ── */
const wvMetricState = {
  capacity: true, adr: true, revenue: true, onlineOffline: true, roomTypes: true,
  avail: true, availAlloc: true, toRates: true,
  dm_rnSold: true, dm_pickup: true, dm_pickup_0: true, dm_pickup_1: true, dm_pickup_2: true,
  dm_avgAdults: true, dm_avgChildren: true, dm_totalAdults: true, dm_totalChildren: true, dm_trevpar: true,
  dm_availRooms: true, dm_availGuar: true,
  dm_avgLos: true, dm_avgLeadTime: true, dm_totalGuests: true,
  bizMix: true, mealsSummary: true,
  cmp_sdly: true, cmp_final_ly: true, cmp_forecast: true, cmp_hotel: true,
  dm_closeouts: true, dm_co_rooms: true, dm_co_boards: true, dm_co_tos: true,
};

function wvAcc(title, section, bodyHtml, badge) {
  const collapsed = wvCollapsed[section];
  // SVG chevrons — up (open) / down (closed) — Figma expand_less / expand_more style
  const chevUp   = '<svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="18 15 12 9 6 15"/></svg>';
  const chevDown = '<svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"/></svg>';
  return '<div class="wv-acc-sect' + (collapsed ? '' : ' wv-acc-open') + '">'
    + '<div class="wv-acc-hdr" data-section="' + section + '">'
    + '<span class="wv-acc-chev">' + (collapsed ? chevDown : chevUp) + '</span>'
    + '<span class="wv-acc-title">' + title + '</span>'
    + (badge ? '<span class="wv-acc-badge">' + badge + '</span>' : '')
    + '</div>'
    + '<div class="wv-acc-body' + (collapsed ? ' wv-body-hidden' : '') + '">'
    + bodyHtml
    + '</div>'
    + '</div>';
}

/* ── Close-out metadata helpers ─────────────────────────────────────────── */
function coFmtDate(iso) {
  if (!iso) return '';
  var d = new Date(iso);
  var MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  var h = d.getHours(), m = d.getMinutes();
  return d.getDate() + ' ' + MONTHS[d.getMonth()] + ', ' + (h < 10 ? '0'+h : h) + ':' + (m < 10 ? '0'+m : m);
}
function coInitials(name) {
  return (name || '').split(' ').map(function(p){ return p[0]||''; }).join('').toUpperCase().slice(0,2);
}
function coMetaHtml(appliedBy, appliedAt, extraStyle) {
  if (!appliedBy) return '';
  var ini = coInitials(appliedBy);
  var dt  = coFmtDate(appliedAt);
  return '<div class="co-rule-meta" style="'+(extraStyle||'')+'">'
    + '<span class="co-meta-avatar">'+ini+'</span>'
    + '<span class="co-meta-text"><strong>'+appliedBy+'</strong>'+(dt ? ' · '+dt : '')+'</span>'
    + '</div>';
}

/* ── Close-out "+N more" toggle ──────────────────────────────────────────── */
window.coToggleMore = function(uid) {
  var moreSpan = document.getElementById('coh_' + uid);
  var plusSpan = document.querySelector('[onclick="coToggleMore(\'' + uid + '\')"]');
  if (!moreSpan || !plusSpan) return;
  var isHidden = moreSpan.style.display === 'none';
  moreSpan.style.display = isHidden ? 'inline' : 'none';
  plusSpan.textContent = isHidden ? '− less' : '+' + plusSpan.dataset.n + ' more';
};

/* ── Room Type → Meal Plan nested accordion renderer ─────────────────────
   For each room type, shows all board types as sub-accordions.
   Each board type shows the full combined metric set, scaled to that segment.
*/
function buildRoomTypeBoardView(dm, dd, hotel, to, adr, rev, v) {
  const isFullyLocked = LOCKED_DAYS.has(dm + '-' + dd);
  const rules = PARTIAL_CLOSURES[dm + '-' + dd] || [];
  const lockSvg = '<svg viewBox="0 0 10 12" fill="none" stroke="currentColor" stroke-width="1.6" width="11" height="13"><rect x="1" y="5" width="8" height="7" rx="1"/><path d="M3 5V3.5a2 2 0 0 1 4 0V5"/></svg>';
  const noCloseSvg = '<svg viewBox="0 0 14 14" fill="none" stroke="#15803d" stroke-width="1.6" width="12" height="12"><path d="M2 7l4 4 6-6"/></svg>';
  const BMAP = {ai:'All Inclusive',bb:'Bed & Breakfast',hb:'Half Board',ro:'Room Only',fb:'Full Board'};

  if (isFullyLocked) {
    var _lm = LOCKED_DAYS_META[dm + '-' + dd] || {};
    return '<div class="wv-closeouts-wrap">'
      + '<div class="co-full-banner">'+lockSvg+' Full Day Close Out — All inventory closed'
      + coMetaHtml(_lm.appliedBy, _lm.appliedAt, 'margin-top:6px;opacity:.85')
      + '</div>'
      + '</div>';
  }

  if (rules.length === 0) {
    return '<div class="wv-closeouts-wrap">'
      + '<div class="co-section-title">Close Outs</div>'
      + '<div class="co-all-open">'+noCloseSvg+' No close outs — all channels open</div>'
      + '</div>';
  }

  // ── Dimension chip helpers ────────────────────────────────────────────
  function chip(label, clr, bg) {
    return '<span class="co-chip" style="background:'+(bg||clr+'18')+';color:'+clr+';border-color:'+clr+'55">'+label+'</span>';
  }
  function toChip(name) { return chip(name, TO_COLORS_MAP[name]||'#dc2626'); }
  function rtChip(name) { return chip(name, RT_NAME_COLORS[name]||'#b45309'); }
  function bdChip(b)    { return chip(BMAP[b]||b, '#7c3aed'); }
  function allChip(label) { return chip(label, '#6b7280', '#f3f4f6'); }

  var _coUidSeq = 0;
  function chipsMore(items, chipFn, max) {
    if (items.length <= max) return items.map(chipFn).join('');
    var uid = 'cob_' + dm + dd + '_' + (++_coUidSeq);
    var shown  = items.slice(0, max).map(chipFn).join('');
    var hidden = items.slice(max).map(chipFn).join('');
    return shown
      + '<span id="coh_' + uid + '" style="display:none">' + hidden + '</span>'
      + '<span class="co-more-pill" onclick="coToggleMore(\'' + uid + '\')" data-n="' + (items.length - max) + '">+' + (items.length - max) + ' more</span>';
  }

  const ruleCards = rules.map(function(rule, ri) {
    const hasTO = rule.tos.length > 0;
    const hasRT = rule.roomTypes.length > 0;
    const hasBD = rule.boards.length > 0;

    const toPart = hasTO
      ? chipsMore(rule.tos, toChip, 2)
      : allChip('All Operators');
    const rtPart = hasRT
      ? chipsMore(rule.roomTypes, rtChip, 2)
      : allChip('All Room Types');
    const bdPart = hasBD
      ? chipsMore(rule.boards, bdChip, 2)
      : allChip('All Meal Plans');

    return '<div class="co-rule-card">'
      + '<div class="co-rule-num">'+lockSvg+' Strategy ' + (ri+1) + coMetaHtml(rule.appliedBy, rule.appliedAt, 'margin-left:auto') + '</div>'
      + '<div class="co-rule-row">'
      + '<div class="co-rule-dim"><span class="co-rule-dim-lbl">Operator</span><div class="co-chips co-chips-inline">'+toPart+'</div></div>'
      + '<div class="co-rule-sep">+</div>'
      + '<div class="co-rule-dim"><span class="co-rule-dim-lbl">Room Type</span><div class="co-chips co-chips-inline">'+rtPart+'</div></div>'
      + '<div class="co-rule-sep">+</div>'
      + '<div class="co-rule-dim"><span class="co-rule-dim-lbl">Meal Plan</span><div class="co-chips co-chips-inline">'+bdPart+'</div></div>'
      + '</div></div>';
  }).join('');

  return '<div class="wv-closeouts-wrap">'
    + '<div class="co-section-title">Close Outs</div>'
    + ruleCards
    + '</div>';
}


// ── Close Out Report View — grouped-by-day card layout ───────────────────────
function buildCoReportView(days) {
  var BMAP    = {ai:'All Inclusive',bb:'Bed & Breakfast',hb:'Half Board',ro:'Room Only',fb:'Full Board'};
  var DOW2    = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
  var MNAMES2 = ['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  var TODAY_R = new Date(2026, 2, 9);

  var lockSvg  = '<svg viewBox="0 0 10 12" fill="none" stroke="currentColor" stroke-width="1.6" width="10" height="12"><rect x="1" y="5" width="8" height="7" rx="1"/><path d="M3 5V3.5a2 2 0 0 1 4 0V5"/></svg>';
  var checkSvg = '<svg viewBox="0 0 14 14" fill="none" stroke="#15803d" stroke-width="2" width="11" height="11"><path d="M2 7l4 4 6-6"/></svg>';

  // ── Chip helpers ─────────────────────────────────────────────────────────────
  function chip(label, clr, bg) {
    return '<span class="co-rpt-chip" style="background:'+(bg||clr+'18')+';color:'+clr+';border-color:'+clr+'44">'+label+'</span>';
  }
  function allDimChip(label) { return chip(label, '#6b7280', '#f3f4f6'); }

  function chipsMore(items, mapFn, max, uid) {
    if (items.length <= max) return items.map(mapFn).join('');
    var shown  = items.slice(0, max).map(mapFn).join('');
    var hidden = items.slice(max).map(mapFn).join('');
    return shown
      + '<span id="coh_'+uid+'" style="display:none">'+hidden+'</span>'
      + '<span class="co-more-pill" onclick="coToggleMore(\''+uid+'\')" data-n="'+(items.length-max)+'">+'+(items.length-max)+' more</span>';
  }

  // ── Day cards ─────────────────────────────────────────────────────────────────
  var cardsHtml = days.map(function(dv) {
    var dm = dv.month, day = dv.day, key = dm+'-'+day;
    var isFullyLocked = LOCKED_DAYS.has(key);
    var rules = PARTIAL_CLOSURES[key] || [];
    var dt  = new Date(2026, dm-1, day);
    var dow = DOW2[dt.getDay()];
    var dba = Math.round((dt - TODAY_R) / 86400000);
    var isToday = dba === 0;
    var evts = (typeof CAL_EVENTS !== 'undefined' && CAL_EVENTS[key]) ? CAL_EVENTS[key] : null;

    var stateClass = isFullyLocked ? 'co-rpt-day--locked' : rules.length > 0 ? 'co-rpt-day--partial' : 'co-rpt-day--open';
    var todayClass = isToday ? ' co-rpt-day--today' : '';

    // ── Day header ──────────────────────────────────────────────────────────
    var dbaLabel = dba === 0 ? 'Today' : dba > 0 ? dba+'d DBA' : '';
    var evtDot   = evts ? '<span class="co-rpt-evt-dot" title="'+evts.map(function(e){return e.name;}).join(', ')+'"></span>' : '';
    var badge = isFullyLocked
      ? '<span class="co-rpt-badge co-rpt-badge--locked">'+lockSvg+' Full Day Closed</span>'
      : rules.length > 0
        ? '<span class="co-rpt-badge co-rpt-badge--partial">'+rules.length+' rule'+(rules.length>1?'s':'')+'</span>'
        : '<span class="co-rpt-badge co-rpt-badge--open">'+checkSvg+' Open</span>';

    var hdr = '<div class="co-rpt-day-hdr">'
      + '<div class="co-rpt-day-date">'+MNAMES2[dm]+' '+day+'</div>'
      + '<div class="co-rpt-day-sub">'+dow+(dbaLabel ? ' <span class="co-rpt-dba">· '+dbaLabel+'</span>' : '')+evtDot+'</div>'
      + badge
      + '</div>';

    // ── Day body ─────────────────────────────────────────────────────────────
    var body;
    if (isFullyLocked) {
      var lm = LOCKED_DAYS_META[key] || {};
      body = '<div class="co-rpt-full-lock">'
        + lockSvg + ' All inventory closed for this date'
        + (lm.appliedBy ? coMetaHtml(lm.appliedBy, lm.appliedAt, 'margin-top:5px') : '')
        + '</div>';
    } else if (rules.length === 0) {
      body = '<div class="co-rpt-all-open">'+checkSvg+' No restrictions — all channels open</div>';
    } else {
      // Column-label header (shown once per card, above the strategy rows)
      body = '<div class="co-rpt-col-hdr">'
        + '<div class="co-rpt-col-hdr-num"></div>'
        + '<div class="co-rpt-col-hdr-dim">Operator</div>'
        + '<div class="co-rpt-col-hdr-sep"></div>'
        + '<div class="co-rpt-col-hdr-dim">Room Type</div>'
        + '<div class="co-rpt-col-hdr-sep"></div>'
        + '<div class="co-rpt-col-hdr-dim">Meal Plan</div>'
        + '<div class="co-rpt-col-hdr-meta">Applied by</div>'
        + '</div>';

      body += rules.map(function(rule, ri) {
        var uid = 'rpt_'+dm+'_'+day+'_'+ri;
        var toFn = function(n) { return chip(n, TO_COLORS_MAP[n]||'#dc2626'); };
        var rtFn = function(n) { return chip(n, RT_NAME_COLORS[n]||'#b45309'); };
        var bdFn = function(b) { return chip(BMAP[b]||b, '#7c3aed'); };

        var toPart = rule.tos.length        ? chipsMore(rule.tos,        toFn, 2, uid+'_to') : allDimChip('All Operators');
        var rtPart = rule.roomTypes.length  ? chipsMore(rule.roomTypes,  rtFn, 2, uid+'_rt') : allDimChip('All Room Types');
        var bdPart = rule.boards.length     ? chipsMore(rule.boards,     bdFn, 2, uid+'_bd') : allDimChip('All Meal Plans');

        var metaPart = rule.appliedBy
          ? '<div class="co-rule-meta">'
            + '<span class="co-meta-avatar" style="width:16px;height:16px;font-size:7.5px">'+coInitials(rule.appliedBy)+'</span>'
            + '<span class="co-meta-text"><strong>'+rule.appliedBy+'</strong>'+(rule.appliedAt ? '<br><span style="font-weight:400">'+coFmtDate(rule.appliedAt)+'</span>' : '')+'</span>'
            + '</div>'
          : '';

        var STRAT_COLORS = ['#dc2626','#b45309','#7c3aed','#0891b2','#16a34a'];
        var sClr = STRAT_COLORS[ri % STRAT_COLORS.length];

        return '<div class="co-rpt-strat-row">'
          + '<div class="co-rpt-strat-num" style="background:'+sClr+'18;color:'+sClr+';border-color:'+sClr+'33">'+(ri+1)+'</div>'
          + '<div class="co-rpt-strat-dim"><div class="co-rpt-chips">'+toPart+'</div></div>'
          + '<div class="co-rpt-strat-sep">+</div>'
          + '<div class="co-rpt-strat-dim"><div class="co-rpt-chips">'+rtPart+'</div></div>'
          + '<div class="co-rpt-strat-sep">+</div>'
          + '<div class="co-rpt-strat-dim"><div class="co-rpt-chips">'+bdPart+'</div></div>'
          + '<div class="co-rpt-strat-meta">'+metaPart+'</div>'
          + '</div>';
      }).join('');
    }

    return '<div class="co-rpt-day-card '+stateClass+todayClass+'">'
      + hdr
      + '<div class="co-rpt-day-body">'+body+'</div>'
      + '</div>';
  }).join('');

  return '<div class="co-rpt-list">'+cardsHtml+'</div>';
}


// ── Close-Out Heat Map by Room Type ──────────────────────────────────────────
function buildCoHeatmap(days) {
  var RT = ['Standard','Superior','Deluxe','Suite','Jr. Suite','Family'];
  var DOW_S = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
  var MNAMES_S = ['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  var lockSvg = '<svg viewBox="0 0 10 12" fill="none" stroke="currentColor" stroke-width="1.4" width="9" height="11"><rect x="1" y="5" width="8" height="7" rx="1"/><path d="M3 5V3.5a2 2 0 0 1 4 0V5"/></svg>';

  // For each day+room type, determine close-out status:
  // 'full' = full day locked, 'closed' = room type specifically closed, 'partial' = day has closures but not this RT, 'open' = no closures
  function getStatus(dm, dd, rt) {
    var key = dm + '-' + dd;
    if (LOCKED_DAYS.has(key)) return 'full';
    var rules = PARTIAL_CLOSURES[key] || [];
    if (rules.length === 0) return 'open';
    // Check if any rule targets this room type (or all room types)
    for (var i = 0; i < rules.length; i++) {
      var r = rules[i];
      if (r.roomTypes.length === 0) return 'closed'; // applies to all room types
      if (r.roomTypes.indexOf(rt) >= 0) return 'closed';
    }
    return 'partial'; // day has closures but not for this room type
  }

  var cellClasses = { full:'co-hm-full', closed:'co-hm-closed', partial:'co-hm-partial', open:'co-hm-open' };
  var cellTitles  = { full:'Full day closed', closed:'Closed out', partial:'Other closures (open for this type)', open:'Open' };

  // Build header row (days)
  var hdrCells = '<div class="co-hm-corner">Room Type</div>' + days.map(function(dv) {
    var dt = new Date(2026, dv.month - 1, dv.day);
    var isToday = dv.month === 3 && dv.day === 9;
    return '<div class="co-hm-hdr' + (isToday ? ' co-hm-today' : '') + '">'
      + '<span class="co-hm-dow">' + DOW_S[dt.getDay()] + '</span>'
      + '<span class="co-hm-date">' + MNAMES_S[dv.month] + ' ' + dv.day + '</span>'
      + '</div>';
  }).join('');

  // Build rows (one per room type)
  var bodyRows = RT.map(function(rt) {
    var rtClr = RT_NAME_COLORS[rt] || '#6b7280';
    var cells = days.map(function(dv) {
      var status = getStatus(dv.month, dv.day, rt);
      var icon = status === 'full' || status === 'closed' ? lockSvg : '';
      return '<div class="co-hm-cell ' + cellClasses[status] + '" title="' + rt + ' — ' + MNAMES_S[dv.month] + ' ' + dv.day + ': ' + cellTitles[status] + '">'
        + icon
        + '</div>';
    }).join('');
    return '<div class="co-hm-row">'
      + '<div class="co-hm-label"><span class="co-hm-dot" style="background:' + rtClr + '"></span>' + rt + '</div>'
      + cells
      + '</div>';
  }).join('');

  // Legend
  var legend = '<div class="co-hm-legend">'
    + '<span class="co-hm-leg"><span class="co-hm-leg-sw co-hm-full"></span>Full Close Out</span>'
    + '<span class="co-hm-leg"><span class="co-hm-leg-sw co-hm-closed"></span>Room Type Closed</span>'
    + '<span class="co-hm-leg"><span class="co-hm-leg-sw co-hm-partial"></span>Other Closures</span>'
    + '<span class="co-hm-leg"><span class="co-hm-leg-sw co-hm-open"></span>Open</span>'
    + '</div>';

  return '<div class="co-hm-wrap">'
    + '<div class="co-hm-title">' + lockSvg + ' Close Out Heat Map — Room Types</div>'
    + '<div class="co-hm-grid" style="grid-template-columns:140px repeat(' + days.length + ',1fr)">'
    + hdrCells + bodyRows
    + '</div>'
    + legend
    + '</div>';
}

// ── Monthly view day close-out checkboxes ────────────────────────────────────
var _moSelectedDays = new Set(); // ISO date strings selected for close-out in monthly view

window.moDayCheck = function(dateStr, cb) {
  if (cb.checked) _moSelectedDays.add(dateStr);
  else _moSelectedDays.delete(dateStr);
  _syncCloseOutBtn();
};

window.moOpenCloseOut = function() {
  var dates = Array.from(_moSelectedDays).sort();
  if (!dates.length) return;
  if (typeof window._coOpenModalDays === 'function') {
    window._coOpenModalDays(dates, 'cal');
  } else if (typeof window._coOpenModal === 'function') {
    window._coOpenModal(dates[0], dates[dates.length - 1], 'cal');
  }
};

// Smart single button: pre-fill dates if cells selected, else open empty
window.moSmartClose = function() {
  var dates = Array.from(_moSelectedDays).sort();
  if (dates.length) {
    if (typeof window._coOpenModalDays === 'function') window._coOpenModalDays(dates, 'cal');
    else if (typeof window._coOpenModal === 'function') window._coOpenModal(dates[0], dates[dates.length - 1], 'cal');
  } else {
    if (typeof window._coOpenModal === 'function') window._coOpenModal('', '', 'cal');
  }
};

window.wvSmartClose = function() {
  var wvDates = Array.from(_wvSelectedDays).sort();
  var wbDates = Array.from(_wbSelectedDays).sort();
  var dates = wvDates.concat(wbDates).filter(function(v,i,a){ return a.indexOf(v)===i; }).sort();
  if (dates.length) {
    if (typeof window._coOpenModalDays === 'function') window._coOpenModalDays(dates, 'wv');
    else if (typeof window._coOpenModal === 'function') window._coOpenModal(dates[0], dates[dates.length - 1], 'wv');
  } else {
    if (typeof window._coOpenModal === 'function') window._coOpenModal('', '', 'wv');
  }
};

// ── Weekly view day close-out checkboxes ─────────────────────────────────────
var _wvSelectedDays = new Set(); // ISO date strings selected for close-out in weekly view

window.wvDayCheck = function(dateStr, cb) {
  if (cb.checked) _wvSelectedDays.add(dateStr);
  else _wvSelectedDays.delete(dateStr);
  _syncCloseOutBtn();
};

window.wvOpenCloseOut = function() {
  var dates = Array.from(_wvSelectedDays).sort();
  if (!dates.length) return;
  if (typeof window._coOpenModalDays === 'function') {
    window._coOpenModalDays(dates, 'wv');
  } else if (typeof window._coOpenModal === 'function') {
    window._coOpenModal(dates[0], dates[dates.length - 1], 'wv');
  }
};

// No-op — single button is always enabled; kept for compatibility
function _syncCloseOutBtn() {}

// ── Daily B View ─────────────────────────────────────────────────────────────
var _wbCollapsed    = {};   // shared collapse state (used by both HTML fallback and AG Grid)
var _wbAllIds       = [];   // all toggleable row IDs in Daily B (populated on each render)
var _wbSelectedDays = new Set(); // ISO date strings selected for close-out in Daily B
var _wbGroupOrder   = null; // null = default; array of group keys for custom Daily B order

var WB_GROUPS_DEF = [
  { key: 'g_closeouts', lbl: 'Close Outs',      clr: '#dc2626' },
  { key: 'g_daily',   lbl: 'Daily Metrics',    clr: '#006461' },
  { key: 'g_more',    lbl: 'More Metrics',     clr: '#2e65e8' },
  { key: 'g_meals',   lbl: 'Meal Plans',       clr: '#f59e0b' },
  { key: 'g_biz',     lbl: 'Business Mix',     clr: '#7c3aed' },
  { key: 'g_avail',   lbl: 'Room Availability',clr: '#0891b2' },
  { key: 'g_torates', lbl: 'Travel Co. Rates', clr: '#0f766e' },
];
var _dailyBGridApi = null;
var _dbAllRows     = [];
var _dbGrpRenderrs = [];

function _getDBVisibleRows() {
  return _dbAllRows.filter(function(r) {
    if (r.type === 'grp') return true;
    if (_wbCollapsed[r.grpKey]) return false;
    if (r.type === 'sect') return true;
    if (_wbCollapsed[r.sectKey]) return false;
    return true;
  }).map(function(r) {
    return { _type:r.type, _lbl:r.lbl, _clr:r.clr||'#374151',
             _dot:r.dot||null, _isRem:r.isRem||false,
             _grpKey:r.grpKey||null, _sectKey:r.sectKey||null, _fn:r.fn||null,
             _noChev:r.noChev||false };
  });
}
function _toggleDBGrp(grpKey) {
  _wbCollapsed[grpKey] = !_wbCollapsed[grpKey];
  _dbGrpRenderrs.forEach(function(gr){ if (gr._grpKey===grpKey) gr._syncChev(); });
  if (_dailyBGridApi) _dailyBGridApi.setGridOption('rowData', _getDBVisibleRows());
}
function _toggleDBSect(sectKey) {
  _wbCollapsed[sectKey] = !_wbCollapsed[sectKey];
  if (_dailyBGridApi) _dailyBGridApi.setGridOption('rowData', _getDBVisibleRows());
}

function wbToggle(id) {
  _wbCollapsed[id] = !_wbCollapsed[id];
  buildWeekGrid(wvMonth, wvWeekStart, wvWeekStart);
}
function wbSetAll(collapse) {
  // Set ALL tracked IDs and any keys already in the collapse map
  _wbAllIds.forEach(function(id) { _wbCollapsed[id] = collapse; });
  for (var k in _wbCollapsed) {
    if (_wbCollapsed.hasOwnProperty(k)) _wbCollapsed[k] = collapse;
  }
  buildWeekGrid(wvMonth, wvWeekStart, wvWeekStart);
}

// ── Daily B day close-out selection ────────────────────────────────────────
window.wbDayToggle = function(dateStr) {
  if (_wbSelectedDays.has(dateStr)) {
    _wbSelectedDays.delete(dateStr);
  } else {
    _wbSelectedDays.add(dateStr);
  }
  // Update just this header cell's selected class (no full rebuild)
  var cell = document.querySelector('.wb-hdr-cell[data-wb-date="' + dateStr + '"]');
  if (cell) cell.classList.toggle('wb-hdr-selected', _wbSelectedDays.has(dateStr));
  _syncCloseOutBtn();
};

window.wbOpenCloseOut = function() {
  var dates = Array.from(_wbSelectedDays).sort();
  if (!dates.length) return;
  if (typeof window._coOpenModalDays === 'function') {
    window._coOpenModalDays(dates, 'wv');
  } else if (typeof window._coOpenModal === 'function') {
    window._coOpenModal(dates[0], dates[dates.length - 1], 'wv');
  }
};

function buildDailyBView(days, month, activeDay) {
  var DOW_SHORT  = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
  var WV_CAP     = 250;
  var RT_NAMES   = ['Standard','Superior','Deluxe','Suite','Jr. Suite','Family'];
  var RT_CAPS    = [51,36,27,12,15,9];
  var TO_NAMES   = ['Sunshine Tours','Global Adv.','Beach Hols','City Breaks','Adventure'];
  var TO_COLORS  = ['#3b82f6','#967EF3','#0ea5e9','#10b981','#f59e0b'];

  // ── Per-day values ────────────────────────────────────────────────────────
  var dd7 = days.map(function(dv) {
    var dm = dv.month, dd = dv.day;
    var hh = getOccupancy(dm, dd); var hotel = hh.hotel, to = hh.to;
    var adr   = 150 + Math.abs((dm*47+dd*31)%130);
    var v     = Math.abs((dm*127+dd*53+dm*dd*7+dd*dd*3))%100;
    var toAdr = Math.max(80, adr - 20 - Math.abs((dm*3+dd*7)%15));
    var toRn  = Math.round(WV_CAP * to / 100);
    var hnRn  = Math.round(WV_CAP * hotel / 100);
    var toRev = Math.floor(toRn * toAdr);
    var hnRev = Math.floor(hnRn * adr);
    var otherPct = Math.max(0, hotel - to);
    var otherRms = Math.round(WV_CAP * otherPct / 100);
    var freeRms  = WV_CAP - toRn - otherRms;
    var onlinePct = Math.max(30, Math.min(80, 45 + Math.abs((dm*13+dd*7)%35)));
    var adrBar = Math.min(90, Math.round(toAdr / 280 * 100));
    var revBar = Math.min(90, Math.round(toRev / 4500000 * 100));
    var sdlyH  = Math.max(5, hotel - 9), lyH = Math.max(5, hotel - 6), fcstH = Math.min(100, hotel + 4);
    var sdlyA  = adr - 8, lyA = adr - 4, fcstA = adr + 6;
    var sdlyRn = Math.round(toRn * 0.88), lyRn = Math.round(toRn * 0.93), fcstRn = Math.round(toRn * 1.06);
    var sdlyR  = Math.floor(Math.round(WV_CAP * sdlyH / 100) * sdlyA);
    var lyR    = Math.floor(hnRev * 0.95), fcstR = Math.floor(hnRev * 1.06);
    var revpar = Math.max(50, (adr+80) - 30 - Math.abs((dm*5+dd*3)%20));
    var hRevpar = Math.round(adr * hotel / 100);
    var sdlyRevpar = Math.max(40, revpar - 8), lyRevpar = Math.max(40, revpar - 4);
    var pickup = Math.max(0, Math.floor((v%25+5)*to/Math.max(1,hotel)));
    var hPickup = Math.floor(v%25+5);
    var avgA  = (1.8+v%3*0.1).toFixed(1), avgC = (0.3+v%2*0.1).toFixed(1);
    var hAvgA = (parseFloat(avgA)+0.3).toFixed(1), hAvgC = (parseFloat(avgC)+0.1).toFixed(1);
    var totAT = Math.round(toRn*parseFloat(avgA)),  totCT = Math.round(toRn*parseFloat(avgC));
    var totAH = Math.round(hnRn*parseFloat(hAvgA)), totCH = Math.round(hnRn*parseFloat(hAvgC));
    var totG  = Math.round(toRn*(parseFloat(avgA)+parseFloat(avgC)));
    var hTotG = Math.round(hnRn*(parseFloat(hAvgA)+parseFloat(hAvgC)));
    var avgLos = (2.8+v%5*0.3).toFixed(1)+'n', hLos = (2.8+v%5*0.3+0.4).toFixed(1)+'n';
    var avgLead = (18+v%60)+'d', hLead = (18+v%60+12)+'d';
    var availRooms = Math.max(0, 102-Math.floor(hotel*1.02));
    var availGuar  = Math.floor(8+v%5);
    var aiPct = Math.max(45, Math.min(68, 55+(dm*7+dd*3)%14));
    var bbPct = Math.max(14, Math.min(28, 20+(dm*11+dd*5)%11));
    var hbPct = Math.max(6,  Math.min(16, 10+(dm*5+dd*7)%9));
    var roPct = 100 - aiPct - bbPct - hbPct;
    var toPct = to / Math.max(1, hotel);
    var toMix   = 28+Math.abs((dm*7+dd*5)%25);
    var dirMix  = 30+Math.abs((dm*5+dd*9)%20);
    var otaMix  = 20+Math.abs((dm*9+dd*3)%18);
    var otherMix = Math.max(0, 100-toMix-dirMix-otaMix);
    function fR(val){return val>=1000000?'$'+(val/1000000).toFixed(1)+'M':'$'+Math.round(val/1000)+'k';}
    return {dm, dd, hotel, to, adr, toAdr, toRn, hnRn, toRev, hnRev,
            otherPct, otherRms, freeRms, onlinePct, adrBar, revBar,
            sdlyH, lyH, fcstH, sdlyA, lyA, fcstA, sdlyRn, lyRn, fcstRn,
            sdlyR, lyR, fcstR, revpar, hRevpar, sdlyRevpar, lyRevpar,
            pickup, hPickup, avgA, avgC, hAvgA, hAvgC,
            totAT, totCT, totAH, totCH, totG, hTotG,
            avgLos, hLos, avgLead, hLead, availRooms, availGuar,
            aiPct, bbPct, hbPct, roPct, toPct, toMix, dirMix, otaMix, otherMix, fR, v};
  });

  // ── Row schema (built per group, then assembled in custom order) ──────────
  var _cmpOrder = ['stly','ly','fcst'], _cmpNames = {stly:'STLY',ly:'LY',fcst:'Fcst'};
  var compLabel = _cmpOrder.filter(function(k){ return wvCompare.has(k); }).map(function(k){ return _cmpNames[k]; }).join('/') || 'STLY';
  var grp = { g_closeouts:[], g_daily:[], g_more:[], g_meals:[], g_biz:[], g_avail:[], g_torates:[] };
  window._wbGrpData = grp; // expose for Table Settings modal

  // Group: Close Outs (details directly under group — summary shown in collapsed header)
  if (wvMetricState.dm_closeouts) {
    grp.g_closeouts.push({type:'top', id:'g_closeouts', label:'Close Outs'});
    if (wvMetricState.dm_co_rooms)  grp.g_closeouts.push({type:'sub', id:'co_rooms',  label:'Room Types',     dot:'#6b7280', parent:'g_closeouts'});
    if (wvMetricState.dm_co_boards) grp.g_closeouts.push({type:'sub', id:'co_boards', label:'Board Types',    dot:'#6b7280', parent:'g_closeouts'});
    if (wvMetricState.dm_co_tos)    grp.g_closeouts.push({type:'sub', id:'co_tos',    label:'Tour Operators', dot:'#6b7280', parent:'g_closeouts'});
  }

  // Group: Daily Metrics
  grp.g_daily.push({type:'top', id:'g_daily', label:'Daily Metrics'});
  if (wvMetricState.capacity) {
    grp.g_daily.push({type:'sect', id:'occ',       label:'Occupancy',               parent:'g_daily'});
    grp.g_daily.push({type:'sub',  id:'occ_tdh',   label:'Travel Distribution Hubs',dot:'#004948', parent:'occ'});
    grp.g_daily.push({type:'sub',  id:'occ_other', label:'Other Segments',          dot:'#52d9ce', parent:'occ'});
    grp.g_daily.push({type:'sub',  id:'occ_stly',  label:compLabel,                 dot:'#C4FF45', parent:'occ'});
    grp.g_daily.push({type:'sub',  id:'occ_rem',   label:'Total Hotel Remaining',   dot:'#445e0d', parent:'occ', isRem:true});
  }
  if (wvMetricState.onlineOffline) {
    grp.g_daily.push({type:'sect', id:'onoff',     label:'Online / Offline', parent:'g_daily'});
    grp.g_daily.push({type:'sub',  id:'onoff_on',  label:'Online',  dot:'#004948', parent:'onoff'});
    grp.g_daily.push({type:'sub',  id:'onoff_off', label:'Offline', dot:'#52d9ce', parent:'onoff'});
  }
  if (wvMetricState.adr) {
    grp.g_daily.push({type:'sect', id:'adr',       label:'ADR',          parent:'g_daily'});
    grp.g_daily.push({type:'sub',  id:'adr_t',     label:'TO ADR',        dot:'#004948', parent:'adr'});
    grp.g_daily.push({type:'sub',  id:'adr_hotel', label:'Hotel ADR',    dot:'#52d9ce', parent:'adr'});
    grp.g_daily.push({type:'sub',  id:'adr_stly',  label:compLabel,      dot:'#C4FF45', parent:'adr'});
  }
  if (wvMetricState.revenue) {
    grp.g_daily.push({type:'sect', id:'rev',       label:'Revenue',       parent:'g_daily'});
    grp.g_daily.push({type:'sub',  id:'rev_t',     label:'TO Revenue',     dot:'#004948', parent:'rev'});
    grp.g_daily.push({type:'sub',  id:'rev_hotel', label:'Hotel Revenue', dot:'#52d9ce', parent:'rev'});
    grp.g_daily.push({type:'sub',  id:'rev_stly',  label:compLabel,       dot:'#C4FF45', parent:'rev'});
  }

  // Group: More Metrics
  var hasMore = wvMetricState.dm_rnSold || wvMetricState.dm_trevpar || wvMetricState.dm_pickup ||
                wvMetricState.dm_avgAdults || wvMetricState.dm_avgChildren ||
                wvMetricState.dm_totalAdults || wvMetricState.dm_totalChildren ||
                wvMetricState.dm_totalGuests || wvMetricState.dm_avgLos ||
                wvMetricState.dm_avgLeadTime || wvMetricState.dm_availRooms || wvMetricState.dm_availGuar;
  if (hasMore) {
    grp.g_more.push({type:'top', id:'g_more', label:'More Metrics'});
    if (wvMetricState.dm_rnSold) {
      grp.g_more.push({type:'sect', id:'rn',       label:'RN Sold',    parent:'g_more'});
      grp.g_more.push({type:'sub',  id:'rn_t',     label:'TO RN',       dot:'#004948', parent:'rn'});
      grp.g_more.push({type:'sub',  id:'rn_hotel', label:'Hotel RN',   dot:'#52d9ce', parent:'rn'});
      grp.g_more.push({type:'sub',  id:'rn_stly',  label:compLabel,    dot:'#C4FF45', parent:'rn'});
    }
    if (wvMetricState.dm_trevpar) {
      grp.g_more.push({type:'sect', id:'revpar_s',    label:'REVPAR',    parent:'g_more'});
      grp.g_more.push({type:'sub',  id:'revpar_h',    label:'Hotel',     dot:'#52d9ce', parent:'revpar_s'});
      grp.g_more.push({type:'sub',  id:'revpar_stly', label:compLabel,   dot:'#C4FF45', parent:'revpar_s'});
    }
    if (wvMetricState.dm_pickup) {
      var _anyPU = pickupDayValues.some(function(dv, i) { return wvMetricState['dm_pickup_' + i] !== false; });
      if (_anyPU) grp.g_more.push({type:'sect', id:'pickup_s', label:'Pickup', parent:'g_more'});
    }
    if (wvMetricState.dm_avgAdults) {
      grp.g_more.push({type:'sect', id:'avga_s', label:'Avg Adults',   parent:'g_more'});
      grp.g_more.push({type:'sub',  id:'avga_t', label:'T Avg Adults', dot:'#004948', parent:'avga_s'});
      grp.g_more.push({type:'sub',  id:'avga_h', label:'Hotel',        dot:'#52d9ce', parent:'avga_s'});
    }
    if (wvMetricState.dm_avgChildren) {
      grp.g_more.push({type:'sect', id:'avgc_s', label:'Avg Children',   parent:'g_more'});
      grp.g_more.push({type:'sub',  id:'avgc_t', label:'T Avg Children', dot:'#004948', parent:'avgc_s'});
      grp.g_more.push({type:'sub',  id:'avgc_h', label:'Hotel',          dot:'#52d9ce', parent:'avgc_s'});
    }
    if (wvMetricState.dm_totalAdults) {
      grp.g_more.push({type:'sect', id:'tota_s', label:'Total Adults',   parent:'g_more'});
      grp.g_more.push({type:'sub',  id:'tota_t', label:'T Total Adults', dot:'#004948', parent:'tota_s'});
      grp.g_more.push({type:'sub',  id:'tota_h', label:'Hotel',          dot:'#52d9ce', parent:'tota_s'});
    }
    if (wvMetricState.dm_totalChildren) {
      grp.g_more.push({type:'sect', id:'totc_s', label:'Total Children',   parent:'g_more'});
      grp.g_more.push({type:'sub',  id:'totc_t', label:'T Total Children', dot:'#004948', parent:'totc_s'});
      grp.g_more.push({type:'sub',  id:'totc_h', label:'Hotel',            dot:'#52d9ce', parent:'totc_s'});
    }
    if (wvMetricState.dm_totalGuests) {
      grp.g_more.push({type:'sect', id:'totg_s', label:'Total Guests', parent:'g_more'});
      grp.g_more.push({type:'sub',  id:'totg_t', label:'T Guests',     dot:'#004948', parent:'totg_s'});
      grp.g_more.push({type:'sub',  id:'totg_h', label:'Hotel',        dot:'#52d9ce', parent:'totg_s'});
    }
    if (wvMetricState.dm_avgLos) {
      grp.g_more.push({type:'sect', id:'los_s', label:'Avg LOS',   parent:'g_more'});
      grp.g_more.push({type:'sub',  id:'los_t', label:'T Avg LOS', dot:'#004948', parent:'los_s'});
      grp.g_more.push({type:'sub',  id:'los_h', label:'Hotel',     dot:'#52d9ce', parent:'los_s'});
    }
    if (wvMetricState.dm_avgLeadTime) {
      grp.g_more.push({type:'sect', id:'lead_s', label:'Lead Time',   parent:'g_more'});
      grp.g_more.push({type:'sub',  id:'lead_t', label:'T Lead Time', dot:'#004948', parent:'lead_s'});
      grp.g_more.push({type:'sub',  id:'lead_h', label:'Hotel',       dot:'#52d9ce', parent:'lead_s'});
    }
    if (wvMetricState.dm_availRooms) grp.g_more.push({type:'sect', id:'avail_s',  label:'Avail Rooms', parent:'g_more'});
    if (wvMetricState.dm_availGuar)  grp.g_more.push({type:'sect', id:'availg_s', label:'Avail Guar.', parent:'g_more'});
  }

  // Group: Meal Plans
  if (wvMetricState.mealsSummary) {
    grp.g_meals.push({type:'top',  id:'g_meals', label:'Meal Plans'});
    grp.g_meals.push({type:'sect', id:'mp_ai',   label:'All Inclusive',   parent:'g_meals'});
    grp.g_meals.push({type:'sub',  id:'mp_ai_t', label:'TO',                        dot:'#004948', parent:'mp_ai'});
    grp.g_meals.push({type:'sub',  id:'mp_ai_h', label:'Hotel',             dot:'#52d9ce', parent:'mp_ai'});
    grp.g_meals.push({type:'sect', id:'mp_bb',   label:'Bed & Breakfast', parent:'g_meals'});
    grp.g_meals.push({type:'sub',  id:'mp_bb_t', label:'TO',                        dot:'#004948', parent:'mp_bb'});
    grp.g_meals.push({type:'sub',  id:'mp_bb_h', label:'Hotel',             dot:'#52d9ce', parent:'mp_bb'});
    grp.g_meals.push({type:'sect', id:'mp_hb',   label:'Half Board',      parent:'g_meals'});
    grp.g_meals.push({type:'sub',  id:'mp_hb_t', label:'TO',                        dot:'#004948', parent:'mp_hb'});
    grp.g_meals.push({type:'sub',  id:'mp_hb_h', label:'Hotel',             dot:'#52d9ce', parent:'mp_hb'});
    grp.g_meals.push({type:'sect', id:'mp_ro',   label:'Room Only',       parent:'g_meals'});
    grp.g_meals.push({type:'sub',  id:'mp_ro_t', label:'TO',                        dot:'#004948', parent:'mp_ro'});
    grp.g_meals.push({type:'sub',  id:'mp_ro_h', label:'Hotel',             dot:'#52d9ce', parent:'mp_ro'});
    grp.g_meals.push({type:'sect', id:'mp_sum',  label:'Summary',         parent:'g_meals'});
  }

  // Group: Business Mix
  if (wvMetricState.bizMix) {
    grp.g_biz.push({type:'top',  id:'g_biz',    label:'Business Mix'});
    grp.g_biz.push({type:'sect', id:'biz',       label:'Channel Mix', parent:'g_biz'});
    grp.g_biz.push({type:'sub',  id:'biz_to',    label:'TO',          dot:'#004948', parent:'biz'});
    grp.g_biz.push({type:'sub',  id:'biz_dir',   label:'Direct',      dot:'#52d9ce', parent:'biz'});
    grp.g_biz.push({type:'sub',  id:'biz_ota',   label:'OTA',         dot:'#D97706', parent:'biz'});
    grp.g_biz.push({type:'sub',  id:'biz_other', label:'Other',       dot:'#9ca3af', parent:'biz'});
  }

  // Group: Room Availability
  if (wvMetricState.avail || wvMetricState.availAlloc) {
    grp.g_avail.push({type:'top', id:'g_avail', label:'Room Availability'});
    RT_NAMES.forEach(function(name, i) {
      grp.g_avail.push({type:'sect', id:'avrt'+i,       label:name,        parent:'g_avail', rtIdx:i});
      grp.g_avail.push({type:'sub',  id:'avrt'+i+'_to', label:'TO Sold',   dot:'#004948', parent:'avrt'+i, rtIdx:i, rtSub:'to'});
      grp.g_avail.push({type:'sub',  id:'avrt'+i+'_ot', label:'Other Segments', dot:'#52d9ce', parent:'avrt'+i, rtIdx:i, rtSub:'other'});
      grp.g_avail.push({type:'sub',  id:'avrt'+i+'_tn', label:'Tentative Sold (Group)',    dot:'#967EF3', parent:'avrt'+i, rtIdx:i, rtSub:'tentative'});
      grp.g_avail.push({type:'sub',  id:'avrt'+i+'_oo', label:'Out-of-Order',             dot:'#ef4444', parent:'avrt'+i, rtIdx:i, rtSub:'ooo'});
      grp.g_avail.push({type:'sub',  id:'avrt'+i+'_al', label:'Alloc Rem.',dot:'#D97706', parent:'avrt'+i, rtIdx:i, rtSub:'alloc'});
      grp.g_avail.push({type:'sub',  id:'avrt'+i+'_av', label:'Total Hotel Remaining', dot:'#445e0d', parent:'avrt'+i, rtIdx:i, rtSub:'avail', isRem:true});
    });
  }

  // Group: Travel Co. Rates
  if (wvMetricState.toRates) {
    grp.g_torates.push({type:'top', id:'g_torates', label:'Travel Co. Rates'});
    TO_NAMES.forEach(function(name, i) {
      grp.g_torates.push({type:'sect', id:'torate'+i, label:name, parent:'g_torates', toIdx:i});
    });
    grp.g_torates.push({type:'sect', id:'torate_base', label:'Base Rate', parent:'g_torates', toBase:true});
  }

  // ── Assemble rows in custom order ─────────────────────────────────────────
  var wbOrder = (_wbGroupOrder && _wbGroupOrder.length) ? _wbGroupOrder : WB_GROUPS_DEF.map(function(g){return g.key;});
  var rows = [];
  wbOrder.forEach(function(key) { if (grp[key]) rows = rows.concat(grp[key]); });

  // ── Helpers ────────────────────────────────────────────────────────────────
  var chevUp   = '<span class="material-icons" style="font-size:14px">expand_less</span>';
  var chevDown = '<span class="material-icons" style="font-size:14px">expand_more</span>';

  // Populate module-level ID list for Open All / Close All
  _wbAllIds = rows.filter(function(r){ return r.type==='top'||r.type==='sect'; }).map(function(r){ return r.id; });

  // Default collapse state: top-groups open, sub-sections closed (only set if not yet toggled by user)
  rows.forEach(function(r) {
    if (r.type === 'sect' && !Object.prototype.hasOwnProperty.call(_wbCollapsed, r.id)) {
      _wbCollapsed[r.id] = true;
    }
  });

  // Flat id→row map for fast grandparent lookup
  var rowMap = {};
  rows.forEach(function(r){ rowMap[r.id] = r; });

  function isHidden(row) {
    if (!row.parent) return false;
    if (_wbCollapsed[row.parent]) return true;
    var par = rowMap[row.parent];
    if (par && par.parent && _wbCollapsed[par.parent]) return true;
    return false;
  }

  var TODAY_WV = new Date(2026, 2, 9);

  function cmpSfx(cmpStr, curr, comp) {
    if (!cmpStr || wvCompare.size === 0) return '';
    var clr = '#9ca3af';
    if (curr != null && comp != null && !isNaN(parseFloat(curr)) && !isNaN(parseFloat(comp))) {
      var c = parseFloat(curr), p = parseFloat(comp);
      if (c > p) clr = '#16a34a';
      else if (c < p) clr = '#dc2626';
    }
    return '<span class="wv-cmp-sep"> / </span><span class="wv-cmp-val-txt" style="color:' + clr + '">' + cmpStr + '</span>';
  }

  // trendBadge wraps badges in a single container so wb-sect-val always has
  // exactly 2 flex children (value + badges), keeping space-between sane with multiple compares
  function trendBadge(curr, stlyComp, lyComp, fcstComp) {
    var b = _wvMultiTrendBadge(curr, stlyComp, lyComp, fcstComp);
    return b ? '<span style="display:inline-flex;gap:2px;flex-wrap:wrap;flex-shrink:0;align-items:center">' + b + '</span>' : '';
  }

  function wbGrad(clr) {
    if (clr==='#004948') return 'linear-gradient(to right,#004948,#007a75)';
    if (clr==='#52d9ce') return 'linear-gradient(to right,#52d9ce,#8aeee8)';
    if (clr==='#445e0d') return 'linear-gradient(to right,#445e0d,#6a9014)';
    if (clr==='#D97706') return 'linear-gradient(to right,#D97706,#F59E0B)';
    if (clr==='#967EF3') return 'linear-gradient(to right,#967EF3,#a78bfa)';
    if (clr==='#16a34a') return 'linear-gradient(to right,#16a34a,#22c55e)';
    if (clr==='#C4FF45') return 'linear-gradient(to right,#C4FF45,#D4FF73)';
    return clr;
  }
  function wbStackBar(segs) {
    return '<div style="height:6px;background:'+wbGrad('#e5e7eb')+';border-radius:2px;display:flex;overflow:hidden;margin-top:3px">'
      + segs.map(function(s){ return '<div style="width:'+s.p+'%;background:'+wbGrad(s.c)+'"></div>'; }).join('')
      + '</div>';
  }

  // Wrap a bar HTML string with a comparison marker line
  function wbBarMark(barHtml, compPct) {
    if (wvCompare.size === 0 || compPct == null || isNaN(compPct)) return barHtml;
    var pct = Math.min(100, Math.max(0, compPct));
    return '<div style="position:relative">'
      + barHtml
      + '<div style="position:absolute;left:'+pct+'%;top:0;height:6px;width:2.5px;background:'+wbGrad('#C4FF45')+';transform:translateX(-50%);z-index:2;border-radius:1px;pointer-events:none;box-shadow:0 0 3px rgba(196,255,69,0.6)"></div>'
      + '</div>';
  }

  // ── Build HTML ─────────────────────────────────────────────────────────────
  var html = '<div class="wb-layout">';

  // Sticky date-header row
  html += '<div class="wb-row wb-hdr-row">';
  html += '<div class="wb-label-cell wb-hdr-label-cell"></div>';
  days.forEach(function(dv) {
    var dt   = new Date(dv.year, dv.month - 1, dv.day);
    var dow  = DOW_SHORT[dt.getDay()];
    var isAct = dv.day === activeDay && dv.month === month;
    var dba  = Math.round((dt - TODAY_WV) / 86400000);
    var dbaStr = dba === 0 ? 'Today' : dba > 0 ? dba + ' DBA' : '';
    var mm2 = String(dv.month).padStart(2,'0'), dd2 = String(dv.day).padStart(2,'0');
    var isoDate = '2026-' + mm2 + '-' + dd2;
    var isSel = _wbSelectedDays.has(isoDate);
    var _coKey2 = dv.month+'-'+dv.day;
    var _coFull2 = LOCKED_DAYS.has(_coKey2);
    var _coPart2 = (PARTIAL_CLOSURES[_coKey2] || []).length > 0;
    var _lockColor = isSel ? '#f43f5e' : _coFull2 ? '#fca5a5' : _coPart2 ? '#fde68a' : 'rgba(255,255,255,0.35)';
    var _lockWidth = (isSel || _coFull2 || _coPart2) ? '1.5' : '1.3';
    var _lockIcon = _coFull2
          ? '<span class="wb-hdr-lock-icon" title="Closed out"><span class="material-icons" style="font-size:13px;color:#fca5a5">lock</span></span>'
          : _coPart2
          ? '<span class="wb-hdr-lock-icon" title="Partially closed out"><span class="material-icons" style="font-size:13px;color:#fde68a">lock_open</span></span>'
          : '';
    var _wbEvtKey = dv.month + '-' + dv.day;
    var _wbHasEvt = (typeof CAL_EVENTS !== 'undefined' && CAL_EVENTS[_wbEvtKey]);
    var _evtIcon = _wbHasEvt
          ? '<span class="wv-event-cal-icon has-events" data-event-key="' + _wbEvtKey + '" onmouseenter="calShowEventTip(event,\'' + _wbEvtKey + '\')" onmouseleave="calHideEventTip()" style="display:inline-flex;align-items:center"><span class="material-icons" style="font-size:14px;color:#c4ff45">today</span></span>'
          : '';
    html += '<div class="wb-data-cell wb-hdr-cell'
          + (isAct ? ' wb-hdr-active' : '')
          + (isSel ? ' wb-hdr-selected' : '')
          + '" data-wb-date="' + isoDate + '" title="Select for close-out">'
          + '<input type="checkbox" class="wv-day-chk wb-day-chk"' + (isSel ? ' checked' : '') + ' onclick="event.stopPropagation();wbDayToggle(\'' + isoDate + '\');this.checked=_wbSelectedDays.has(\'' + isoDate + '\')" title="Select for close-out">'
          + '<span class="wb-hdr-dow">' + dow + '</span>'
          + '<span class="wb-hdr-date">' + dv.day + '/' + dv.month + '</span>'
          + _evtIcon
          + (dbaStr ? '<span style="font-size:10px;background:rgba(255,255,255,0.2);border-radius:3px;padding:0 4px;color:#fff;white-space:nowrap">'+dbaStr+'</span>' : '')
          + _lockIcon
          + '</div>';
  });
  html += '</div>';

  // Data rows
  rows.forEach(function(row) {
    var collapsed = !!_wbCollapsed[row.id];
    var hidden    = isHidden(row);
    var rowCls    = 'wb-row wb-row-' + row.type + (hidden ? ' wb-row-hidden' : '');

    html += '<div class="' + rowCls + '" data-wb-id="' + row.id + '"'
          + (row.parent ? ' data-wb-parent="' + row.parent + '"' : '') + '>';

    // ── Label cell ──────────────────────────────────────────────────────────
    if (row.type === 'top') {
      html += '<div class="wb-label-cell wb-grp-hdr" onclick="wbToggle(\'' + row.id + '\')">'
            + '<span class="wb-chev">' + (collapsed ? chevDown : chevUp) + '</span>'
            + '<span class="wb-grp-label">' + row.label + '</span>'
            + '</div>';
    } else if (row.type === 'sect') {
      if (row.id === 'pickup_s') {
        html += '<div class="wb-label-cell wb-sect-lbl">'
              + '<span class="wb-sect-label">' + row.label + '</span>'
              + '</div>';
      } else {
        html += '<div class="wb-label-cell wb-sect-lbl" onclick="wbToggle(\'' + row.id + '\')">'
              + '<span class="wb-chev">' + (collapsed ? chevDown : chevUp) + '</span>'
              + '<span class="wb-sect-label">' + row.label + '</span>'
              + '</div>';
      }
    } else {
      var dotHtml = row.dot ? '<span class="wb-sub-dot" style="background:' + row.dot + '"></span>' : '';
      html += '<div class="wb-label-cell wb-sub-lbl-cell">'
            + dotHtml
            + '<span class="wb-sub-label' + (row.isRem ? ' wb-sub-lbl-rem' : '') + '">' + row.label + '</span>'
            + '</div>';
    }

    // ── Data cells (one per day) ────────────────────────────────────────────
    days.forEach(function(dv, i) {
      var d = dd7[i];
      var cellContent = '';

      if (row.type === 'top') {
        // Close Outs group shows summary in collapsed state
        if (row.id === 'g_closeouts' && collapsed) {
          var _coKey3 = d.dm+'-'+d.dd;
          var _coFull3 = LOCKED_DAYS.has(_coKey3);
          var _coPart3 = PARTIAL_CLOSURES[_coKey3] || [];
          var _lockIco3 = '<span class="material-icons" style="font-size:13px;vertical-align:middle;margin-right:3px">';
          if (_coFull3) {
            cellContent = '<div style="display:flex;align-items:center;gap:4px;padding:2px 0">'+_lockIco3+'lock</span><span style="font-size:12px;font-weight:600;color:var(--text-primary)">Full Close Out</span></div>';
          } else if (_coPart3.length > 0) {
            cellContent = '<div style="display:flex;align-items:center;gap:4px;padding:2px 0">'+_lockIco3+'lock</span><span style="font-size:12px;font-weight:600;color:var(--text-primary)">'+_coPart3.length+' rule'+(_coPart3.length>1?'s':'')+'</span></div>';
          } else {
            cellContent = '<div style="display:flex;align-items:center;gap:4px;padding:2px 0"><span class="material-icons" style="font-size:13px;color:#059669;vertical-align:middle;margin-right:3px">check_circle</span><span style="font-size:12px;color:var(--text-primary)">Open</span></div>';
          }
        } else {
          cellContent = '';
        }

      } else if (row.type === 'sect') {
        var cs = '';

        // ── Room Availability (dynamic rtIdx) ────────────────────────────────
        if (row.rtIdx !== undefined) {
          var inv  = RT_CAPS[row.rtIdx];
          var sold = Math.min(inv, Math.floor(inv * d.hotel / 110));
          var toS  = Math.min(sold, Math.round(sold * d.to / Math.max(1, d.hotel)));
          var otS  = sold - toS;
          var tent = Math.max(0, Math.floor(2+Math.abs((d.dm*(row.rtIdx+4)+d.dd*(row.rtIdx+2))%6)));
          var alloc = Math.floor(inv * 0.8 + Math.abs((d.dm*(row.rtIdx+3)+d.dd*(row.rtIdx+5))%15));
          var allocRem = Math.max(0, alloc - toS);
          var avRm = Math.max(0, inv - sold - tent);
          var toP  = Math.round(toS/inv*100), otP = Math.round(otS/inv*100);
          var tnP  = Math.round(tent/inv*100);
          var alP  = Math.round(allocRem/inv*100), avP = Math.max(0, 100-toP-otP-tnP-alP);
          var avClr = avRm <= 0 ? '#dc2626' : '#004948';
          cellContent = '<div class="wb-sect-val"><span class="wv-occ-total" style="color:'+(avRm<=0?'#16a34a':avClr)+'">'
            + (avRm <= 0 ? '0 available' : avRm+' avail') + '</span>'
            + '<span style="font-size:12px;color:#9ca3af;margin-left:4px">/ '+inv+'</span></div>'
            + '<div class="wv-occ-bar-track">'
            + '<div style="width:'+toP+'%;background:'+wbGrad('#004948')+';height:6px"></div>'
            + '<div style="width:'+otP+'%;background:'+wbGrad('#52d9ce')+';height:6px"></div>'
            + '<div style="width:'+tnP+'%;background:'+wbGrad('#967EF3')+';height:6px"></div>'
            + '<div style="width:'+alP+'%;background:'+wbGrad('#D97706')+';height:6px"></div>'
            + '<div style="width:'+avP+'%;background:'+wbGrad('#d7f7ed')+';height:6px"></div>'
            + '</div>';

        // ── Travel Co. Rates (dynamic toIdx) ──────────────────────────────────
        } else if (row.toIdx !== undefined) {
          var toRate  = d.adr - 15 + Math.abs((d.dm*(row.toIdx+3)+d.dd*(row.toIdx+5))%50);
          var toAllot = 5 + Math.abs((d.dm*(row.toIdx+2)+d.dd*(row.toIdx+3))%20);
          var toUsed  = Math.max(0, toAllot - Math.floor(d.hotel/20));
          var toRem   = toAllot - toUsed;
          var barPct  = Math.round(toUsed/toAllot*100);
          var isEbb   = (new Date(2026,d.dm-1,d.dd)).getDay() < 3;
          var promoTxt = isEbb ? 'EBB 10%' : 'Contract';
          var promoClr = isEbb ? '#16a34a' : '#2563eb';
          cellContent = '<div class="wb-sect-val" style="justify-content:space-between">'
            + '<span class="wv-occ-total">$'+toRate+'</span>'
            + '<span style="font-size:12px;color:#9ca3af">'+toRem+'r</span>'
            + '<span style="font-size:11px;font-weight:700;padding:1px 5px;border-radius:3px;background:'+promoClr+'22;color:'+promoClr+';border:1px solid '+promoClr+'44">'+promoTxt+'</span>'
            + '</div>'
            + '<div class="wv-occ-bar-track"><div style="width:'+barPct+'%;background:'+wbGrad('#004948')+';height:6px"></div></div>';

        } else if (row.toBase) {
          var baseRate = d.adr + 8;
          cellContent = '<div class="wb-sect-val"><span class="wv-occ-total" style="font-weight:700">$'+baseRate+'</span></div>'
            + '<div class="wv-occ-bar-track"><div style="width:'+Math.min(90,Math.round(baseRate/280*100))+'%;background:'+wbGrad('#004948')+';height:6px"></div></div>';

        } else {
        // colors are read from the first sub-row's dot for each section
        function wbBar(pct, clr) {
          return '<div class="wv-occ-bar-track"><div style="width:'+pct+'%;background:'+wbGrad(clr)+';height:6px"></div></div>';
        }
        switch (row.id) {
          // ── Daily Metrics ──────────────────────────────────────────────────
          case 'occ': {
            cs = _wvMultiCmpSfx(d.hotel, d.sdlyH, d.lyH, d.fcstH, function(v){ return v+'%'; });
            var _cv0 = wvCompare.has('stly')?d.sdlyH:wvCompare.has('ly')?d.lyH:wvCompare.has('fcst')?d.fcstH:null;
            cellContent = '<div class="wb-sect-val"><span class="wv-occ-total">'+d.hotel+'%</span>'+trendBadge(d.hotel,d.sdlyH,d.lyH,d.fcstH)+'</div>'
              + (cs ? '<div class="wb-cmp-line">'+cs+'</div>' : '')
              + wbBarMark('<div class="wv-occ-bar-track">'
                + '<div style="width:'+d.to+'%;background:'+wbGrad('#004948')+';height:6px"></div>'
                + '<div style="width:'+d.otherPct+'%;background:'+wbGrad('#52d9ce')+';height:6px"></div>'
                + '</div>', _cv0);
            break;
          }
          case 'onoff':
            cellContent = '<div class="wb-sect-val"><span class="wv-occ-total">'+d.onlinePct+'%</span></div>'
              + '<div class="wv-occ-bar-track">'
              + '<div style="width:'+d.onlinePct+'%;background:'+wbGrad('#004948')+';height:6px"></div>'
              + '<div style="width:'+(100-d.onlinePct)+'%;background:'+wbGrad('#52d9ce')+';height:6px"></div>'
              + '</div>';
            break;
          case 'adr': {
            cs = _wvMultiCmpSfx(d.toAdr, d.sdlyA, d.lyA, d.fcstA, function(v){ return '$'+v; });
            var _cv0 = wvCompare.has('stly')?d.sdlyA:wvCompare.has('ly')?d.lyA:wvCompare.has('fcst')?d.fcstA:null;
            var cvPct = _cv0!=null?Math.min(90,Math.round(_cv0/280*100)):null;
            cellContent = '<div class="wb-sect-val"><span class="wv-occ-total">$'+d.toAdr+'</span>'+trendBadge(d.toAdr,d.sdlyA,d.lyA,d.fcstA)+'</div>'
              + (cs ? '<div class="wb-cmp-line">'+cs+'</div>' : '')
              + wbBarMark(wbBar(d.adrBar, '#004948'), cvPct);
            break;
          }
          case 'rev': {
            cs = _wvMultiCmpSfx(d.toRev, d.sdlyR, d.lyR, d.fcstR, d.fR);
            var _cv0 = wvCompare.has('stly')?d.sdlyR:wvCompare.has('ly')?d.lyR:wvCompare.has('fcst')?d.fcstR:null;
            var cvPct = _cv0!=null?Math.min(90,Math.round(_cv0/4500000*100)):null;
            cellContent = '<div class="wb-sect-val"><span class="wv-occ-total">'+d.fR(d.toRev)+'</span>'+trendBadge(d.toRev,d.sdlyR,d.lyR,d.fcstR)+'</div>'
              + (cs ? '<div class="wb-cmp-line">'+cs+'</div>' : '')
              + wbBarMark(wbBar(d.revBar, '#004948'), cvPct);
            break;
          }
          // ── More Metrics ───────────────────────────────────────────────────
          case 'rn': {
            cs = _wvMultiCmpSfx(d.toRn, d.sdlyRn, d.lyRn, d.fcstRn, String);
            var _cv0 = wvCompare.has('stly')?d.sdlyRn:wvCompare.has('ly')?d.lyRn:wvCompare.has('fcst')?d.fcstRn:null;
            var cvPct = _cv0!=null?Math.round(_cv0/WV_CAP*100):null;
            cellContent = '<div class="wb-sect-val"><span class="wv-occ-total">'+d.toRn+'</span>'+trendBadge(d.toRn,d.sdlyRn,d.lyRn,d.fcstRn)+'</div>'
              + (cs ? '<div class="wb-cmp-line">'+cs+'</div>' : '')
              + wbBarMark(wbBar(Math.round(d.toRn/WV_CAP*100), '#004948'), cvPct) + wbBar(Math.round(d.hnRn/WV_CAP*100), '#52d9ce');
            break;
          }
          case 'revpar_s': {
            var _cv0 = wvCompare.has('stly')?d.sdlyRevpar:wvCompare.has('ly')?d.lyRevpar:null;
            var cvPct = _cv0!=null?Math.min(90,Math.round(_cv0/4)):null;
            var _rcs = _wvMultiCmpSfx(d.hRevpar,d.sdlyRevpar,d.lyRevpar,null,function(v){return '$'+v;});
            cellContent = '<div class="wb-sect-val"><span class="wv-occ-total">$'+d.hRevpar+'</span>'+trendBadge(d.hRevpar,d.sdlyRevpar,d.lyRevpar,null)+'</div>'
              + (_rcs ? '<div class="wb-cmp-line">'+_rcs+'</div>' : '')
              + wbBarMark(wbBar(Math.min(90,Math.round(d.hRevpar/4)), '#004948'), cvPct);
            break;
          }
          case 'pickup_s': {
            var _activePUw = pickupDayValues.filter(function(dv, i) { return wvMetricState['dm_pickup_' + i] !== false; });
            var _pun = _activePUw.length;
            if (!_pun) { cellContent = ''; break; }
            var _puHdrs = '', _puVals = '';
            _activePUw.forEach(function(dv, idx) {
              var _sc = dv<=1?0.3:dv<=3?0.6:dv<=7?1:Math.min(2,dv/7);
              var _pv = Math.max(0, Math.round(d.pickup * _sc));
              var _pvBar = Math.min(90, _pv * 3);
              var _hpvBar = Math.min(90, d.hPickup * _sc * 3);
              var _bL = idx===0 ? '' : 'border-left:1px solid #e0e0e0;';
              _puHdrs += '<div class="wv-pu-fig-cell" style="'+_bL+'">'+dv+'</div>';
              _puVals += '<div class="wv-pu-fig-val-cell" style="'+_bL+'">'
                + '<span class="wv-pu-fig-num">+'+_pv+'</span>'
                + '<div class="wv-occ-bar-track" style="height:4px;margin:2px 0 1px"><div style="width:'+_pvBar+'%;background:'+wbGrad('#004948')+';height:4px"></div></div>'
                + '<div class="wv-occ-bar-track" style="height:4px"><div style="width:'+_hpvBar+'%;background:'+wbGrad('#52d9ce')+';height:4px"></div></div>'
                + '</div>';
            });
            cellContent = '<div class="wv-pu-fig-wrap">'
              + '<div class="wv-pu-fig-hdr-row" style="grid-template-columns:repeat('+_pun+',1fr)">'+_puHdrs+'</div>'
              + '<div class="wv-pu-fig-val-row" style="grid-template-columns:repeat('+_pun+',1fr)">'+_puVals+'</div>'
              + '</div>';
            break;
          }
          case 'avga_s':
            cellContent = '<div class="wb-sect-val"><span class="wv-occ-total">'+d.avgA+'</span></div>'
              + wbBar(Math.min(90,parseFloat(d.avgA)/3*100), '#004948') + wbBar(Math.min(90,parseFloat(d.hAvgA)/3*100), '#52d9ce');
            break;
          case 'avgc_s':
            cellContent = '<div class="wb-sect-val"><span class="wv-occ-total">'+d.avgC+'</span></div>'
              + wbBar(Math.min(90,parseFloat(d.avgC)/2*100), '#004948') + wbBar(Math.min(90,parseFloat(d.hAvgC)/2*100), '#52d9ce');
            break;
          case 'tota_s':
            cellContent = '<div class="wb-sect-val"><span class="wv-occ-total">'+d.totAT+'</span></div>'
              + wbBar(Math.min(90,Math.round(d.totAT/500*100)), '#004948') + wbBar(Math.min(90,Math.round(d.totAH/500*100)), '#52d9ce');
            break;
          case 'totc_s':
            cellContent = '<div class="wb-sect-val"><span class="wv-occ-total">'+d.totCT+'</span></div>'
              + wbBar(Math.min(90,Math.round(d.totCT/100*100)), '#004948') + wbBar(Math.min(90,Math.round(d.totCH/100*100)), '#52d9ce');
            break;
          case 'totg_s':
            cellContent = '<div class="wb-sect-val"><span class="wv-occ-total">'+d.totG+'</span></div>'
              + wbBar(Math.min(90,Math.round(d.totG/600*100)), '#004948') + wbBar(Math.min(90,Math.round(d.hTotG/600*100)), '#52d9ce');
            break;
          case 'los_s':
            cellContent = '<div class="wb-sect-val"><span class="wv-occ-total">'+d.avgLos+'</span></div>'
              + wbBar(Math.min(90,parseFloat(d.avgLos)/10*100), '#004948') + wbBar(Math.min(90,parseFloat(d.hLos)/10*100), '#52d9ce');
            break;
          case 'lead_s':
            cellContent = '<div class="wb-sect-val"><span class="wv-occ-total">'+d.avgLead+'</span></div>'
              + wbBar(Math.min(90,parseInt(d.avgLead)/90*100), '#004948') + wbBar(Math.min(90,parseInt(d.hLead)/90*100), '#52d9ce');
            break;
          case 'avail_s':
            cellContent = '<div class="wb-sect-val"><span class="wv-occ-total">'+d.availRooms+' rm</span></div>'
              + wbBar(Math.min(90,Math.round(d.availRooms/WV_CAP*100)), '#16a34a');
            break;
          case 'availg_s':
            cellContent = '<div class="wb-sect-val"><span class="wv-occ-total">'+d.availGuar+' rm</span></div>'
              + wbBar(Math.min(90,Math.round(d.availGuar/20*100)), '#004948');
            break;
          // ── Meal Plans — each uses its own dot color ────────────────────────
          case 'mp_ai':
            { var aiRn=Math.round(d.hnRn*d.aiPct/100),aiSeats=Math.round(aiRn*(parseFloat(d.hAvgA)+parseFloat(d.hAvgC)));
            cellContent = '<div class="wb-sect-val"><span class="wv-occ-total">'+d.aiPct+'%</span><span style="font-size:11px;color:#6b7280;margin-left:4px">'+aiRn+' rms</span></div>'
              +'<div style="font-size:11px;color:#374151;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;margin-top:-2px">'+aiSeats+' seats</div>'
              + wbBar(d.aiPct, '#004948'); }
            break;
          case 'mp_bb':
            { var bbRn=Math.round(d.hnRn*d.bbPct/100),bbSeats=Math.round(bbRn*(parseFloat(d.hAvgA)+parseFloat(d.hAvgC)));
            cellContent = '<div class="wb-sect-val"><span class="wv-occ-total">'+d.bbPct+'%</span><span style="font-size:11px;color:#6b7280;margin-left:4px">'+bbRn+' rms</span></div>'
              +'<div style="font-size:11px;color:#374151;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;margin-top:-2px">'+bbSeats+' seats</div>'
              + wbBar(d.bbPct, '#004948'); }
            break;
          case 'mp_hb':
            { var hbRn=Math.round(d.hnRn*d.hbPct/100),hbSeats=Math.round(hbRn*(parseFloat(d.hAvgA)+parseFloat(d.hAvgC)));
            cellContent = '<div class="wb-sect-val"><span class="wv-occ-total">'+d.hbPct+'%</span><span style="font-size:11px;color:#6b7280;margin-left:4px">'+hbRn+' rms</span></div>'
              +'<div style="font-size:11px;color:#374151;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;margin-top:-2px">'+hbSeats+' seats</div>'
              + wbBar(d.hbPct, '#004948'); }
            break;
          case 'mp_ro':
            { var roRn=Math.round(d.hnRn*d.roPct/100),roSeats=Math.round(roRn*(parseFloat(d.hAvgA)+parseFloat(d.hAvgC)));
            cellContent = '<div class="wb-sect-val"><span class="wv-occ-total">'+d.roPct+'%</span><span style="font-size:11px;color:#6b7280;margin-left:4px">'+roRn+' rms</span></div>'
              +'<div style="font-size:11px;color:#374151;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;margin-top:-2px">'+roSeats+' seats</div>'
              + wbBar(d.roPct, '#004948'); }
            break;
          case 'mp_sum':
            { var _sumGPR=parseFloat(d.hAvgA)+parseFloat(d.hAvgC);
              var _aiR=Math.round(d.hnRn*d.aiPct/100),_aiSt=Math.round(_aiR*_sumGPR);
              var _bbR=Math.round(d.hnRn*d.bbPct/100),_bbSt=Math.round(_bbR*_sumGPR);
              var _hbR=Math.round(d.hnRn*d.hbPct/100),_hbSt=Math.round(_hbR*_sumGPR);
              var _roR=Math.round(d.hnRn*d.roPct/100),_roSt=Math.round(_roR*_sumGPR);
            cellContent = wbStackBar([{p:d.aiPct,c:'#004948'},{p:d.bbPct,c:'#52d9ce'},{p:d.hbPct,c:'#D97706'},{p:d.roPct,c:'#d7f7ed'}])
              + '<div style="display:flex;gap:4px;flex-wrap:wrap;margin-top:3px">'
              + '<span style="font-size:12px;font-family:Lato,sans-serif;color:#004948">AI '+d.aiPct+'% · '+_aiSt+' seats</span>'
              + '<span style="font-size:12px;font-family:Lato,sans-serif;color:#52d9ce">BB '+d.bbPct+'% · '+_bbSt+' seats</span>'
              + '<span style="font-size:12px;font-family:Lato,sans-serif;color:#D97706">HB '+d.hbPct+'% · '+_hbSt+' seats</span>'
              + '<span style="font-size:12px;font-family:Lato,sans-serif;color:#6b7280">RO '+d.roPct+'% · '+_roSt+' seats</span>'
              + '</div>'; }
            break;
          // (co_summary removed — summary now shown in collapsed group header)
          // ── Business Mix — stacked using sub-row dot colors ─────────────────
          case 'biz':
            cellContent = wbStackBar([{p:d.toMix,c:'#004948'},{p:d.dirMix,c:'#52d9ce'},{p:d.otaMix,c:'#D97706'},{p:d.otherMix,c:'#9ca3af'}])
              + '<div style="display:flex;gap:4px;flex-wrap:wrap;margin-top:3px">'
              + '<span style="font-size:12px;font-family:Lato,sans-serif;color:#004948">TO '+d.toMix+'%</span>'
              + '<span style="font-size:12px;font-family:Lato,sans-serif;color:#52d9ce">D '+d.dirMix+'%</span>'
              + '<span style="font-size:12px;font-family:Lato,sans-serif;color:#D97706">OTA '+d.otaMix+'%</span>'
              + '<span style="font-size:12px;font-family:Lato,sans-serif;color:#9ca3af">Oth '+d.otherMix+'%</span>'
              + '</div>';
            break;
        }
        }

      } else {
        // Sub-row
        var v1 = '', v2 = '', remCls = row.isRem ? ' wb-sub-val-rem' : '';
        // Room availability sub-rows (dynamic)
        if (row.rtSub !== undefined) {
          var inv2  = RT_CAPS[row.rtIdx];
          var sold2 = Math.min(inv2, Math.floor(inv2 * d.hotel / 110));
          var toS2  = Math.min(sold2, Math.round(sold2 * d.to / Math.max(1, d.hotel)));
          var otS2  = sold2 - toS2;
          var alloc2 = Math.floor(inv2 * 0.8 + Math.abs((d.dm*(row.rtIdx+3)+d.dd*(row.rtIdx+5))%15));
          var alRem2 = Math.max(0, alloc2 - toS2);
          var avRm2  = Math.max(0, inv2 - sold2);
          if      (row.rtSub === 'to')    v1 = toS2 + ' rm';
          else if (row.rtSub === 'other') v1 = otS2 + ' rm';
          else if (row.rtSub === 'tentative') { var tent2 = Math.max(0, Math.floor(2+Math.abs((d.dm*(row.rtIdx+4)+d.dd*(row.rtIdx+2))%6))); v1 = tent2 + ' rm'; }
          else if (row.rtSub === 'ooo') { var ooo2 = Math.max(0, Math.floor(Math.abs((d.dm*(row.rtIdx+1)+d.dd*(row.rtIdx+3))%4))); v1 = ooo2 + ' rm'; }
          else if (row.rtSub === 'alloc') v1 = alRem2 + ' rm';
          else if (row.rtSub === 'avail') { v1 = avRm2 + ' rm'; remCls = avRm2 === 0 ? ' wb-sub-val-rem' : ''; }
        } else {
        switch (row.id) {
          // close outs
          // close out details
          case 'co_rooms': {
            var _crKey = d.dm+'-'+d.dd;
            if (LOCKED_DAYS.has(_crKey)) { v1 = 'All'; }
            else { var _crR = PARTIAL_CLOSURES[_crKey] || [], _crRt = [];
              _crR.forEach(function(r){ _crRt = _crRt.concat(r.roomTypes); });
              _crRt = _crRt.filter(function(v,i,a){ return a.indexOf(v)===i; });
              v1 = _crRt.length > 0 ? _crRt.join(', ') : '—';
            }
          } break;
          case 'co_boards': {
            var _cbKey = d.dm+'-'+d.dd;
            var bdMap = {ai:'AI',bb:'B&B',hb:'HB',ro:'RO'};
            if (LOCKED_DAYS.has(_cbKey)) { v1 = 'All'; }
            else { var _cbR = PARTIAL_CLOSURES[_cbKey] || [], _cbBd = [];
              _cbR.forEach(function(r){ _cbBd = _cbBd.concat(r.boards); });
              _cbBd = _cbBd.filter(function(v,i,a){ return a.indexOf(v)===i; }).map(function(b){ return bdMap[b]||b; });
              v1 = _cbBd.length > 0 ? _cbBd.join(', ') : '—';
            }
          } break;
          case 'co_tos': {
            var _ctKey = d.dm+'-'+d.dd;
            if (LOCKED_DAYS.has(_ctKey)) { v1 = 'All'; }
            else { var _ctR = PARTIAL_CLOSURES[_ctKey] || [], _ctTo = [];
              _ctR.forEach(function(r){ _ctTo = _ctTo.concat(r.tos); });
              _ctTo = _ctTo.filter(function(v,i,a){ return a.indexOf(v)===i; });
              v1 = _ctTo.length > 0 ? _ctTo.join(', ') : '—';
            }
          } break;
          // occupancy
          case 'occ_tdh':    v1 = d.toRn+' rms';    v2 = d.to+'%';                         break;
          case 'occ_other':  v1 = d.otherRms+' rms'; v2 = d.otherPct+'%';                  break;
          case 'occ_stly':   { var cRn=wvCompare.has('stly')?d.sdlyRn:wvCompare.has('ly')?d.lyRn:wvCompare.has('fcst')?d.fcstRn:d.sdlyRn; var cH=wvCompare.has('stly')?d.sdlyH:wvCompare.has('ly')?d.lyH:wvCompare.has('fcst')?d.fcstH:d.sdlyH; v1=cRn+' rms'; v2=cH+'%'; } break;
          case 'occ_rem':    v1 = d.freeRms+' rms';  v2 = Math.max(0,100-d.hotel)+'%';     break;
          // online/offline
          case 'onoff_on':   v1 = d.onlinePct+'%';                                          break;
          case 'onoff_off':  v1 = (100-d.onlinePct)+'%';                                    break;
          // adr
          case 'adr_t':      v1 = '$'+d.toAdr;                                              break;
          case 'adr_hotel':  v1 = '$'+d.adr;                                                break;
          case 'adr_stly':   v1 = '$'+(wvCompare.has('stly')?d.sdlyA:wvCompare.has('ly')?d.lyA:wvCompare.has('fcst')?d.fcstA:d.sdlyA); break;
          // revenue
          case 'rev_t':      v1 = d.fR(d.toRev);                                            break;
          case 'rev_hotel':  v1 = d.fR(d.hnRev);                                            break;
          case 'rev_stly':   v1 = d.fR(wvCompare.has('stly')?d.sdlyR:wvCompare.has('ly')?d.lyR:wvCompare.has('fcst')?d.fcstR:d.sdlyR); break;
          // rn sold
          case 'rn_t':       v1 = d.toRn+' rms';                                            break;
          case 'rn_hotel':   v1 = d.hnRn+' rms';                                            break;
          case 'rn_stly':    v1 = (wvCompare.has('stly')?d.sdlyRn:wvCompare.has('ly')?d.lyRn:wvCompare.has('fcst')?d.fcstRn:d.sdlyRn)+' rms'; break;
          // revpar
          case 'revpar_h':   v1 = '$'+d.hRevpar;                                            break;
          case 'revpar_stly':v1 = '$'+(wvCompare.has('stly')?d.sdlyRevpar:wvCompare.has('ly')?d.lyRevpar:d.sdlyRevpar); break;
          // pickup
          case 'pickup_t':   v1 = '+'+d.pickup;                                             break;
          case 'pickup_h':   v1 = '+'+d.hPickup;                                            break;
          // avg adults / children
          case 'avga_t':     v1 = d.avgA;                                                   break;
          case 'avga_h':     v1 = d.hAvgA;                                                  break;
          case 'avgc_t':     v1 = d.avgC;                                                   break;
          case 'avgc_h':     v1 = d.hAvgC;                                                  break;
          // total adults / children / guests
          case 'tota_t':     v1 = d.totAT;                                                  break;
          case 'tota_h':     v1 = d.totAH;                                                  break;
          case 'totc_t':     v1 = d.totCT;                                                  break;
          case 'totc_h':     v1 = d.totCH;                                                  break;
          case 'totg_t':     v1 = d.totG;                                                   break;
          case 'totg_h':     v1 = d.hTotG;                                                  break;
          // avg los / lead time
          case 'los_t':      v1 = d.avgLos;                                                 break;
          case 'los_h':      v1 = d.hLos;                                                   break;
          case 'lead_t':     v1 = d.avgLead;                                                break;
          case 'lead_h':     v1 = d.hLead;                                                  break;
          // meal plans (% · rooms)
          case 'mp_ai_h':    { var _aiHRm=Math.round(d.hnRn*d.aiPct/100),_aiHSt=Math.round(_aiHRm*(parseFloat(d.hAvgA)+parseFloat(d.hAvgC))); v1=d.aiPct+'% · '+_aiHRm+'r · '+_aiHSt+' seats'; } break;
          case 'mp_ai_t':    { var _aiTp=Math.max(0,Math.round(d.aiPct*d.toPct*0.9)),_aiTRm=Math.round(d.toRn*d.aiPct/100),_aiTSt=Math.round(_aiTRm*(parseFloat(d.avgA)+parseFloat(d.avgC))); v1=_aiTp+'% · '+_aiTRm+'r · '+_aiTSt+' seats'; } break;
          case 'mp_bb_h':    { var _bbHRm=Math.round(d.hnRn*d.bbPct/100),_bbHSt=Math.round(_bbHRm*(parseFloat(d.hAvgA)+parseFloat(d.hAvgC))); v1=d.bbPct+'% · '+_bbHRm+'r · '+_bbHSt+' seats'; } break;
          case 'mp_bb_t':    { var _bbTp=Math.max(0,Math.round(d.bbPct*d.toPct*0.9)),_bbTRm=Math.round(d.toRn*d.bbPct/100),_bbTSt=Math.round(_bbTRm*(parseFloat(d.avgA)+parseFloat(d.avgC))); v1=_bbTp+'% · '+_bbTRm+'r · '+_bbTSt+' seats'; } break;
          case 'mp_hb_h':    { var _hbHRm=Math.round(d.hnRn*d.hbPct/100),_hbHSt=Math.round(_hbHRm*(parseFloat(d.hAvgA)+parseFloat(d.hAvgC))); v1=d.hbPct+'% · '+_hbHRm+'r · '+_hbHSt+' seats'; } break;
          case 'mp_hb_t':    { var _hbTp=Math.max(0,Math.round(d.hbPct*d.toPct*0.9)),_hbTRm=Math.round(d.toRn*d.hbPct/100),_hbTSt=Math.round(_hbTRm*(parseFloat(d.avgA)+parseFloat(d.avgC))); v1=_hbTp+'% · '+_hbTRm+'r · '+_hbTSt+' seats'; } break;
          case 'mp_ro_h':    { var _roHRm=Math.round(d.hnRn*d.roPct/100),_roHSt=Math.round(_roHRm*(parseFloat(d.hAvgA)+parseFloat(d.hAvgC))); v1=d.roPct+'% · '+_roHRm+'r · '+_roHSt+' seats'; } break;
          case 'mp_ro_t':    { var _roTp=Math.max(0,Math.round(d.roPct*d.toPct*0.9)),_roTRm=Math.round(d.toRn*d.roPct/100),_roTSt=Math.round(_roTRm*(parseFloat(d.avgA)+parseFloat(d.avgC))); v1=_roTp+'% · '+_roTRm+'r · '+_roTSt+' seats'; } break;
          // business mix
          case 'biz_to':     v1 = d.toMix+'%';                                              break;
          case 'biz_dir':    v1 = d.dirMix+'%';                                             break;
          case 'biz_ota':    v1 = d.otaMix+'%';                                             break;
          case 'biz_other':  v1 = d.otherMix+'%';                                           break;
        }
        } // end rtSub else
        // Compare chip for sub-rows (Fcst / LY / STLY — mirrors Forecast behaviour)
        var fcstChip = '';
        if (wvCompare.size > 0 && v1 && row.id.indexOf('_stly') < 0) {
          var _fSeed = Math.abs((d.dm * 7 + d.dd * 13 + (row.rtIdx||0) * 5 + row.id.charCodeAt(row.id.length-1)) % 20);
          var _fNum = parseFloat(String(v1).replace(/[^0-9.\-]/g, ''));
          if (!isNaN(_fNum) && _fNum !== 0) {
            var _fCmpDefs = [{k:'stly',m:0.84+_fSeed*0.004,l:'STLY'},{k:'ly',m:0.89+_fSeed*0.004,l:'LY'},{k:'fcst',m:0.92+_fSeed*0.008,l:'Fc'}];
            _fCmpDefs.filter(function(x){ return wvCompare.has(x.k); }).forEach(function(x) {
              var _fVal = Math.round(_fNum * x.m), _fDiff = _fNum - _fVal;
              var _fClr = _fDiff > 0 ? '#059669' : _fDiff < 0 ? '#dc2626' : '#6b7280';
              var _fIco = _fDiff > 0 ? 'trending_up' : _fDiff < 0 ? 'trending_down' : 'remove';
              fcstChip += '<span style="font-size:10px;color:'+_fClr+';margin-left:4px;display:inline-flex;align-items:center;gap:1px;opacity:0.85">'
                + '<span class="material-icons" style="font-size:11px">'+_fIco+'</span>'
                + x.l+':'+_fVal+'</span>';
            });
          }
        }
        cellContent = '<div class="wb-sub-vals' + remCls + '">'
                    + '<span class="wb-sub-v1">' + v1 + fcstChip + '</span>'
                    + (v2 ? '<span class="wb-sub-v2">' + v2 + '</span>' : '')
                    + '</div>';
      }

      var _dayLocked = LOCKED_DAYS.has(dv.month+'-'+dv.day);
      html += '<div class="wb-data-cell wb-' + row.type + '-cell' + (_dayLocked ? ' wb-col-locked' : '') + '">' + cellContent + '</div>';
    });

    html += '</div>'; // wb-row
  });

  html += '</div>'; // wb-layout
  return html;
}

/* ── Daily B AG Grid ─────────────────────────────────────────────────────── */
function initDailyBGrid(days, month, activeDay, containerEl) {
  var AG = _realAgGrid;
  if (!AG || typeof AG.createGrid !== 'function') {
    containerEl.style.cssText = 'display:flex;flex-direction:column;overflow-x:auto;';
    containerEl.innerHTML = buildDailyBView(days, month, activeDay);
    return;
  }
  if (_dailyBGridApi) { try { _dailyBGridApi.destroy(); } catch(e){} _dailyBGridApi = null; }
  containerEl.innerHTML = '';
  containerEl.style.cssText = '';
  containerEl.style.padding = '0';
  var wrapper = document.createElement('div');
  wrapper.className = 'ag-theme-quartz daily-b-ag-wrap';
  containerEl.appendChild(wrapper);

  var DOW_SHORT = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
  var MNAMES_S  = ['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  var TODAY_WV  = new Date(2026, 2, 9);
  var WV_CAP    = 250;
  var RT_NAMES  = ['Standard','Superior','Deluxe','Suite','Jr. Suite','Family'];
  var RT_CAPS   = [51,36,27,12,15,9];
  var TO_NAMES  = ['Sunshine Tours','Global Adv.','Beach Hols','City Breaks','Adventure'];

  // Per-day data
  var dd7 = days.map(function(dv) {
    var dm = dv.month, dd = dv.day;
    var hh = getOccupancy(dm, dd); var hotel = hh.hotel, to = hh.to;
    var adr   = 150 + Math.abs((dm*47+dd*31)%130);
    var v     = Math.abs((dm*127+dd*53+dm*dd*7+dd*dd*3))%100;
    var toAdr = Math.max(80, adr - 20 - Math.abs((dm*3+dd*7)%15));
    var toRn  = Math.round(WV_CAP * to / 100);
    var hnRn  = Math.round(WV_CAP * hotel / 100);
    var toRev = Math.floor(toRn * toAdr);
    var hnRev = Math.floor(hnRn * adr);
    var otherPct = Math.max(0, hotel - to);
    var otherRms = Math.round(WV_CAP * otherPct / 100);
    var freeRms  = WV_CAP - toRn - otherRms;
    var onlinePct = Math.max(30, Math.min(80, 45 + Math.abs((dm*13+dd*7)%35)));
    var adrBar = Math.min(90, Math.round(toAdr / 280 * 100));
    var revBar = Math.min(90, Math.round(toRev / 4500000 * 100));
    var sdlyH  = Math.max(5, hotel - 9), lyH = Math.max(5, hotel - 6), fcstH = Math.min(100, hotel + 4);
    var sdlyA  = adr - 8, lyA = adr - 4, fcstA = adr + 6;
    var sdlyRn = Math.round(toRn * 0.88), lyRn = Math.round(toRn * 0.93), fcstRn = Math.round(toRn * 1.06);
    var sdlyR  = Math.floor(Math.round(WV_CAP * sdlyH / 100) * sdlyA);
    var lyR    = Math.floor(hnRev * 0.95), fcstR = Math.floor(hnRev * 1.06);
    var revpar = Math.max(50, (adr+80) - 30 - Math.abs((dm*5+dd*3)%20));
    var hRevpar = Math.round(adr * hotel / 100);
    var sdlyRevpar = Math.max(40, revpar - 8), lyRevpar = Math.max(40, revpar - 4);
    var pickup = Math.max(0, Math.floor((v%25+5)*to/Math.max(1,hotel)));
    var hPickup = Math.floor(v%25+5);
    var avgA = (1.8+v%3*0.1).toFixed(1), avgC = (0.3+v%2*0.1).toFixed(1);
    var hAvgA = (parseFloat(avgA)+0.3).toFixed(1), hAvgC = (parseFloat(avgC)+0.1).toFixed(1);
    var totAT = Math.round(toRn*parseFloat(avgA)),  totCT = Math.round(toRn*parseFloat(avgC));
    var totAH = Math.round(hnRn*parseFloat(hAvgA)), totCH = Math.round(hnRn*parseFloat(hAvgC));
    var totG  = Math.round(toRn*(parseFloat(avgA)+parseFloat(avgC)));
    var hTotG = Math.round(hnRn*(parseFloat(hAvgA)+parseFloat(hAvgC)));
    var avgLos = (2.8+v%5*0.3).toFixed(1)+'n', hLos = (2.8+v%5*0.3+0.4).toFixed(1)+'n';
    var avgLead = (18+v%60)+'d', hLead = (18+v%60+12)+'d';
    var availRooms = Math.max(0, 102-Math.floor(hotel*1.02));
    var availGuar  = Math.floor(8+v%5);
    var aiPct = Math.max(45, Math.min(68, 55+(dm*7+dd*3)%14));
    var bbPct = Math.max(14, Math.min(28, 20+(dm*11+dd*5)%11));
    var hbPct = Math.max(6,  Math.min(16, 10+(dm*5+dd*7)%9));
    var roPct = 100 - aiPct - bbPct - hbPct;
    var toPct = to / Math.max(1, hotel);
    var toMix = 28+Math.abs((dm*7+dd*5)%25), dirMix = 30+Math.abs((dm*5+dd*9)%20), otaMix = 20+Math.abs((dm*9+dd*3)%18);
    var otherMix = Math.max(0, 100-toMix-dirMix-otaMix);
    var tcRates = [0,1,2,3,4].map(function(i){ return adr-15+Math.abs((dm*(i+3)+dd*(i+5))%50); });
    var baseRate = adr + 8;
    var isEbbDay = (new Date(2026,dm-1,dd)).getDay() < 3;
    function fR(val){return val>=1000000?'$'+(val/1000000).toFixed(1)+'M':'$'+Math.round(val/1000)+'k';}
    return {dm,dd,hotel,to,adr,toAdr,toRn,hnRn,toRev,hnRev,otherPct,otherRms,freeRms,
            onlinePct,adrBar,revBar,sdlyH,lyH,fcstH,sdlyA,lyA,fcstA,sdlyRn,lyRn,fcstRn,
            sdlyR,lyR,fcstR,revpar,hRevpar,sdlyRevpar,lyRevpar,pickup,hPickup,
            avgA,avgC,hAvgA,hAvgC,totAT,totCT,totAH,totCH,totG,hTotG,
            avgLos,hLos,avgLead,hLead,availRooms,availGuar,
            aiPct,bbPct,hbPct,roPct,toPct,toMix,dirMix,otaMix,otherMix,
            tcRates,baseRate,isEbbDay,fR,v};
  });

  // ── Render helpers ────────────────────────────────────────────────────────
  var C1='#004948', C2='#52d9ce', C3='#D97706', C4='#d7f7ed', CSTLY='#C4FF45', CREM='#445e0d';
  function cmpSfx(s, curr, comp) {
    if (!s || wvCompare.size === 0) return '';
    var clr = '#9ca3af';
    if (curr != null && comp != null && !isNaN(parseFloat(curr)) && !isNaN(parseFloat(comp))) {
      var c = parseFloat(curr), p = parseFloat(comp);
      if (c > p) clr = '#16a34a';
      else if (c < p) clr = '#dc2626';
    }
    return '<span class="wv-cmp-sep"> / </span><span class="wv-cmp-val-txt" style="color:'+clr+'">'+s+'</span>';
  }
  function wbGrad2(clr) {
    if (clr==='#004948') return 'linear-gradient(to right,#004948,#007a75)';
    if (clr==='#52d9ce') return 'linear-gradient(to right,#52d9ce,#8aeee8)';
    if (clr==='#445e0d') return 'linear-gradient(to right,#445e0d,#6a9014)';
    if (clr==='#967EF3') return 'linear-gradient(to right,#967EF3,#a78bfa)';
    if (clr==='#D97706') return 'linear-gradient(to right,#D97706,#F59E0B)';
    if (clr==='#16a34a') return 'linear-gradient(to right,#16a34a,#22c55e)';
    if (clr==='#C4FF45') return 'linear-gradient(to right,#C4FF45,#D4FF73)';
    return clr;
  }
  function bar(pct, clr) {
    return '<div style="display:flex;height:6px;border-radius:2px;overflow:hidden;background:#e5e7eb;width:100%"><div style="width:'+pct+'%;background:'+wbGrad2(clr)+';height:6px"></div></div>';
  }
  function sBar(segs) {
    return '<div style="display:flex;height:6px;border-radius:2px;overflow:hidden;background:#e5e7eb;width:100%">'
      + segs.map(function(s){ return '<div style="width:'+s.p+'%;background:'+wbGrad2(s.c)+';height:6px"></div>'; }).join('')+'</div>';
  }
  function sCell(val, barHtml) {
    return '<div style="font-size:14px;font-weight:400;color:#111827;margin-bottom:4px;padding:0 10px">'+val+'</div>'
      +'<div style="padding:0">'+barHtml+'</div>';
  }
  function rCell(v1, v2, rem) {
    var c = rem ? '#388c3f' : '#111827', c2 = rem ? '#388c3f' : '#6b7280';
    return '<div style="display:flex;justify-content:flex-end;align-items:center;gap:8px;height:100%;padding:0 2px">'
      +'<span style="font-size:14px;color:'+c+'">'+v1+'</span>'
      +(v2?'<span style="font-size:14px;color:'+c2+'">'+v2+'</span>':'')
      +'</div>';
  }
  // Compare chip for TO sub-rows in board view — loops over all active compares
  function _wbCmpChip(numStr, seedN, d) {
    if (wvCompare.size === 0) return '';
    var _s = Math.abs((d.dm * 7 + d.dd * 13 + seedN) % 20);
    var _n = parseFloat(String(numStr).replace(/[^0-9.\-]/g, ''));
    if (isNaN(_n) || _n === 0) return '';
    var _defs = [{k:'stly',m:0.84+_s*0.004,l:'STLY'},{k:'ly',m:0.89+_s*0.004,l:'LY'},{k:'fcst',m:0.92+_s*0.008,l:'Fc'}];
    return _defs.filter(function(x){ return wvCompare.has(x.k); }).map(function(x) {
      var _v = Math.round(_n * x.m), _d2 = _n - _v;
      var _c = _d2 > 0 ? '#059669' : _d2 < 0 ? '#dc2626' : '#6b7280';
      var _i = _d2 > 0 ? 'trending_up' : _d2 < 0 ? 'trending_down' : 'remove';
      return '<span style="font-size:10px;color:'+_c+';margin-left:3px;display:inline-flex;align-items:center;gap:1px;opacity:0.85">'
        +'<span class="material-icons" style="font-size:11px">'+_i+'</span>'+x.l+':'+_v+'</span>';
    }).join('');
  }

  // ── Row builder ───────────────────────────────────────────────────────────
  _dbAllRows = []; _dbGrpRenderrs = [];
  var _gi=0, _si=0, _curGK=null, _curSK=null;
  function grp(lbl,clr){ _curGK='G'+(_gi++); _curSK=null; _dbAllRows.push({type:'grp',lbl:lbl,clr:clr||C1,grpKey:_curGK}); }
  function sect(lbl,clr,dot,fn,noChev){ _curSK='S'+(_si++); _dbAllRows.push({type:'sect',lbl:lbl,clr:clr||C1,dot:dot,fn:fn,grpKey:_curGK,sectKey:_curSK,noChev:!!noChev}); }
  function sub(lbl,dot,isRem,fn){ _dbAllRows.push({type:'sub',lbl:lbl,dot:dot,isRem:isRem||false,fn:fn,grpKey:_curGK,sectKey:_curSK}); }

  // ── Close Outs ─────────────────────────────────────────────────────────────
  grp('Close Outs', '#dc2626');
  var bdMap = {ai:'AI',bb:'B&B',hb:'HB',ro:'RO'};
  sub('Room Types', '#6b7280', false, function(d){
    var k=d.dm+'-'+d.dd;
    if (LOCKED_DAYS.has(k)) return rCell('All');
    var rules=PARTIAL_CLOSURES[k]||[], rt=[];
    rules.forEach(function(r){rt=rt.concat(r.roomTypes);});
    rt=rt.filter(function(v,i,a){return a.indexOf(v)===i;});
    return rCell(rt.length?rt.join(', '):'—');
  });
  sub('Board Types', '#6b7280', false, function(d){
    var k=d.dm+'-'+d.dd;
    if (LOCKED_DAYS.has(k)) return rCell('All');
    var rules=PARTIAL_CLOSURES[k]||[], bd=[];
    rules.forEach(function(r){bd=bd.concat(r.boards);});
    bd=bd.filter(function(v,i,a){return a.indexOf(v)===i;}).map(function(b){return bdMap[b]||b;});
    return rCell(bd.length?bd.join(', '):'—');
  });
  sub('Tour Operators', '#6b7280', false, function(d){
    var k=d.dm+'-'+d.dd;
    if (LOCKED_DAYS.has(k)) return rCell('All');
    var rules=PARTIAL_CLOSURES[k]||[], to=[];
    rules.forEach(function(r){to=to.concat(r.tos);});
    to=to.filter(function(v,i,a){return a.indexOf(v)===i;});
    return rCell(to.length?to.join(', '):'—');
  });

  // ── Daily Metrics ─────────────────────────────────────────────────────────
  grp('Daily Metrics', C1);
  if (wvMetricState.capacity) {
    sect('Occupancy', C1, C1, function(d){ var cs=_wvMultiCmpSfx(d.hotel,d.sdlyH,d.lyH,d.fcstH,function(v){return v+'%';}); return sCell(d.hotel+'%'+cs, sBar([{p:d.to,c:C1},{p:d.otherPct,c:C2}])); });
    sub('Travel Distribution Hubs', C1, false, function(d){ return rCell(d.toRn+' rms', d.to+'%'+_wbCmpChip(d.to, 1, d)); });
    sub('Other Segments', C2, false, function(d){ return rCell(d.otherRms+' rms',d.otherPct+'%'); });
    sub('STLY', CSTLY, false, function(d){ return rCell(d.sdlyRn+' rms',d.sdlyH+'%'); });
    sub('Total Hotel Remaining', CREM, true, function(d){ return rCell(d.freeRms+' rms',Math.max(0,100-d.hotel)+'%',true); });
  }
  if (wvMetricState.onlineOffline) {
    sect('Online / Offline', C1, C1, function(d){ return sCell(d.onlinePct+'%', sBar([{p:d.onlinePct,c:C1},{p:100-d.onlinePct,c:C2}])); });
    sub('Online',  C1, false, function(d){ return rCell(d.onlinePct+'%'); });
    sub('Offline', C2, false, function(d){ return rCell((100-d.onlinePct)+'%'); });
  }
  if (wvMetricState.adr) {
    sect('ADR', C1, C1, function(d){ var cs=_wvMultiCmpSfx(d.toAdr,d.sdlyA,d.lyA,d.fcstA,function(v){return '$'+v;}); return sCell('$'+d.toAdr+cs, bar(d.adrBar,C1)); });
    sub('TO ADR',    C1,    false, function(d){ return rCell('$'+d.toAdr + _wbCmpChip(d.toAdr, 2, d)); });
    sub('Hotel ADR', C2,   false, function(d){ return rCell('$'+d.adr); });
    sub('STLY',     CSTLY, false, function(d){ return rCell('$'+d.sdlyA); });
  }
  if (wvMetricState.revenue) {
    sect('Revenue', C1, C1, function(d){ var cs=_wvMultiCmpSfx(d.toRev,d.sdlyR,d.lyR,d.fcstR,d.fR); return sCell(d.fR(d.toRev)+cs, bar(d.revBar,C1)); });
    sub('TO Revenue',     C1,    false, function(d){ return rCell(d.fR(d.toRev) + _wbCmpChip(d.toRev, 3, d)); });
    sub('Hotel Revenue', C2,    false, function(d){ return rCell(d.fR(d.hnRev)); });
    sub('STLY',          CSTLY, false, function(d){ return rCell(d.fR(d.sdlyR)); });
  }

  // ── More Metrics ──────────────────────────────────────────────────────────
  var hasMore = wvMetricState.dm_rnSold||wvMetricState.dm_trevpar||wvMetricState.dm_pickup||
    wvMetricState.dm_avgAdults||wvMetricState.dm_avgChildren||wvMetricState.dm_totalAdults||
    wvMetricState.dm_totalChildren||wvMetricState.dm_totalGuests||wvMetricState.dm_avgLos||
    wvMetricState.dm_avgLeadTime||wvMetricState.dm_availRooms||wvMetricState.dm_availGuar;
  if (hasMore) {
    grp('More Metrics', C1);
    if (wvMetricState.dm_rnSold) {
      sect('RN Sold', C1, C1, function(d){ var cs=_wvMultiCmpSfx(d.toRn,d.sdlyRn,d.lyRn,d.fcstRn,String); return sCell(d.toRn+cs, bar(Math.round(d.toRn/WV_CAP*100),C1)+'<div style="margin-top:2px">'+bar(Math.round(d.hnRn/WV_CAP*100),C2)+'</div>'); });
      sub('TO RN',     C1,    false, function(d){ return sCell(d.toRn+' rms'+_wbCmpChip(d.toRn, 4, d), bar(Math.round(d.toRn/WV_CAP*100),C1)); });
      sub('Hotel RN', C2,    false, function(d){ return sCell(d.hnRn+' rms', bar(Math.round(d.hnRn/WV_CAP*100),C2)); });
      sub('STLY',     CSTLY, false, function(d){ return sCell(d.sdlyRn+' rms', bar(Math.round(d.sdlyRn/WV_CAP*100),CSTLY)); });
    }
    if (wvMetricState.dm_trevpar) {
      sect('REVPAR', C1, C1, function(d){ var cs=_wvMultiCmpSfx(d.hRevpar,d.sdlyRevpar,d.lyRevpar,null,function(v){return '$'+v;}); return sCell('$'+d.hRevpar+cs, bar(Math.min(90,Math.round(d.hRevpar/4)),C1)); });
      sub('Hotel',    C2,    false, function(d){ return sCell('$'+d.hRevpar, bar(Math.min(90,Math.round(d.hRevpar/4)),C2)); });
      sub('STLY',     CSTLY, false, function(d){ return sCell('$'+d.sdlyRevpar, bar(Math.min(90,Math.round(d.sdlyRevpar/4)),CSTLY)); });
    }
    if (wvMetricState.dm_pickup) {
      var _dhActivePU = pickupDayValues.filter(function(dv, i) { return wvMetricState['dm_pickup_' + i] !== false; });
      if (_dhActivePU.length) {
        (function(activePU) {
          sect('Pickup', C1, C1, function(d) {
            var n = activePU.length;
            var hdrs = '', vals = '';
            activePU.forEach(function(dv, idx) {
              var sc = dv<=1?0.3:dv<=3?0.6:dv<=7?1:Math.min(2,dv/7);
              var pv = Math.max(0, Math.round(d.pickup * sc));
              var pvBar = Math.min(90, pv * 3);
              var hpvBar = Math.min(90, d.hPickup * sc * 3);
              var bL = idx===0?'':'border-left:1px solid #e0e0e0;';
              hdrs += '<div class="wv-pu-fig-cell" style="'+bL+'">'+dv+'</div>';
              vals += '<div class="wv-pu-fig-val-cell" style="'+bL+'">'
                + '<span class="wv-pu-fig-num">+'+pv+'</span>'
                + '<div class="wv-occ-bar-track" style="height:4px;margin:2px 0 1px"><div style="width:'+pvBar+'%;background:'+wbGrad2(C1)+';height:4px"></div></div>'
                + '<div class="wv-occ-bar-track" style="height:4px"><div style="width:'+hpvBar+'%;background:'+wbGrad2(C2)+';height:4px"></div></div>'
                + '</div>';
            });
            return '<div class="wv-pu-fig-wrap">'
              + '<div class="wv-pu-fig-hdr-row" style="grid-template-columns:repeat('+n+',1fr)">'+hdrs+'</div>'
              + '<div class="wv-pu-fig-val-row" style="grid-template-columns:repeat('+n+',1fr)">'+vals+'</div>'
              + '</div>';
          }, true);
        })(_dhActivePU);
      }
    }
    if (wvMetricState.dm_avgAdults) {
      sect('Avg Adults', C1, C1, function(d){ return sCell(d.avgA, bar(Math.min(90,parseFloat(d.avgA)/3*100),C1)+'<div style="margin-top:2px">'+bar(Math.min(90,parseFloat(d.hAvgA)/3*100),C2)+'</div>'); });
      sub('T', C1, false, function(d){ return sCell(d.avgA, bar(Math.min(90,parseFloat(d.avgA)/3*100),C1)); });
      sub('Hotel', C2, false, function(d){ return sCell(d.hAvgA, bar(Math.min(90,parseFloat(d.hAvgA)/3*100),C2)); });
    }
    if (wvMetricState.dm_avgChildren) {
      sect('Avg Children', C1, C1, function(d){ return sCell(d.avgC, bar(Math.min(90,parseFloat(d.avgC)/2*100),C1)+'<div style="margin-top:2px">'+bar(Math.min(90,parseFloat(d.hAvgC)/2*100),C2)+'</div>'); });
      sub('T', C1, false, function(d){ return sCell(d.avgC, bar(Math.min(90,parseFloat(d.avgC)/2*100),C1)); });
      sub('Hotel', C2, false, function(d){ return sCell(d.hAvgC, bar(Math.min(90,parseFloat(d.hAvgC)/2*100),C2)); });
    }
    if (wvMetricState.dm_totalAdults) {
      sect('Total Adults', C1, C1, function(d){ return sCell(String(d.totAT), bar(Math.min(90,Math.round(d.totAT/500*100)),C1)+'<div style="margin-top:2px">'+bar(Math.min(90,Math.round(d.totAH/500*100)),C2)+'</div>'); });
      sub('T', C1, false, function(d){ return sCell(String(d.totAT), bar(Math.min(90,Math.round(d.totAT/500*100)),C1)); });
      sub('Hotel', C2, false, function(d){ return sCell(String(d.totAH), bar(Math.min(90,Math.round(d.totAH/500*100)),C2)); });
    }
    if (wvMetricState.dm_totalChildren) {
      sect('Total Children', C1, C1, function(d){ return sCell(String(d.totCT), bar(Math.min(90,Math.round(d.totCT/100*100)),C1)+'<div style="margin-top:2px">'+bar(Math.min(90,Math.round(d.totCH/100*100)),C2)+'</div>'); });
      sub('T', C1, false, function(d){ return sCell(String(d.totCT), bar(Math.min(90,Math.round(d.totCT/100*100)),C1)); });
      sub('Hotel', C2, false, function(d){ return sCell(String(d.totCH), bar(Math.min(90,Math.round(d.totCH/100*100)),C2)); });
    }
    if (wvMetricState.dm_totalGuests) {
      sect('Total Guests', C1, C1, function(d){ return sCell(String(d.totG), bar(Math.min(90,Math.round(d.totG/600*100)),C1)+'<div style="margin-top:2px">'+bar(Math.min(90,Math.round(d.hTotG/600*100)),C2)+'</div>'); });
      sub('T', C1, false, function(d){ return sCell(String(d.totG), bar(Math.min(90,Math.round(d.totG/600*100)),C1)); });
      sub('Hotel', C2, false, function(d){ return sCell(String(d.hTotG), bar(Math.min(90,Math.round(d.hTotG/600*100)),C2)); });
    }
    if (wvMetricState.dm_avgLos) {
      sect('Avg LOS', C1, C1, function(d){ return sCell(d.avgLos, bar(Math.min(90,parseFloat(d.avgLos)/10*100),C1)+'<div style="margin-top:2px">'+bar(Math.min(90,parseFloat(d.hLos)/10*100),C2)+'</div>'); });
      sub('T', C1, false, function(d){ return sCell(d.avgLos, bar(Math.min(90,parseFloat(d.avgLos)/10*100),C1)); });
      sub('Hotel', C2, false, function(d){ return sCell(d.hLos, bar(Math.min(90,parseFloat(d.hLos)/10*100),C2)); });
    }
    if (wvMetricState.dm_avgLeadTime) {
      sect('Lead Time', C1, C1, function(d){ return sCell(d.avgLead, bar(Math.min(90,parseInt(d.avgLead)/90*100),C1)+'<div style="margin-top:2px">'+bar(Math.min(90,parseInt(d.hLead)/90*100),C2)+'</div>'); });
      sub('T', C1, false, function(d){ return sCell(d.avgLead, bar(Math.min(90,parseInt(d.avgLead)/90*100),C1)); });
      sub('Hotel', C2, false, function(d){ return sCell(d.hLead, bar(Math.min(90,parseInt(d.hLead)/90*100),C2)); });
    }
    if (wvMetricState.dm_availRooms) {
      sect('Avail Rooms', '#16a34a', '#16a34a', function(d){ return sCell(d.availRooms+' rm', bar(Math.min(90,Math.round(d.availRooms/WV_CAP*100)),'#16a34a')); });
    }
    if (wvMetricState.dm_availGuar) {
      sect('Avail Guar.', C1, C1, function(d){ return sCell(d.availGuar+' rm', bar(Math.min(90,Math.round(d.availGuar/20*100)),C1)); });
    }
  }

  // ── Meal Plans ────────────────────────────────────────────────────────────
  if (wvMetricState.mealsSummary) {
    grp('Meal Plans', C1);
    [['All Inclusive','aiPct'],['Bed & Breakfast','bbPct'],['Half Board','hbPct'],['Room Only','roPct']].forEach(function(mp){
      var k=mp[1];
      sect(mp[0], C1, C1, (function(k){ return function(d){ var rn=Math.round(d.hnRn*d[k]/100); return sCell(d[k]+'% · '+rn+' rms', bar(d[k],C1)); }; })(k));
      sub('TO', C1, false, (function(k){ return function(d){ var pct=Math.max(0,Math.round(d[k]*d.toPct*0.9)); var rn=Math.round(d.toRn*d[k]/100); return sCell(pct+'% · '+rn+' rms', bar(pct,C1)); }; })(k));
      sub('Hotel', C2, false, (function(k){ return function(d){ var rn=Math.round(d.hnRn*d[k]/100); return sCell(d[k]+'% · '+rn+' rms', bar(d[k],C2)); }; })(k));
    });
    sect('Summary', C1, C1, function(d){
      var _aiR=Math.round(d.hnRn*d.aiPct/100), _bbR=Math.round(d.hnRn*d.bbPct/100), _hbR=Math.round(d.hnRn*d.hbPct/100), _roR=Math.round(d.hnRn*d.roPct/100);
      return sBar([{p:d.aiPct,c:C1},{p:d.bbPct,c:C2},{p:d.hbPct,c:C3},{p:d.roPct,c:C4}])
        +'<div style="display:flex;gap:5px;flex-wrap:wrap;margin-top:3px">'
        +'<span style="font-size:12px;color:'+C1+'">AI '+d.aiPct+'% · '+_aiR+'</span>'
        +'<span style="font-size:12px;color:'+C2+'">BB '+d.bbPct+'% · '+_bbR+'</span>'
        +'<span style="font-size:12px;color:'+C3+'">HB '+d.hbPct+'% · '+_hbR+'</span>'
        +'<span style="font-size:12px;color:#6b7280">RO '+d.roPct+'% · '+_roR+'</span></div>';
    });
  }

  // ── Business Mix ──────────────────────────────────────────────────────────
  if (wvMetricState.bizMix) {
    grp('Business Mix', C1);
    sect('Channel Mix', C1, C1, function(d){
      return sBar([{p:d.toMix,c:C1},{p:d.dirMix,c:C2},{p:d.otaMix,c:C3},{p:d.otherMix,c:'#9ca3af'}])
        +'<div style="display:flex;gap:5px;flex-wrap:wrap;margin-top:3px">'
        +'<span style="font-size:12px;color:'+C1+'">TO '+d.toMix+'%</span>'
        +'<span style="font-size:12px;color:'+C2+'">D '+d.dirMix+'%</span>'
        +'<span style="font-size:12px;color:'+C3+'">OTA '+d.otaMix+'%</span>'
        +'<span style="font-size:12px;color:#9ca3af">Oth '+d.otherMix+'%</span></div>';
    });
    sub('TO', C1, false, function(d){ return rCell(d.toMix+'%'); });
    sub('Direct', C2, false, function(d){ return rCell(d.dirMix+'%'); });
    sub('OTA', C3, false, function(d){ return rCell(d.otaMix+'%'); });
    sub('Other', '#9ca3af', false, function(d){ return rCell(d.otherMix+'%'); });
  }

  // ── Room Availability ─────────────────────────────────────────────────────
  if (wvMetricState.avail || wvMetricState.availAlloc) {
    grp('Room Availability', C1);
    RT_NAMES.forEach(function(name, rtI) {
      var inv = RT_CAPS[rtI];
      sect(name, C1, C1, (function(inv,rtI){ return function(d) {
        var sold=Math.min(inv,Math.floor(inv*d.hotel/110));
        var toS=Math.min(sold,Math.round(sold*d.to/Math.max(1,d.hotel)));
        var otS=sold-toS;
        var tent=Math.max(0,Math.floor(2+Math.abs((d.dm*(rtI+4)+d.dd*(rtI+2))%6)));
        var alRem=Math.max(0,Math.floor(inv*0.8+Math.abs((d.dm*(rtI+3)+d.dd*(rtI+5))%15))-toS);
        var avRm=Math.max(0,inv-sold-tent);
        var toP=Math.round(toS/inv*100),otP=Math.round(otS/inv*100),tnP=Math.round(tent/inv*100),alP=Math.round(alRem/inv*100);
        var avClr=avRm<=0?'#16a34a':C1;
        return '<div style="font-size:14px;color:'+avClr+';margin-bottom:5px">'
          +(avRm<=0?'0 available':avRm+' avail')
          +'<span style="font-size:12px;color:#9ca3af;margin-left:4px">/ '+inv+'</span></div>'
          +sBar([{p:toP,c:C1},{p:otP,c:C2},{p:tnP,c:'#967EF3'},{p:alP,c:C3},{p:Math.max(0,100-toP-otP-tnP-alP),c:C4}]);
      }; })(inv,rtI));
      sub('TO Sold', C1, false, (function(inv){ return function(d){ var s=Math.min(inv,Math.floor(inv*d.hotel/110)); return rCell(Math.min(s,Math.round(s*d.to/Math.max(1,d.hotel)))+' rm'); }; })(inv));
      sub('Other Segments', C2, false, (function(inv){ return function(d){ var s=Math.min(inv,Math.floor(inv*d.hotel/110)); var t=Math.min(s,Math.round(s*d.to/Math.max(1,d.hotel))); return rCell((s-t)+' rm'); }; })(inv));
      sub('Tentative Sold (Group)', '#967EF3', false, (function(inv,rtI){ return function(d){ var tent=Math.max(0,Math.floor(2+Math.abs((d.dm*(rtI+4)+d.dd*(rtI+2))%6))); return rCell(tent+' rm'); }; })(inv,rtI));
      sub('Out-of-Order', '#ef4444', false, (function(inv,rtI){ return function(d){ var ooo=Math.max(0,Math.floor(Math.abs((d.dm*(rtI+1)+d.dd*(rtI+3))%4))); return rCell(ooo+' rm'); }; })(inv,rtI));
      sub('Total Hotel Remaining', CREM, true, (function(inv){ return function(d){ var avRm=Math.max(0,inv-Math.min(inv,Math.floor(inv*d.hotel/110))); return sCell(avRm+' rm', bar(Math.min(90,Math.round(avRm/inv*100)),'#16a34a')); }; })(inv));
    });
  }

  // ── Travel Co. Rates ──────────────────────────────────────────────────────
  if (wvMetricState.toRates) {
    grp('Travel Co. Rates', C1);
    TO_NAMES.forEach(function(name, toI) {
      sect(name, C1, C1, (function(toI){ return function(d) {
        var toRate=d.adr-15+Math.abs((d.dm*(toI+3)+d.dd*(toI+5))%50);
        var toAllot=5+Math.abs((d.dm*(toI+2)+d.dd*(toI+3))%20);
        var toUsed=Math.max(0,toAllot-Math.floor(d.hotel/20));
        var promoTxt=d.isEbbDay?'EBB 10%':'Contract', promoClr=d.isEbbDay?'#16a34a':'#2563eb';
        return '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:5px">'
          +'<span style="font-size:14px;color:#111827">$'+toRate+'</span>'
          +'<span style="font-size:12px;color:#9ca3af">'+(toAllot-toUsed)+'r</span>'
          +'<span style="font-size:11px;font-weight:700;padding:1px 5px;border-radius:3px;background:'+promoClr+'22;color:'+promoClr+';border:1px solid '+promoClr+'44">'+promoTxt+'</span></div>'
          +bar(Math.round(toUsed/toAllot*100),C1);
      }; })(toI));
    });
    sect('Base Rate', C1, C1, function(d){ return sCell('$'+d.baseRate, bar(Math.min(90,Math.round(d.baseRate/280*100)),C1)); });
  }

  // ── Group header renderer ─────────────────────────────────────────────────
  function GrpRenderer() { this._iconEl=null; this._grpKey=null; }
  GrpRenderer.prototype.init = function(p) {
    var self=this, r=p.data;
    this._grpKey=r._grpKey;
    var isC=!!_wbCollapsed[r._grpKey], clr=r._clr;
    this.gui=document.createElement('div');
    this.gui.style.cssText='display:flex;align-items:center;gap:9px;width:100%;height:100%;box-sizing:border-box;cursor:pointer;user-select:none;padding:0 14px;background:#f8f9fd;border-left:3px solid '+clr+';border-top:1px solid #dde1e2;border-bottom:1px solid #dde1e2;';
    var iconWrap=document.createElement('span');
    this._iconEl=iconWrap;
    iconWrap.style.cssText='display:inline-flex;align-items:center;justify-content:center;flex-shrink:0;width:20px;height:20px;border-radius:5px;background:'+clr+'22;color:'+clr+';box-shadow:0 1px 3px '+clr+'33;transform:rotate('+(isC?'-90deg':'0deg')+');transition:transform .2s ease;';
    iconWrap.innerHTML='<span class="material-icons" style="font-size:12px">expand_more</span>';
    var label=document.createElement('span');
    label.style.cssText='font-size:10.5px;font-weight:700;text-transform:uppercase;letter-spacing:.65px;color:#1e2d3a;';
    label.textContent=r._lbl;
    this.gui.appendChild(iconWrap); this.gui.appendChild(label);
    this.gui.addEventListener('click', function(){ _toggleDBGrp(r._grpKey); });
    _dbGrpRenderrs.push(self);
  };
  GrpRenderer.prototype._syncChev=function(){ if(this._iconEl) this._iconEl.style.transform='rotate('+(_wbCollapsed[this._grpKey]?'-90deg':'0deg')+')'; };
  GrpRenderer.prototype.getGui=function(){ return this.gui; };
  GrpRenderer.prototype.destroy=function(){ var i=_dbGrpRenderrs.indexOf(this); if(i!==-1) _dbGrpRenderrs.splice(i,1); };

  // ── Day header factory ────────────────────────────────────────────────────
  function makeDayHeader(dv, isActive, isToday, isLocked, dba, evts) {
    var dm=dv.month, dd=dv.day;
    var bg=isLocked?'#374151':isActive?'#006461':isToday?'#125756':'#1a5e5b';
    var tb=isActive?'3px solid rgba(255,255,255,0.5)':isToday?'3px solid rgba(255,255,255,0.3)':isLocked?'3px solid #dc2626':'3px solid transparent';
    var dc=isLocked?'#fca5a5':'#fff', sc=isLocked?'rgba(252,165,165,0.85)':'rgba(255,255,255,0.75)';
    var dbaStr=dba===0?'Today':dba>0?dba+' DBA':'';
    function H(){}
    H.prototype.init=function(){ this.gui=document.createElement('div'); this.gui.style.cssText='background:'+bg+';width:100%;height:100%;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:4px 6px;box-sizing:border-box;gap:2px;border-top:'+tb+';border-right:1px solid rgba(255,255,255,0.12);'; this.gui.innerHTML='<div style="font-weight:700;font-size:13px;color:'+dc+'">'+(isLocked?'🔒 ':'')+DOW_SHORT[new Date(2026,dm-1,dd).getDay()]+' '+dd+'</div><div style="font-size:11px;color:'+sc+';display:flex;align-items:center;gap:4px"><span>'+MNAMES_S[dm]+'</span>'+(dbaStr?'<span style="background:rgba(255,255,255,0.2);border-radius:3px;padding:0 4px;font-size:10px;color:#fff">'+dbaStr+'</span>':'')+(evts?'<span style="width:6px;height:6px;border-radius:50%;background:rgba(255,255,255,0.9);display:inline-block"></span>':'')+'</div>'; };
    H.prototype.getGui=function(){ return this.gui; };
    H.prototype.destroy=function(){};
    return H;
  }

  // ── Column defs ───────────────────────────────────────────────────────────
  var colDefs = [];
  colDefs.push({
    field:'_lbl', headerName:'', pinned:'left', lockPinned:true, width:190,
    suppressMovable:true, resizable:false,
    cellRenderer: function(p) {
      var r=p.data;
      if (r._type==='sect') {
        var el=document.createElement('div');
        el.style.cssText='display:flex;align-items:center;gap:6px;width:100%;height:100%;box-sizing:border-box;user-select:none;padding:0 10px 0 8px;'+(r._noChev?'':'cursor:pointer;');
        if (!r._noChev) {
          var iw=document.createElement('span');
          var isC=!!_wbCollapsed[r._sectKey];
          iw.style.cssText='display:inline-flex;align-items:center;justify-content:center;flex-shrink:0;width:16px;height:16px;border-radius:4px;background:'+r._clr+'18;color:'+r._clr+';transform:rotate('+(isC?'-90deg':'0deg')+');transition:transform .2s ease;';
          iw.innerHTML='<span class="material-icons" style="font-size:11px">expand_more</span>';
          el.appendChild(iw);
          el.addEventListener('click', function(){ _toggleDBSect(r._sectKey); });
        }
        var lbl=document.createElement('span');
        lbl.style.cssText='font-size:14px;font-weight:400;color:#1c1c1c;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;font-family:Lato,sans-serif;';
        lbl.textContent=r._lbl;
        el.appendChild(lbl);
        return el;
      }
      var el2=document.createElement('div');
      el2.style.cssText='display:flex;align-items:center;gap:5px;padding:0 8px 0 26px;width:100%;height:100%;box-sizing:border-box;font-family:Lato,sans-serif;';
      if (r._dot) { var dot=document.createElement('span'); dot.style.cssText='width:9px;height:9px;border-radius:2px;background:'+r._dot+';flex-shrink:0;display:inline-block;'; el2.appendChild(dot); }
      var lb2=document.createElement('span');
      lb2.style.cssText='font-size:14px;color:'+(r._isRem?'#388c3f':'#1c1c1c')+';font-weight:'+(r._isRem?'600':'400')+';white-space:nowrap;overflow:hidden;text-overflow:ellipsis;';
      lb2.textContent=r._lbl; el2.appendChild(lb2);
      return el2;
    },
    cellStyle: function(p) {
      var r=p.data;
      if (r._type==='sect') return {background:'#fff',borderBottom:'1px solid #dde1e2',borderRight:'2px solid #dde1e2',display:'flex',alignItems:'center'};
      return {background:r._isRem?'#fff8f5':'#fff',borderBottom:'1px solid #f3f4f6',borderRight:'2px solid #dde1e2'};
    },
  });

  days.forEach(function(dv, di) {
    var dm=dv.month, dd=dv.day;
    var isToday=dm===3&&dd===9, isActive=dm===month&&dd===activeDay;
    var isLocked=LOCKED_DAYS.has(dm+'-'+dd);
    var dt=new Date(2026,dm-1,dd), dba=Math.round((dt-TODAY_WV)/86400000);
    var evts=(typeof CAL_EVENTS!=='undefined'&&CAL_EVENTS[dm+'-'+dd])?CAL_EVENTS[dm+'-'+dd]:null;
    colDefs.push({
      field:'day'+di, flex:1, minWidth:90, suppressMovable:true, resizable:false,
      headerComponent: makeDayHeader(dv,isActive,isToday,isLocked,dba,evts),
      cellRenderer: function(p) {
        var r=p.data; if(r._type==='grp') return '';
        var d=dd7[di]; return r._fn?r._fn(d):'';
      },
      cellStyle: function(p) {
        var r=p.data;
        if (r._type==='sect') return {background:'#fff',padding:'7px 0 5px',borderBottom:'1px solid #dde1e2',borderRight:'1px solid #e5e7eb'};
        return {background:r._isRem?'#fff8f5':'#fff',padding:'0 10px',borderBottom:'1px solid #f3f4f6',borderRight:'1px solid #e5e7eb',display:'flex',alignItems:'center'};
      },
    });
  });

  // ── Create grid ───────────────────────────────────────────────────────────
  _dailyBGridApi = AG.createGrid(wrapper, {
    columnDefs: colDefs,
    rowData: _getDBVisibleRows(),
    headerHeight: 50,
    domLayout: 'autoHeight',
    suppressHorizontalScroll: false,
    alwaysShowHorizontalScroll: true,
    suppressCellFocus: true,
    suppressRowClickSelection: true,
    getRowHeight: function(p) {
      if (p.data._type==='grp')  return 40;
      if (p.data._type==='sect') return 58;
      return 33;
    },
    isFullWidthRow: function(p) { return p.rowNode.data._type==='grp'; },
    fullWidthCellRenderer: GrpRenderer,
    defaultColDef: { sortable:false, resizable:false },
    getRowStyle: function(p) {
      var t=p.data._type;
      if (t==='grp')  return {borderTop:'1px solid #dde1e2',borderBottom:'1px solid #dde1e2'};
      if (t==='sect') return {background:'#fff'};
      return {background:p.data._isRem?'#fff8f5':'#fff'};
    },
  });
}

// ── Daily H View — horizontal layout with sticky label column ─────────────────
function buildDailyHView(days, activeMonth, activeDay) {
  var DOW_SHORT = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
  var MNAMES_S  = ['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  var TODAY_WV  = new Date(2026, 2, 9);
  var WV_CAP    = 250;

  // ── Per-day computed data ─────────────────────────────────────────────────
  var dayData = days.map(function(dv) {
    var dm = dv.month, dd = dv.day;
    var hh = getOccupancy(dm, dd); var hotel = hh.hotel, to = hh.to;
    var adr = 150 + Math.abs((dm*47+dd*31)%130);
    var v   = Math.abs((dm*127+dd*53+dm*dd*7+dd*dd*3))%100;
    var toAdr = Math.max(80, adr-20-Math.abs((dm*3+dd*7)%15));
    var toRn  = Math.round(WV_CAP * to / 100);
    var hnRn  = Math.round(WV_CAP * hotel / 100);
    var toRev = Math.floor(toRn * toAdr);
    var hnRev = Math.floor(hnRn * adr);
    var otherPct = Math.max(0, hotel - to), freePct = Math.max(0, 100 - hotel);
    var toRms = toRn, otherRms = Math.round(WV_CAP*otherPct/100);
    var freeRms = WV_CAP - toRms - otherRms;
    var fitPct = Math.round(to*0.45), dynPct = Math.round(to*0.35), serPct = to - fitPct - dynPct;
    var onlinePct = Math.max(30, Math.min(80, 45+Math.abs((dm*13+dd*7)%35)));
    var adrBar = Math.min(95, 40+Math.abs((dm*11+dd*19)%55));
    var revBar = Math.min(95, 35+Math.abs((dm*17+dd*13)%60));
    var revpar = Math.max(50, (adr+80)-30-Math.abs((dm*5+dd*3)%20));
    var pickup = Math.max(0, Math.floor((v%25+5)*to/Math.max(1,hotel)));
    var hPickup = Math.floor(v%25+5);
    var sdlyH=Math.max(5,hotel-9), lyH=Math.max(5,hotel-6), fcstH=Math.min(100,hotel+4);
    var sdlyA=adr-8, lyA=adr-4, fcstA=adr+6;
    var sdlyR=Math.floor(hnRev*0.9), lyR=Math.floor(hnRev*0.95), fcstR=Math.floor(hnRev*1.06);
    function fR(v){return v>=1000000?'$'+(v/1000000).toFixed(1)+'M':'$'+Math.round(v/1000)+'k';}
    var avgA=(1.8+v%3*0.1).toFixed(1), avgC=(0.3+v%2*0.1).toFixed(1);
    var hAvgA=(parseFloat(avgA)+0.3).toFixed(1), hAvgC=(parseFloat(avgC)+0.1).toFixed(1);
    var totG=Math.round(toRn*(parseFloat(avgA)+parseFloat(avgC)));
    var hTotG=Math.round(hnRn*(parseFloat(hAvgA)+parseFloat(hAvgC)));
    var totAT=Math.round(toRn*parseFloat(avgA)), totCT=Math.round(toRn*parseFloat(avgC));
    var totAH=Math.round(hnRn*parseFloat(hAvgA)), totCH=Math.round(hnRn*parseFloat(hAvgC));
    var avgLos=(2.8+v%5*0.3).toFixed(1)+'n', hLos=(2.8+v%5*0.3+0.4).toFixed(1)+'n';
    var avgLead=(18+v%60)+'d', hLead=(18+v%60+12)+'d';
    var availRooms=Math.max(0,102-Math.floor(hotel*1.02));
    var availGuar=Math.floor(8+v%5);
    var aiPct=Math.max(45,Math.min(68,55+(dm*7+dd*3)%14));
    var bbPct=Math.max(14,Math.min(28,20+(dm*11+dd*5)%11));
    var hbPct=Math.max(6,Math.min(16,10+(dm*5+dd*7)%9));
    var roPct=100-aiPct-bbPct-hbPct;
    var toPct=to/Math.max(1,hotel);
    var toMix=28+Math.abs((dm*7+dd*5)%25), dirMix=30+Math.abs((dm*5+dd*9)%20), otaMix=20+Math.abs((dm*9+dd*3)%18);
    var otherMix=Math.max(0,100-toMix-dirMix-otaMix);
    var tcRates=[0,1,2,3,4].map(function(i){return adr-15+Math.abs((dm*(i+3)+dd*(i+5))%50);});
    var baseRate=adr+8;
    var isEbbDay=(new Date(2026,dm-1,dd)).getDay()<3;
    var sdlyRn=Math.round(toRn*0.88), lyRn=Math.round(toRn*0.93), fcstRn=Math.round(toRn*1.06);
    var sdlyRevpar=Math.max(40,revpar-8), lyRevpar=Math.max(40,revpar-4);
    return {dm,dd,hotel,to,adr,toAdr,toRn,hnRn,toRev,hnRev,otherPct,freePct,toRms,otherRms,freeRms,
      fitPct,dynPct,serPct,onlinePct,adrBar,revBar,revpar,sdlyRevpar,lyRevpar,pickup,hPickup,
      sdlyH,lyH,fcstH,sdlyA,lyA,fcstA,sdlyR,lyR,fcstR,fR,
      avgA,avgC,hAvgA,hAvgC,totG,hTotG,totAT,totCT,totAH,totCH,
      avgLos,hLos,avgLead,hLead,availRooms,availGuar,
      aiPct,bbPct,hbPct,roPct,toPct,toMix,dirMix,otaMix,otherMix,
      tcRates,baseRate,isEbbDay,sdlyRn,lyRn,fcstRn};
  });

  // ── Cell renderers ─────────────────────────────────────────────────────────
  // Segmented bar (like wv-occ-bar-track)
  function segBar(segs, ticks) {
    return '<div style="height:6px;background:#e5e7eb;border-radius:3px;overflow:hidden;position:relative;margin:2px 0">'
      + segs.map(function(s){return '<div style="position:absolute;top:0;left:'+s.o+'%;width:'+s.w+'%;height:100%;background:'+s.c+'"></div>';}).join('')
      + (ticks||'')
      + '</div>';
  }
  // Simple horizontal bar (T solid, Hotel faded behind)
  function dualBar(tPct, hPct, clr) {
    return '<div style="height:4px;background:#e5e7eb;border-radius:2px;position:relative;margin:3px 0">'
      + (hPct!=null?'<div style="position:absolute;top:0;left:0;height:100%;width:'+Math.min(92,hPct)+'%;background:#d1d5db;border-radius:2px"></div>':'')
      + '<div style="position:absolute;top:0;left:0;height:100%;width:'+Math.min(92,tPct)+'%;background:'+clr+';border-radius:2px"></div>'
      + '</div>';
  }
  // Stacked bar (for meal plans / biz mix)
  function stackBar(segs) {
    return '<div style="height:5px;background:#e5e7eb;border-radius:3px;display:flex;overflow:hidden;margin:3px 0">'
      + segs.map(function(s){return '<div style="width:'+s.p+'%;background:'+s.c+'"></div>';}).join('')
      + '</div>';
  }
  // Ref chips row
  function refChips(pairs) {
    var CSS = {stly:'background:#e0e7ff;color:#4338ca',ly:'background:#dcfce7;color:#15803d',fcst:'background:#fef9c3;color:#a16207'};
    return '<div style="display:flex;gap:2px;flex-wrap:wrap;margin-top:2px">'
      + pairs.filter(Boolean).map(function(p){
          var s=CSS[p.k]||'background:#f3f4f6;color:#374151';
          return '<span style="font-size:7.5px;font-weight:700;padding:1px 4px;border-radius:3px;'+s+'">'+p.lbl+' '+p.v+'</span>';
        }).join('')
      + '</div>';
  }
  // Promo badge
  function promoBadge(d) {
    var clr = d.isEbbDay?'#16a34a':'#2563eb';
    var lbl = d.isEbbDay?'EBB 10%':'Contract';
    return '<span style="font-size:7.5px;font-weight:700;padding:1px 5px;border-radius:3px;background:'+clr+'20;color:'+clr+';border:1px solid '+clr+'44">'+lbl+'</span>';
  }
  // Value + Hotel chip inline
  function valH(tVal, hVal, tClr) {
    return '<div style="display:flex;align-items:baseline;gap:4px;justify-content:space-between">'
      + '<span style="font-size:8px;color:#9ca3af">'+hVal+'</span>'
      + '<span style="font-size:11px;font-weight:800;color:'+tClr+'">'+tVal+'</span>'
      + '</div>';
  }

  // ── Row definitions ─────────────────────────────────────────────────────────
  var ROWS = [];
  function sec(lbl,clr)  { ROWS.push({type:'sec',lbl:lbl,clr:clr||'#374151'}); }
  function par(lbl,clr)  { ROWS.push({type:'par',lbl:lbl,clr:clr||'#374151'}); }
  function row(lbl,fn)   { ROWS.push({type:'row',lbl:lbl,fn:fn}); }

  // ── DAILY METRICS ─────────────────────────────────────────────────────────
  sec('Daily Metrics','#006461');
  par('Occupancy','#006461');
  row('T / Hotel', function(d){
    var segs=[
      {o:0,      w:d.fitPct,   c:'#006461'},
      {o:d.fitPct,w:d.dynPct,  c:'#0891b2'},
      {o:d.fitPct+d.dynPct,w:d.serPct,c:'#6366f1'},
      {o:d.to,   w:Math.max(0,d.otherPct), c:'#5883ed'},
    ];
    return valH(d.to+'%', d.hotel+'%', '#006461')
      + segBar(segs)
      + refChips([
          {k:'stly',lbl:'STLY',v:d.sdlyH+'%'},
          {k:'ly',  lbl:'LY',  v:d.lyH+'%'},
          {k:'fcst',lbl:'Fcst',v:d.fcstH+'%'},
        ]);
  });
  row('Online / Offline', function(d){
    return stackBar([{p:d.onlinePct,c:'#3b82f6'},{p:100-d.onlinePct,c:'#f97316'}])
      +'<div style="display:flex;justify-content:space-between">'
      +'<span style="font-size:8px;color:#3b82f6">'+d.onlinePct+'% online</span>'
      +'<span style="font-size:8px;color:#f97316">'+(100-d.onlinePct)+'% offline</span>'
      +'</div>';
  });

  par('ADR','#7c3aed');
  row('T / Hotel', function(d){
    var df=d.toAdr-d.adr;
    return valH('$'+d.toAdr,'$'+d.adr,'#7c3aed')
      + dualBar(d.adrBar, Math.min(95,d.adrBar+12), '#7c3aed')
      + '<div style="display:flex;align-items:center;gap:4px;margin-top:2px">'
      + '<span style="font-size:8px;font-weight:700;color:'+(df<=0?'#16a34a':'#dc2626')+'">'+(df>=0?'+':'−')+'$'+Math.abs(df)+' diff</span>'
      + '</div>'
      + refChips([{k:'stly',lbl:'STLY',v:'$'+d.sdlyA},{k:'ly',lbl:'LY',v:'$'+d.lyA},{k:'fcst',lbl:'Fcst',v:'$'+d.fcstA}]);
  });

  par('Revenue','#ea580c');
  row('T / Hotel', function(d){
    return valH(d.fR(d.toRev), d.fR(d.hnRev), '#ea580c')
      + dualBar(d.revBar, Math.min(95,d.revBar+10), '#ea580c')
      + refChips([{k:'stly',lbl:'STLY',v:d.fR(d.sdlyR)},{k:'ly',lbl:'LY',v:d.fR(d.lyR)},{k:'fcst',lbl:'Fcst',v:d.fR(d.fcstR)}]);
  });

  par('REVPAR','#9333ea');
  row('T / Hotel', function(d){
    return valH('$'+d.revpar, '$'+(d.revpar+22), '#9333ea')
      + dualBar(Math.round(d.revpar/4), Math.round((d.revpar+22)/4), '#9333ea')
      + refChips([{k:'stly',lbl:'STLY',v:'$'+d.sdlyRevpar},{k:'ly',lbl:'LY',v:'$'+d.lyRevpar}]);
  });

  par('Pickup','#16a34a');
  row('T / Hotel', function(d){
    return valH('+'+d.pickup, '+'+d.hPickup, '#16a34a');
  });

  // Segments
  par('Segments (T)','#0891b2');
  row('FIT / Dyn / Series', function(d){
    return stackBar([{p:d.fitPct,c:'#006461'},{p:d.dynPct,c:'#0891b2'},{p:d.serPct,c:'#6366f1'}])
      +'<div style="display:flex;gap:6px;font-size:8px;flex-wrap:wrap;margin-top:1px">'
      +'<span style="color:#006461">FIT '+d.fitPct+'%</span>'
      +'<span style="color:#0891b2">Dyn '+d.dynPct+'%</span>'
      +'<span style="color:#6366f1">Series '+d.serPct+'%</span>'
      +'</div>';
  });

  // ── MORE METRICS ──────────────────────────────────────────────────────────
  sec('More Metrics','#2e65e8');
  par('RN Sold','#2e65e8');
  row('T / Hotel', function(d){
    return valH(d.toRn, d.hnRn, '#2e65e8')
      + dualBar(Math.round(d.toRn/WV_CAP*100), Math.round(d.hnRn/WV_CAP*100), '#2e65e8')
      + refChips([{k:'stly',lbl:'STLY',v:d.sdlyRn},{k:'ly',lbl:'LY',v:d.lyRn},{k:'fcst',lbl:'Fcst',v:d.fcstRn}]);
  });

  par('Avg Adults','#2e65e8');
  row('T / Hotel', function(d){ return valH(d.avgA, d.hAvgA, '#2e65e8'); });

  par('Avg Children','#d33030');
  row('T / Hotel', function(d){ return valH(d.avgC, d.hAvgC, '#d33030'); });

  par('Total Adults','#2e65e8');
  row('T / Hotel', function(d){ return valH(d.totAT, d.totAH, '#2e65e8'); });

  par('Total Children','#d33030');
  row('T / Hotel', function(d){ return valH(d.totCT, d.totCH, '#d33030'); });

  par('Total Guests','#0369a1');
  row('T / Hotel', function(d){ return valH(d.totG, d.hTotG, '#0369a1'); });

  par('Avg LOS','#0891b2');
  row('T / Hotel', function(d){ return valH(d.avgLos, d.hLos, '#0891b2'); });

  par('Lead Time','#6366f1');
  row('T / Hotel', function(d){ return valH(d.avgLead, d.hLead, '#6366f1'); });

  par('Avail Rooms','#16a34a');
  row('Hotel', function(d){ return '<span style="font-size:11px;font-weight:800;color:#16a34a">'+d.availRooms+' rm</span>'; });

  par('Avail Guar.','#ea580c');
  row('T', function(d){ return '<span style="font-size:11px;font-weight:800;color:#ea580c">'+d.availGuar+' rm</span>'; });

  // ── MEAL PLANS ────────────────────────────────────────────────────────────
  sec('Meal Plans','#967EF3');
  var mpDefs = [['All Inclusive','#006461','aiPct'],['Bed & Bkfst','#3b82f6','bbPct'],['Half Board','#967EF3','hbPct'],['Room Only','#f59e0b','roPct']];
  mpDefs.forEach(function(mp){
    par(mp[0],mp[1]);
    var key=mp[2];
    row('Hotel / TO %', function(d){
      var hPct = d[key];
      var toPct2 = Math.max(0, Math.round(hPct * d.toPct * 0.9));
      return valH(hPct+'%', 'TO '+toPct2+'%', mp[1])
        + dualBar(hPct, null, mp[1]);
    });
  });
  // Summary stacked bar
  par('Summary','#967EF3');
  row('AI / BB / HB / RO', function(d){
    return stackBar([{p:d.aiPct,c:'#006461'},{p:d.bbPct,c:'#3b82f6'},{p:d.hbPct,c:'#967EF3'},{p:d.roPct,c:'#f59e0b'}])
      +'<div style="display:flex;gap:5px;font-size:8px;flex-wrap:wrap">'
      +'<span style="color:#006461">AI '+d.aiPct+'%</span>'
      +'<span style="color:#3b82f6">BB '+d.bbPct+'%</span>'
      +'<span style="color:#967EF3">HB '+d.hbPct+'%</span>'
      +'<span style="color:#f59e0b">RO '+d.roPct+'%</span>'
      +'</div>';
  });

  // ── BUSINESS MIX ─────────────────────────────────────────────────────────
  sec('Business Mix','#0284c7');
  par('TO / Direct / OTA','#0284c7');
  row('Mix %', function(d){
    return stackBar([{p:d.toMix,c:'#006461'},{p:d.dirMix,c:'#0284c7'},{p:d.otaMix,c:'#D97706'},{p:d.otherMix,c:'#9ca3af'}])
      +'<div style="display:flex;gap:5px;font-size:8px;flex-wrap:wrap">'
      +'<span style="color:#006461">TO '+d.toMix+'%</span>'
      +'<span style="color:#0284c7">Direct '+d.dirMix+'%</span>'
      +'<span style="color:#D97706">OTA '+d.otaMix+'%</span>'
      +'<span style="color:#9ca3af">Other '+d.otherMix+'%</span>'
      +'</div>';
  });

  // ── TRAVEL CO. RATES ─────────────────────────────────────────────────────
  sec('Travel Co. Rates','#0f766e');
  var toOps=[['Sunshine Tours','#3b82f6'],['Global Adv.','#967EF3'],['Beach Hols','#0ea5e9'],['City Breaks','#10b981'],['Adventure','#f59e0b']];
  toOps.forEach(function(op,i){
    par(op[0],op[1]);
    row('Rate / Promo', function(d){
      return '<div style="display:flex;align-items:center;justify-content:space-between;gap:4px">'
        +'<span style="font-size:11px;font-weight:800;color:'+op[1]+'">$'+d.tcRates[i]+'</span>'
        +promoBadge(d)
        +'</div>';
    });
  });
  par('Base Rate','#9333ea');
  row('Rate', function(d){
    return '<span style="font-size:11px;font-weight:800;color:#9333ea">$'+d.baseRate+'</span>';
  });

  // ── Build table ─────────────────────────────────────────────────────────────
  var LABEL_W = '140px';
  var thBase = 'padding:5px 8px;font-size:10px;font-weight:700;text-align:center;border-left:1px solid rgba(255,255,255,.2)';

  var hdrRow = '<tr>'
    +'<th style="position:sticky;left:0;z-index:6;background:#1a5e5b;color:#fff;padding:6px 10px;min-width:'+LABEL_W+';text-align:left;border-right:2px solid #006461;font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.4px">Metric</th>';
  days.forEach(function(dv) {
    var dm=dv.month, dd=dv.day;
    var isToday=dm===3&&dd===9, isActive=dm===activeMonth&&dd===activeDay;
    var isLocked=LOCKED_DAYS.has(dm+'-'+dd);
    var dt=new Date(2026,dm-1,dd);
    var dba=Math.round((dt-TODAY_WV)/86400000);
    var dbaStr=dba===0?'Today':dba>0?dba+' DBA':'';
    var evts=(typeof CAL_EVENTS!=='undefined'&&CAL_EVENTS[dm+'-'+dd])?CAL_EVENTS[dm+'-'+dd]:null;
    var bg=isLocked?'#dc2626':isActive?'#006461':isToday?'#0d8a87':'#1a5e5b';
    var bl=isActive?'2px solid #C4FF45':isToday?'2px solid rgba(255,255,255,.5)':'1px solid rgba(255,255,255,.15)';
    hdrRow+='<th style="'+thBase+';background:'+bg+';border-left:'+bl+';min-width:130px;color:#fff;vertical-align:top">'
      +'<div style="font-weight:800;font-size:11px">'+(isLocked?'🔒 ':'')+MNAMES_S[dm]+' '+dd+'</div>'
      +'<div style="display:flex;align-items:center;justify-content:center;gap:4px;font-size:8px;opacity:.85">'
      +'<span>'+DOW_SHORT[dt.getDay()]+'</span>'
      +(dbaStr?'<span>'+dbaStr+'</span>':'')
      +(evts?'<span style="width:7px;height:7px;border-radius:2px;background:#C4FF45;display:inline-block" title="'+evts.map(function(e){return e.name;}).join(', ')+'"></span>':'')
      +'</div>'
      +'</th>';
  });
  hdrRow += '</tr>';

  var dataRows = ROWS.map(function(r, ri) {
    if (r.type === 'sec') {
      return '<tr><td colspan="'+(days.length+1)+'" style="background:'+r.clr+';color:#fff;font-size:9px;font-weight:800;text-transform:uppercase;letter-spacing:.6px;padding:5px 10px;position:sticky;left:0">'+r.lbl+'</td></tr>';
    }
    if (r.type === 'par') {
      return '<tr style="background:#f8fafc"><td style="position:sticky;left:0;z-index:4;background:#f1f5f9;padding:4px 10px 4px 14px;font-size:9px;font-weight:700;color:'+r.clr+';border-right:2px solid #006461;border-bottom:1px solid #e5e7eb;white-space:nowrap">'
        +r.lbl+'</td>'
        +days.map(function(){return '<td style="background:#f8fafc;border-bottom:1px solid #e5e7eb;border-left:1px solid #f0f0f0"></td>';}).join('')
        +'</tr>';
    }
    // value row — alternating bg
    var bg = ri%2===0?'#fff':'#fafafa';
    var cells = days.map(function(dv, di) {
      var d = dayData[di];
      var isLocked = LOCKED_DAYS.has(d.dm+'-'+d.dd);
      if (isLocked) {
        return '<td style="text-align:center;padding:5px 8px;font-size:11px;color:#9ca3af;border-left:1px solid #f3f4f6;border-bottom:1px solid #f3f4f6">—</td>';
      }
      var html = r.fn(d);
      return '<td style="padding:5px 8px;vertical-align:top;border-left:1px solid #f3f4f6;border-bottom:1px solid #f3f4f6;min-width:130px">'+html+'</td>';
    }).join('');
    return '<tr style="background:'+bg+'">'
      +'<td style="position:sticky;left:0;z-index:4;background:'+bg+';padding:4px 10px 4px 24px;font-size:8.5px;font-weight:600;color:#6b7280;border-right:2px solid #006461;border-bottom:1px solid #f3f4f6;white-space:nowrap">'+r.lbl+'</td>'
      +cells+'</tr>';
  }).join('');

  return '<div class="wv-report-wrap"><table class="wv-report-tbl"><thead>'+hdrRow+'</thead><tbody>'+dataRows+'</tbody></table></div>';
}

window.calToggleMonthlySummary = function() {
  var detail = document.getElementById('calMsDetail');
  var chev   = document.getElementById('calMsChev');
  if (!detail) return;
  var open = detail.style.display === 'none' || !detail.style.display;
  detail.style.display = open ? 'block' : 'none';
  if (chev) chev.style.transform = open ? 'rotate(180deg)' : '';
  window._calMsSummaryOpen = open;
};

var _wv7dAccState = {};
var _wv7dSummaryData = null;

window.wv7dToggle = function(id) {
  _wv7dAccState[id] = !_wv7dAccState[id];
  var c = document.getElementById('wvSummaryContainer');
  if (c && _wv7dSummaryData) c.innerHTML = _buildWv7dSummaryHtml(_wv7dSummaryData);
};

window._buildWv7dSummaryHtml = function(d) {
  var WV = 250;
  var tcOps = [['Sunshine Tours','#3b82f6'],['Global Adv.','#967EF3'],['Beach Hols','#0ea5e9'],['City Breaks','#10b981'],['Adventure','#f59e0b']];
  var chevUp   = '<span class="material-icons" style="font-size:16px">expand_less</span>';
  var chevDown = '<span class="material-icons" style="font-size:16px">expand_more</span>';

  function bar(pct, clr) {
    return '<div class="wv-occ-bar-track"><div style="width:'+Math.min(92,pct)+'%;background:'+clr+';height:6px"></div></div>';
  }
  function stackBar(segs) {
    return '<div class="wv-occ-bar-track">'
      +segs.map(function(s){return '<div style="width:'+s.p+'%;background:'+s.c+';height:6px"></div>';}).join('')+'</div>';
  }
  function refChips(pairs) {
    var CSS={stly:'background:#e0e7ff;color:#4338ca',ly:'background:#dcfce7;color:#15803d',fcst:'background:#fef9c3;color:#a16207'};
    return '<div style="display:flex;gap:2px;flex-wrap:wrap;margin-top:2px">'
      +pairs.map(function(p){return '<span style="font-size:7px;font-weight:700;padding:1px 4px;border-radius:3px;'+CSS[p.k]+'">'+p.l+' '+p.v+'</span>';}).join('')+'</div>';
  }

  // Row definitions (same groups as monthly)
  var rows = [];
  rows.push({type:'top', id:'wv7d_co', label:'Close Outs'});
  rows.push({type:'sect', id:'mos_co_full', label:'Full Close Out', parent:'wv7d_co'});
  rows.push({type:'sect', id:'mos_co_part', label:'Partial Lock', parent:'wv7d_co'});

  rows.push({type:'top', id:'wv7d_daily', label:'Daily Metrics'});
  rows.push({type:'sect', id:'mos_occ', label:'Occupancy', parent:'wv7d_daily'});
  rows.push({type:'sub', id:'mos_occ_to',   label:'TO',    dot:'#004948', parent:'mos_occ', gp:'wv7d_daily'});
  rows.push({type:'sub', id:'mos_occ_htl',  label:'Hotel', dot:'#52d9ce', parent:'mos_occ', gp:'wv7d_daily'});
  rows.push({type:'sub', id:'mos_occ_stly', label:'STLY',  dot:'#818cf8', parent:'mos_occ', gp:'wv7d_daily'});
  rows.push({type:'sect', id:'mos_adr', label:'ADR', parent:'wv7d_daily'});
  rows.push({type:'sub', id:'mos_adr_to',  label:'TO ADR',    dot:'#004948', parent:'mos_adr', gp:'wv7d_daily'});
  rows.push({type:'sub', id:'mos_adr_htl', label:'Hotel ADR', dot:'#52d9ce', parent:'mos_adr', gp:'wv7d_daily'});
  rows.push({type:'sect', id:'mos_rev', label:'Revenue', parent:'wv7d_daily'});
  rows.push({type:'sub', id:'mos_rev_to',  label:'TO Revenue',    dot:'#004948', parent:'mos_rev', gp:'wv7d_daily'});
  rows.push({type:'sub', id:'mos_rev_htl', label:'Hotel Revenue', dot:'#52d9ce', parent:'mos_rev', gp:'wv7d_daily'});
  rows.push({type:'sect', id:'mos_revpar', label:'REVPAR', parent:'wv7d_daily'});
  rows.push({type:'sub', id:'mos_revpar_stly', label:'STLY', dot:'#818cf8', parent:'mos_revpar', gp:'wv7d_daily'});
  rows.push({type:'sect', id:'mos_pickup', label:'Pickup', parent:'wv7d_daily'});
  rows.push({type:'sect', id:'mos_onoff', label:'Online / Offline', parent:'wv7d_daily'});
  rows.push({type:'sub', id:'mos_onoff_on',  label:'Online',  dot:'#3b82f6', parent:'mos_onoff', gp:'wv7d_daily'});
  rows.push({type:'sub', id:'mos_onoff_off', label:'Offline', dot:'#f97316', parent:'mos_onoff', gp:'wv7d_daily'});

  rows.push({type:'top', id:'wv7d_seg', label:'Segment Mix'});
  rows.push({type:'sect', id:'mos_segbar', label:'Summary', parent:'wv7d_seg'});
  rows.push({type:'sub', id:'mos_seg_fit', label:'Static FIT',  dot:'#006461', parent:'mos_segbar', gp:'wv7d_seg'});
  rows.push({type:'sub', id:'mos_seg_dyn', label:'TO Dynamic',  dot:'#0891b2', parent:'mos_segbar', gp:'wv7d_seg'});
  rows.push({type:'sub', id:'mos_seg_ser', label:'Tour Series', dot:'#6366f1', parent:'mos_segbar', gp:'wv7d_seg'});
  rows.push({type:'sub', id:'mos_seg_oth', label:'Other Segs',  dot:'#5883ed', parent:'mos_segbar', gp:'wv7d_seg'});
  rows.push({type:'sub', id:'mos_seg_rem', label:'Remaining',   dot:'#9ca3af', parent:'mos_segbar', gp:'wv7d_seg', isRem:true});

  rows.push({type:'top', id:'wv7d_more', label:'More Metrics'});
  rows.push({type:'sect', id:'mos_rn',    label:'RN Sold',       parent:'wv7d_more'});
  rows.push({type:'sub',  id:'mos_rn_stly', label:'STLY', dot:'#818cf8', parent:'mos_rn', gp:'wv7d_more'});
  rows.push({type:'sect', id:'mos_avga',  label:'Avg Adults',    parent:'wv7d_more'});
  rows.push({type:'sect', id:'mos_avgc',  label:'Avg Children',  parent:'wv7d_more'});
  rows.push({type:'sect', id:'mos_tota',  label:'Total Adults',  parent:'wv7d_more'});
  rows.push({type:'sect', id:'mos_totc',  label:'Total Children',parent:'wv7d_more'});
  rows.push({type:'sect', id:'mos_totg',  label:'Total Guests',  parent:'wv7d_more'});
  rows.push({type:'sect', id:'mos_los',   label:'Avg LOS',       parent:'wv7d_more'});
  rows.push({type:'sect', id:'mos_lead',  label:'Lead Time',     parent:'wv7d_more'});
  rows.push({type:'sect', id:'mos_avail', label:'Avail Rooms',   parent:'wv7d_more'});
  rows.push({type:'sect', id:'mos_availg',label:'Avail Guar.',   parent:'wv7d_more'});

  rows.push({type:'top', id:'wv7d_meals', label:'Meal Plans'});
  rows.push({type:'sect', id:'mos_mpsum', label:'Summary', parent:'wv7d_meals'});
  rows.push({type:'sub', id:'mos_mp_ai', label:'All Inclusive',  dot:'#006461', parent:'mos_mpsum', gp:'wv7d_meals'});
  rows.push({type:'sub', id:'mos_mp_bb', label:'Bed & Breakfast',dot:'#3b82f6', parent:'mos_mpsum', gp:'wv7d_meals'});
  rows.push({type:'sub', id:'mos_mp_hb', label:'Half Board',     dot:'#967EF3', parent:'mos_mpsum', gp:'wv7d_meals'});
  rows.push({type:'sub', id:'mos_mp_ro', label:'Room Only',      dot:'#f59e0b', parent:'mos_mpsum', gp:'wv7d_meals'});

  rows.push({type:'top', id:'wv7d_biz', label:'Business Mix'});
  rows.push({type:'sect', id:'mos_bizbar', label:'Summary', parent:'wv7d_biz'});
  rows.push({type:'sub', id:'mos_biz_to',  label:'TO',     dot:'#006461', parent:'mos_bizbar', gp:'wv7d_biz'});
  rows.push({type:'sub', id:'mos_biz_dir', label:'Direct', dot:'#0284c7', parent:'mos_bizbar', gp:'wv7d_biz'});
  rows.push({type:'sub', id:'mos_biz_ota', label:'OTA',    dot:'#D97706', parent:'mos_bizbar', gp:'wv7d_biz'});
  rows.push({type:'sub', id:'mos_biz_oth', label:'Other',  dot:'#9ca3af', parent:'mos_bizbar', gp:'wv7d_biz'});

  rows.push({type:'top', id:'wv7d_tc', label:'Travel Co. Rates'});
  tcOps.forEach(function(op,i){ rows.push({type:'sect', id:'mos_tc'+i, label:op[0], parent:'wv7d_tc', toIdx:i, toClr:op[1]}); });
  rows.push({type:'sect', id:'mos_tcbase', label:'Base Seg. Rate', parent:'wv7d_tc', toBase:true});

  function isHidden(row) {
    if(row.type==='top') return false;
    if(row.type==='sect' && _wv7dAccState[row.parent]) return true;
    if(row.type==='sub'){ if(_wv7dAccState[row.gp]) return true; if(_wv7dAccState[row.parent]) return true; }
    return false;
  }

  var promoLabel = d.isEbbWeek ? 'EBB 10%' : 'Contract';
  var promoClr   = d.isEbbWeek ? '#16a34a' : '#2563eb';

  var html = '<div class="wb-layout">';
  rows.forEach(function(row) {
    var collapsed = !!_wv7dAccState[row.id];
    var hidden = isHidden(row);
    html += '<div class="wb-row wb-row-'+row.type+(hidden?' wb-row-hidden':'')+'">';

    // Label cell
    if(row.type==='top'){
      html += '<div class="wb-label-cell wb-grp-hdr" onclick="wv7dToggle(\''+row.id+'\')">'
             +'<span class="wb-chev">'+(collapsed?chevDown:chevUp)+'</span>'
             +'<span class="wb-grp-label">'+row.label+'</span></div>';
    } else if(row.type==='sect'){
      html += '<div class="wb-label-cell wb-sect-lbl" onclick="wv7dToggle(\''+row.id+'\')">'
             +'<span class="wb-chev">'+(collapsed?chevDown:chevUp)+'</span>'
             +'<span class="wb-sect-label">'+row.label+'</span></div>';
    } else {
      html += '<div class="wb-label-cell wb-sub-lbl-cell">'
             +(row.dot?'<span class="wb-sub-dot" style="background:'+row.dot+'"></span>':'')
             +'<span class="wb-sub-label'+(row.isRem?' wb-sub-lbl-rem':'')+'">'+(row.label)+'</span></div>';
    }

    // Data cell
    var cc = '';
    if(row.type==='top'){
      cc = '';
    } else if(row.type==='sect'){
      switch(row.id){
        case 'mos_co_full':
          cc = d.fullCoCount7>0
            ? '<div class="wb-sect-val"><span class="material-icons" style="font-size:13px;color:#fca5a5;vertical-align:middle;margin-right:3px">lock</span><span class="wv-occ-total" style="color:#ef4444">'+d.fullCoCount7+' day'+(d.fullCoCount7!==1?'s':'')+'</span><span style="font-size:10px;color:#9ca3af;margin-left:6px">/ '+d.n7+'</span></div>'+bar(Math.min(90,Math.round(d.fullCoCount7/d.n7*100)),'#ef4444')
            : '<div class="wb-sect-val" style="color:#9ca3af;font-size:12px">None</div>';
          break;
        case 'mos_co_part':
          cc = d.partCoCount7>0
            ? '<div class="wb-sect-val"><span class="material-icons" style="font-size:13px;color:#fde68a;vertical-align:middle;margin-right:3px">lock_open</span><span class="wv-occ-total" style="color:#d97706">'+d.partCoCount7+' day'+(d.partCoCount7!==1?'s':'')+'</span><span style="font-size:10px;color:#9ca3af;margin-left:6px">/ '+d.n7+'</span></div>'+bar(Math.min(90,Math.round(d.partCoCount7/d.n7*100)),'#f59e0b')
            : '<div class="wb-sect-val" style="color:#9ca3af;font-size:12px">None</div>';
          break;
        case 'mos_occ':
          cc = '<div class="wb-sect-val"><span class="wv-occ-total">'+d.avgHotel+'%</span></div>'
             +'<div class="wv-occ-bar-track"><div style="width:'+d.avgTo+'%;background:#004948;height:6px"></div><div style="width:'+Math.max(0,d.avgHotel-d.avgTo)+'%;background:#52d9ce;height:6px"></div></div>'
             +refChips([{k:'stly',l:'STLY',v:d.sdlyTo+'%'},{k:'ly',l:'LY',v:d.lyTo+'%'},{k:'fcst',l:'Fcst',v:d.fcstTo+'%'}]);
          break;
        case 'mos_adr':
          cc = '<div class="wb-sect-val"><span class="wv-occ-total">$'+d.avgToAdr+'</span></div>'
             +'<div class="wv-occ-bar-track"><div style="width:'+Math.round(d.avgToAdr/3.5)+'%;background:#004948;height:6px"></div></div>'
             +refChips([{k:'stly',l:'STLY',v:'$'+d.sdlyAdr},{k:'ly',l:'LY',v:'$'+d.lyAdr},{k:'fcst',l:'Fcst',v:'$'+d.fcstAdr}]);
          break;
        case 'mos_rev':
          cc = '<div class="wb-sect-val"><span class="wv-occ-total">'+d.totalRevStr+'</span></div>'
             +refChips([{k:'stly',l:'STLY',v:d.sdlyRev},{k:'ly',l:'LY',v:d.lyRev},{k:'fcst',l:'Fcst',v:d.fcstRev}]);
          break;
        case 'mos_revpar':
          cc = '<div class="wb-sect-val"><span class="wv-occ-total">$'+d.avgRevpar+'</span></div>'
             +'<div class="wv-occ-bar-track"><div style="width:'+Math.round(d.avgRevpar/4)+'%;background:#004948;height:6px"></div></div>'
             +refChips([{k:'stly',l:'STLY',v:'$'+d.sdlyRevpar},{k:'ly',l:'LY',v:'$'+d.lyRevpar}]);
          break;
        case 'mos_pickup':
          cc = '<div class="wb-sect-val"><span class="wv-occ-total">+'+d.sumPickup+'</span><span style="font-size:11px;color:#9ca3af;margin-left:6px">H: +'+d.sumHotelPickup+'</span></div>'
             +'<div class="wv-occ-bar-track"><div style="width:'+Math.min(90,d.sumPickup/10)+'%;background:#004948;height:6px"></div></div>';
          break;
        case 'mos_onoff':
          cc = '<div class="wb-sect-val"><span class="wv-occ-total">'+d.avgOnline+'% online</span></div>'
             +'<div class="wv-occ-bar-track"><div style="width:'+d.avgOnline+'%;background:#004948;height:6px"></div><div style="width:'+(100-d.avgOnline)+'%;background:#52d9ce;height:6px"></div></div>';
          break;
        case 'mos_segbar':
          cc = stackBar([{p:d.avgFitPct,c:'#006461'},{p:d.avgDynPct,c:'#0891b2'},{p:d.avgSerPct,c:'#6366f1'},{p:d.avgOtherPct,c:'#5883ed'},{p:d.avgFreePct,c:'#e5e7eb'}])
             +'<div style="display:flex;gap:4px;flex-wrap:wrap;margin-top:3px">'
             +'<span style="font-size:12px;font-family:Lato,sans-serif;color:#006461">FIT '+d.avgFitPct+'%</span>'
             +'<span style="font-size:12px;font-family:Lato,sans-serif;color:#0891b2">Dyn '+d.avgDynPct+'%</span>'
             +'<span style="font-size:12px;font-family:Lato,sans-serif;color:#6366f1">Ser '+d.avgSerPct+'%</span></div>';
          break;
        case 'mos_rn':
          cc = '<div class="wb-sect-val"><span class="wv-occ-total">'+d.sumRn+' rn</span><span style="font-size:11px;color:#9ca3af;margin-left:6px">H: '+d.avgRnH+'</span></div>'
             +'<div class="wv-occ-bar-track"><div style="width:'+Math.min(92,Math.round(d.sumRn/WV*100))+'%;background:#004948;height:6px"></div></div>'
             +refChips([{k:'stly',l:'STLY',v:d.sdlyRn+' rn'},{k:'ly',l:'LY',v:d.lyRn+' rn'},{k:'fcst',l:'Fcst',v:d.fcstRn+' rn'}]);
          break;
        case 'mos_avga':
          cc = '<div class="wb-sect-val"><span class="wv-occ-total">'+(d.avgTotAdults/Math.max(1,d.sumRn)).toFixed(1)+'</span><span style="font-size:11px;color:#9ca3af;margin-left:6px">H: '+(d.avgHotelTotAdults/Math.max(1,d.avgRnH*d.n7)).toFixed(1)+'</span></div>'
             +'<div class="wv-occ-bar-track"><div style="width:'+Math.min(90,(d.avgTotAdults/Math.max(1,d.sumRn))/3*100)+'%;background:#004948;height:6px"></div></div>';
          break;
        case 'mos_avgc':
          cc = '<div class="wb-sect-val"><span class="wv-occ-total">'+(d.avgTotChildren/Math.max(1,d.sumRn)).toFixed(1)+'</span><span style="font-size:11px;color:#9ca3af;margin-left:6px">H: '+(d.avgHotelTotChildren/Math.max(1,d.avgRnH*d.n7)).toFixed(1)+'</span></div>'
             +'<div class="wv-occ-bar-track"><div style="width:'+Math.min(90,(d.avgTotChildren/Math.max(1,d.sumRn))/2*100)+'%;background:#d33030;height:6px"></div></div>';
          break;
        case 'mos_tota':
          cc = '<div class="wb-sect-val"><span class="wv-occ-total">'+d.avgTotAdults+'</span><span style="font-size:11px;color:#9ca3af;margin-left:6px">H: '+d.avgHotelTotAdults+'</span></div>'; break;
        case 'mos_totc':
          cc = '<div class="wb-sect-val"><span class="wv-occ-total">'+d.avgTotChildren+'</span><span style="font-size:11px;color:#9ca3af;margin-left:6px">H: '+d.avgHotelTotChildren+'</span></div>'; break;
        case 'mos_totg':
          cc = '<div class="wb-sect-val"><span class="wv-occ-total">'+d.avgTotGuests+'</span><span style="font-size:11px;color:#9ca3af;margin-left:6px">H: '+d.avgHotelTotGuests+'</span></div>'; break;
        case 'mos_los':
          cc = '<div class="wb-sect-val"><span class="wv-occ-total">'+d.avgLos+'</span><span style="font-size:11px;color:#9ca3af;margin-left:6px">H: '+d.avgHotelLos+'</span></div>'
             +'<div class="wv-occ-bar-track"><div style="width:'+Math.min(90,parseFloat(d.avgLos)/10*100)+'%;background:#004948;height:6px"></div></div>'; break;
        case 'mos_lead':
          cc = '<div class="wb-sect-val"><span class="wv-occ-total">'+d.avgLead+'</span><span style="font-size:11px;color:#9ca3af;margin-left:6px">H: '+d.avgHotelLead+'</span></div>'
             +'<div class="wv-occ-bar-track"><div style="width:'+Math.min(90,parseInt(d.avgLead)/90*100)+'%;background:#004948;height:6px"></div></div>'; break;
        case 'mos_avail':
          cc = '<div class="wb-sect-val"><span class="wv-occ-total">'+d.avgAvailRooms+' rm</span></div>'
             +'<div class="wv-occ-bar-track"><div style="width:'+Math.min(90,Math.round(d.avgAvailRooms/WV*100))+'%;background:#16a34a;height:6px"></div></div>'; break;
        case 'mos_availg':
          cc = '<div class="wb-sect-val"><span class="wv-occ-total">'+d.avgAvailGuar+' rm</span></div>'
             +'<div class="wv-occ-bar-track"><div style="width:'+Math.min(90,Math.round(d.avgAvailGuar/20*100))+'%;background:#004948;height:6px"></div></div>'; break;
        case 'mos_mpsum':
          { var _7sGPR=d.avgHotelTotGuests/Math.max(1,d.avgRnH*d.n7);
            var _7sAiR=Math.round(d.avgRnH*d.avgAiPct/100),_7sAiSt=Math.round(_7sAiR*_7sGPR);
            var _7sBbR=Math.round(d.avgRnH*d.avgBbPct/100),_7sBbSt=Math.round(_7sBbR*_7sGPR);
            var _7sHbR=Math.round(d.avgRnH*d.avgHbPct/100),_7sHbSt=Math.round(_7sHbR*_7sGPR);
            var _7sRoR=Math.round(d.avgRnH*d.avgRoPct/100),_7sRoSt=Math.round(_7sRoR*_7sGPR);
          cc = stackBar([{p:d.avgAiPct,c:'#006461'},{p:d.avgBbPct,c:'#3b82f6'},{p:d.avgHbPct,c:'#967EF3'},{p:d.avgRoPct,c:'#f59e0b'}])
             +'<div style="display:flex;gap:4px;flex-wrap:wrap;margin-top:3px">'
             +'<span style="font-size:12px;font-family:Lato,sans-serif;color:#006461">AI '+d.avgAiPct+'% · '+_7sAiSt+' seats</span>'
             +'<span style="font-size:12px;font-family:Lato,sans-serif;color:#3b82f6">BB '+d.avgBbPct+'% · '+_7sBbSt+' seats</span>'
             +'<span style="font-size:12px;font-family:Lato,sans-serif;color:#967EF3">HB '+d.avgHbPct+'% · '+_7sHbSt+' seats</span>'
             +'<span style="font-size:12px;font-family:Lato,sans-serif;color:#f59e0b">RO '+d.avgRoPct+'% · '+_7sRoSt+' seats</span></div>'; }
          break;
        case 'mos_bizbar':
          cc = stackBar([{p:d.avgToMix,c:'#006461'},{p:d.avgDirMix,c:'#0284c7'},{p:d.avgOtaMix,c:'#D97706'},{p:d.avgOtherMix,c:'#9ca3af'}])
             +'<div style="display:flex;gap:4px;flex-wrap:wrap;margin-top:3px">'
             +'<span style="font-size:12px;font-family:Lato,sans-serif;color:#006461">TO '+d.avgToMix+'%</span>'
             +'<span style="font-size:12px;font-family:Lato,sans-serif;color:#0284c7">D '+d.avgDirMix+'%</span>'
             +'<span style="font-size:12px;font-family:Lato,sans-serif;color:#D97706">OTA '+d.avgOtaMix+'%</span></div>';
          break;
        case 'mos_tcbase':
          cc = '<div class="wb-sect-val"><span class="wv-occ-total" style="font-weight:700;color:#1C1C1C">$'+(d.avgHotelAdr+8)+'</span></div>'
             +'<div class="wv-occ-bar-track"><div style="width:'+Math.min(90,Math.round((d.avgHotelAdr+8)/280*100))+'%;background:#004948;height:6px"></div></div>';
          break;
        default:
          if(row.toIdx!==undefined){
            cc = '<div class="wb-sect-val" style="justify-content:space-between">'
               +'<span class="wv-occ-total" style="color:#1C1C1C">$'+d.avgTcRates[row.toIdx]+'</span>'
               +'<span style="font-size:11px;font-weight:700;padding:1px 5px;border-radius:3px;background:'+promoClr+'22;color:'+promoClr+';border:1px solid '+promoClr+'44">'+promoLabel+'</span>'
               +'</div>'
               +'<div class="wv-occ-bar-track"><div style="width:'+Math.min(90,Math.round(d.avgTcRates[row.toIdx]/280*100))+'%;background:#004948;height:6px"></div></div>';
          }
          break;
      }
    } else {
      // Sub rows — same as monthly
      var v1 = '';
      switch(row.id){
        case 'mos_occ_to':     v1 = d.avgTo+'%'; break;
        case 'mos_occ_htl':    v1 = d.avgHotel+'%'; break;
        case 'mos_occ_stly':   v1 = d.sdlyTo+'%'; break;
        case 'mos_adr_to':     v1 = '$'+d.avgToAdr; break;
        case 'mos_adr_htl':    v1 = '$'+d.avgHotelAdr; break;
        case 'mos_rev_to':     v1 = d.totalRevStr; break;
        case 'mos_rev_htl':    v1 = d.totalHotelRevStr; break;
        case 'mos_revpar_stly':v1 = '$'+d.sdlyRevpar; break;
        case 'mos_rn_stly':    v1 = d.sdlyRn+' rn'; break;
        case 'mos_onoff_on':   v1 = d.avgOnline+'%'; break;
        case 'mos_onoff_off':  v1 = (100-d.avgOnline)+'%'; break;
        case 'mos_seg_fit':    v1 = d.avgFitPct+'% · '+d.avgFitRms+' rm'; break;
        case 'mos_seg_dyn':    v1 = d.avgDynPct+'% · '+d.avgDynRms+' rm'; break;
        case 'mos_seg_ser':    v1 = d.avgSerPct+'% · '+d.avgSerRms+' rm'; break;
        case 'mos_seg_oth':    v1 = d.avgOtherPct+'% · '+d.avgOtherRms+' rm'; break;
        case 'mos_seg_rem':    v1 = d.avgFreePct+'% · '+d.avgFreeRms+' rm'; break;
        case 'mos_mp_ai':      { var _7gpr=d.avgHotelTotGuests/Math.max(1,d.avgRnH*d.n7),_7aiRm=Math.round(d.avgRnH*d.avgAiPct/100),_7aiSt=Math.round(_7aiRm*_7gpr); v1=d.avgAiPct+'% · '+_7aiRm+'r · '+_7aiSt+' seats'; } break;
        case 'mos_mp_bb':      { var _7gprb=d.avgHotelTotGuests/Math.max(1,d.avgRnH*d.n7),_7bbRm=Math.round(d.avgRnH*d.avgBbPct/100),_7bbSt=Math.round(_7bbRm*_7gprb); v1=d.avgBbPct+'% · '+_7bbRm+'r · '+_7bbSt+' seats'; } break;
        case 'mos_mp_hb':      { var _7gprh=d.avgHotelTotGuests/Math.max(1,d.avgRnH*d.n7),_7hbRm=Math.round(d.avgRnH*d.avgHbPct/100),_7hbSt=Math.round(_7hbRm*_7gprh); v1=d.avgHbPct+'% · '+_7hbRm+'r · '+_7hbSt+' seats'; } break;
        case 'mos_mp_ro':      { var _7gprr=d.avgHotelTotGuests/Math.max(1,d.avgRnH*d.n7),_7roRm=Math.round(d.avgRnH*d.avgRoPct/100),_7roSt=Math.round(_7roRm*_7gprr); v1=d.avgRoPct+'% · '+_7roRm+'r · '+_7roSt+' seats'; } break;
        case 'mos_biz_to':     v1 = d.avgToMix+'%'; break;
        case 'mos_biz_dir':    v1 = d.avgDirMix+'%'; break;
        case 'mos_biz_ota':    v1 = d.avgOtaMix+'%'; break;
        case 'mos_biz_oth':    v1 = d.avgOtherMix+'%'; break;
      }
      cc = '<span class="wb-sub-val">'+v1+'</span>';
    }

    html += '<div class="wb-data-cell">'+cc+'</div>';
    html += '</div>'; // close wb-row
  });
  html += '</div>'; // close wb-layout

  // Wrap in outer accordion (same as monthly "Overview")
  var ovCollapsed = _wv7dAccState['wv7d_overview'] === true;
  var ovChev = ovCollapsed
    ? '<span class="material-icons" style="font-size:16px">expand_more</span>'
    : '<span class="material-icons" style="font-size:16px">expand_less</span>';
  return '<div class="cal-summary-wrap" style="background:#fff">'
    +'<div class="wv-acc-sect'+(ovCollapsed?'':' wv-acc-open')+'" style="border:1px solid #dde1e2;border-radius:0;overflow:hidden">'
    +'<div class="wv-acc-hdr" onclick="wv7dToggle(\'wv7d_overview\')" style="background:#fff;border-bottom:none;border-radius:0">'
    +'<span class="wv-acc-chev" style="color:#006461">'+ovChev+'</span>'
    +'<span class="wv-acc-title" style="font-weight:700">7 Day Metrics Summary</span>'
    +'</div>'
    +'<div class="wv-acc-body'+(ovCollapsed?' wv-body-hidden':'')+'" style="padding:0;background:#fff">'
    +html
    +'</div></div></div>';
};

/* ── Daily-H AG Grid ─────────────────────────────────────────────────────── */
var _dailyHGridApi  = null;
var _dhCollapsed    = {};   // persists collapse state between tab/week switches
var _dhAllRows      = [];   // flat ROWS array; rebuilt each time initDailyHGrid runs
var _dhSecRenderers = [];   // live SecRenderer instances → used for Open/Close All chevron sync
var _dhParCells     = [];   // live par cell DOM refs for chevron sync
var _dhMetricOrder  = null; // null = default; array of parKey strings = custom order
var _dhLastInitArgs = null; // saved args for grid rebuild after reorder
var _wvSectionOrder = null; // null = default; array of section keys for combined view
var _drColOrder     = null; // null = default; array of group names for Daily R
var _drLastInitArgs = null; // saved args for Daily R rebuild

// Section/group definitions used by the reorder modal
var WV_SECTIONS_DEF = [
  { key:'daily',        lbl:'Daily Metrics',        clr:'#006461' },
  { key:'detailed',     lbl:'More Metrics',          clr:'#2e65e8' },
  { key:'meals',        lbl:'Meal Plans',            clr:'#7c3aed' },
  { key:'mealsSummary', lbl:'Meal Plans Summary',    clr:'#7c3aed' },
  { key:'avail',        lbl:'Room Availability',     clr:'#16a34a' },
  { key:'toRates',      lbl:'Travel Company Rates',  clr:'#0f766e' },
  { key:'bizMix',       lbl:'Business Mix',          clr:'#0284c7' },
];
var DR_GROUPS_DEF = [
  { key:'Daily Metrics',    clr:'#006461' },
  { key:'Room Avail.',      clr:'#16a34a' },
  { key:'Segments (T)',     clr:'#0891b2' },
  { key:'Business Mix',     clr:'#0284c7' },
  { key:'Meal Plans',       clr:'#7c3aed' },
  { key:'Travel Co. Rates', clr:'#0f766e' },
];

// Reorder ROWS by a custom par-key sequence (preserves sec headers above their first par)
function reorderDHRows(rows, order) {
  // Parse into groups: { sec, pars: [{par, rows:[]}] }
  var sections = [], curSec = null, curPar = null;
  rows.forEach(function(r) {
    if (r.type === 'sec') { curSec = { sec: r, pars: [] }; sections.push(curSec); curPar = null; }
    else if (r.type === 'par') { curPar = { par: r, rows: [] }; if (curSec) curSec.pars.push(curPar); }
    else if (r.type === 'row') { if (curPar) curPar.rows.push(r); }
  });
  // Build lookup: parKey → { section, parGroup }
  var parMap = {};
  sections.forEach(function(sec) {
    sec.pars.forEach(function(pg) { parMap[pg.par.parKey] = { sec: sec, pg: pg }; });
  });
  // Emit rows in specified order, emitting sec header the first time it appears
  var result = [], usedSecs = {};
  order.forEach(function(pk) {
    var entry = parMap[pk];
    if (!entry) return;
    var secKey = entry.sec.sec.secKey;
    if (!usedSecs[secKey]) { result.push(entry.sec.sec); usedSecs[secKey] = true; }
    result.push(entry.pg.par);
    entry.pg.rows.forEach(function(r) { result.push(r); });
  });
  return result;
}

function _getDHVisibleRowData() {
  return _dhAllRows.filter(function(r) {
    if (r.type === 'sec') return true;
    if (_dhCollapsed[r.secKey]) return false;
    if (r.type === 'par') return true;
    if (_dhCollapsed[r.parKey]) return false;
    return true;
  }).map(function(r) {
    return { _type:r.type, _lbl:r.lbl, _clr:r.clr||'#374151', _fn:r.fn||null,
             _secKey:r.secKey||null, _parKey:r.parKey||null };
  });
}
function _toggleDHSection(secKey) {
  _dhCollapsed[secKey] = !_dhCollapsed[secKey];
  if (_dailyHGridApi) _dailyHGridApi.setGridOption('rowData', _getDHVisibleRowData());
}
function _toggleDHPar(parKey) {
  _dhCollapsed[parKey] = !_dhCollapsed[parKey];
  if (_dailyHGridApi) _dailyHGridApi.setGridOption('rowData', _getDHVisibleRowData());
}
function dhSetAll(collapse) {
  // Set all known rows
  _dhAllRows.forEach(function(r) {
    if (r.type === 'sec') _dhCollapsed[r.secKey] = collapse;
    if (r.type === 'par') _dhCollapsed[r.parKey] = collapse;
  });
  // Also set any keys already in the collapse map to catch stragglers
  for (var k in _dhCollapsed) {
    if (_dhCollapsed.hasOwnProperty(k)) _dhCollapsed[k] = collapse;
  }
  // Sync chevrons on still-visible sec/par renderers before rebuilding rowData
  _dhSecRenderers.forEach(function(sr) { sr._syncChevron(); });
  _dhParCells.forEach(function(pc) {
    var rot = _dhCollapsed[pc._parKey] ? '-90deg' : '0deg';
    if (pc._iconEl) pc._iconEl.style.transform = 'rotate(' + rot + ')';
  });
  if (_dailyHGridApi) _dailyHGridApi.setGridOption('rowData', _getDHVisibleRowData());
}

function initDailyHGrid(days, activeMonth, activeDay, containerEl) {
  _dhLastInitArgs = { days: days, month: activeMonth, day: activeDay, container: containerEl };
  var AG = _realAgGrid;
  if (!AG || typeof AG.createGrid !== 'function') {
    containerEl.innerHTML = buildDailyHView(days, activeMonth, activeDay);
    return;
  }

  if (_dailyHGridApi) { try { _dailyHGridApi.destroy(); } catch(e){} _dailyHGridApi = null; }
  containerEl.innerHTML = '';
  containerEl.style.padding = '0';

  var wrapper = document.createElement('div');
  wrapper.className = 'ag-theme-quartz daily-h-ag-wrap';
  containerEl.appendChild(wrapper);

  var DOW_SHORT = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
  var MNAMES_S  = ['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  var TODAY_WV  = new Date(2026, 2, 9);
  var WV_CAP    = 250;

  // ── Per-day computed data ─────────────────────────────────────────────────
  var dayData = days.map(function(dv) {
    var dm = dv.month, dd = dv.day;
    var hh = getOccupancy(dm, dd); var hotel = hh.hotel, to = hh.to;
    var adr = 150 + Math.abs((dm*47+dd*31)%130);
    var v   = Math.abs((dm*127+dd*53+dm*dd*7+dd*dd*3))%100;
    var toAdr = Math.max(80, adr-20-Math.abs((dm*3+dd*7)%15));
    var toRn  = Math.round(WV_CAP * to / 100);
    var hnRn  = Math.round(WV_CAP * hotel / 100);
    var toRev = Math.floor(toRn * toAdr);
    var hnRev = Math.floor(hnRn * adr);
    var otherPct = Math.max(0, hotel - to), freePct = Math.max(0, 100 - hotel);
    var toRms = toRn, otherRms = Math.round(WV_CAP*otherPct/100);
    var freeRms = WV_CAP - toRms - otherRms;
    var fitPct = Math.round(to*0.45), dynPct = Math.round(to*0.35), serPct = to - fitPct - dynPct;
    var onlinePct = Math.max(30, Math.min(80, 45+Math.abs((dm*13+dd*7)%35)));
    var adrBar = Math.min(95, 40+Math.abs((dm*11+dd*19)%55));
    var revBar = Math.min(95, 35+Math.abs((dm*17+dd*13)%60));
    var revpar = Math.max(50, (adr+80)-30-Math.abs((dm*5+dd*3)%20));
    var pickup = Math.max(0, Math.floor((v%25+5)*to/Math.max(1,hotel)));
    var hPickup = Math.floor(v%25+5);
    var sdlyH=Math.max(5,hotel-9), lyH=Math.max(5,hotel-6), fcstH=Math.min(100,hotel+4);
    var sdlyA=adr-8, lyA=adr-4, fcstA=adr+6;
    var sdlyR=Math.floor(hnRev*0.9), lyR=Math.floor(hnRev*0.95), fcstR=Math.floor(hnRev*1.06);
    function fR(v){return v>=1000000?'$'+(v/1000000).toFixed(1)+'M':'$'+Math.round(v/1000)+'k';}
    var avgA=(1.8+v%3*0.1).toFixed(1), avgC=(0.3+v%2*0.1).toFixed(1);
    var hAvgA=(parseFloat(avgA)+0.3).toFixed(1), hAvgC=(parseFloat(avgC)+0.1).toFixed(1);
    var totG=Math.round(toRn*(parseFloat(avgA)+parseFloat(avgC)));
    var hTotG=Math.round(hnRn*(parseFloat(hAvgA)+parseFloat(hAvgC)));
    var totAT=Math.round(toRn*parseFloat(avgA)), totCT=Math.round(toRn*parseFloat(avgC));
    var totAH=Math.round(hnRn*parseFloat(hAvgA)), totCH=Math.round(hnRn*parseFloat(hAvgC));
    var avgLos=(2.8+v%5*0.3).toFixed(1)+'n', hLos=(2.8+v%5*0.3+0.4).toFixed(1)+'n';
    var avgLead=(18+v%60)+'d', hLead=(18+v%60+12)+'d';
    var availRooms=Math.max(0,102-Math.floor(hotel*1.02));
    var availGuar=Math.floor(8+v%5);
    var aiPct=Math.max(45,Math.min(68,55+(dm*7+dd*3)%14));
    var bbPct=Math.max(14,Math.min(28,20+(dm*11+dd*5)%11));
    var hbPct=Math.max(6,Math.min(16,10+(dm*5+dd*7)%9));
    var roPct=100-aiPct-bbPct-hbPct;
    var toPct=to/Math.max(1,hotel);
    var toMix=28+Math.abs((dm*7+dd*5)%25), dirMix=30+Math.abs((dm*5+dd*9)%20), otaMix=20+Math.abs((dm*9+dd*3)%18);
    var otherMix=Math.max(0,100-toMix-dirMix-otaMix);
    var tcRates=[0,1,2,3,4].map(function(i){return adr-15+Math.abs((dm*(i+3)+dd*(i+5))%50);});
    var baseRate=adr+8;
    var isEbbDay=(new Date(2026,dm-1,dd)).getDay()<3;
    var sdlyRn=Math.round(toRn*0.88), lyRn=Math.round(toRn*0.93), fcstRn=Math.round(toRn*1.06);
    var sdlyRevpar=Math.max(40,revpar-8), lyRevpar=Math.max(40,revpar-4);
    return {dm,dd,hotel,to,adr,toAdr,toRn,hnRn,toRev,hnRev,otherPct,freePct,toRms,otherRms,freeRms,
      fitPct,dynPct,serPct,onlinePct,adrBar,revBar,revpar,sdlyRevpar,lyRevpar,pickup,hPickup,
      sdlyH,lyH,fcstH,sdlyA,lyA,fcstA,sdlyR,lyR,fcstR,fR,
      avgA,avgC,hAvgA,hAvgC,totG,hTotG,totAT,totCT,totAH,totCH,
      avgLos,hLos,avgLead,hLead,availRooms,availGuar,
      aiPct,bbPct,hbPct,roPct,toPct,toMix,dirMix,otaMix,otherMix,
      tcRates,baseRate,isEbbDay,sdlyRn,lyRn,fcstRn};
  });

  // ── Cell render helpers ───────────────────────────────────────────────────
  function segBar(segs) {
    return '<div style="height:6px;background:#e5e7eb;border-radius:3px;overflow:hidden;position:relative;margin:2px 0">'
      + segs.map(function(s){return '<div style="position:absolute;top:0;left:'+s.o+'%;width:'+s.w+'%;height:100%;background:'+s.c+'"></div>';}).join('')
      + '</div>';
  }
  function dualBar(tPct, hPct, clr) {
    return '<div style="height:4px;background:#e5e7eb;border-radius:2px;position:relative;margin:3px 0">'
      + (hPct!=null?'<div style="position:absolute;top:0;left:0;height:100%;width:'+Math.min(92,hPct)+'%;background:#d1d5db;border-radius:2px"></div>':'')
      + '<div style="position:absolute;top:0;left:0;height:100%;width:'+Math.min(92,tPct)+'%;background:'+clr+';border-radius:2px"></div>'
      + '</div>';
  }
  function stackBar(segs) {
    return '<div style="height:5px;background:#e5e7eb;border-radius:3px;display:flex;overflow:hidden;margin:3px 0">'
      + segs.map(function(s){return '<div style="width:'+s.p+'%;background:'+s.c+'"></div>';}).join('')
      + '</div>';
  }
  function refChips(pairs) {
    var CSS = {stly:'background:#e0e7ff;color:#4338ca',ly:'background:#dcfce7;color:#15803d',fcst:'background:#fef9c3;color:#a16207'};
    return '<div style="display:flex;gap:3px;flex-wrap:wrap;margin-top:3px">'
      + pairs.filter(Boolean).map(function(p){
          var s=CSS[p.k]||'background:#f3f4f6;color:#374151';
          return '<span style="font-size:8.5px;font-weight:700;padding:1px 5px;border-radius:3px;'+s+'">'+p.lbl+' '+p.v+'</span>';
        }).join('')
      + '</div>';
  }
  function promoBadge(d) {
    var clr = d.isEbbDay?'#16a34a':'#2563eb';
    var lbl = d.isEbbDay?'EBB 10%':'Contract';
    return '<span style="font-size:8.5px;font-weight:700;padding:1px 6px;border-radius:3px;background:'+clr+'20;color:'+clr+';border:1px solid '+clr+'44">'+lbl+'</span>';
  }
  function valH(tVal, hVal, tClr) {
    return '<div style="display:flex;align-items:baseline;gap:4px;justify-content:space-between">'
      + '<span style="font-size:9px;color:#9ca3af">'+hVal+'</span>'
      + '<span style="font-size:12px;font-weight:800;color:'+tClr+'">'+tVal+'</span>'
      + '</div>';
  }

  // ── Row definitions ───────────────────────────────────────────────────────
  var ROWS = [];
  var _dhSecIdx = 0, _dhParIdx = 0;
  var _dhCurSecKey = null, _dhCurParKey = null;

  function sec(lbl, clr) {
    _dhCurSecKey = 'S' + (_dhSecIdx++);
    _dhCurParKey = null;
    ROWS.push({ type:'sec', lbl:lbl, clr:clr||'#374151', secKey:_dhCurSecKey });
  }
  function par(lbl, clr) {
    _dhCurParKey = 'P' + (_dhParIdx++);
    ROWS.push({ type:'par', lbl:lbl, clr:clr||'#374151', secKey:_dhCurSecKey, parKey:_dhCurParKey });
  }
  function row(lbl, fn) {
    ROWS.push({ type:'row', lbl:lbl, fn:fn, secKey:_dhCurSecKey, parKey:_dhCurParKey });
  }

  // Delegate to module-level (so Open All / Close All can reach them)
  var getDHVisibleRowData = _getDHVisibleRowData;
  var toggleDHSection     = _toggleDHSection;
  var toggleDHPar         = _toggleDHPar;

  // Reset live renderer refs for this grid instance
  _dhSecRenderers = [];
  _dhParCells     = [];

  sec('Daily Metrics','#006461');
  par('Occupancy','#006461');
  row('T / Hotel', function(d){
    var segs=[{o:0,w:d.fitPct,c:'#006461'},{o:d.fitPct,w:d.dynPct,c:'#0891b2'},{o:d.fitPct+d.dynPct,w:d.serPct,c:'#6366f1'},{o:d.to,w:Math.max(0,d.otherPct),c:'#5883ed'}];
    return valH(d.to+'%',d.hotel+'%','#006461')+segBar(segs)+refChips([{k:'stly',lbl:'STLY',v:d.sdlyH+'%'},{k:'ly',lbl:'LY',v:d.lyH+'%'},{k:'fcst',lbl:'Fcst',v:d.fcstH+'%'}]);
  });
  row('Online / Offline', function(d){
    return stackBar([{p:d.onlinePct,c:'#3b82f6'},{p:100-d.onlinePct,c:'#f97316'}])
      +'<div style="display:flex;justify-content:space-between"><span style="font-size:8px;color:#3b82f6">'+d.onlinePct+'% online</span><span style="font-size:8px;color:#f97316">'+(100-d.onlinePct)+'% offline</span></div>';
  });
  par('ADR','#7c3aed');
  row('T / Hotel', function(d){
    var df=d.toAdr-d.adr;
    return valH('$'+d.toAdr,'$'+d.adr,'#7c3aed')+dualBar(d.adrBar,Math.min(95,d.adrBar+12),'#7c3aed')
      +'<div style="display:flex;align-items:center;gap:4px;margin-top:2px"><span style="font-size:8px;font-weight:700;color:'+(df<=0?'#16a34a':'#dc2626')+'">'+(df>=0?'+':'−')+'$'+Math.abs(df)+' diff</span></div>'
      +refChips([{k:'stly',lbl:'STLY',v:'$'+d.sdlyA},{k:'ly',lbl:'LY',v:'$'+d.lyA},{k:'fcst',lbl:'Fcst',v:'$'+d.fcstA}]);
  });
  par('Revenue','#ea580c');
  row('T / Hotel', function(d){
    return valH(d.fR(d.toRev),d.fR(d.hnRev),'#ea580c')+dualBar(d.revBar,Math.min(95,d.revBar+10),'#ea580c')+refChips([{k:'stly',lbl:'STLY',v:d.fR(d.sdlyR)},{k:'ly',lbl:'LY',v:d.fR(d.lyR)},{k:'fcst',lbl:'Fcst',v:d.fR(d.fcstR)}]);
  });
  par('REVPAR','#9333ea');
  row('T / Hotel', function(d){
    return valH('$'+d.revpar,'$'+(d.revpar+22),'#9333ea')+dualBar(Math.round(d.revpar/4),Math.round((d.revpar+22)/4),'#9333ea')+refChips([{k:'stly',lbl:'STLY',v:'$'+d.sdlyRevpar},{k:'ly',lbl:'LY',v:'$'+d.lyRevpar}]);
  });
  par('Pickup','#16a34a');
  row('T / Hotel', function(d){ return valH('+'+d.pickup,'+'+d.hPickup,'#16a34a'); });
  par('Segments (T)','#0891b2');
  row('FIT / Dyn / Series', function(d){
    return stackBar([{p:d.fitPct,c:'#006461'},{p:d.dynPct,c:'#0891b2'},{p:d.serPct,c:'#6366f1'}])
      +'<div style="display:flex;gap:6px;font-size:8px;flex-wrap:wrap;margin-top:1px"><span style="color:#006461">FIT '+d.fitPct+'%</span><span style="color:#0891b2">Dyn '+d.dynPct+'%</span><span style="color:#6366f1">Series '+d.serPct+'%</span></div>';
  });

  sec('More Metrics','#2e65e8');
  par('RN Sold','#2e65e8');
  row('T / Hotel', function(d){
    return valH(d.toRn,d.hnRn,'#2e65e8')+dualBar(Math.round(d.toRn/WV_CAP*100),Math.round(d.hnRn/WV_CAP*100),'#2e65e8')+refChips([{k:'stly',lbl:'STLY',v:d.sdlyRn},{k:'ly',lbl:'LY',v:d.lyRn},{k:'fcst',lbl:'Fcst',v:d.fcstRn}]);
  });
  par('Avg Adults','#2e65e8');    row('T / Hotel', function(d){ return valH(d.avgA,d.hAvgA,'#2e65e8'); });
  par('Avg Children','#d33030');  row('T / Hotel', function(d){ return valH(d.avgC,d.hAvgC,'#d33030'); });
  par('Total Adults','#2e65e8'); row('T / Hotel', function(d){ return valH(d.totAT,d.totAH,'#2e65e8'); });
  par('Total Children','#d33030'); row('T / Hotel', function(d){ return valH(d.totCT,d.totCH,'#d33030'); });
  par('Total Guests','#0369a1'); row('T / Hotel', function(d){ return valH(d.totG,d.hTotG,'#0369a1'); });
  par('Avg LOS','#0891b2');       row('T / Hotel', function(d){ return valH(d.avgLos,d.hLos,'#0891b2'); });
  par('Lead Time','#6366f1');     row('T / Hotel', function(d){ return valH(d.avgLead,d.hLead,'#6366f1'); });
  par('Avail Rooms','#16a34a');   row('Hotel', function(d){ return '<span style="font-size:11px;font-weight:800;color:#16a34a">'+d.availRooms+' rm</span>'; });
  par('Avail Guar.','#ea580c');   row('T', function(d){ return '<span style="font-size:11px;font-weight:800;color:#ea580c">'+d.availGuar+' rm</span>'; });

  sec('Meal Plans','#967EF3');
  var mpDefs=[['All Inclusive','#006461','aiPct'],['Bed & Bkfst','#3b82f6','bbPct'],['Half Board','#967EF3','hbPct'],['Room Only','#f59e0b','roPct']];
  mpDefs.forEach(function(mp){
    par(mp[0],mp[1]);
    var key=mp[2];
    row('Hotel / TO %', function(d){
      var hPct=d[key], toPct2=Math.max(0,Math.round(hPct*d.toPct*0.9));
      return valH(hPct+'%','TO '+toPct2+'%',mp[1])+dualBar(hPct,null,mp[1]);
    });
  });
  par('Summary','#967EF3');
  row('AI / BB / HB / RO', function(d){
    return stackBar([{p:d.aiPct,c:'#006461'},{p:d.bbPct,c:'#3b82f6'},{p:d.hbPct,c:'#967EF3'},{p:d.roPct,c:'#f59e0b'}])
      +'<div style="display:flex;gap:5px;font-size:8px;flex-wrap:wrap"><span style="color:#006461">AI '+d.aiPct+'%</span><span style="color:#3b82f6">BB '+d.bbPct+'%</span><span style="color:#967EF3">HB '+d.hbPct+'%</span><span style="color:#f59e0b">RO '+d.roPct+'%</span></div>';
  });

  sec('Business Mix','#0284c7');
  par('TO / Direct / OTA','#0284c7');
  row('Mix %', function(d){
    return stackBar([{p:d.toMix,c:'#006461'},{p:d.dirMix,c:'#0284c7'},{p:d.otaMix,c:'#D97706'},{p:d.otherMix,c:'#9ca3af'}])
      +'<div style="display:flex;gap:5px;font-size:8px;flex-wrap:wrap"><span style="color:#006461">TO '+d.toMix+'%</span><span style="color:#0284c7">Direct '+d.dirMix+'%</span><span style="color:#D97706">OTA '+d.otaMix+'%</span><span style="color:#9ca3af">Other '+d.otherMix+'%</span></div>';
  });

  sec('Travel Co. Rates','#0f766e');
  var toOps=[['Sunshine Tours','#3b82f6'],['Global Adv.','#967EF3'],['Beach Hols','#0ea5e9'],['City Breaks','#10b981'],['Adventure','#f59e0b']];
  toOps.forEach(function(op,i){
    par(op[0],op[1]);
    row('Rate / Promo', (function(op,i){ return function(d){
      return '<div style="display:flex;align-items:center;justify-content:space-between;gap:4px"><span style="font-size:11px;font-weight:800;color:'+op[1]+'">$'+d.tcRates[i]+'</span>'+promoBadge(d)+'</div>';
    };})(op,i));
  });
  par('Base Rate','#9333ea');
  row('Rate', function(d){ return '<span style="font-size:11px;font-weight:800;color:#9333ea">$'+d.baseRate+'</span>'; });

  // ── Day column header component factory ───────────────────────────────────
  function makeDayHeader(dv, isActive, isToday, isLocked, dba, evts) {
    var dm = dv.month, dd = dv.day;
    var bg       = isLocked ? '#374151' : isActive ? '#006461' : isToday ? '#125756' : '#1a5e5b';
    var topBorder= isActive ? '3px solid rgba(255,255,255,0.5)' : isToday ? '3px solid rgba(255,255,255,0.3)' : isLocked ? '3px solid #dc2626' : '3px solid transparent';
    var dayClr   = isLocked ? '#fca5a5' : '#fff';
    var subClr   = isLocked ? 'rgba(252,165,165,0.85)' : 'rgba(255,255,255,0.75)';
    var dbaStr   = dba === 0 ? 'Today' : dba > 0 ? dba + ' DBA' : '';
    function H() {}
    H.prototype.init = function(p) {
      this.gui = document.createElement('div');
      this.gui.style.cssText = 'background:'+bg+';width:100%;height:100%;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:4px 6px;box-sizing:border-box;gap:2px;border-top:'+topBorder+';border-right:1px solid rgba(255,255,255,0.12);';
      this.gui.innerHTML =
        '<div style="font-weight:700;font-size:12px;color:'+dayClr+'">'+(isLocked?'🔒 ':'')+DOW_SHORT[new Date(2026,dm-1,dd).getDay()]+' '+dd+'</div>'
        +'<div style="font-size:10px;color:'+subClr+';display:flex;align-items:center;gap:4px">'
        +'<span>'+MNAMES_S[dm]+'</span>'
        +(dbaStr?'<span style="background:rgba(255,255,255,0.2);border-radius:3px;padding:0 4px;font-size:9px;color:#fff">'+dbaStr+'</span>':'')
        +(evts?'<span style="width:6px;height:6px;border-radius:50%;background:rgba(255,255,255,0.9);display:inline-block"></span>':'')
        +'</div>';
    };
    H.prototype.getGui = function() { return this.gui; };
    H.prototype.destroy = function() {};
    return H;
  }

  // ── Full-width section-header renderer — HR demo-inspired ────────────────
  function SecRenderer() { this._iconEl = null; this._secKey = null; }
  SecRenderer.prototype.init = function(p) {
    var self = this;
    var r    = p.data;
    this._secKey = r._secKey;

    var isCollapsed = !!_dhCollapsed[r._secKey];
    var clr = r._clr;

    this.gui = document.createElement('div');
    this.gui.style.cssText =
      'display:flex;align-items:center;gap:9px;width:100%;height:100%;box-sizing:border-box;'
      + 'cursor:pointer;user-select:none;padding:0 14px;'
      + 'background:#f8f9fd;'
      + 'border-left:3px solid ' + clr + ';'
      + 'border-top:1px solid #dde1e2;border-bottom:1px solid #dde1e2;';

    // Rounded icon badge (HR demo style)
    var iconWrap = document.createElement('span');
    this._iconEl = iconWrap;
    iconWrap.style.cssText =
      'display:inline-flex;align-items:center;justify-content:center;flex-shrink:0;'
      + 'width:20px;height:20px;border-radius:5px;'
      + 'background:' + clr + '22;color:' + clr + ';'
      + 'box-shadow:0 1px 3px ' + clr + '33;'
      + 'transform:rotate(' + (isCollapsed ? '-90deg' : '0deg') + ');'
      + 'transition:transform .2s ease;';
    iconWrap.innerHTML = '<span class="material-icons" style="font-size:12px">expand_more</span>';

    var label = document.createElement('span');
    label.style.cssText =
      'font-size:10.5px;font-weight:700;text-transform:uppercase;letter-spacing:.65px;color:#1e2d3a';
    label.textContent = r._lbl;

    this.gui.appendChild(iconWrap);
    this.gui.appendChild(label);
    this.gui.addEventListener('click', function() { toggleDHSection(r._secKey); });

    // Register for Open All / Close All sync
    _dhSecRenderers.push(self);
  };
  SecRenderer.prototype._syncChevron = function() {
    if (this._iconEl) {
      this._iconEl.style.transform = 'rotate(' + (_dhCollapsed[this._secKey] ? '-90deg' : '0deg') + ')';
    }
  };
  SecRenderer.prototype.getGui    = function() { return this.gui; };
  SecRenderer.prototype.destroy   = function() {
    var i = _dhSecRenderers.indexOf(this);
    if (i !== -1) _dhSecRenderers.splice(i, 1);
  };

  // ── Column defs ───────────────────────────────────────────────────────────
  var colDefs = [];

  // Metric label (pinned left)
  colDefs.push({
    field: '_lbl',
    headerName: 'Metric',
    pinned: 'left',
    lockPinned: true,
    width: 170,
    suppressMovable: true,
    resizable: false,
    cellRenderer: function(p) {
      var r = p.data;
      if (r._type === 'par') {
        var el = document.createElement('div');
        var isCollapsed = !!_dhCollapsed[r._parKey];
        var clr = r._clr;
        el.style.cssText = 'display:flex;align-items:center;gap:7px;cursor:pointer;width:100%;height:100%;box-sizing:border-box;user-select:none;padding:0 10px 0 16px;';

        // Small rounded icon badge — HR demo sub-group style
        var iconWrap = document.createElement('span');
        iconWrap.style.cssText =
          'display:inline-flex;align-items:center;justify-content:center;flex-shrink:0;'
          + 'width:16px;height:16px;border-radius:4px;'
          + 'background:' + clr + '18;color:' + clr + ';'
          + 'transform:rotate(' + (isCollapsed ? '-90deg' : '0deg') + ');'
          + 'transition:transform .2s ease;';
        iconWrap.innerHTML = '<span class="material-icons" style="font-size:11px">expand_more</span>';

        var label = document.createElement('span');
        label.style.cssText = 'font-size:10px;font-weight:600;color:#374151;letter-spacing:.15px';
        label.textContent = r._lbl;

        el.appendChild(iconWrap);
        el.appendChild(label);
        el.addEventListener('click', function() { toggleDHPar(r._parKey); });

        // Register for Open All / Close All sync
        var ref = { _parKey: r._parKey, _iconEl: iconWrap };
        _dhParCells.push(ref);
        return el;
      }
      return '<span style="font-size:9.5px;font-weight:500;color:#536271;padding-left:24px">'+r._lbl+'</span>';
    },
    cellStyle: function(p) {
      var r = p.data, idx = p.node.rowIndex;
      if (r._type === 'par') return { background:'#f8f9fd', display:'flex', alignItems:'center', padding:'4px 10px 4px 0', borderBottom:'1px solid #dde1e2', borderRight:'2px solid #dde1e2' };
      return { background:'#fff', padding:'6px 12px 6px 28px', borderRight:'2px solid #dde1e2' };
    },
  });

  // One column per day
  days.forEach(function(dv, di) {
    var dm = dv.month, dd = dv.day;
    var isToday  = dm===3 && dd===9;
    var isActive = dm===activeMonth && dd===activeDay;
    var isLocked = LOCKED_DAYS.has(dm+'-'+dd);
    var dt  = new Date(2026, dm-1, dd);
    var dba = Math.round((dt - TODAY_WV) / 86400000);
    var evts = (typeof CAL_EVENTS!=='undefined' && CAL_EVENTS[dm+'-'+dd]) ? CAL_EVENTS[dm+'-'+dd] : null;

    colDefs.push({
      field: 'day'+di,
      width: 148,
      suppressMovable: true,
      resizable: false,
      headerComponent: makeDayHeader(dv, isActive, isToday, isLocked, dba, evts),
      cellRenderer: function(p) {
        var r = p.data;
        if (r._type !== 'row') return '';
        var d = dayData[di];
        if (LOCKED_DAYS.has(d.dm+'-'+d.dd)) return '<span style="color:#9ca3af;font-size:12px">—</span>';
        return r._fn(d);
      },
      cellStyle: function(p) {
        var r = p.data, idx = p.node.rowIndex;
        if (r._type === 'par') return { background:'#f8f9fd', padding:'4px 8px', borderBottom:'1px solid #dde1e2', borderRight:'1px solid #dde1e2' };
        return { background:'#fff', padding:'6px 10px', borderRight:'1px solid #dde1e2' };
      },
    });
  });

  // ── Apply custom metric order if set ─────────────────────────────────────
  if (_dhMetricOrder && _dhMetricOrder.length) {
    ROWS = reorderDHRows(ROWS, _dhMetricOrder);
  }

  // ── Expose rows to module level (for Open All / Close All) ──────────────
  _dhAllRows = ROWS;

  // Default collapse state: sections open, parent-rows closed (only set if not yet toggled by user)
  ROWS.forEach(function(r) {
    if (r.type === 'par' && !Object.prototype.hasOwnProperty.call(_dhCollapsed, r.parKey)) {
      _dhCollapsed[r.parKey] = true;
    }
  });

  // ── Build row data (respecting accordion collapse state) ─────────────────
  var rowData = getDHVisibleRowData();

  // ── Create grid ───────────────────────────────────────────────────────────
  _dailyHGridApi = AG.createGrid(wrapper, {
    columnDefs: colDefs,
    rowData: rowData,
    headerHeight: 54,
    domLayout: 'autoHeight',
    suppressHorizontalScroll: false,
    alwaysShowHorizontalScroll: true,
    suppressCellFocus: true,
    suppressRowClickSelection: true,
    getRowHeight: function(p) {
      if (p.data._type === 'sec') return 36;
      if (p.data._type === 'par') return 30;
      // 'row' type: let autoHeight per column determine height
    },
    isFullWidthRow: function(p) {
      return p.rowNode.data._type === 'sec';
    },
    fullWidthCellRenderer: SecRenderer,
    defaultColDef: {
      sortable: false,
      resizable: false,
      autoHeight: true,
    },
    getRowStyle: function(p) {
      var t = p.data._type, idx = p.node.rowIndex;
      if (t === 'sec') return { background:'#f8f9fd', borderTop:'1px solid #dde1e2', borderBottom:'1px solid #dde1e2' };
      if (t === 'par') return { background:'#f8f9fd' };
      if (idx%2===0) return { background:'#fff' };
      return { background:'#fff' };
    },
  });
}

/* ── Close Out Report AG Grid ────────────────────────────────────────────── */
var _coReportGridApi = null;

function initCoReportGrid(days, containerEl) {
  var AG = _realAgGrid;
  if (!AG || typeof AG.createGrid !== 'function') {
    containerEl.innerHTML = buildCoReportView(days);
    return;
  }

  if (_coReportGridApi) { try { _coReportGridApi.destroy(); } catch(e){} _coReportGridApi = null; }
  containerEl.innerHTML = '';
  containerEl.style.padding = '0';

  var wrapper = document.createElement('div');
  wrapper.className = 'ag-theme-quartz co-report-ag-wrap';
  containerEl.appendChild(wrapper);

  var BMAP   = {ai:'All Inclusive', bb:'B&B', hb:'Half Board', ro:'Room Only', fb:'Full Board'};
  var MNAMES = ['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  var DOW    = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
  var TODAY  = new Date(2026, 2, 9);
  var STRAT_COLORS = ['#dc2626','#b45309','#7c3aed','#0891b2','#16a34a'];

  // ── determine max strategies ────────────────────────────────────────────
  var maxStrats = 0;
  days.forEach(function(dv) {
    var r = PARTIAL_CLOSURES[dv.month + '-' + dv.day];
    if (r && r.length > maxStrats) maxStrats = r.length;
  });
  if (maxStrats === 0) maxStrats = 1;

  // ── chip renderer ───────────────────────────────────────────────────────
  function chip(label, color) {
    return '<span style="display:inline-flex;align-items:center;font-size:9px;font-weight:600;padding:1px 6px;border-radius:3px;background:'+color+'18;color:'+color+';border:1px solid '+color+'44;white-space:nowrap;margin:1px 2px">'+label+'</span>';
  }
  var lockSvg = '<svg viewBox="0 0 10 12" fill="none" stroke="currentColor" stroke-width="1.6" width="9" height="10" style="flex-shrink:0"><rect x="1" y="5" width="8" height="7" rx="1"/><path d="M3 5V3.5a2 2 0 0 1 4 0V5"/></svg>';
  var openSvg = '<svg viewBox="0 0 14 14" fill="none" stroke="#15803d" stroke-width="1.8" width="10" height="10"><path d="M2 7l4 4 6-6"/></svg>';

  // ── row data ────────────────────────────────────────────────────────────
  var rowData = days.map(function(dv) {
    var dm = dv.month, dd = dv.day;
    var key = dm + '-' + dd;
    var isFullyLocked = LOCKED_DAYS.has(key);
    var rules = PARTIAL_CLOSURES[key] || [];
    var dt  = new Date(2026, dm - 1, dd);
    var dba = Math.round((dt - TODAY) / 86400000);
    var row = {
      _date:     MNAMES[dm] + ' ' + dd,
      _dow:      DOW[dt.getDay()],
      _dba:      dba === 0 ? 'Today' : dba > 0 ? dba + ' DBA' : '',
      _isToday:  dm === 3 && dd === 9,
      _isLocked: isFullyLocked,
      _evts:     (typeof CAL_EVENTS !== 'undefined' && CAL_EVENTS[key]) ? CAL_EVENTS[key] : null,
    };
    for (var si = 0; si < maxStrats; si++) {
      var rule = isFullyLocked ? null : rules[si];
      row['s' + si + '_hasRule']  = !!rule;
      row['s' + si + '_isLocked'] = isFullyLocked && si === 0; // lock shown once in first strat
      row['s' + si + '_tos']    = rule ? rule.tos    : [];
      row['s' + si + '_rooms']  = rule ? rule.roomTypes : [];
      row['s' + si + '_boards'] = rule ? rule.boards  : [];
    }
    return row;
  });

  // ── custom colored group header ─────────────────────────────────────────
  function makeStratHeader(color) {
    function H() {}
    H.prototype.init = function(p) {
      this.gui = document.createElement('div');
      this.gui.style.cssText = 'background:' + color + ';color:#fff;font-size:10px;font-weight:700;letter-spacing:.3px;display:flex;align-items:center;padding:0 10px;width:100%;height:100%';
      this.gui.textContent = p.displayName;
    };
    H.prototype.getGui = function() { return this.gui; };
    H.prototype.destroy = function() {};
    return H;
  }

  // ── column defs ─────────────────────────────────────────────────────────
  var colDefs = [
    {
      headerName: 'Stay Date', pinned: 'left', lockPinned: true, width: 118, suppressSizeToFit: true,
      sortable: false, resizable: false,
      cellStyle: { padding: '4px 8px', lineHeight: '1.4', display: 'flex', alignItems: 'center' },
      cellRenderer: function(p) {
        var d = p.data;
        return '<div style="line-height:1.5">'
          + '<div style="font-weight:800;font-size:11px;color:' + (d._isLocked ? '#dc2626' : '#111827') + '">'
          + (d._isLocked ? lockSvg + ' ' : '') + d._date + '</div>'
          + '<div style="font-size:9px;color:#6b7280;display:flex;align-items:center;gap:4px;margin-top:1px">'
          + '<span>' + d._dow + '</span>'
          + (d._dba ? '<span style="color:#006461;font-weight:700">' + d._dba + '</span>' : '')
          + (d._evts ? '<span title="' + d._evts.map(function(e){return e.name;}).join(', ') + '" style="width:7px;height:7px;border-radius:2px;background:#C4FF45;display:inline-block;flex-shrink:0"></span>' : '')
          + '</div>'
          + (d._isToday ? '<div style="width:20px;height:2px;background:#006461;border-radius:1px;margin-top:2px"></div>' : '')
          + '</div>';
      }
    }
  ];

  // Strategy column groups
  for (var si = 0; si < maxStrats; si++) {
    (function(idx) {
      var clr  = STRAT_COLORS[idx % STRAT_COLORS.length];
      var pfx  = 's' + idx;

      function opCell(p) {
        var d = p.data;
        if (d[pfx + '_isLocked']) {
          return '<span style="color:#dc2626;font-size:10px;font-weight:700;display:flex;align-items:center;gap:4px">' + lockSvg + ' Full Day Closed</span>';
        }
        if (!d[pfx + '_hasRule']) {
          return idx === 0 && !d._isLocked
            ? '<span style="color:#15803d;display:flex;align-items:center;gap:4px">' + openSvg + '<span style="font-size:10px">Open</span></span>'
            : '';
        }
        var tos = d[pfx + '_tos'];
        return tos.length
          ? tos.map(function(n){ return chip(n, TO_COLORS_MAP[n] || '#dc2626'); }).join('')
          : '<span style="font-size:9px;color:#9ca3af;font-style:italic">All operators</span>';
      }

      function rtCell(p) {
        var d = p.data;
        if (!d[pfx + '_hasRule'] || d[pfx + '_isLocked']) return '';
        var rooms = d[pfx + '_rooms'];
        return rooms.length
          ? rooms.map(function(n){ return chip(n, RT_NAME_COLORS[n] || '#b45309'); }).join('')
          : '<span style="font-size:9px;color:#9ca3af;font-style:italic">All rooms</span>';
      }

      function bdCell(p) {
        var d = p.data;
        if (!d[pfx + '_hasRule'] || d[pfx + '_isLocked']) return '';
        var bds = d[pfx + '_boards'];
        return bds.length
          ? bds.map(function(b){ return chip(BMAP[b] || b, '#7c3aed'); }).join('')
          : '<span style="font-size:9px;color:#9ca3af;font-style:italic">All plans</span>';
      }

      colDefs.push({
        headerName: 'Strategy ' + (idx + 1),
        headerGroupComponent: makeStratHeader(clr),
        children: [
          { headerName: 'Operator',  field: pfx + '_tos',    width: 150, sortable: false, resizable: true, cellRenderer: opCell },
          { headerName: 'Room Type', field: pfx + '_rooms',  width: 130, sortable: false, resizable: true, cellRenderer: rtCell },
          { headerName: 'Meal Plan', field: pfx + '_boards', width: 120, sortable: false, resizable: true, cellRenderer: bdCell },
        ]
      });
    })(si);
  }

  // Full Day column
  colDefs.push({
    headerName: 'Full Day', width: 90, sortable: false, resizable: false,
    cellStyle: { display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '11px' },
    cellRenderer: function(p) {
      return p.data._isLocked
        ? '<span style="color:#dc2626;font-weight:700;display:flex;align-items:center;gap:4px">' + lockSvg + ' Closed</span>'
        : '<span style="color:#15803d;display:flex;align-items:center;gap:4px">' + openSvg + '<span style="font-size:10px">Open</span></span>';
    }
  });

  // ── create grid ─────────────────────────────────────────────────────────
  _coReportGridApi = AG.createGrid(wrapper, {
    columnDefs: colDefs,
    rowData: rowData,
    rowHeight: 44,
    headerHeight: 28,
    groupHeaderHeight: 28,
    domLayout: 'autoHeight',
    suppressHorizontalScroll: false,
    alwaysShowHorizontalScroll: true,
    defaultColDef: {
      sortable: false,
      resizable: true,
      cellStyle: { fontSize: '11px', display: 'flex', alignItems: 'center', flexWrap: 'wrap', padding: '4px 8px' },
    },
    getRowStyle: function(p) {
      if (p.data._isLocked) return { background: '#fef2f2' };
      if (p.data._isToday)  return { background: 'rgba(0,100,97,0.06)' };
      if (p.node.rowIndex % 2) return { background: '#fafafa' };
    },
  });
}

/* ── Daily Revenue AG Grid ───────────────────────────────────────────────── */
var _dailyRevGridApi = null;

function initDailyRevGrid(days, containerEl) {
  _drLastInitArgs = { days: days, container: containerEl };
  var AG = _realAgGrid;
  if (!AG || typeof AG.createGrid !== 'function') {
    containerEl.innerHTML = buildReportView(days);
    return;
  }

  // destroy previous
  if (_dailyRevGridApi) { try { _dailyRevGridApi.destroy(); } catch(e){} _dailyRevGridApi = null; }
  containerEl.innerHTML = '';
  containerEl.style.padding = '0';

  var wrapper = document.createElement('div');
  wrapper.className = 'ag-theme-quartz daily-rev-ag-wrap';
  containerEl.appendChild(wrapper);

  // ── per-day data ──────────────────────────────────────────────────────────
  var WV_CAP = 250;
  var MNAMES = ['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  var DOW    = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
  var TODAY  = new Date(2026,2,9);

  var rowData = days.map(function(dv){
    var dm=dv.month, d2=dv.day;
    var hh=getOccupancy(dm,d2); var hotel=hh.hotel, to=hh.to;
    var adr=150+Math.abs((dm*47+d2*31)%130);
    var v=Math.abs((dm*127+d2*53+dm*d2*7+d2*d2*3))%100;
    var toAdr=Math.max(80,adr-20-Math.abs((dm*3+d2*7)%15));
    var toRn=Math.round(WV_CAP*to/100), hnRn=Math.round(WV_CAP*hotel/100);
    var toRev=Math.floor(toRn*toAdr), hnRev=Math.floor(hnRn*adr);
    var otherPct=Math.max(0,hotel-to), freePct=Math.max(0,100-hotel);
    var toRms=toRn, otherRms=Math.round(WV_CAP*otherPct/100);
    var freeRms=WV_CAP-toRms-otherRms;
    var fitPct=Math.round(to*0.45), dynPct=Math.round(to*0.35), serPct=to-fitPct-dynPct;
    var revpar=Math.max(50,(adr+80)-30-Math.abs((dm*5+d2*3)%20));
    var hotelRevpar=adr+80;
    var pickup=Math.max(0,Math.floor((v%25+5)*to/Math.max(1,hotel)));
    var hotelPickup=Math.floor(v%25+5);
    var sdlyH=Math.max(5,hotel-9), lyH=Math.max(5,hotel-6), fcstH=Math.min(100,hotel+4);
    var sdlyA=adr-8, lyA=adr-4, fcstA=adr+6;
    var sdlyR=Math.floor(hnRev*0.9), lyR=Math.floor(hnRev*0.95), fcstR=Math.floor(hnRev*1.06);
    var sdlyRevpar=Math.max(40,revpar-8), lyRevpar=Math.max(40,revpar-4);
    var toMix=28+Math.abs((dm*7+d2*5)%25), dirMix=30+Math.abs((dm*5+d2*9)%20);
    var otaMix=20+Math.abs((dm*9+d2*3)%18), otherMix=Math.max(0,100-toMix-dirMix-otaMix);
    var aiPct=Math.max(45,Math.min(68,55+(dm*7+d2*3)%14));
    var bbPct=Math.max(14,Math.min(28,20+(dm*11+d2*5)%11));
    var hbPct=Math.max(6,Math.min(16,10+(dm*5+d2*7)%9));
    var roPct=100-aiPct-bbPct-hbPct;
    var tcRates=[0,1,2,3,4].map(function(i){return adr-15+Math.abs((dm*(i+3)+d2*(i+5))%50);});
    var baseRate=adr+8;
    var dt=new Date(2026,dm-1,d2);
    var dow=DOW[dt.getDay()];
    var dba=Math.round((dt-TODAY)/86400000);
    var dbaStr=dba===0?'Today':dba>0?dba+' DBA':'';
    function fR(x){return x>=1000000?'$'+(x/1000000).toFixed(1)+'M':'$'+Math.round(x/1000)+'k';}
    var onPct=Math.max(30,Math.min(80,45+Math.abs((dm*13+d2*7)%35)));
    return {
      _date:MNAMES[dm]+' '+d2, _dow:dow, _dba:dbaStr, _isToday:dm===3&&d2===9,
      occ_t:to+'%',    occ_h:hotel+'%',  occ_stly:sdlyH+'%', occ_ly:lyH+'%', occ_fcst:fcstH+'%',
      adr_t:'$'+toAdr, adr_h:'$'+adr,   adr_diff:(toAdr-adr>=0?'+':'')+(toAdr-adr),
      adr_stly:'$'+sdlyA, adr_ly:'$'+lyA, adr_fcst:'$'+fcstA,
      rev_t:fR(toRev),  rev_h:fR(hnRev), rev_stly:fR(sdlyR), rev_ly:fR(lyR), rev_fcst:fR(fcstR),
      rp_t:'$'+revpar,  rp_h:'$'+hotelRevpar, rp_stly:'$'+sdlyRevpar, rp_ly:'$'+lyRevpar,
      pk_t:'+'+pickup,  pk_h:'+'+hotelPickup,
      td_rms:toRms+' rm',    td_pct:to+'%',
      os_rms:otherRms+' rm', os_pct:otherPct+'%',
      rem_rms:freeRms+' rm', rem_pct:Math.max(0,Math.round(freePct))+'%',
      on_on:onPct+'%', on_off:(100-onPct)+'%',
      fit_rms:Math.round(250*fitPct/100)+' rm', fit_pct:fitPct+'%',
      dyn_rms:Math.round(250*dynPct/100)+' rm', dyn_pct:dynPct+'%',
      ser_rms:Math.round(250*serPct/100)+' rm', ser_pct:serPct+'%',
      biz_to:toMix+'%', biz_dir:dirMix+'%', biz_ota:otaMix+'%', biz_oth:otherMix+'%',
      mp_ai_h:Math.round(hnRn*aiPct/100), mp_ai_t:Math.round(toRn*aiPct/100), mp_ai_pct:aiPct+'%',
      mp_bb_h:Math.round(hnRn*bbPct/100), mp_bb_t:Math.round(toRn*bbPct/100), mp_bb_pct:bbPct+'%',
      mp_hb_h:Math.round(hnRn*hbPct/100), mp_hb_t:Math.round(toRn*hbPct/100), mp_hb_pct:hbPct+'%',
      mp_ro_h:Math.round(hnRn*roPct/100), mp_ro_t:Math.round(toRn*roPct/100), mp_ro_pct:roPct+'%',
      tc_0:'$'+tcRates[0], tc_1:'$'+tcRates[1], tc_2:'$'+tcRates[2],
      tc_3:'$'+tcRates[3], tc_4:'$'+tcRates[4], tc_base:'$'+baseRate,
    };
  });

  // ── helpers ───────────────────────────────────────────────────────────────
  function cs(color, bold){ return { color:color, fontWeight:bold?'700':'400', display:'flex', alignItems:'center' }; }
  function csFn(fn){ return function(p){ return fn(p); }; }
  var BASE_COL = {
    sortable: false, resizable: true, suppressMovable: false,
    cellStyle: { fontSize:'11px', display:'flex', alignItems:'center' },
    headerComponentParams: {}
  };

  // ── column defs ───────────────────────────────────────────────────────────
  // Group colDefs keyed by headerName for reorder support
  var drGroupColDefs = {
    'Daily Metrics':
    { headerName:'Daily Metrics', headerClass:'drg-top drg-daily', openByDefault:true, children:[
      { headerName:'Occupancy', children:[
        {field:'occ_t',    headerName:'T',     width:65, cellStyle:cs('#006461',true)},
        {field:'occ_h',    headerName:'Hotel', width:70, cellStyle:cs('#374151')},
        {field:'occ_stly', headerName:'STLY',  width:68, cellStyle:cs('#9ca3af')},
        {field:'occ_ly',   headerName:'LY',    width:65, cellStyle:cs('#9ca3af')},
        {field:'occ_fcst', headerName:'Fcst',  width:68, cellStyle:cs('#f59e0b')},
      ]},
      { headerName:'ADR', children:[
        {field:'adr_t',    headerName:'T',     width:70, cellStyle:cs('#7c3aed',true)},
        {field:'adr_h',    headerName:'Hotel', width:70, cellStyle:cs('#374151')},
        {field:'adr_diff', headerName:'Diff',  width:65, cellStyle:csFn(function(p){return {color:p.value&&p.value.charAt(0)==='-'?'#dc2626':'#16a34a',display:'flex',alignItems:'center'};})},
        {field:'adr_stly', headerName:'STLY',  width:70, cellStyle:cs('#9ca3af')},
        {field:'adr_ly',   headerName:'LY',    width:70, cellStyle:cs('#9ca3af')},
        {field:'adr_fcst', headerName:'Fcst',  width:70, cellStyle:cs('#f59e0b')},
      ]},
      { headerName:'Revenue', children:[
        {field:'rev_t',    headerName:'T',     width:80, cellStyle:cs('#ea580c',true)},
        {field:'rev_h',    headerName:'Hotel', width:80, cellStyle:cs('#374151')},
        {field:'rev_stly', headerName:'STLY',  width:80, cellStyle:cs('#9ca3af')},
        {field:'rev_ly',   headerName:'LY',    width:80, cellStyle:cs('#9ca3af')},
        {field:'rev_fcst', headerName:'Fcst',  width:80, cellStyle:cs('#f59e0b')},
      ]},
      { headerName:'RevPAR', children:[
        {field:'rp_t',    headerName:'T',     width:70, cellStyle:cs('#9333ea',true)},
        {field:'rp_h',    headerName:'Hotel', width:70, cellStyle:cs('#374151')},
        {field:'rp_stly', headerName:'STLY',  width:70, cellStyle:cs('#9ca3af')},
        {field:'rp_ly',   headerName:'LY',    width:70, cellStyle:cs('#9ca3af')},
      ]},
      { headerName:'Pickup', children:[
        {field:'pk_t', headerName:'T',     width:65, cellStyle:cs('#16a34a',true)},
        {field:'pk_h', headerName:'Hotel', width:70, cellStyle:cs('#374151')},
      ]},
    ]},
    'Room Avail.':
    { headerName:'Room Avail.', headerClass:'drg-top drg-avail', openByDefault:true, children:[
      { headerName:'T Dist. Hubs', children:[
        {field:'td_rms', headerName:'Rooms', width:85, cellStyle:cs('#006461',true)},
        {field:'td_pct', headerName:'%',     width:60, cellStyle:cs('#006461')},
      ]},
      { headerName:'Other Segs', children:[
        {field:'os_rms', headerName:'Rooms', width:85, cellStyle:cs('#5883ed')},
        {field:'os_pct', headerName:'%',     width:60, cellStyle:cs('#5883ed')},
      ]},
      { headerName:'Remaining', children:[
        {field:'rem_rms', headerName:'Rooms', width:85, cellStyle:cs('#16a34a',true)},
        {field:'rem_pct', headerName:'%',     width:60, cellStyle:cs('#16a34a')},
      ]},
      { headerName:'Online/Offline', children:[
        {field:'on_on',  headerName:'Online',  width:75, cellStyle:cs('#3b82f6')},
        {field:'on_off', headerName:'Offline', width:75, cellStyle:cs('#f97316')},
      ]},
    ]},
    'Segments (T)':
    { headerName:'Segments (T)', headerClass:'drg-top drg-segs', openByDefault:true, children:[
      { headerName:'Static FIT', children:[
        {field:'fit_rms', headerName:'Rooms', width:85, cellStyle:cs('#006461',true)},
        {field:'fit_pct', headerName:'%',     width:60},
      ]},
      { headerName:'TO Dynamic', children:[
        {field:'dyn_rms', headerName:'Rooms', width:85, cellStyle:cs('#0891b2',true)},
        {field:'dyn_pct', headerName:'%',     width:60},
      ]},
      { headerName:'Tour Series', children:[
        {field:'ser_rms', headerName:'Rooms', width:85, cellStyle:cs('#6366f1',true)},
        {field:'ser_pct', headerName:'%',     width:60},
      ]},
    ]},
    'Business Mix':
    { headerName:'Business Mix', headerClass:'drg-top drg-biz', openByDefault:true, children:[
      {field:'biz_to',  headerName:'TO',     width:75, cellStyle:cs('#006461',true)},
      {field:'biz_dir', headerName:'Direct', width:85, cellStyle:cs('#0284c7',true)},
      {field:'biz_ota', headerName:'OTA',    width:75, cellStyle:cs('#D97706',true)},
      {field:'biz_oth', headerName:'Other',  width:75, cellStyle:cs('#9ca3af')},
    ]},
    'Meal Plans':
    { headerName:'Meal Plans', headerClass:'drg-top drg-meals', openByDefault:false, children:[
      { headerName:'AI', children:[
        {field:'mp_ai_h',   headerName:'Hotel', width:72, cellStyle:cs('#374151')},
        {field:'mp_ai_t',   headerName:'T',     width:65, cellStyle:cs('#006461',true)},
        {field:'mp_ai_pct', headerName:'Occ',   width:65},
      ]},
      { headerName:'BB', children:[
        {field:'mp_bb_h',   headerName:'Hotel', width:72, cellStyle:cs('#374151')},
        {field:'mp_bb_t',   headerName:'T',     width:65, cellStyle:cs('#3b82f6',true)},
        {field:'mp_bb_pct', headerName:'Occ',   width:65},
      ]},
      { headerName:'HB', children:[
        {field:'mp_hb_h',   headerName:'Hotel', width:72, cellStyle:cs('#374151')},
        {field:'mp_hb_t',   headerName:'T',     width:65, cellStyle:cs('#967EF3',true)},
        {field:'mp_hb_pct', headerName:'Occ',   width:65},
      ]},
      { headerName:'RO', children:[
        {field:'mp_ro_h',   headerName:'Hotel', width:72, cellStyle:cs('#374151')},
        {field:'mp_ro_t',   headerName:'T',     width:65, cellStyle:cs('#f59e0b',true)},
        {field:'mp_ro_pct', headerName:'Occ',   width:65},
      ]},
    ]},
    'Travel Co. Rates':
    { headerName:'Travel Co. Rates', headerClass:'drg-top drg-tc', openByDefault:true, children:[
      {field:'tc_0',    headerName:'Sunshine',    width:92, cellStyle:cs('#3b82f6',true)},
      {field:'tc_1',    headerName:'Global Adv.', width:105, cellStyle:cs('#967EF3',true)},
      {field:'tc_2',    headerName:'Beach Hols',  width:100, cellStyle:cs('#0ea5e9',true)},
      {field:'tc_3',    headerName:'City Breaks',  width:100, cellStyle:cs('#10b981',true)},
      {field:'tc_4',    headerName:'Adventure',   width:92, cellStyle:cs('#f59e0b',true)},
      {field:'tc_base', headerName:'Base Rate',   width:92, cellStyle:cs('#9333ea',true)},
    ]},
  };

  // Build ordered colDefs (pinned date + groups in custom or default order)
  var groupOrder = (_drColOrder && _drColOrder.length) ? _drColOrder : DR_GROUPS_DEF.map(function(g){return g.key;});
  var colDefs = [{
    headerName:'Stay Date', pinned:'left', lockPinned:true, width:140, suppressSizeToFit:true,
    cellRenderer: function(p){
      var d=p.data;
      return '<div style="line-height:1.4;padding:2px 0">'
        +'<div style="font-weight:800;font-size:11px;color:#111827">'+d._date+'</div>'
        +'<div style="font-size:9px;color:#6b7280;display:flex;align-items:center;gap:4px">'
        +'<span>'+d._dow+'</span>'
        +(d._dba?'<span style="color:#006461;font-weight:700">'+d._dba+'</span>':'')
        +'</div>'
        +(d._isToday?'<div style="width:24px;height:2px;background:#006461;border-radius:1px;margin-top:2px"></div>':'')
        +'</div>';
    }
  }];
  groupOrder.forEach(function(k){ if(drGroupColDefs[k]) colDefs.push(drGroupColDefs[k]); });

  // ── create grid ───────────────────────────────────────────────────────────
  _dailyRevGridApi = AG.createGrid(wrapper, {
    columnDefs: colDefs,
    rowData: rowData,
    rowHeight: 42,
    headerHeight: 26,
    groupHeaderHeight: 26,
    domLayout: 'autoHeight',
    suppressHorizontalScroll: false,
    alwaysShowHorizontalScroll: true,
    defaultColDef: Object.assign({}, BASE_COL),
    getRowStyle: function(p){
      if(p.data._isToday) return {background:'rgba(0,100,97,0.06)',fontWeight:'700'};
      if(p.node.rowIndex%2) return {background:'#fafafa'};
    },
    onGridReady: function(e){ e.api.sizeColumnsToFit && false; /* keep explicit widths */ },
  });
}

// ── Report View ──────────────────────────────────────────────────────────────
function buildReportView(days) {
  var WV_CAP2 = 250;
  var collapsed = window._rptCollapsed || {};

  // ── Per-day data helper ───────────────────────────────────────────────────
  function dd(dm, dd2) {
    var hh = getOccupancy(dm, dd2); var hotel=hh.hotel, to=hh.to;
    var adr  = 150+Math.abs((dm*47+dd2*31)%130);
    var v    = Math.abs((dm*127+dd2*53+dm*dd2*7+dd2*dd2*3))%100;
    var toAdr= Math.max(80,adr-20-Math.abs((dm*3+dd2*7)%15));
    var toRn = Math.round(WV_CAP2*to/100);
    var hnRn = Math.round(WV_CAP2*hotel/100);
    var toRev= Math.floor(toRn*toAdr);
    var hnRev= Math.floor(hnRn*adr);
    var otherPct=Math.max(0,hotel-to), freePct=100-hotel;
    var toRms=Math.round(WV_CAP2*to/100), otherRms=Math.round(WV_CAP2*otherPct/100);
    var freeRms=WV_CAP2-toRms-otherRms;
    var fitPct=Math.round(to*0.45),dynPct=Math.round(to*0.35),serPct=to-fitPct-dynPct;
    var onlinePct=Math.max(30,Math.min(80,45+Math.abs((dm*13+dd2*7)%35)));
    var adrBar=Math.min(95,40+Math.abs((dm*11+dd2*19)%55));
    var revBar=Math.min(95,35+Math.abs((dm*17+dd2*13)%60));
    var revpar=Math.max(50,(adr+80)-30-Math.abs((dm*5+dd2*3)%20));
    var hotelRevpar=adr+80;
    var pickup=Math.max(0,Math.floor((v%25+5)*to/Math.max(1,hotel)));
    var hotelPickup=Math.floor(v%25+5);
    var avgA=(1.8+v%3*0.1).toFixed(1), avgC=(0.3+v%2*0.1).toFixed(1);
    var hAvgA=(parseFloat(avgA)+0.3).toFixed(1), hAvgC=(parseFloat(avgC)+0.1).toFixed(1);
    var totG=Math.round(toRn*(parseFloat(avgA)+parseFloat(avgC)));
    var hTotG=Math.round(hnRn*(parseFloat(hAvgA)+parseFloat(hAvgC)));
    var avgLos=(2.8+v%5*0.3).toFixed(1)+'n', hLos=(2.8+v%5*0.3+0.4).toFixed(1)+'n';
    var avgLead=(18+v%60)+'d', hLead=(18+v%60+12)+'d';
    var availRooms=Math.max(0,102-Math.floor(hotel*1.02));
    var availGuar=Math.floor(8+v%5);
    var aiPct=Math.max(45,Math.min(68,55+(dm*7+dd2*3)%14));
    var bbPct=Math.max(14,Math.min(28,20+(dm*11+dd2*5)%11));
    var hbPct=Math.max(6,Math.min(16,10+(dm*5+dd2*7)%9));
    var roPct=100-aiPct-bbPct-hbPct;
    var toPct=to/Math.max(1,hotel);
    var mealPlans=[
      {n:'AI',s:'All Inclusive',pct:aiPct,toPct:Math.round(aiPct*toPct*(0.9+(dm+dd2)%3*0.05)),c:'#006461'},
      {n:'BB',s:'Bed & Breakfast',pct:bbPct,toPct:Math.round(bbPct*toPct*(0.85+(dm*3+dd2)%3*0.05)),c:'#3b82f6'},
      {n:'HB',s:'Half Board',pct:hbPct,toPct:Math.round(hbPct*toPct*(0.8+(dm+dd2*2)%3*0.05)),c:'#967EF3'},
      {n:'RO',s:'Room Only',pct:roPct,toPct:Math.round(roPct*toPct*(0.95+(dm*2+dd2)%3*0.03)),c:'#f59e0b'},
    ];
    var toMix=28+Math.abs((dm*7+dd2*5)%25),dirMix=30+Math.abs((dm*5+dd2*9)%20),otaMix=20+Math.abs((dm*9+dd2*3)%18);
    var otherMix=Math.max(0,100-toMix-dirMix-otaMix);
    var tcRates=[0,1,2,3,4].map(function(i){return adr-15+Math.abs((dm*(i+3)+dd2*(i+5))%50);});
    var baseRate=adr+8;
    var sdlyH=Math.max(5,hotel-9),lyH=Math.max(5,hotel-6),fcstH=Math.min(100,hotel+4);
    var sdlyA=adr-8,lyA=adr-4,fcstA=adr+6;
    var sdlyR=Math.floor(hnRev*0.9),lyR=Math.floor(hnRev*0.95),fcstR=Math.floor(hnRev*1.06);
    // EBB 10% for first 3 days of week, Contract for other 4
    var dayOfWeek2=(new Date(2026,dm-1,dd2)).getDay(); // 0=Sun
    var isEbb2=dayOfWeek2<3; // Sun/Mon/Tue → EBB, rest → Contract
    var ebbPromo={n:'Early Bird 10%',t:'EBB 10%',d:10,c:'#16a34a'};
    var contractPromo={n:'Contract Rate',t:'Contract',d:0,c:'#2563eb'};
    var tcPromos=[0,1,2,3,4].map(function(i){
      return isEbb2 ? ebbPromo : contractPromo;
    });
      var sdlyRevpar=Math.max(40,revpar-8), lyRevpar=Math.max(40,revpar-4);
      var sdlyRn=Math.round(toRn*0.88), lyRn=Math.round(toRn*0.93), fcstRn=Math.round(toRn*1.06);
      var totAdultsT=Math.round(toRn*parseFloat(avgA)), totChildrenT=Math.round(toRn*parseFloat(avgC));
      var totAdultsH=Math.round(hnRn*parseFloat(hAvgA)), totChildrenH=Math.round(hnRn*parseFloat(hAvgC));
      // Meal plan guest counts (rooms * avg occupancy)
      mealPlans=mealPlans.map(function(p,pi){
        var hotelRms=Math.round(hnRn*p.pct/100);
        var toRmsPlan=Math.max(0,Math.round(toRn*Math.max(0,p.toPct)/100));
        var hAdults=Math.round(hotelRms*parseFloat(hAvgA)), hChildren=Math.round(hotelRms*parseFloat(hAvgC));
        var tAdults=Math.round(toRmsPlan*parseFloat(avgA)), tChildren=Math.round(toRmsPlan*parseFloat(avgC));
        // ADR Gross = adr, ADR Net = adr * 0.88 (net of commission)
        var hAdrGross=adr+[0,4,-2,6][pi%4], hAdrNet=Math.round(hAdrGross*0.88);
        var tAdrGross=toAdr+[0,3,-1,5][pi%4], tAdrNet=Math.round(tAdrGross*0.88);
        var hRev=Math.floor(hotelRms*hAdrGross), tRev=Math.floor(toRmsPlan*tAdrGross);
        return Object.assign({},p,{
          toPct:Math.max(0,p.toPct),
          hotelRms:hotelRms, toRms:toRmsPlan,
          hAdults:hAdults,   tAdults:tAdults,
          hChildren:hChildren, tChildren:tChildren,
          hGuests:hAdults+hChildren, tGuests:tAdults+tChildren,
          hRev:hRev, tRev:tRev,
          hAdrGross:hAdrGross, tAdrGross:tAdrGross,
          hAdrNet:hAdrNet, tAdrNet:tAdrNet,
          hotelGuests:hAdults+hChildren, toGuests:tAdults+tChildren
        });
      });
        var RT2=[['Standard',51],['Superior',36],['Deluxe',27],['Suite',12],['Jr. Suite',15],['Family',9]];
        var rtRows2=RT2.map(function(r,ri){
          var inv=r[1];
          var totalSold=Math.min(inv,Math.floor(inv*hotel/110));
          var toSold=Math.min(totalSold,Math.round(totalSold*to/Math.max(1,hotel)));
          var otherSold=totalSold-toSold;
          var toAlloc=Math.floor(inv*0.8+Math.abs((dm*(ri+3)+dd2*(ri+5))%15));
          var avail=Math.max(0,inv-totalSold);
          return{name:r[0],cap:inv,toRn:toSold,otherRn:otherSold,alloc:toAlloc,avail:avail};
        });
        return {hotel,to,adr,toAdr,toRn,hnRn,toRev,hnRev,otherPct,freePct,toRms,otherRms,freeRms,
          fitPct,dynPct,serPct,onlinePct,revpar,hotelRevpar,sdlyRevpar,lyRevpar,pickup,hotelPickup,
          avgA,avgC,hAvgA,hAvgC,totG,hTotG,avgLos,hLos,avgLead,hLead,availRooms,availGuar,
          aiPct,bbPct,hbPct,roPct,mealPlans,toMix,dirMix,otaMix,otherMix,tcRates,tcPromos,baseRate,
          sdlyH,lyH,fcstH,sdlyA,lyA,fcstA,sdlyR,lyR,fcstR,
          sdlyRn,lyRn,fcstRn,totAdultsT,totChildrenT,totAdultsH,totChildrenH,rtRows2,v};
  }

  function fmtRev(v){return v>=1000000?'$'+(v/1000000).toFixed(1)+'M':'$'+Math.round(v/1000)+'k';}
  var GROUPS = [
    // ── Daily Metrics ────────────────────────────────────────────────────────
    { id:'daily', label:'Daily Metrics', clr:'#006461', metrics:[
      { lbl:'Occupancy', cols:[
        {child:'T',     fn:function(d){return{t:d.to+'%',      clr:'#006461',bold:true};}},
        {child:'Hotel', fn:function(d){return{t:d.hotel+'%',   clr:'#374151'};}},
        {child:'STLY',  fn:function(d){return{t:d.sdlyH+'%',   clr:'#9ca3af'};}},
        {child:'LY',    fn:function(d){return{t:d.lyH+'%',     clr:'#9ca3af'};}},
        {child:'Fcst',  fn:function(d){return{t:d.fcstH+'%',   clr:'#f59e0b'};}},
      ]},
      { lbl:'T Dist. Hubs', cols:[
        {child:'Rooms', fn:function(d){return{t:d.toRms+' rm', clr:'#006461',bold:true};}},
        {child:'%',     fn:function(d){return{t:d.to+'%',      clr:'#006461'};}},
      ]},
      { lbl:'Other Segs', cols:[
        {child:'Rooms', fn:function(d){return{t:d.otherRms+' rm',clr:'#5883ed'};}},
        {child:'%',     fn:function(d){return{t:d.otherPct+'%', clr:'#5883ed'};}},
      ]},
      { lbl:'Remaining', cols:[
        {child:'Rooms', fn:function(d){return{t:d.freeRms+' rm',clr:'#16a34a',bold:true};}},
        {child:'%',     fn:function(d){return{t:Math.max(0,Math.round(d.freePct))+'%',clr:'#16a34a'};}},
      ]},
      { lbl:'Online/Offline', cols:[
        {child:'Online', fn:function(d){return{t:d.onlinePct+'%',      clr:'#3b82f6'};}},
        {child:'Offline',fn:function(d){return{t:(100-d.onlinePct)+'%',clr:'#f97316'};}},
      ]},
      { lbl:'ADR', cols:[
        {child:'T',     fn:function(d){return{t:'$'+d.toAdr,  clr:'#7c3aed',bold:true};}},
        {child:'Hotel', fn:function(d){return{t:'$'+d.adr,    clr:'#374151'};}},
        {child:'Diff',  fn:function(d){var df=d.toAdr-d.adr;return{t:(df>=0?'+':'')+df,clr:df>=0?'#16a34a':'#dc2626'};}},
        {child:'STLY',  fn:function(d){return{t:'$'+d.sdlyA,  clr:'#9ca3af'};}},
        {child:'LY',    fn:function(d){return{t:'$'+d.lyA,    clr:'#9ca3af'};}},
        {child:'Fcst',  fn:function(d){return{t:'$'+d.fcstA,  clr:'#f59e0b'};}},
      ]},
      { lbl:'Revenue', cols:[
        {child:'T',     fn:function(d){return{t:fmtRev(d.toRev), clr:'#ea580c',bold:true};}},
        {child:'Hotel', fn:function(d){return{t:fmtRev(d.hnRev), clr:'#374151'};}},
        {child:'STLY',  fn:function(d){return{t:fmtRev(d.sdlyR), clr:'#9ca3af'};}},
        {child:'LY',    fn:function(d){return{t:fmtRev(d.lyR),   clr:'#9ca3af'};}},
        {child:'Fcst',  fn:function(d){return{t:fmtRev(d.fcstR), clr:'#f59e0b'};}},
      ]},
      { lbl:'REVPAR', cols:[
        {child:'T',     fn:function(d){return{t:'$'+d.revpar,      clr:'#9333ea',bold:true};}},
        {child:'Hotel', fn:function(d){return{t:'$'+d.hotelRevpar, clr:'#374151'};}},
        {child:'STLY',  fn:function(d){return{t:'$'+d.sdlyRevpar,  clr:'#9ca3af'};}},
        {child:'LY',    fn:function(d){return{t:'$'+d.lyRevpar,    clr:'#9ca3af'};}},
      ]},
      { lbl:'Pickup', cols:[
        {child:'T',     fn:function(d){return{t:'+'+d.pickup,      clr:'#16a34a',bold:true};}},
        {child:'Hotel', fn:function(d){return{t:'+'+d.hotelPickup, clr:'#374151'};}},
      ]},
    ]},

    // ── Segments (T only) ────────────────────────────────────────────────────
    { id:'segs', label:'Segments (T)', clr:'#0891b2', metrics:[
      { lbl:'Static FIT', cols:[
        {child:'Rooms', fn:function(d){return{t:Math.round(250*d.fitPct/100)+' rm',clr:'#006461',bold:true};}},
        {child:'%',     fn:function(d){return{t:d.fitPct+'%', clr:'#006461'};}},
      ]},
      { lbl:'TO Dynamic', cols:[
        {child:'Rooms', fn:function(d){return{t:Math.round(250*d.dynPct/100)+' rm',clr:'#0891b2',bold:true};}},
        {child:'%',     fn:function(d){return{t:d.dynPct+'%', clr:'#0891b2'};}},
      ]},
      { lbl:'Tour Series', cols:[
        {child:'Rooms', fn:function(d){return{t:Math.round(250*d.serPct/100)+' rm',clr:'#6366f1',bold:true};}},
        {child:'%',     fn:function(d){return{t:d.serPct+'%', clr:'#6366f1'};}},
      ]},
    ]},

    // ── More Metrics ─────────────────────────────────────────────────────────
    { id:'more', label:'More Metrics', clr:'#2e65e8', metrics:[
      { lbl:'RN Sold', cols:[
        {child:'T',     fn:function(d){return{t:d.toRn,        clr:'#2e65e8',bold:true};}},
        {child:'Hotel', fn:function(d){return{t:d.hnRn,        clr:'#374151'};}},
        {child:'STLY',  fn:function(d){return{t:d.sdlyRn,      clr:'#9ca3af'};}},
        {child:'LY',    fn:function(d){return{t:d.lyRn,        clr:'#9ca3af'};}},
        {child:'Fcst',  fn:function(d){return{t:d.fcstRn,      clr:'#f59e0b'};}},
      ]},
      { lbl:'Avg Adults', cols:[
        {child:'T',     fn:function(d){return{t:d.avgA,        clr:'#2e65e8',bold:true};}},
        {child:'Hotel', fn:function(d){return{t:d.hAvgA,       clr:'#374151'};}},
      ]},
      { lbl:'Avg Children', cols:[
        {child:'T',     fn:function(d){return{t:d.avgC,        clr:'#d33030',bold:true};}},
        {child:'Hotel', fn:function(d){return{t:d.hAvgC,       clr:'#374151'};}},
      ]},
      { lbl:'Total Adults', cols:[
        {child:'T',     fn:function(d){return{t:d.totAdultsT,  clr:'#2e65e8',bold:true};}},
        {child:'Hotel', fn:function(d){return{t:d.totAdultsH,  clr:'#374151'};}},
      ]},
      { lbl:'Total Children', cols:[
        {child:'T',     fn:function(d){return{t:d.totChildrenT,clr:'#d33030',bold:true};}},
        {child:'Hotel', fn:function(d){return{t:d.totChildrenH,clr:'#374151'};}},
      ]},
      { lbl:'Tot. Guests', cols:[
        {child:'T',     fn:function(d){return{t:d.totG,        clr:'#0369a1',bold:true};}},
        {child:'Hotel', fn:function(d){return{t:d.hTotG,       clr:'#374151'};}},
      ]},
      { lbl:'Avg LOS', cols:[
        {child:'T',     fn:function(d){return{t:d.avgLos,      clr:'#0891b2',bold:true};}},
        {child:'Hotel', fn:function(d){return{t:d.hLos,        clr:'#374151'};}},
      ]},
      { lbl:'Lead Time', cols:[
        {child:'T',     fn:function(d){return{t:d.avgLead,     clr:'#6366f1',bold:true};}},
        {child:'Hotel', fn:function(d){return{t:d.hLead,       clr:'#374151'};}},
      ]},
    ]},

    // ── Room Availability (per room type) ─────────────────────────────────────
    { id:'avail', label:'Room Availability', clr:'#16a34a',
      get metrics(){
        var cols=['Standard','Superior','Deluxe','Suite','Jr. Suite','Family'];
        var mets=cols.map(function(nm,ri){
          var cap2=[51,36,27,12,15,9][ri];
          return { lbl:nm+' ('+cap2+')', cols:[
            {child:'TO RN', fn:function(d){var r=d.rtRows2[ri];return{t:r.toRn,   clr:'#006461',bold:true};}},
            {child:'Other', fn:function(d){var r=d.rtRows2[ri];return{t:r.otherRn,clr:'#5883ed'};}},
            {child:'Alloc', fn:function(d){var r=d.rtRows2[ri];return{t:r.alloc,  clr:'#fb923c'};}},
            {child:'Avail', fn:function(d){var r=d.rtRows2[ri];var av=r.avail;return{t:av,clr:av===0?'#ef4444':'#16a34a',bold:true};}},
          ]};
        });
        // Add totals row
        mets.push({ lbl:'TOTAL', cols:[
          {child:'Cap',   fn:function(d){return{t:d.rtRows2.reduce(function(s,r){return s+r.cap;},0),    clr:'#374151',bold:true};}},
          {child:'TO RN', fn:function(d){return{t:d.rtRows2.reduce(function(s,r){return s+r.toRn;},0),   clr:'#006461',bold:true};}},
          {child:'Other', fn:function(d){return{t:d.rtRows2.reduce(function(s,r){return s+r.otherRn;},0),clr:'#5883ed',bold:true};}},
          {child:'Alloc', fn:function(d){return{t:d.rtRows2.reduce(function(s,r){return s+r.alloc;},0),  clr:'#fb923c',bold:true};}},
          {child:'Avail', fn:function(d){var av=d.rtRows2.reduce(function(s,r){return s+r.avail;},0);return{t:av,clr:av===0?'#ef4444':'#16a34a',bold:true};}},
        ]});
        return mets;
      }
    },

    // ── Meal Plans — 4 separate collapsible groups ────────────────────────────
    { id:'mp_ai', label:'Meal Plan: AI', clr:'#006461', metrics:[
      { lbl:'Rooms Sold', cols:[
        {child:'Hotel', fn:function(d){return{t:d.mealPlans[0].hotelRms,  clr:'#374151'};}},
        {child:'T',     fn:function(d){return{t:d.mealPlans[0].toRms,    clr:'#006461',bold:true};}},
      ]},
      { lbl:'Occ %', cols:[
        {child:'Hotel', fn:function(d){return{t:d.aiPct+'%',                             clr:'#374151'};}},
        {child:'TO',    fn:function(d){return{t:Math.max(0,d.mealPlans[0].toPct)+'%',   clr:'#006461',bold:true};}},
      ]},
      { lbl:'Adults', cols:[
        {child:'Hotel', fn:function(d){return{t:d.mealPlans[0].hAdults,  clr:'#374151'};}},
        {child:'T',     fn:function(d){return{t:d.mealPlans[0].tAdults,  clr:'#006461',bold:true};}},
      ]},
      { lbl:'Children', cols:[
        {child:'Hotel', fn:function(d){return{t:d.mealPlans[0].hChildren,clr:'#374151'};}},
        {child:'T',     fn:function(d){return{t:d.mealPlans[0].tChildren,clr:'#006461',bold:true};}},
      ]},
      { lbl:'Tot. Guests', cols:[
        {child:'Hotel', fn:function(d){return{t:d.mealPlans[0].hGuests,  clr:'#374151'};}},
        {child:'T',     fn:function(d){return{t:d.mealPlans[0].tGuests,  clr:'#006461',bold:true};}},
      ]},
      { lbl:'Revenue', cols:[
        {child:'Hotel', fn:function(d){return{t:fmtRev(d.mealPlans[0].hRev),clr:'#374151'};}},
        {child:'T',     fn:function(d){return{t:fmtRev(d.mealPlans[0].tRev),clr:'#006461',bold:true};}},
      ]},
      { lbl:'ADR Gross', cols:[
        {child:'Hotel', fn:function(d){return{t:'$'+d.mealPlans[0].hAdrGross,clr:'#374151'};}},
        {child:'T',     fn:function(d){return{t:'$'+d.mealPlans[0].tAdrGross,clr:'#006461',bold:true};}},
      ]},
      { lbl:'ADR Net', cols:[
        {child:'Hotel', fn:function(d){return{t:'$'+d.mealPlans[0].hAdrNet, clr:'#374151'};}},
        {child:'T',     fn:function(d){return{t:'$'+d.mealPlans[0].tAdrNet, clr:'#006461',bold:true};}},
      ]},
    ]},
    { id:'mp_bb', label:'Meal Plan: BB', clr:'#3b82f6', metrics:[
      { lbl:'Rooms Sold', cols:[
        {child:'Hotel', fn:function(d){return{t:d.mealPlans[1].hotelRms,  clr:'#374151'};}},
        {child:'T',     fn:function(d){return{t:d.mealPlans[1].toRms,    clr:'#3b82f6',bold:true};}},
      ]},
      { lbl:'Occ %', cols:[
        {child:'Hotel', fn:function(d){return{t:d.bbPct+'%',                             clr:'#374151'};}},
        {child:'TO',    fn:function(d){return{t:Math.max(0,d.mealPlans[1].toPct)+'%',   clr:'#3b82f6',bold:true};}},
      ]},
      { lbl:'Adults', cols:[
        {child:'Hotel', fn:function(d){return{t:d.mealPlans[1].hAdults,  clr:'#374151'};}},
        {child:'T',     fn:function(d){return{t:d.mealPlans[1].tAdults,  clr:'#3b82f6',bold:true};}},
      ]},
      { lbl:'Children', cols:[
        {child:'Hotel', fn:function(d){return{t:d.mealPlans[1].hChildren,clr:'#374151'};}},
        {child:'T',     fn:function(d){return{t:d.mealPlans[1].tChildren,clr:'#3b82f6',bold:true};}},
      ]},
      { lbl:'Tot. Guests', cols:[
        {child:'Hotel', fn:function(d){return{t:d.mealPlans[1].hGuests,  clr:'#374151'};}},
        {child:'T',     fn:function(d){return{t:d.mealPlans[1].tGuests,  clr:'#3b82f6',bold:true};}},
      ]},
      { lbl:'Revenue', cols:[
        {child:'Hotel', fn:function(d){return{t:fmtRev(d.mealPlans[1].hRev),clr:'#374151'};}},
        {child:'T',     fn:function(d){return{t:fmtRev(d.mealPlans[1].tRev),clr:'#3b82f6',bold:true};}},
      ]},
      { lbl:'ADR Gross', cols:[
        {child:'Hotel', fn:function(d){return{t:'$'+d.mealPlans[1].hAdrGross,clr:'#374151'};}},
        {child:'T',     fn:function(d){return{t:'$'+d.mealPlans[1].tAdrGross,clr:'#3b82f6',bold:true};}},
      ]},
      { lbl:'ADR Net', cols:[
        {child:'Hotel', fn:function(d){return{t:'$'+d.mealPlans[1].hAdrNet, clr:'#374151'};}},
        {child:'T',     fn:function(d){return{t:'$'+d.mealPlans[1].tAdrNet, clr:'#3b82f6',bold:true};}},
      ]},
    ]},
    { id:'mp_hb', label:'Meal Plan: HB', clr:'#967EF3', metrics:[
      { lbl:'Rooms Sold', cols:[
        {child:'Hotel', fn:function(d){return{t:d.mealPlans[2].hotelRms,  clr:'#374151'};}},
        {child:'T',     fn:function(d){return{t:d.mealPlans[2].toRms,    clr:'#967EF3',bold:true};}},
      ]},
      { lbl:'Occ %', cols:[
        {child:'Hotel', fn:function(d){return{t:d.hbPct+'%',                             clr:'#374151'};}},
        {child:'TO',    fn:function(d){return{t:Math.max(0,d.mealPlans[2].toPct)+'%',   clr:'#967EF3',bold:true};}},
      ]},
      { lbl:'Adults', cols:[
        {child:'Hotel', fn:function(d){return{t:d.mealPlans[2].hAdults,  clr:'#374151'};}},
        {child:'T',     fn:function(d){return{t:d.mealPlans[2].tAdults,  clr:'#967EF3',bold:true};}},
      ]},
      { lbl:'Children', cols:[
        {child:'Hotel', fn:function(d){return{t:d.mealPlans[2].hChildren,clr:'#374151'};}},
        {child:'T',     fn:function(d){return{t:d.mealPlans[2].tChildren,clr:'#967EF3',bold:true};}},
      ]},
      { lbl:'Tot. Guests', cols:[
        {child:'Hotel', fn:function(d){return{t:d.mealPlans[2].hGuests,  clr:'#374151'};}},
        {child:'T',     fn:function(d){return{t:d.mealPlans[2].tGuests,  clr:'#967EF3',bold:true};}},
      ]},
      { lbl:'Revenue', cols:[
        {child:'Hotel', fn:function(d){return{t:fmtRev(d.mealPlans[2].hRev),clr:'#374151'};}},
        {child:'T',     fn:function(d){return{t:fmtRev(d.mealPlans[2].tRev),clr:'#967EF3',bold:true};}},
      ]},
      { lbl:'ADR Gross', cols:[
        {child:'Hotel', fn:function(d){return{t:'$'+d.mealPlans[2].hAdrGross,clr:'#374151'};}},
        {child:'T',     fn:function(d){return{t:'$'+d.mealPlans[2].tAdrGross,clr:'#967EF3',bold:true};}},
      ]},
      { lbl:'ADR Net', cols:[
        {child:'Hotel', fn:function(d){return{t:'$'+d.mealPlans[2].hAdrNet, clr:'#374151'};}},
        {child:'T',     fn:function(d){return{t:'$'+d.mealPlans[2].tAdrNet, clr:'#967EF3',bold:true};}},
      ]},
    ]},
    { id:'mp_ro', label:'Meal Plan: RO', clr:'#f59e0b', metrics:[
      { lbl:'Rooms Sold', cols:[
        {child:'Hotel', fn:function(d){return{t:d.mealPlans[3].hotelRms,  clr:'#374151'};}},
        {child:'T',     fn:function(d){return{t:d.mealPlans[3].toRms,    clr:'#f59e0b',bold:true};}},
      ]},
      { lbl:'Occ %', cols:[
        {child:'Hotel', fn:function(d){return{t:d.roPct+'%',                             clr:'#374151'};}},
        {child:'TO',    fn:function(d){return{t:Math.max(0,d.mealPlans[3].toPct)+'%',   clr:'#f59e0b',bold:true};}},
      ]},
      { lbl:'Adults', cols:[
        {child:'Hotel', fn:function(d){return{t:d.mealPlans[3].hAdults,  clr:'#374151'};}},
        {child:'T',     fn:function(d){return{t:d.mealPlans[3].tAdults,  clr:'#f59e0b',bold:true};}},
      ]},
      { lbl:'Children', cols:[
        {child:'Hotel', fn:function(d){return{t:d.mealPlans[3].hChildren,clr:'#374151'};}},
        {child:'T',     fn:function(d){return{t:d.mealPlans[3].tChildren,clr:'#f59e0b',bold:true};}},
      ]},
      { lbl:'Tot. Guests', cols:[
        {child:'Hotel', fn:function(d){return{t:d.mealPlans[3].hGuests,  clr:'#374151'};}},
        {child:'T',     fn:function(d){return{t:d.mealPlans[3].tGuests,  clr:'#f59e0b',bold:true};}},
      ]},
      { lbl:'Revenue', cols:[
        {child:'Hotel', fn:function(d){return{t:fmtRev(d.mealPlans[3].hRev),clr:'#374151'};}},
        {child:'T',     fn:function(d){return{t:fmtRev(d.mealPlans[3].tRev),clr:'#f59e0b',bold:true};}},
      ]},
      { lbl:'ADR Gross', cols:[
        {child:'Hotel', fn:function(d){return{t:'$'+d.mealPlans[3].hAdrGross,clr:'#374151'};}},
        {child:'T',     fn:function(d){return{t:'$'+d.mealPlans[3].tAdrGross,clr:'#f59e0b',bold:true};}},
      ]},
      { lbl:'ADR Net', cols:[
        {child:'Hotel', fn:function(d){return{t:'$'+d.mealPlans[3].hAdrNet, clr:'#374151'};}},
        {child:'T',     fn:function(d){return{t:'$'+d.mealPlans[3].tAdrNet, clr:'#f59e0b',bold:true};}},
      ]},
    ]},

    // ── Business Mix ─────────────────────────────────────────────────────────
    { id:'biz', label:'Business Mix', clr:'#0284c7', metrics:[
      { lbl:'TO',     cols:[{child:'%',fn:function(d){return{t:d.toMix+'%',   clr:'#006461',bold:true};}}]},
      { lbl:'Direct', cols:[{child:'%',fn:function(d){return{t:d.dirMix+'%',  clr:'#0284c7',bold:true};}}]},
      { lbl:'OTA',    cols:[{child:'%',fn:function(d){return{t:d.otaMix+'%',  clr:'#D97706',bold:true};}}]},
      { lbl:'Other',  cols:[{child:'%',fn:function(d){return{t:d.otherMix+'%',clr:'#9ca3af',bold:true};}}]},
    ]},

    // ── Travel Co. Rates ─────────────────────────────────────────────────────
    { id:'tc', label:'Travel Co. Rates', clr:'#0f766e', metrics:[
      { lbl:'Sunshine',    cols:[{child:'Rate',fn:function(d){return{t:'$'+d.tcRates[0],clr:'#3b82f6',bold:true,badge:d.tcPromos[0]};}}]},
      { lbl:'Global Adv.', cols:[{child:'Rate',fn:function(d){return{t:'$'+d.tcRates[1],clr:'#967EF3',bold:true,badge:d.tcPromos[1]};}}]},
      { lbl:'Beach Hols',  cols:[{child:'Rate',fn:function(d){return{t:'$'+d.tcRates[2],clr:'#0ea5e9',bold:true,badge:d.tcPromos[2]};}}]},
      { lbl:'City Breaks', cols:[{child:'Rate',fn:function(d){return{t:'$'+d.tcRates[3],clr:'#10b981',bold:true,badge:d.tcPromos[3]};}}]},
      { lbl:'Adventure',   cols:[{child:'Rate',fn:function(d){return{t:'$'+d.tcRates[4],clr:'#f59e0b',bold:true,badge:d.tcPromos[4]};}}]},
      { lbl:'Base Rate',   cols:[{child:'Rate',fn:function(d){return{t:'$'+d.baseRate,  clr:'#9333ea',bold:true};}}]},
    ]},
  ];

  // ── Flatten cols for data rendering ──────────────────────────────────────
  GROUPS.forEach(function(g){
    g.cols = [];
    g.metrics.forEach(function(m){ m.cols.forEach(function(c){ g.cols.push(c); }); });
  });

  // ── Expose toggle ─────────────────────────────────────────────────────────
  window._rptCollapsed = collapsed;
  window.wvRptToggleGroup = function(gid){
    window._rptCollapsed[gid] = !window._rptCollapsed[gid];
    var grid=document.getElementById('weekGrid');
    if(grid) grid.innerHTML = buildReportView(window._wvRptDays||days);
  };
  window._wvRptDays = days;

  // ── Build 3-level header rows: Section > Metric > Child ──────────────────
  var DOW2=['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
  var MNAMES2=['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  var TODAY_R=new Date(2026,2,9);

  // Row 1: section groups
  var hRow1='<tr class="wv-rpt-grp-hdr"><th class="wv-rpt-th-date" rowspan="3" style="min-width:110px">Stay Date<br><span style="font-size:8px;font-weight:500;opacity:.7">DOW · DBA</span></th>';
  // Row 2: metric parents
  var hRow2='<tr style="background:#f8fafc;border-top:1px solid #e5e7eb">';
  // Row 3: child labels
  var hRow3='<tr class="wv-rpt-sub-hdr">';

  GROUPS.forEach(function(g){
    var vis=!collapsed[g.id];
    var totalCols=vis?g.cols.length:1;
    // Row 1: section header spanning all child cols
    hRow1+='<th colspan="'+totalCols+'" onclick="wvRptToggleGroup(\''+g.id+'\')" '
      +'style="background:'+g.clr+';border-left:2px solid rgba(255,255,255,.25);cursor:pointer">'
      +(vis?'':'▶ ')+g.label+(vis?' ▾':' ◀')+'</th>';

    if(!vis){
      hRow2+='<th style="background:#f5f5f5;color:'+g.clr+';border-left:2px solid #e5e7eb;cursor:pointer;text-align:center;font-size:11px" onclick="wvRptToggleGroup(\''+g.id+'\')">···</th>';
      hRow3+='<th style="background:#f5f5f5;border-left:2px solid #e5e7eb"></th>';
    } else {
      // Row 2: metric parents — each spanning its children
      var isFirst=true;
      g.metrics.forEach(function(m){
        var bl=isFirst?'border-left:2px solid rgba(0,0,0,.10);':'border-left:1px solid #e9ecef;';
        isFirst=false;
        var mClr=g.clr;
        hRow2+='<th colspan="'+m.cols.length+'" style="'+bl+'background:#f1f5f9;color:#374151;font-size:8.5px;font-weight:700;text-align:center;padding:3px 6px;text-transform:none;letter-spacing:0;border-bottom:1px solid #d1d5db">'+m.lbl+'</th>';
        // Row 3: children
        m.cols.forEach(function(c,ci){
          var cbl=ci===0?bl:'';
          hRow3+='<th style="'+cbl+'font-size:8px;color:'+mClr+';font-weight:700">'+c.child+'</th>';
        });
      });
    }
  });
  hRow1+='</tr>'; hRow2+='</tr>'; hRow3+='</tr>';
  var hdrHtml = hRow1+hRow2+hRow3;


  // ── Build data rows ───────────────────────────────────────────────────────
  var rows=days.map(function(dv,idx){
    var dm=dv.month,day=dv.day;
    var data=dd(dm,day);
    var isToday=dm===3&&day===9;
    var isLocked=LOCKED_DAYS.has(dm+'-'+day);
    var dt=new Date(2026,dm-1,day);
    var dow=DOW2[dt.getDay()];
    var dba=Math.round((dt-TODAY_R)/86400000);
    var dbaStr=dba===0?'Today':dba>0?dba+' DBA':'';
    var occBg=data.hotel>=85?'rgba(94,131,237,.15)':data.hotel>=70?'rgba(94,131,237,.08)':data.hotel>=55?'rgba(94,131,237,.04)':'#fff';

    var trCls=isToday?'wv-rpt-row-today':isLocked?'wv-rpt-row-locked':idx%2?'wv-rpt-row-odd':'';
    var trCls=isToday?'wv-rpt-row-today':isLocked?'wv-rpt-row-locked':idx%2?'wv-rpt-row-odd':'';
    var r='<tr class="'+trCls+'">';
    var evtKey2=dm+'-'+day;
    var evts2=(typeof CAL_EVENTS!=='undefined'&&CAL_EVENTS[evtKey2])?CAL_EVENTS[evtKey2]:null;
    // Date cell
    r+='<td class="wv-rpt-cell-date" style="background:'+occBg+'">'
      +'<div style="font-weight:800;font-size:11px">'+(isLocked?'🔒 ':'')+MNAMES2[dm]+' '+day+'</div>'
      +'<div style="display:flex;align-items:center;gap:4px;font-size:8px;color:#6b7280">'
      +'<span>'+dow+'</span>'
      +(dbaStr?'<span style="color:#006461;font-weight:700">'+dbaStr+'</span>':'')
      +(evts2?'<span title="'+evts2.map(function(e){return e.name;}).join(', ')+'" '
        +'style="width:8px;height:8px;border-radius:2px;background:#C4FF45;flex-shrink:0;cursor:help;display:inline-block"></span>':'')
      +'</div>'
      +(isToday?'<div style="width:24px;height:2px;background:#006461;border-radius:1px;margin-top:2px"></div>':'')
      +'</td>';

    GROUPS.forEach(function(g){
      if(collapsed[g.id]){
        r+='<td class="wv-rpt-td" style="text-align:center;color:'+g.clr+';border-left:2px solid #e5e7eb;font-size:13px">···</td>';
        return;
      }
      g.cols.forEach(function(c,ci){
        var res=c.fn(data);
        var bl=ci===0?'border-left:2px solid rgba(0,0,0,.06);':'';
        var clr=res.clr||(!res.plain&&!res.bar?g.clr:'#006461');
        r+='<td class="wv-rpt-td" style="'+bl+'">';
        if(res.badge){
          r+='<div style="display:flex;align-items:center;gap:3px;justify-content:flex-end">'
            +'<span style="font-size:7px;font-weight:700;background:'+res.badge.c+';color:#fff;border-radius:2px;padding:1px 3px">'+res.badge.t+'</span>'
            +'<span style="font-weight:'+(res.bold?'800':'700')+';color:'+clr+'">'+res.t+'</span>'
            +'</div>';
        } else {
          // No bars — show Hotel (grey) then T (coloured) side by side
          var tClr = res.clr || (res.bar ? res.bar.c : g.clr);
          if(res.h) r+='<span style="font-size:8.5px;color:#9ca3af;margin-right:4px">'+res.h+'</span>';
          r+='<span style="font-size:9px;font-weight:'+(res.bold?'800':'700')+';color:'+tClr+'">'+res.t+'</span>';
        }
        r+='</td>';
      });
    });
    r+='</tr>';
    return r;
  }).join('');
  return '<div class="wv-report-wrap"><table class="wv-report-tbl"><thead>'+hdrHtml+'</thead><tbody>'+rows+'</tbody></table></div>';
}


window.wvSetSegMode = function(mode) {
  wvSegMode = mode;
  document.getElementById('wvSegCombined').classList.toggle('active', mode === 'combined');
  document.getElementById('wvSegIndividual').classList.toggle('active', mode === 'individual');
  buildWeekGrid(wvMonth, wvWeekStart, wvWeekStart);
};

// DOM-reorder .wv-acc-sect elements within each .wv-acc-group per _wvSectionOrder
function applyWvSectionOrder(grid) {
  if (!_wvSectionOrder || !_wvSectionOrder.length) return;
  grid.querySelectorAll('.wv-acc-group').forEach(function(group) {
    var sects = {};
    group.querySelectorAll('.wv-acc-sect').forEach(function(el) {
      var hdr = el.querySelector('.wv-acc-hdr[data-section]');
      if (hdr) sects[hdr.dataset.section] = el;
    });
    // Append in desired order; any key missing from sects is silently skipped
    _wvSectionOrder.forEach(function(key) {
      if (sects[key]) group.appendChild(sects[key]);
    });
  });
}

function buildWeekGrid(month, weekStart, activeDay) {
  const days = getWeekDays(2026, month, weekStart);
  const rangeEl = document.getElementById('wvRange');
  const m0 = days[0], m6 = days[6];
  const MNAMES = ['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  rangeEl.textContent = m0.month === m6.month
    ? `${MNAMES[m0.month]} ${m0.day} – ${m6.day}, 2026`
    : `${MNAMES[m0.month]} ${m0.day} – ${MNAMES[m6.month]} ${m6.day}, 2026`;
  // Re-render picker grid if panel is open (keeps highlight in sync with prev/next nav)
  var _pp = document.getElementById('wvWeekPickPanel');
  if (_pp && _pp.style.display !== 'none') { wvwpViewMonth = month; wvwpViewYear = wvYear; wvwpRender(); }

  const grid = document.getElementById('weekGrid');

  // Always hide section panel immediately — only shown at end for 'combined' tab
  (function(){ var p = document.getElementById('wvSectionPanel'); if (p) p.style.display = 'none'; })();

  // destroy daily rev grid when switching away
  if (wvGroupBy !== 'report' && _dailyRevGridApi) {
    try { _dailyRevGridApi.destroy(); } catch(e) {}
    _dailyRevGridApi = null;
  }
  // destroy co report grid when switching away
  if (wvGroupBy !== 'coReport' && _coReportGridApi) {
    try { _coReportGridApi.destroy(); } catch(e) {}
    _coReportGridApi = null;
  }
  // destroy daily-H grid when switching away
  if (wvGroupBy !== 'dailyH' && _dailyHGridApi) {
    try { _dailyHGridApi.destroy(); } catch(e) {}
    _dailyHGridApi = null;
  }
  // ── Report view: AG Grid ──────────────────────────────────────────────────
  if (wvGroupBy === 'report') {
    initDailyRevGrid(days, grid);
    return;
  }
  if (wvGroupBy === 'coReport') {
    initCoReportGrid(days, grid);
    return;
  }
  if (wvGroupBy === 'dailyH') {
    initDailyHGrid(days, month, activeDay, grid);
    return;
  }
  // ── Build 7-day aggregate summary (runs for all views including dailyB) ──
  var sumRn=0,sumHotelRn=0,sumRev=0,sumHotelRev=0,sumAdr=0,sumHotelAdr=0;
  var sumHotel=0,sumTo=0,sumPickup=0,sumHotelPickup=0;
  var sumAvgLos=0,sumHotelLos=0,sumAvgLead=0,sumHotelLead=0;
  var sumTotAdults=0,sumTotChildren=0,sumHotelAdults=0,sumHotelChildren=0;
  var sumRevpar=0,sumHotelRevpar=0,sumAvailRooms=0,sumAvailGuar=0;
  var sumAiPct=0,sumBbPct=0,sumHbPct=0,sumRoPct=0;
  var sumToMix=0,sumDirMix=0,sumOtaMix=0;
  var sumFitPct=0,sumDynPct=0,sumSerPct=0,sumOnline=0,sumOtherPct=0,sumFreePct=0;
  var sumTcRates=[[0,0],[0,0],[0,0],[0,0],[0,0]]; // [sum,count] per operator
  days.forEach(function(dv) {
    const dm2=dv.month,dd2=dv.day;
    const {hotel:h2,to:t2}=getOccupancy(dm2,dd2);
    const adr2=150+Math.abs((dm2*47+dd2*31)%130);
    const v2=Math.abs((dm2*127+dd2*53+dm2*dd2*7+dd2*dd2*3))%100;
    const toAdr2=Math.max(80,adr2-20-Math.abs((dm2*3+dd2*7)%15));
    const rn2=Math.round(250*t2/100);
    const rnH2=Math.round(250*h2/100);
    const rev2=Math.floor(rn2*toAdr2);
    const revH2=Math.floor(rnH2*adr2);
    sumRn+=rn2; sumHotelRn+=rnH2;
    sumRev+=rev2; sumHotelRev+=revH2;
    sumAdr+=toAdr2; sumHotelAdr+=adr2;
    sumHotel+=h2; sumTo+=t2;
    sumPickup+=Math.max(0,Math.floor((v2%25+5)*t2/Math.max(1,h2)));
    sumHotelPickup+=Math.floor(v2%25+5);
    sumAvgLos+=2.8+v2%5*0.3; sumHotelLos+=2.8+v2%5*0.3+0.4;
    sumAvgLead+=18+v2%60; sumHotelLead+=18+v2%60+12;
    const avgA2=(1.8+v2%3*0.1); const avgC2=(0.3+v2%2*0.1);
    sumTotAdults+=Math.round(rn2*avgA2); sumTotChildren+=Math.round(rn2*avgC2);
    sumHotelAdults+=Math.round(rnH2*(avgA2+0.3)); sumHotelChildren+=Math.round(rnH2*(avgC2+0.1));
    sumRevpar+=Math.max(50,(adr2+80)-30-Math.abs((dm2*5+dd2*3)%20));
    sumHotelRevpar+=adr2+80;
    sumAvailRooms+=Math.max(0,102-Math.floor(h2*1.02));
    sumAvailGuar+=Math.floor(8+v2%5);
    sumAiPct+=Math.max(45,Math.min(68,55+(dm2*7+dd2*3)%14));
    sumBbPct+=Math.max(14,Math.min(28,20+(dm2*11+dd2*5)%11));
    sumHbPct+=Math.max(6,Math.min(16,10+(dm2*5+dd2*7)%9));
    sumToMix+=28+Math.abs((dm2*7+dd2*5)%25);
    sumDirMix+=30+Math.abs((dm2*5+dd2*9)%20);
    sumOtaMix+=20+Math.abs((dm2*9+dd2*3)%18);
    sumFitPct+=Math.round(t2*0.45); sumDynPct+=Math.round(t2*0.35);
    sumSerPct+=t2-Math.round(t2*0.45)-Math.round(t2*0.35);
    sumOtherPct+=Math.max(0,h2-t2); sumFreePct+=Math.max(0,100-h2);
    sumOnline+=Math.max(30,Math.min(80,45+Math.abs((dm2*13+dd2*7)%35)));
    [0,1,2,3,4].forEach(function(i){
      sumTcRates[i][0]+=adr2-15+Math.abs((dm2*(i+3)+dd2*(i+5))%50);
      sumTcRates[i][1]++;
    });
  });
  const n7=days.length;
  const avgToAdr=Math.round(sumAdr/n7),avgHotelAdr=Math.round(sumHotelAdr/n7);
  const avgHotel=Math.round(sumHotel/n7),avgTo=Math.round(sumTo/n7);
  const revStr=s=>s>=1000000?'$'+(s/1000000).toFixed(1)+'M':'$'+Math.round(s/1000)+'k';
  const totalRevStr=revStr(sumRev),totalHotelRevStr=revStr(sumHotelRev);
  const avgLos=(sumAvgLos/n7).toFixed(1)+'n',avgHotelLos=(sumHotelLos/n7).toFixed(1)+'n';
  const avgLead=Math.round(sumAvgLead/n7)+'d',avgHotelLead=Math.round(sumHotelLead/n7)+'d';
  const avgRevpar=Math.round(sumRevpar/n7),avgHotelRevpar=Math.round(sumHotelRevpar/n7);
  const avgAvailRooms=Math.round(sumAvailRooms/n7),avgAvailGuar=Math.round(sumAvailGuar/n7);
  const avgAiPct=Math.round(sumAiPct/n7),avgBbPct=Math.round(sumBbPct/n7);
  const avgHbPct=Math.round(sumHbPct/n7),avgRoPct=100-avgAiPct-avgBbPct-avgHbPct;
  const avgToMix=Math.round(sumToMix/n7),avgDirMix=Math.round(sumDirMix/n7);
  const avgOtaMix=Math.round(sumOtaMix/n7);
  const avgOtherMix=Math.max(0,100-avgToMix-avgDirMix-avgOtaMix);
  const avgRnH=Math.round(sumHotelRn/n7);
  const avgFitPct=Math.round(sumFitPct/n7),avgDynPct=Math.round(sumDynPct/n7),avgSerPct=Math.round(sumSerPct/n7);
  const avgOnline=Math.round(sumOnline/n7);
  const avgOtherPct=Math.round(sumOtherPct/n7),avgFreePct=Math.round(sumFreePct/n7);
  const WV_SUM=250;
  const avgFitRms=Math.round(WV_SUM*avgFitPct/100),avgDynRms=Math.round(WV_SUM*avgDynPct/100),avgSerRms=Math.round(WV_SUM*avgSerPct/100);
  const avgOtherRms=Math.round(WV_SUM*avgOtherPct/100),avgFreeRms=Math.round(WV_SUM*avgFreePct/100);
  const avgTcRates=sumTcRates.map(function(r){return Math.round(r[0]/Math.max(1,r[1]));});
  // STLY/LY/Fcst approximations
  const sdlyTo=Math.max(5,avgTo-9),lyTo=Math.max(5,avgTo-6),fcstTo=Math.min(100,avgTo+4);
  const sdlyAdr=avgToAdr-8,lyAdr=avgToAdr-4,fcstAdr=avgToAdr+6;
  const sdlyRev=revStr(Math.floor(sumRev*0.9)),lyRev=revStr(Math.floor(sumRev*0.95)),fcstRev=revStr(Math.floor(sumRev*1.06));
  const sdlyRn=Math.round(sumRn*0.88),lyRn=Math.round(sumRn*0.93),fcstRn=Math.round(sumRn*1.06);
  const sdlyRevpar=Math.max(40,avgRevpar-8),lyRevpar=Math.max(40,avgRevpar-4);
  const avgTotGuests=sumTotAdults+sumTotChildren,avgHotelTotGuests=sumHotelAdults+sumHotelChildren;
  const avgTotAdults=sumTotAdults,avgTotChildren=sumTotChildren;
  const avgHotelTotAdults=sumHotelAdults,avgHotelTotChildren=sumHotelChildren;
  // Promo
  const firstDay=days[0]; const isEbbWeek=(new Date(2026,firstDay.month-1,firstDay.day)).getDay()<3;

  // ── Helpers ─────────────────────────────────────────────────────────────
  function mBar(pct,col,col2,pct2){
    return '<div style="height:3px;background:#e5e7eb;border-radius:2px;margin-top:3px;position:relative">'
      +'<div style="height:100%;width:'+pct+'%;background:'+col+';border-radius:2px;position:absolute"></div>'
      +(col2&&pct2?'<div style="height:100%;width:'+pct2+'%;background:'+col2+';border-radius:2px;position:absolute;opacity:0.4"></div>':'')
      +'</div>';
  }
  function dualBar(tPct,hPct,clr){
    return '<div style="height:3px;border-radius:2px;margin-top:2px;background:#e5e7eb;position:relative">'
      +(hPct!=null?'<div style="height:100%;width:'+Math.min(92,hPct)+'%;background:#d1d5db;border-radius:2px;position:absolute"></div>':'')
      +'<div style="height:100%;width:'+Math.min(92,tPct)+'%;background:'+clr+';border-radius:2px;position:absolute"></div>'
      +'</div>';
  }
  function stackBar(segs){
    return '<div style="height:5px;background:#e5e7eb;border-radius:3px;margin:3px 0;display:flex;overflow:hidden">'
      +segs.map(function(s){return '<div style="width:'+s.p+'%;background:'+s.c+'"></div>';}).join('')
      +'</div>';
  }
  function sumRefRow(pairs){
    var CSS={stly:'background:#e0e7ff;color:#4338ca',ly:'background:#dcfce7;color:#15803d',fcst:'background:#fef9c3;color:#a16207'};
    return '<div style="display:flex;gap:2px;flex-wrap:wrap;margin-top:2px">'
      +pairs.map(function(p){return '<span style="font-size:7px;font-weight:700;padding:1px 4px;border-radius:3px;'+CSS[p.k]+'">'+p.l+' '+p.v+'</span>';}).join('')
      +'</div>';
  }
  function mRow2(lbl,tVal,hVal,barT,barHPct,clr,refs){
    var barHtml=barT!=null?dualBar(barT,barHPct,clr):'';
    return '<div style="display:flex;align-items:center;gap:3px;margin-bottom:4px">'
      +'<span style="font-size:8px;color:#6b7280;flex:1;min-width:0">'+lbl+'</span>'
      +(hVal&&hVal!=='\u2014'
        ?'<span style="display:flex;flex-direction:column;align-items:flex-end">'
         +'<span style="font-size:6.5px;font-weight:700;color:#9ca3af;letter-spacing:.3px;text-transform:uppercase">Hotel</span>'
         +'<span style="font-size:8px;font-weight:600;color:#6b7280">'+hVal+'</span>'
         +'</span>':''
      )
      +'<span style="display:flex;flex-direction:column;align-items:flex-end;margin-left:6px">'
      +'<span style="font-size:6.5px;font-weight:700;color:'+clr+';letter-spacing:.3px;text-transform:uppercase">TO</span>'
      +'<span style="font-size:9px;font-weight:800;color:'+clr+'">'+tVal+'</span>'
      +'</span>'
      +'</div>'
      +barHtml
      +(refs?sumRefRow(refs):'');
  }
  function colHdr(tClr){
    return '<div style="display:flex;justify-content:flex-end;gap:12px;margin-bottom:5px;padding-bottom:3px;border-bottom:1px solid #f3f4f6">'
      +'<span style="font-size:6.5px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.3px">Hotel</span>'
      +'<span style="font-size:6.5px;font-weight:700;color:'+(tClr||'#006461')+';text-transform:uppercase;letter-spacing:.3px;min-width:24px;text-align:right">TO</span>'
      +'</div>';
  }
  function sumSec(title,content){
    return '<div style="background:#fff;border:1px solid #e5e7eb;border-radius:8px;padding:9px 11px">'
      +'<div style="font-size:8px;font-weight:800;color:#374151;text-transform:uppercase;letter-spacing:.5px;margin-bottom:6px;padding-bottom:4px;border-bottom:1px solid #f3f4f6">'+title+'</div>'
      +content
      +'</div>';
  }
  function dotLegend(items){
    return '<div style="display:grid;grid-template-columns:1fr 1fr;gap:2px 8px;margin-top:4px">'
      +items.map(function(it){
        return '<div style="display:flex;align-items:center;gap:4px"><span style="width:6px;height:6px;border-radius:50%;background:'+it[2]+';flex-shrink:0"></span><span style="font-size:8px;color:#374151">'+it[0]+' '+it[1]+'</span></div>';
      }).join('')
      +'</div>';
  }

  // ── Sections ─────────────────────────────────────────────────────────────
  const secOcc = sumSec('Daily Metrics',
    colHdr('#006461')
    +mRow2('Occupancy',avgTo+'%',avgHotel+'%',avgTo,Math.min(92,avgHotel),'#006461',
        [{k:'stly',l:'STLY',v:sdlyTo+'%'},{k:'ly',l:'LY',v:lyTo+'%'},{k:'fcst',l:'Fcst',v:fcstTo+'%'}])
    +mRow2('ADR','$'+avgToAdr,'$'+avgHotelAdr,Math.round(avgToAdr/3.5),Math.round(avgHotelAdr/3.5),'#7c3aed',
        [{k:'stly',l:'STLY',v:'$'+sdlyAdr},{k:'ly',l:'LY',v:'$'+lyAdr},{k:'fcst',l:'Fcst',v:'$'+fcstAdr}])
    +mRow2('Revenue',totalRevStr,totalHotelRevStr,Math.min(92,Math.round(sumRev/sumHotelRev*70)),70,'#ea580c',
        [{k:'stly',l:'STLY',v:sdlyRev},{k:'ly',l:'LY',v:lyRev},{k:'fcst',l:'Fcst',v:fcstRev}])
    +mRow2('REVPAR','$'+avgRevpar,'$'+avgHotelRevpar,Math.round(avgRevpar/4),Math.round(avgHotelRevpar/4),'#9333ea',
        [{k:'stly',l:'STLY',v:'$'+sdlyRevpar},{k:'ly',l:'LY',v:'$'+lyRevpar}])
    +mRow2('Pickup','+'+sumPickup,'+'+sumHotelPickup,null,null,'#16a34a')
    +stackBar([{p:avgFitPct,c:'#006461'},{p:avgDynPct,c:'#0891b2'},{p:avgSerPct,c:'#6366f1'},{p:Math.max(0,avgTo-avgFitPct-avgDynPct-avgSerPct),c:'#5883ed'}])
    +'<div style="margin-top:4px;display:flex;flex-direction:column;gap:2px">'
    +[['Static FIT',avgFitPct,avgFitRms,'#006461'],['TO Dynamic',avgDynPct,avgDynRms,'#0891b2'],['Tour Series',avgSerPct,avgSerRms,'#6366f1'],['Other Segs',avgOtherPct,avgOtherRms,'#5883ed'],['Remaining',avgFreePct,avgFreeRms,'#16a34a']].map(function(s){
      return '<div style="display:flex;align-items:center;gap:4px"><span style="width:6px;height:6px;border-radius:50%;background:'+s[3]+';flex-shrink:0"></span>'
        +'<span style="font-size:8px;color:#374151;flex:1">'+s[0]+'</span>'
        +'<span style="font-size:8px;color:#9ca3af">'+s[2]+' rm</span>'
        +'<span style="font-size:8px;font-weight:700;color:'+s[3]+';min-width:24px;text-align:right">'+s[1]+'%</span>'
        +'</div>';
    }).join('')
    +'</div>'
    +'<div style="display:flex;justify-content:space-between;margin-top:5px;padding-top:4px;border-top:1px solid #f3f4f6">'
    +'<span style="font-size:8px;color:#3b82f6">🌐 '+avgOnline+'% online</span>'
    +'<span style="font-size:8px;color:#f97316">📴 '+(100-avgOnline)+'% offline</span>'
    +'</div>'
  );

  const secMore = sumSec('More Metrics',
    colHdr('#2e65e8')
    +mRow2('RN Sold',sumRn+' rn',sumHotelRn+' rn',Math.min(92,Math.round(sumRn/sumHotelRn*70)),70,'#2e65e8',
        [{k:'stly',l:'STLY',v:sdlyRn},{k:'ly',l:'LY',v:lyRn},{k:'fcst',l:'Fcst',v:fcstRn}])
    +mRow2('Avg Adults',(sumTotAdults/sumRn).toFixed(1),(sumHotelAdults/sumHotelRn).toFixed(1),null,null,'#2e65e8')
    +mRow2('Avg Children',(sumTotChildren/sumRn).toFixed(1),(sumHotelChildren/sumHotelRn).toFixed(1),null,null,'#d33030')
    +mRow2('Total Adults',avgTotAdults,avgHotelTotAdults,null,null,'#2e65e8')
    +mRow2('Total Children',avgTotChildren,avgHotelTotChildren,null,null,'#d33030')
    +mRow2('Total Guests',avgTotGuests,avgHotelTotGuests,null,null,'#0369a1')
    +mRow2('Avg LOS',avgLos,avgHotelLos,null,null,'#0891b2')
    +mRow2('Avg Lead',avgLead,avgHotelLead,null,null,'#6366f1')
    +mRow2('Avail Rooms',avgAvailRooms+' rm','\u2014',Math.min(92,Math.round(avgAvailRooms/250*100)),null,'#16a34a')
    +mRow2('Avail Guar.',avgAvailGuar+' rm','\u2014',null,null,'#ea580c')
  );

  // Meal Plans with per-plan T%
  const toPct7=avgTo/Math.max(1,avgHotel);
  const aiToP=Math.max(0,Math.round(avgAiPct*toPct7*0.9));
  const bbToP=Math.max(0,Math.round(avgBbPct*toPct7*0.85));
  const hbToP=Math.max(0,Math.round(avgHbPct*toPct7*0.80));
  const roToP=Math.max(0,Math.round(avgRoPct*toPct7*0.95));
  const secMeals = sumSec('Meal Plans',
    stackBar([{p:avgAiPct,c:'#006461'},{p:avgBbPct,c:'#3b82f6'},{p:avgHbPct,c:'#967EF3'},{p:avgRoPct,c:'#f59e0b'}])
    +'<div style="font-size:7px;color:#9ca3af;text-align:right;margin-bottom:2px">Hotel % · TO %</div>'
    +[['AI (All Inclusive)',avgAiPct,aiToP,'#006461'],['BB (Bed & Bkfst)',avgBbPct,bbToP,'#3b82f6'],
      ['HB (Half Board)',avgHbPct,hbToP,'#967EF3'],['RO (Room Only)',avgRoPct,roToP,'#f59e0b']].map(function(p){
      return '<div style="display:flex;align-items:center;gap:4px;margin-bottom:2px">'
        +'<span style="width:6px;height:6px;border-radius:50%;background:'+p[3]+';flex-shrink:0"></span>'
        +'<span style="font-size:8px;color:#374151;flex:1">'+p[0]+'</span>'
        +'<span style="font-size:8px;color:#6b7280">'+p[1]+'%</span>'
        +'<span style="font-size:8px;font-weight:700;color:'+p[3]+';min-width:22px;text-align:right">'+p[2]+'%</span>'
        +'</div>';
    }).join('')
  );

  const secBiz = sumSec('Business Mix',
    stackBar([{p:avgToMix,c:'#006461'},{p:avgDirMix,c:'#0284c7'},{p:avgOtaMix,c:'#D97706'},{p:avgOtherMix,c:'#9ca3af'}])
    +dotLegend([['TO',avgToMix+'%','#006461'],['Direct',avgDirMix+'%','#0284c7'],['OTA',avgOtaMix+'%','#D97706'],['Other',avgOtherMix+'%','#9ca3af']])
  );

  // Room Types section
  const RT2_DEF = [['Standard',51],['Superior',36],['Deluxe',27],['Suite',12],['Jr. Suite',15],['Family',9]];
  const RT_SUM_COLORS = ['#006461','#0891b2','#6366f1','#f59e0b','#ec4899','#10b981'];
  const totalCap = RT2_DEF.reduce(function(s,r){return s+r[1];},0);
  const secRoomTypes = sumSec('Room Types',
    stackBar(RT2_DEF.map(function(r,ri){return {p:Math.round(r[1]/totalCap*100),c:RT_SUM_COLORS[ri]};}))
    +RT2_DEF.map(function(r,ri){
      var cap=r[1];
      var totalSold=Math.min(cap,Math.floor(cap*avgHotel/110));
      var toRn=Math.min(totalSold,Math.round(totalSold*avgTo/Math.max(1,avgHotel)));
      var avail=Math.max(0,cap-totalSold);
      var occPct=Math.round(totalSold/cap*100);
      return '<div style="display:flex;align-items:center;gap:4px;margin-bottom:3px">'
        +'<span style="width:6px;height:6px;border-radius:50%;background:'+RT_SUM_COLORS[ri]+';flex-shrink:0"></span>'
        +'<span style="font-size:8px;color:#374151;flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">'+r[0]+'</span>'
        +'<span style="font-size:7.5px;color:#9ca3af;min-width:32px;text-align:right">'+avail+' avail</span>'
        +'<span style="font-size:8px;font-weight:700;color:'+RT_SUM_COLORS[ri]+';min-width:26px;text-align:right">'+toRn+'rn</span>'
        +'</div>';
    }).join('')
    +'<div style="margin-top:4px;padding-top:4px;border-top:1px solid #f3f4f6;display:flex;justify-content:space-between">'
    +'<span style="font-size:7.5px;color:#9ca3af">Cap: '+totalCap+' rooms</span>'
    +'<span style="font-size:7.5px;font-weight:700;color:#006461">'+Math.round(RT2_DEF.reduce(function(s,r){return s+Math.min(r[1],Math.floor(r[1]*avgHotel/110));},0)/totalCap*100)+'% sold</span>'
    +'</div>'
  );

  const tcOps=[['Sunshine Tours','#3b82f6'],['Global Adv.','#967EF3'],['Beach Hols','#0ea5e9'],['City Breaks','#10b981'],['Adventure','#f59e0b']];
  const promoLabel=isEbbWeek?'EBB 10%':'Contract';
  const promoClr  =isEbbWeek?'#16a34a':'#2563eb';
  const secTC = sumSec('Travel Co. Rates',
    tcOps.map(function(op,i){
      return '<div style="display:flex;align-items:center;gap:5px;margin-bottom:3px">'
        +'<span style="width:6px;height:6px;border-radius:50%;background:'+op[1]+';flex-shrink:0"></span>'
        +'<span style="font-size:8px;color:#374151;flex:1">'+op[0]+'</span>'
        +'<span style="font-size:7px;font-weight:700;padding:1px 4px;border-radius:3px;background:'+promoClr+'20;color:'+promoClr+';border:1px solid '+promoClr+'44">'+promoLabel+'</span>'
        +'<span style="font-size:9px;font-weight:700;color:'+op[1]+'">$'+avgTcRates[i]+'</span>'
        +'</div>';
    }).join('')
    +'<div style="margin-top:5px;padding-top:4px;border-top:1px solid #e5e7eb;display:flex;justify-content:space-between;align-items:center">'
    +'<span style="font-size:8px;color:#9333ea;font-weight:700">Base Seg. Rate</span>'
    +'<span style="font-size:9px;font-weight:800;color:#9333ea">$'+(avgHotelAdr+8)+'</span>'
    +'</div>'
  );

  // Close-out counts for the 7 days
  var fullCoCount7=0, partCoCount7=0;
  days.forEach(function(dv){
    if(LOCKED_DAYS.has(dv.month+'-'+dv.day)) fullCoCount7++;
    var pr=PARTIAL_CLOSURES[dv.month+'-'+dv.day];
    if(pr&&pr.length) partCoCount7++;
  });

  // Init accordion state once
  if(!_wv7dAccState._init){
    _wv7dAccState._init=true;
    _wv7dAccState['wv7d_overview']=true; // outer accordion closed on first load
    ['wv7d_co','wv7d_daily','wv7d_seg','wv7d_more','wv7d_meals','wv7d_biz','wv7d_tc',
     'mos_co_full','mos_co_part','mos_occ','mos_adr','mos_rev','mos_revpar','mos_pickup','mos_onoff',
     'mos_segbar','mos_rn','mos_avga','mos_avgc','mos_tota','mos_totc','mos_totg','mos_los','mos_lead',
     'mos_avail','mos_availg','mos_mpsum','mos_bizbar',
     'mos_tc0','mos_tc1','mos_tc2','mos_tc3','mos_tc4','mos_tcbase'
    ].forEach(function(k){_wv7dAccState[k]=false;});
  }

  _wv7dSummaryData = {
    avgTo:avgTo, avgHotel:avgHotel,
    avgToAdr:avgToAdr, avgHotelAdr:avgHotelAdr,
    totalRevStr:totalRevStr, totalHotelRevStr:totalHotelRevStr,
    avgRevpar:avgRevpar, avgHotelRevpar:avgHotelRevpar,
    sumPickup:sumPickup, sumHotelPickup:sumHotelPickup,
    sumRn:sumRn, avgRnH:avgRnH,
    avgLos:avgLos, avgHotelLos:avgHotelLos,
    avgLead:avgLead, avgHotelLead:avgHotelLead,
    avgAvailRooms:avgAvailRooms, avgAvailGuar:avgAvailGuar,
    avgTotGuests:avgTotGuests, avgHotelTotGuests:avgHotelTotGuests,
    avgTotAdults:avgTotAdults, avgTotChildren:avgTotChildren,
    avgHotelTotAdults:avgHotelTotAdults, avgHotelTotChildren:avgHotelTotChildren,
    avgAiPct:avgAiPct, avgBbPct:avgBbPct, avgHbPct:avgHbPct, avgRoPct:avgRoPct,
    aiToP:aiToP, bbToP:bbToP, hbToP:hbToP, roToP:roToP,
    avgToMix:avgToMix, avgDirMix:avgDirMix, avgOtaMix:avgOtaMix, avgOtherMix:avgOtherMix,
    avgFitPct:avgFitPct, avgDynPct:avgDynPct, avgSerPct:avgSerPct, avgOtherPct:avgOtherPct, avgFreePct:avgFreePct,
    avgFitRms:avgFitRms, avgDynRms:avgDynRms, avgSerRms:avgSerRms, avgOtherRms:avgOtherRms, avgFreeRms:avgFreeRms,
    avgOnline:avgOnline, avgTcRates:avgTcRates,
    sdlyTo:sdlyTo, lyTo:lyTo, fcstTo:fcstTo,
    sdlyAdr:sdlyAdr, lyAdr:lyAdr, fcstAdr:fcstAdr,
    sdlyRev:sdlyRev, lyRev:lyRev, fcstRev:fcstRev,
    sdlyRn:sdlyRn, lyRn:lyRn, fcstRn:fcstRn,
    sdlyRevpar:sdlyRevpar, lyRevpar:lyRevpar,
    isEbbWeek:isEbbWeek, fullCoCount7:fullCoCount7, partCoCount7:partCoCount7, n7:n7
  };

  var summaryContainer = document.getElementById('wvSummaryContainer');
  if(summaryContainer) summaryContainer.innerHTML = _buildWv7dSummaryHtml(_wv7dSummaryData);

  // Close-out heat map removed
  var coHeatmapContainer = document.getElementById('coHeatmapContainer');
  if (coHeatmapContainer) coHeatmapContainer.innerHTML = '';

  // ── Render grid based on view mode ──────────────────────────────────────
  if (wvGroupBy === 'dailyB') {
    grid.style.cssText = 'display:flex;flex-direction:column;flex:1;min-width:0;';
    grid.innerHTML = buildDailyBView(days, month, activeDay);
    return;
  }
  grid.style.cssText = '';

  grid.innerHTML = days.map(({ month: dm, day: dd }) => {
    const isToday  = dm === 3 && dd === 9;
    const isActive = dm === month && dd === activeDay;
    const isLocked = LOCKED_DAYS.has(`${dm}-${dd}`);
    const { hotel, to } = getOccupancy(dm, dd);
    const d0 = new Date(2026, dm - 1, dd);
    const dowName = DOW_FULL[d0.getDay()];
    const adr = 150 + Math.abs((dm * 47 + dd * 31) % 130);
    const rev = (hotel * adr * 1.1).toFixed(0);
    const v = Math.abs((dm * 127 + dd * 53 + dm * dd * 7 + dd * dd * 3)) % 100;

    const colClass = ['wv-col', isActive ? 'wv-active' : '', isToday ? 'wv-today' : ''].filter(Boolean).join(' ');

    const onlinePct = Math.max(30, Math.min(80, 45 + Math.abs((dm * 13 + dd * 7) % 35)));
    const adrBar = Math.min(95, 40 + Math.abs((dm * 11 + dd * 19) % 55));
    const revBar = Math.min(95, 35 + Math.abs((dm * 17 + dd * 13) % 60));

    const rtAvail = RT_NAMES.map((n, i) => {
      const alloc = 10 + Math.abs((dm * (i+3) + dd * (i+7)) % 35);
      const avail = Math.max(0, alloc - Math.floor(hotel / 15));
      return { n, alloc, avail };
    });

    // Weekly range selection classes
    const wvDv = dm * 100 + dd;
    const wvSv = wvSelStart ? wvSelStart.month * 100 + wvSelStart.day : null;
    const wvEv = wvSelEnd   ? wvSelEnd.month   * 100 + wvSelEnd.day   : null;
    const wvLo = (wvSv !== null && wvEv !== null) ? Math.min(wvSv, wvEv) : wvSv;
    const wvHi = (wvSv !== null && wvEv !== null) ? Math.max(wvSv, wvEv) : null;
    const wvSelClass = wvLo !== null && wvDv === wvLo ? ' wv-sel-lo'
                     : wvHi !== null && wvDv === wvHi ? ' wv-sel-hi'
                     : wvHi !== null && wvDv > wvLo && wvDv < wvHi ? ' wv-sel-mid'
                     : '';

    // DBA: days from today (March 9 2026) to the stay date
    const TODAY_WV = new Date(2026, 2, 9);  // March 9 2026 prototype today
    const stayDate = new Date(2026, dm - 1, dd);
    const dba = Math.round((stayDate - TODAY_WV) / 86400000);
    const dbaStr = dba === 0 ? 'Today' : dba > 0 ? dba + ' DBA' : '';
    const hdrDateStr = `${dowName} ${dm}/${dd}/25`;
    // Event icon for this day
    const wvEventKey = `${dm}-${dd}`;
    const wvEvents = (typeof CAL_EVENTS !== 'undefined' && CAL_EVENTS[wvEventKey]) ? CAL_EVENTS[wvEventKey] : null;
    const wvEventIconHtml = wvEvents
      ? `<span class="wv-event-cal-icon has-events" data-event-key="${wvEventKey}" onmouseenter="calShowEventTip(event,'${wvEventKey}')" onmouseleave="calHideEventTip()"><span class="material-icons" style="font-size:14px;color:#c4ff45">today</span></span>`
      : '';
    const isActionNeeded = hotel >= 65 && to < 40 && !isLocked;
    let wvMetricVal = hotel;
    if (wvActiveTab === 'pickup') wvMetricVal = getPickupPct(dm, dd);
    else if (wvActiveTab === 'guarantees') wvMetricVal = getGuaranteeFill(dm, dd);
    const hdrHeatClass = isLocked ? '' : (isActionNeeded ? getSegClass(to) : getHotelClass(wvMetricVal));
    const clHtml = !isLocked ? buildClosuresHtml(dm, dd) : '';
    const hasColCl = clHtml.length > 0;
    const restrictPanelId = `wvrp-${dm}-${dd}`;
    const wvIso = `2026-${String(dm).padStart(2,'0')}-${String(dd).padStart(2,'0')}`;
    const wvChk = _wvSelectedDays.has(wvIso) ? ' checked' : '';
    return `<div class="${colClass}${isLocked ? ' wv-locked' : ''}${wvSelClass}" data-dm="${dm}" data-dd="${dd}">
      <div class="wv-col-hdr ${hdrHeatClass}${isLocked ? ' closed' : ''}${isToday ? ' wv-col-hdr-today' : ''}">
        <input type="checkbox" class="wv-day-chk" data-wv-date="${wvIso}"${wvChk} onclick="event.stopPropagation();wvDayCheck('${wvIso}',this)" title="Select for close-out">
        <div style="display:flex;flex-direction:column;gap:1px;min-width:0;flex:1">
          <div style="display:flex;align-items:center;gap:4px">
            <span class="wv-col-hdr-date">${hdrDateStr}</span>
            ${wvEventIconHtml}
          </div>
          <span class="wv-col-hdr-dba" style="font-size:9px;font-weight:600;color:#fff;opacity:.75;letter-spacing:.2px">${dbaStr}</span>
        </div>
        ${isLocked ? `<svg class="wv-lock-icon" viewBox="0 0 10 12" fill="none" stroke="#dc2626" stroke-width="1.6" width="11" height="13"><rect x="1" y="5" width="8" height="7" rx="1"/><path d="M3 5V3.5a2 2 0 0 1 4 0V5"/></svg>` : ''}
        ${hasColCl ? `<button class="wv-partial-lock-btn" data-restrict-id="${restrictPanelId}" title="View closed out"><svg viewBox="0 0 10 12" fill="none" stroke="currentColor" stroke-width="1.6" width="10" height="12"><rect x="1" y="5" width="8" height="7" rx="1"/><path d="M3 5V3.5a2 2 0 0 1 4 0V5"/></svg></button>` : ''}
        ${isToday ? `<span class="wv-today-badge">TODAY</span>` : ''}
      </div>
      ${hasColCl ? `<div class="wv-restrict-panel" id="${restrictPanelId}">${clHtml}</div>` : ''}
      <div class="wv-acc-group">
      ${wvGroupBy === 'roomType' ? buildRoomTypeBoardView(dm, dd, hotel, to, adr, rev, v) : ''}
      ${wvGroupBy === 'combined' ? (wvMetricState.capacity||wvMetricState.onlineOffline||wvMetricState.adr||wvMetricState.revenue) ? wvAcc('Daily Metrics', 'daily', (function(){
        const showS = wvMetricState.cmp_sdly, showL = wvMetricState.cmp_final_ly, showF = wvMetricState.cmp_forecast, showH = wvMetricState.cmp_hotel;
        // Reference values — Hotel STLY / Final LY / Forecast
        const sdlyH = Math.max(5, hotel-9),   lyH = Math.max(5, hotel-6),   fcstH = Math.min(100, hotel+4);
        const sdlyA = adr-8,                   lyA = adr-4,                  fcstA = adr+6;
        const sdlyR = Math.floor(rev*0.9),     lyR = Math.floor(rev*0.95),   fcstR = Math.floor(rev*1.06);
        // Operator (TO) compare values — used by the Compare dropdown (wvCompare)
        const toAdrV0 = Math.max(80, adr - 20 - Math.abs((dm*3+dd*7)%15));
        const toRns0  = Math.round(250 * to / 100);
        const toRevV0 = Math.floor(toRns0 * toAdrV0);
        const sdlyTo = Math.max(5, to-9),      lyTo = Math.max(5, to-6),     fcstTo = Math.min(100, to+4);
        const sdlyToAdr = toAdrV0-5,           lyToAdr = toAdrV0-3,          fcstToAdr = toAdrV0+4;
        const sdlyToRev = Math.floor(toRevV0*0.9), lyToRev = Math.floor(toRevV0*0.95), fcstToRev = Math.floor(toRevV0*1.06);
        // Helper: compact compare chips for an operator metric (respects wvCompare dropdown)
        function _opCmp(curr, sdly, ly, fcst, fmt) {
          if (wvCompare.size === 0) return '';
          var defs = [{k:'stly',l:'STLY',v:sdly},{k:'ly',l:'LY',v:ly},{k:'fcst',l:'Fcst',v:fcst}];
          var chips = defs.filter(function(x){return wvCompare.has(x.k)&&x.v!=null;}).map(function(x){
            var n=parseFloat(curr),p=parseFloat(x.v),clr=!isNaN(n)&&!isNaN(p)?(n>p?'#059669':n<p?'#dc2626':'#8A9096'):'#8A9096';
            return '<span style="font-size:8.5px;padding:1px 4px;border-radius:3px;background:'+clr+'18;color:'+clr+';font-weight:700;white-space:nowrap">'+x.l+' '+fmt(x.v)+'</span>';
          }).join('');
          return chips ? '<div style="display:flex;gap:3px;margin-top:2px;flex-wrap:wrap;padding-left:14px">'+chips+'</div>' : '';
        }
        const adrBarRef = Math.max(3, adrBar-15), revBarRef = Math.max(3, revBar-15);
        // Build multi-colored tick marks
        function occTicks() {
          return (showS?`<div class="wv-occ-ref-tick wv-tick-sdly" style="left:${sdlyH}%" title="STLY ${sdlyH}%"></div>`:'')
               + (showL?`<div class="wv-occ-ref-tick wv-tick-ly"   style="left:${lyH}%"   title="LY ${lyH}%"></div>`:'')
               + (showF?`<div class="wv-occ-ref-tick wv-tick-fcst" style="left:${fcstH}%" title="Fcst ${fcstH}%"></div>`:'');
        }
        function barTicks(sP, lP, fP) {
          return (showS?`<div class="wv-bar-ref-tick wv-tick-sdly" style="left:${sP}%"></div>`:'')
               + (showL?`<div class="wv-bar-ref-tick wv-tick-ly"   style="left:${lP}%"></div>`:'')
               + (showF?`<div class="wv-bar-ref-tick wv-tick-fcst" style="left:${fP}%"></div>`:'');
        }
        // Build ref-tag rows
        function refRow(sv, lv, fv) {
          const parts = [];
          if (showS && sv != null) parts.push('<span class="wv-ref-tag wv-ref-sdly">STLY '+sv+'</span>');
          if (showL && lv != null) parts.push('<span class="wv-ref-tag wv-ref-ly">LY '+lv+'</span>');
          if (showF) parts.push('<span class="wv-ref-tag wv-ref-fcst">Fcst '+fv+'</span>');
          return parts.length ? '<div class="wv-ref-row">'+parts.join('')+'</div>' : '';
        }
        return `<div class="wv-quick">
          ${wvMetricState.capacity ? (function(){
            const WV_CAP   = 250;
            const otherPct = Math.max(0, hotel - to);
            const freePct  = 100 - hotel;
            const toRms    = Math.round(WV_CAP * to       / 100);
            const otherRms = Math.round(WV_CAP * otherPct / 100);
            const freeRms  = WV_CAP - toRms - otherRms;

            // Bar track
            var barInner;
            if (wvSegMode === 'individual') {
              const fitPct2 = Math.round(to * 0.45), dynPct2 = Math.round(to * 0.35), serPct2 = to - fitPct2 - dynPct2;
              barInner = '<div class="wv-occ-seg" style="width:'+fitPct2+'%;background:#006461" title="Static FIT: '+fitPct2+'%"></div>'
                       + '<div class="wv-occ-seg" style="width:'+dynPct2+'%;background:#0891b2" title="TO Dynamic: '+dynPct2+'%"></div>'
                       + '<div class="wv-occ-seg" style="width:'+serPct2+'%;background:#6366f1" title="Tour Series: '+serPct2+'%"></div>'
                       + '<div class="wv-occ-seg wv-occ-other" style="width:'+otherPct+'%" title="Other: '+otherPct+'%"></div>';
            } else {
              barInner = '<div class="wv-occ-seg wv-occ-to" style="width:'+to+'%" title="TO: '+to+'%"></div>'
                       + '<div class="wv-occ-seg wv-occ-other" style="width:'+otherPct+'%" title="Other: '+otherPct+'%"></div>';
            }

            // Breakdown rows
            var bdRows;
            if (wvSegMode === 'individual') {
              const fitPct = Math.round(to * 0.45), dynPct = Math.round(to * 0.35), serPct = to - fitPct - dynPct;
              const fitRms = Math.round(WV_CAP * fitPct / 100), dynRms = Math.round(WV_CAP * dynPct / 100), serRms = Math.round(WV_CAP * serPct / 100);
              // Individual segment compare values (derived from TO compare)
              const sdlyFit = Math.round(sdlyTo*0.45), lyFit = Math.round(lyTo*0.45), fcstFit = Math.round(fcstTo*0.45);
              const sdlyDyn = Math.round(sdlyTo*0.35), lyDyn = Math.round(lyTo*0.35), fcstDyn = Math.round(fcstTo*0.35);
              const sdlySer = Math.max(0,sdlyTo-sdlyFit-sdlyDyn), lySer = Math.max(0,lyTo-lyFit-lyDyn), fcstSer = Math.max(0,fcstTo-fcstFit-fcstDyn);
              function brRow(clr,lbl,rms,pct,extra,rmsCls){
                return '<div class="wv-occ-br-row'+(extra?' '+extra:'')+'"><div class="wv-occ-br-left"><span class="wv-occ-br-dot" style="background:'+clr+'"></span><span class="wv-occ-br-lbl">'+lbl+'</span></div><div class="wv-occ-br-right"><span class="wv-occ-br-rms'+(rmsCls?' '+rmsCls:'')+'">'+rms+' rms</span><span class="wv-occ-br-pct">'+pct+'%</span></div></div>';
              }
              bdRows = brRow('#006461','Static FIT Rates',fitRms,fitPct)
                +_opCmp(fitPct, sdlyFit, lyFit, fcstFit, function(v){return v+'%';})
                +brRow('#0891b2','TO Dynamic',dynRms,dynPct)
                +_opCmp(dynPct, sdlyDyn, lyDyn, fcstDyn, function(v){return v+'%';})
                +brRow('#6366f1','Tour Series',serRms,serPct)
                +_opCmp(serPct, sdlySer, lySer, fcstSer, function(v){return v+'%';})
                +brRow('#47c5bc','Other Segments',otherRms,otherPct)
                +brRow('#388C3F','Remaining',freeRms,freePct,'wv-occ-br-remain','wv-remain-count');
            } else {
              function brRow(clr,lbl,rms,pct,extra,rmsCls){
                return '<div class="wv-occ-br-row'+(extra?' '+extra:'')+'"><div class="wv-occ-br-left"><span class="wv-occ-br-dot" style="background:'+clr+'"></span><span class="wv-occ-br-lbl">'+lbl+'</span></div><div class="wv-occ-br-right"><span class="wv-occ-br-rms'+(rmsCls?' '+rmsCls:'')+'">'+rms+' rms</span><span class="wv-occ-br-pct">'+pct+'%</span></div></div>';
              }
              bdRows = brRow('#006461','Travel Distribution Hubs',toRms,to)
                +_opCmp(to, sdlyTo, lyTo, fcstTo, function(v){return v+'%';})
                +brRow('#47c5bc','Other Segments',otherRms,otherPct)
                +brRow('#388C3F','Remaining',freeRms,freePct,'wv-occ-br-remain','wv-remain-count');
            }

            return '<div class="wv-occ-bar-wrap">'
              +'<div class="wv-occ-bar-labels"><span class="wv-q-label">Occupancy</span>'+wvHdrRight(hotel+'%',sdlyH+'%',lyH+'%',fcstH+'%')+'</div>'
              +'<div class="wv-occ-bar-track" style="position:relative">'+barInner+occTicks()+'</div>'
              +'<div class="wv-occ-breakdown">'+bdRows+'</div>'
              +'</div>';
          })() : ''}
          ${wvMetricState.onlineOffline ? `<div><div class="wv-occ-bar-labels"><span class="wv-q-label">Online / Offline</span><span class="wv-occ-total">${onlinePct}%</span></div><div class="wv-occ-bar-track"><div style="width:${onlinePct}%;background:#3b82f6;height:7px"></div><div style="width:${100-onlinePct}%;background:#f97316;height:7px"></div></div><div class="wv-occ-breakdown"><div class="wv-occ-br-row"><div class="wv-occ-br-left"><span class="wv-occ-br-dot" style="background:#3b82f6"></span><span class="wv-occ-br-lbl">Online</span></div><div class="wv-occ-br-right"><span class="wv-occ-br-rms">${onlinePct}%</span></div></div><div class="wv-occ-br-row"><div class="wv-occ-br-left"><span class="wv-occ-br-dot" style="background:#f97316"></span><span class="wv-occ-br-lbl">Offline</span></div><div class="wv-occ-br-right"><span class="wv-occ-br-rms">${100-onlinePct}%</span></div></div></div></div>` : ''}
          ${wvMetricState.adr ? (function(){
            const toAdrV    = Math.max(80, adr - 20 - Math.abs((dm*3+dd*7)%15));
            const hotelAdrV = adr;
            const hotelAdrTick = Math.min(95, adrBar + 12);
            const diff      = hotelAdrV - toAdrV;
            const diffSign  = diff >= 0 ? '+$'+diff : '-$'+Math.abs(diff);
            const diffColor = diff >= 0 ? '#16a34a' : '#dc2626';
            var segRows = '';
            if (showH) {
              if (wvSegMode === 'individual') {
                const fitAdr = Math.round(toAdrV * 0.97), dynAdr = Math.round(toAdrV * 1.04), serAdr = Math.round(toAdrV * 0.91);
                const sdlyFitAdr = Math.round(sdlyToAdr * 0.97), lyFitAdr = Math.round(lyToAdr * 0.97), fcstFitAdr = Math.round(fcstToAdr * 0.97);
                const sdlyDynAdr = Math.round(sdlyToAdr * 1.04), lyDynAdr = Math.round(lyToAdr * 1.04), fcstDynAdr = Math.round(fcstToAdr * 1.04);
                const sdlySerAdr = Math.round(sdlyToAdr * 0.91), lySerAdr = Math.round(lyToAdr * 0.91), fcstSerAdr = Math.round(fcstToAdr * 0.91);
                segRows = '<div class="wv-occ-br-row"><div class="wv-occ-br-left"><span class="wv-occ-br-dot" style="background:#006461"></span><span class="wv-occ-br-lbl">Static FIT Rates</span></div><div class="wv-occ-br-right"><span class="wv-occ-br-rms">$'+fitAdr+'</span></div></div>'
                  +_opCmp(fitAdr, sdlyFitAdr, lyFitAdr, fcstFitAdr, function(v){return '$'+v;})
                  +'<div class="wv-occ-br-row"><div class="wv-occ-br-left"><span class="wv-occ-br-dot" style="background:#0891b2"></span><span class="wv-occ-br-lbl">TO Dynamic</span></div><div class="wv-occ-br-right"><span class="wv-occ-br-rms">$'+dynAdr+'</span></div></div>'
                  +_opCmp(dynAdr, sdlyDynAdr, lyDynAdr, fcstDynAdr, function(v){return '$'+v;})
                  +'<div class="wv-occ-br-row"><div class="wv-occ-br-left"><span class="wv-occ-br-dot" style="background:#6366f1"></span><span class="wv-occ-br-lbl">Tour Series</span></div><div class="wv-occ-br-right"><span class="wv-occ-br-rms">$'+serAdr+'</span></div></div>'
                  +_opCmp(serAdr, sdlySerAdr, lySerAdr, fcstSerAdr, function(v){return '$'+v;});
              } else {
                segRows = '<div class="wv-occ-br-row"><div class="wv-occ-br-left"><span class="wv-occ-br-dot" style="background:#94b1f5"></span><span class="wv-occ-br-lbl">TO ADR</span></div><div class="wv-occ-br-right"><span class="wv-occ-br-rms">$'+toAdrV+'</span></div></div>'
                  +_opCmp(toAdrV, sdlyToAdr, lyToAdr, fcstToAdr, function(v){return '$'+v;});
              }
              segRows += '<div class="wv-occ-br-row"><div class="wv-occ-br-left"><span class="wv-occ-br-dot" style="background:#e5e7eb"></span><span class="wv-occ-br-lbl">Difference</span></div><div class="wv-occ-br-right"><span class="wv-occ-br-rms" style="color:'+diffColor+'">'+diffSign+'</span></div></div>';
            }
            const hdrRow = '<div class="wv-occ-br-row"><div class="wv-occ-br-left"><span class="wv-occ-br-dot" style="background:#7c3aed"></span><span class="wv-occ-br-lbl">Hotel ADR</span></div><div class="wv-occ-br-right"><span class="wv-occ-br-rms">$'+hotelAdrV+'</span></div></div>';
            const htick  = showH ? '<div class="wv-bar-ref-tick" style="left:'+hotelAdrTick+'%;background:#7c3aed;width:2px;position:absolute;top:0;bottom:0"></div>' : '';
            return '<div>'
              +'<div class="wv-occ-bar-labels"><span class="wv-q-label">TO ADR</span>'+wvHdrRight('$'+toAdrV,'$'+sdlyA,'$'+(lyA??sdlyA),'$'+fcstA)+'</div>'
              +'<div class="wv-occ-bar-track" style="position:relative;overflow:visible"><div style="width:'+adrBar+'%;height:7px;background:#94b1f5;border-radius:4px;flex-shrink:0"></div>'+htick+barTicks(Math.max(3,adrBarRef-5),adrBarRef,Math.min(92,adrBarRef+5))+'</div>'
              +(showH ? '<div class="wv-occ-breakdown">'+hdrRow+segRows+'</div>' : '')
              +'</div>';
          })() : ''}
          ${wvMetricState.revenue ? (function(){
            const toAdrV2   = Math.max(80, adr - 20 - Math.abs((dm*3+dd*7)%15));
            const toRns     = Math.round(250 * to / 100);
            const toRevV    = Math.floor(toRns * toAdrV2);
            const hotelRevV = Number(rev);
            const hotelRevTick = Math.min(95, revBar + 10);
            const toShare   = hotelRevV > 0 ? Math.round(toRevV / hotelRevV * 100) : 0;
            const toRevStr  = '$'+Math.round(toRevV/1000)+'k';
            const hotRevStr = '$'+Math.round(hotelRevV/1000)+'k';
            var segRows = '';
            if (showH) {
              if (wvSegMode === 'individual') {
                const fitRev = Math.round(toRevV * 0.45), dynRev = Math.round(toRevV * 0.35), serRev = toRevV - fitRev - Math.round(toRevV*0.35);
                const fStr = fitRev>=1000 ? '$'+Math.round(fitRev/1000)+'k' : '$'+fitRev;
                const dStr = dynRev>=1000 ? '$'+Math.round(dynRev/1000)+'k' : '$'+dynRev;
                const sStr = serRev>=1000 ? '$'+Math.round(serRev/1000)+'k' : '$'+serRev;
                const sdlyFitRev = Math.round(sdlyToRev * 0.45), lyFitRev = Math.round(lyToRev * 0.45), fcstFitRev = Math.round(fcstToRev * 0.45);
                const sdlyDynRev = Math.round(sdlyToRev * 0.35), lyDynRev = Math.round(lyToRev * 0.35), fcstDynRev = Math.round(fcstToRev * 0.35);
                const sdlySerRev = sdlyToRev - sdlyFitRev - sdlyDynRev, lySerRev = lyToRev - lyFitRev - lyDynRev, fcstSerRev = fcstToRev - fcstFitRev - fcstDynRev;
                function fmtRev(v){return v>=1000?'$'+Math.round(v/1000)+'k':'$'+v;}
                segRows = '<div class="wv-occ-br-row"><div class="wv-occ-br-left"><span class="wv-occ-br-dot" style="background:#006461"></span><span class="wv-occ-br-lbl">Static FIT Rates</span></div><div class="wv-occ-br-right"><span class="wv-occ-br-rms">'+fStr+'</span></div></div>'
                  +_opCmp(fitRev, sdlyFitRev, lyFitRev, fcstFitRev, fmtRev)
                  +'<div class="wv-occ-br-row"><div class="wv-occ-br-left"><span class="wv-occ-br-dot" style="background:#0891b2"></span><span class="wv-occ-br-lbl">TO Dynamic</span></div><div class="wv-occ-br-right"><span class="wv-occ-br-rms">'+dStr+'</span></div></div>'
                  +_opCmp(dynRev, sdlyDynRev, lyDynRev, fcstDynRev, fmtRev)
                  +'<div class="wv-occ-br-row"><div class="wv-occ-br-left"><span class="wv-occ-br-dot" style="background:#6366f1"></span><span class="wv-occ-br-lbl">Tour Series</span></div><div class="wv-occ-br-right"><span class="wv-occ-br-rms">'+sStr+'</span></div></div>'
                  +_opCmp(serRev, sdlySerRev, lySerRev, fcstSerRev, fmtRev);
              } else {
                segRows = '<div class="wv-occ-br-row"><div class="wv-occ-br-left"><span class="wv-occ-br-dot" style="background:#eba2a2"></span><span class="wv-occ-br-lbl">TO Revenue</span></div><div class="wv-occ-br-right"><span class="wv-occ-br-rms">'+toRevStr+'</span></div></div>'
                  +_opCmp(toRevV, sdlyToRev, lyToRev, fcstToRev, function(v){return '$'+Math.round(v/1000)+'k';});
              }
            }
            const hotRevRow = '<div class="wv-occ-br-row"><div class="wv-occ-br-left"><span class="wv-occ-br-dot" style="background:#ea580c"></span><span class="wv-occ-br-lbl">Hotel Revenue</span></div><div class="wv-occ-br-right"><span class="wv-occ-br-rms">'+hotRevStr+'</span></div></div>';
            const htick     = showH ? '<div class="wv-bar-ref-tick" style="left:'+hotelRevTick+'%;background:#ea580c;width:2px;position:absolute;top:0;bottom:0"></div>' : '';
            return '<div>'
              +'<div class="wv-occ-bar-labels"><span class="wv-q-label">TO Revenue</span>'+wvHdrRight(toRevStr,'$'+Math.floor(sdlyR/1000)+'k','$'+Math.floor((lyR??sdlyR)/1000)+'k','$'+Math.floor(fcstR/1000)+'k')+'</div>'
              +'<div class="wv-occ-bar-track" style="position:relative;overflow:visible"><div style="width:'+revBar+'%;height:7px;background:#eba2a2;border-radius:4px;flex-shrink:0"></div>'+htick+barTicks(Math.max(3,revBarRef-5),revBarRef,Math.min(92,revBarRef+5))+'</div>'
              +(showH ? '<div class="wv-occ-breakdown">'+hotRevRow+segRows+'</div>' : '')
              +'</div>';
          })() : ''}
        </div>`;
      })()) : '' : ''}
      ${wvGroupBy === 'combined' ? ['dm_rnSold','dm_pickup_0','dm_pickup_1','dm_pickup_2','dm_avgAdults','dm_avgChildren','dm_totalAdults','dm_totalChildren','dm_trevpar','dm_availRooms','dm_availGuar'].some(function(k){return wvMetricState[k];}) ? wvAcc('More Metrics', 'detailed', (function(){
          const showS = wvMetricState.cmp_sdly, showL = wvMetricState.cmp_final_ly, showF = wvMetricState.cmp_forecast, showH = wvMetricState.cmp_hotel;
          const availRooms = Math.max(0, 102 - Math.floor(hotel * 1.02));
          // All three reference sets
          const S = { rn: Math.floor(hotel*0.82), adr: adr-8, rev: Math.floor(rev*0.9/1000), pkup: '+0', adA: '1.9', adC: '0.4', trev: Math.floor(adr*0.92), avR: availRooms+3, avG: Math.floor(8+v%5)+2 };
          const L = { rn: Math.floor(hotel*0.90), adr: adr-4, rev: Math.floor(rev*0.95/1000), pkup: '+2', adA: '1.95', adC: '0.35', trev: Math.floor(adr*0.96), avR: availRooms+2, avG: Math.floor(8+v%5)+1 };
          const F = { rn: Math.min(102,Math.floor(hotel*1.1)+3), adr: adr+6, rev: Math.floor(rev*1.06/1000), pkup: '+'+Math.floor(v%10+8), adA: '2.0', adC: '0.4', trev: Math.floor(adr*1.08), avR: availRooms-1, avG: Math.floor(8+v%5)-1 };
          // Hotel indicator values (higher than TO)
          const dmToAdr = Math.max(80, adr - 20 - Math.abs((dm*3+dd*7)%15));
          const dmToRn  = Math.round(HOTEL_CAPACITY * to / 100);
          const toFrac   = hotel > 0 ? to / hotel : 0; // fraction of hotel rooms that are TO
          const dmHotelRn = Math.round(HOTEL_CAPACITY * hotel / 100);
          const dmToRev = Math.floor(dmToRn * dmToAdr);
          const dmToPickup = Math.max(0, Math.floor((v%25+5) * to / Math.max(1, hotel)));
          const dmHotelPickup = Math.floor(v%25+5);
          const dmToTrev = Math.max(50, (adr+80) - 30 - Math.abs((dm*5+dd*3)%20));
          function dmRefRow(sv, lv, fv, hv) {
            const parts = [];
            if (showS && sv != null) parts.push('<span class="wv-ref-tag wv-ref-sdly">STLY '+sv+'</span>');
            if (showL && lv != null) parts.push('<span class="wv-ref-tag wv-ref-ly">LY '+lv+'</span>');
            if (showF && fv != null) parts.push('<span class="wv-ref-tag wv-ref-fcst">Fcst '+fv+'</span>');
            // Hotel chip removed — Hotel shown in breakdown rows below
            return parts.length ? '<div class="wv-ref-row">'+parts.join('')+'</div>' : '';
          }
          function dmTicks(bp, hbp) {
            const sp = Math.max(3, bp-15), lp = Math.max(3, bp-10), fp = Math.min(92, bp+5);
            return (showS?'<div class="wv-dm-sdly-mark wv-tick-sdly" style="left:'+sp+'%"></div>':'')
                 + (showL?'<div class="wv-dm-sdly-mark wv-tick-ly" style="left:'+lp+'%"></div>':'')
                 + (showF?'<div class="wv-dm-sdly-mark wv-tick-fcst" style="left:'+fp+'%"></div>':'')
                 // Hotel tick removed — shown in breakdown rows
          }
          // rows: [lbl, toVal, sv, lv, fv, barClr, barPct, key, hotelVal, hotelBarPct]
          const dmAvgAdults   = (1.8+v%3*.1).toFixed(1);
          const dmAvgChildren = (0.3+v%2*.1).toFixed(1);
          const dmTotalAdults   = Math.round(dmToRn * parseFloat(dmAvgAdults));
          const dmTotalChildren = Math.round(dmToRn * parseFloat(dmAvgChildren));
          // Hotel-level values for guest/LOS/lead time metrics
          const hotelAvgAdults   = (parseFloat(dmAvgAdults)   + 0.3).toFixed(1); // hotel slightly higher
          const hotelAvgChildren = (parseFloat(dmAvgChildren) + 0.1).toFixed(1);
          const hotelTotalAdults   = Math.round(dmHotelRn * parseFloat(hotelAvgAdults));
          const hotelTotalChildren = Math.round(dmHotelRn * parseFloat(hotelAvgChildren));
          const hotelAvgLos      = ((2.8+v%5*.3) + 0.4).toFixed(1)+'n';
          const hotelAvgLeadTime = (18+v%60+12)+'d';
          const dmTotalGuestsT   = Math.round(dmToRn * (parseFloat(dmAvgAdults) + parseFloat(dmAvgChildren)));
          const hotelTotalGuests = Math.round(dmHotelRn * (parseFloat(hotelAvgAdults) + parseFloat(hotelAvgChildren)));
          return [
            ['RN Sold',        dmToRn,         S.rn,  L.rn,  F.rn,  '#2e65e8', Math.min(92, 55+(v%37)),             'dm_rnSold',       dmHotelRn,         Math.min(92, 55+(v%37)+10)],
            {__type:'pickup_group', __dmToPickup:dmToPickup, __dmHotelPickup:dmHotelPickup, __toFrac:toFrac, __v:v},
            ['Avg Adults',     dmAvgAdults,    null,  null,  null,  '#2e65e8', Math.min(92, 55+v%30),               'dm_avgAdults',    hotelAvgAdults,    Math.min(92, 55+v%30+8)],
            ['Avg Children',   dmAvgChildren,  null,  null,  null,  '#d33030', Math.min(92, 20+v%40),               'dm_avgChildren',  hotelAvgChildren,  Math.min(92, 20+v%40+8)],
            ['Total Adults',   dmTotalAdults,  null,  null,  null,  '#2e65e8', Math.min(92, 60+v%28),               'dm_totalAdults',  hotelTotalAdults,  Math.min(92, 60+v%28+8)],
            ['Total Children', dmTotalChildren,null,  null,  null,  '#d33030', Math.min(92, 15+v%35),               'dm_totalChildren',hotelTotalChildren,Math.min(92, 15+v%35+8)],
            ['REVPAR',         '$'+dmToTrev,   '$'+S.trev, '$'+L.trev, null,  '#2e65e8', Math.min(92, 65+v%25),    'dm_trevpar',      '$'+(adr+80),      Math.min(92, 65+v%25+10)],
            ['Avail Rooms',    availRooms,     null,  null,  null,  '#16a34a', Math.min(92, Math.max(5, hotel*0.8)),'dm_availRooms',   '__hotelOnly',     null],
            ['Avail Guar.',    Math.floor(8+v%5), null,null, null,  '#2e65e8', Math.min(92, 10+v%50),               'dm_availGuar',    null,              null],
            ['Avg LOS',        (2.8+v%5*.3).toFixed(1)+'n', null,null,null,'#0891b2', Math.min(92, 40+v%40), 'dm_avgLos',       hotelAvgLos,       Math.min(92, 40+v%40+8)],
            ['Avg Lead Time',  (18+v%60)+'d',               null,null,null,'#6366f1', Math.min(92, 25+v%55), 'dm_avgLeadTime',  hotelAvgLeadTime,  Math.min(92, 25+v%55+8)],
            ['Total Guests',   dmTotalGuestsT, null,null,null,'#0369a1', Math.min(92, 55+v%35),               'dm_totalGuests',  hotelTotalGuests,  Math.min(92, 55+v%35+8)],
          ].filter(function(row){
            if (row.__type==='pickup_group') return wvMetricState.dm_pickup;
            return wvMetricState[row[7]];
          }).map(function(row){
            // ── Grouped pickup cells (one per active window) ─────────────
            if (row.__type === 'pickup_group') {
              return pickupDayValues.map(function(dv, i) {
                if (!wvMetricState['dm_pickup_' + i]) return '';
                var scale = dv<=1?0.3:dv<=3?0.6:dv<=7?1:Math.min(2,dv/7);
                var toP  = Math.max(0, Math.round(row.__dmToPickup * scale));
                var htlP = Math.max(0, Math.round(row.__dmHotelPickup * scale));
                var bp   = Math.min(92, 30 + row.__v % 50);
                var tPct = Math.round(bp * Math.min(1, row.__toFrac));
                var hPct = Math.round(bp * 1.1);
                var dualBar = '<div class="wv-dm-bar-wrap" style="position:relative">'
                  +'<div class="wv-dm-bar-fill" style="width:'+hPct+'%;background:#006461;opacity:0.2"></div>'
                  +'<div class="wv-dm-bar-fill" style="width:'+tPct+'%;background:#006461"></div>'
                  +'</div>';
                var bdRows = '<div class="wv-occ-br-row"><div class="wv-occ-br-left"><span class="wv-occ-br-dot" style="background:#006461;opacity:.45"></span><span class="wv-occ-br-lbl">Hotel</span></div><div class="wv-occ-br-right"><span class="wv-occ-br-rms">+'+htlP+'</span></div></div>'
                  +'<div class="wv-occ-br-row"><div class="wv-occ-br-left"><span class="wv-occ-br-dot" style="background:#006461"></span><span class="wv-occ-br-lbl" style="color:#006461">TO</span></div><div class="wv-occ-br-right"><span class="wv-occ-br-rms" style="color:#006461">+'+toP+'</span></div></div>';
                return '<div>'
                  +'<div class="wv-occ-bar-labels"><span class="wv-q-label">Pickup '+dv+'</span>'
                  +'<div class="wv-hdr-right"><span class="wv-occ-total" style="color:#006461">+'+toP+'</span></div></div>'
                  +dualBar
                  +'<div class="wv-occ-breakdown" style="margin-top:2px">'+bdRows+'</div>'
                  +'</div>';
              }).join('');
            }
            const lbl=row[0],val=row[1],sv=row[2],lv=row[3],fv=row[4],barClr=row[5],barPct=row[6],hv=row[8],hbp=row[9];
            const isHotelOnly = hv === '__hotelOnly';
            const hvDisplay   = (hv && hv !== '__hotelOnly') ? hv : null;
            // Bar: Hotel faded full + T teal solid
            const tBarPct  = isHotelOnly ? 0 : Math.round(barPct * Math.min(1, toFrac));
            const hBarPct2 = hbp != null ? hbp : barPct;
            const dualBar  = '<div class="wv-dm-bar-wrap" style="position:relative">'
              +'<div class="wv-dm-bar-fill" style="width:'+hBarPct2+'%;background:'+barClr+(isHotelOnly?'':';opacity:0.25')+'"></div>'
              +(isHotelOnly ? '' : '<div class="wv-dm-bar-fill" style="width:'+tBarPct+'%;background:#006461"></div>')
              +dmTicks(barPct,hbp)
              +'</div>';
            // Breakdown rows
            var bdRows = '';
            if (isHotelOnly) {
              bdRows = '<div class="wv-occ-br-row"><div class="wv-occ-br-left"><span class="wv-occ-br-dot" style="background:'+barClr+'"></span><span class="wv-occ-br-lbl">Hotel</span></div><div class="wv-occ-br-right"><span class="wv-occ-br-rms">'+val+'</span></div></div>';
            } else {
              if (hvDisplay != null) {
                bdRows += '<div class="wv-occ-br-row"><div class="wv-occ-br-left"><span class="wv-occ-br-dot" style="background:'+barClr+';opacity:.45"></span><span class="wv-occ-br-lbl">Hotel</span></div><div class="wv-occ-br-right"><span class="wv-occ-br-rms">'+hvDisplay+'</span></div></div>';
              }
              const canIndiv = (row[7]==='dm_rnSold'||row[7]==='dm_pickup'||row[7]==='dm_trevpar'||row[7]==='dm_avgAdults'||row[7]==='dm_avgChildren'||row[7]==='dm_totalAdults'||row[7]==='dm_totalChildren'||row[7]==='dm_avgLos'||row[7]==='dm_avgLeadTime'||row[7]==='dm_totalGuests');
              if (wvSegMode === 'individual' && canIndiv) {
                const rawVal = parseFloat(String(val).replace(/[^0-9.-]/g,'')) || 0;
                const isPickup = row[7]==='dm_pickup', isRev = row[7]==='dm_trevpar',
                      isLos = row[7]==='dm_avgLos', isLead = row[7]==='dm_avgLeadTime';
                const fmtSeg = function(f){ return isRev ? '$'+Math.round(rawVal*f) : isLos ? (rawVal*f).toFixed(1)+'n' : isLead ? Math.round(rawVal*f)+'d' : isPickup ? (rawVal*f>=0?'+':'')+Math.round(rawVal*f) : String(Math.round(rawVal*f)); };
                bdRows += '<div class="wv-occ-br-row"><div class="wv-occ-br-left"><span class="wv-occ-br-dot" style="background:#006461"></span><span class="wv-occ-br-lbl" style="color:#006461">Static FIT</span></div><div class="wv-occ-br-right"><span class="wv-occ-br-rms" style="color:#006461">'+fmtSeg(0.45)+'</span></div></div>'
                  +'<div class="wv-occ-br-row"><div class="wv-occ-br-left"><span class="wv-occ-br-dot" style="background:#0891b2"></span><span class="wv-occ-br-lbl" style="color:#0891b2">TO Dynamic</span></div><div class="wv-occ-br-right"><span class="wv-occ-br-rms" style="color:#0891b2">'+fmtSeg(0.35)+'</span></div></div>'
                  +'<div class="wv-occ-br-row"><div class="wv-occ-br-left"><span class="wv-occ-br-dot" style="background:#6366f1"></span><span class="wv-occ-br-lbl" style="color:#6366f1">Tour Series</span></div><div class="wv-occ-br-right"><span class="wv-occ-br-rms" style="color:#6366f1">'+fmtSeg(0.20)+'</span></div></div>';
              } else {
                bdRows += '<div class="wv-occ-br-row"><div class="wv-occ-br-left"><span class="wv-occ-br-dot" style="background:#006461"></span><span class="wv-occ-br-lbl" style="color:#006461">TO</span></div><div class="wv-occ-br-right"><span class="wv-occ-br-rms" style="color:#006461">'+val+'</span></div></div>';
              }
            }
            const headerClr = isHotelOnly ? '#181d1f' : '#006461';
            const dmHdrRight = (sv != null || lv != null || fv != null)
              ? wvHdrRight(val, sv != null ? String(sv) : null, lv != null ? String(lv) : null, fv != null ? String(fv) : null)
              : '<div class="wv-hdr-right"><span class="wv-occ-total" style="color:'+headerClr+'">'+val+'</span></div>';
            return '<div>'
              +'<div class="wv-occ-bar-labels"><span class="wv-q-label">'+lbl+'</span>'+dmHdrRight+'</div>'
              +dualBar
              +'<div class="wv-occ-breakdown" style="margin-top:2px">'+bdRows+'</div>'
              +'</div>';
          }).join('');
        })())  : '' : ''}
      ${wvGroupBy === 'combined' ? wvAcc('Meal Plans', 'meals', (function(){
          const aiPct  = Math.max(45, Math.min(68, 55 + (dm*7+dd*3)%14));
          const bbPct  = Math.max(14, Math.min(28, 20 + (dm*11+dd*5)%11));
          const hbPct  = Math.max(6,  Math.min(16, 10 + (dm*5+dd*7)%9));
          const roPct  = Math.max(2,  100 - aiPct - bbPct - hbPct);
          const totalRooms = Math.floor(hotel * 1.1);
          // TO meal plan percentages (subset of hotel)
          const toPct = to / Math.max(1, hotel); // fraction of rooms that are TO
          const toAiPct  = Math.round(aiPct  * toPct * (0.9 + (dm+dd)%3 * 0.05));
          const toBbPct  = Math.round(bbPct  * toPct * (0.85 + (dm*3+dd)%3 * 0.05));
          const toHbPct  = Math.round(hbPct  * toPct * (0.8  + (dm+dd*2)%3 * 0.05));
          const toRoPct  = Math.round(roPct  * toPct * (0.95 + (dm*2+dd)%3 * 0.03));
          const plans = [
            { name:'All Inclusive',   short:'AI', pct:aiPct, toPct:toAiPct,  color:'#004948' },
            { name:'Bed & Breakfast', short:'BB', pct:bbPct, toPct:toBbPct,  color:'#52d9ce' },
            { name:'Half Board',      short:'HB', pct:hbPct, toPct:toHbPct,  color:'#C4FF45' },
            { name:'Room Only',       short:'RO', pct:roPct, toPct:toRoPct,  color:'#D97706' },
          ];
          const colHdr = '<div style="display:flex;justify-content:flex-end;gap:10px;padding:1px 8px 0;margin-bottom:-2px">'
            +'<span style="font-size:7px;font-weight:700;color:#6b7280;text-transform:uppercase;letter-spacing:.3px;width:30px;text-align:right">Hotel</span>'
            +'<span style="font-size:7px;font-weight:700;color:#006461;text-transform:uppercase;letter-spacing:.3px;width:30px;text-align:right">TO</span>'
            +'</div>';
          const barHtml = '<div class="wv-meals-bar">'
            + plans.map(p=>'<div style="width:'+p.pct+'%;background:'+p.color+';height:100%"></div>').join('')
            + '</div>';
          // Per-plan base values
          const baseAdr     = adr;
          const baseAdrNet  = Math.round(adr * 0.88);
          const toAdrGross  = Math.round(adr * 0.82);
          const toAdrNet    = Math.round(toAdrGross * 0.91);
          const avgAdultsV  = 1.8 + (dm*11+dd*7)%3 * 0.1;
          const avgChildrenV= 0.3 + (dm*7+dd*13)%5 * 0.1;

          const rowsHtml = plans.map(function(p){
            const totalPlanRooms = Math.round(totalRooms * p.pct / 100);
            const toRoomsAmt     = Math.round(totalRooms * p.toPct / 100);
            const hotelOnlyRooms = Math.max(0, totalPlanRooms - toRoomsAmt);
            const hAdults   = Math.round(totalPlanRooms * avgAdultsV);
            const hChildren = Math.round(totalPlanRooms * avgChildrenV);
            const hGuests   = hAdults + hChildren;
            const hRev      = Math.round(totalPlanRooms * baseAdr);
            const hRevStr   = hRev >= 1000 ? '$'+Math.round(hRev/1000)+'k' : '$'+hRev;
            const tAdults   = Math.round(toRoomsAmt * avgAdultsV);
            const tChildren = Math.round(toRoomsAmt * avgChildrenV);
            const tGuests   = tAdults + tChildren;
            const tRev      = Math.round(toRoomsAmt * toAdrGross);
            const tRevStr   = tRev >= 1000 ? '$'+Math.round(tRev/1000)+'k' : '$'+tRev;
            const toShare   = totalPlanRooms > 0 ? Math.round(toRoomsAmt / totalPlanRooms * 100) : 0;
            // Summary row (occ-br style: dot | name | pct | hotel rm | TO rm)
            const summaryRow = '<div class="wv-occ-br-row" style="grid-template-columns:8px 1fr 28px 30px 30px;padding:3px 8px">'
              +'<span class="wv-occ-br-dot" style="background:'+p.color+'"></span>'
              +'<span class="wv-occ-br-lbl" style="font-weight:700">'+p.short+' <span style="font-weight:400;color:#9ca3af">'+p.name+'</span></span>'
              +'<span class="wv-occ-br-pct">'+p.pct+'%</span>'
              +'<span class="wv-occ-br-rms" title="Hotel rooms">'+totalPlanRooms+'</span>'
              +'<span class="wv-occ-br-rms" style="color:#006461" title="TO rooms">'+toRoomsAmt+'</span>'
              +'</div>';
            // Detail rows (dm-row style)
            function dmRow(lbl, hVal, tVal) {
              return '<div class="wv-dm-row" style="padding:0 8px 0 20px">'
                +'<div class="wv-dm-top">'
                +'<span class="wv-dm-label">'+lbl+'</span>'
                +'<span style="display:flex;gap:10px">'
                +'<span class="wv-dm-val" style="color:#374151">'+hVal+'</span>'
                +'<span class="wv-dm-val" style="color:#006461">'+tVal+'</span>'
                +'</span>'
                +'</div>'
                +'</div>';
            }
            const detailRows =
              dmRow('Rooms', totalPlanRooms+' rm', toRoomsAmt+' rm')
              +dmRow('Adults', hAdults, tAdults)
              +dmRow('Children', hChildren, tChildren)
              +dmRow('Guests', hGuests, tGuests)
              +dmRow('Revenue', hRevStr, tRevStr)
              +dmRow('ADR Gross', '$'+baseAdr, '$'+toAdrGross)
              +dmRow('ADR Net', '$'+baseAdrNet, '$'+toAdrNet)
              +dmRow('TO Share', toShare+'%', (100-toShare)+'% hotel');
            return summaryRow + detailRows;
          }).join('');
          return barHtml + colHdr + rowsHtml;
        })()) : ''}
      ${wvGroupBy === 'combined' ? wvMetricState.mealsSummary ? (function(){
        const aiPct2 = Math.max(45, Math.min(68, 55 + (dm*7+dd*3)%14));
        const bbPct2 = Math.max(14, Math.min(28, 20 + (dm*11+dd*5)%11));
        const hbPct2 = Math.max(6,  Math.min(16, 10 + (dm*5+dd*7)%9));
        const roPct2 = Math.max(2,  100 - aiPct2 - bbPct2 - hbPct2);
        const totalRooms2 = Math.floor(hotel * 1.1);
        const toPct2 = to / Math.max(1, hotel);
        const plans2 = [
          { name:'All Inclusive',   short:'AI', pct:aiPct2, toPct: Math.round(aiPct2 * toPct2 * (0.9  + (dm+dd)%3 * 0.05)), color:'#004948' },
          { name:'Bed & Breakfast', short:'BB', pct:bbPct2, toPct: Math.round(bbPct2 * toPct2 * (0.85 + (dm*3+dd)%3 * 0.05)), color:'#52d9ce' },
          { name:'Half Board',      short:'HB', pct:hbPct2, toPct: Math.round(hbPct2 * toPct2 * (0.8  + (dm+dd*2)%3 * 0.05)), color:'#C4FF45' },
          { name:'Room Only',       short:'RO', pct:roPct2, toPct: Math.round(roPct2 * toPct2 * (0.95 + (dm*2+dd)%3 * 0.03)), color:'#D97706' },
        ];
        const avgAdV2   = 1.8 + (dm*11+dd*7)%3 * 0.1;
        const avgChV2   = 0.3 + (dm*7+dd*13)%5 * 0.1;
        const rows2 = plans2.map(function(p){
          const planRooms = Math.round(totalRooms2 * p.pct / 100);
          const toRooms   = Math.round(totalRooms2 * p.toPct / 100);
          const guests    = Math.round(planRooms * (avgAdV2 + avgChV2));
          const toGuestPct= planRooms > 0 ? Math.round(toRooms / planRooms * 100) : 0;
          return '<div class="wv-occ-br-row" style="grid-template-columns:8px 1fr 28px 40px;padding:2px 8px">'            +'<span class="wv-occ-br-dot" style="background:'+p.color+'"></span>'            +'<span class="wv-occ-br-lbl" style="font-weight:700">'+p.short+'</span>'            +'<span class="wv-occ-br-pct">'+p.pct+'%</span>'            +'<span class="wv-occ-br-rms" style="color:#374151">'+guests+' TG</span>'            +'</div>';
        }).join('');
        const hdr2 = '<div style="display:flex;justify-content:flex-end;gap:6px;padding:1px 8px 2px">'          +'<span style="font-size:7px;font-weight:700;color:#9ca3af;text-transform:uppercase;letter-spacing:.3px;width:40px;text-align:right">TG</span>'          +'</div>';
        return wvAcc('Meal Plans Summary', 'mealsSummary', hdr2 + rows2);
      })() : '' : ''}
      ${wvGroupBy === 'combined' ? (wvMetricState.avail || wvMetricState.availAlloc) ? wvAcc('Room Availability', 'avail', (function(){
          const RT = [['Standard',51],['Superior',36],['Deluxe',27],['Suite',12],['Jr. Suite',15],['Family',9]];
          const totalCap = RT.reduce(function(s,r){ return s+r[1]; }, 0);
          // Per room type calculations
          const rtData = RT.map(function(r, i){
            const inv = r[1];
            const totalSoldRt = Math.min(inv, Math.floor(inv * hotel / 110));
            const toSoldRt    = Math.min(totalSoldRt, Math.round(totalSoldRt * to / Math.max(1, hotel)));
            const otherSoldRt = totalSoldRt - toSoldRt;
            const toAlloc     = Math.floor(inv * 0.8 + Math.abs((dm*(i+3)+dd*(i+5))%15));
            const toAllocRem  = Math.max(0, toAlloc - toSoldRt);
            const avail       = Math.max(0, inv - totalSoldRt);
            return { inv, totalSoldRt, toSoldRt, otherSoldRt, toAlloc, toAllocRem, avail };
          });
          const totalSold      = rtData.reduce(function(s,d){ return s+d.totalSoldRt; }, 0);
          const totalToSold    = rtData.reduce(function(s,d){ return s+d.toSoldRt; }, 0);
          const totalOtherSold = rtData.reduce(function(s,d){ return s+d.otherSoldRt; }, 0);
          const totalAvail     = totalCap - totalSold;
          const totalToAllocRem= rtData.reduce(function(s,d){ return s+d.toAllocRem; }, 0);
          // Stacked capacity bar
          const toSoldPct    = Math.round(totalToSold    / totalCap * 100);
          const otherSoldPct = Math.round(totalOtherSold / totalCap * 100);
          const toAllocPct   = Math.round(totalToAllocRem/ totalCap * 100);
          const availPct     = Math.max(0, 100 - toSoldPct - otherSoldPct - toAllocPct);
          const capBar = '<div class="wv-cap-bar-wrap">'
            +'<div class="wv-cap-bar">'
            +'<div style="width:'+toSoldPct+'%;background:#006461;height:100%" title="TO Sold"></div>'
            +'<div style="width:'+otherSoldPct+'%;background:#3b82f6;height:100%" title="Other Sold"></div>'
            +'<div style="width:'+toAllocPct+'%;background:#fb923c;opacity:0.6;height:100%" title="T Alloc Remaining"></div>'
            +'<div style="width:'+availPct+'%;background:#d1fae5;height:100%" title="Available"></div>'
            +'</div>'
            +'<div class="wv-cap-legend">'
            +'<span class="wv-cap-leg-item"><span class="wv-cap-leg-dot" style="background:#006461"></span>TO Sold<b>'+totalToSold+'</b></span>'
            +'<span class="wv-cap-leg-item"><span class="wv-cap-leg-dot" style="background:#3b82f6"></span>Other <b>'+totalOtherSold+'</b></span>'
            +'<span class="wv-cap-leg-item"><span class="wv-cap-leg-dot" style="background:#fb923c"></span>T Alloc Rem. <b>'+totalToAllocRem+'</b></span>'
            +'<span class="wv-cap-leg-item"><span class="wv-cap-leg-dot" style="background:#16a34a"></span>Avail <b>'+totalAvail+'</b></span>'
            +'</div>'
            +'<div class="wv-cap-total">Capacity: <b>'+totalCap+' rooms</b> · '+Math.round(totalSold/totalCap*100)+'% occupied</div>'
            +'</div>';
          const tblHdr = '<div class="wv-cap-tbl-hdr">'
            +'<span class="wv-cap-th-type">Room Type</span>'
            +'<span class="wv-cap-th">Cap</span>'
            +'<span class="wv-cap-th" style="color:#006461">TO</span>'
            +'<span class="wv-cap-th" style="color:#3b82f6">Other</span>'
            +'<span class="wv-cap-th" style="color:#fb923c">Alloc↑</span>'
            +'<span class="wv-cap-th" style="color:#16a34a">Avail</span>'
            +'</div>';
          const rows = RT.map(function(r, i){
            const d = rtData[i];
            const availClr = d.avail === 0 ? '#ef4444' : '#16a34a';
            const toSoldPctRt    = d.inv > 0 ? Math.round(d.toSoldRt    / d.inv * 100) : 0;
            const otherSoldPctRt = d.inv > 0 ? Math.round(d.otherSoldRt / d.inv * 100) : 0;
            const toAllocPctRt   = d.inv > 0 ? Math.round(d.toAllocRem  / d.inv * 100) : 0;
            const availPctRt     = Math.max(0, 100 - toSoldPctRt - otherSoldPctRt - toAllocPctRt);
            return '<div class="wv-cap-rt-row">'
              +'<div class="wv-cap-rt-name">'
              +'<span class="wv-cap-rt-sw" style="background:'+RT_COLORS[i]+'"></span>'
              +'<span class="wv-cap-rt-lbl">'+r[0]+(d.avail===0?' <span class="wv-rt-closed-badge">CLOSED</span>':'')+'</span>'
              +'</div>'
              +'<span class="wv-cap-td">'+d.inv+'</span>'
              +'<span class="wv-cap-td" style="color:#006461">'+d.toSoldRt+'</span>'
              +'<span class="wv-cap-td" style="color:#3b82f6">'+d.otherSoldRt+'</span>'
              +'<span class="wv-cap-td" style="color:#fb923c">'+d.toAllocRem+'</span>'
              +'<span class="wv-cap-td" style="color:'+availClr+'">'+d.avail+'</span>'
              +'</div>';
          }).join('');
          return capBar + tblHdr + rows;
        })()) : '' : ''}
      ${wvGroupBy === 'combined' ? wvMetricState.toRates ? (function(){
        const seed = (dm * 37 + dd * 17) % 100;
        // EBB 10% for first 3 days of week (Sun/Mon/Tue), Contract for other 4
        const dayOfWeekD = (new Date(2026, dm-1, dd)).getDay();
        const isEbbDay   = dayOfWeekD < 3;
        const allPromos  = isEbbDay
          ? [{ name:'Early Bird 10%', type:'EBB 10%', discount:10, color:'#16a34a' }]
          : [{ name:'Contract Rate',  type:'Contract',discount:0,  color:'#2563eb' }];
        const toOperators = [
          { name:'Sunshine Tours', color:'#3b82f6' },
          { name:'Global Adv.',    color:'#967EF3' },
          { name:'Beach Hols',     color:'#0ea5e9' },
          { name:'City Breaks',    color:'#10b981' },
          { name:'Adventure',      color:'#f59e0b' },
        ];
        const baseSegRate = adr + 8;  // Base Segment Selling Rate (higher than TO rates)
        const rows = toOperators.map(function(op, i) {
          const toRate  = adr - 15 + Math.abs((dm*(i+3) + dd*(i+5)) % 50);
          const toAllot = 5  + Math.abs((dm*(i+2) + dd*(i+3)) % 20);
          const toUsed  = Math.max(0, toAllot - Math.floor(hotel / 20));
          const barPct  = Math.round((toUsed / toAllot) * 100);
          const barCls  = barPct >= 90 ? 'wv-to-bar-high' : barPct >= 60 ? 'wv-to-bar-mid' : 'wv-to-bar-low';
          // All operators get same promo (EBB or Contract based on day of week)
          const hasPromo = true;
          const promo = allPromos[0];
          const promoTag = hasPromo
            ? `<span class="wv-to-promo-tag" style="background:${promo.color}" data-tooltip="${promo.name}${promo.discount>0?' (−'+promo.discount+'%)':''}">${promo.type}</span>`
            : `<span class="wv-to-promo-none">—</span>`;
          const tooltipText = hasPromo ? `Promo: ${promo.name}` : '';
          return `<div class="wv-to-rate-row" title="${tooltipText}">
            <span class="wv-to-dot" style="background:${op.color}"></span>
            <span class="wv-to-name">${op.name}</span>
            ${promoTag}
            <span class="wv-to-rate">$${toRate}</span>
            <span class="wv-to-allot">${toAllot - toUsed}r</span>
          </div>`;
        }).join('');
        const baseRateLine = '<div class="wv-to-rate-row" style="border-top:1px solid #e5e7eb;margin-top:4px;padding-top:4px">'
          +'<span class="wv-to-dot" style="background:#9333ea"></span>'
          +'<span class="wv-to-name" style="font-weight:700;color:#374151">Base Segment Rate</span>'
          +'<span style="flex:1"></span>'
          +'<span class="wv-to-rate" style="font-weight:700;color:#9333ea">$'+baseSegRate+'</span>'
          +'</div>';
        return wvAcc('Travel Company Rates', 'toRates', rows + baseRateLine);
      })() : '' : ''}
      ${wvGroupBy === 'combined' ? wvMetricState.bizMix ? (function(){
        const toMix    = 28 + Math.abs((dm*7+dd*5)%25);
        const directMix= 30 + Math.abs((dm*5+dd*9)%20);
        const otaMix   = 20 + Math.abs((dm*9+dd*3)%18);
        const otherMix = Math.max(0, 100 - toMix - directMix - otaMix);
        const segments = [
          { name:'Operator', short:'TO',     pct: toMix,    color:'#006461' },
          { name:'Direct',        short:'Direct', pct: directMix,color:'#0284c7' },
          { name:'OTA',           short:'OTA',    pct: otaMix,   color:'#D97706' },
          { name:'Other',         short:'Other',  pct: otherMix, color:'#9ca3af' },
        ];
        const barHtml = '<div class="wv-meals-bar">'
          + segments.map(s=>'<div style="width:'+s.pct+'%;background:'+s.color+';height:100%"></div>').join('')
          + '</div>';
        const rowsHtml = segments.map(function(s){
          // Individual mode: replace TO row with 3 sub-segments; all other rows unchanged
          if (wvSegMode === 'individual' && s.short === 'TO') {
            const fitP2 = Math.round(s.pct * 0.45), dynP2 = Math.round(s.pct * 0.35), serP2 = s.pct - fitP2 - dynP2;
            return '<div class="wv-meal-row"><span class="wv-meal-dot" style="background:#006461"></span><span class="wv-meal-name">Static FIT</span><span class="wv-meal-pct">'+fitP2+'%</span></div>'
              +'<div class="wv-meal-row"><span class="wv-meal-dot" style="background:#0891b2"></span><span class="wv-meal-name">TO Dynamic</span><span class="wv-meal-pct">'+dynP2+'%</span></div>'
              +'<div class="wv-meal-row"><span class="wv-meal-dot" style="background:#6366f1"></span><span class="wv-meal-name">Tour Series</span><span class="wv-meal-pct">'+serP2+'%</span></div>';
          }
          return '<div class="wv-meal-row"><span class="wv-meal-dot" style="background:'+s.color+'"></span><span class="wv-meal-name">'+s.short+'</span><span class="wv-meal-pct">'+s.pct+'%</span></div>';
        }).join('');
        return wvAcc('Business Mix', 'bizMix', barHtml + rowsHtml);
      })() : '' : ''}
      </div>
    </div>`;
  }).join('');

  if (wvGroupBy === 'report' || wvGroupBy === 'coReport' || wvGroupBy === 'dailyH' || wvGroupBy === 'dailyB') return;

  // Apply custom section order for combined view
  if (wvGroupBy === 'combined' && _wvSectionOrder) applyWvSectionOrder(grid);

  // ── Equalize section body heights across day columns ──────────────────────
  // Returns a map of section → max body height (for panel sync)
  var sectionHeightMap = {};
  (function equalizeWvBodies() {
    var sections = ['daily','detailed','meals','mealsSummary','avail','toRates','bizMix'];
    if (wvGroupBy === 'roomType') sections = ['rt_0','rt_1','rt_2','rt_3','rt_4','rt_5'];
    sections.forEach(function(sec) {
      var bodies = [];
      grid.querySelectorAll('.wv-acc-hdr[data-section="' + sec + '"]').forEach(function(hdr) {
        var b = hdr.nextElementSibling;
        if (b && b.classList.contains('wv-acc-body')) bodies.push(b);
      });
      bodies.forEach(function(b) { b.style.minHeight = ''; });
      var maxH = 0;
      bodies.forEach(function(b) {
        if (!b.classList.contains('wv-body-hidden') && b.scrollHeight > maxH) maxH = b.scrollHeight;
      });
      if (maxH > 0) {
        bodies.forEach(function(b) {
          if (!b.classList.contains('wv-body-hidden')) b.style.minHeight = maxH + 'px';
        });
        sectionHeightMap[sec] = maxH;
      }
    });
  })();

  // ── Build left section panel (combined mode only) ─────────────────────────
  var panel = document.getElementById('wvSectionPanel');
  if (!panel) return;

  if (wvGroupBy !== 'combined') {
    panel.style.display = 'none';
    return;
  }

  // Collect sections in order from first day column
  var sectionItems = [];
  var seen = {};
  grid.querySelectorAll('.wv-acc-hdr[data-section]').forEach(function(hdr) {
    var sec = hdr.dataset.section;
    if (seen[sec]) return;
    seen[sec] = true;
    var titleEl = hdr.querySelector('.wv-acc-title');
    var badgeEl = hdr.querySelector('.wv-acc-badge');
    sectionItems.push({
      sec: sec,
      title: titleEl ? titleEl.textContent : sec,
      badge: badgeEl ? badgeEl.textContent : ''
    });
  });

  var spChevUp   = '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><polyline points="18 15 12 9 6 15"/></svg>';
  var spChevDown = '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"/></svg>';

  // Measure col-hdr height so we can add a matching spacer at the top of the panel
  var firstColHdr = grid.querySelector('.wv-col-hdr');
  var colHdrH = firstColHdr ? Math.round(firstColHdr.getBoundingClientRect().height) : 0;

  panel.innerHTML = '<div class="wvsp-hdr-spacer" style="height:' + colHdrH + 'px;flex-shrink:0;border-bottom:2px solid #006461"></div>'
    + sectionItems.map(function(item) {
    var collapsed = !!wvCollapsed[item.sec];
    return '<div class="wvsp-item' + (collapsed ? '' : ' wvsp-open') + '" onclick="wvToggleSection(\'' + item.sec + '\')" data-section="' + item.sec + '">'
      + '<span class="wvsp-chev">' + (collapsed ? spChevDown : spChevUp) + '</span>'
      + '<span class="wvsp-title">' + item.title + '</span>'
      + (item.badge ? '<span class="wv-acc-badge">' + item.badge + '</span>' : '')
      + '</div>';
  }).join('');

  panel.style.display = '';

  // Sync close-out button with weekly checkbox state
  _syncCloseOutBtn();

  // Sync panel item heights to actual rendered section heights (not body.scrollHeight
  // which can diverge from layout height once flex/grid stretching is applied)
  sectionItems.forEach(function(item) {
    var panelItem = panel.querySelector('[data-section="' + item.sec + '"]');
    if (!panelItem) return;
    var collapsed = !!wvCollapsed[item.sec];
    if (collapsed) {
      panelItem.style.height = '';
      panelItem.style.minHeight = '';
      return;
    }
    // Use the actual rendered section height from the grid (tallest column wins)
    var maxSectH = 0;
    grid.querySelectorAll('.wv-acc-hdr[data-section="' + item.sec + '"]').forEach(function(hdr) {
      var sect = hdr.closest('.wv-acc-sect');
      if (sect) {
        var h = Math.round(sect.getBoundingClientRect().height);
        if (h > maxSectH) maxSectH = h;
      }
    });
    if (maxSectH > 0) {
      panelItem.style.height = maxSectH + 'px';
      panelItem.style.minHeight = maxSectH + 'px';
    }
  });
}

// ── Section panel toggle ──────────────────────────────────────────────────
window.wvToggleSection = function(section) {
  if (!wvCollapsed.hasOwnProperty(section)) wvCollapsed[section] = true;
  wvCollapsed[section] = !wvCollapsed[section];
  buildWeekGrid(wvMonth, wvWeekStart, wvWeekStart);
};

// ── Open All / Close All (works for both combined and room & board views) ─
const COMBINED_SECTIONS = ['daily','detailed','meals','avail','availAlloc','toRates','promos'];

// ── Room Type & Board default state: RT open, BT and sub-sections closed ────
(function initRtDefaults() {
  [0,1,2,3,4,5].forEach(function(ri) {
    if (!Object.prototype.hasOwnProperty.call(wvCollapsed, 'rt_' + ri)) wvCollapsed['rt_' + ri] = false;
    [0,1,2,3].forEach(function(bi) {
      ['bt_','btdet_','btavail_','btto_'].forEach(function(pfx) {
        var k = pfx + ri + '_' + bi;
        if (!Object.prototype.hasOwnProperty.call(wvCollapsed, k)) wvCollapsed[k] = true;
      });
    });
  });
}());
function setAllAccordions(collapse) {
  if (wvGroupBy === 'dailyH') {
    dhSetAll(collapse);
    return;
  }
  if (wvGroupBy === 'dailyB') {
    wbSetAll(collapse);
    return;
  }
  if (wvGroupBy === 'combined') {
    COMBINED_SECTIONS.forEach(function(k) { wvCollapsed[k] = collapse; });
    // Also set any other keys in wvCollapsed so inner accordions collapse too
    for (var ck in wvCollapsed) {
      if (wvCollapsed.hasOwnProperty(ck)) wvCollapsed[ck] = collapse;
    }
  } else {
    // RT level
    [0,1,2,3,4,5].forEach(function(ri) {
      wvCollapsed['rt_' + ri] = collapse;
      // BT level
      [0,1,2,3].forEach(function(bi) {
        wvCollapsed['bt_' + ri + '_' + bi] = collapse;
        // Sub-section level
        wvCollapsed['btdet_'   + ri + '_' + bi] = collapse;
        wvCollapsed['btavail_' + ri + '_' + bi] = collapse;
        wvCollapsed['btto_'    + ri + '_' + bi] = collapse;
      });
    });
  }
  buildWeekGrid(wvMonth, wvWeekStart, wvWeekStart);
}
document.getElementById('wvRtCloseAll')?.addEventListener('click', function() { setAllAccordions(true); });
document.getElementById('wvRtOpenAll')?.addEventListener('click',  function() { setAllAccordions(false); });

function _updateAccBtnState() {
  var disabled = (wvGroupBy === 'roomType' || wvGroupBy === 'coReport');
  // Open/Close All
  ['wvRtCloseAll','wvRtOpenAll'].forEach(function(id) {
    var btn = document.getElementById(id);
    if (!btn) return;
    btn.disabled = disabled;
    btn.style.opacity = disabled ? '0.35' : '';
    btn.style.cursor  = disabled ? 'not-allowed' : '';
  });
  // Table Settings & Filters buttons
  ['wvTableSettingsBtn','wvFiltersBtn'].forEach(function(id) {
    var btn = document.getElementById(id);
    if (!btn) return;
    btn.disabled = disabled;
    btn.style.opacity      = disabled ? '0.35' : '';
    btn.style.cursor       = disabled ? 'not-allowed' : '';
    btn.style.pointerEvents = disabled ? 'none' : '';
  });
  // Compare pills
  var pillsWrap = document.getElementById('wvCmpPills');
  if (pillsWrap) {
    pillsWrap.querySelectorAll('.wv-cmp-pill').forEach(function(p) {
      p.disabled = disabled;
      p.style.opacity       = disabled ? '0.35' : '';
      p.style.cursor        = disabled ? 'not-allowed' : '';
      p.style.pointerEvents = disabled ? 'none' : '';
    });
  }
}

// ── Reorder Modal (shared across Daily, Daily H, Daily R) ─────────────────
var _tsDragEl = null;

// Mapping: sect/group IDs → wvMetricState keys
var _tsMetricMap = {
  // Group-level (top)
  g_closeouts: ['dm_closeouts'],
  g_daily:   ['capacity','onlineOffline','adr','revenue'],
  g_more:    ['dm_rnSold','dm_trevpar','dm_pickup','dm_avgAdults','dm_avgChildren',
              'dm_totalAdults','dm_totalChildren','dm_totalGuests','dm_avgLos','dm_avgLeadTime',
              'dm_availRooms','dm_availGuar'],
  g_meals:   ['mealsSummary'],
  g_biz:     ['bizMix'],
  g_avail:   ['avail','availAlloc'],
  g_torates: ['toRates'],
  // Close-outs sub-rows
  co_rooms: ['dm_co_rooms'], co_boards: ['dm_co_boards'], co_tos: ['dm_co_tos'],
  // Sect-level (child)
  occ: ['capacity'], onoff: ['onlineOffline'], adr: ['adr'], rev: ['revenue'],
  rn: ['dm_rnSold'], revpar_s: ['dm_trevpar'], pickup_s: ['dm_pickup'],
  avga_s: ['dm_avgAdults'], avgc_s: ['dm_avgChildren'],
  tota_s: ['dm_totalAdults'], totc_s: ['dm_totalChildren'],
  totg_s: ['dm_totalGuests'], los_s: ['dm_avgLos'], lead_s: ['dm_avgLeadTime'],
  avail_s: ['dm_availRooms'], availg_s: ['dm_availGuar'],
  biz: ['bizMix'],
  mp_ai: ['mealsSummary'], mp_bb: ['mealsSummary'], mp_hb: ['mealsSummary'],
  mp_ro: ['mealsSummary'], mp_sum: ['mealsSummary']
};
var _tsCheckSvg = '<svg viewBox="0 0 18 18" width="14" height="14" fill="none"><path d="M3.5 9l3.5 3.5 7-7" stroke="#fff" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"/></svg>';
var _tsDragSvg = '<svg width="24" height="24" viewBox="0 0 24 24" fill="currentColor">'
  + '<circle cx="9" cy="5" r="1.5"/><circle cx="15" cy="5" r="1.5"/>'
  + '<circle cx="9" cy="10" r="1.5"/><circle cx="15" cy="10" r="1.5"/>'
  + '<circle cx="9" cy="15" r="1.5"/><circle cx="15" cy="15" r="1.5"/>'
  + '<circle cx="9" cy="20" r="1.5"/><circle cx="15" cy="20" r="1.5"/></svg>';

function _tsToggleCb(cb) {
  cb.classList.toggle('unchecked');
  cb.innerHTML = cb.classList.contains('unchecked') ? '' : _tsCheckSvg;
}

function _tsAddRow(list, key, label, depth, draggable, checked) {
  if (checked === undefined) checked = true;
  var row = document.createElement('div');
  row.className = 'ts-tree-row';
  row.style.paddingLeft = (8 + depth * 24) + 'px';
  if (draggable) row.draggable = true;
  row.dataset.parKey = key;
  row.dataset.depth = depth;

  // Checkbox
  var cb = document.createElement('span');
  cb.className = 'ts-checkbox' + (checked ? '' : ' unchecked');
  cb.innerHTML = checked ? _tsCheckSvg : '';
  cb.addEventListener('click', function(e) {
    e.stopPropagation();
    _tsToggleCb(cb);
    // If parent, toggle all children too
    if (depth === 0) {
      var isChecked = !cb.classList.contains('unchecked');
      var sib = row.nextElementSibling;
      while (sib && parseInt(sib.dataset.depth || 0) > 0) {
        var childCb = sib.querySelector('.ts-checkbox');
        if (childCb) {
          if (isChecked) { childCb.classList.remove('unchecked'); childCb.innerHTML = _tsCheckSvg; }
          else { childCb.classList.add('unchecked'); childCb.innerHTML = ''; }
        }
        sib = sib.nextElementSibling;
      }
    }
  });

  // Label
  var lbl = document.createElement('span');
  lbl.className = 'ts-tree-lbl';
  lbl.textContent = label;
  if (depth === 0) lbl.style.fontWeight = '600';

  // Drag handle (right side)
  var handle = document.createElement('span');
  handle.className = 'ts-drag-handle';
  handle.innerHTML = _tsDragSvg;

  row.appendChild(cb);
  row.appendChild(lbl);
  row.appendChild(handle);

  // Drag events (only for draggable rows)
  if (draggable) {
    row.addEventListener('dragstart', function(e) {
      _tsDragEl = row; row.classList.add('dragging');
      e.dataTransfer.effectAllowed = 'move';
    });
    row.addEventListener('dragend', function() {
      row.classList.remove('dragging'); _tsDragEl = null;
      list.querySelectorAll('.ts-tree-row').forEach(function(r) { r.classList.remove('drop-above','drop-below'); });
    });
  }
  row.addEventListener('dragover', function(e) {
    e.preventDefault();
    if (!_tsDragEl || _tsDragEl === row) return;
    var mid = row.getBoundingClientRect().top + row.offsetHeight / 2;
    row.classList.remove('drop-above','drop-below');
    row.classList.add(e.clientY < mid ? 'drop-above' : 'drop-below');
  });
  row.addEventListener('dragleave', function() { row.classList.remove('drop-above','drop-below'); });
  row.addEventListener('drop', function(e) {
    e.preventDefault();
    row.classList.remove('drop-above','drop-below');
    if (!_tsDragEl || _tsDragEl === row) return;
    var mid = row.getBoundingClientRect().top + row.offsetHeight / 2;
    list.insertBefore(_tsDragEl, e.clientY < mid ? row : row.nextSibling);
  });

  list.appendChild(row);
  return row;
}

function _tsIsChecked(key) {
  var mks = _tsMetricMap[key];
  if (!mks || mks.length === 0) return true;
  for (var i = 0; i < mks.length; i++) {
    if (wvMetricState[mks[i]]) return true;
  }
  return false;
}

function _buildReorderList(list, items) {
  // items: [{ key, lbl, children: [{ key, lbl, children: [...] }] }]
  list.innerHTML = '';
  items.forEach(function(p) {
    _tsAddRow(list, p.key, p.lbl, 0, true, _tsIsChecked(p.key));
    if (p.children) {
      p.children.forEach(function(c) {
        _tsAddRow(list, c.key, c.lbl, 1, false, _tsIsChecked(c.key));
        if (c.children) {
          c.children.forEach(function(sc) {
            _tsAddRow(list, sc.key, sc.lbl, 2, false, _tsIsChecked(sc.key));
          });
        }
      });
    }
  });
}

// Select All / Deselect All for Table Settings
window.tsCheckAll = function(checked) {
  var list = document.getElementById('dhReorderList');
  if (!list) return;
  list.querySelectorAll('.ts-checkbox').forEach(function(cb) {
    if (checked) { cb.classList.remove('unchecked'); cb.innerHTML = _tsCheckSvg; }
    else { cb.classList.add('unchecked'); cb.innerHTML = ''; }
  });
};

window.dhOpenReorder = function() {
  var modal = document.getElementById('dhReorderModal');
  var list  = document.getElementById('dhReorderList');
  if (!modal || !list) return;

  // Title is now fixed in HTML as "Table Settings"

  if (wvGroupBy === 'dailyH') {
    // Build hierarchical items: sec → par as children
    var items = [], curSec = null;
    _dhAllRows.forEach(function(r) {
      if (r.type === 'sec') {
        curSec = { key: r.secKey, lbl: r.lbl, children: [] };
        items.push(curSec);
      }
      if (r.type === 'par' && curSec) {
        curSec.children.push({ key: r.parKey, lbl: r.lbl });
      }
    });
    _buildReorderList(list, items);

  } else if (wvGroupBy === 'combined') {
    var renderedSecs = {};
    var g = document.getElementById('weekGrid');
    if (g) g.querySelectorAll('.wv-acc-hdr[data-section]').forEach(function(el) { renderedSecs[el.dataset.section] = true; });
    var curOrder = (_wvSectionOrder && _wvSectionOrder.length) ? _wvSectionOrder : WV_SECTIONS_DEF.map(function(s){return s.key;});
    var items = [];
    curOrder.forEach(function(k) {
      if (!renderedSecs[k]) return;
      var def = WV_SECTIONS_DEF.filter(function(s){return s.key===k;})[0];
      if (def) items.push({ key: k, lbl: def.lbl });
    });
    WV_SECTIONS_DEF.forEach(function(s) {
      if (renderedSecs[s.key] && curOrder.indexOf(s.key) === -1)
        items.push({ key: s.key, lbl: s.lbl });
    });
    _buildReorderList(list, items);

  } else if (wvGroupBy === 'report') {
    var curOrder = (_drColOrder && _drColOrder.length) ? _drColOrder : DR_GROUPS_DEF.map(function(g){return g.key;});
    var items = curOrder.map(function(k) {
      var def = DR_GROUPS_DEF.filter(function(g){return g.key===k;})[0] || { clr:'#374151' };
      return { key: k, lbl: k };
    });
    _buildReorderList(list, items);
  } else if (wvGroupBy === 'dailyB') {
    var curOrder = (_wbGroupOrder && _wbGroupOrder.length) ? _wbGroupOrder : WB_GROUPS_DEF.map(function(g){return g.key;});
    var gd = window._wbGrpData || {};
    var items = curOrder.map(function(k) {
      var def = WB_GROUPS_DEF.filter(function(g){return g.key===k;})[0] || { lbl: k, clr: '#374151' };
      var children = [];
      var rows = gd[k] || [];
      // Build sect → sub hierarchy
      var curSect = null;
      rows.forEach(function(r) {
        if (r.type === 'top') return;
        if (r.type === 'sect') {
          curSect = { key: r.id, lbl: r.label, children: [] };
          children.push(curSect);
        } else if (r.type === 'sub' && curSect) {
          curSect.children.push({ key: r.id, lbl: r.label });
        } else if (r.type === 'sub') {
          children.push({ key: r.id, lbl: r.label });
        }
      });
      return { key: k, lbl: def.lbl, children: children };
    });
    _buildReorderList(list, items);
  }

  modal.style.display = 'flex';
};

window.dhReorderModalBg = function(e) {
  if (e.target.id === 'dhReorderModal') e.target.style.display = 'none';
};

window.dhApplyReorder = function() {
  var list = document.getElementById('dhReorderList');
  var order = [];
  list.querySelectorAll('.ts-tree-row[data-depth="0"]').forEach(function(li) { order.push(li.dataset.parKey); });

  // Update wvMetricState from checkbox states
  if (wvGroupBy === 'dailyB' || wvGroupBy === 'dailyH') {
    list.querySelectorAll('.ts-tree-row').forEach(function(row) {
      var key = row.dataset.parKey;
      var isChecked = !row.querySelector('.ts-checkbox').classList.contains('unchecked');
      var metricKeys = _tsMetricMap[key];
      if (metricKeys) {
        metricKeys.forEach(function(mk) { wvMetricState[mk] = isChecked; });
      }
    });
  }

  document.getElementById('dhReorderModal').style.display = 'none';
  if (wvGroupBy === 'dailyH') {
    _dhMetricOrder = order;
    var a = _dhLastInitArgs;
    if (a) initDailyHGrid(a.days, a.month, a.day, a.container);
  } else if (wvGroupBy === 'combined') {
    _wvSectionOrder = order;
    buildWeekGrid(wvMonth, wvWeekStart, wvWeekStart);
  } else if (wvGroupBy === 'report') {
    _drColOrder = order;
    var a = _drLastInitArgs;
    if (a) initDailyRevGrid(a.days, a.container);
  } else if (wvGroupBy === 'dailyB') {
    _wbGroupOrder = order;
    buildWeekGrid(wvMonth, wvWeekStart, wvWeekStart);
  }
};

window.dhResetReorder = function() {
  document.getElementById('dhReorderModal').style.display = 'none';
  if (wvGroupBy === 'dailyH') {
    _dhMetricOrder = null;
    var a = _dhLastInitArgs;
    if (a) initDailyHGrid(a.days, a.month, a.day, a.container);
  } else if (wvGroupBy === 'combined') {
    _wvSectionOrder = null;
    buildWeekGrid(wvMonth, wvWeekStart, wvWeekStart);
  } else if (wvGroupBy === 'report') {
    _drColOrder = null;
    var a = _drLastInitArgs;
    if (a) initDailyRevGrid(a.days, a.container);
  } else if (wvGroupBy === 'dailyB') {
    _wbGroupOrder = null;
    buildWeekGrid(wvMonth, wvWeekStart, wvWeekStart);
  }
};

// ── Group-by toggle ───────────────────────────────────────────────────────
document.querySelectorAll('#weekView .wv-groupby-btn').forEach(function(btn) {
  btn.addEventListener('click', function() {
    // "Monthly" tab in weekly bar → switch back to month view
    if (this.dataset.groupby === 'monthly') {
      goToMonthView();
      return;
    }
    wvGroupBy = this.dataset.groupby;
    document.querySelectorAll('#weekView .wv-groupby-btn').forEach(function(b) { b.classList.remove('active'); });
    this.classList.add('active');
    _updateAccBtnState();
    // Table Settings inline button visibility handled by topbar
    buildWeekGrid(wvMonth, wvWeekStart, wvWeekStart);
  });
});

// Return to monthly calendar view
window.goToMonthView = function() {
  // Sync weekly filters → calendar so selections persist across views
  syncFiltersWvToCal();
  applyFilterUI('calFiltersDropdown');
  _syncPickupBtnUI('cal');

  document.getElementById('demand-calendar').style.display = '';
  document.getElementById('weekView').classList.remove('visible');
  var backArrow = document.getElementById('wvBack');
  if (backArrow) backArrow.style.display = 'none';
  var hdrCtr = document.getElementById('wvHeaderCenter');
  if (hdrCtr) hdrCtr.style.display = 'none';
  // Show monthly tab bar and reset "Monthly" as active tab
  var moBar = document.getElementById('moGroupbyBar');
  if (moBar) moBar.style.display = (calDisplayView <= 3) ? '' : 'none';
  document.querySelectorAll('.mo-grp-btn').forEach(function(b) {
    b.classList.toggle('active', b.dataset.mogroupby === 'monthly');
  });
  renderCalendar();
};

// Week nav + back (legacy arrow in cal-header also calls goToMonthView)
document.getElementById('wvBack')?.addEventListener('click', goToMonthView);
document.getElementById('wvPrev')?.addEventListener('click', () => {
  const dim = [0,31,28,31,30,31,30,31,31,30,31,30,31];
  wvWeekStart -= 1;
  if (wvWeekStart < 1) { wvMonth--; if (wvMonth < 1) wvMonth = 12; wvWeekStart = dim[wvMonth]; }
  buildWeekGrid(wvMonth, wvWeekStart, wvWeekStart);
});
document.getElementById('wvNext')?.addEventListener('click', () => {
  const dim = [0,31,28,31,30,31,30,31,31,30,31,30,31];
  wvWeekStart += 1;
  if (wvWeekStart > dim[wvMonth]) { wvMonth++; if (wvMonth > 12) wvMonth = 1; wvWeekStart = 1; }
  buildWeekGrid(wvMonth, wvWeekStart, wvWeekStart);
});

// ── Week date picker ──────────────────────────────────────────────
// ── Week date-picker popup ────────────────────────────────────────
var wvwpViewMonth = 3, wvwpViewYear = 2026;

function wvWeekPickToggle() {
  var panel = document.getElementById('wvWeekPickPanel');
  if (!panel) return;
  if (panel.style.display !== 'none') { panel.style.display = 'none'; return; }
  wvwpViewMonth = wvMonth; wvwpViewYear = wvYear;
  wvwpRender();
  var btn = document.getElementById('wvWeekPickBtn');
  var rect = btn.getBoundingClientRect();
  panel.style.left = rect.left + 'px';
  panel.style.top  = (rect.bottom + 4) + 'px';
  panel.style.display = 'block';
}
function wvwpNav(dir) {
  wvwpViewMonth += dir;
  if (wvwpViewMonth < 1)  { wvwpViewMonth = 12; wvwpViewYear--; }
  if (wvwpViewMonth > 12) { wvwpViewMonth = 1;  wvwpViewYear++; }
  wvwpRender();
}
function wvwpDayIdx(m, d) {
  var dim = [0,31,28,31,30,31,30,31,31,30,31,30,31];
  var idx = d; for (var i = 1; i < m; i++) idx += dim[i]; return idx;
}
function wvwpRender() {
  var MNAMES = ['','January','February','March','April','May','June','July','August','September','October','November','December'];
  var DNAMES = ['Mo','Tu','We','Th','Fr','Sa','Su'];
  var dim    = [0,31,28,31,30,31,30,31,31,30,31,30,31];
  document.getElementById('wvwpTitle').textContent = MNAMES[wvwpViewMonth] + ' ' + wvwpViewYear;
  var startIdx = wvwpDayIdx(wvMonth, wvWeekStart), endIdx = startIdx + 6;
  var firstDow  = new Date(wvwpViewYear, wvwpViewMonth - 1, 1).getDay(); // 0=Sun
  var startOff  = (firstDow + 6) % 7; // Mon-based offset
  var html = '';
  DNAMES.forEach(function(n){ html += '<div class="wvwp-day-hdr">'+n+'</div>'; });
  for (var i = 0; i < startOff; i++) html += '<div class="wvwp-day wvwp-empty"><div class="wvwp-day-bg"></div><div class="wvwp-day-lbl"></div></div>';
  for (var d = 1; d <= dim[wvwpViewMonth]; d++) {
    var idx = wvwpDayIdx(wvwpViewMonth, d);
    var cls = 'wvwp-day';
    if (idx === startIdx) cls += ' wvwp-week-start';
    if (idx === endIdx)   cls += ' wvwp-week-end';
    if (idx > startIdx && idx < endIdx) cls += ' wvwp-in-week';
    html += '<div class="'+cls+'" onclick="wvPickWeekDay('+wvwpViewMonth+','+d+')">'
          + '<div class="wvwp-day-bg"></div><div class="wvwp-day-lbl">'+d+'</div></div>';
  }
  document.getElementById('wvwpGrid').innerHTML = html;
}
function wvPickWeekDay(m, d) {
  wvMonth = m; wvWeekStart = d;
  buildWeekGrid(wvMonth, wvWeekStart, wvWeekStart);
  document.getElementById('wvWeekPickPanel').style.display = 'none';
}
// Close picker on outside click
document.addEventListener('click', function(e) {
  var wrap = document.getElementById('wvWeekPickWrap');
  if (wrap && !wrap.contains(e.target)) {
    var p = document.getElementById('wvWeekPickPanel');
    if (p) p.style.display = 'none';
  }
});

// ── Partial closure padlock toggle ───────────────────────────────
document.getElementById('weekGrid')?.addEventListener('click', function(e) {
  const btn = e.target.closest('.wv-partial-lock-btn');
  if (!btn) return;
  e.stopPropagation();
  const panel = document.getElementById(btn.dataset.restrictId);
  if (!panel) return;
  const open = panel.classList.toggle('wv-restrict-open');
  btn.classList.toggle('wv-partial-lock-active', open);
});

// ── Orange padlock TO detail toggle ──────────────────────────────────────
document.getElementById('weekGrid')?.addEventListener('click', function(e) {
  const btn = e.target.closest('.wv-tos-btn');
  if (!btn) return;
  e.stopPropagation();
  const key = btn.dataset.toskey;
  wvTosOpen[key] = !wvTosOpen[key];
  buildWeekGrid(wvMonth, wvWeekStart, wvWeekStart);
});

// ── Weekly Range Selection ────────────────────────────────────────
(function() {
  const MNAMES = ['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const svgCalSm = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" width="16" height="16"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="3" y1="10" x2="21" y2="10"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="16" y1="2" x2="16" y2="6"/></svg>`;

  // wvSelBtn removed — range selection now triggered via wvCloseSelectRange()

  document.getElementById('weekGrid')?.addEventListener('click', function(e) {
    if (!wvSelPicking && !wvSelStart) return;
    const col = e.target.closest('.wv-col[data-dm]');
    if (!col) return;
    const dm = +col.dataset.dm, dd = +col.dataset.dd;

    if (!wvSelStart) {
      wvSelStart = { month: dm, day: dd };
      wvSelPicking = true;
      const btn = document.getElementById('wvSelBtn');
      if (btn) btn.innerHTML = svgCalSm + ' Pick end…';
    } else {
      wvSelEnd = { month: dm, day: dd };
      wvSelPicking = false;
      // Open Close Out modal pre-populated with range
      (function() {
        var s = wvSelStart, en = wvSelEnd;
        var startV = s.month * 100 + s.day, endV = en.month * 100 + en.day;
        var lo = startV <= endV ? s : en, hi = startV <= endV ? en : s;
        var pad = function(n){ return String(n).padStart(2,'0'); };
        var fromStr = '2026-' + pad(lo.month) + '-' + pad(lo.day);
        var toStr   = '2026-' + pad(hi.month) + '-' + pad(hi.day);
        if (typeof window._coOpenModal === 'function') window._coOpenModal(fromStr, toStr, 'wv');
      })();
    }
    buildWeekGrid(wvMonth, wvWeekStart, wvWeekStart);
  });

  document.addEventListener('keydown', function(e) {
    if (e.key === 'Escape' && wvSelStart) clearWvSelection();
  });
})();

/* ─── WEEK VIEW CLOSE DROPDOWN ─── */
(function() {
  window.wvCloseDropdownToggle = function(e) {
    e.stopPropagation();
    var dd = document.getElementById('wvCloseDropdown');
    if (!dd) return;
    dd.style.display = dd.style.display !== 'none' ? 'none' : 'block';
  };

  window.wvCloseSelectRange = function() {
    document.getElementById('wvCloseDropdown').style.display = 'none';
    if (wvSelStart) { clearWvSelection(); return; }
    wvSelPicking = true;
    const right = document.querySelector('.wv-topbar-right');
    if (right) right.classList.add('range-mode');
    // Highlight the Close button to show picking state
    const closeBtn = document.getElementById('wvCloseOutBtn');
    if (closeBtn) { closeBtn.style.background = '#006461'; closeBtn.style.color = '#fff'; }
  };

  window.wvCloseCustom = function() {
    var dd = document.getElementById('wvCloseDropdown');
    if (dd) dd.style.display = 'none';
    if (typeof window._coOpenModal === 'function') window._coOpenModal('', '', 'wv');
  };

  // Close dropdown on outside click
  document.addEventListener('click', function(e) {
    var wrap = document.getElementById('wvCloseWrap');
    var dd   = document.getElementById('wvCloseDropdown');
    if (dd && wrap && !wrap.contains(e.target)) dd.style.display = 'none';
  });
})();

// ── Weekly section collapse (roomType mode — headers still in grid) ───────
document.getElementById('weekGrid')?.addEventListener('click', function(e) {
  const el = e.target.closest('.wv-acc-hdr, .wv-collapse-btn');
  if (!el) return;
  if (wvGroupBy === 'combined') return; // handled by section panel
  e.stopPropagation();
  const section = el.dataset.section;
  if (section) {
    if (!wvCollapsed.hasOwnProperty(section)) wvCollapsed[section] = true;
    wvCollapsed[section] = !wvCollapsed[section];
  }
  buildWeekGrid(wvMonth, wvWeekStart, wvWeekStart);
});

// ── Metrics selector ─────────────────────────────────────────────
function updateMetricCheckboxes() {
  document.querySelectorAll('.wv-ms-cb[data-key]').forEach(function(cb) {
    cb.classList.toggle('checked', !!wvMetricState[cb.dataset.key]);
  });
}

document.addEventListener('click', function(e) {
  const btn = e.target.closest('#wvMetricsBtn');
  if (btn) {
    const dd = document.getElementById('wvMetricsDropdown');
    const isOpen = dd.style.display !== 'none';
    dd.style.display = isOpen ? 'none' : 'block';
    if (!isOpen) { renderPickupMetricItems(); updateMetricCheckboxes(); }
    e.stopPropagation(); return;
  }
  const cb = e.target.closest('.wv-ms-cb[data-key]');
  if (cb) {
    const key = cb.dataset.key;
    wvMetricState[key] = !wvMetricState[key];
    // Keep dm_pickup master in sync with individual window toggles
    if (key.startsWith('dm_pickup_')) {
      wvMetricState.dm_pickup = [0,1,2].some(function(i){ return !!wvMetricState['dm_pickup_'+i]; });
    }
    updateMetricCheckboxes();
    buildWeekGrid(wvMonth, wvWeekStart, wvWeekStart);
    return;
  }
  if (e.target.id === 'wvMsClearCard') {
    ['capacity','adr','revenue','onlineOffline','roomTypes','avail','availAlloc','toRates'].forEach(function(k){ wvMetricState[k] = false; });
    updateMetricCheckboxes();
    buildWeekGrid(wvMonth, wvWeekStart, wvWeekStart); return;
  }
  if (e.target.id === 'wvMsClearDetail') {
    ['dm_rnSold','dm_pickup','dm_pickup_0','dm_pickup_1','dm_pickup_2','dm_avgAdults','dm_avgChildren','dm_totalAdults','dm_totalChildren','dm_trevpar','dm_availRooms','dm_availGuar'].forEach(function(k){ wvMetricState[k] = false; });
    updateMetricCheckboxes();
    buildWeekGrid(wvMonth, wvWeekStart, wvWeekStart); return;
  }
  if (!e.target.closest('#wvMetricsWrap')) {
    const dd = document.getElementById('wvMetricsDropdown');
    if (dd) dd.style.display = 'none';
  }
});

// "View Week" button handled inside popup click listener (see popup IIFE above)

/* ─── RE-OPEN SALES (week view per-day + monthly bulk) ─── */
document.addEventListener('click', e => {
  const btn = e.target.closest('.wv-reopen-btn');
  if (btn) {
    const m = +btn.dataset.month, d = +btn.dataset.day;
    const key = `${m}-${d}`;
    if (confirm(`Re-open sales for ${['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][m]} ${d}?`)) {
      LOCKED_DAYS.delete(key);
      buildWeekGrid(wvMonth, wvWeekStart, wvWeekStart);
      renderCalendar();
    }
  }
});

// Monthly "Re-open Visible" button
document.querySelector('.btn-reopen')?.addEventListener('click', () => {
  if (LOCKED_DAYS.size) {
    if (confirm(`Re-open all ${LOCKED_DAYS.size} locked date(s)?`)) {
      LOCKED_DAYS.clear();
      renderCalendar();
    }
  } else {
    alert('No locked dates currently.');
  }
});

/* ─── DATE RANGE PICKER ─── */
(function() {
  const MONTHS = ['January','February','March','April','May','June',
                  'July','August','September','October','November','December'];
  const DAYS_IN = (y, m) => new Date(y, m, 0).getDate();
  const FIRST_DOW = (y, m) => new Date(y, m - 1, 1).getDay(); // 0=Sun

  const trigger   = document.getElementById('dpTrigger');
  const dropdown  = document.getElementById('dpDropdown');
  const monthsCnt = document.getElementById('dpMonths');
  const rangeLabel= document.getElementById('dpRangeLabel');
  const trigText  = document.getElementById('dpTriggerText');
  const cancelBtn = document.getElementById('dpCancel');
  const applyBtn  = document.getElementById('dpApply');
  if (!trigger) return;

  // Current displayed months (left = viewYear/viewMonth, right = next)
  let viewYear  = 2025;
  let viewMonth = 7; // July

  // Selection state
  let startDate = { y:2025, m:7, d:17 };
  let endDate   = { y:2025, m:7, d:25 };
  let picking   = null; // null | 'start' | 'end'
  let hoverDate = null;

  function fmt(d) {
    return `${String(d.m).padStart(2,'0')}/${String(d.d).padStart(2,'0')}/${d.y}`;
  }
  function dateVal(y, m, d) { return y * 10000 + m * 100 + d; }
  function startVal() { return startDate ? dateVal(startDate.y, startDate.m, startDate.d) : null; }
  function endVal()   { return endDate   ? dateVal(endDate.y,   endDate.m,   endDate.d)   : null; }
  function hoverVal() { return hoverDate ? dateVal(hoverDate.y, hoverDate.m, hoverDate.d) : null; }

  function renderDropdown() {
    monthsCnt.innerHTML = '';
    // Two months: left + right
    const months = [
      { y: viewYear, m: viewMonth },
      viewMonth === 12 ? { y: viewYear + 1, m: 1 } : { y: viewYear, m: viewMonth + 1 }
    ];
    months.forEach((mo, idx) => {
      const col = document.createElement('div');
      col.className = 'dp-month';

      // Navigation row
      const navRow = document.createElement('div');
      navRow.className = 'dp-month-nav';

      // Left month: show << < ; Right month: show > >>
      if (idx === 0) {
        navRow.innerHTML = `
          <button class="dp-month-nav-btn" data-act="first">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="11 17 6 12 11 7"/><line x1="18" y1="12" x2="6" y2="12"/></svg>
          </button>
          <button class="dp-month-nav-btn" data-act="prev">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="15 18 9 12 15 6"/></svg>
          </button>
          <span class="dp-month-title">${MONTHS[mo.m - 1]} ${mo.y}</span>
          <button class="dp-month-nav-btn hidden"></button>
          <button class="dp-month-nav-btn hidden"></button>`;
      } else {
        navRow.innerHTML = `
          <button class="dp-month-nav-btn hidden"></button>
          <button class="dp-month-nav-btn hidden"></button>
          <span class="dp-month-title">${MONTHS[mo.m - 1]} ${mo.y}</span>
          <button class="dp-month-nav-btn" data-act="next">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="9 18 15 12 9 6"/></svg>
          </button>
          <button class="dp-month-nav-btn" data-act="last">
            <span class="material-icons" style="font-size:20px">arrow_forward</span>
          </button>`;
      }
      navRow.querySelectorAll('[data-act]').forEach(btn => {
        btn.addEventListener('click', e => {
          const act = e.currentTarget.dataset.act;
          if (act === 'prev' || act === 'first') {
            if (viewMonth === 1) { viewMonth = 12; viewYear--; } else { viewMonth--; }
          } else {
            if (viewMonth === 12) { viewMonth = 1; viewYear++; } else { viewMonth++; }
          }
          renderDropdown();
        });
      });
      col.appendChild(navRow);

      // Calendar
      const cal = document.createElement('div');
      cal.className = 'dp-cal';

      // Weekday header
      const legend = document.createElement('div');
      legend.className = 'dp-week-legend';
      ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'].forEach(d => {
        const s = document.createElement('span'); s.textContent = d;
        legend.appendChild(s);
      });
      cal.appendChild(legend);

      const totalDays = DAYS_IN(mo.y, mo.m);
      const first = FIRST_DOW(mo.y, mo.m);
      let week = document.createElement('div');
      week.className = 'dp-week';

      // Empty cells at start
      for (let i = 0; i < first; i++) {
        const cell = document.createElement('div');
        cell.className = 'dp-day empty';
        week.appendChild(cell);
      }

      for (let d = 1; d <= totalDays; d++) {
        if ((first + d - 1) % 7 === 0 && d > 1) {
          cal.appendChild(week);
          week = document.createElement('div');
          week.className = 'dp-week';
        }
        const cell = document.createElement('div');
        cell.className = 'dp-day';
        const dv = dateVal(mo.y, mo.m, d);
        const sv = startVal();
        const ev = endVal();
        const hv = picking === 'end' ? hoverVal() : null;
        const effectiveEnd = (picking === 'end' && hv && sv) ? (hv > sv ? hv : sv) : ev;
        const effectiveStart = sv;

        const inner = document.createElement('span');
        inner.className = 'dp-inner';
        inner.textContent = d;
        cell.appendChild(inner);

        if (effectiveStart && dv === effectiveStart) {
          cell.classList.add('range-start');
          inner.textContent = d;
        } else if (effectiveEnd && effectiveStart && dv === effectiveEnd && dv !== effectiveStart) {
          cell.classList.add('range-end');
          inner.textContent = d;
        } else if (effectiveStart && effectiveEnd && dv > effectiveStart && dv < effectiveEnd) {
          cell.classList.add('range-mid');
          // row-first / row-last for pill shape
          const posInWeek = (first + d - 1) % 7;
          if (posInWeek === 0) cell.classList.add('row-first');
          if (posInWeek === 6) cell.classList.add('row-last');
        }

        cell.addEventListener('mouseenter', () => {
          if (picking === 'end') {
            if (!hoverDate || hoverDate.y !== mo.y || hoverDate.m !== mo.m || hoverDate.d !== d) {
              hoverDate = { y: mo.y, m: mo.m, d };
              renderDropdown();
            }
          }
        });
        cell.addEventListener('click', () => {
          if (!picking || picking === 'start') {
            startDate = { y: mo.y, m: mo.m, d };
            endDate = null;
            picking = 'end';
            hoverDate = null;
          } else {
            const clickedVal = dateVal(mo.y, mo.m, d);
            if (clickedVal >= startVal()) {
              endDate = { y: mo.y, m: mo.m, d };
            } else {
              endDate = startDate;
              startDate = { y: mo.y, m: mo.m, d };
            }
            picking = null;
            hoverDate = null;
          }
          updateRangeLabel();
          renderDropdown();
        });

        week.appendChild(cell);
      }
      // Fill remaining
      cal.appendChild(week);
      col.appendChild(cal);
      monthsCnt.appendChild(col);
    });
    /* Clear hover preview when mouse leaves the calendar grid */
    monthsCnt.addEventListener('mouseleave', () => {
      if (picking === 'end' && hoverDate) { hoverDate = null; renderDropdown(); }
    }, { once: true });
    updateRangeLabel();
  }

  function updateRangeLabel() {
    if (startDate && endDate) {
      const label = `${fmt(startDate)} – ${fmt(endDate)}`;
      rangeLabel.textContent = label;
    } else if (startDate) {
      rangeLabel.textContent = `${fmt(startDate)} – ...`;
    }
  }

  function openPicker() {
    picking = null; hoverDate = null;
    renderDropdown();
    dropdown.classList.add('open');
    trigger.classList.add('open');
  }
  function closePicker() {
    dropdown.classList.remove('open');
    trigger.classList.remove('open');
    picking = null;
  }

  trigger.addEventListener('click', e => {
    e.stopPropagation();
    dropdown.classList.contains('open') ? closePicker() : openPicker();
  });
  cancelBtn.addEventListener('click', closePicker);
  applyBtn.addEventListener('click', () => {
    if (startDate && endDate) {
      trigText.textContent = `${fmt(startDate)} – ${fmt(endDate)}`;
      closePicker();
      if (window.dpOnApply) window.dpOnApply(startDate, endDate);
    }
    // If only start selected, don't close — wait for end date
  });

  // Close on Escape
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape' && dropdown.classList.contains('open')) closePicker();
  });
  // Close on outside click
  document.addEventListener('click', e => {
    if (dropdown.classList.contains('open') && !dropdown.contains(e.target) && !trigger.contains(e.target)) {
      closePicker();
    }
  });

  // Init trigger display
  trigText.textContent = `${fmt(startDate)} – ${fmt(endDate)}`;
})();

/* ─── Contracts & Promotions date-driven summaries ─── */
function updateContractsStats(start, end) {
  var BASE_DAYS = 365; // annual baseline
  var d1   = new Date(start.y, start.m - 1, start.d);
  var d2   = new Date(end.y,   end.m   - 1, end.d);
  var days = Math.max(1, Math.round((d2 - d1) / 864e5) + 1);
  var f    = days / BASE_DAYS;
  function set(id, val) { var el = document.getElementById(id); if (el) el.textContent = val; }
  // Contracts inventory
  var alloc   = Math.max(1,  Math.round(920   * f));
  var sold    = Math.min(alloc, Math.max(0, Math.round(662 * f)));
  var remain  = alloc - sold;
  var soldPct = (sold / alloc * 100).toFixed(1);
  var rev     = Math.round(121700 * f);
  set('acStatAlloc',   alloc.toLocaleString());
  set('acStatSold',    sold.toLocaleString());
  set('acStatRemain',  remain.toLocaleString());
  set('acStatSoldPct', soldPct + '%');
  set('acStatRev',     fmtStatK(rev));
  // Key insights text
  var insEl = document.querySelector('.ac-insights-list');
  if (insEl) insEl.innerHTML =
    '<div class="ac-insight"><span class="ac-insight-dot"></span>Overall Utilization: ' + soldPct + '% of allocated rooms sold</div>'
    + '<div class="ac-insight"><span class="ac-insight-dot"></span>Avg ADR: $184 across all contracted rooms</div>'
    + '<div class="ac-insight"><span class="ac-insight-dot"></span>Available Inventory: ' + remain.toLocaleString() + ' rooms still available</div>';
  // Promotions stats — bookings/revenue scale; active/total stay anchored
  var pmBook   = Math.max(1, Math.round(324    * f));
  var pmRev    = Math.round(388800 * f);
  var pmActive = Math.min(18, Math.max(1, Math.round(17 * Math.pow(f, 0.6))));
  set('pmStatActive', pmActive);
  set('pmStatTotal',  18);
  set('pmStatBook',   pmBook.toLocaleString());
  set('pmStatRev',    fmtStatK(pmRev));
}
window.dpOnApply = updateContractsStats;

/* ─── INIT ─── */
initTargetsGrid();
updateChart();
buildRoomTypeTable();
buildCalendar();
updateRevStats();
updateContractsStats({ y:2025, m:7, d:17 }, { y:2025, m:7, d:25 });

/* ─── MOBILE SIDEBAR TOGGLE ─── */
(function () {
  const hamburger = document.getElementById('hamburger');
  const sidebar   = document.querySelector('.sidebar');
  const overlay   = document.getElementById('sidebarOverlay');
  if (!hamburger || !sidebar || !overlay) return;
  function openSidebar()  { sidebar.classList.add('open');  overlay.classList.add('visible'); }
  function closeSidebar() { sidebar.classList.remove('open'); overlay.classList.remove('visible'); }
  hamburger.addEventListener('click', () => sidebar.classList.contains('open') ? closeSidebar() : openSidebar());
  overlay.addEventListener('click', closeSidebar);
})();

/* ─── SEARCH CLEAR BUTTON ─── */
(function () {
  const input = document.getElementById('searchInput');
  const clearBtn = document.getElementById('searchClear');
  if (!input || !clearBtn) return;
  input.addEventListener('input', () => {
    clearBtn.style.display = input.value ? 'flex' : 'none';
  });
  clearBtn.addEventListener('click', () => {
    input.value = '';
    clearBtn.style.display = 'none';
    input.dispatchEvent(new Event('input'));
    input.focus();
  });
})();

/* ─── CLOSE OUT SALES MODAL ─── */
(function () {
  const overlay  = document.getElementById('closeOutOverlay');
  const modal    = document.getElementById('closeOutModal');
  const closeBtn = document.getElementById('closeOutClose');
  const cancelBtn = document.getElementById('closeOutCancel');
  if (!overlay || !modal) return;

  // ── Static data ────────────────────────────────────────────────
  const OPERATORS  = ['TUI Group','Thomas Cook','Sunwing','Club Med','Jet2 Holidays'];
  const ROOM_TYPES = ['Standard Double','Superior Double','Junior Suite','Suite','Deluxe Ocean View'];
  const BOARD_TYPES = ['All Inclusive','Full Board','Half Board','Bed & Breakfast','Room Only'];

  let drIdSeq   = 0;  // date-range id counter
  let ruleIdSeq = 0;  // rule id counter

  // ── Helpers ────────────────────────────────────────────────────
  function pad(n) { return String(n).padStart(2,'0'); }

  function buildChips(items, selectedSet, ruleId, field) {
    var allActive = selectedSet.has('all');
    var html = '<div class="co-chips-wrap">';
    var allCls = 'co-chip co-chip-all' + (allActive ? ' active' : '');
    html += '<span class="' + allCls + '" data-rid="' + ruleId + '" data-fld="' + field + '" data-val="__all__" onclick="coChipClick(this)">'
      + (allActive ? '&#10003; ' : '') + 'All</span>';
    items.forEach(function(item) {
      var active = !allActive && selectedSet.has(item);
      var cls = 'co-chip' + (active ? ' active' : '');
      var safe = item.replace(/&/g,'&amp;').replace(/"/g,'&quot;');
      html += '<span class="' + cls + '" data-rid="' + ruleId + '" data-fld="' + field + '" data-val="' + safe + '" onclick="coChipClick(this)">'
        + (active ? '&#10003; ' : '') + item + '</span>';
    });
    html += '</div>';
    return html;
  }

  // ── Date Range rows ────────────────────────────────────────────
  var dateRanges = []; // [{id, from, to}]

  window.coAddDateRange = function(fromVal, toVal) {
    var id = ++drIdSeq;
    dateRanges.push({ id: id, from: fromVal || '', to: toVal || '' });
    renderDateRanges();
  };

  window.coRemoveDateRange = function(id) {
    dateRanges = dateRanges.filter(function(dr) { return dr.id !== id; });
    renderDateRanges();
  };

  window.coDRChange = function(id, field, val) {
    var dr = dateRanges.find(function(d) { return d.id === id; });
    if (dr) dr[field] = val;
  };

  function fmtDRDate(iso) {
    if (!iso) return '';
    var parts = iso.split('-');
    return parts[1] + '/' + parts[2] + '/' + parts[0];
  }

  // ── Close-out date range picker (shared panel) ─────────────
  var _coDRPickState = { drId: null, from: null, to: null, pickingTo: false, hover: null, viewYear: 2026, viewMonth: 2 };
  var _MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  var _DOWS = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];

  function _coDRFmt(d) { return (d.getMonth()+1)+'/'+d.getDate()+'/'+d.getFullYear(); }
  function _coDRSame(a,b) { return a&&b&&a.getFullYear()===b.getFullYear()&&a.getMonth()===b.getMonth()&&a.getDate()===b.getDate(); }

  function _coDRBuildMonth(y,m) {
    var first = new Date(y,m,1).getDay(), days = new Date(y,m+1,0).getDate();
    var h = '<div><div style="display:grid;grid-template-columns:repeat(7,36px);gap:2px 0;margin-bottom:2px">';
    _DOWS.forEach(function(d){ h+='<div class="caldr-dow">'+d+'</div>'; });
    h += '</div><div style="display:grid;grid-template-columns:repeat(7,36px);gap:2px 0">';
    for(var i=0;i<first;i++) h+='<div class="caldr-day caldr-empty"></div>';
    for(var d=1;d<=days;d++){
      var ts=new Date(y,m,d).getTime();
      h+='<div class="caldr-day" data-ts="'+ts+'" onclick="coDRPickDay('+ts+')">'+d+'</div>';
    }
    h+='</div></div>'; return h;
  }

  function _coDRRefresh() {
    var s=_coDRPickState, grids=document.getElementById('coDRPickGrids');
    if(!grids) return;
    var rangeEnd=s.pickingTo?(s.hover||null):s.to;
    grids.querySelectorAll('.caldr-day[data-ts]').forEach(function(el){
      var dt=new Date(parseInt(el.dataset.ts));
      el.className='caldr-day';
      var today=new Date(2026,2,29);
      if(_coDRSame(dt,today)) el.classList.add('caldr-today');
      if(s.from&&_coDRSame(dt,s.from)) el.classList.add('caldr-start');
      if(rangeEnd&&_coDRSame(dt,rangeEnd)) el.classList.add('caldr-end');
      if(s.from&&rangeEnd&&!_coDRSame(s.from,rangeEnd)){
        var lo=s.from<rangeEnd?s.from:rangeEnd, hi=s.from<rangeEnd?rangeEnd:s.from;
        if(dt>=lo&&dt<=hi) el.classList.add('caldr-in-range');
      }
    });
    var banner=document.getElementById('coDRPickBanner');
    if(banner) banner.style.display=s.pickingTo?'':'none';
    var foot=document.getElementById('coDRPickFooter');
    if(foot){
      if(s.from&&(s.to||rangeEnd)) foot.textContent=_coDRFmt(s.from)+' \u2013 '+_coDRFmt(rangeEnd||s.to);
      else if(s.from) foot.textContent=_coDRFmt(s.from)+' \u2013 ... (click end date)';
      else foot.textContent='Select start date';
    }
  }

  function _coDRRender() {
    var s=_coDRPickState;
    var lbl1=document.getElementById('coDRLeftLbl'), lbl2=document.getElementById('coDRRightLbl');
    var grids=document.getElementById('coDRPickGrids');
    if(!grids) return;
    var m2m=s.viewMonth+1, m2y=s.viewYear;
    if(m2m>11){m2m=0;m2y++;}
    if(lbl1) lbl1.textContent=_MONTHS[s.viewMonth]+' '+s.viewYear;
    if(lbl2) lbl2.textContent=_MONTHS[m2m]+' '+m2y;
    grids.innerHTML=_coDRBuildMonth(s.viewYear,s.viewMonth)+_coDRBuildMonth(m2y,m2m);
    _coDRRefresh();
    grids.querySelectorAll('.caldr-day[data-ts]').forEach(function(el){
      el.addEventListener('mouseenter',function(){
        if(s.pickingTo){s.hover=new Date(parseInt(this.dataset.ts));_coDRRefresh();}
      });
    });
    grids.addEventListener('mouseleave',function(){if(s.pickingTo){s.hover=null;_coDRRefresh();}});
  }

  window.coDRPickDay = function(ts) {
    var s=_coDRPickState, dt=new Date(ts);
    if(!s.pickingTo){s.from=dt;s.to=null;s.hover=null;s.pickingTo=true;}
    else{
      if(_coDRSame(dt,s.from)){s.to=dt;}
      else if(dt<s.from){s.to=s.from;s.from=dt;}
      else{s.to=dt;}
      s.pickingTo=false;s.hover=null;
    }
    _coDRRefresh();
  };

  window.coDRPickNav = function(delta) {
    var s=_coDRPickState;
    s.viewMonth+=delta;
    while(s.viewMonth>11){s.viewMonth-=12;s.viewYear++;}
    while(s.viewMonth<0){s.viewMonth+=12;s.viewYear--;}
    _coDRRender();
  };

  window.coDRPickCancel = function() {
    document.getElementById('coDRPanel').style.display='none';
    _coDRPickState.pickingTo=false;_coDRPickState.hover=null;
  };

  window.coDRPickApply = function() {
    var s=_coDRPickState;
    if(!s.from||!s.to) return;
    document.getElementById('coDRPanel').style.display='none';
    s.pickingTo=false;
    var dr=dateRanges.find(function(d){return d.id===s.drId;});
    if(dr){
      var pad2=function(n){return n<10?'0'+n:''+n;};
      dr.from=s.from.getFullYear()+'-'+pad2(s.from.getMonth()+1)+'-'+pad2(s.from.getDate());
      dr.to=s.to.getFullYear()+'-'+pad2(s.to.getMonth()+1)+'-'+pad2(s.to.getDate());
    }
    renderDateRanges();
  };

  window.coDRPickPreset = function(key) {
    var s=_coDRPickState;
    var from=new Date(2026,2,29);from.setHours(0,0,0,0);
    var to=new Date(from);
    if(key==='today'){to=new Date(from);}
    else if(key==='7d'){to.setDate(to.getDate()+6);}
    else if(key==='14d'){to.setDate(to.getDate()+13);}
    else if(key==='1m'){to.setMonth(to.getMonth()+1);to.setDate(to.getDate()-1);}
    else if(key==='2m'){to.setMonth(to.getMonth()+2);to.setDate(to.getDate()-1);}
    else if(key==='3m'){to.setMonth(to.getMonth()+3);to.setDate(to.getDate()-1);}
    s.from=from;s.to=to;s.pickingTo=false;s.hover=null;
    _coDRRender();
  };

  function openCoDRPicker(drId, triggerEl) {
    var s=_coDRPickState;
    s.drId=drId;
    var dr=dateRanges.find(function(d){return d.id===drId;});
    if(dr&&dr.from){var p=dr.from.split('-');s.from=new Date(+p[0],+p[1]-1,+p[2]);s.viewYear=+p[0];s.viewMonth=+p[1]-1;}
    else{s.from=null;s.viewYear=2026;s.viewMonth=2;}
    if(dr&&dr.to){var p2=dr.to.split('-');s.to=new Date(+p2[0],+p2[1]-1,+p2[2]);}else{s.to=null;}
    s.pickingTo=false;s.hover=null;
    var panel=document.getElementById('coDRPanel');
    if(!panel) return;
    var rect=triggerEl.getBoundingClientRect();
    var panelW=Math.min(720,window.innerWidth*0.95);
    var left=rect.left;
    if(left+panelW>window.innerWidth-8) left=Math.max(8,window.innerWidth-panelW-8);
    var top=rect.bottom+6;
    if(top+400>window.innerHeight) top=Math.max(8,rect.top-410);
    panel.style.left=left+'px';
    panel.style.top=top+'px';
    panel.style.width=panelW+'px';
    panel.style.display='block';
    _coDRRender();
  }

  // Close picker on outside click
  document.addEventListener('click',function(e){
    var panel=document.getElementById('coDRPanel');
    if(!panel||panel.style.display==='none') return;
    if(panel.contains(e.target)) return;
    if(e.target.closest('.co2-dr-trigger')) return;
    panel.style.display='none';
    _coDRPickState.pickingTo=false;
  },true);

  function renderDateRanges() {
    var list = document.getElementById('coDateRangeList');
    if (!list) return;
    list.innerHTML = dateRanges.map(function(dr, idx) {
      var label = (fmtDRDate(dr.from) || 'Start') + ' - ' + (fmtDRDate(dr.to) || 'End');
      return '<div class="co2-dr-wrap">'
        + '<span class="co2-dr-label">Date Range ' + (idx + 1) + '</span>'
        + '<div class="co2-dr-trigger" data-drid="' + dr.id + '">'
        + '<span class="material-icons co2-dr-cal-ico">calendar_today</span>'
        + '<span class="co2-dr-text">' + label + '</span>'
        + (dateRanges.length > 1 ? '<button type="button" class="co2-dr-remove" data-drid="' + dr.id + '" onclick="event.stopPropagation();coRemoveDateRange(+this.dataset.drid)" title="Remove">&times;</button>' : '')
        + '</div></div>';
    }).join('');

    // Click trigger opens the date picker popup
    list.querySelectorAll('.co2-dr-trigger').forEach(function(trig) {
      trig.addEventListener('click', function(e) {
        if (e.target.closest('.co2-dr-remove')) return;
        var drId = parseInt(trig.dataset.drid);
        openCoDRPicker(drId, trig);
      });
    });
  }

  // ── Restriction Strategies ─────────────────────────────────────────
  var rules = [];

  window.coAddRule = function() {
    var id = ++ruleIdSeq;
    rules.push({ id: id, ops: new Set(['all']), rooms: new Set(['all']), boards: new Set(['all']) });
    renderRules();
  };

  window.coRemoveRule = function(id) {
    rules = rules.filter(function(r) { return r.id !== id; });
    renderRules();
  };

  // Chip click — uses data attributes, no quoting issues
  window.coChipClick = function(el) {
    var ruleId = parseInt(el.dataset.rid);
    var field  = el.dataset.fld;
    var value  = el.dataset.val === '__all__' ? 'all' : el.dataset.val;
    var rule = rules.find(function(r) { return r.id === ruleId; });
    if (!rule) return;
    var set = rule[field];
    if (value === 'all') {
      set.clear(); set.add('all');
    } else {
      set.delete('all');
      if (set.has(value)) set.delete(value);
      else set.add(value);
      if (set.size === 0) set.add('all');
    }
    renderRules();
  };

  function buildMSDropdown(items, selectedSet, ruleId, field) {
    var trigText = selectedSet.has('all') ? 'All'
      : selectedSet.size <= 2 ? Array.from(selectedSet).join(', ')
      : selectedSet.size + ' selected';
    var ddItems = items.map(function(item) {
      var isOn = selectedSet.has(item);
      return '<label class="co2-ms-item"><input type="checkbox" value="' + item + '"'
        + (isOn ? ' checked' : '')
        + ' data-rid="' + ruleId + '" data-fld="' + field + '"'
        + ' onchange="coMSChange(this)">' + item + '</label>';
    }).join('');
    return '<div class="co2-ms-wrap" data-rid="' + ruleId + '" data-fld="' + field + '">'
      + '<div class="co2-ms-trigger" onclick="coMSToggle(this)">'
      + '<span class="co2-ms-text">' + trigText + '</span>'
      + '<span class="material-icons co2-select-arrow">arrow_drop_down</span></div>'
      + '<div class="co2-ms-list">' + ddItems + '</div>'
      + '</div>';
  }

  function buildChipsForSet(selectedSet, ruleId, field) {
    if (selectedSet.has('all') || selectedSet.size === 0) return '';
    return '<div class="co2-ms-chips">' + Array.from(selectedSet).map(function(v) {
      return '<span class="co2-ms-chip">' + v
        + '<span class="co2-ms-chip-x" data-rid="' + ruleId + '" data-fld="' + field + '" data-val="' + v + '" onclick="coMSRemoveChip(this)">&times;</span>'
        + '</span>';
    }).join('') + '</div>';
  }

  function renderRules() {
    var list = document.getElementById('coRuleList');
    if (!list) return;
    list.innerHTML = rules.map(function(rule, idx) {
      return '<div class="co2-strategy-group">'
        + (rules.length > 1 ? '<button type="button" class="co2-strategy-remove" data-ruleid="' + rule.id + '" onclick="coRemoveRule(+this.dataset.ruleid)" title="Remove strategy">&times; Remove</button>' : '')
        + '<div class="co2-field-group"><label class="co2-field-label">Operators</label>'
        + buildMSDropdown(OPERATORS, rule.ops, rule.id, 'ops')
        + buildChipsForSet(rule.ops, rule.id, 'ops') + '</div>'
        + '<div class="co2-field-group"><label class="co2-field-label">Room Types</label>'
        + buildMSDropdown(ROOM_TYPES, rule.rooms, rule.id, 'rooms')
        + buildChipsForSet(rule.rooms, rule.id, 'rooms') + '</div>'
        + '<div class="co2-field-group"><label class="co2-field-label">Meal Plans</label>'
        + buildMSDropdown(BOARD_TYPES, rule.boards, rule.id, 'boards')
        + buildChipsForSet(rule.boards, rule.id, 'boards') + '</div>'
        + '</div>';
    }).join('');
  }

  window.coMSToggle = function(trigger) {
    var wrap = trigger.closest('.co2-ms-wrap');
    var list = wrap.querySelector('.co2-ms-list');
    var isOpen = list.classList.contains('open');
    // Close all others
    document.querySelectorAll('.co2-ms-list.open').forEach(function(l) { l.classList.remove('open'); });
    document.querySelectorAll('.co2-ms-trigger.open').forEach(function(t) { t.classList.remove('open'); });
    if (!isOpen) {
      list.classList.add('open');
      trigger.classList.add('open');
    }
  };

  window.coMSChange = function(cb) {
    var ruleId = parseInt(cb.dataset.rid);
    var field = cb.dataset.fld;
    var rule = rules.find(function(r) { return r.id === ruleId; });
    if (!rule) return;
    var val = cb.value;
    if (cb.checked) {
      rule[field].delete('all');
      rule[field].add(val);
    } else {
      rule[field].delete(val);
    }
    if (rule[field].size === 0) rule[field].add('all');
    renderRules();
  };

  window.coMSRemoveChip = function(el) {
    var ruleId = parseInt(el.dataset.rid);
    var field = el.dataset.fld;
    var val = el.dataset.val;
    var rule = rules.find(function(r) { return r.id === ruleId; });
    if (!rule) return;
    rule[field].delete(val);
    if (rule[field].size === 0) rule[field].add('all');
    renderRules();
  };

  // Close dropdowns on click outside
  document.addEventListener('click', function(e) {
    if (!e.target.closest('.co2-ms-wrap')) {
      document.querySelectorAll('.co2-ms-list.open').forEach(function(l) { l.classList.remove('open'); });
      document.querySelectorAll('.co2-ms-trigger.open').forEach(function(t) { t.classList.remove('open'); });
    }
  });


  // ── Open modal ─────────────────────────────────────────────────
  function resetModalState() {
    const title = document.getElementById('closeOutTitle');
    if (title) title.textContent = 'Close or re-open sales';
    const confirmBtn = document.getElementById('coConfirmBtn');
    if (confirmBtn) { confirmBtn.textContent = 'Close Out'; confirmBtn.style.background = ''; confirmBtn.style.borderColor = ''; }
    document.querySelectorAll('.co2-type-card').forEach(function(c) { c.classList.remove('active'); });
    const firstCard = document.querySelector('.co2-type-card[data-type="full"]');
    if (firstCard) firstCard.classList.add('active');
    const losField = document.getElementById('coLosField');
    if (losField) losField.style.display = 'none';
    // Reset send action radio to email
    var emailRadio = document.querySelector('input[name="coSendAction"][value="email"]');
    if (emailRadio) emailRadio.checked = true;
  }

  function openModal(fromDate, toDate, ctx) {
    overlay.classList.add('open');
    resetModalState();

    // Reset date ranges
    drIdSeq = 0; dateRanges = [];
    var from = fromDate, to = toDate;
    if (!from || !to) {
      // Auto-fill from calendar or weekly selection
      const rs = (calSelStart && calSelEnd) ? calSelStart : (wvSelStart && wvSelEnd) ? wvSelStart : null;
      const re = (calSelStart && calSelEnd) ? calSelEnd   : (wvSelStart && wvSelEnd) ? wvSelEnd   : null;
      if (rs && re) {
        const sv = rs.month*100+rs.day, ev = re.month*100+re.day;
        const lo = sv<=ev?rs:re, hi = sv<=ev?re:rs;
        from = '2026-'+pad(lo.month)+'-'+pad(lo.day);
        to   = '2026-'+pad(hi.month)+'-'+pad(hi.day);
      }
    }
    coAddDateRange(from || '', to || '');

    // Pre-populate rules from active filter state
    ruleIdSeq = 0; rules = [];
    _coPrePopulateRule(ctx || 'wv');
  }

  function closeModal() { overlay.classList.remove('open'); }

  // ── Event wiring ───────────────────────────────────────────────
  document.addEventListener('click', e => {
    const _coTrigger = e.target.closest('.popup-btn-closeout, .wv-lock-btn');
    if (_coTrigger) {
      // Detect which view triggered the modal: cal = monthly calendar, wv = weekly/daily view
      var _coCtx = _coTrigger.classList.contains('popup-btn-closeout') ? 'cal' : 'wv';
      openModal(undefined, undefined, _coCtx);
    }
  });

  closeBtn && closeBtn.addEventListener('click', closeModal);
  cancelBtn && cancelBtn.addEventListener('click', closeModal);
  overlay.addEventListener('click', e => { if (e.target === overlay) closeModal(); });

  // ── Action selector (now radio buttons) ────────────────────────
  function getSelectedAction() {
    var checked = document.querySelector('input[name="coSendAction"]:checked');
    return checked ? checked.value : 'email';
  }

  // ── Restriction type card toggle ──────────────────────────────
  document.getElementById('coTypeGroup')?.addEventListener('click', e => {
    const btn = e.target.closest('.co2-type-card');
    if (!btn) return;
    document.querySelectorAll('.co2-type-card').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    const type = btn.dataset.type;
    const losField = document.getElementById('coLosField');
    if (losField) losField.style.display = type === 'los' ? '' : 'none';
    const isReopen = type === 'reopen';
    const title = document.getElementById('closeOutTitle');
    if (title) title.textContent = 'Close or re-open sales';
    const confirmBtn = document.getElementById('coConfirmBtn');
    if (confirmBtn) {
      confirmBtn.textContent = isReopen ? 'Re-Open' : 'Close Out';
      confirmBtn.style.background = '';
      confirmBtn.style.borderColor = '';
    }
  });

  // ── Confirm ────────────────────────────────────────────────────
  document.getElementById('coConfirmBtn').addEventListener('click', () => {
    const activeCard = document.querySelector('.co2-type-card.active');
    const isReopen   = activeCard && activeCard.dataset.type === 'reopen';
    const action     = getSelectedAction();
    const email      = document.getElementById('coEmail')?.value || '';
    const message    = document.getElementById('coMessage')?.value || '';

    // Parse ISO date string as local time (not UTC) to avoid off-by-one day in non-UTC zones
    function parseLocalDate(iso) {
      var p = iso.split('-');
      return new Date(+p[0], +p[1]-1, +p[2]);
    }

    // Apply calendar lock/unlock for each date range
    dateRanges.forEach(function(dr) {
      if (!dr.from || !dr.to) return;
      var cur = parseLocalDate(dr.from), end = parseLocalDate(dr.to);
      // Ensure correct order
      if (cur > end) { var tmp = cur; cur = end; end = tmp; }
      while (cur <= end) {
        const m = cur.getMonth()+1, d = cur.getDate();
        if (isReopen) LOCKED_DAYS.delete(m+'-'+d);
        else          LOCKED_DAYS.add(m+'-'+d);
        cur.setDate(cur.getDate()+1);
      }
    });

    // Simulate actions (prototype feedback)
    if (action === 'email' || action === 'both') {
      const rulesSummary = rules.map(function(r, i) {
        const ops = r.ops.has('all') ? 'All Operators' : Array.from(r.ops).join(', ');
        return 'Strategy ' + (i+1) + ': ' + ops;
      }).join(' | ');
      console.log('[Close Out] Email action — recipients:', rulesSummary, '| message:', message || '(none)');
    }
    if (action === 'internal' || action === 'both') {
      console.log('[Close Out] Internal note recorded — message:', message || '(none)');
    }

    // Clear selected days after close-out
    _moSelectedDays.clear();
    _wbSelectedDays.clear();
    _wvSelectedDays.clear();
    _syncCloseOutBtn();

    renderCalendar();
    buildWeekGrid(wvMonth, wvWeekStart, wvWeekStart);
    closeModal();
  });

  // ── Shared: build a pre-populated rule from the active view's filter state ──
  function _coPrePopulateRule(ctx) {
    var _TO_MAP    = { sunwing: 'Sunwing', tui: 'TUI Group', 'thomas-cook': 'Thomas Cook', 'club-med': 'Club Med' };
    var _ROOM_MAP  = { standard: 'Standard Double', superior: 'Superior Double', deluxe: 'Deluxe Ocean View', suite: 'Suite' };
    var _BOARD_MAP = { ai: 'All Inclusive', bb: 'Bed & Breakfast', ro: 'Room Only', hb: 'Half Board', fb: 'Full Board' };
    var _opsSet = new Set(['all']), _roomsSet = new Set(['all']), _boardsSet = new Set(['all']);
    if (typeof filterState !== 'undefined') {
      // 'cal' = monthly calendar view, 'wv' = weekly/daily view
      var _isCal = (ctx === 'cal');
      var _filtTO    = _isCal ? filterState.cal.calFiltTO    : filterState.wv.wvFiltTO;
      var _filtRoom  = _isCal ? filterState.cal.calFiltRoom  : filterState.wv.wvFiltRoom;
      var _filtBoard = _isCal ? filterState.cal.calFiltBoard : filterState.wv.wvFiltBoard;
      // Handle comma-separated multi-select values
      function _mapMulti(raw, map) {
        if (!raw || raw === 'all') return new Set(['all']);
        var mapped = raw.split(',').map(function(v){ return map[v.trim()]; }).filter(Boolean);
        return mapped.length ? new Set(mapped) : new Set(['all']);
      }
      _opsSet   = _mapMulti(_filtTO,    _TO_MAP);
      _roomsSet = _mapMulti(_filtRoom,  _ROOM_MAP);
      _boardsSet = _mapMulti(_filtBoard, _BOARD_MAP);
    }
    rules.push({ id: ++ruleIdSeq, ops: _opsSet, rooms: _roomsSet, boards: _boardsSet });
    renderRules();
  }

  // Open modal with specific individual days (not a range)
  function openModalDays(daysArr, ctx) {
    overlay.classList.add('open');
    resetModalState();

    // Add each selected day as its own date range (from=to)
    drIdSeq = 0; dateRanges = [];
    daysArr.forEach(function(d) {
      coAddDateRange(d, d);
    });

    // Pre-populate rules from active filter state
    ruleIdSeq = 0; rules = [];
    _coPrePopulateRule(ctx || 'wv');
  }

  // Expose openModal globally so inline calls work
  window._coOpenModal = openModal;
  window._coOpenModalDays = openModalDays;
})();

/* ─── MONTHLY FILTER DROPDOWN — toggle open/close ─── */
(function() {
  const filtersBtn = document.getElementById('calFiltersBtn');
  const filtersDd  = document.getElementById('calFiltersDropdown');
  if (!filtersBtn || !filtersDd) return;

  filtersBtn.addEventListener('click', function(e) {
    e.stopPropagation();
    const isOpen = filtersDd.style.display !== 'none';
    if (!isOpen) _closeAllDropdowns('calFiltersDropdown');
    filtersDd.style.display = isOpen ? 'none' : '';
    filtersBtn.classList.toggle('active', !isOpen);
  });

  document.addEventListener('click', function(e) {
    if (!document.getElementById('calFiltersWrap')?.contains(e.target)) {
      if (filtersDd) filtersDd.style.display = 'none';
      filtersBtn.classList.remove('active');
    }
  });
})();

/* ─── FILTER STATE (for list-form filters) ─── */
const filterState = {
  cal: { calFiltTO: 'all', calFiltRoom: 'all', calFiltBoard: 'all', calFiltMarket: 'all', calFiltPickup: 'all' },
  wv:  { wvFiltTO:  'all', wvFiltRoom:  'all', wvFiltBoard:  'all', wvFiltMarket:  'all', wvFiltPickup:  'all' },
};

/* Update pickup button/input active state to match current filterState */
function _syncPickupBtnUI(panel) {
  var val    = panel === 'cal' ? filterState.cal.calFiltPickup : filterState.wv.wvFiltPickup;
  var wrap   = document.getElementById(panel === 'cal' ? 'calPickupBtns'  : 'wvPickupBtns');
  var lbl    = document.getElementById(panel === 'cal' ? 'calPickupLabel' : 'wvPickupLabel');
  var hidden = document.getElementById(panel === 'cal' ? 'calFiltPickup'  : 'wvFiltPickup');
  var isAll  = (!val || val === 'all' || val === '365');
  if (hidden) hidden.value = isAll ? 'all' : String(val);
  if (lbl)    lbl.textContent = isAll ? 'All time' : (val == '1' ? '1 day' : val + ' days');
  if (wrap) {
    // All button
    wrap.querySelectorAll('.pickup-day-btn').forEach(function(b) {
      b.classList.toggle('active', isAll && b.dataset.val === 'all');
    });
    // Number inputs — activate whichever input currently shows this value
    var matched = false;
    wrap.querySelectorAll('.pickup-day-input').forEach(function(inp) {
      var matches = !isAll && String(inp.value) === String(val);
      inp.classList.toggle('active', matches);
      if (matches) matched = true;
    });
    // If a custom value (not matching any preset input), activate first input and set its value
    if (!isAll && !matched) {
      var inputs = wrap.querySelectorAll('.pickup-day-input');
      if (inputs.length) {
        inputs.forEach(function(i) { i.classList.remove('active'); });
        inputs[0].value = val;
        inputs[0].classList.add('active');
      }
    }
  }
}

/* Sync filters between cal↔wv so switching views preserves selections */
function syncFiltersCalToWv() {
  filterState.wv.wvFiltTO     = filterState.cal.calFiltTO;
  filterState.wv.wvFiltRoom   = filterState.cal.calFiltRoom;
  filterState.wv.wvFiltBoard  = filterState.cal.calFiltBoard;
  filterState.wv.wvFiltMarket = filterState.cal.calFiltMarket;
  filterState.wv.wvFiltPickup = filterState.cal.calFiltPickup;
}
function syncFiltersWvToCal() {
  filterState.cal.calFiltTO     = filterState.wv.wvFiltTO;
  filterState.cal.calFiltRoom   = filterState.wv.wvFiltRoom;
  filterState.cal.calFiltBoard  = filterState.wv.wvFiltBoard;
  filterState.cal.calFiltMarket = filterState.wv.wvFiltMarket;
  filterState.cal.calFiltPickup = filterState.wv.wvFiltPickup;
  calFiltTO = (filterState.cal.calFiltTO === 'all' || !filterState.cal.calFiltTO) ? 'all' : filterState.cal.calFiltTO.split(',')[0];
}

function getFilterVal(id) {
  const ctx = (id.startsWith('wv')) ? filterState.wv : filterState.cal;
  return ctx[id] || 'all';
}

function applyFilterUI(dropdownId) {
  const dd = document.getElementById(dropdownId);
  if (!dd) return;
  dd.querySelectorAll('.wv-fi-rb').forEach(function(rb) {
    const fid = rb.dataset.fid, val = rb.dataset.val;
    const ctx = (fid.startsWith('wv')) ? filterState.wv : filterState.cal;
    const cur = ctx[fid] || 'all';
    if (val === 'all') {
      rb.classList.toggle('checked', cur === 'all');
    } else {
      rb.classList.toggle('checked', cur !== 'all' && cur.split(',').indexOf(val) !== -1);
    }
  });
  // update filter count badge
  const countId = dropdownId === 'calFiltersDropdown' ? 'calFilterCount' : 'wvFilterCount';
  const badge = document.getElementById(countId);
  if (badge) {
    const ctx = dropdownId === 'calFiltersDropdown' ? filterState.cal : filterState.wv;
    const n = Object.keys(ctx).filter(function(k){ return ctx[k] !== 'all'; }).length;
    badge.textContent = n;
    badge.style.display = n > 0 ? '' : 'none';
  }
}

// Multiselect checkbox click handler — attached directly to each filter dropdown
// so it fires before any document-level outside-click close handler can see the event.
function _handleWvFiRbClick(e) {
  const rb = e.target.closest('.wv-fi-rb');
  if (!rb) return;
  e.stopPropagation(); // keep dropdown open
  const fid = rb.dataset.fid;
  const val = rb.dataset.val;
  if (!fid || !val) return;
  const ctx = (fid.startsWith('wv')) ? filterState.wv : filterState.cal;
  const dd = rb.closest('.cal-filters-dropdown');

  if (val === 'all') {
    ctx[fid] = 'all';
  } else {
    const current = ctx[fid] || 'all';
    const parts = current === 'all' ? [] : current.split(',');
    const idx = parts.indexOf(val);
    if (idx === -1) parts.push(val);
    else            parts.splice(idx, 1);
    ctx[fid] = parts.length > 0 ? parts.join(',') : 'all';
  }

  if (dd) {
    dd.querySelectorAll('.wv-fi-rb[data-fid="' + fid + '"]').forEach(function(item) {
      const v = item.dataset.val;
      const cur = ctx[fid] || 'all';
      if (v === 'all') {
        item.classList.toggle('checked', cur === 'all');
      } else {
        item.classList.toggle('checked', cur !== 'all' && cur.split(',').indexOf(v) !== -1);
      }
    });
    const countId = dd.id === 'calFiltersDropdown' ? 'calFilterCount' : 'wvFilterCount';
    const badge = document.getElementById(countId);
    if (badge) {
      const stateCtx = dd.id === 'calFiltersDropdown' ? filterState.cal : filterState.wv;
      const n = Object.keys(stateCtx).filter(function(k){ return stateCtx[k] !== 'all' && stateCtx[k] !== '365'; }).length;
      badge.textContent = n;
      badge.style.display = n > 0 ? '' : 'none';
    }
  }

  if (fid === 'calFiltTO') {
    const cur = ctx[fid];
    calFiltTO = (cur === 'all' || !cur) ? 'all' : cur.split(',')[0];
    renderCalendar();
  }
}
// Attach to each filter dropdown element so stopPropagation prevents outside-close handlers
['calFiltersDropdown', 'wvFiltersDropdown'].forEach(function(id) {
  var el = document.getElementById(id);
  if (el) el.addEventListener('click', _handleWvFiRbClick);
});

// Apply button — re-renders
document.getElementById('calFilterApply')?.addEventListener('click', function() {
  calFiltTO = filterState.cal.calFiltTO;
  const dd = document.getElementById('calFiltersDropdown');
  if (dd) dd.style.display = 'none';
  document.getElementById('calFiltersBtn')?.classList.remove('active');
  renderCalendar();
});
document.getElementById('wvFilterApply')?.addEventListener('click', function() {
  buildWeekGrid(wvMonth, wvWeekStart, wvWeekStart);
  const dd = document.getElementById('wvFiltersDropdown');
  if (dd) dd.style.display = 'none';
  document.getElementById('wvFiltersBtn')?.classList.remove('active');
});

// Reset buttons
document.getElementById('calFilterReset')?.addEventListener('click', function() {
  Object.keys(filterState.cal).forEach(function(k){ filterState.cal[k] = 'all'; });
  calFiltTO = 'all';
  pickupBtnReset('cal');
  applyFilterUI('calFiltersDropdown');
  renderCalendar();
});
document.getElementById('wvFilterReset')?.addEventListener('click', function() {
  Object.keys(filterState.wv).forEach(function(k){ filterState.wv[k] = 'all'; });
  pickupBtnReset('wv');
  applyFilterUI('wvFiltersDropdown');
  buildWeekGrid(wvMonth, wvWeekStart, wvWeekStart);
});

/* ─── WEEKLY METRIC TABS ─── */
document.getElementById('wvMetricTabs')?.addEventListener('click', function(e) {
  const btn = e.target.closest('.cal-metric-tab');
  if (!btn) return;
  document.querySelectorAll('#wvMetricTabs .cal-metric-tab').forEach(function(b){ b.classList.remove('active'); });
  btn.classList.add('active');
  wvActiveTab = btn.dataset.metric || 'occupancy';
  buildWeekGrid(wvMonth, wvWeekStart, wvWeekStart);
});

/* ─── COMP MODE PILLS (delegated) ─── */
document.addEventListener('click', function(e) {
  const pill = e.target.closest('.wv-comp-pill');
  if (!pill) return;
  e.stopPropagation();
  wvCompMode = pill.dataset.comp || 'sdly';
  buildWeekGrid(wvMonth, wvWeekStart, wvWeekStart);
});

/* ─── METRIC TABS ─── */
document.getElementById('calMetricTabs')?.addEventListener('click', e => {
  const btn = e.target.closest('.cal-metric-tab');
  if (!btn) return;
  document.querySelectorAll('.cal-metric-tab').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  calMetric = btn.dataset.metric;
  renderCalendar();
});

// ── Monthly cell metrics selector ──────────────────────────
document.addEventListener('click', function(e) {
  if (e.target.closest('#calMetricsBtn')) {
    var btn = document.getElementById('calMetricsBtn');
    var rect = btn.getBoundingClientRect();

    // ── Daily R mode: show AG Grid column panel ──────────────
    if (wvGroupBy === 'report' && _dailyRevGridApi) {
      var drp = document.getElementById('dailyRevColPanel');
      if (drp) {
        var isOpen = drp.style.display !== 'none';
        drp.style.display = isOpen ? 'none' : 'block';
        if (!isOpen) {
          drp.style.top  = (rect.bottom + 4) + 'px';
          drp.style.left = rect.left + 'px';
        }
        e.stopPropagation(); return;
      }
    }

    // ── Normal calendar mode ──────────────────────────────────
    var dd  = document.getElementById('calMetricsDropdown');
    var isOpen = dd.style.display !== 'none';
    if (isOpen) { dd.style.display = 'none'; e.stopPropagation(); return; }
    var maxH  = Math.min(window.innerHeight - rect.bottom - 12, window.innerHeight * 0.75);
    dd.style.top     = (rect.bottom + 4) + 'px';
    dd.style.left    = rect.left + 'px';
    dd.style.width = '320px';
    dd.style.maxHeight = maxH + 'px';
    dd.style.display = 'block';
    e.stopPropagation(); return;
  }
  var cb = e.target.closest('.cal-md-cb[data-cm-key]');
  if (cb) {
    e.stopPropagation(); // keep metrics dropdown open
    var key = cb.dataset.cmKey;
    var idx = calCellMetrics.indexOf(key);
    if (idx > -1) {
      if (calCellMetrics.length > 1) { calCellMetrics.splice(idx, 1); cb.classList.remove('checked'); }
    } else {
      if (calCellMetrics.length < 4) { calCellMetrics.push(key); cb.classList.add('checked'); }
      else { alert('Maximum 4 metrics. Remove one first.'); }
    }
    renderCalendar(); return;
  }
  if (!e.target.closest('#calMetricsWrap') && !e.target.closest('#dailyRevColPanel')) {
    var dd2 = document.getElementById('calMetricsDropdown');
    if (dd2) dd2.style.display = 'none';
    var drp2 = document.getElementById('dailyRevColPanel');
    if (drp2) drp2.style.display = 'none';
  }
});

// ── Daily R column group toggle ──────────────────────────────
var _drColVisibility = {
  daily: true, avail: true, segs: true, biz: true, meals: false, tc: true
};
var _drColFields = {
  daily: ['occ_t','occ_h','occ_stly','occ_ly','occ_fcst','adr_t','adr_h','adr_diff','adr_stly','adr_ly','adr_fcst','rev_t','rev_h','rev_stly','rev_ly','rev_fcst','rp_t','rp_h','rp_stly','rp_ly','pk_t','pk_h'],
  avail: ['td_rms','td_pct','os_rms','os_pct','rem_rms','rem_pct','on_on','on_off'],
  segs:  ['fit_rms','fit_pct','dyn_rms','dyn_pct','ser_rms','ser_pct'],
  biz:   ['biz_to','biz_dir','biz_ota','biz_oth'],
  meals: ['mp_ai_h','mp_ai_t','mp_ai_pct','mp_bb_h','mp_bb_t','mp_bb_pct','mp_hb_h','mp_hb_t','mp_hb_pct','mp_ro_h','mp_ro_t','mp_ro_pct'],
  tc:    ['tc_0','tc_1','tc_2','tc_3','tc_4','tc_base']
};
window.drColToggle = function(group, checkbox) {
  if (!_dailyRevGridApi) return;
  var visible = checkbox.checked;
  _drColVisibility[group] = visible;
  _dailyRevGridApi.setColumnsVisible(_drColFields[group], visible);
};


// ── Bulk select ──────────────────────────────────────────────
document.getElementById('calBulkBtn')?.addEventListener('click', function() {
  bulkSelectMode = !bulkSelectMode;
  this.classList.toggle('active', bulkSelectMode);
  if (!bulkSelectMode) { bulkSelected.clear(); }
  document.getElementById('bulkBanner').style.display = bulkSelectMode ? 'flex' : 'none';
  renderCalendar();
});
document.getElementById('bulkCancelBtn')?.addEventListener('click', function() {
  bulkSelectMode = false; bulkSelected.clear();
  document.getElementById('calBulkBtn')?.classList.remove('active');
  document.getElementById('bulkBanner').style.display = 'none';
  renderCalendar();
});
document.getElementById('bulkReopenBtn')?.addEventListener('click', function() {
  bulkSelected.forEach(function(key) { LOCKED_DAYS.delete(key); });
  bulkSelected.clear(); bulkSelectMode = false;
  document.getElementById('calBulkBtn')?.classList.remove('active');
  document.getElementById('bulkBanner').style.display = 'none';
  renderCalendar();
});

// ── Filter chips ─────────────────────────────────────────────
document.addEventListener('click', function(e) {
  var chip = e.target.closest('.fchip[data-fq-filter]');
  if (!chip) return;
  var filterId = chip.dataset.fqFilter;
  var val = chip.dataset.fqValue;
  var el = document.getElementById(filterId);
  if (el) { el.value = val; el.dispatchEvent(new Event('change')); }
  // Update active state
  document.querySelectorAll('.fchip[data-fq-filter="' + filterId + '"]').forEach(function(c) {
    c.classList.toggle('fchip-active', c.dataset.fqValue === val);
  });
  // Show/hide active filter tags
  updateActiveChipTags();
});

function updateActiveChipTags() {
  var wrap = document.getElementById('fchipsActiveWrap');
  if (!wrap) return;
  var defs = [
    { id: 'calFiltRoom',   label: 'Room' },
    { id: 'calFiltBoard',  label: 'Board' },
    { id: 'calFiltMarket', label: 'Source Geo' },
    { id: 'calFiltPickup', label: 'Pickup' },
  ];
  var html = defs.map(function(d) {
    var el = document.getElementById(d.id);
    if (!el || el.value === 'all') return '';
    var text = (el.options && el.options[el.selectedIndex]) ? el.options[el.selectedIndex].text : el.value;
    return '<span class="fchip fchip-tag">' + d.label + ': ' + text + '<button class="fchip-x" data-close-filter="' + d.id + '">×</button></span>';
  }).filter(Boolean).join('');
  wrap.innerHTML = html;
  var bar = document.getElementById('filterChipsBar');
  if (bar) bar.style.display = html ? '' : 'none';
}

document.addEventListener('click', function(e) {
  var xBtn = e.target.closest('.fchip-x[data-close-filter]');
  if (!xBtn) return;
  var id = xBtn.dataset.closeFilter;
  var el = document.getElementById(id);
  if (el) { el.value = 'all'; el.dispatchEvent(new Event('change')); }
  updateActiveChipTags();
});

// ── Pickup day inputs + All button ───────────────────────────
function _pickupSetFilter(panel, val) {
  var isAll = (val === 'all');
  var labelText = isAll ? 'All time' : (val == 1 ? '1 day' : val + ' days');
  var hiddenVal = isAll ? 'all' : String(val);
  var lbl    = document.getElementById(panel === 'cal' ? 'calPickupLabel' : 'wvPickupLabel');
  var hidden = document.getElementById(panel === 'cal' ? 'calFiltPickup'  : 'wvFiltPickup');
  if (lbl)    lbl.textContent = labelText;
  if (hidden) { hidden.value = hiddenVal; hidden.dispatchEvent(new Event('change')); }
}
window.pickupInputFocus = function(input, panel) {
  var wrap = input.closest('.pickup-btns-wrap');
  if (wrap) {
    wrap.querySelectorAll('.pickup-day-btn').forEach(function(b)  { b.classList.remove('active'); });
    wrap.querySelectorAll('.pickup-day-input').forEach(function(i) { i.classList.remove('active'); });
    input.classList.add('active');
  }
  var val = parseInt(input.value) || 1;
  _pickupSetFilter(panel, val);
};
window.pickupInputChange = function(input, panel) {
  var val = parseInt(input.value);
  if (!val || val < 1) return;
  if (val > 365) { input.value = 365; val = 365; }
  _pickupSetFilter(panel, val);
  if (typeof renderPickupMetricItems === 'function') renderPickupMetricItems();
};
window.pickupBtnClick = function(btn, panel) {
  var wrap = btn.closest('.pickup-btns-wrap');
  if (wrap) {
    wrap.querySelectorAll('.pickup-day-btn').forEach(function(b)  { b.classList.remove('active'); });
    wrap.querySelectorAll('.pickup-day-input').forEach(function(i) { i.classList.remove('active'); });
    btn.classList.add('active');
  }
  _pickupSetFilter(panel, 'all');
};
function pickupBtnReset(panel) {
  var wrap = document.getElementById(panel === 'cal' ? 'calPickupBtns' : 'wvPickupBtns');
  if (wrap) {
    wrap.querySelectorAll('.pickup-day-input').forEach(function(i) { i.classList.remove('active'); });
    wrap.querySelectorAll('.pickup-day-btn').forEach(function(b) {
      b.classList.toggle('active', b.dataset.val === 'all');
    });
  }
  var lbl = document.getElementById(panel === 'cal' ? 'calPickupLabel' : 'wvPickupLabel');
  if (lbl) lbl.textContent = 'All time';
}

// ── Pickup metric items (dynamic labels based on input values) ───────────
var pickupDayValues = [1, 3, 7];
window.pickupDayValues = pickupDayValues; // expose globally for renderCalendar (which is at global scope)

// Build a 2-row grid: window numbers on top, values below
function _mkPickupGrid(getValFn) {
  var hdrs = '', vals = '', n = 0;
  pickupDayValues.forEach(function(dv, i) {
    if (!wvMetricState['dm_pickup_' + i]) return;
    n++;
    hdrs += '<div class="wv-pickup-hdr-cell">' + dv + '</div>';
    vals += '<div class="wv-pickup-val-cell">' + getValFn(dv, i) + '</div>';
  });
  if (!n) return '';
  return '<div class="wv-pickup-grid" style="grid-template-columns:repeat(' + n + ',1fr)">'
    + hdrs + vals + '</div>';
}
function getPickupInputValues() {
  // Try wv inputs first, then cal inputs
  var wrap = document.getElementById('wvPickupBtns') || document.getElementById('calPickupBtns');
  if (!wrap) return [1, 3, 7];
  var vals = [];
  wrap.querySelectorAll('.pickup-day-input').forEach(function(inp) {
    var v = parseInt(inp.value);
    if (v && v > 0) vals.push(v);
  });
  return vals.length === 3 ? vals : [1, 3, 7];
}
window.renderPickupMetricItems = function() {
  pickupDayValues = getPickupInputValues();
  window.pickupDayValues = pickupDayValues; // keep global in sync
  var container = document.getElementById('wvPickupMetricItems');
  if (!container) return;
  container.innerHTML = pickupDayValues.map(function(d, i) {
    var key   = 'dm_pickup_' + i;
    var label = 'Pickup · ' + d;
    if (!(key in wvMetricState)) wvMetricState[key] = true;
    var checked = wvMetricState[key] !== false;
    return '<label class="wv-ms-item"><span class="wv-ms-cb' + (checked ? ' checked' : '') + '" data-key="' + key + '"></span><span class="wv-ms-label">' + label + '</span></label>';
  }).join('');
  // Keep dm_pickup master in sync
  wvMetricState.dm_pickup = [0,1,2].some(function(i){ return !!wvMetricState['dm_pickup_'+i]; });
};
setTimeout(function() { window.renderPickupMetricItems(); }, 600);

/* ─── TOUR OPERATORS & CONTRACTS PAGE ─── */
(function() {
  const TO_DATA = {
    sunshine:  { name:'Sunshine Tours',     contact:'John Smith',    status:'active',   contracts:2, revenue:'$125,000', bookings:45, rooms:'33%' },
    global:    { name:'Global Adventures',  contact:'Maria Garcia',  status:'active',   contracts:5, revenue:'$210,000', bookings:78, rooms:'41%' },
    beach:     { name:'Beach Holidays Ltd', contact:'Sarah Johnson', status:'active',   contracts:2, revenue:'$89,000',  bookings:32, rooms:'28%' },
    city:      { name:'City Break Tours',   contact:'Thomas Mueller',status:'active',   contracts:4, revenue:'$156,000', bookings:55, rooms:'37%' },
    adventure: { name:'Adventure Seekers',  contact:'Emma Wilson',   status:'inactive', contracts:1, revenue:'$22,000',  bookings:8,  rooms:'12%' },
  };

  // Sidebar navigation handled by the global nav handler below

  // Tab switching
  document.getElementById('tourOperators')?.addEventListener('click', e => {
    const tab = e.target.closest('.to-tab');
    if (tab) {
      document.querySelectorAll('.to-tab').forEach(t => t.classList.remove('active'));
      tab.classList.add('active');
      const name = tab.dataset.tab;
      document.querySelectorAll('.to-tab-panel').forEach(p => p.style.display = 'none');
      const panel = document.getElementById(`toPanel-${name}`);
      if (panel) panel.style.display = '';
    }

    // Row click → update detail panel
    const row = e.target.closest('.to-row');
    if (row) {
      document.querySelectorAll('.to-row').forEach(r => r.classList.remove('active-row'));
      row.classList.add('active-row');
      const op = TO_DATA[row.dataset.op];
      if (op) {
        document.getElementById('toDetailName').textContent    = op.name;
        document.getElementById('toDetailContact').textContent = op.contact;
        const statusEl = document.getElementById('toDetailStatus');
        statusEl.textContent = op.status.toUpperCase();
        statusEl.className = `to-status ${op.status}`;
        document.getElementById('toStatContracts').textContent = op.contracts;
        document.getElementById('toStatRevenue').textContent   = op.revenue;
        document.getElementById('toStatBookings').textContent  = op.bookings;
        document.getElementById('toStatRooms').textContent     = op.rooms;
        document.getElementById('toContractsSub').textContent  = `All contracts for ${op.name}`;
      }
    }

  });

  // Search filter
  document.getElementById('toSearch')?.addEventListener('input', e => {
    const q = e.target.value.toLowerCase();
    document.querySelectorAll('.to-row').forEach(row => {
      row.style.display = row.textContent.toLowerCase().includes(q) ? '' : 'none';
    });
  });
})();

/* ─── ANALYSIS & NOTES PAGE NAV ─── */
(function() {
  const pages = {
    0: document.querySelector('.page-content:not(#toPage):not(#analysisPage):not(#notesPage):not(#settingsPage)'),
    1: document.getElementById('toPage'),
    2: document.getElementById('analysisPage'),
    3: document.getElementById('notesPage'),
    4: document.getElementById('settingsPage'),
  };

  // Override sidebar nav to handle all 4 pages
  document.querySelectorAll('.nav-item').forEach(function(item, i) {
    item.addEventListener('click', function(e) {
      e.preventDefault();
      document.querySelectorAll('.nav-item').forEach(function(n) { n.classList.remove('active'); });
      item.classList.add('active');
      Object.values(pages).forEach(function(p) { if (p) p.style.display = 'none'; });
      if (pages[i]) pages[i].style.display = '';
      else if (pages[0]) pages[0].style.display = ''; // fallback to dashboard
    });
  });

  // Notes tab switching
  document.querySelectorAll('.cn-comms-tab').forEach(function(tab) {
    tab.addEventListener('click', function() {
      document.querySelectorAll('.cn-comms-tab').forEach(function(t) { t.classList.remove('active'); });
      tab.classList.add('active');
    });
  });

  // Add Note button stub
  document.getElementById('btnAddNote')?.addEventListener('click', function() {
    alert('Add Note form coming soon.');
  });
})();

/* ─── SETTINGS PAGE ─── */
(function() {
  const settingsPage = document.getElementById('settingsPage');
  if (!settingsPage) return;

  // Tab switching
  settingsPage.addEventListener('click', e => {
    const tab = e.target.closest('.st-tab');
    if (tab) {
      settingsPage.querySelectorAll('.st-tab').forEach(t => t.classList.remove('active'));
      tab.classList.add('active');
      settingsPage.querySelectorAll('.st-panel').forEach(p => p.style.display = 'none');
      const panel = settingsPage.querySelector(`#stPanel-${tab.dataset.stab}`);
      if (panel) panel.style.display = '';
      // Initialise board-types grid when operators tab is revealed
      if (tab.dataset.stab === 'operators') setTimeout(_btInit, 50);
    }

    // Role simulator
    const roleBtn = e.target.closest('.st-role-btn');
    if (roleBtn) {
      settingsPage.querySelectorAll('.st-role-btn').forEach(b => b.classList.remove('active'));
      roleBtn.classList.add('active');
      const badge = settingsPage.querySelector('.st-role-badge');
      if (badge) badge.lastChild.textContent = roleBtn.dataset.role.toUpperCase();
    }
  });

  // ── Calendar months selectors (update label when select changes) ─────────
  ['calMonthsSelect', 'wvMonthsSelect'].forEach(function(id) {
    var sel = document.getElementById(id);
    var labelId = id === 'calMonthsSelect' ? 'calMonthsLabel' : 'wvMonthsLabel';
    var lbl = document.getElementById(labelId);
    if (sel && lbl) {
      sel.addEventListener('change', function() {
        lbl.textContent = sel.options[sel.selectedIndex].text;
      });
    }
  });

  // ── Role select dropdown (syncs with hidden #stRoleGroup buttons) ───────
  var roleSelect = document.getElementById('stRoleSelect');
  var roleSelectWrap = roleSelect ? roleSelect.closest('.st-role-select-wrap') : null;
  if (roleSelect) {
    // Set initial data-label for CSS ::after display
    if (roleSelectWrap) roleSelectWrap.setAttribute('data-label', roleSelect.options[roleSelect.selectedIndex].text);
    roleSelect.addEventListener('change', function() {
      var selectedRole = this.value;
      var selectedText = this.options[this.selectedIndex].text;
      // Update CSS ::after label
      if (roleSelectWrap) roleSelectWrap.setAttribute('data-label', selectedText);
      // Sync with hidden role buttons so existing role logic fires
      var roleGroup = document.getElementById('stRoleGroup');
      if (roleGroup) {
        var hiddenBtn = roleGroup.querySelector('.st-role-btn[data-role="'+selectedRole+'"]');
        if (hiddenBtn) {
          roleGroup.querySelectorAll('.st-role-btn').forEach(b => b.classList.remove('active'));
          hiddenBtn.classList.add('active');
          // Update the role badge in the page header
          var badge = settingsPage.querySelector('.st-role-badge');
          if (badge) {
            var badgeText = badge.querySelector('span') || badge.lastChild;
            if (badgeText) badgeText.textContent = selectedText;
          }
        }
      }
    });
    // Trigger once on load to set initial label
    roleSelect.dispatchEvent(new Event('change'));
  }

  // Edit Targets toggle
  let editing = false;
  document.getElementById('stEditTargets')?.addEventListener('click', () => {
    editing = !editing;
    document.querySelectorAll('.st-target-input').forEach(inp => {
      inp.readOnly = !editing;
      inp.classList.toggle('editable', editing);
    });
    const saveRow = document.getElementById('stSaveRow');
    const editBtn = document.getElementById('stEditTargets');
    if (saveRow) saveRow.style.display = editing ? 'flex' : 'none';
    if (editBtn) editBtn.textContent = editing ? 'Editing…' : 'Edit targets';
  });

  document.getElementById('stCancelEdit')?.addEventListener('click', () => {
    editing = false;
    document.querySelectorAll('.st-target-input').forEach(inp => {
      inp.readOnly = true;
      inp.classList.remove('editable');
    });
    const saveRow = document.getElementById('stSaveRow');
    const editBtn = document.getElementById('stEditTargets');
    if (saveRow) saveRow.style.display = 'none';
    if (editBtn) editBtn.textContent = 'Edit targets';
  });

  document.getElementById('stSaveTargets')?.addEventListener('click', () => {
    editing = false;
    document.querySelectorAll('.st-target-input').forEach(inp => {
      inp.readOnly = true;
      inp.classList.remove('editable');
    });
    const saveRow = document.getElementById('stSaveRow');
    const editBtn = document.getElementById('stEditTargets');
    if (saveRow) saveRow.style.display = 'none';
    if (editBtn) editBtn.textContent = 'Edit targets';
    // Show brief save confirmation
    if (editBtn) { editBtn.textContent = 'Saved ✓'; setTimeout(() => { if (editBtn) editBtn.textContent = 'Edit targets'; }, 2000); }
  });
})();

/* ─── MEGA MENU ─── */
(function() {
  const trigger = document.getElementById('tourOpTab');
  const menu = document.getElementById('tourOpMegaMenu');
  if (!trigger || !menu) return;

  function positionMenu() {
    var rect = trigger.getBoundingClientRect();
    menu.style.left = rect.left + 'px';
  }

  trigger.addEventListener('click', function(e) {
    e.preventDefault();
    positionMenu();
    menu.classList.toggle('open');
  });

  // Close on outside click
  document.addEventListener('click', function(e) {
    if (!trigger.contains(e.target) && !menu.contains(e.target)) {
      menu.classList.remove('open');
    }
  });

  // Mega menu links navigate to the appropriate page
  menu.addEventListener('click', function(e) {
    const link = e.target.closest('.mega-link');
    if (!link) return;
    e.preventDefault();
    const pageIdx = parseInt(link.dataset.navpage, 10);
    const navItems = document.querySelectorAll('.nav-item');
    const pages = {
      0: document.querySelector('.page-content:not(#toPage):not(#analysisPage):not(#notesPage):not(#settingsPage)'),
      1: document.getElementById('toPage'),
      2: document.getElementById('analysisPage'),
      3: document.getElementById('notesPage'),
      4: document.getElementById('settingsPage'),
    };
    navItems.forEach(function(n) { n.classList.remove('active'); });
    if (navItems[pageIdx]) navItems[pageIdx].classList.add('active');
    Object.values(pages).forEach(function(p) { if (p) p.style.display = 'none'; });
    if (pages[pageIdx]) pages[pageIdx].style.display = '';
    menu.classList.remove('open');
  });
})();

/* ─── ANALYSIS AG-GRID ─── */
(function() {
  const el = document.getElementById('analysisGrid');
  if (!el || typeof agGrid === 'undefined') return;

  // ── Generate Q1 2026 demo data ─────────────────────────────────────────
  var TOUR_OPS = [
    {name:'TUI Group',           baseSold:24, baseAdr:76},
    {name:'Thomas Cook',         baseSold:28, baseAdr:76},
    {name:'Jet2holidays',        baseSold:32, baseAdr:76},
    {name:'FTI Group',           baseSold:33, baseAdr:80},
    {name:'DER Touristik',       baseSold:28, baseAdr:78},
    {name:'Alltours',            baseSold:25, baseAdr:76},
    {name:'Neckermann',          baseSold:28, baseAdr:76},
    {name:'Schauinsland Reisen', baseSold:29, baseAdr:77},
  ];
  var MN = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  var DN = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
  var RAW_DATA = [];
  for (var m = 0; m < 3; m++) {
    var dIM = new Date(2026, m+1, 0).getDate();
    for (var d = 1; d <= dIM; d++) {
      var dt = new Date(2026, m, d);
      TOUR_OPS.forEach(function(op) {
        var sold = op.baseSold + Math.round(Math.sin(d * 0.31 + m * 1.1) * 6);
        var adr  = op.baseAdr  + Math.round(Math.cos(d * 0.47 + m * 0.9) * 10);
        if (sold < 10) sold = 10;
        if (adr  < 60) adr  = 60;
        var rev   = sold * adr;
        var lyRev = rev * (0.88 + Math.abs(Math.sin(d * 0.23 + m)) * 0.18);
        var iso = '2026-'+(m+1<10?'0':'')+(m+1)+'-'+(d<10?'0':'')+d;
        RAW_DATA.push({ date: iso, day: DN[dt.getDay()],
          roomType:'Standard Room', board:'Room Only', segment:'Travel Distribution Hub', tourOp: op.name,
          _sold:sold, _avail:150, _adr:adr, _rev:rev, _restRev:0, _lyRev:Math.round(lyRev),
          _m:m, _d:d, _y:2026 });
      });
      // OTA + Corporate rows
      var iso2 = '2026-'+(m+1<10?'0':'')+(m+1)+'-'+(d<10?'0':'')+d;
      [{seg:'OTA',op:'–',sold:28,avail:14,adr:85.5,rr:0},
       {seg:'Corporate',op:'–',sold:18,avail:150,adr:83.6,rr:10.58}].forEach(function(r) {
        var rev = r.sold * r.adr;
        RAW_DATA.push({ date: iso2, day: DN[dt.getDay()],
          roomType:'Standard Room', board:'Room Only', segment:r.seg, tourOp:r.op,
          _sold:r.sold, _avail:r.avail, _adr:r.adr, _rev:rev, _restRev:r.rr, _lyRev:rev*0.93,
          _m:m, _d:d, _y:2026 });
      });
    }
  }

  // ── Helpers ────────────────────────────────────────────────────────────
  function fmtE(v) { return '€'+v.toFixed(2); }
  function fmtP(n,d) { return d>0 ? Math.round(n/d*100)+'%' : '–'; }
  function deltaStyle(p) {
    var v = parseFloat(p.value);
    return isNaN(v) ? {} : (v>=0 ? {color:'#16a34a',fontWeight:'500'} : {color:'#dc2626',fontWeight:'500'});
  }
  function toRow(r) {
    var delta = r._lyRev>0 ? ((r._rev-r._lyRev)/r._lyRev*100).toFixed(1)+'%' : '–';
    return { date:r.date, day:r.day, roomType:r.roomType, board:r.board, segment:r.segment, tourOp:r.tourOp,
      roomsSold:r._sold, available:r._avail, occ:fmtP(r._sold,r._avail),
      adr:fmtE(r._adr), revenue:fmtE(r._rev), restRev:fmtE(r._restRev),
      revpar:fmtE(r._rev/r._avail), lyRevenue:fmtE(r._lyRev), lyDelta:delta };
  }

  // ── Aggregation ────────────────────────────────────────────────────────
  function periodLabel(r, period) {
    if (period==='month')   return MN[r._m]+' '+r._y;
    if (period==='quarter') return 'Q'+(Math.floor(r._m/3)+1)+' '+r._y;
    if (period==='year')    return String(r._y);
    return r.date;
  }
  function aggregate(rows, period) {
    if (period==='custom') return rows.map(toRow);
    var groups = {};
    rows.forEach(function(r) {
      var key = periodLabel(r,period)+'|'+r.roomType+'|'+r.board+'|'+r.segment+'|'+r.tourOp;
      if (!groups[key]) groups[key] = {pKey:periodLabel(r,period), roomType:r.roomType, board:r.board,
        segment:r.segment, tourOp:r.tourOp, sold:0, avail:r._avail, rev:0, restRev:0, lyRev:0, days:0};
      var g = groups[key];
      g.sold+=r._sold; g.rev+=r._rev; g.restRev+=r._restRev; g.lyRev+=r._lyRev; g.days++;
    });
    return Object.values(groups).map(function(g) {
      var cap = g.avail * g.days;
      var adr = g.sold>0 ? g.rev/g.sold : 0;
      var delta = g.lyRev>0 ? ((g.rev-g.lyRev)/g.lyRev*100).toFixed(1)+'%' : '–';
      return { date:g.pKey, day:g.days+' days', roomType:g.roomType, board:g.board, segment:g.segment, tourOp:g.tourOp,
        roomsSold:g.sold, available:cap, occ:fmtP(g.sold,cap),
        adr:fmtE(adr), revenue:fmtE(g.rev), restRev:fmtE(g.restRev),
        revpar:fmtE(g.rev/cap), lyRevenue:fmtE(g.lyRev), lyDelta:delta };
    });
  }

  // ── State ──────────────────────────────────────────────────────────────
  var currentPeriod = 'custom';
  var filterRoom='', filterBoard='', filterSeg='', filterTourOp='';
  var drpDateStart = '', drpDateEnd = ''; // ISO strings e.g. '2026-01-15'

  function getDisplayData() {
    var rows = RAW_DATA.filter(function(r) {
      if (filterRoom   && r.roomType!==filterRoom)   return false;
      if (filterBoard  && r.board!==filterBoard)     return false;
      if (filterSeg    && r.segment!==filterSeg)     return false;
      if (filterTourOp && r.tourOp!==filterTourOp)   return false;
      if (drpDateStart && r.date < drpDateStart)     return false;
      if (drpDateEnd   && r.date > drpDateEnd)       return false;
      return true;
    });
    return aggregate(rows, currentPeriod);
  }

  // Exposed so the drp picker can trigger a filtered refresh
  window.anApplyDateRange = function(isoStart, isoEnd) {
    drpDateStart = isoStart || '';
    drpDateEnd   = isoEnd   || '';
    gridApi.setGridOption('rowData', getDisplayData());
  };

  // ── Column defs ────────────────────────────────────────────────────────
  var colDefs = [
    { field:'date',      headerName:'Date',              filter:'agTextColumnFilter',   width:120 },
    { field:'roomType',  headerName:'Room Type',         filter:'agTextColumnFilter',   flex:1.4 },
    { field:'board',     headerName:'Meal Plan',        filter:'agTextColumnFilter',   flex:1.2 },
    { field:'segment',   headerName:'Segment',           filter:'agTextColumnFilter',   flex:1.2 },
    { field:'roomsSold', headerName:'Rooms Sold',        filter:'agNumberColumnFilter', width:115, type:'numericColumn' },
    { field:'available', headerName:'Rooms Available',   filter:'agNumberColumnFilter', width:140, type:'numericColumn' },
    { field:'occ',       headerName:'Occupancy %',       filter:'agTextColumnFilter',   width:120, cellStyle:{textAlign:'right'} },
    { field:'adr',       headerName:'ADR',               filter:'agTextColumnFilter',   width:100, cellStyle:{textAlign:'right'} },
    { field:'revenue',   headerName:'Revenue',           filter:'agTextColumnFilter',   flex:1,    cellStyle:{textAlign:'right'} },
    { field:'restRev',   headerName:'Restaurant Rev',    filter:'agTextColumnFilter',   flex:1.1,  cellStyle:{textAlign:'right'} },
    { field:'lyRevenue', headerName:'LY Revenue',        filter:'agTextColumnFilter',   flex:1,    cellStyle:{textAlign:'right'}, hide:true },
    { field:'lyDelta',   headerName:'vs LY %',           filter:'agTextColumnFilter',   width:90,  cellStyle:deltaStyle, hide:true },
  ];

  // ── New theming API (ag-Grid v33+) ─────────────────────────────────────
  var anTheme = sharedTheme;

  // ── Grid init ──────────────────────────────────────────────────────────
  var gridApi = agGrid.createGrid(el, {
    theme: anTheme,
    columnDefs: colDefs,
    rowData: getDisplayData(),
    rowHeight: 42,
    headerHeight: 42,
    suppressMovableColumns: true,
    suppressCellFocus: true,
    tooltipShowDelay: 300,
    defaultColDef: { sortable: true, resizable: true, floatingFilter: true, minWidth: 200 },
  });

  // ── Period tabs ────────────────────────────────────────────────────────
  var periodTabs = document.getElementById('anPeriodTabs');
  if (periodTabs) {
    periodTabs.addEventListener('click', function(e) {
      var btn = e.target.closest('.an-period-btn');
      if (!btn) return;
      currentPeriod = btn.dataset.period;
      periodTabs.querySelectorAll('.an-period-btn').forEach(function(b) {
        b.classList.toggle('active', b===btn);
      });
      var drpWrap = document.getElementById('drpWrap');
      if (drpWrap) drpWrap.style.display = currentPeriod==='custom' ? '' : 'none';
      // Clear date filter when leaving custom range
      if (currentPeriod !== 'custom') { drpDateStart = ''; drpDateEnd = ''; }
      gridApi.setGridOption('rowData', getDisplayData());
    });
  }

  // ── Period dropdown (header UI — syncs with hidden #anPeriodTabs) ────────
  var periodTrigger = document.getElementById('anPeriodTrigger');
  var periodDropdown = document.getElementById('anPeriodDropdown');
  var periodLabel = document.getElementById('anPeriodLabel');
  if (periodTrigger && periodDropdown) {
    // Open/close dropdown
    periodTrigger.addEventListener('click', function(e) {
      e.stopPropagation();
      var isOpen = periodDropdown.style.display !== 'none';
      periodDropdown.style.display = isOpen ? 'none' : '';
      periodTrigger.classList.toggle('active', !isOpen);
    });
    // Close on outside click
    document.addEventListener('click', function(e) {
      if (!periodTrigger.contains(e.target) && !periodDropdown.contains(e.target)) {
        periodDropdown.style.display = 'none';
        periodTrigger.classList.remove('active');
      }
    });
    // Option click — update label, sync with hidden tabs
    periodDropdown.addEventListener('click', function(e) {
      var opt = e.target.closest('.an-period-opt');
      if (!opt) return;
      var period = opt.dataset.period;
      // Update dropdown active state
      periodDropdown.querySelectorAll('.an-period-opt').forEach(function(o) {
        o.classList.toggle('active', o === opt);
      });
      if (periodLabel) periodLabel.textContent = opt.textContent;
      periodDropdown.style.display = 'none';
      periodTrigger.classList.remove('active');
      // Trigger the hidden period tabs listener to handle data update
      var hiddenBtn = document.querySelector('#anPeriodTabs .an-period-btn[data-period="'+period+'"]');
      if (hiddenBtn) hiddenBtn.click();
    });
  }

  // ── Compare to LY toggle ───────────────────────────────────────────────
  var lyToggle = document.getElementById('anCompareLY');
  if (lyToggle) {
    lyToggle.addEventListener('change', function() {
      gridApi.setColumnsVisible(['lyRevenue','lyDelta'], this.checked);
    });
  }

  // ── Filter bar ─────────────────────────────────────────────────────────
  function onFilterChange() {
    filterRoom   = document.getElementById('anFbRoomType')?.value || '';
    filterBoard  = document.getElementById('anFbBoard')?.value   || '';
    filterSeg    = document.getElementById('anFbSegment')?.value  || '';
    filterTourOp = document.getElementById('anFbTourOp')?.value   || '';
    gridApi.setGridOption('rowData', getDisplayData());
  }
  ['anFbRoomType','anFbBoard','anFbSegment','anFbTourOp'].forEach(function(id) {
    document.getElementById(id)?.addEventListener('change', onFilterChange);
  });

  // ── Search input ───────────────────────────────────────────────────────
  var searchInput = document.getElementById('anSearchInput');
  if (searchInput) {
    searchInput.addEventListener('input', function() {
      gridApi.setGridOption('quickFilterText', this.value);
    });
  }
})();

/* ─── TOUR OPERATORS GRID ─── */
var _tourOpsGridApi = null;
(function() {
  var TO_OPS_DATA = [
    { name:'Sunshine Tours',    contact:'John Smith',     email:'john@sunshinetours.com',        country:'USA',       contracts:3, status:'ACTIVE',
      emails:{ stopSale:{contact:'John Smith',email:'stopsale@sunshinetours.com'}, release:{contact:'John Smith',email:'release@sunshinetours.com'}, promotions:{contact:'Lisa Brown',email:'promos@sunshinetours.com'}, contract:{contact:'John Smith',email:'contracts@sunshinetours.com'} } },
    { name:'Global Adventures', contact:'Maria Garcia',   email:'maria@globaladv.com',           country:'Spain',     contracts:5, status:'ACTIVE',
      emails:{ stopSale:{contact:'Maria Garcia',email:'stopsale@globaladv.com'}, release:{contact:'Carlos Ruiz',email:'release@globaladv.com'}, promotions:{contact:'Maria Garcia',email:'promos@globaladv.com'}, contract:{contact:'Carlos Ruiz',email:'contracts@globaladv.com'} } },
    { name:'Beach Holidays Ltd',contact:'Sarah Johnson',  email:'sarah@beachholidays.co.uk',     country:'UK',        contracts:2, status:'ACTIVE',
      emails:{ stopSale:{contact:'Sarah Johnson',email:'stopsale@beachholidays.co.uk'}, release:{contact:'Sarah Johnson',email:'release@beachholidays.co.uk'}, promotions:{contact:'Tom Clarke',email:'promos@beachholidays.co.uk'}, contract:{contact:'Tom Clarke',email:'contracts@beachholidays.co.uk'} } },
    { name:'City Break Tours',  contact:'Thomas Mueller', email:'thomas@citybreak.de',           country:'Germany',   contracts:4, status:'ACTIVE',
      emails:{ stopSale:{contact:'Thomas Mueller',email:'stopsale@citybreak.de'}, release:{contact:'Anna Schmidt',email:'release@citybreak.de'}, promotions:{contact:'Anna Schmidt',email:'promos@citybreak.de'}, contract:{contact:'Thomas Mueller',email:'contracts@citybreak.de'} } },
    { name:'Adventure Seekers', contact:'Emma Wilson',    email:'emma@adventureseekers.com.au',  country:'Australia', contracts:1, status:'INACTIVE',
      emails:{ stopSale:{contact:'Emma Wilson',email:'stopsale@adventureseekers.com.au'}, release:{contact:'',email:''}, promotions:{contact:'',email:''}, contract:{contact:'Emma Wilson',email:'contracts@adventureseekers.com.au'} } },
  ];

  var TO_DATA_MAP = {
    'Sunshine Tours':    { contracts:2, revenue:'$125,000', bookings:45, rooms:'33%' },
    'Global Adventures': { contracts:5, revenue:'$210,000', bookings:78, rooms:'41%' },
    'Beach Holidays Ltd':{ contracts:2, revenue:'$89,000',  bookings:32, rooms:'28%' },
    'City Break Tours':  { contracts:4, revenue:'$156,000', bookings:55, rooms:'37%' },
    'Adventure Seekers': { contracts:1, revenue:'$22,000',  bookings:8,  rooms:'12%' },
  };

  function updateDetailPanel(data) {
    var op   = TO_DATA_MAP[data.name] || { contracts:0, revenue:'—', bookings:0, rooms:'—' };
    var eml  = data.emails || {};
    var set  = function(id, v) { var el=document.getElementById(id); if(el) el.textContent = v||'—'; };
    var cls  = data.status==='ACTIVE' ? 'active' : 'inactive';
    var stEl = document.getElementById('toDetailStatus');
    if (stEl) { stEl.textContent = data.status; stEl.className = 'to-status ' + cls; }
    set('toDetailName',    data.name);
    set('toDetailContact', data.contact || '—');
    set('toStatContracts', op.contracts);
    set('toStatRevenue',   op.revenue);
    set('toStatBookings',  op.bookings);
    set('toStatRooms',     op.rooms);
    set('toContractsSub',  'All contracts for ' + data.name);
    set('toEmailStopSaleContact', eml.stopSale    && eml.stopSale.contact);
    set('toEmailStopSale',        eml.stopSale    && eml.stopSale.email);
    set('toEmailReleaseContact',  eml.release     && eml.release.contact);
    set('toEmailRelease',         eml.release     && eml.release.email);
    set('toEmailPromoContact',    eml.promotions  && eml.promotions.contact);
    set('toEmailPromo',           eml.promotions  && eml.promotions.email);
    set('toEmailContractContact', eml.contract    && eml.contract.contact);
    set('toEmailContract',        eml.contract    && eml.contract.email);
  }

  var el = document.getElementById('tourOpsGrid');
  if (!el) return;

  var AG = _realAgGrid;
  if (!AG || typeof AG.createGrid !== 'function') return;

  el.innerHTML = '';
  el.style.padding = '0';
  var wrapper = document.createElement('div');
  wrapper.className = 'ag-theme-quartz tour-ops-ag-wrap';
  el.appendChild(wrapper);

  var colDefs = [
    { field: 'name', headerName: 'Operator', flex: 2, minWidth: 160, pinned: 'left', lockPinned: true,
      cellStyle: { fontWeight: '700', fontSize: '13px', display: 'flex', alignItems: 'center' } },
    { field: 'contact', headerName: 'Contact name', flex: 1.5, minWidth: 140,
      cellStyle: { fontSize: '12px', display: 'flex', alignItems: 'center' } },
    { field: 'email', headerName: 'Email', flex: 2, minWidth: 180,
      cellStyle: { fontSize: '12px', display: 'flex', alignItems: 'center', color: '#374151' } },
    { field: 'country', headerName: 'Country', width: 110,
      cellStyle: { fontSize: '12px', display: 'flex', alignItems: 'center' } },
    { field: 'contracts', headerName: 'Contracts', width: 100, type: 'numericColumn',
      cellStyle: { fontSize: '12px', display: 'flex', alignItems: 'center', justifyContent: 'center' } },
    { field: 'status', headerName: 'Status', width: 110,
      cellRenderer: function(p) {
        var cls = p.value === 'ACTIVE' ? 'active' : 'inactive';
        return '<span class="to-status ' + cls + '">' + p.value + '</span>';
      },
      cellStyle: { display: 'flex', alignItems: 'center', justifyContent: 'center' } },
  ];

  _tourOpsGridApi = AG.createGrid(wrapper, {
    columnDefs: colDefs,
    rowData: TO_OPS_DATA,
    rowHeight: 48,
    headerHeight: 38,
    domLayout: 'autoHeight',
    suppressMovableColumns: true,
    suppressCellFocus: true,
    rowSelection: 'single',
    defaultColDef: { sortable: true, resizable: true },
    getRowStyle: function(p) {
      return p.node.isSelected() ? { background: 'rgba(0,100,97,0.08)' } : {};
    },
    onRowClicked: function(params) {
      params.node.setSelected(true);
      updateDetailPanel(params.data);
    },
    onFirstDataRendered: function(params) {
      params.api.getRowNode(0)?.setSelected(true);
    },
  });

  // Show first operator by default
  updateDetailPanel(TO_OPS_DATA[0]);

  /* ── New Operator Modal ── */
  function openModal() {
    // reset form
    ['tonName','tonContact','tonPhone','tonCountry',
     'tonEmailStopSaleContact','tonEmailStopSale',
     'tonEmailReleaseContact','tonEmailRelease',
     'tonEmailPromoContact','tonEmailPromo',
     'tonEmailContractContact','tonEmailContract'].forEach(function(id){
      var el = document.getElementById(id); if(el) el.value = '';
    });
    document.querySelector('input[name="tonStatus"][value="ACTIVE"]').checked = true;
    var errEl = document.getElementById('tonNameErr'); if(errEl) errEl.style.display = 'none';
    document.getElementById('toNewOpOverlay').classList.add('open');
    setTimeout(function(){ var el=document.getElementById('tonName'); if(el) el.focus(); }, 80);
  }
  function closeModal() {
    document.getElementById('toNewOpOverlay').classList.remove('open');
  }
  function saveOperator() {
    var name = (document.getElementById('tonName').value || '').trim();
    if (!name) {
      var errEl = document.getElementById('tonNameErr'); if(errEl) errEl.style.display = '';
      document.getElementById('tonName').focus();
      return;
    }
    var status   = document.querySelector('input[name="tonStatus"]:checked').value;
    var contact  = (document.getElementById('tonContact').value || '').trim();
    var phone    = (document.getElementById('tonPhone').value || '').trim();
    var country  = (document.getElementById('tonCountry').value || '').trim();
    var stopSaleContact  = (document.getElementById('tonEmailStopSaleContact').value || '').trim();
    var stopSale         = (document.getElementById('tonEmailStopSale').value || '').trim();
    var releaseContact   = (document.getElementById('tonEmailReleaseContact').value || '').trim();
    var release          = (document.getElementById('tonEmailRelease').value || '').trim();
    var promotionsContact= (document.getElementById('tonEmailPromoContact').value || '').trim();
    var promotions       = (document.getElementById('tonEmailPromo').value || '').trim();
    var contractContact  = (document.getElementById('tonEmailContractContact').value || '').trim();
    var contract         = (document.getElementById('tonEmailContract').value || '').trim();

    var newOp = {
      name: name, contact: contact,
      email: stopSale || release || promotions || contract || '—',
      country: country, contracts: 0, status: status,
      emails: {
        stopSale:   { contact: stopSaleContact,   email: stopSale },
        release:    { contact: releaseContact,    email: release },
        promotions: { contact: promotionsContact, email: promotions },
        contract:   { contact: contractContact,   email: contract },
      },
    };

    // Add to data and refresh grid
    TO_OPS_DATA.push(newOp);
    TO_DATA_MAP[name] = { contracts: 0, revenue: '—', bookings: 0, rooms: '—' };
    if (_tourOpsGridApi) {
      _tourOpsGridApi.setGridOption('rowData', TO_OPS_DATA);
      // select the new row
      setTimeout(function() {
        var lastNode = _tourOpsGridApi.getRowNode(TO_OPS_DATA.length - 1);
        if (lastNode) lastNode.setSelected(true);
      }, 0);
    }

    // Update count badge
    var badge = document.querySelector('#toPanel-operators .to-count-badge');
    if (badge) badge.textContent = TO_OPS_DATA.length;

    // Select the new operator
    updateDetailPanel(newOp);
    closeModal();
  }

  document.getElementById('btnNewOperator')?.addEventListener('click', openModal);
  document.getElementById('toNewOpClose')?.addEventListener('click', closeModal);
  document.getElementById('toNewOpCancel')?.addEventListener('click', closeModal);
  document.getElementById('toNewOpSave')?.addEventListener('click', saveOperator);
  document.getElementById('toNewOpOverlay')?.addEventListener('click', function(e) {
    if (e.target === this) closeModal();
  });
  document.getElementById('toNewOpModal')?.addEventListener('keydown', function(e) {
    if (e.key === 'Escape') closeModal();
    if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) saveOperator();
  });
  document.getElementById('tonName')?.addEventListener('input', function() {
    var errEl = document.getElementById('tonNameErr'); if(errEl) errEl.style.display = 'none';
  });
})();

/* ─── ALL CONTRACTS GRID ─── */
(function() {
  var CONTRACTS_DATA = [
    {id:'CT-005', name:'test',                  operator:'Sunshine Tours',    bookingWindow:'—', stayWindow:'—', totalValue:'$0', bookings:0, revenue:'$0', roomsSold:'0/—', status:'ACCEPTED'},
    {id:'CT-001', name:'Q1 2026 Package',       operator:'Sunshine Tours',    bookingWindow:'—', stayWindow:'—', totalValue:'$0', bookings:0, revenue:'$0', roomsSold:'0/—', status:'ACCEPTED'},
    {id:'CT-010', name:'European Tour Package', operator:'Global Adventures', bookingWindow:'—', stayWindow:'—', totalValue:'$0', bookings:0, revenue:'$0', roomsSold:'0/—', status:'ACCEPTED'},
    {id:'CT-015', name:'Beach Paradise 2026',   operator:'Beach Holidays Ltd',bookingWindow:'—', stayWindow:'—', totalValue:'$0', bookings:0, revenue:'$0', roomsSold:'0/—', status:'ACCEPTED'},
    {id:'CT-020', name:'City Explorer Package', operator:'City Break Tours',  bookingWindow:'—', stayWindow:'—', totalValue:'$0', bookings:0, revenue:'$0', roomsSold:'0/—', status:'ACCEPTED'},
  ];

  var viewSvg = '<svg viewBox="0 0 14 14" fill="none" width="13" height="13"><circle cx="7" cy="7" r="5" stroke="currentColor" stroke-width="1.3"/><circle cx="7" cy="7" r="2" stroke="currentColor" stroke-width="1.3"/></svg>';
  var tagSvg  = '<svg viewBox="0 0 14 14" fill="none" width="13" height="13"><path d="M2 2h5l5 5-5 5-5-5V2z" stroke="currentColor" stroke-width="1.3" stroke-linejoin="round"/></svg>';

  var el = document.getElementById('contractsGrid');
  if (!el || typeof agGrid === 'undefined') return;

  var colDefs = [
    { field: 'id',           headerName: 'Contract ID',      width: 100,
      cellRenderer: function(p) { return '<span class="ac-id">' + p.value + '</span>'; } },
    { field: 'name',         headerName: 'Contract name',    flex: 1.5,
      cellRenderer: function(p) { return '<strong>' + p.value + '</strong>'; } },
    { field: 'operator',     headerName: 'Operator',         flex: 1.2 },
    { field: 'bookingWindow',headerName: 'Booking window',   flex: 1 },
    { field: 'stayWindow',   headerName: 'Stay date window', flex: 1 },
    { field: 'totalValue',   headerName: 'Total value',      width: 100, cellStyle: { textAlign: 'right' } },
    { field: 'bookings',     headerName: 'Bookings',         width: 90,  type: 'numericColumn' },
    { field: 'revenue',      headerName: 'Revenue',          width: 100, cellStyle: { textAlign: 'right' } },
    { field: 'roomsSold',    headerName: 'Rooms sold',       width: 100 },
    { field: 'status',       headerName: 'Status',           width: 110,
      cellRenderer: function(p) { return '<span class="to-badge accepted">' + p.value + '</span>'; } },
    { headerName: 'Actions', width: 90, sortable: false, resizable: false,
      cellRenderer: function() {
        return '<div class="ac-actions-cell"><button class="to-icon-btn" title="View">' + viewSvg + '</button>' +
               '<button class="to-icon-btn" title="Tag">' + tagSvg + '</button></div>';
      }
    },
  ];

  var contractsGridApi = agGrid.createGrid(el, {
    theme: sharedTheme,
    columnDefs: colDefs,
    rowData: CONTRACTS_DATA,
    rowHeight: 42,
    headerHeight: 42,
    suppressMovableColumns: true,
    suppressCellFocus: true,
    defaultColDef: { sortable: true, resizable: true, floatingFilter: false, minWidth: 200 },
  });
  window._contractsGridApi = contractsGridApi;
  window._contractsData = CONTRACTS_DATA;
})();

/* ─── CONTRACTED ROOMS INVENTORY GRID ─── */
(function() {
  var INVENTORY_DATA = [
    {roomType:'Standard King',    operator:'Sunshine Tours',    contractName:'Q1 2026 Package',      contractId:'CT-001', allocated:100, sold:75,  remaining:25, soldPct:75.0,  revenue:'$11.3k', adr:'$150', period:'Jan 1 – Mar 31, 2026',    pctClass:'high'},
    {roomType:'Deluxe Room',      operator:'Sunshine Tours',    contractName:'Q1 2026 Package',      contractId:'CT-001', allocated:80,  sold:62,  remaining:18, soldPct:77.5,  revenue:'$12.4k', adr:'$200', period:'Jan 1 – Mar 31, 2026',    pctClass:'high'},
    {roomType:'Standard King',    operator:'Global Adventures', contractName:'European Tour Package', contractId:'CT-010', allocated:120, sold:98,  remaining:22, soldPct:81.7,  revenue:'$17.6k', adr:'$180', period:'Jan 1 – Dec 31, 2026',    pctClass:'high'},
    {roomType:'Premium Suite',    operator:'Global Adventures', contractName:'European Tour Package', contractId:'CT-010', allocated:50,  sold:45,  remaining:5,  soldPct:90.0,  revenue:'$11.3k', adr:'$250', period:'Jan 1 – Dec 31, 2026',    pctClass:'very-high'},
    {roomType:'Deluxe Room',      operator:'Paradise Getaways', contractName:'Winter Escape 2026',   contractId:'CT-015', allocated:60,  sold:42,  remaining:18, soldPct:70.0,  revenue:'$9.2k',  adr:'$220', period:'Jan 15 – Mar 15, 2026',   pctClass:'mid'},
    {roomType:'Family Room',      operator:'Paradise Getaways', contractName:'Winter Escape 2026',   contractId:'CT-015', allocated:45,  sold:38,  remaining:7,  soldPct:84.4,  revenue:'$7.6k',  adr:'$200', period:'Jan 15 – Mar 15, 2026',   pctClass:'high'},
    {roomType:'Executive Suite',  operator:'Luxury Journeys',   contractName:'Premium Experience',   contractId:'CT-023', allocated:30,  sold:28,  remaining:2,  soldPct:93.3,  revenue:'$8.4k',  adr:'$300', period:'Jan 1 – Jun 30, 2026',    pctClass:'very-high'},
    {roomType:'Standard King',    operator:'Budget Travels',    contractName:'Value Package 2026',   contractId:'CT-031', allocated:150, sold:89,  remaining:61, soldPct:59.3,  revenue:'$11.1k', adr:'$125', period:'Jan 1 – Apr 30, 2026',    pctClass:'low'},
    {roomType:'Standard Double',  operator:'Budget Travels',    contractName:'Value Package 2026',   contractId:'CT-031', allocated:120, sold:76,  remaining:44, soldPct:63.3,  revenue:'$9.5k',  adr:'$125', period:'Jan 1 – Apr 30, 2026',    pctClass:'mid'},
    {roomType:'Premium Suite',    operator:'Sunshine Tours',    contractName:'Summer Special 2026',  contractId:'CT-002', allocated:40,  sold:8,   remaining:32, soldPct:20.0,  revenue:'$2.0k',  adr:'$250', period:'Jan 1 – Aug 31, 2026',    pctClass:'very-low'},
    {roomType:'Deluxe Room',      operator:'Coastal Escapes',   contractName:'Beach Season Deal',    contractId:'CT-018', allocated:70,  sold:54,  remaining:16, soldPct:77.1,  revenue:'$11.9k', adr:'$220', period:'Feb 1 – May 31, 2026',    pctClass:'high'},
    {roomType:'Family Room',      operator:'Coastal Escapes',   contractName:'Beach Season Deal',    contractId:'CT-018', allocated:55,  sold:47,  remaining:8,  soldPct:85.5,  revenue:'$9.4k',  adr:'$200', period:'Feb 1 – May 31, 2026',    pctClass:'high'},
  ];

  var INVENTORY_TOTAL = {
    roomType:'TOTAL', operator:'', contractName:'', contractId:'', allocated:920, sold:662, remaining:258, soldPct:72.0,
    revenue:'$121.7k', adr:'$184', period:'—', pctClass:'high',
  };

  var el = document.getElementById('inventoryGrid');
  if (!el || typeof agGrid === 'undefined') return;

  var colDefs = [
    { field: 'roomType',     headerName: 'Room type',     flex: 1 },
    { field: 'operator',     headerName: 'Tour operator', flex: 1.2 },
    { field: 'contractName', headerName: 'Contract',      flex: 1.5,
      cellRenderer: function(p) {
        if (!p.data.contractId) return '';
        return '<span class="ac-contract-name">' + p.value + '</span><br><span class="ac-contract-id">' + p.data.contractId + '</span>';
      }
    },
    { field: 'allocated',    headerName: 'Allocated',     width: 100, type: 'numericColumn' },
    { field: 'sold',         headerName: 'Sold',          width: 80,  type: 'numericColumn' },
    { field: 'remaining',    headerName: 'Remaining',     width: 100, type: 'numericColumn' },
    { field: 'soldPct',      headerName: 'SOLD %',        width: 100,
      cellRenderer: function(p) {
        return '<span class="ac-pct ' + p.data.pctClass + '">' + p.value.toFixed(1) + '%</span>';
      }
    },
    { field: 'revenue',      headerName: 'Revenue',       flex: 1,    cellStyle: { textAlign: 'right' } },
    { field: 'adr',          headerName: 'ADR',           width: 80,  cellStyle: { textAlign: 'right' } },
    { field: 'period',       headerName: 'Period',        flex: 1.2 },
  ];

  agGrid.createGrid(el, {
    theme: sharedTheme,
    columnDefs: colDefs,
    rowData: INVENTORY_DATA,
    pinnedBottomRowData: [INVENTORY_TOTAL],
    rowHeight: 42,
    headerHeight: 42,
    suppressMovableColumns: true,
    suppressCellFocus: true,
    defaultColDef: { sortable: true, resizable: true, floatingFilter: false, minWidth: 200 },
  });
})();

/* ─── PROMOTIONS GRID ─── */
(function() {
  var viewSvg   = '<svg viewBox="0 0 14 14" fill="none" width="13" height="13"><circle cx="7" cy="7" r="5" stroke="currentColor" stroke-width="1.3"/><circle cx="7" cy="7" r="2" stroke="currentColor" stroke-width="1.3"/></svg>';
  var deleteSvg = '<svg viewBox="0 0 14 14" fill="none" width="13" height="13"><polyline points="2,4 12,4" stroke="currentColor" stroke-width="1.3" stroke-linecap="round"/><path d="M5 4V3h4v1M4 4l1 8h4l1-8" stroke="currentColor" stroke-width="1.3" stroke-linecap="round" stroke-linejoin="round"/></svg>';
  var sendSvg   = '<svg viewBox="0 0 14 14" fill="none" width="13" height="13"><path d="M2 2l10 5-10 5V8.5l7-1.5-7-1.5V2z" fill="currentColor"/></svg>';

  var PROMOTIONS_DATA = [
    {id:'PM-001', name:'Summer Kids Free Campaign',  tag:'Kids Free',           operators:'Global Adventures, Sunshine Tours +1 more', bookingWindow:'Jan 7 – Mar 13, 2026',         stayDates:'Mar 24 – May 17, 2026', expiry:'—',  rooms:null,  matched:null, revenue:'—',      status:'template'},
    {id:'PM-002', name:'Early Bird Spring Special',  tag:'Early Bird',          operators:'City Break Tours, Global Adventures',        bookingWindow:'Dec 1, 2025 – Mar 5, 2026',    stayDates:'Apr 26 – Jul 1, 2026',  expiry:'—',  rooms:null,  matched:null, revenue:'—',      status:'template'},
    {id:'PM-003', name:'Last Minute Winter Deals',   tag:'Last Minute',         operators:'Sunshine Tours, Beach Holidays Ltd +1 more', bookingWindow:'Jan 7 – Mar 21, 2026',         stayDates:'Mar 28 – Jun 22, 2026', expiry:'67d',rooms:23,    matched:23,  revenue:'$27,600', status:'active'},
    {id:'PM-004', name:'Extended Stay Autumn Package',tag:'Extended Stay',      operators:'Sunshine Tours, Adventure Seekers +1 more',  bookingWindow:'Dec 16, 2025 – Jan 21, 2026',  stayDates:'Feb 18 – Apr 18, 2026', expiry:'19d',rooms:23,    matched:23,  revenue:'$27,600', status:'active'},
    {id:'PM-005', name:'Weekend Getaway Promotion',  tag:'Weekend Special',     operators:'City Break Tours',                          bookingWindow:'Dec 1, 2025 – Feb 5, 2026',    stayDates:'Mar 8 – May 5, 2026',   expiry:'27d',rooms:28,    matched:28,  revenue:'$33,600', status:'active'},
    {id:'PM-006', name:'Mid-Week Business Traveler', tag:'Mid-Week Discount',   operators:'City Break Tours, Sunshine Tours',          bookingWindow:'Dec 28, 2025 – Feb 28, 2026',  stayDates:'Mar 15 – Jun 7, 2026',  expiry:'41d',rooms:19,    matched:19,  revenue:'$22,800', status:'active'},
    {id:'PM-007', name:'Senior Citizen Discount',    tag:'Senior Discount',     operators:'Beach Holidays Ltd, Sunshine Tours',        bookingWindow:'Dec 15, 2025 – Apr 4, 2026',   stayDates:'Mar 12 – Jun 27, 2026', expiry:'80d',rooms:12,    matched:12,  revenue:'$14,400', status:'active'},
  ];

  var el = document.getElementById('promotionsGrid');
  if (!el || typeof agGrid === 'undefined') return;

  var colDefs = [
    { field: 'id',            headerName: 'ID',              width: 80,
      cellRenderer: function(p) { return '<span class="ac-id">' + p.value + '</span>'; } },
    { field: 'name',          headerName: 'Promotion name',  flex: 1.5 },
    { field: 'tag',           headerName: 'Offer tag',       width: 140,
      cellRenderer: function(p) { return '<span class="pm-tag">' + p.value + '</span>'; } },
    { field: 'operators',     headerName: 'Tour operators',  flex: 1.5 },
    { field: 'bookingWindow', headerName: 'Booking window',  flex: 1.2 },
    { field: 'stayDates',     headerName: 'Stay dates',      flex: 1.2 },
    { field: 'expiry',        headerName: 'Expiry',          width: 80,
      cellRenderer: function(p) {
        if (!p.value || p.value === '—') return '—';
        return '<span class="pm-expiry">' + p.value + '</span>';
      }
    },
    { field: 'rooms',         headerName: 'Rooms',           width: 80,  type: 'numericColumn',
      valueFormatter: function(p) { return p.value !== null ? p.value : '—'; } },
    { field: 'revenue',       headerName: 'Revenue',         width: 100, cellStyle: { textAlign: 'right' } },
    { field: 'status',        headerName: 'Status',          width: 110,
      cellRenderer: function(p) {
        return '<span class="pm-status ' + p.value + '">' + p.value.charAt(0).toUpperCase() + p.value.slice(1) + '</span>';
      }
    },
    { headerName: 'Actions',  width: 120, sortable: false, resizable: false,
      cellRenderer: function(p) {
        if (p.data.status === 'active') {
          return '<div class="ac-actions-cell">' +
                 '<button class="to-icon-btn pm-send-btn" title="Send">' + sendSvg + '</button>' +
                 '<button class="to-icon-btn" title="View">' + viewSvg + '</button>' +
                 '<button class="to-icon-btn" title="Delete">' + deleteSvg + '</button></div>';
        }
        return '<div class="ac-actions-cell">' +
               '<button class="to-icon-btn" title="View">' + viewSvg + '</button>' +
               '<button class="to-icon-btn" title="Delete">' + deleteSvg + '</button></div>';
      }
    },
  ];

  agGrid.createGrid(el, {
    theme: sharedTheme,
    columnDefs: colDefs,
    rowData: PROMOTIONS_DATA,
    rowHeight: 42,
    headerHeight: 42,
    suppressMovableColumns: true,
    suppressCellFocus: true,
    defaultColDef: { sortable: true, resizable: true, floatingFilter: false, minWidth: 200 },
  });
})();

/* ─── COMMUNICATIONS GRID ─── */
(function() {
  var contractSvg = '<svg viewBox="0 0 24 24" fill="currentColor" width="13" height="13"><path d="M14 2H6c-1.1 0-2 .9-2 2v16c0 1.1.9 2 2 2h12c1.1 0 2-.9 2-2V8l-6-6zm2 16H8v-2h8v2zm0-4H8v-2h8v2zm-3-5V3.5L18.5 9H13z"/></svg>';
  var promoSvg    = '<svg viewBox="0 0 24 24" fill="currentColor" width="13" height="13"><path d="M21.41 11.58l-9-9C12.05 2.22 11.55 2 11 2H4c-1.1 0-2 .9-2 2v7c0 .55.22 1.05.59 1.42l9 9c.36.36.86.58 1.41.58s1.05-.22 1.41-.59l7-7c.37-.36.59-.86.59-1.41s-.23-1.06-.59-1.42zM5.5 7C4.67 7 4 6.33 4 5.5S4.67 4 5.5 4 7 4.67 7 5.5 6.33 7 5.5 7z"/></svg>';

  var COMMS_DATA = [
    {type:'contract', typeLabel:'Contract Sent',   datetime:'Feb 5, 2026 06:23 PM', from:'Sales Team',     to:'City Break Tours',   message:'Contract ready for signature: City Explorer Package (Contract ID: CT-020). This 10-month agreement (February 1 – November 30, 2026) includes 120 room...'},
    {type:'contract', typeLabel:'Contract Sent',   datetime:'Feb 5, 2026 06:29 PM', from:'Revenue Manager',to:'Beach Holidays Ltd',  message:'We are delighted to present the Beach Paradise 2026 contract (Contract ID: CT-015) for your review. This seasonal agreement runs from March 1 – October 31...'},
    {type:'contract', typeLabel:'Contract Sent',   datetime:'Feb 5, 2026 06:29 PM', from:'Sales Manager',  to:'Global Adventures',  message:'Your annual contract for European Tour Package (Contract ID: CT-010) is ready for review. This agreement covers 150 room allocations for the entire year 2026.'},
    {type:'contract', typeLabel:'Contract Sent',   datetime:'Feb 5, 2026 06:20 PM', from:'Sales Manager',  to:'Sunshine Tours',     message:'Please find attached the contract agreement for Q1 2026 Package (Contract ID: CT-001). This contract includes 100 room allocations for the period January 1...'},
    {type:'promo',    typeLabel:'Promotion Sent',  datetime:'Feb 5, 2026 06:20 PM', from:'Sales Manager',  to:'Beach Holidays Ltd',  message:'Book now and save 15% on all room types! Early Bird promotion available for bookings made 90+ days in advance. Valid for stays March 1 – October 31.'},
    {type:'promo',    typeLabel:'Promotion Sent',  datetime:'Feb 5, 2026 06:20 PM', from:'Revenue Manager',to:'Global Adventures',  message:'Exciting news! We are pleased to announce our Summer Kids Free promotion. Book 7 nights and kids stay free. Valid for bookings made until June 30, 2026 for stays...'},
    {type:'promo',    typeLabel:'Promotion Sent',  datetime:'Feb 5, 2026 06:20 PM', from:'Revenue Manager',to:'Sunshine Tours',     message:'Exciting news! We are pleased to announce our Summer Kids Free promotion. Book 7 nights and kids stay free. Valid for bookings made until June 30, 2026 for stays...'},
    {type:'contract', typeLabel:'Contract Sent',   datetime:'Feb 5, 2026 06:27 PM', from:'Sales Manager',  to:'Sunshine Tours',     message:'Contract has been created and is ready for review. Period: 2026-02-01 to 2026-02-22. Total allocation: 120 rooms. Commission / Payment terms:...'},
  ];

  var el = document.getElementById('commsGrid');
  if (!el || typeof agGrid === 'undefined') return;

  var colDefs = [
    { field: 'typeLabel', headerName: 'Type',         width: 150,
      cellRenderer: function(p) {
        var iconCls = p.data.type === 'contract' ? 'cn-type-contract' : 'cn-type-promo';
        var svg     = p.data.type === 'contract' ? contractSvg : promoSvg;
        return '<div style="display:flex;flex-direction:column;align-items:center;gap:2px">' +
               '<div class="cn-type-icon ' + iconCls + '">' + svg + '</div>' +
               '<span class="cn-type-label">' + p.value + '</span></div>';
      }
    },
    { field: 'datetime', headerName: 'Date & time',  width: 180 },
    { field: 'from',     headerName: 'From',         flex: 1 },
    { field: 'to',       headerName: 'To',           flex: 1 },
    { field: 'message',  headerName: 'Message',      flex: 3,
      cellRenderer: function(p) {
        return '<div class="cn-msg-preview">' + p.value + '</div>';
      }
    },
  ];

  agGrid.createGrid(el, {
    theme: sharedTheme,
    columnDefs: colDefs,
    rowData: COMMS_DATA,
    rowHeight: 42,
    headerHeight: 42,
    suppressMovableColumns: true,
    suppressCellFocus: true,
    defaultColDef: { sortable: true, resizable: true, floatingFilter: true, minWidth: 200 },
  });
})();

/* ─── CONTRACT WIZARD ─── */
(function() {
  const overlay   = document.getElementById('contractWizard');
  const previewMd = document.getElementById('contractPreviewModal');
  const sendMd    = document.getElementById('sendContractModal');
  if (!overlay) return;

  var currentStep = 1;

  function showWizard() {
    overlay.style.display = '';
    goStep(1);
  }
  function hideWizard() {
    overlay.style.display = 'none';
  }
  function goStep(n) {
    currentStep = n;
    [1,2,3,4].forEach(function(i) {
      var panel = document.getElementById('cwPanel' + i);
      if (panel) panel.style.display = i === n ? '' : 'none';
    });
    document.querySelectorAll('.cw-step-item').forEach(function(item) {
      var s = parseInt(item.dataset.step, 10);
      item.classList.toggle('active', s === n);
      item.classList.toggle('done', s < n);
    });
    var lbl = document.getElementById('cwStepLabel');
    if (lbl) lbl.textContent = 'Step ' + n + ' of 4';
    overlay.scrollTo(0, 0);
  }

  // Open wizard from any NEW CONTRACT button
  document.addEventListener('click', function(e) {
    var btn = e.target.closest('button, a');
    if (btn && /new\s*contract/i.test(btn.textContent)) {
      showWizard();
    }
  });

  // Back button in header
  document.getElementById('cwBackBtn')?.addEventListener('click', function() {
    if (currentStep > 1) goStep(currentStep - 1);
    else hideWizard();
  });

  // Step nav
  ['cwCancel1','cwCancel2','cwCancel3'].forEach(function(id) {
    document.getElementById(id)?.addEventListener('click', hideWizard);
  });
  document.getElementById('cwNext1')?.addEventListener('click', function() { goStep(2); });
  document.getElementById('cwBack2')?.addEventListener('click', function() { goStep(1); });
  document.getElementById('cwNext2')?.addEventListener('click', function() { goStep(3); });
  document.getElementById('cwBack3')?.addEventListener('click', function() { goStep(2); });
  document.getElementById('cwNext3')?.addEventListener('click', function() { goStep(4); updateReview(); });
  document.getElementById('cwBack4')?.addEventListener('click', function() { goStep(3); });

  function updateReview() {
    var to = document.getElementById('cwTourOp')?.value || '–';
    var name = document.getElementById('cwContractName')?.value || '–';
    var start = document.getElementById('cwStartDate')?.value || '–';
    var end = document.getElementById('cwEndDate')?.value || '–';
    var el;
    if ((el = document.getElementById('revTourOp'))) el.textContent = to;
    if ((el = document.getElementById('revContractName'))) el.textContent = name;
    if ((el = document.getElementById('revStartDate'))) el.textContent = start;
    if ((el = document.getElementById('revEndDate'))) el.textContent = end;
    if ((el = document.getElementById('previewTourOp'))) el.textContent = to;
    if ((el = document.getElementById('previewContractName'))) el.textContent = name;
  }

  // Add season rows
  document.getElementById('cwAddSeason')?.addEventListener('click', function() {
    var cont = document.getElementById('cwSeasons');
    var row = document.createElement('div');
    row.className = 'cw-season-row';
    row.innerHTML = '<input class="cw-input cw-season-name" type="text" placeholder="Season Name *"/>'
      + '<input class="cw-input cw-season-from" type="date"/>'
      + '<input class="cw-input cw-season-to" type="date"/>'
      + '<button class="cw-remove-btn">REMOVE</button>';
    cont.appendChild(row);
  });

  // Remove season rows (delegation)
  document.getElementById('cwSeasons')?.addEventListener('click', function(e) {
    if (e.target.classList.contains('cw-remove-btn')) {
      e.target.closest('.cw-season-row').remove();
    }
  });

  // Add blackout row
  document.getElementById('cwAddBlackout')?.addEventListener('click', function() {
    var cont = document.getElementById('cwBlackouts');
    var row = document.createElement('div');
    row.className = 'cw-blackout-row';
    row.innerHTML = '<div class="cw-blackout-field"><label class="cw-label">From Date <span class="cw-req">*</span></label><input class="cw-input" type="date"/></div>'
      + '<div class="cw-blackout-field"><label class="cw-label">To Date <span class="cw-req">*</span></label><input class="cw-input" type="date"/></div>'
      + '<div class="cw-blackout-field cw-blackout-reason"><label class="cw-label">Reason</label><input class="cw-input" type="text" placeholder="e.g., National Holiday"/></div>'
      + '<button class="cw-remove-outline-btn">REMOVE</button>';
    cont.appendChild(row);
  });
  document.getElementById('cwBlackouts')?.addEventListener('click', function(e) {
    if (e.target.classList.contains('cw-remove-outline-btn')) {
      e.target.closest('.cw-blackout-row').remove();
    }
  });

  // Collapse buttons
  overlay.addEventListener('click', function(e) {
    var btn = e.target.closest('.cw-collapse-btn');
    if (btn) {
      var target = document.getElementById(btn.dataset.target);
      if (target) target.style.display = target.style.display === 'none' ? '' : 'none';
    }
  });

  // Preview contract
  document.getElementById('cwPreview')?.addEventListener('click', function() {
    updateReview();
    var start = document.getElementById('cwStartDate')?.value || '–';
    var end = document.getElementById('cwEndDate')?.value || '–';
    var pf = document.getElementById('previewFrom');
    var pt = document.getElementById('previewTo');
    if (pf) pf.textContent = start;
    if (pt) pt.textContent = end;
    if (previewMd) previewMd.style.display = 'flex';
  });
  document.getElementById('previewClose')?.addEventListener('click', function() { previewMd.style.display = 'none'; });
  document.getElementById('previewCloseBtn')?.addEventListener('click', function() { previewMd.style.display = 'none'; });

  // Save & Send
  document.getElementById('cwSaveSend')?.addEventListener('click', function() {
    if (sendMd) sendMd.style.display = 'flex';
  });
  document.getElementById('sendEmailCancel')?.addEventListener('click', function() { sendMd.style.display = 'none'; });
  document.getElementById('sendEmailBtn')?.addEventListener('click', function() {
    sendMd.style.display = 'none';
    hideWizard();
    alert('Contract sent successfully!');
  });

  // Save contract — add to tables and close
  document.getElementById('cwSave')?.addEventListener('click', function() {
    var to       = document.getElementById('cwTourOp')?.value || '—';
    var name     = document.getElementById('cwContractName')?.value;
    var start    = document.getElementById('cwStartDate')?.value || '';
    var end      = document.getElementById('cwEndDate')?.value || '';

    if (!name || !to) { hideWizard(); return; }

    // Generate next contract ID
    var ids = (window._contractsData || []).map(function(r) {
      return parseInt((r.id || 'CT-000').replace('CT-',''), 10);
    }).filter(function(n) { return !isNaN(n); });
    var nextId = 'CT-' + String((ids.length ? Math.max.apply(null, ids) : 0) + 1).padStart(3, '0');

    var newRow = {
      id: nextId, name: name, operator: to,
      bookingWindow: start && end ? start + ' – ' + end : '—',
      stayWindow: start && end ? start + ' – ' + end : '—',
      totalValue: '$0', bookings: 0, revenue: '$0', roomsSold: '0/—', status: 'ACCEPTED'
    };

    if (window._contractsData) window._contractsData.push(newRow);
    if (window._contractsGridApi) window._contractsGridApi.setGridOption('rowData', window._contractsData);

    hideWizard();
  });

  // Also Save & Send should add the contract
  var _origSendBtn = document.getElementById('sendEmailBtn');
  if (_origSendBtn) {
    _origSendBtn.addEventListener('click', function() {
      document.getElementById('cwSave')?.click();
    });
  }
})();

/* ─── DATE RANGE PICKER ─── */
(function() {
  var trigger = document.getElementById('drpTrigger');
  var dropdown = document.getElementById('drpDropdown');
  var label = document.getElementById('drpLabel');
  var rangeText = document.getElementById('drpRangeText');
  if (!trigger || !dropdown) return;

  var MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  var today = new Date();
  var viewYear1 = today.getFullYear();
  var viewMonth1 = today.getMonth(); // 0-indexed
  var startDate = null;
  var endDate = null;
  var activeShortcut = null;
  var selecting = false; // true after first click, before second

  function dateKey(d) { return d.getFullYear() + '-' + d.getMonth() + '-' + d.getDate(); }
  function fmtDate(d) {
    return (d.getMonth()+1).toString().padStart(2,'0') + '/' + d.getDate().toString().padStart(2,'0') + '/' + d.getFullYear();
  }
  function sameDay(a, b) { return a && b && a.getFullYear()===b.getFullYear() && a.getMonth()===b.getMonth() && a.getDate()===b.getDate(); }
  function isBetween(d, a, b) {
    if (!a || !b) return false;
    var s = a < b ? a : b, e = a < b ? b : a;
    return d > s && d < e;
  }

  function renderMonth(year, month, daysId, titleId) {
    var title = document.getElementById(titleId);
    var container = document.getElementById(daysId);
    if (!title || !container) return;
    title.textContent = MONTHS[month] + ' ' + year;
    container.innerHTML = '';
    var firstDay = new Date(year, month, 1).getDay();
    var daysInMonth = new Date(year, month+1, 0).getDate();
    // Empty cells
    for (var i = 0; i < firstDay; i++) {
      var e = document.createElement('div');
      e.className = 'drp-day empty';
      container.appendChild(e);
    }
    for (var d = 1; d <= daysInMonth; d++) {
      var cell = document.createElement('div');
      cell.className = 'drp-day';
      var numSpan = document.createElement('span');
      numSpan.textContent = d;
      cell.appendChild(numSpan);
      var thisDate = new Date(year, month, d);
      cell.dataset.ts = thisDate.getTime();
      if (sameDay(thisDate, today)) cell.classList.add('today');
      if (startDate && sameDay(thisDate, startDate)) cell.classList.add('range-start');
      if (endDate && sameDay(thisDate, endDate)) cell.classList.add('range-end');
      if (startDate && endDate && isBetween(thisDate, startDate, endDate)) cell.classList.add('in-range');
      container.appendChild(cell);
    }
  }

  function renderBoth() {
    var year2 = viewMonth1 === 11 ? viewYear1+1 : viewYear1;
    var month2 = viewMonth1 === 11 ? 0 : viewMonth1+1;
    renderMonth(viewYear1, viewMonth1, 'drpDays1', 'drpTitle1');
    renderMonth(year2, month2, 'drpDays2', 'drpTitle2');
  }

  function updateFooter() {
    if (startDate && endDate) {
      rangeText.textContent = fmtDate(startDate) + ' - ' + fmtDate(endDate);
    } else if (startDate) {
      rangeText.textContent = fmtDate(startDate) + ' - ...';
    } else {
      rangeText.textContent = '';
    }
  }

  function openDropdown() {
    dropdown.style.display = '';
    trigger.classList.add('active');
    renderBoth();
    updateFooter();
  }
  function closeDropdown() {
    dropdown.style.display = 'none';
    trigger.classList.remove('active');
  }

  trigger.addEventListener('click', function(e) {
    e.stopPropagation();
    if (dropdown.style.display === 'none' || !dropdown.style.display) {
      openDropdown();
    } else {
      closeDropdown();
    }
  });

  // Click on calendar days
  dropdown.addEventListener('click', function(e) {
    var day = (e.target.closest('.drp-day') && !e.target.closest('.drp-day').classList.contains('empty')) ? e.target.closest('.drp-day') : null;
    if (day) {
      var ts = parseInt(day.dataset.ts, 10);
      var clicked = new Date(ts);
      // Clear shortcut highlight
      dropdown.querySelectorAll('.drp-shortcut').forEach(function(b) { b.classList.remove('active'); });
      activeShortcut = null;
      if (!startDate || (startDate && endDate)) {
        // Start new selection
        startDate = clicked;
        endDate = null;
        selecting = true;
      } else {
        // Set end date
        if (clicked < startDate) {
          endDate = startDate;
          startDate = clicked;
        } else {
          endDate = clicked;
        }
        selecting = false;
      }
      renderBoth();
      updateFooter();
    }
  });

  // Navigation
  document.getElementById('drpPrev')?.addEventListener('click', function() {
    viewMonth1--; if (viewMonth1 < 0) { viewMonth1 = 11; viewYear1--; }
    renderBoth();
  });
  document.getElementById('drpNext')?.addEventListener('click', function() {
    viewMonth1++; if (viewMonth1 > 11) { viewMonth1 = 0; viewYear1++; }
    renderBoth();
  });
  document.getElementById('drpPrevPrev')?.addEventListener('click', function() {
    viewYear1--; renderBoth();
  });
  document.getElementById('drpNextNext')?.addEventListener('click', function() {
    viewYear1++; renderBoth();
  });

  // Shortcuts
  dropdown.addEventListener('click', function(e) {
    var btn = e.target.closest('.drp-shortcut');
    if (!btn) return;
    var range = btn.dataset.range;
    var s = new Date(today), en = new Date(today);
    if (range === 'today') {
      s = new Date(today.getFullYear(), today.getMonth(), today.getDate());
      en = new Date(s);
    } else if (range === 'tomorrow') {
      s = new Date(today.getFullYear(), today.getMonth(), today.getDate()+1);
      en = new Date(s);
    } else if (range === 'cur-week') {
      var dow = today.getDay();
      s = new Date(today.getFullYear(), today.getMonth(), today.getDate()-dow);
      en = new Date(s.getFullYear(), s.getMonth(), s.getDate()+6);
    } else if (range === 'cur-month') {
      s = new Date(today.getFullYear(), today.getMonth(), 1);
      en = new Date(today.getFullYear(), today.getMonth()+1, 0);
    } else if (range === 'next-7') {
      s = new Date(today.getFullYear(), today.getMonth(), today.getDate());
      en = new Date(today.getFullYear(), today.getMonth(), today.getDate()+6);
    } else if (range === 'next-14') {
      s = new Date(today.getFullYear(), today.getMonth(), today.getDate());
      en = new Date(today.getFullYear(), today.getMonth(), today.getDate()+13);
    } else if (range === 'next-30') {
      s = new Date(today.getFullYear(), today.getMonth(), today.getDate());
      en = new Date(today.getFullYear(), today.getMonth(), today.getDate()+29);
    } else if (range === 'next-90') {
      s = new Date(today.getFullYear(), today.getMonth(), today.getDate());
      en = new Date(today.getFullYear(), today.getMonth(), today.getDate()+89);
    }
    startDate = s; endDate = en;
    viewYear1 = s.getFullYear(); viewMonth1 = s.getMonth();
    dropdown.querySelectorAll('.drp-shortcut').forEach(function(b) { b.classList.remove('active'); });
    btn.classList.add('active');
    activeShortcut = range;
    renderBoth();
    updateFooter();
  });

  // Apply
  document.getElementById('drpApply')?.addEventListener('click', function() {
    if (startDate) {
      var txt = fmtDate(startDate);
      var isoStart = startDate.getFullYear() + '-' +
        String(startDate.getMonth()+1).padStart(2,'0') + '-' +
        String(startDate.getDate()).padStart(2,'0');
      var isoEnd = isoStart;
      if (endDate && !sameDay(startDate, endDate)) {
        txt += ' \u2013 ' + fmtDate(endDate);
        isoEnd = endDate.getFullYear() + '-' +
          String(endDate.getMonth()+1).padStart(2,'0') + '-' +
          String(endDate.getDate()).padStart(2,'0');
      }
      label.textContent = txt;
      if (typeof window.anApplyDateRange === 'function') {
        window.anApplyDateRange(isoStart, isoEnd);
      }
    }
    closeDropdown();
  });

  // Cancel
  document.getElementById('drpCancel')?.addEventListener('click', function() {
    closeDropdown();
  });

  // Close on outside click
  document.addEventListener('click', function(e) {
    if (!document.getElementById('drpWrap')?.contains(e.target)) {
      closeDropdown();
    }
  });
})();

/* ─── SIDEBAR COLLAPSE ─── */
(function() {
  var btn = document.getElementById('sidebarCollapseBtn');
  var sidebar = document.querySelector('.sidebar');
  var main = document.querySelector('.main');
  if (!btn || !sidebar) return;

  btn.addEventListener('click', function() {
    sidebar.classList.toggle('collapsed');
    if (main) {
      if (sidebar.classList.contains('collapsed')) {
        main.style.marginLeft = '52px';
      } else {
        main.style.marginLeft = '220px';
      }
    }
  });
})();

/* ─── FIGMA EXPORT TOOLBAR ─── */
(function() {
  var triggerBtn  = document.getElementById('figmaExportBtn');
  var fab         = document.getElementById('figmaFab');
  var toolbar     = document.getElementById('figmaToolbar');
  var overlay     = document.getElementById('figmaTbOverlay');
  var closeBtn    = document.getElementById('figmaTbClose');
  var sendBtn     = document.getElementById('figmaTbSend');
  var svgBtn      = document.getElementById('figmaTbSvg');
  var statusEl    = document.getElementById('figmaTbStatus');
  var connDot     = document.getElementById('figmaTbConnDot');
  var connLabel   = document.getElementById('figmaTbConnLabel');
  var connHint    = document.getElementById('figmaTbConnHint');
  var optEls      = document.querySelectorAll('.figma-tb-opt');

  if (!toolbar) return;

  /* ── Figma connection polling ── */
  function isFigmaConnected() {
    return !!(window.figma && typeof window.figma.captureForDesign === 'function');
  }
  function updateConnectionUI() {
    var connected = isFigmaConnected();
    if (connDot) {
      connDot.className = 'figma-tb-conn-dot ' + (connected ? 'connected' : 'disconnected');
    }
    if (connLabel) {
      connLabel.className = 'figma-tb-conn-label ' + (connected ? 'connected' : '');
      connLabel.textContent = connected ? 'Figma desktop connected' : 'Figma desktop not detected';
    }
    if (connHint) {
      connHint.style.display = connected ? 'none' : '';
    }
    if (sendBtn) {
      sendBtn.title = connected ? 'Send directly to Figma' : 'Requires Figma desktop + HTML to Design plugin';
    }
    /* update FAB dot + class */
    if (fab) {
      fab.classList.toggle('connected', connected);
    }
  }
  updateConnectionUI();
  setInterval(updateConnectionUI, 2000);

  /* ── open / close ── */
  function openToolbar() {
    toolbar.classList.add('open');
    overlay.classList.add('open');
    toolbar.setAttribute('aria-hidden', 'false');
    if (triggerBtn) triggerBtn.setAttribute('aria-expanded', 'true');
    if (fab) fab.classList.add('active');
  }
  function closeToolbar() {
    toolbar.classList.remove('open');
    overlay.classList.remove('open');
    toolbar.setAttribute('aria-hidden', 'true');
    if (triggerBtn) triggerBtn.setAttribute('aria-expanded', 'false');
    if (fab) fab.classList.remove('active');
  }

  if (triggerBtn) triggerBtn.addEventListener('click', function() {
    toolbar.classList.contains('open') ? closeToolbar() : openToolbar();
  });
  if (fab) fab.addEventListener('click', function() {
    toolbar.classList.contains('open') ? closeToolbar() : openToolbar();
  });
  closeBtn.addEventListener('click', closeToolbar);
  overlay.addEventListener('click', closeToolbar);
  document.addEventListener('keydown', function(e) {
    if (e.key === 'Escape' && toolbar.classList.contains('open')) closeToolbar();
  });

  /* ── keep active class in sync with radio selection ── */
  optEls.forEach(function(opt) {
    opt.addEventListener('click', function() {
      optEls.forEach(function(o) { o.classList.remove('figma-tb-opt--active'); });
      opt.classList.add('figma-tb-opt--active');
    });
  });

  /* ── section map ── */
  var sectionMap = {
    'full':           null,
    'targets':        '#targets',
    'revenue-trend':  '#revenue-trend',
    'demand-calendar':'#demand-calendar',
    'weekView':       '#weekView',
    'room-type':      '#room-type'
  };

  /* ── layered SVG export ── */
  function exportLayeredSVG(rootEl, filename) {
    statusEl.className = 'figma-tb-status';
    statusEl.textContent = 'Building layers…';

    var rootRect = rootEl.getBoundingClientRect();
    var ox = rootRect.left + window.scrollX;
    var oy = rootRect.top + window.scrollY;
    var W  = Math.round(rootRect.width);
    var H  = Math.round(rootRect.height);

    var svgNS = 'http://www.w3.org/2000/svg';
    var svg = document.createElementNS(svgNS, 'svg');
    svg.setAttribute('xmlns', svgNS);
    svg.setAttribute('width',   W);
    svg.setAttribute('height',  H);
    svg.setAttribute('viewBox', '0 0 ' + W + ' ' + H);

    var usedIds = {};
    function uniqueId(base) {
      var safe = (base || 'layer').replace(/[^a-zA-Z0-9_-]/g, '_').replace(/^_+/, '') || 'layer';
      if (!usedIds[safe]) { usedIds[safe] = 1; return safe; }
      return safe + '_' + (++usedIds[safe]);
    }

    function rgbaToHex(rgba) {
      var m = rgba.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?\)/);
      if (!m) return null;
      var a = m[4] !== undefined ? parseFloat(m[4]) : 1;
      if (a === 0) return null;
      return { hex: '#' + [m[1],m[2],m[3]].map(function(n){ return ('0'+parseInt(n).toString(16)).slice(-2); }).join(''), alpha: a };
    }

    function processEl(el) {
      var rect = el.getBoundingClientRect();
      if (rect.width < 1 || rect.height < 1) return;

      var cs = getComputedStyle(el);
      if (cs.display === 'none' || cs.visibility === 'hidden') return;
      var opacity = parseFloat(cs.opacity);
      if (opacity === 0) return;

      var x = Math.round(rect.left + window.scrollX - ox);
      var y = Math.round(rect.top  + window.scrollY - oy);
      var w = Math.round(rect.width);
      var h = Math.round(rect.height);

      /* skip elements entirely outside root bounds */
      if (x + w < 0 || y + h < 0 || x > W || y > H) return;

      /* clamp */
      var cx = Math.max(0, x), cy = Math.max(0, y);
      var cw = Math.min(w, W - cx), ch = Math.min(h, H - cy);
      if (cw < 1 || ch < 1) return;

      var labelBase = el.id || (el.dataset && el.dataset.layer) ||
                      (el.className && typeof el.className === 'string' ? el.className.trim().split(/\s+/)[0] : '') ||
                      el.tagName.toLowerCase();
      var g = document.createElementNS(svgNS, 'g');
      g.setAttribute('id', uniqueId(labelBase));
      if (opacity < 1) g.setAttribute('opacity', opacity);

      var added = false;

      /* background fill */
      var bg = rgbaToHex(cs.backgroundColor);
      if (bg) {
        var bgR = document.createElementNS(svgNS, 'rect');
        bgR.setAttribute('x', cx); bgR.setAttribute('y', cy);
        bgR.setAttribute('width', cw); bgR.setAttribute('height', ch);
        bgR.setAttribute('fill', bg.hex);
        if (bg.alpha < 1) bgR.setAttribute('fill-opacity', bg.alpha);
        var br = parseFloat(cs.borderTopLeftRadius) || 0;
        if (br > 0) { bgR.setAttribute('rx', br); bgR.setAttribute('ry', br); }
        g.appendChild(bgR);
        added = true;
      }

      /* border */
      var bw = parseFloat(cs.borderTopWidth) || 0;
      var bc = bw > 0 ? rgbaToHex(cs.borderTopColor) : null;
      if (bw > 0 && bc) {
        var bR = document.createElementNS(svgNS, 'rect');
        var bx = cx + bw/2, by = cy + bw/2;
        bR.setAttribute('x', bx); bR.setAttribute('y', by);
        bR.setAttribute('width',  Math.max(0, cw - bw));
        bR.setAttribute('height', Math.max(0, ch - bw));
        bR.setAttribute('fill',   'none');
        bR.setAttribute('stroke', bc.hex);
        bR.setAttribute('stroke-width', bw);
        if (bc.alpha < 1) bR.setAttribute('stroke-opacity', bc.alpha);
        var br2 = parseFloat(cs.borderTopLeftRadius) || 0;
        if (br2 > 0) { bR.setAttribute('rx', br2); bR.setAttribute('ry', br2); }
        g.appendChild(bR);
        added = true;
      }

      /* text (only direct text nodes) */
      var textContent = '';
      el.childNodes.forEach(function(node) {
        if (node.nodeType === 3) textContent += node.textContent.trim();
      });
      if (textContent) {
        var fontSize   = parseFloat(cs.fontSize)   || 14;
        var fontWeight = cs.fontWeight             || '400';
        var fontFamily = cs.fontFamily             || 'sans-serif';
        var color      = rgbaToHex(cs.color);
        var ta         = cs.textAlign;
        var pl         = parseFloat(cs.paddingLeft)  || 0;
        var pr         = parseFloat(cs.paddingRight) || 0;

        var tx, anchor;
        if (ta === 'center') { tx = cx + cw/2; anchor = 'middle'; }
        else if (ta === 'right') { tx = cx + cw - pr; anchor = 'end'; }
        else { tx = cx + pl; anchor = 'start'; }

        var tEl = document.createElementNS(svgNS, 'text');
        tEl.setAttribute('x', tx);
        tEl.setAttribute('y', cy + ch/2 + fontSize * 0.36);
        tEl.setAttribute('font-family', fontFamily);
        tEl.setAttribute('font-size',   fontSize);
        tEl.setAttribute('font-weight', fontWeight);
        tEl.setAttribute('text-anchor', anchor);
        tEl.setAttribute('fill', color ? color.hex : '#000000');
        if (color && color.alpha < 1) tEl.setAttribute('fill-opacity', color.alpha);
        /* clip text to element width */
        tEl.setAttribute('textLength',       Math.max(1, cw - pl - pr));
        tEl.setAttribute('lengthAdjust',     'spacingAndGlyphs');
        tEl.textContent = textContent;
        g.appendChild(tEl);
        added = true;
      }

      if (added) svg.appendChild(g);
    }

    /* process root then all descendants in DOM order */
    processEl(rootEl);
    rootEl.querySelectorAll('*').forEach(processEl);

    var svgStr = new XMLSerializer().serializeToString(svg);
    var blob = new Blob([svgStr], { type: 'image/svg+xml;charset=utf-8' });
    var url  = URL.createObjectURL(blob);
    var a    = document.createElement('a');
    a.download = filename + '.svg';
    a.href = url;
    a.click();
    URL.revokeObjectURL(url);

    statusEl.className = 'figma-tb-status success';
    statusEl.textContent = '✓ SVG downloaded — File → Import in Figma for editable layers';
    sendBtn.disabled = false;
    setTimeout(function() { statusEl.textContent = ''; statusEl.className = 'figma-tb-status'; }, 6000);
  }

  /* ── send ── */
  sendBtn.addEventListener('click', function() {
    var selected = document.querySelector('input[name="figmaScope"]:checked');
    var scope    = selected ? selected.value : 'full';
    var selector = sectionMap[scope];
    var targetEl = selector ? document.querySelector(selector) : document.body;
    var filename = 'duetto-' + (scope === 'full' ? 'dashboard' : scope);

    statusEl.className = 'figma-tb-status';
    statusEl.textContent = 'Preparing…';
    sendBtn.disabled = true;

    /* scroll section into view first */
    if (selector && targetEl) {
      targetEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }

    setTimeout(function() {
      if (isFigmaConnected()) {
        /* Send directly to Figma desktop via HTML to Design plugin */
        try {
          var p = window.figma.captureForDesign(selector ? { selector: selector } : undefined);
          Promise.resolve(p).then(function() {
            statusEl.className = 'figma-tb-status success';
            statusEl.textContent = '✓ Added to Figma';
            sendBtn.disabled = false;
            setTimeout(function() { statusEl.textContent = ''; statusEl.className = 'figma-tb-status'; }, 4000);
          }).catch(function(err) {
            statusEl.className = 'figma-tb-status error';
            statusEl.textContent = (err && err.message) || 'Figma capture failed — check the plugin is open';
            sendBtn.disabled = false;
          });
        } catch (e) {
          statusEl.className = 'figma-tb-status error';
          statusEl.textContent = 'Figma plugin error — reopen the HTML to Design plugin';
          sendBtn.disabled = false;
        }
      } else {
        statusEl.className = 'figma-tb-status error';
        statusEl.textContent = 'Open Figma desktop with the HTML to Design plugin first';
        sendBtn.disabled = false;
      }
    }, selector ? 500 : 0);
  });

  /* ── SVG download button ── */
  if (svgBtn) {
    svgBtn.addEventListener('click', function() {
      var selected = document.querySelector('input[name="figmaScope"]:checked');
      var scope    = selected ? selected.value : 'full';
      var selector = sectionMap[scope];
      var targetEl = selector ? document.querySelector(selector) : document.body;
      var filename = 'duetto-' + (scope === 'full' ? 'dashboard' : scope);
      svgBtn.disabled = true;
      if (selector && targetEl) targetEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
      setTimeout(function() {
        exportLayeredSVG(targetEl || document.body, filename);
        svgBtn.disabled = false;
      }, selector ? 500 : 0);
    });
  }
})();

/* ─── COPY ENTIRE SCREEN TO FIGMA ─── */
(function() {
  var copyBtn = document.getElementById('figmaCopyScreen');
  var toast   = document.getElementById('figmaToast');
  if (!copyBtn || !toast) return;

  function showToast(msg, type) {
    toast.textContent = msg;
    toast.className = 'figma-toast show' + (type ? ' ' + type : '');
    clearTimeout(toast._t);
    toast._t = setTimeout(function() {
      toast.className = 'figma-toast';
    }, 3500);
  }

  copyBtn.addEventListener('click', function() {
    copyBtn.disabled = true;
    copyBtn.textContent = 'Capturing…';

    /* Try direct Figma push first */
    if (window.figma && typeof window.figma.captureForDesign === 'function') {
      try {
        var p = window.figma.captureForDesign();
        Promise.resolve(p).then(function() {
          showToast('✓ Sent to Figma', 'success');
        }).catch(function(err) {
          showToast((err && err.message) || 'Figma capture failed', 'error');
        }).finally(function() {
          copyBtn.disabled = false;
          copyBtn.innerHTML = '<svg width="15" height="15" viewBox="0 0 15 15" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><rect x="1" y="1" width="9" height="9" rx="1.5"/><rect x="5" y="5" width="9" height="9" rx="1.5"/></svg> Copy screen';
        });
        return;
      } catch (e) {
        /* captureForDesign threw synchronously — fall through to PNG fallback */
      }
    }

    /* Fallback: PNG via html2canvas → clipboard */
    if (typeof html2canvas !== 'undefined') {
      html2canvas(document.body, { scale: 1, useCORS: true, logging: false }).then(function(canvas) {
        canvas.toBlob(function(blob) {
          if (navigator.clipboard && navigator.clipboard.write) {
            navigator.clipboard.write([new ClipboardItem({ 'image/png': blob })]).then(function() {
              showToast('✓ Screenshot copied — paste into Figma', 'success');
            }).catch(function() {
              /* Clipboard blocked, fall back to download */
              var url = URL.createObjectURL(blob);
              var a = document.createElement('a');
              a.href = url; a.download = 'dashboard-screen.png'; a.click();
              URL.revokeObjectURL(url);
              showToast('✓ PNG downloaded — drag into Figma', 'success');
            });
          } else {
            var url = URL.createObjectURL(blob);
            var a = document.createElement('a');
            a.href = url; a.download = 'dashboard-screen.png'; a.click();
            URL.revokeObjectURL(url);
            showToast('✓ PNG downloaded — drag into Figma', 'success');
          }
        }, 'image/png');
      }).catch(function() {
        showToast('Screenshot failed', 'error');
      }).finally(function() {
        copyBtn.disabled = false;
        copyBtn.innerHTML = '<svg width="15" height="15" viewBox="0 0 15 15" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><rect x="1" y="1" width="9" height="9" rx="1.5"/><rect x="5" y="5" width="9" height="9" rx="1.5"/></svg> Copy screen';
      });
    } else {
      showToast('html2canvas not loaded', 'error');
      copyBtn.disabled = false;
      copyBtn.innerHTML = '<svg width="15" height="15" viewBox="0 0 15 15" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><rect x="1" y="1" width="9" height="9" rx="1.5"/><rect x="5" y="5" width="9" height="9" rx="1.5"/></svg> Copy screen';
    }
  });
})();

/* ─── ANALYSIS FILTERS DROPDOWN ─── */
(function() {
  var btn      = document.getElementById('anFiltersBtn');
  var dropdown = document.getElementById('anFiltersDropdown');
  var badge    = document.getElementById('anFilterBadge');
  var resetBtn = document.getElementById('anFiltersReset');
  var selects  = ['anFbRoomType','anFbBoard','anFbSegment','anFbTourOp'];

  if (!btn || !dropdown) return;

  function updateBadge() {
    var count = selects.filter(function(id) {
      var el = document.getElementById(id);
      return el && el.value !== '';
    }).length;
    if (badge) {
      badge.textContent = count;
      badge.style.display = count ? '' : 'none';
    }
    if (count > 0) btn.classList.add('active');
    else btn.classList.toggle('active', dropdown.style.display !== 'none');
  }

  btn.addEventListener('click', function(e) {
    e.stopPropagation();
    var open = dropdown.style.display !== 'none';
    dropdown.style.display = open ? 'none' : '';
    btn.classList.toggle('active', !open);
    updateBadge();
  });

  document.addEventListener('click', function(e) {
    if (!dropdown.contains(e.target) && e.target !== btn) {
      dropdown.style.display = 'none';
      updateBadge();
    }
  });

  selects.forEach(function(id) {
    document.getElementById(id)?.addEventListener('change', updateBadge);
  });

  if (resetBtn) {
    resetBtn.addEventListener('click', function() {
      selects.forEach(function(id) {
        var el = document.getElementById(id);
        if (el) el.value = '';
      });
      updateBadge();
    });
  }
})();

/* ─── DESIGN SYSTEM SEARCH FIELDS ─── */
document.querySelectorAll('.ds-search-field').forEach(function(wrap) {
  var input = wrap.querySelector('.ds-sf-input');
  var clear = wrap.querySelector('.ds-sf-clear');
  if (!input || !clear) return;
  input.addEventListener('input', function() {
    clear.style.display = this.value ? '' : 'none';
  });
  clear.addEventListener('click', function() {
    input.value = '';
    clear.style.display = 'none';
    input.focus();
  });
});

/* ─── WEEK VIEW FILTER DROPDOWN — toggle open/close ─── */
(function() {
  const filtersBtn = document.getElementById('wvFiltersBtn');
  const filtersDd  = document.getElementById('wvFiltersDropdown');
  if (!filtersBtn || !filtersDd) return;

  filtersBtn.addEventListener('click', function(e) {
    e.stopPropagation();
    const isOpen = filtersDd.style.display !== 'none';
    if (!isOpen) _closeAllDropdowns('wvFiltersDropdown');
    filtersDd.style.display = isOpen ? 'none' : '';
    filtersBtn.classList.toggle('active', !isOpen);
  });

  document.addEventListener('click', function(e) {
    if (!document.getElementById('wvFiltersWrap')?.contains(e.target)) {
      filtersDd.style.display = 'none';
      filtersBtn.classList.remove('active');
    }
  });
})();

/* ─── WEEK VIEW RE-OPEN — now handled via Close Out modal Re-open type ─── */

/* ─── TOUR OPERATORS TABS (page-level) ─── */
(function() {
  const tabsEl = document.getElementById('toPageTabs');
  if (!tabsEl) return;
  tabsEl.addEventListener('click', e => {
    const tab = e.target.closest('.to-tab');
    if (!tab) return;
    document.querySelectorAll('#toPageTabs .to-tab').forEach(t => t.classList.remove('active'));
    tab.classList.add('active');
    const name = tab.dataset.tab;
    document.querySelectorAll('.to-tab-panel').forEach(p => p.style.display = 'none');
    const panel = document.getElementById(`toPanel-${name}`);
    if (panel) panel.style.display = '';
    /* Show Expand/Collapse only on Insights tab */
    var ecWrap = document.getElementById('insExpandCollapseWrap');
    if (ecWrap) ecWrap.style.display = name === 'insights' ? 'flex' : 'none';
  });
})();



/* ═══════════════════════════════════════════════════════════════
   INSIGHTS TAB v3
═══════════════════════════════════════════════════════════════ */
(function() {

  // ── Static data ───────────────────────────────────────────────
  const INS_TOS = [
    { id:'TUI Group',    name:'TUI Group',    color:'#0ea5e9' },
    { id:'Thomas Cook',  name:'Thomas Cook',  color:'#f59e0b' },
    { id:'Jet2holidays', name:'Jet2holidays', color:'#10b981' },
    { id:'FTI Group',    name:'FTI Group',    color:'#967EF3' },
    { id:'Sunwing',      name:'Sunwing',      color:'#ec4899' },
    { id:'Club Med',     name:'Club Med',     color:'#f97316' },
  ];
  const INS_CONTRACTS = [
    { id:'C-2025-TUI-001', name:'Summer 2025',    tos:['TUI Group','Thomas Cook'],  type:'Static', season:'Summer', releaseDate:'2025-04-15' },
    { id:'C-2025-TUI-002', name:'Winter 2025',    tos:['TUI Group','Jet2holidays'], type:'Static', season:'Winter', releaseDate:'2025-09-01' },
    { id:'C-2025-TC-001',  name:'Winter 2025 II', tos:['Thomas Cook','FTI Group'],  type:'Static', season:'Winter', releaseDate:'2025-09-20' },
    { id:'C-2025-J2-001',  name:'Annual 2025',    tos:['Jet2holidays','Sunwing'],   type:'Static', season:'Annual', releaseDate:'2025-01-10' },
    { id:'C-2025-FTI-001', name:'Q1 2025',        tos:['FTI Group','Club Med'],     type:'Static', season:'Q1',     releaseDate:'2024-12-01' },
    { id:'C-2025-SW-001',  name:'Spring 2025',    tos:['Sunwing','Thomas Cook'],    type:'Static', season:'Spring', releaseDate:'2025-02-28' },
  ];
  const INS_PROMOS = [
    { id:'EBB',   name:'Early Bird Booking', discount:12 },
    { id:'SPOFF', name:'Special Offers',      discount:8  },
    { id:'LAST',  name:'Last Minute',         discount:15 },
  ];
  const ALL_ORIGINS = ['UK','SP','US','MX'];
  const ALL_MEALS   = ['HB','AI','RO'];

  // ── State ─────────────────────────────────────────────────────
  let insView    = 'to';
  let insMetric   = 'rn';
  let insTopN     = 5;        // 3–15
  let insTopMetric= 'rn';     // 'rn'|'rev'|'los'|'avgAdults'|'leadTime'|'srcGeo'
  let insCompare  = false;
  let insCmpType   = 'ly';   // 'ly'|'stly'|'both'|'none'
  let insCmpMethod = 'dow';  // 'dow'|'date'|'custom'

  window.insToggleCmpType = function(t){
    insCmpType=t;
    ['ly','stly','both','none'].forEach(function(k){
      var b=document.getElementById('insCmp'+k.charAt(0).toUpperCase()+k.slice(1)+'Btn');
      if(b) b.classList.toggle('active',k===t);
    });
    insRender();
  };
  window.insToggleCmpMethod = function(m){
    insCmpMethod=m;
    ['DOW','Date','Custom'].forEach(function(k){
      var b=document.getElementById('insCmpMethod'+k);
      if(b) b.classList.toggle('active',k.toLowerCase()===m);
    });
    var cw=document.getElementById('insCmpCustomWrap');
    if(cw) cw.style.display=m==='custom'?'flex':'none';
    insRender();
  };

  let collOpen   = {};
  let promoOpen  = {};
  const insSelected = { meal:[], origin:[], to:[], contract:[] };

  // ── Dropdown multi-select logic ───────────────────────────────
  const DD_LABELS = {
    meal:     { all:'All Meal Plans', items:{ HB:'Half Board', AI:'All Inclusive', RO:'Room Only' } },
    origin:   { all:'All Origins',    items:{ UK:'UK', SP:'Spain', US:'US', MX:'Mexico' } },
    to:       { all:'All TOs',        items:{ 'TUI Group':'TUI Group','Thomas Cook':'Thomas Cook','Jet2holidays':'Jet2holidays','FTI Group':'FTI Group','Sunwing':'Sunwing','Club Med':'Club Med' } },
    contract: { all:'All Contracts',  items:{ 'C-2025-TUI-001':'Summer 2025','C-2025-TUI-002':'Winter 2025','C-2025-TC-001':'Winter 2025 II','C-2025-J2-001':'Annual 2025','C-2025-FTI-001':'Q1 2025','C-2025-SW-001':'Spring 2025' } },
  };

  function updateDDLabel(group) {
    const sel = insSelected[group];
    const lbl = document.getElementById('ins'+group.charAt(0).toUpperCase()+group.slice(1)+'DDLabel');
    const btn = document.getElementById('ins'+group.charAt(0).toUpperCase()+group.slice(1)+'DDBtn');
    if (!lbl) return;
    if (!sel || sel.length === 0) {
      lbl.textContent = DD_LABELS[group].all;
      if (btn) btn.classList.remove('has-selection');
    } else if (sel.length === 1) {
      lbl.textContent = DD_LABELS[group].items[sel[0]] || sel[0];
      if (btn) btn.classList.add('has-selection');
    } else {
      lbl.textContent = sel.length + ' selected';
      if (btn) btn.classList.add('has-selection');
    }
  }

  window.insToggleDD = function(group) {
    const panelId = 'ins'+group.charAt(0).toUpperCase()+group.slice(1)+'DDPanel';
    const panel = document.getElementById(panelId);
    if (!panel) return;
    const isOpen = panel.style.display !== 'none';
    // close all panels first
    ['meal','origin','to','contract'].forEach(function(g) {
      const p = document.getElementById('ins'+g.charAt(0).toUpperCase()+g.slice(1)+'DDPanel');
      if (p) p.style.display = 'none';
    });
    if (!isOpen) panel.style.display = 'block';
  };

  window.insDDChange = function(group, checkbox) {
    const val = checkbox.value;
    const panel = checkbox.closest('.ins-dd-panel');
    if (val === 'all') {
      // uncheck all specifics, mark all as checked
      panel.querySelectorAll('input[type=checkbox]').forEach(function(cb){ cb.checked = false; });
      checkbox.checked = true;
      insSelected[group] = [];
    } else {
      // uncheck All
      const allCb = panel.querySelector('input[value="all"]');
      if (allCb) allCb.checked = false;
      // rebuild selection
      const sel = [];
      panel.querySelectorAll('input[type=checkbox]:not([value="all"])').forEach(function(cb){
        if (cb.checked) sel.push(cb.value);
      });
      if (sel.length === 0) {
        // nothing left → revert to All
        if (allCb) allCb.checked = true;
        insSelected[group] = [];
      } else {
        insSelected[group] = sel;
      }
    }
    updateDDLabel(group);
    insRender();
  };

  // Keep ins dropdown open when clicking items inside it
  ['insMealDDPanel','insOriginDDPanel','insToDDPanel','insContractDDPanel'].forEach(function(pid) {
    var el = document.getElementById(pid);
    if (el) el.addEventListener('click', function(e) { e.stopPropagation(); });
  });

  // Close dropdowns when clicking outside
  document.addEventListener('click', function(e) {
    if (!e.target.closest('.ins-dd-wrap')) {
      ['meal','origin','to','contract'].forEach(function(g) {
        const p = document.getElementById('ins'+g.charAt(0).toUpperCase()+g.slice(1)+'DDPanel');
        if (p) p.style.display = 'none';
      });
    }
  });

  // ── Expand / Collapse all ─────────────────────────────────────
  window.insExpandAll = function(open) {
    Object.keys(collOpen).forEach(function(k){ collOpen[k] = open; });
    // Also set defaults for all current TOs
    INS_TOS.forEach(function(t){ collOpen['details_'+t.id] = open; });
    INS_CONTRACTS.forEach(function(c){ collOpen['cstay_'+c.id] = open; });
    insRender();
  };

  // ── Data helpers ──────────────────────────────────────────────
  // Per-TO base profiles — gives each operator meaningfully different numbers
  const TO_PROFILES = {
    'TUI Group':    { rnBase:420, adrBase:112, lyMult:0.91, slyMult:0.84, losMod:3.8, leadMod:52 },
    'Thomas Cook':  { rnBase:185, adrBase:148, lyMult:0.79, slyMult:0.72, losMod:2.2, leadMod:28 },
    'Jet2holidays': { rnBase:310, adrBase:96,  lyMult:0.93, slyMult:0.87, losMod:3.1, leadMod:41 },
    'FTI Group':    { rnBase:265, adrBase:105, lyMult:0.88, slyMult:0.82, losMod:4.0, leadMod:35 },
    'Sunwing':      { rnBase:190, adrBase:122, lyMult:0.85, slyMult:0.78, losMod:5.2, leadMod:60 },
    'Club Med':     { rnBase:145, adrBase:195, lyMult:0.94, slyMult:0.89, losMod:6.5, leadMod:75 },
  };

  function toData(toName, origin, mealPlan) {
    const prof = TO_PROFILES[toName] || { rnBase:260, adrBase:100, lyMult:0.88, slyMult:0.82, losMod:3.0, leadMod:40 };
    const s = (origin||'X').charCodeAt(0) + (mealPlan||'A').charCodeAt(0);
    const rn  = prof.rnBase  + s * 3 % 80;
    const adr = prof.adrBase + s % 30;
    const rev = rn * adr;
    const ly  = prof.lyMult;
    const sly = prof.slyMult;
    const cs  = 0.52 + s % 22 * 0.01;
    const ds  = 0.33 + s % 18 * 0.01;
    return {
      rn, adr, rev,
      lyRn:Math.round(rn*ly),   lyAdr:Math.round(adr*ly),   lyRev:Math.round(rev*ly),
      stlyRn:Math.round(rn*sly),stlyAdr:Math.round(adr*sly),stlyRev:Math.round(rev*sly),
      revpar:Math.round(adr*0.72),lyRevpar:Math.round(adr*0.72*ly),stlyRevpar:Math.round(adr*0.72*sly),
      los:(prof.losMod + s%3*0.2).toFixed(1),
      avgAdults:(1.7+s%3*0.2).toFixed(1),
      avgChildren:(0.3+s%2*0.2).toFixed(1),
      avgLeadTime: prof.leadMod + s%15,
      contractShare:cs, promoShare:1-cs, dynamicShare:ds, staticShare:1-ds,
      projRn:Math.round(rn*1.08), projAdr:Math.round(adr*1.05), projRev:Math.round(rev*1.14),
    };
  }

  function getDateFactor() {
    const from=new Date(document.getElementById('insDateFrom')?.value||'2025-01-01');
    const to=new Date(document.getElementById('insDateTo')?.value||'2025-12-31');
    return Math.min(1,Math.max(0.01,(to-from)/86400000/365));
  }

  function getAggData(toName) {
    const meals  =insSelected.meal.length   ?insSelected.meal   :ALL_MEALS;
    const origins=insSelected.origin.length ?insSelected.origin :ALL_ORIGINS;
    const f=getDateFactor();
    let rn=0,rev=0,lyRn=0,lyRev=0,stlyRn=0,stlyRev=0,los=0,avgA=0,avgC=0,lead=0,cs=0,ds=0,cnt=0;
    origins.forEach(function(o){meals.forEach(function(m){
      const d=toData(toName,o,m);
      rn+=d.rn;rev+=d.rev;lyRn+=d.lyRn;lyRev+=d.lyRev;stlyRn+=d.stlyRn;stlyRev+=d.stlyRev;
      los+=parseFloat(d.los);avgA+=parseFloat(d.avgAdults);avgC+=parseFloat(d.avgChildren);
      lead+=d.avgLeadTime;cs+=d.contractShare;ds+=d.dynamicShare;cnt++;
    });});
    const n=cnt||1;
    rn=Math.round(rn*f);rev=Math.round(rev*f);lyRn=Math.round(lyRn*f);lyRev=Math.round(lyRev*f);
    stlyRn=Math.round(stlyRn*f);stlyRev=Math.round(stlyRev*f);
    const adr=Math.round(rev/Math.max(rn,1)),lyAdr=Math.round(lyRev/Math.max(lyRn,1)),stlyAdr=Math.round(stlyRev/Math.max(stlyRn,1));
    return {
      rn,rev,adr,lyRn,lyRev,lyAdr,stlyRn,stlyRev,stlyAdr,
      revpar:Math.round(adr*0.72),lyRevpar:Math.round(lyAdr*0.72),stlyRevpar:Math.round(stlyAdr*0.72),
      los:(los/n).toFixed(1),avgAdults:(avgA/n).toFixed(1),avgChildren:(avgC/n).toFixed(1),avgLeadTime:Math.round(lead/n),
      contractShare:cs/n,promoShare:1-cs/n,dynamicShare:ds/n,staticShare:1-ds/n,
      projRn:Math.round(rn*1.08),projAdr:Math.round(adr*1.05),projRev:Math.round(rev*1.14),
    };
  }

  // ── Format helpers ────────────────────────────────────────────
  function fmt(n){return n>=1000?'€'+(n/1000).toFixed(1)+'k':'€'+n;}
  function fmtN(n){return n.toLocaleString();}
  function fmtE(n){return '€'+n;}
  function delta(cur,ref){if(!ref)return'–';const p=Math.round((cur-ref)/ref*100);return(p>=0?'+':'')+p+'%';}
  function dCls(c,r){return c>=r?'ins-badge-pos':'ins-badge-neg';}
  function metricVal(d){return insMetric==='rn'?d.rn:d.rev;}
  function metricLY(d){return insMetric==='rn'?d.lyRn:d.lyRev;}
  function metricSTLY(d){return insMetric==='rn'?d.stlyRn:d.stlyRev;}
  function metricFmt(v){return insMetric==='rn'?fmtN(v):fmt(v);}
  function metricLabel(){return insMetric==='rn'?'Room Nights':'Revenue';}
  function chevron(open){return '<svg class="ins-coll-arrow'+(open?' open':'')+'" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" width="20" height="20"><polyline points="6 9 12 15 18 9"/></svg>';}

  // ── KPI card (LY/STLY stacked below value) ───────────────────
  function kpiCard(label,val,ly,stly,prefix){
    prefix=prefix||'';
    return '<div class="ins-kpi">'
      +'<div class="ins-kpi-label">'+label+'</div>'
      +'<div class="ins-kpi-val">'+prefix+val.toLocaleString()+'</div>'
      +'<div class="ins-kpi-sub">'
      +'<span class="ins-kpi-badge ins-badge-ly">LY '+delta(val,ly)+'</span>'
      +'<span class="ins-kpi-badge '+dCls(val,stly)+'">STLY '+delta(val,stly)+'</span>'
      +'</div></div>';
  }

  // ── Bar row ───────────────────────────────────────────────────
  function barRow(label,val,lyV,stlyV,maxV,color){
    const pct=Math.min(94,Math.round(val/maxV*100));
    const lyP=lyV!=null?Math.min(94,Math.round(lyV/maxV*100)):null;
    const slP=stlyV!=null?Math.min(94,Math.round(stlyV/maxV*100)):null;
    return '<div class="ins-bar-row">'
      +'<div class="ins-bar-lbl">'+label+'</div>'
      +'<div class="ins-bar-track">'
      +'<div class="ins-bar-fill" style="width:'+pct+'%;background:'+color+'"></div>'
      +(lyP!=null?'<div class="ins-bar-ref ins-bar-ref-ly" style="left:'+lyP+'%"></div>':'')
      +(slP!=null?'<div class="ins-bar-ref ins-bar-ref-stly" style="left:'+slP+'%"></div>':'')
      +'</div>'
      +'<div class="ins-bar-val">'+metricFmt(val)+'</div>'
      +'<div class="ins-bar-tags">'
      +(lyV!=null?'<span class="ins-bar-tag '+dCls(val,lyV)+'">LY '+delta(val,lyV)+'</span>':'')
      +(stlyV!=null?'<span class="ins-bar-tag '+dCls(val,stlyV)+'">STLY '+delta(val,stlyV)+'</span>':'')
      +'</div></div>';
  }

  // ── Pie chart SVG for contract/promo share ───────────────────
  function pieSvg(cShare, pShare) {
    const r=44, cx=54, cy=54, circ=2*Math.PI*r;
    const cLen=circ*cShare, pLen=circ*pShare;
    const cPct=Math.round(cShare*100), pPct=Math.round(pShare*100);
    // arc helper: sweep from startAngle for fraction
    function arcPath(startFrac, endFrac, fill) {
      const startA=(startFrac*2*Math.PI)-Math.PI/2;
      const endA  =(endFrac  *2*Math.PI)-Math.PI/2;
      const x1=cx+r*Math.cos(startA), y1=cy+r*Math.sin(startA);
      const x2=cx+r*Math.cos(endA),   y2=cy+r*Math.sin(endA);
      const large=endFrac-startFrac>0.5?1:0;
      return '<path d="M'+cx+','+cy+' L'+x1.toFixed(1)+','+y1.toFixed(1)+' A'+r+','+r+' 0 '+large+',1 '+x2.toFixed(1)+','+y2.toFixed(1)+' Z" fill="'+fill+'" opacity=".9"/>';
    }
    return '<div class="ins-donut-wrap">'
      +'<svg width="108" height="108" viewBox="0 0 108 108">'
      +arcPath(0, cShare, '#0ea5e9')
      +arcPath(cShare, 1, '#967EF3')
      +'<circle cx="'+cx+'" cy="'+cy+'" r="24" fill="var(--surface-1)"/>'
      +'<text x="'+cx+'" y="'+(cy-5)+'" text-anchor="middle" font-size="11" font-weight="700" fill="var(--text-primary)">'+cPct+'%</text>'
      +'<text x="'+cx+'" y="'+(cy+8)+'" text-anchor="middle" font-size="9" fill="var(--text-muted)">Contracts</text>'
      +'</svg>'
      +'<div class="ins-donut-legend">'
      +'<div class="ins-donut-leg-item"><div class="ins-donut-leg-dot" style="background:#0ea5e9"></div><span>'+cPct+'% From Contracts</span></div>'
      +'<div class="ins-donut-leg-item"><div class="ins-donut-leg-dot" style="background:#967EF3"></div><span>'+pPct+'% From Promotions</span></div>'
      +'</div></div>';
  }

  // ── Donut for channel share ───────────────────────────────────
  function donutSvg(dShare){
    const r=34,cx=42,cy=42,circ=2*Math.PI*r;
    const dLen=circ*dShare,cLen=circ-dLen;
    return '<svg width="84" height="84" viewBox="0 0 84 84">'
      +'<circle cx="'+cx+'" cy="'+cy+'" r="'+r+'" fill="none" stroke="var(--border-sub)" stroke-width="11"/>'
      +'<circle cx="'+cx+'" cy="'+cy+'" r="'+r+'" fill="none" stroke="#0ea5e9" stroke-width="11"'
      +' stroke-dasharray="'+dLen.toFixed(1)+' '+cLen.toFixed(1)+'" stroke-dashoffset="'+(circ*0.25).toFixed(1)+'" stroke-linecap="round"/>'
      +'<circle cx="'+cx+'" cy="'+cy+'" r="'+r+'" fill="none" stroke="#967EF3" stroke-width="11"'
      +' stroke-dasharray="'+cLen.toFixed(1)+' '+dLen.toFixed(1)+'" stroke-dashoffset="'+(circ*0.25-dLen).toFixed(1)+'" stroke-linecap="round" opacity=".85"/>'
      +'<text x="'+cx+'" y="'+(cy-3)+'" text-anchor="middle" font-size="11" font-weight="700" fill="var(--text-primary)">'+Math.round(dShare*100)+'%</text>'
      +'<text x="'+cx+'" y="'+(cy+10)+'" text-anchor="middle" font-size="9" fill="var(--text-muted)">Online</text>'
      +'</svg>';
  }

  // ── Collapsible ───────────────────────────────────────────────
  function collapsible(key,title,bodyHtml){
    const open=!!collOpen[key];
    return '<div class="ins-collapsible">'
      +'<div class="ins-coll-header" onclick="insToggleColl(\''+key+'\')">'
      +'<span>'+title+'</span>'+chevron(open)
      +'</div>'
      +'<div class="ins-coll-body" style="display:'+(open?'block':'none')+'">'+bodyHtml+'</div>'
      +'</div>';
  }

  // ── Stay details card only (no grid wrapper) ─────────────────
  function stayDetailsHtml(d){
    return '<div class="ins-detail-card">'
      +'<div class="ins-detail-card-title">Stay Details</div>'
      +'<div class="ins-detail-row"><span class="ins-detail-key">Avg LOS</span><span class="ins-detail-val">'+d.los+' nights</span></div>'
      +'<div class="ins-detail-row"><span class="ins-detail-key">Avg Adults</span><span class="ins-detail-val">'+d.avgAdults+'</span></div>'
      +'<div class="ins-detail-row"><span class="ins-detail-key">Avg Children</span><span class="ins-detail-val">'+d.avgChildren+'</span></div>'
      +'<div class="ins-detail-row"><span class="ins-detail-key">Avg Lead Time</span><span class="ins-detail-val">'+d.avgLeadTime+' days</span></div>'
      +'</div>';
  }

  // ── Multi-segment pie chart ───────────────────────────────────
  function multiPieSvg(segments){
    const r=44, cx=54, cy=54;
    function arcPath(startFrac, endFrac, fill){
      if(endFrac-startFrac>=1) endFrac=startFrac+0.9999;
      const startA=(startFrac*2*Math.PI)-Math.PI/2;
      const endA  =(endFrac  *2*Math.PI)-Math.PI/2;
      const x1=cx+r*Math.cos(startA), y1=cy+r*Math.sin(startA);
      const x2=cx+r*Math.cos(endA),   y2=cy+r*Math.sin(endA);
      const large=endFrac-startFrac>0.5?1:0;
      return '<path d="M'+cx+','+cy+' L'+x1.toFixed(1)+','+y1.toFixed(1)+' A'+r+','+r+' 0 '+large+',1 '+x2.toFixed(1)+','+y2.toFixed(1)+' Z" fill="'+fill+'" opacity=".9"/>';
    }
    var total=segments.reduce(function(s,sg){ return s+sg.pct; },0)||1;
    var start=0;
    var paths=segments.map(function(sg){
      var frac=sg.pct/total;
      var p=arcPath(start, start+frac, sg.color);
      start+=frac;
      return p;
    }).join('');
    var legend=segments.map(function(sg){
      return '<div class="ins-donut-leg-item"><div class="ins-donut-leg-dot" style="background:'+sg.color+'"></div><span>'+sg.pct+'% '+sg.name+'</span></div>';
    }).join('');
    return '<div class="ins-donut-wrap">'
      +'<svg width="108" height="108" viewBox="0 0 108 108">'
      +paths
      +'<circle cx="'+cx+'" cy="'+cy+'" r="20" fill="var(--surface-1)"/>'
      +'</svg>'
      +'<div class="ins-donut-legend">'+legend+'</div>'
      +'</div>';
  }

  // ── Room Type Mix pie card ────────────────────────────────────
  function roomTypeMixCard(toKey){
    const RT_NAMES=['Standard','Superior','Deluxe','Suite','Jr. Suite','Family'];
    const RT_COLORS=['#006461','#0891b2','#6366f1','#f59e0b','#ec4899','#10b981'];
    var segs=RT_NAMES.map(function(rtn,ri){
      const seed=(toKey.charCodeAt(0)||1)*(ri+1)*7;
      const pct=Math.max(5,Math.min(35,10+(seed*7+ri*13)%25));
      return {name:rtn, pct:pct, color:RT_COLORS[ri]};
    });
    return '<div class="ins-detail-card">'
      +'<div class="ins-detail-card-title">Room Type Mix</div>'
      +multiPieSvg(segs)
      +'</div>';
  }

  // ── Meal Plan Mix pie card ────────────────────────────────────
  function mealPlanMixCard(toKey){
    const MP_DEFS=[
      {name:'All Inclusive',color:'#004948'},
      {name:'Half Board',   color:'#C4FF45'},
      {name:'Bed & Breakfast',color:'#52d9ce'},
      {name:'Room Only',    color:'#D97706'},
      {name:'Full Board',   color:'#d7f7ed'},
    ];
    var segs=MP_DEFS.map(function(mp,mi){
      const seed2=(toKey.charCodeAt(1)||2)*(mi+1)*3;
      const pct=Math.max(4,Math.min(40,8+(seed2*11+mi*17)%32));
      return {name:mp.name, pct:pct, color:mp.color};
    });
    return '<div class="ins-detail-card">'
      +'<div class="ins-detail-card-title">Meal Plan Mix</div>'
      +multiPieSvg(segs)
      +'</div>';
  }

  // ── Promotions deep dive ──────────────────────────────────────
  function promosHtml(toName,d){
    return INS_PROMOS.map(function(p,i){
      const pRn=Math.round(d.rn*d.promoShare*(0.22+i*0.09));
      const pRev=Math.round(d.rev*d.promoShare*(0.22+i*0.09));
      const pAdr=Math.round(d.adr*(1-p.discount/100));
      const key='promo_'+toName+'_'+p.id;
      const open=promoOpen[key];
      return '<div class="ins-promo-accordion">'
        +'<div class="ins-promo-header" onclick="insTogglePromo(\''+key+'\')">'
        +'<span>'+p.name+' <span style="font-size:10px;font-weight:500;color:var(--text-muted)">−'+p.discount+'%</span></span>'
        +'<svg class="'+(open?'ins-coll-arrow open':'ins-coll-arrow')+'" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" width="20" height="20"><polyline points="6 9 12 15 18 9"/></svg>'
        +'</div>'
        +'<div class="ins-promo-body" style="display:'+(open?'block':'none')+'">'
        +'<div class="ins-promo-kpi-row">'
        +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Room Nights</div><div class="ins-promo-kpi-val">'+fmtN(pRn)+'</div></div>'
        +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Revenue</div><div class="ins-promo-kpi-val">'+fmt(pRev)+'</div></div>'
        +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">ADR</div><div class="ins-promo-kpi-val">'+fmtE(pAdr)+'</div></div>'
        +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Adults</div><div class="ins-promo-kpi-val">'+Math.round(pRn*parseFloat(d.avgAdults)).toLocaleString()+'</div></div>'
        +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Children</div><div class="ins-promo-kpi-val">'+Math.round(pRn*parseFloat(d.avgChildren)).toLocaleString()+'</div></div>'
        +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">RevPAR</div><div class="ins-promo-kpi-val">'+fmtE(Math.round(pAdr*0.72))+'</div></div>'
        +'</div></div></div>';
    }).join('');
  }

  // ── TO section ────────────────────────────────────────────────
  function renderTOSection(t,d){
    const toKey = t.id;
    // ── Contracts for this TO with LY/STLY ───────────────────────
    const toContracts = INS_CONTRACTS.filter(function(c){ return (c.tos||[c.to]).indexOf(toKey)>=0; });
    const contractsHtml = toContracts.length > 0
      ? '<div class="ins-section-title" style="margin-top:10px">Active Contracts</div>'
        + toContracts.map(function(c){
            const sM = c.season==='Summer'?1.12:c.season==='Winter'?0.88:c.season==='Q1'?0.82:1.0;
            const cRn=Math.round(d.rn*sM*0.45), cRev=Math.round(d.rev*sM*0.45);
            const cAdr=Math.round(d.adr*(0.96+(c.id.charCodeAt(c.id.length-1)%8)*0.005));
            const cLyRn=Math.round(cRn*0.9), cStlyRn=Math.round(cRn*0.85);
            const cLyRev=Math.round(cRev*0.9), cStlyRev=Math.round(cRev*0.85);
            const relDate=c.releaseDate?new Date(c.releaseDate):null;
            const TODAY_INS=new Date(2026,2,9);
            const dLeft=relDate?Math.round((relDate-TODAY_INS)/86400000):null;
            const relBadge=dLeft===null?'':(dLeft>0
              ?'<span style="font-size:9px;font-weight:700;color:#f59e0b;margin-left:6px">⏳'+dLeft+'d to release</span>'
              :'<span style="font-size:9px;font-weight:700;color:#16a34a;margin-left:6px">✓ Released</span>');
            return '<div style="margin-bottom:8px;padding:10px 12px;border-radius:8px;border:1px solid '+t.color+'33;background:var(--surface-2)">'
              +'<div style="display:flex;align-items:center;gap:6px;margin-bottom:8px">'
              +'<span style="font-size:12px;font-weight:700;color:'+t.color+'">'+c.name+'</span>'
              +'<span style="font-size:10px;background:#eff6ff;color:#1d4ed8;padding:2px 6px;border-radius:4px">'+c.type+'</span>'
              +relBadge
              +'</div>'
              +'<div style="display:flex;gap:10px;flex-wrap:wrap">'
              +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Room Nights</div><div class="ins-promo-kpi-val">'+fmtN(cRn)+'</div></div>'
              +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">LY RN</div><div class="ins-promo-kpi-val" style="color:#6b7280">'+fmtN(cLyRn)+'</div></div>'
              +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">STLY RN</div><div class="ins-promo-kpi-val" style="color:#6b7280">'+fmtN(cStlyRn)+'</div></div>'
              +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Revenue</div><div class="ins-promo-kpi-val">'+fmt(cRev)+'</div></div>'
              +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">LY Rev</div><div class="ins-promo-kpi-val" style="color:#6b7280">'+fmt(cLyRev)+'</div></div>'
              +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">ADR</div><div class="ins-promo-kpi-val">'+fmtE(cAdr)+'</div></div>'
              +'</div></div>';
          }).join('')
      : '';

    // ── Room Type breakdown ────────────────────────────────────────
    const RT_NAMES=['Standard','Superior','Deluxe','Suite','Jr. Suite','Family'];
    const RT_COLORS=['#006461','#0891b2','#6366f1','#f59e0b','#ec4899','#10b981'];
    const rtRows=RT_NAMES.map(function(rtn,ri){
      const seed=(toKey.charCodeAt(0)||1)*(ri+1)*7;
      const rtRnPct=Math.max(5,Math.min(35,10+(seed*7+ri*13)%25));
      const rtRn=Math.round(d.rn*rtRnPct/100);
      const rtAdr=Math.round(d.adr*(0.8+ri*0.08+(seed%10)*0.01));
      const rtRev=rtRn*rtAdr, rtRevpar=Math.round(rtAdr*0.72);
      const rtLos=(parseFloat(d.los)*(0.9+ri*0.05)).toFixed(1);
      const rtAvgA=(parseFloat(d.avgAdults)*(0.95+ri*0.03)).toFixed(1);
      const rtAvgC=(parseFloat(d.avgChildren)*(0.9+ri*0.04)).toFixed(1);
      const rtGuests=Math.round(rtRn*(parseFloat(rtAvgA)+parseFloat(rtAvgC)));
      return '<div style="padding:8px 10px;border-radius:7px;border:1px solid '+RT_COLORS[ri]+'33;background:var(--surface-2);margin-bottom:6px">'
        +'<div style="display:flex;align-items:center;gap:6px;margin-bottom:7px">'
        +'<span style="width:8px;height:8px;border-radius:50%;background:'+RT_COLORS[ri]+';flex-shrink:0"></span>'
        +'<span style="font-size:12px;font-weight:700;color:'+RT_COLORS[ri]+'">'+rtn+'</span>'
        +'<span style="margin-left:auto;font-size:10px;font-weight:700;padding:1px 6px;border-radius:8px;background:'+RT_COLORS[ri]+'20;color:'+RT_COLORS[ri]+'">'+rtRnPct+'% mix</span>'
        +'</div>'
        +'<div style="display:flex;gap:8px;flex-wrap:wrap">'
        +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Room Nights</div><div class="ins-promo-kpi-val">'+fmtN(rtRn)+'</div></div>'
        +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Revenue</div><div class="ins-promo-kpi-val">'+fmt(rtRev)+'</div></div>'
        +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">ADR</div><div class="ins-promo-kpi-val">'+fmtE(rtAdr)+'</div></div>'
        +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">REVPAR</div><div class="ins-promo-kpi-val">'+fmtE(rtRevpar)+'</div></div>'
        +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Avg LOS</div><div class="ins-promo-kpi-val">'+rtLos+' n</div></div>'
        +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Avg Adults</div><div class="ins-promo-kpi-val">'+rtAvgA+'</div></div>'
        +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Avg Children</div><div class="ins-promo-kpi-val">'+rtAvgC+'</div></div>'
        +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Total Guests</div><div class="ins-promo-kpi-val">'+fmtN(rtGuests)+'</div></div>'
        +'</div></div>';
    }).join('');

    // ── Meal Plan breakdown ────────────────────────────────────────
    const MP_DEFS=[
      {key:'AI',name:'All Inclusive',color:'#004948'},
      {key:'HB',name:'Half Board',   color:'#C4FF45'},
      {key:'BB',name:'Bed & Breakfast',color:'#52d9ce'},
      {key:'RO',name:'Room Only',    color:'#D97706'},
      {key:'FB',name:'Full Board',   color:'#d7f7ed'},
    ];
    const mpRows=MP_DEFS.map(function(mp,mi){
      const seed2=(toKey.charCodeAt(1)||2)*(mi+1)*3;
      const mpPct=Math.max(4,Math.min(40,8+(seed2*11+mi*17)%32));
      const mpRn=Math.round(d.rn*mpPct/100);
      const mpAdrNet=Math.round(d.adr*(0.78+mi*0.06+(seed2%8)*0.01));
      const mpAdrGross=Math.round(mpAdrNet*(1.12+mi*0.02));
      const mpRev=mpRn*mpAdrGross, mpRevpar=Math.round(mpAdrGross*0.72);
      const mpLos=(parseFloat(d.los)*(0.88+mi*0.07)).toFixed(1);
      const mpAvgA=(parseFloat(d.avgAdults)*(0.93+mi*0.04)).toFixed(1);
      const mpAvgC=(parseFloat(d.avgChildren)*(0.85+mi*0.06)).toFixed(1);
      const mpAvgG=(parseFloat(mpAvgA)+parseFloat(mpAvgC)).toFixed(1);
      const mpTotG=Math.round(mpRn*parseFloat(mpAvgG));
      return '<div style="padding:8px 10px;border-radius:7px;border:1px solid '+mp.color+'33;background:var(--surface-2);margin-bottom:6px">'
        +'<div style="display:flex;align-items:center;gap:6px;margin-bottom:7px">'
        +'<span style="width:8px;height:8px;border-radius:50%;background:'+mp.color+';flex-shrink:0"></span>'
        +'<span style="font-size:12px;font-weight:700;color:'+mp.color+'">'+mp.name+'</span>'
        +'<span style="margin-left:auto;font-size:10px;font-weight:700;padding:1px 6px;border-radius:8px;background:'+mp.color+'20;color:'+mp.color+'">'+mpPct+'% mix</span>'
        +'</div>'
        +'<div style="display:flex;gap:8px;flex-wrap:wrap">'
        +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Room Nights</div><div class="ins-promo-kpi-val">'+fmtN(mpRn)+'</div></div>'
        +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Revenue</div><div class="ins-promo-kpi-val">'+fmt(mpRev)+'</div></div>'
        +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">ADR</div><div class="ins-promo-kpi-val">'+fmtE(mpAdrNet)+'</div></div>'
        +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">ADR Net</div><div class="ins-promo-kpi-val">'+fmtE(mpAdrNet)+'</div></div>'
        +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">ADR Gross</div><div class="ins-promo-kpi-val">'+fmtE(mpAdrGross)+'</div></div>'
        +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">REVPAR</div><div class="ins-promo-kpi-val">'+fmtE(mpRevpar)+'</div></div>'
        +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Avg LOS</div><div class="ins-promo-kpi-val">'+mpLos+' n</div></div>'
        +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Avg Adults</div><div class="ins-promo-kpi-val">'+mpAvgA+'</div></div>'
        +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Avg Children</div><div class="ins-promo-kpi-val">'+mpAvgC+'</div></div>'
        +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Avg Guest</div><div class="ins-promo-kpi-val">'+mpAvgG+'</div></div>'
        +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Total Guests</div><div class="ins-promo-kpi-val">'+fmtN(mpTotG)+'</div></div>'
        +'</div></div>';
    }).join('');

    // ── Source Geo ─────────────────────────────────────────────────
    const GEO_DEFS=[
      {key:'UK',name:'United Kingdom',color:'#006461'},
      {key:'SP',name:'Spain',         color:'#0891b2'},
      {key:'US',name:'United States', color:'#6366f1'},
      {key:'MX',name:'Mexico',        color:'#f59e0b'},
      {key:'DE',name:'Germany',       color:'#ec4899'},
      {key:'FR',name:'France',        color:'#10b981'},
    ];
    const geoSeed=(toKey.charCodeAt(2)||1);
    var geoTotal=0;
    const geoPcts=GEO_DEFS.map(function(g,gi){ const p=Math.max(3,Math.min(40,6+(geoSeed*(gi+2)+gi*11)%34)); geoTotal+=p; return p; });
    const geoRows=GEO_DEFS.map(function(g,gi){
      const pct=Math.round(geoPcts[gi]/geoTotal*100);
      return '<div style="margin-bottom:6px">'
        +'<div style="display:flex;align-items:center;gap:6px;margin-bottom:2px">'
        +'<span style="width:7px;height:7px;border-radius:50%;background:'+g.color+';flex-shrink:0"></span>'
        +'<span style="font-size:11px;color:var(--text-primary);flex:1">'+g.name+'</span>'
        +'<span style="font-size:11px;font-weight:700;color:'+g.color+'">'+pct+'%</span>'
        +'</div>'
        +'<div style="height:5px;background:#e5e7eb;border-radius:3px;overflow:hidden">'
        +'<div style="height:100%;width:'+Math.max(2,pct)+'%;background:'+g.color+';border-radius:3px"></div>'
        +'</div></div>';
    }).join('');

    // ── Channel Share — now includes Tour Series ───────────────────
    const tourSeriesPct = Math.round(d.staticShare * 100 * 0.35);
    const dynamicPct    = Math.round(d.dynamicShare * 100);
    const staticFITPct  = Math.round(d.staticShare * 100 * 0.65);
    const CHANNELS = [
      {name:'Dynamic / Online', pct:dynamicPct,   color:'#0ea5e9'},
      {name:'Static FIT',       pct:staticFITPct, color:'#967EF3'},
      {name:'Tour Series',      pct:tourSeriesPct,color:'#ec4899'},
    ];
    const chDetailRows = CHANNELS.map(function(ch){
      const chRn  = Math.round(d.rn  * ch.pct / 100);
      const chRev = Math.round(d.rev * ch.pct / 100);
      const chAdr = Math.round(d.adr * (0.9 + ch.pct * 0.002));
      const chRevpar = Math.round(chAdr * 0.72);
      const chLos = (parseFloat(d.los) * (0.9 + ch.pct * 0.003)).toFixed(1);
      const chAvgA = (parseFloat(d.avgAdults) * 0.97).toFixed(1);
      const chAvgC = (parseFloat(d.avgChildren) * 0.95).toFixed(1);
      const chGuests = Math.round(chRn * (parseFloat(chAvgA)+parseFloat(chAvgC)));
      return '<div style="padding:8px 10px;border-radius:7px;border:1px solid '+ch.color+'33;background:var(--surface-2);margin-bottom:6px">'
        +'<div style="display:flex;align-items:center;gap:6px;margin-bottom:7px">'
        +'<span style="width:8px;height:8px;border-radius:50%;background:'+ch.color+';flex-shrink:0"></span>'
        +'<span style="font-size:12px;font-weight:700;color:'+ch.color+'">'+ch.name+'</span>'
        +'<span style="margin-left:auto;font-size:10px;font-weight:700;padding:1px 6px;border-radius:8px;background:'+ch.color+'20;color:'+ch.color+'">'+ch.pct+'% share</span>'
        +'</div>'
        +'<div style="display:flex;gap:8px;flex-wrap:wrap">'
        +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Room Nights</div><div class="ins-promo-kpi-val">'+fmtN(chRn)+'</div></div>'
        +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Revenue</div><div class="ins-promo-kpi-val">'+fmt(chRev)+'</div></div>'
        +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">ADR</div><div class="ins-promo-kpi-val">'+fmtE(chAdr)+'</div></div>'
        +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">REVPAR</div><div class="ins-promo-kpi-val">'+fmtE(chRevpar)+'</div></div>'
        +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Avg LOS</div><div class="ins-promo-kpi-val">'+chLos+' n</div></div>'
        +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Avg Adults</div><div class="ins-promo-kpi-val">'+chAvgA+'</div></div>'
        +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Avg Children</div><div class="ins-promo-kpi-val">'+chAvgC+'</div></div>'
        +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Total Guests</div><div class="ins-promo-kpi-val">'+fmtN(chGuests)+'</div></div>'
        +'</div></div>';
    }).join('');

    const channelHtml =
      '<div class="ins-detail-card">'
      +'<div class="ins-detail-card-title">Channel Share</div>'
      +'<div class="ins-donut-wrap">'
      +donutSvg(d.dynamicShare)
      +'<div class="ins-donut-legend">'
      +'<div class="ins-donut-leg-item"><div class="ins-donut-leg-dot" style="background:#0ea5e9"></div><span>'+dynamicPct+'% Dynamic / Online</span></div>'
      +'<div class="ins-donut-leg-item"><div class="ins-donut-leg-dot" style="background:#967EF3"></div><span>'+staticFITPct+'% Static FIT</span></div>'
      +'<div class="ins-donut-leg-item"><div class="ins-donut-leg-dot" style="background:#ec4899"></div><span>'+tourSeriesPct+'% Tour Series</span></div>'
      +'</div></div>'
      +'</div>'
      +collapsible('ch_detail_'+toKey,'Channel Stay Details',chDetailRows);

    const moreBody =
      '<div class="ins-detail-grid">'+stayDetailsHtml(d)+channelHtml+'</div>'
      +'<div class="ins-detail-grid">'+roomTypeMixCard(toKey)+mealPlanMixCard(toKey)+'</div>'
      +collapsible('rt_to_'+toKey, 'Room Type Breakdown',
          '<div class="ins-detail-card" style="margin-top:0">'+rtRows+'</div>')
      +collapsible('mp_to_'+toKey, 'Meal Plan Breakdown',
          '<div class="ins-detail-card" style="margin-top:0">'+mpRows+'</div>')
      +collapsible('geo_to_'+toKey,'Source Geo',
          '<div class="ins-detail-card" style="margin-top:0">'+geoRows+'</div>')
      +'<div class="ins-section-title" style="margin-top:12px">Contracts &amp; Promotions Share</div>'
      +pieSvg(d.contractShare,d.promoShare)
      +contractsHtml
      +'<div class="ins-section-title" style="margin-top:12px">Projected <span class="ins-projected-badge">📋 Contract</span></div>'
      +'<div class="ins-kpi-row">'
      +'<div class="ins-kpi"><div class="ins-kpi-label">Proj. Room Nights</div><div class="ins-kpi-val">'+fmtN(d.projRn)+'</div></div>'
      +'<div class="ins-kpi"><div class="ins-kpi-label">Proj. Revenue</div><div class="ins-kpi-val">€'+(d.projRev/1000).toFixed(1)+'k</div></div>'
      +'<div class="ins-kpi"><div class="ins-kpi-label">Proj. ADR</div><div class="ins-kpi-val">€'+d.projAdr+'</div></div>'
      +'<div class="ins-kpi"><div class="ins-kpi-label">Proj. RevPAR</div><div class="ins-kpi-val">€'+Math.round(d.projAdr*0.72)+'</div></div>'
      +'</div>'
      +'<div class="ins-section-title" style="margin-top:12px">Promotion Deep Dive</div>'
      +promosHtml(t.name,d);

    return '<div class="ins-section" style="border-left:3px solid '+t.color+';padding-left:12px">'
      +'<div style="display:flex;align-items:center;gap:8px;margin-bottom:14px">'
      +'<div style="width:10px;height:10px;border-radius:50%;background:'+t.color+'"></div>'
      +'<span style="font-size:15px;font-weight:700;color:var(--text-primary)">'+t.name+'</span>'
      +'</div>'
      +'<div class="ins-section-title">Actuals vs LY &amp; STLY</div>'
      +'<div class="ins-kpi-row">'
      +kpiCard('Room Nights',d.rn,d.lyRn,d.stlyRn)
      +kpiCard('Revenue',d.rev,d.lyRev,d.stlyRev,'€')
      +kpiCard('ADR',d.adr,d.lyAdr,d.stlyAdr,'€')
      +kpiCard('RevPAR',d.revpar,d.lyRevpar,d.stlyRevpar,'€')
      +kpiCard('Total Adults',Math.round(d.rn*parseFloat(d.avgAdults)),Math.round(d.lyRn*parseFloat(d.avgAdults)),Math.round(d.stlyRn*parseFloat(d.avgAdults)))
      +kpiCard('Total Children',Math.round(d.rn*parseFloat(d.avgChildren)),Math.round(d.lyRn*parseFloat(d.avgChildren)),Math.round(d.stlyRn*parseFloat(d.avgChildren)))
      +'</div>'
      +collapsible('details_'+t.id,'More Details',moreBody)
      +'</div>';
  }

  // ── Side-by-side compare table ────────────────────────────────
  function renderCompare(t1,d1,t2,d2){
    const rows=[
      {label:metricLabel(),v1:metricVal(d1),v2:metricVal(d2),fmt:metricFmt,ly1:metricLY(d1),ly2:metricLY(d2),stly1:metricSTLY(d1),stly2:metricSTLY(d2)},
      {label:'ADR',v1:d1.adr,v2:d2.adr,fmt:fmtE,ly1:d1.lyAdr,ly2:d2.lyAdr,stly1:d1.stlyAdr,stly2:d2.stlyAdr},
      {label:'RevPAR',v1:d1.revpar,v2:d2.revpar,fmt:fmtE,ly1:d1.lyRevpar,ly2:d2.lyRevpar,stly1:d1.stlyRevpar,stly2:d2.stlyRevpar},
      {label:'Avg LOS',v1:parseFloat(d1.los),v2:parseFloat(d2.los),fmt:function(v){return v.toFixed(1)+' nights';},ly1:null,ly2:null,stly1:null,stly2:null},
      {label:'Lead Time',v1:d1.avgLeadTime,v2:d2.avgLeadTime,fmt:function(v){return v+' days';},ly1:null,ly2:null,stly1:null,stly2:null},
      {label:'Online %',v1:Math.round(d1.dynamicShare*100),v2:Math.round(d2.dynamicShare*100),fmt:function(v){return v+'%';},ly1:null,ly2:null,stly1:null,stly2:null},
      {label:'Contracts %',v1:Math.round(d1.contractShare*100),v2:Math.round(d2.contractShare*100),fmt:function(v){return v+'%';},ly1:null,ly2:null,stly1:null,stly2:null},
    ];

    let html='<div class="ins-section"><div class="ins-section-title">Side-by-Side Comparison</div>'
      +'<div class="ins-chart-wrap" style="padding:0;overflow:hidden">'
      +'<table class="ins-cmp-table"><thead><tr>'
      +'<th style="width:120px">Metric</th>'
      +'<th class="ins-cmp-metric"><span class="ins-cmp-to-dot" style="background:'+t1.color+'"></span>'+t1.name+'</th>'
      +'<th class="ins-cmp-metric"><span class="ins-cmp-to-dot" style="background:'+t2.color+'"></span>'+t2.name+'</th>'
      +'</tr></thead><tbody>';

    rows.forEach(function(r){
      const w=r.v1>=r.v2?'t1':'t2';
      html+='<tr>'
        +'<td class="ins-cmp-row-label">'+r.label+'</td>'
        +'<td class="ins-cmp-val" style="'+(w==='t1'?'background:rgba(14,165,233,.08)':'')+'">'
        +r.fmt(r.v1)
        +(r.ly1!=null?'<div style="display:flex;gap:4px;justify-content:center;margin-top:4px"><span class="ins-kpi-badge ins-badge-ly">LY '+delta(r.v1,r.ly1)+'</span><span class="ins-kpi-badge '+dCls(r.v1,r.stly1)+'">STLY '+delta(r.v1,r.stly1)+'</span></div>':'')
        +'</td>'
        +'<td class="ins-cmp-val" style="'+(w==='t2'?'background:rgba(139,92,246,.08)':'')+'">'
        +r.fmt(r.v2)
        +(r.ly2!=null?'<div style="display:flex;gap:4px;justify-content:center;margin-top:4px"><span class="ins-kpi-badge ins-badge-ly">LY '+delta(r.v2,r.ly2)+'</span><span class="ins-kpi-badge '+dCls(r.v2,r.stly2)+'">STLY '+delta(r.v2,r.stly2)+'</span></div>':'')
        +'</td></tr>';
    });

    html+='</tbody></table></div>'
      +'<div style="font-size:10px;color:var(--text-muted);margin-top:6px;padding:0 4px">'
      +'<span style="display:inline-flex;align-items:center;gap:4px"><span style="width:12px;height:8px;border-radius:2px;background:rgba(14,165,233,.15);display:inline-block"></span>Leading value highlighted</span>'
      +'</div></div>';

    // Bar chart below
    const maxV=Math.max(metricVal(d1),metricVal(d2),metricLY(d1),metricLY(d2),1);
    html+='<div class="ins-section"><div class="ins-chart-wrap"><div class="ins-chart-title">'+metricLabel()+' Visual Comparison</div>'
      +barRow(t1.name,metricVal(d1),metricLY(d1),metricSTLY(d1),maxV,t1.color)
      +barRow(t2.name,metricVal(d2),metricLY(d2),metricSTLY(d2),maxV,t2.color)
      +'<div style="display:flex;gap:12px;margin-top:6px;font-size:10px;color:var(--text-muted)">'
      +'<span style="display:flex;align-items:center;gap:4px"><span style="width:14px;height:2px;background:#1d4ed8;display:inline-block"></span>LY</span>'
      +'<span style="display:flex;align-items:center;gap:4px"><span style="width:14px;height:2px;background:#15803d;display:inline-block"></span>STLY</span>'
      +'</div></div></div>';

    return html;
  }

  // ── Main render ───────────────────────────────────────────────
  function insRender(){
    const container=document.getElementById('insContent');
    if(!container) return;
    try {
    let html='';

    if(insView==='to'){
      const toList=insSelected.to.length
        ?INS_TOS.filter(function(t){return insSelected.to.indexOf(t.id)>=0;})
        :INS_TOS;

      // ── Top N bar chart ─────────────────────────────────────────
      function topMetricVal(d) {
        if(insTopMetric==='rev')       return d.rev;
        if(insTopMetric==='los')       return parseFloat(d.los);
        if(insTopMetric==='avgAdults') return parseFloat(d.avgAdults);
        if(insTopMetric==='leadTime')  return d.avgLeadTime;
        if(insTopMetric==='srcGeo')    return Math.round(d.dynamicShare*100);
        return d.rn;
      }
      function topMetricFmt(v) {
        if(insTopMetric==='rev')       return fmt(v);
        if(insTopMetric==='los')       return v.toFixed(1)+' n';
        if(insTopMetric==='avgAdults') return v.toFixed(1);
        if(insTopMetric==='leadTime')  return v+' d';
        if(insTopMetric==='srcGeo')    return v+'%';
        return fmtN(Math.round(v));
      }
      function topMetricLY(d)   { return (insTopMetric==='rev'?d.lyRev:(insTopMetric==='rn'?d.lyRn:null)); }
      function topMetricSTLY(d) { return (insTopMetric==='rev'?d.stlyRev:(insTopMetric==='rn'?d.stlyRn:null)); }

      const sortedTOs = toList.slice().sort(function(a,b){ return topMetricVal(getAggData(b.id))-topMetricVal(getAggData(a.id)); });
      const topN    = Math.min(insTopN, sortedTOs.length);
      const topTOs  = sortedTOs.slice(0, topN);
      const restTOs = sortedTOs.slice(topN);
      const maxV    = Math.max.apply(null, topTOs.map(function(t){ return topMetricVal(getAggData(t.id)); }).concat([1]));
      const allDataSum  = sortedTOs.reduce(function(acc,t){ return acc+topMetricVal(getAggData(t.id)); },0)||1;
      const restOpsSum  = restTOs.reduce(function(acc,t){ return acc+topMetricVal(getAggData(t.id)); },0);
      const hotelOtherV = Math.round(allDataSum * 0.30);

      const nOpts = [3,5,7,10,15];
      const tmOpts = [{v:'rn',l:'RN'},{v:'rev',l:'Revenue'},{v:'los',l:'LOS'},{v:'avgAdults',l:'Avg Adults'},{v:'leadTime',l:'Lead Time'},{v:'srcGeo',l:'Src Geo'}];

      html+='<div class="ins-section">'
        +'<div class="ins-chart-wrap">'
        +'<div style="display:flex;align-items:center;flex-wrap:wrap;gap:8px;margin-bottom:14px">'
        +'<div class="ins-chart-title" style="margin:0">Top '+topN+' Operators</div>'
        +'<div style="margin-left:auto;display:flex;gap:6px;flex-wrap:wrap;align-items:center">'
        +'<span style="font-size:10px;color:var(--text-muted)">Show:</span>'
        +'<div style="display:flex;gap:1px;background:#f3f4f6;border-radius:6px;padding:2px">'
        +nOpts.map(function(n){ return '<button class="rev-gran-btn'+(insTopN===n?' active':'')+'" onclick="insSetTopN('+n+')">'+n+'</button>'; }).join('')
        +'</div>'
        +'<span style="font-size:10px;color:var(--text-muted);margin-left:4px">By:</span>'
        +'<div style="display:flex;gap:1px;background:#f3f4f6;border-radius:6px;padding:2px">'
        +tmOpts.map(function(o){ return '<button class="rev-gran-btn'+(insTopMetric===o.v?' active':'')+'" onclick="insSetTopMetric(&apos;'+o.v+'&apos;)" style="font-size:9px;padding:3px 7px">'+o.l+'</button>'; }).join('')
        +'</div></div></div>'
        +topTOs.map(function(t){ const d=getAggData(t.id); return barRow(t.name,topMetricVal(d),topMetricLY(d),topMetricSTLY(d),maxV,t.color); }).join('')
        +(restTOs.length>0
          ? '<div style="margin-top:6px;padding:5px 8px;border-radius:6px;background:#f8fafc;border:1px dashed #d1d5db">'
            +'<div style="display:flex;align-items:center;gap:6px">'
            +'<span style="font-size:11px;color:var(--text-muted);flex:1">Other operators ('+restTOs.length+')</span>'
            +'<span style="font-size:11px;font-weight:700;color:#6b7280">'+topMetricFmt(restOpsSum)+'</span>'
            +'<span style="font-size:10px;color:#9ca3af;margin-left:4px">('+Math.round(restOpsSum/allDataSum*100)+'%)</span>'
            +'</div></div>'
          : '')
        +'<div style="margin-top:4px;padding:5px 8px;border-radius:6px;background:#f1f5f9;border:1px dashed #cbd5e1">'
        +'<div style="display:flex;align-items:center;gap:6px">'
        +'<span style="font-size:11px;color:var(--text-muted);flex:1">Hotel – other segments (non-operator)</span>'
        +'<span style="font-size:11px;font-weight:700;color:#475569">'+topMetricFmt(hotelOtherV)+'</span>'
        +'<span style="font-size:10px;color:#9ca3af;margin-left:4px">('+Math.round(hotelOtherV/(allDataSum+hotelOtherV)*100)+'%)</span>'
        +'</div></div>'
        +'<div style="display:flex;gap:12px;margin-top:8px;font-size:10px;color:var(--text-muted)">'
        +'<span style="display:flex;align-items:center;gap:4px"><span style="width:14px;height:2px;background:#1d4ed8;display:inline-block"></span>LY</span>'
        +'<span style="display:flex;align-items:center;gap:4px"><span style="width:14px;height:2px;background:#15803d;display:inline-block"></span>STLY</span>'
        +'</div></div></div>';

      // Compare
      if(insCompare&&toList.length>=1){
        const cmpId=document.getElementById('insCmpTO')?.value;
        const t1=toList[0],t2=INS_TOS.find(function(t){return t.id===cmpId;})||INS_TOS[1];
        if(t1&&t2&&t1.id!==t2.id) html+=renderCompare(t1,getAggData(t1.id),t2,getAggData(t2.id));
      }

      // Per-TO sections
      toList.forEach(function(t){html+=renderTOSection(t,getAggData(t.id));});

    } else if(insView==='contract') {
      // Contract view — no metric selector, no channel/source
      const contractList=insSelected.contract.length
        ?INS_CONTRACTS.filter(function(c){return insSelected.contract.indexOf(c.id)>=0;})
        :INS_CONTRACTS;

      contractList.forEach(function(c){
        // Use the first TO's color for the contract border
        const primaryTO = c.tos ? c.tos[0] : (c.to||'');
        const t = INS_TOS.find(function(x){return x.id===primaryTO;})||{color:'#6b7280',name:primaryTO};
        // Aggregate data across all TOs on this contract
        const allTONames = c.tos || [c.to];
        const meals  = insSelected.meal.length   ? insSelected.meal   : ALL_MEALS;
        const origins= insSelected.origin.length ? insSelected.origin : ALL_ORIGINS;
        const f = getDateFactor();
        // Combined totals across all TOs
        let aggRn=0,aggRev=0,aggLyRn=0,aggLyRev=0,aggStlyRn=0,aggStlyRev=0;
        allTONames.forEach(function(tn){
          origins.forEach(function(o){ meals.forEach(function(m){
            const d=toData(tn,o,m);
            aggRn+=d.rn; aggRev+=d.rev;
            aggLyRn+=d.lyRn; aggLyRev+=d.lyRev;
            aggStlyRn+=d.stlyRn; aggStlyRev+=d.stlyRev;
          }); });
        });
        aggRn=Math.round(aggRn*f); aggRev=Math.round(aggRev*f);
        aggLyRn=Math.round(aggLyRn*f); aggLyRev=Math.round(aggLyRev*f);
        aggStlyRn=Math.round(aggStlyRn*f); aggStlyRev=Math.round(aggStlyRev*f);
        const aggAdr=Math.round(aggRev/Math.max(aggRn,1));
        const aggLyAdr=Math.round(aggLyRev/Math.max(aggLyRn,1));
        const aggStlyAdr=Math.round(aggStlyRev/Math.max(aggStlyRn,1));
        const aggRevpar=Math.round(aggAdr*0.72);

        // TO tags row
        const toTagsHtml = allTONames.map(function(tn){
          const tc=INS_TOS.find(function(x){return x.id===tn;})||{color:'#6b7280'};
          return '<span style="display:inline-flex;align-items:center;gap:4px;padding:2px 8px;border-radius:12px;background:'+tc.color+'18;border:1px solid '+tc.color+'44;font-size:11px;font-weight:600;color:'+tc.color+'">'+tn+'</span>';
        }).join(' ');

        // Per-TO breakdown for contracts with 2 TOs
        const perTOHtml = allTONames.map(function(tn){
          const d2=getAggData(tn);
          const tc=INS_TOS.find(function(x){return x.id===tn;})||{color:'#6b7280'};
          return '<div style="margin-bottom:8px;padding:10px 12px;border-radius:8px;border:1px solid '+tc.color+'33;background:var(--surface-2)">'
            +'<div style="font-size:12px;font-weight:700;color:'+tc.color+';margin-bottom:8px">'+tn+'</div>'
            +'<div style="display:flex;gap:10px;flex-wrap:wrap">'
            +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Room Nights</div><div class="ins-promo-kpi-val">'+fmtN(d2.rn)+'</div></div>'
            +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Revenue</div><div class="ins-promo-kpi-val">'+fmt(d2.rev)+'</div></div>'
            +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">ADR</div><div class="ins-promo-kpi-val">'+fmtE(d2.adr)+'</div></div>'
            +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">RevPAR</div><div class="ins-promo-kpi-val">'+fmtE(d2.revpar)+'</div></div>'
            +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Avg LOS</div><div class="ins-promo-kpi-val">'+d2.los+' n</div></div>'
            +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Lead Time</div><div class="ins-promo-kpi-val">'+d2.avgLeadTime+' d</div></div>'
            +'</div></div>';
        }).join('');

        // Get stay details using first TO
        const d1=getAggData(primaryTO);

        html+='<div class="ins-section" style="border-left:3px solid '+t.color+';padding-left:12px">'
          +'<div style="display:flex;align-items:center;gap:10px;margin-bottom:10px">'
          +'<div style="width:10px;height:10px;border-radius:50%;background:'+t.color+'"></div>'
          +'<div>'
          +'<div style="font-size:15px;font-weight:700;color:var(--text-primary)">'+c.name+'</div>'
          +'<div style="font-size:11px;color:var(--text-muted);margin-top:2px">Type: '+c.type+' &nbsp;·&nbsp; ID: '+c.id+'</div>'
          +'</div></div>'
          +'<div style="display:flex;gap:6px;flex-wrap:wrap;margin-bottom:14px">'+toTagsHtml+'</div>'
          +(function(){
            var showLY   = insCmpType==='ly'  || insCmpType==='both';
            var showSTLY = insCmpType==='stly' || insCmpType==='both';
            var methodLbl = insCmpMethod==='dow'?'DOW':insCmpMethod==='date'?'Date':'Custom';
            var cmpBadge = insCmpType!=='none'
              ? '<span style="font-size:9px;font-weight:600;padding:1px 6px;border-radius:8px;background:#e0f2fe;color:#0369a1;margin-left:6px">'+methodLbl+'</span>'
              : '';
            var title = insCmpType==='none' ? 'Actuals' : 'Actuals vs '+(showLY&&showSTLY?'LY &amp; STLY':showLY?'LY':'STLY');
            var lyRnRef   = showLY   ? aggLyRn   : null;
            var lyRevRef  = showLY   ? aggLyRev  : null;
            var lyAdrRef  = showLY   ? aggLyAdr  : null;
            var stlyRnRef = showSTLY ? aggStlyRn : null;
            var stlyRevRef= showSTLY ? aggStlyRev: null;
            var stlyAdrRef= showSTLY ? aggStlyAdr: null;
            return '<div class="ins-section-title">'+title+cmpBadge+'</div>'
              +'<div class="ins-kpi-row">'
              +kpiCard('Room Nights',aggRn,lyRnRef,stlyRnRef)
              +kpiCard('Revenue',aggRev,lyRevRef,stlyRevRef,'€')
              +kpiCard('ADR',aggAdr,lyAdrRef,stlyAdrRef,'€')
              +kpiCard('RevPAR',aggRevpar,showLY?Math.round(aggLyAdr*0.72):null,showSTLY?Math.round(aggStlyAdr*0.72):null,'€')
              +'</div>';
          })()
          +collapsible('cstay_'+c.id,'More Details',(function(){
            // ── Stay Details (above operator names) ──────────────────
            const totalGuests1=Math.round(d1.rn*(parseFloat(d1.avgAdults)+parseFloat(d1.avgChildren)));
            const stayDetailsHtml=
              '<div class="ins-detail-card" style="margin-bottom:12px">'
              +'<div class="ins-detail-card-title">Stay Details</div>'
              +'<div class="ins-detail-row"><span class="ins-detail-key">Avg LOS</span><span class="ins-detail-val">'+d1.los+' nights</span></div>'
              +'<div class="ins-detail-row"><span class="ins-detail-key">Avg Adults</span><span class="ins-detail-val">'+d1.avgAdults+'</span></div>'
              +'<div class="ins-detail-row"><span class="ins-detail-key">Avg Children</span><span class="ins-detail-val">'+d1.avgChildren+'</span></div>'
              +'<div class="ins-detail-row"><span class="ins-detail-key">Total Guests</span><span class="ins-detail-val">'+fmtN(totalGuests1)+'</span></div>'
              +'<div class="ins-detail-row"><span class="ins-detail-key">Avg Lead Time</span><span class="ins-detail-val">'+d1.avgLeadTime+' days</span></div>'
              +'</div>';

            // ── Room Type breakdown ───────────────────────────────────
            const RT_NAMES=['Standard','Superior','Deluxe','Suite','Jr. Suite','Family'];
            const RT_COLORS=['#006461','#0891b2','#6366f1','#f59e0b','#ec4899','#10b981'];
            const rtTotal=aggRn||1;
            const rtRows=RT_NAMES.map(function(rtn,ri){
              const seed=(c.id.charCodeAt(4)||1)*(ri+1);
              const rtRnPct=Math.max(5,Math.min(35,10+(seed*7+ri*13)%25));
              const rtRn=Math.round(aggRn*rtRnPct/100);
              const rtAdr=Math.round(aggAdr*(0.8+ri*0.08+(seed%10)*0.01));
              const rtRev=rtRn*rtAdr;
              const rtRevpar=Math.round(rtAdr*0.72);
              const rtLos=(parseFloat(d1.los)*(0.9+ri*0.05+(seed%5)*0.02)).toFixed(1);
              const rtAvgA=(parseFloat(d1.avgAdults)*(0.95+ri*0.03)).toFixed(1);
              const rtAvgC=(parseFloat(d1.avgChildren)*(0.9+ri*0.04)).toFixed(1);
              const rtGuests=Math.round(rtRn*(parseFloat(rtAvgA)+parseFloat(rtAvgC)));
              const rtMix=rtRnPct;
              return '<div style="padding:8px 10px;border-radius:7px;border:1px solid '+RT_COLORS[ri]+'33;background:var(--surface-2);margin-bottom:6px">'
                +'<div style="display:flex;align-items:center;gap:6px;margin-bottom:7px">'
                +'<span style="width:8px;height:8px;border-radius:50%;background:'+RT_COLORS[ri]+';flex-shrink:0"></span>'
                +'<span style="font-size:12px;font-weight:700;color:'+RT_COLORS[ri]+'">'+rtn+'</span>'
                +'<span style="margin-left:auto;font-size:10px;font-weight:700;padding:1px 6px;border-radius:8px;background:'+RT_COLORS[ri]+'20;color:'+RT_COLORS[ri]+'">'+rtMix+'% mix</span>'
                +'</div>'
                +'<div style="display:flex;gap:8px;flex-wrap:wrap">'
                +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Room Nights</div><div class="ins-promo-kpi-val">'+fmtN(rtRn)+'</div></div>'
                +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Revenue</div><div class="ins-promo-kpi-val">'+fmt(rtRev)+'</div></div>'
                +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">ADR</div><div class="ins-promo-kpi-val">'+fmtE(rtAdr)+'</div></div>'
                +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">REVPAR</div><div class="ins-promo-kpi-val">'+fmtE(rtRevpar)+'</div></div>'
                +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Avg LOS</div><div class="ins-promo-kpi-val">'+rtLos+' n</div></div>'
                +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Avg Adults</div><div class="ins-promo-kpi-val">'+rtAvgA+'</div></div>'
                +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Avg Children</div><div class="ins-promo-kpi-val">'+rtAvgC+'</div></div>'
                +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Total Guests</div><div class="ins-promo-kpi-val">'+fmtN(rtGuests)+'</div></div>'
                +'</div></div>';
            }).join('');

            // ── Room Type Breakdown — own block, collapsible ─────────
            const rtBlockHtml=
              '<div class="ins-detail-card" style="margin-top:0">'
              +rtRows
              +'</div>';
            const rtCollHtml=collapsible('crt_'+c.id,'Room Type Breakdown',rtBlockHtml);

            // ── Meal Plan Breakdown — collapsible ─────────────────────
            const MP_DEFS=[
              {key:'AI',name:'All Inclusive',  color:'#004948'},
              {key:'HB',name:'Half Board',     color:'#C4FF45'},
              {key:'BB',name:'Bed & Breakfast',color:'#52d9ce'},
              {key:'RO',name:'Room Only',      color:'#D97706'},
              {key:'FB',name:'Full Board',     color:'#d7f7ed'},
            ];
            const mpRows=MP_DEFS.map(function(mp,mi){
              const seed2=(c.id.charCodeAt(5)||2)*(mi+1)*3;
              const mpRnPct=Math.max(4,Math.min(40,8+(seed2*11+mi*17)%32));
              const mpRn=Math.round(aggRn*mpRnPct/100);
              const mpAdrNet=Math.round(aggAdr*(0.78+mi*0.06+(seed2%8)*0.01));
              const mpAdrGross=Math.round(mpAdrNet*(1.12+mi*0.02+(seed2%5)*0.01));
              const mpAdr=mpAdrNet;
              const mpRev=mpRn*mpAdrGross;
              const mpRevpar=Math.round(mpAdrGross*0.72);
              const mpLos=(parseFloat(d1.los)*(0.88+mi*0.07+(seed2%4)*0.02)).toFixed(1);
              const mpAvgA=(parseFloat(d1.avgAdults)*(0.93+mi*0.04)).toFixed(1);
              const mpAvgC=(parseFloat(d1.avgChildren)*(0.85+mi*0.06)).toFixed(1);
              const mpAvgGuest=(parseFloat(mpAvgA)+parseFloat(mpAvgC)).toFixed(1);
              const mpTotGuests=Math.round(mpRn*parseFloat(mpAvgGuest));
              return '<div style="padding:8px 10px;border-radius:7px;border:1px solid '+mp.color+'33;background:var(--surface-2);margin-bottom:6px">'
                +'<div style="display:flex;align-items:center;gap:6px;margin-bottom:7px">'
                +'<span style="width:8px;height:8px;border-radius:50%;background:'+mp.color+';flex-shrink:0"></span>'
                +'<span style="font-size:12px;font-weight:700;color:'+mp.color+'">'+mp.name+'</span>'
                +'<span style="margin-left:auto;font-size:10px;font-weight:700;padding:1px 6px;border-radius:8px;background:'+mp.color+'20;color:'+mp.color+'">'+mpRnPct+'% mix</span>'
                +'</div>'
                +'<div style="display:flex;gap:8px;flex-wrap:wrap">'
                +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Room Nights</div><div class="ins-promo-kpi-val">'+fmtN(mpRn)+'</div></div>'
                +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Revenue</div><div class="ins-promo-kpi-val">'+fmt(mpRev)+'</div></div>'
                +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">ADR</div><div class="ins-promo-kpi-val">'+fmtE(mpAdr)+'</div></div>'
                +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">ADR Net</div><div class="ins-promo-kpi-val">'+fmtE(mpAdrNet)+'</div></div>'
                +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">ADR Gross</div><div class="ins-promo-kpi-val">'+fmtE(mpAdrGross)+'</div></div>'
                +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">REVPAR</div><div class="ins-promo-kpi-val">'+fmtE(mpRevpar)+'</div></div>'
                +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Avg LOS</div><div class="ins-promo-kpi-val">'+mpLos+' n</div></div>'
                +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Avg Adults</div><div class="ins-promo-kpi-val">'+mpAvgA+'</div></div>'
                +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Avg Children</div><div class="ins-promo-kpi-val">'+mpAvgC+'</div></div>'
                +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Avg Guest</div><div class="ins-promo-kpi-val">'+mpAvgGuest+'</div></div>'
                +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Total Guests</div><div class="ins-promo-kpi-val">'+fmtN(mpTotGuests)+'</div></div>'
                +'</div></div>';
            }).join('');
            const mpBlockHtml='<div class="ins-detail-card" style="margin-top:0">'+mpRows+'</div>';
            const mpCollHtml=collapsible('cmp_'+c.id,'Meal Plan Breakdown',mpBlockHtml);

            // ── By Operator — collapsible ─────────────────────────────
            const opCollHtml=collapsible('cop_'+c.id,'By Operator',perTOHtml);

            // ── Source Geo breakdown (after Stay Details) ────────────
            const GEO_DEFS=[
              {key:'UK',name:'United Kingdom',color:'#006461'},
              {key:'SP',name:'Spain',          color:'#0891b2'},
              {key:'US',name:'United States',  color:'#6366f1'},
              {key:'MX',name:'Mexico',         color:'#f59e0b'},
              {key:'DE',name:'Germany',        color:'#ec4899'},
              {key:'FR',name:'France',         color:'#10b981'},
            ];
            const geoSeed=(c.id.charCodeAt(6)||1);
            var geoTotal=0;
            const geoPcts=GEO_DEFS.map(function(g,gi){
              const p=Math.max(3,Math.min(40,6+(geoSeed*(gi+2)+gi*11)%34));
              geoTotal+=p; return p;
            });
            // Normalise to 100%
            const geoRows=GEO_DEFS.map(function(g,gi){
              const pct=Math.round(geoPcts[gi]/geoTotal*100);
              const barW=Math.max(2,pct);
              return '<div style="margin-bottom:6px">'
                +'<div style="display:flex;align-items:center;gap:6px;margin-bottom:2px">'
                +'<span style="width:7px;height:7px;border-radius:50%;background:'+g.color+';flex-shrink:0"></span>'
                +'<span style="font-size:11px;color:var(--text-primary);flex:1">'+g.name+'</span>'
                +'<span style="font-size:11px;font-weight:700;color:'+g.color+'">'+pct+'%</span>'
                +'</div>'
                +'<div style="height:5px;background:#e5e7eb;border-radius:3px;overflow:hidden">'
                +'<div style="height:100%;width:'+barW+'%;background:'+g.color+';border-radius:3px"></div>'
                +'</div></div>';
            }).join('');
            const geoBlockHtml='<div class="ins-detail-card" style="margin-top:0">'+geoRows+'</div>';
            const geoCollHtml=collapsible('cgeo_'+c.id,'Source Geo',geoBlockHtml);

            // ── Projected with release date ───────────────────────────
            const relDate=c.releaseDate ? new Date(c.releaseDate) : null;
            const TODAY_INS=new Date(2026,2,9);
            const daysLeft=relDate ? Math.round((relDate-TODAY_INS)/86400000) : null;
            const relFmt=relDate ? (relDate.getDate()+'/'+(relDate.getMonth()+1)+'/'+relDate.getFullYear()) : '—';
            const daysLeftHtml=daysLeft===null ? '' : daysLeft>0
              ? '<span style="font-size:10px;font-weight:700;color:#f59e0b;margin-left:8px">⏳ '+daysLeft+' day'+(daysLeft!==1?'s':'')+' to release</span>'
              : '<span style="font-size:10px;font-weight:700;color:#16a34a;margin-left:8px">✓ Released</span>';
            const projHeaderHtml='<div class="ins-section-title" style="margin-top:12px;display:flex;align-items:center;flex-wrap:wrap;gap:4px">'
              +'Projected <span class="ins-projected-badge">📋 Based on Contract</span>'
              +'</div>'
              +'<div style="font-size:11px;color:var(--text-muted);margin-bottom:6px">'
              +'Release date: <strong style="color:var(--text-primary)">'+relFmt+'</strong>'
              +daysLeftHtml
              +'</div>';

            return stayDetailsHtml
              +geoCollHtml
              +rtCollHtml
              +mpCollHtml
              +opCollHtml
              +projHeaderHtml
              +'<div class="ins-kpi-row">'
              +'<div class="ins-kpi"><div class="ins-kpi-label">Proj. RN</div><div class="ins-kpi-val">'+fmtN(Math.round(aggRn*1.08))+'</div></div>'
              +'<div class="ins-kpi"><div class="ins-kpi-label">Proj. Revenue</div><div class="ins-kpi-val">€'+(aggRev*1.14/1000).toFixed(1)+'k</div></div>'
              +'<div class="ins-kpi"><div class="ins-kpi-label">Proj. ADR</div><div class="ins-kpi-val">€'+Math.round(aggAdr*1.05)+'</div></div>'
              +'<div class="ins-kpi"><div class="ins-kpi-label">Proj. RevPAR</div><div class="ins-kpi-val">€'+Math.round(aggAdr*1.05*0.72)+'</div></div>'
              +'</div>';
          })())
          +'</div>';
      });
    }

    if(insView==='overall') {
      // ── Aggregate data across ALL operators ─────────────────────
      function getOverallData(factor) {
        const meals  =insSelected.meal.length   ?insSelected.meal   :ALL_MEALS;
        const origins=insSelected.origin.length ?insSelected.origin :ALL_ORIGINS;
        const f=factor!==undefined?factor:getDateFactor();
        let rn=0,rev=0,lyRn=0,lyRev=0,stlyRn=0,stlyRev=0,los=0,avgA=0,avgC=0,lead=0,cs=0,ds=0,cnt=0;
        INS_TOS.forEach(function(t){
          origins.forEach(function(o){ meals.forEach(function(m){
            const d=toData(t.id,o,m);
            rn+=d.rn;rev+=d.rev;lyRn+=d.lyRn;lyRev+=d.lyRev;stlyRn+=d.stlyRn;stlyRev+=d.stlyRev;
            los+=parseFloat(d.los);avgA+=parseFloat(d.avgAdults);avgC+=parseFloat(d.avgChildren);
            lead+=d.avgLeadTime;cs+=d.contractShare;ds+=d.dynamicShare;cnt++;
          });});
        });
        const n=cnt||1;
        rn=Math.round(rn*f);rev=Math.round(rev*f);lyRn=Math.round(lyRn*f);lyRev=Math.round(lyRev*f);
        stlyRn=Math.round(stlyRn*f);stlyRev=Math.round(stlyRev*f);
        const adr=Math.round(rev/Math.max(rn,1)),lyAdr=Math.round(lyRev/Math.max(lyRn,1)),stlyAdr=Math.round(stlyRev/Math.max(stlyRn,1));
        return {
          rn,rev,adr,lyRn,lyRev,lyAdr,stlyRn,stlyRev,stlyAdr,
          revpar:Math.round(adr*0.72),lyRevpar:Math.round(lyAdr*0.72),stlyRevpar:Math.round(stlyAdr*0.72),
          los:(los/n).toFixed(1),avgAdults:(avgA/n).toFixed(1),avgChildren:(avgC/n).toFixed(1),avgLeadTime:Math.round(lead/n),
          contractShare:cs/n,promoShare:1-cs/n,dynamicShare:ds/n,staticShare:1-ds/n,
          projRn:Math.round(rn*1.08),projAdr:Math.round(adr*1.05),projRev:Math.round(rev*1.14),
        };
      }

      const ovData = getOverallData();
      const tourSeriesPctOv = Math.round(ovData.staticShare * 100 * 0.35);
      const dynamicPctOv    = Math.round(ovData.dynamicShare * 100);
      const staticFITPctOv  = Math.round(ovData.staticShare * 100 * 0.65);

      // ── KPI Summary ─────────────────────────────────────────────
      html+='<div class="ins-section">'
        +'<div style="display:flex;align-items:center;gap:8px;margin-bottom:14px">'
        +'<div style="width:10px;height:10px;border-radius:50%;background:#006461"></div>'
        +'<span style="font-size:15px;font-weight:700;color:var(--text-primary)">All Tour Operators — Overall View</span>'
        +'</div>'
        +'<div class="ins-section-title">Actuals vs LY &amp; STLY</div>'
        +'<div class="ins-kpi-row">'
        +kpiCard('Room Nights',ovData.rn,ovData.lyRn,ovData.stlyRn)
        +kpiCard('Revenue',ovData.rev,ovData.lyRev,ovData.stlyRev,'€')
        +kpiCard('ADR',ovData.adr,ovData.lyAdr,ovData.stlyAdr,'€')
        +kpiCard('RevPAR',ovData.revpar,ovData.lyRevpar,ovData.stlyRevpar,'€')
        +kpiCard('Total Adults',Math.round(ovData.rn*parseFloat(ovData.avgAdults)),Math.round(ovData.lyRn*parseFloat(ovData.avgAdults)),Math.round(ovData.stlyRn*parseFloat(ovData.avgAdults)))
        +kpiCard('Total Children',Math.round(ovData.rn*parseFloat(ovData.avgChildren)),Math.round(ovData.lyRn*parseFloat(ovData.avgChildren)),Math.round(ovData.stlyRn*parseFloat(ovData.avgChildren)))
        +'</div></div>';

      // ── More Details (Stay Details + Channel) ───────────────────
      const chDetailRowsOv = [
        {name:'Dynamic / Online', pct:dynamicPctOv,   color:'#0ea5e9'},
        {name:'Static FIT',       pct:staticFITPctOv, color:'#967EF3'},
        {name:'Tour Series',      pct:tourSeriesPctOv,color:'#ec4899'},
      ].map(function(ch){
        const chRn=Math.round(ovData.rn*ch.pct/100), chRev=Math.round(ovData.rev*ch.pct/100);
        const chAdr=Math.round(ovData.adr*(0.9+ch.pct*0.002));
        return '<div style="padding:8px 10px;border-radius:7px;border:1px solid '+ch.color+'33;background:var(--surface-2);margin-bottom:6px">'
          +'<div style="display:flex;align-items:center;gap:6px;margin-bottom:7px">'
          +'<span style="width:8px;height:8px;border-radius:50%;background:'+ch.color+'"></span>'
          +'<span style="font-size:12px;font-weight:700;color:'+ch.color+'">'+ch.name+'</span>'
          +'<span style="margin-left:auto;font-size:10px;font-weight:700;padding:1px 6px;border-radius:8px;background:'+ch.color+'20;color:'+ch.color+'">'+ch.pct+'% share</span>'
          +'</div>'
          +'<div style="display:flex;gap:8px;flex-wrap:wrap">'
          +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Room Nights</div><div class="ins-promo-kpi-val">'+fmtN(chRn)+'</div></div>'
          +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Revenue</div><div class="ins-promo-kpi-val">'+fmt(chRev)+'</div></div>'
          +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">ADR</div><div class="ins-promo-kpi-val">'+fmtE(chAdr)+'</div></div>'
          +'</div></div>';
      }).join('');

      const channelCardOv='<div class="ins-detail-card">'
        +'<div class="ins-detail-card-title">Channel Share</div>'
        +'<div class="ins-donut-wrap">'
        +donutSvg(ovData.dynamicShare)
        +'<div class="ins-donut-legend">'
        +'<div class="ins-donut-leg-item"><div class="ins-donut-leg-dot" style="background:#0ea5e9"></div><span>'+dynamicPctOv+'% Dynamic / Online</span></div>'
        +'<div class="ins-donut-leg-item"><div class="ins-donut-leg-dot" style="background:#967EF3"></div><span>'+staticFITPctOv+'% Static FIT</span></div>'
        +'<div class="ins-donut-leg-item"><div class="ins-donut-leg-dot" style="background:#ec4899"></div><span>'+tourSeriesPctOv+'% Tour Series</span></div>'
        +'</div></div></div>'
        +collapsible('ov_ch_detail','Channel Stay Details',chDetailRowsOv);

      // ── Room Type Mix (using averaged seed) ─────────────────────
      const ovRtSegs = ['Standard','Superior','Deluxe','Suite','Jr. Suite','Family'].map(function(n,i){
        return {name:n, pct:Math.max(5,Math.min(30,10+(i*17+3)%20)), color:['#006461','#0891b2','#6366f1','#f59e0b','#ec4899','#10b981'][i]};
      });
      const ovRtMixCard = '<div class="ins-detail-card"><div class="ins-detail-card-title">Room Type Mix</div>'+multiPieSvg(ovRtSegs)+'</div>';

      // ── Meal Plan Mix ─────────────────────────────────────────────
      const ovMpSegs = [{name:'All Inclusive',pct:32,color:'#006461'},{name:'Half Board',pct:25,color:'#0891b2'},{name:'Bed & Bkfst',pct:28,color:'#6366f1'},{name:'Room Only',pct:15,color:'#f59e0b'}];
      const ovMpMixCard = '<div class="ins-detail-card"><div class="ins-detail-card-title">Meal Plan Mix</div>'+multiPieSvg(ovMpSegs)+'</div>';

      // ── Source Geo ────────────────────────────────────────────────
      const GEO_OV = [{key:'UK',name:'United Kingdom',color:'#006461'},{key:'SP',name:'Spain',color:'#0891b2'},{key:'US',name:'United States',color:'#6366f1'},{key:'MX',name:'Mexico',color:'#f59e0b'},{key:'DE',name:'Germany',color:'#ec4899'},{key:'FR',name:'France',color:'#10b981'}];
      const geoTotalOv = 100;
      const geoPctsOv = [28,22,18,14,10,8];
      const geoRowsOv = GEO_OV.map(function(g,i){
        return '<div style="margin-bottom:6px">'
          +'<div style="display:flex;align-items:center;gap:6px;margin-bottom:2px">'
          +'<span style="width:7px;height:7px;border-radius:50%;background:'+g.color+'"></span>'
          +'<span style="font-size:11px;color:var(--text-primary);flex:1">'+g.name+'</span>'
          +'<span style="font-size:11px;font-weight:700;color:'+g.color+'">'+geoPctsOv[i]+'%</span>'
          +'</div>'
          +'<div style="height:5px;background:#e5e7eb;border-radius:3px"><div style="height:100%;width:'+geoPctsOv[i]+'%;background:'+g.color+';border-radius:3px"></div></div>'
          +'</div>';
      }).join('');

      // ── Room Type Breakdown rows ──────────────────────────────────
      const ovRtRows = ovRtSegs.map(function(seg){
        const rtRn=Math.round(ovData.rn*seg.pct/100), rtRev=Math.round(ovData.rev*seg.pct/100);
        const rtAdr=Math.round(ovData.adr*(0.9+ovRtSegs.indexOf(seg)*0.06));
        return '<div style="padding:8px 10px;border-radius:7px;border:1px solid '+seg.color+'33;background:var(--surface-2);margin-bottom:6px">'
          +'<div style="display:flex;align-items:center;gap:6px;margin-bottom:7px">'
          +'<span style="width:8px;height:8px;border-radius:50%;background:'+seg.color+'"></span>'
          +'<span style="font-size:12px;font-weight:700;color:'+seg.color+'">'+seg.name+'</span>'
          +'<span style="margin-left:auto;font-size:10px;font-weight:700;padding:1px 6px;border-radius:8px;background:'+seg.color+'20;color:'+seg.color+'">'+seg.pct+'% mix</span>'
          +'</div>'
          +'<div style="display:flex;gap:8px;flex-wrap:wrap">'
          +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Room Nights</div><div class="ins-promo-kpi-val">'+fmtN(rtRn)+'</div></div>'
          +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Revenue</div><div class="ins-promo-kpi-val">'+fmt(rtRev)+'</div></div>'
          +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">ADR</div><div class="ins-promo-kpi-val">'+fmtE(rtAdr)+'</div></div>'
          +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">RevPAR</div><div class="ins-promo-kpi-val">'+fmtE(Math.round(rtAdr*0.72))+'</div></div>'
          +'</div></div>';
      }).join('');

      // ── Meal Plan Breakdown rows ──────────────────────────────────
      const ovMpRows = ovMpSegs.map(function(seg){
        const mpRn=Math.round(ovData.rn*seg.pct/100), mpRev=Math.round(ovData.rev*seg.pct/100);
        const mpAdr=Math.round(ovData.adr*(0.85+ovMpSegs.indexOf(seg)*0.05));
        return '<div style="padding:8px 10px;border-radius:7px;border:1px solid '+seg.color+'33;background:var(--surface-2);margin-bottom:6px">'
          +'<div style="display:flex;align-items:center;gap:6px;margin-bottom:7px">'
          +'<span style="width:8px;height:8px;border-radius:50%;background:'+seg.color+'"></span>'
          +'<span style="font-size:12px;font-weight:700;color:'+seg.color+'">'+seg.name+'</span>'
          +'<span style="margin-left:auto;font-size:10px;font-weight:700;padding:1px 6px;border-radius:8px;background:'+seg.color+'20;color:'+seg.color+'">'+seg.pct+'% mix</span>'
          +'</div>'
          +'<div style="display:flex;gap:8px;flex-wrap:wrap">'
          +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Room Nights</div><div class="ins-promo-kpi-val">'+fmtN(mpRn)+'</div></div>'
          +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Revenue</div><div class="ins-promo-kpi-val">'+fmt(mpRev)+'</div></div>'
          +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">ADR Net</div><div class="ins-promo-kpi-val">'+fmtE(mpAdr)+'</div></div>'
          +'<div class="ins-promo-kpi"><div class="ins-promo-kpi-label">Avg LOS</div><div class="ins-promo-kpi-val">'+ovData.los+' n</div></div>'
          +'</div></div>';
      }).join('');

      const moreBodyOv =
        '<div class="ins-detail-grid">'+stayDetailsHtml(ovData)+channelCardOv+'</div>'
        +'<div class="ins-detail-grid">'+ovRtMixCard+ovMpMixCard+'</div>'
        +collapsible('ov_rt','Room Type Breakdown','<div class="ins-detail-card" style="margin-top:0">'+ovRtRows+'</div>')
        +collapsible('ov_mp','Meal Plan Breakdown','<div class="ins-detail-card" style="margin-top:0">'+ovMpRows+'</div>')
        +collapsible('ov_geo','Source Geo','<div class="ins-detail-card" style="margin-top:0">'+geoRowsOv+'</div>');

      html+='<div class="ins-section" style="border-left:3px solid #006461;padding-left:12px">'
        +collapsible('ov_more','More Details',moreBodyOv)
        +'</div>';

      // ── By Month Summary ─────────────────────────────────────────
      const fromDate=new Date(document.getElementById('insDateFrom')?.value||'2025-01-01');
      const toDate=new Date(document.getElementById('insDateTo')?.value||'2025-12-31');
      const MONTH_NAMES=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
      var months=[];
      var cur=new Date(fromDate.getFullYear(),fromDate.getMonth(),1);
      while(cur<=toDate){
        months.push({year:cur.getFullYear(),month:cur.getMonth()});
        cur=new Date(cur.getFullYear(),cur.getMonth()+1,1);
      }
      if(months.length===0) months.push({year:fromDate.getFullYear(),month:fromDate.getMonth()});

      var monthRowsHtml = months.map(function(mo){
        // factor = fraction of month within the selected date range
        var mStart=new Date(mo.year,mo.month,1);
        var mEnd=new Date(mo.year,mo.month+1,0); // last day of month
        var overlapStart=mStart<fromDate?fromDate:mStart;
        var overlapEnd=mEnd>toDate?toDate:mEnd;
        var days=Math.max(1,(overlapEnd-overlapStart)/86400000+1);
        var totalDays=mEnd.getDate();
        var mFactor=days/totalDays;
        // Seasonal multiplier — peak summer/shoulder spring/autumn, trough Jan/Feb
        var SEASONAL=[0.72,0.75,0.88,0.93,0.97,1.08,1.15,1.12,1.04,0.95,0.82,0.78];
        var seasonal=SEASONAL[mo.month]||1.0;
        var d=getOverallData(mFactor*seasonal);
        var label=MONTH_NAMES[mo.month]+' '+mo.year;
        var occ=Math.round(d.rn/(totalDays*INS_TOS.length)*100/4); // approximate occ%

        return '<div style="border:1px solid var(--border);border-radius:8px;padding:14px 16px;margin-bottom:10px;background:var(--surface-1)">'
          // Month header
          +'<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:12px">'
          +'<span style="font-size:14px;font-weight:700;color:#006461">'+label+'</span>'
          +'<span style="font-size:10px;color:var(--text-muted);background:#f1f5f9;padding:2px 8px;border-radius:10px">'+(months.length > 1 ? 'period segment' : 'full period')+'</span>'
          +'</div>'
          // KPI row
          +'<div style="display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin-bottom:12px">'
          +[
            ['Room Nights', fmtN(d.rn), 'LY '+delta(d.rn,d.lyRn), d.rn>=d.lyRn],
            ['Revenue',     fmt(d.rev),  'LY '+delta(d.rev,d.lyRev), d.rev>=d.lyRev],
            ['ADR',         fmtE(d.adr), 'LY '+delta(d.adr,d.lyAdr), d.adr>=d.lyAdr],
            ['RevPAR',      fmtE(d.revpar),'LY '+delta(d.revpar,d.lyRevpar), d.revpar>=d.lyRevpar],
          ].map(function(k){
            return '<div style="background:var(--surface-2);border-radius:6px;padding:8px 10px">'
              +'<div style="font-size:10px;color:var(--text-muted);margin-bottom:2px">'+k[0]+'</div>'
              +'<div style="font-size:14px;font-weight:700;color:var(--text-primary)">'+k[1]+'</div>'
              +'<div style="font-size:10px;font-weight:600;color:'+(k[3]?'#16a34a':'#dc2626')+';margin-top:2px">'+k[2]+'</div>'
              +'</div>';
          }).join('')
          +'</div>'
          // Secondary metrics
          +'<div style="display:grid;grid-template-columns:repeat(4,1fr);gap:6px;margin-bottom:10px">'
          +[
            ['Avg LOS', d.los+' n'],
            ['Lead Time', d.avgLeadTime+' d'],
            ['Avg Adults', d.avgAdults],
            ['Avg Children', d.avgChildren],
          ].map(function(k){
            return '<div style="display:flex;align-items:center;justify-content:space-between;font-size:11px;padding:4px 8px;background:#f8fafc;border-radius:5px;border:1px solid #e5e7eb">'
              +'<span style="color:var(--text-muted)">'+k[0]+'</span>'
              +'<span style="font-weight:600;color:var(--text-primary)">'+k[1]+'</span>'
              +'</div>';
          }).join('')
          +'</div>'
          // Channel mini-bar
          +'<div style="font-size:10px;font-weight:700;color:var(--text-muted);text-transform:uppercase;letter-spacing:.4px;margin-bottom:4px">Channel Mix</div>'
          +'<div style="display:flex;height:8px;border-radius:4px;overflow:hidden;margin-bottom:4px">'
          +'<div style="flex:'+dynamicPctOv+';background:#0ea5e9"></div>'
          +'<div style="flex:'+staticFITPctOv+';background:#967EF3"></div>'
          +'<div style="flex:'+tourSeriesPctOv+';background:#ec4899"></div>'
          +'</div>'
          +'<div style="display:flex;gap:12px;font-size:10px;color:var(--text-muted)">'
          +'<span><span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:#0ea5e9;margin-right:3px"></span>'+dynamicPctOv+'% Dynamic</span>'
          +'<span><span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:#967EF3;margin-right:3px"></span>'+staticFITPctOv+'% Static FIT</span>'
          +'<span><span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:#ec4899;margin-right:3px"></span>'+tourSeriesPctOv+'% Tour Series</span>'
          +'</div>'
          +'</div>';
      }).join('');

      html+='<div class="ins-section">'
        +collapsible('ov_monthly','By Month Summary',monthRowsHtml)
        +'</div>';
    }

    container.innerHTML=html||'<div style="padding:40px;text-align:center;color:var(--text-muted)">No data for selected filters.</div>';
    } catch(e) { container.innerHTML='<div style="padding:20px;color:#dc2626;font-family:monospace;font-size:12px"><strong>Render error:</strong> '+e.message+'<br><pre>'+e.stack+'</pre></div>'; }
  }

  // ── Public API ────────────────────────────────────────────────
  window.insSetView=function(v){
    insView=v;
    document.getElementById('insByTO').classList.toggle('active',v==='to');
    document.getElementById('insByContract').classList.toggle('active',v==='contract');
    var ovBtn=document.getElementById('insByOverall'); if(ovBtn) ovBtn.classList.toggle('active',v==='overall');
    document.getElementById('insTOFilterWrap').style.display       =v==='to'      ?'':'none';
    document.getElementById('insContractFilterWrap').style.display =v==='contract'?'':'none';
    var ccw=document.getElementById('insContractCmpWrap');
    if(ccw) ccw.style.display=v==='contract'?'flex':'none';
    const cg=document.getElementById('insCompareGroup');
    if(cg) cg.style.display=(v==='contract'||v==='overall')?'none':'';
    if(v==='contract'||v==='overall'){insCompare=false;const w=document.getElementById('insCompareWrap');if(w)w.style.display='none';const b=document.getElementById('insCompareBtn');if(b){b.textContent='+ Add TO';b.classList.remove('active');}}
    insRender();
  };

  window.insSetMetric=function(m){
    insMetric=m;
    // Keep Top N metric in sync when switching RN/Revenue
    if(m==='rn'||m==='rev') insTopMetric=m;
    insRender();
  };
  window.insSetTopN=function(n){ insTopN=n; insRender(); };
  window.insSetTopMetric=function(m){ insTopMetric=m; insRender(); };

  window.insToggleCompare=function(){
    insCompare=!insCompare;
    const btn=document.getElementById('insCompareBtn');
    const wrap=document.getElementById('insCompareWrap');
    if(btn){btn.textContent=insCompare?'✕ Remove':'+ Add TO';btn.classList.toggle('active',insCompare);}
    if(wrap){wrap.style.display=insCompare?'flex':'none';}
    insRender();
  };

  // ── Expose collapsible toggles (needed for inline onclick in dynamic HTML)
  window.insToggleColl = function(key) {
    collOpen[key] = !collOpen[key];
    insRender();
  };

  window.insTogglePromo = function(key) {
    promoOpen[key] = !promoOpen[key];
    insRender();
  };

  // ── Wire Insights tab click → insRender
  if (!document._insTabPatched) {
    document._insTabPatched = true;
    document.addEventListener('click', function(e) {
      var btn = e.target.closest('[data-tab="insights"]');
      if (btn) { setTimeout(insRender, 50); }
    });
  }

  // ── Initial render if already on Insights tab
  setTimeout(function() {
    var panel = document.getElementById('toPanel-insights');
    if (panel && panel.style.display !== 'none') insRender();
  }, 200);

})();



/* ═══ AUTOPILOT CLOSE OUT STRATEGIES ═══ */
(function() {
  var strategies = [
    {
      id: 1, name: 'Close Out Strategy', active: true,
      stayFrom: '4/5/2026', stayTo: '4/5/2026',
      dowStay: [true,true,true,true,true,true,true],
      dbaOp: 'less than', dbaVal: 50,
      demOp: 'more than', demVal: 55,
      comOp: 'more than', comVal: 55,
      roomTypes: 'All Rooms', mealPlans: 'All Plans',
      activeFrom: '4/5/2026', activeTo: '4/5/2026',
      dowActive: [true,true,true,true,true,true,true],
      updatedBy: 'Admin Karina Demo Hotel', updatedAt: '4/5/2026 8:32 pm'
    }
  ];
  var editingId = null;
  var nextId = 2;
  var DOW = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];

  function getFormDow(groupId) {
    var boxes = document.querySelectorAll('#'+groupId+' input[type=checkbox]');
    return Array.from(boxes).map(function(b){ return b.checked; });
  }
  function dowStr(arr) {
    var all = arr.every(Boolean);
    if (all) return 'Every day';
    var days = DOW.filter(function(_,i){ return arr[i]; });
    return days.join(', ') || 'None';
  }

  window.apNewStrategy = function() {
    editingId = null;
    document.getElementById('apFormTitle').textContent = 'New Close Out Strategy';
    document.getElementById('apStrategyName').value = '';
    var today = new Date().toISOString().split('T')[0];
    document.getElementById('apStayFrom').value = today;
    document.getElementById('apStayTo').value = today;
    document.getElementById('apActiveFrom').value = today;
    document.getElementById('apActiveTo').value = today;
    document.getElementById('apDbaVal').value = 50;
    document.getElementById('apDemVal').value = 55;
    document.getElementById('apComVal').value = 55;
    document.getElementById('apStrategyForm').style.display = 'block';
    document.getElementById('apStrategyForm').scrollIntoView({behavior:'smooth'});
  };

  window.apCancelForm = function() {
    document.getElementById('apStrategyForm').style.display = 'none';
    editingId = null;
  };

  window.apSaveStrategy = function() {
    var name = document.getElementById('apStrategyName').value.trim() || 'Unnamed Strategy';
    var s = {
      id: editingId || nextId++,
      name: name, active: true,
      stayFrom: document.getElementById('apStayFrom').value,
      stayTo:   document.getElementById('apStayTo').value,
      dowStay:  getFormDow('apDowStay'),
      dbaOp:    document.getElementById('apDbaOp').value,
      dbaVal:   document.getElementById('apDbaVal').value,
      demOp:    document.getElementById('apDemOp').value,
      demVal:   document.getElementById('apDemVal').value,
      comOp:    document.getElementById('apComOp').value,
      comVal:   document.getElementById('apComVal').value,
      activeFrom: document.getElementById('apActiveFrom').value,
      activeTo:   document.getElementById('apActiveTo').value,
      dowActive:  getFormDow('apDowActive'),
      roomTypes: 'All Rooms', mealPlans: 'All Plans',
      updatedBy: 'Admin', updatedAt: new Date().toLocaleDateString()
    };
    if (editingId) {
      strategies = strategies.map(function(x){ return x.id===editingId ? s : x; });
    } else {
      strategies.push(s);
    }
    document.getElementById('apStrategyForm').style.display = 'none';
    editingId = null;
    renderStrategies();
  };

  window.apEditStrategy = function(id) {
    var s = strategies.find(function(x){ return x.id===id; });
    if (!s) return;
    editingId = id;
    document.getElementById('apFormTitle').textContent = 'Edit Strategy: ' + s.name;
    document.getElementById('apStrategyName').value = s.name;
    document.getElementById('apStayFrom').value = s.stayFrom;
    document.getElementById('apStayTo').value = s.stayTo;
    document.getElementById('apActiveFrom').value = s.activeFrom;
    document.getElementById('apActiveTo').value = s.activeTo;
    document.getElementById('apDbaVal').value = s.dbaVal;
    document.getElementById('apDemVal').value = s.demVal;
    document.getElementById('apComVal').value = s.comVal;
    document.getElementById('apStrategyForm').style.display = 'block';
    document.getElementById('apStrategyForm').scrollIntoView({behavior:'smooth'});
  };

  window.apDeleteStrategy = function(id) {
    if (!confirm('Delete this strategy?')) return;
    strategies = strategies.filter(function(x){ return x.id!==id; });
    renderStrategies();
  };

  window.apToggleStrategy = function(id, cb) {
    strategies = strategies.map(function(x){ return x.id===id ? Object.assign({},x,{active:cb.checked}) : x; });
    renderStrategies();
  };

  function renderStrategies() {
    var el = document.getElementById('apStrategyList');
    if (!el) return;
    if (strategies.length === 0) {
      el.innerHTML = '<div style="text-align:center;padding:40px;color:var(--text-muted);border:1px dashed var(--border);border-radius:10px">No strategies yet. Click <strong>New Strategy</strong> to create one.</div>';
      return;
    }
    el.innerHTML = strategies.map(function(s) {
      var criteriaStr = 'DBA ' + s.dbaOp + ' ' + s.dbaVal +
        ' &amp; Demand Occ ' + s.demOp + ' ' + s.demVal + '%' +
        ' &amp; Committed Occ ' + s.comOp + ' ' + s.comVal + '%';
      return '<div style="padding:16px 0;border-bottom:1px solid var(--border)">'
        + '<div style="display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:10px">'
        +   '<div style="display:flex;align-items:center;gap:10px">'
        +     '<span style="font-size:16px;font-weight:700;color:var(--text-primary)">' + s.name + '</span>'
        +     '<span style="font-size:11px;background:'+(s.active?'#d1fae5':'#f3f4f6')+';color:'+(s.active?'#065f46':'#6b7280')+';padding:2px 8px;border-radius:10px;font-weight:600">'+(s.active?'ACTIVE':'INACTIVE')+'</span>'
        +   '</div>'
        +   '<div style="display:flex;align-items:center;gap:8px">'
        +     '<label class="ds-radio-label" style="font-size:13px">'
        +       '<input type="checkbox" class="ds-checkbox" '+(s.active?'checked':'')+' onchange="apToggleStrategy('+s.id+',this)"> Active'
        +     '</label>'
        +     '<button onclick="apEditStrategy('+s.id+')" class="ds-btn ds-btn-ghost">Edit</button>'
        +     '<button onclick="apDeleteStrategy('+s.id+')" class="ds-btn ds-btn-danger">Delete</button>'
        +   '</div>'
        + '</div>'
        + '<div style="display:grid;grid-template-columns:auto 1fr;gap:4px 12px;font-size:13px">'
        +   '<span style="color:var(--text-muted);font-weight:600">Stay Date:</span><span style="color:var(--text-primary)">' + s.stayFrom + (s.stayTo && s.stayTo!==s.stayFrom ? ' – '+s.stayTo : '') + ' &nbsp;·&nbsp; ' + dowStr(s.dowStay) + '</span>'
        +   '<span style="color:var(--text-muted);font-weight:600">Criteria:</span><span style="color:var(--text-primary)">' + criteriaStr + '</span>'
        +   '<span style="color:var(--text-muted);font-weight:600">Active:</span><span style="color:var(--text-primary)">' + s.activeFrom + (s.activeTo && s.activeTo!==s.activeFrom ? ' – '+s.activeTo : '') + '</span>'
        +   (s.updatedBy ? '<span style="color:var(--text-muted);font-weight:600">Updated:</span><span style="color:var(--text-primary)">' + s.updatedAt + ' by ' + s.updatedBy + '</span>' : '')
        + '</div>'
        + '</div>';
    }).join('');
  }

  // Init on settings tab open
  document.addEventListener('click', function(e) {
    var btn = e.target.closest('[data-stab]');
    if (btn && btn.dataset.stab === 'autopilot') setTimeout(renderStrategies, 50);
  });
  setTimeout(renderStrategies, 300);
})();

/* ═══ BOARD TYPES EDITOR — AG Grid ═══ */
(function() {
  var _btGridApi = null;

  var _btRowData = [
    { name: 'Room Only',       codes: 'RO' },
    { name: 'Bed & Breakfast', codes: 'BB' },
    { name: 'Half Board',      codes: 'HB' },
    { name: 'Full Board',      codes: 'FB' },
    { name: 'All Inclusive',   codes: 'AI' },
  ];

  function _btUpdateLimitUI() {
    var count = 0;
    if (_btGridApi) {
      _btGridApi.forEachNode(function() { count++; });
    }
    var noteEl = document.getElementById('btMaxNote');
    var addBtn = document.getElementById('btAddBtn');
    if (noteEl) noteEl.style.display = count >= 10 ? 'inline' : 'none';
    if (addBtn) addBtn.style.display = count >= 10 ? 'none' : '';
  }

  function _btDeleteRow(rowId) {
    if (!_btGridApi) return;
    var count = 0;
    _btGridApi.forEachNode(function() { count++; });
    if (count <= 1) return; // keep at least one row
    var node = _btGridApi.getRowNode(rowId);
    if (node) _btGridApi.applyTransaction({ remove: [node.data] });
    _btUpdateLimitUI();
  }
  window._btDeleteRow = _btDeleteRow;

  function _btInit() {
    var el = document.getElementById('boardTypesGrid');
    if (!el || _btGridApi) return;

    var colDefs = [
      {
        field: 'name',
        headerName: 'Meal Plan Name',
        flex: 2,
        editable: true,
        cellStyle: { fontSize: '13px' },
      },
      {
        field: 'codes',
        headerName: 'Code(s)',
        headerTooltip: 'Comma-separated codes (e.g. AI, AI1)',
        flex: 1,
        editable: true,
        cellStyle: { fontSize: '13px', color: '#6b7280' },
        cellRenderer: function(p) {
          if (!p.value) return '<span style="color:#d1d5db;font-style:italic">e.g. AI, AI1</span>';
          return p.value.split(',').map(function(c) {
            return '<span style="background:#f0fffe;color:#006461;border:1px solid #b2f0ec;border-radius:4px;padding:1px 7px;font-size:12px;font-weight:600;margin-right:3px">' + c.trim() + '</span>';
          }).join('');
        },
      },
      {
        headerName: '',
        width: 56,
        suppressSizeToFit: true,
        sortable: false,
        resizable: false,
        cellRenderer: function(p) {
          var btn = document.createElement('button');
          btn.innerHTML = '<span class="material-icons" style="font-size:16px;line-height:1">delete_outline</span>';
          btn.title = 'Remove';
          btn.style.cssText = 'border:none;background:none;color:#dc2626;cursor:pointer;display:flex;align-items:center;justify-content:center;width:28px;height:28px;border-radius:4px;';
          btn.addEventListener('mouseenter', function() { btn.style.background = '#fff1f2'; });
          btn.addEventListener('mouseleave', function() { btn.style.background = 'none'; });
          btn.addEventListener('click', function() { _btDeleteRow(String(p.node.id)); });
          return btn;
        },
        cellStyle: { display: 'flex', alignItems: 'center', justifyContent: 'center' },
      },
    ];

    var AG = _realAgGrid || agGrid;
    _btGridApi = AG.createGrid(el, {
      theme: sharedTheme,
      columnDefs: colDefs,
      rowData: _btRowData,
      rowHeight: 44,
      headerHeight: 40,
      domLayout: 'autoHeight',
      singleClickEdit: true,
      onCellValueChanged: function() { _markConfigDirty(); },
      stopEditingWhenCellsLoseFocus: true,
      suppressMovableColumns: true,
      suppressCellFocus: false,
      defaultColDef: { resizable: false, sortable: false },
    });
  }

  window.btAddRow = function() {
    // Check row limit only if grid is ready
    if (_btGridApi) {
      var count = 0;
      _btGridApi.forEachNode(function() { count++; });
      if (count >= 10) { _btUpdateLimitUI(); return; }
    }
    // Always open the modal
    var overlay = document.getElementById('addMealPlanOverlay');
    if (!overlay) return;
    document.getElementById('mpNameInput').value = '';
    document.getElementById('mpCodeInput').value = '';
    overlay.classList.add('open');
    setTimeout(function() { document.getElementById('mpNameInput').focus(); }, 80);
  };

  window.btModalClose = function() {
    var overlay = document.getElementById('addMealPlanOverlay');
    if (overlay) overlay.classList.remove('open');
  };

  window.btModalSave = function() {
    var name = (document.getElementById('mpNameInput').value || '').trim();
    var codes = (document.getElementById('mpCodeInput').value || '').trim().toUpperCase();
    if (!name) { document.getElementById('mpNameInput').focus(); return; }
    if (!_btGridApi) return;
    _btGridApi.applyTransaction({ add: [{ name: name, codes: codes }] });
    _btUpdateLimitUI();
    btModalClose();
    _markConfigDirty();
  };

  window.btSave = function() { configPageSave(); };

  // Init when DOM is ready
  document.addEventListener('DOMContentLoaded', function() {
    // May need a tick if settingsPage is initially hidden
    setTimeout(_btInit, 100);
  });
})();

/* ═══ STOP RULES AG Grid ═══ */
(function() {
  var _srGridApi = null;

  var _srRowData = [
    { roomType: 'Standard Room', threshold: '85% sold', action: 'Alert only', frequency: 'Daily',   enabled: true  },
    { roomType: 'Superior Room', threshold: '70% sold', action: 'Alert only', frequency: 'Daily',   enabled: true  },
    { roomType: 'Deluxe Room',   threshold: '80% sold', action: 'Alert only', frequency: 'Weekly',  enabled: true  },
    { roomType: 'Suite',         threshold: '70% sold', action: 'Alert only', frequency: 'None',    enabled: false },
    { roomType: 'Family Room',   threshold: '90% sold', action: 'Alert only', frequency: 'Weekly',  enabled: true  },
  ];

  function _srInit() {
    var el = document.getElementById('stopRulesGrid');
    if (!el || _srGridApi) return;

    var AG = _realAgGrid || agGrid;

    var colDefs = [
      {
        field: 'roomType',
        headerName: 'Room Type',
        flex: 2,
        editable: false,
        cellStyle: { fontSize: '13px', fontWeight: '600' },
      },
      {
        field: 'threshold',
        headerName: 'Alert Threshold',
        flex: 1,
        editable: true,
        cellEditor: 'agSelectCellEditor',
        cellEditorParams: { values: ['70% sold', '80% sold', '85% sold', '90% sold'] },
        cellStyle: { fontSize: '13px' },
      },
      {
        field: 'action',
        headerName: 'Action',
        flex: 1,
        editable: true,
        cellEditor: 'agSelectCellEditor',
        cellEditorParams: { values: ['Alert only', 'Auto stop sale'] },
        cellStyle: { fontSize: '13px' },
      },
      {
        field: 'frequency',
        headerName: 'Alert Frequency',
        flex: 1,
        editable: true,
        cellEditor: 'agSelectCellEditor',
        cellEditorParams: { values: ['Daily', 'Weekly', 'Monthly', 'None'] },
        cellStyle: { fontSize: '13px' },
      },
      {
        field: 'enabled',
        headerName: 'Enabled',
        width: 96,
        editable: false,
        cellRenderer: function(p) {
          var wrap = document.createElement('div');
          wrap.style.cssText = 'display:flex;align-items:center;justify-content:center;height:100%';
          var cb = document.createElement('input');
          cb.type = 'checkbox';
          cb.checked = !!p.value;
          cb.className = 'ds-checkbox';
          cb.addEventListener('change', function() {
            p.node.setDataValue('enabled', cb.checked);
            _markConfigDirty();
          });
          wrap.appendChild(cb);
          return wrap;
        },
        cellStyle: { display: 'flex', alignItems: 'center', justifyContent: 'center' },
      },
    ];

    _srGridApi = AG.createGrid(el, {
      theme: sharedTheme,
      columnDefs: colDefs,
      rowData: _srRowData,
      rowHeight: 44,
      headerHeight: 40,
      domLayout: 'autoHeight',
      singleClickEdit: true,
      onCellValueChanged: function() { _markConfigDirty(); },
      stopEditingWhenCellsLoseFocus: true,
      suppressMovableColumns: true,
      defaultColDef: { resizable: false, sortable: false },
    });
  }

  // Expose so configPageSave can stop editing
  window._srGridApi = function() { return _srGridApi; };

  // Init when stopsales tab is clicked
  document.addEventListener('click', function(e) {
    var btn = e.target.closest('[data-stab]');
    if (btn && btn.dataset.stab === 'stopsales') setTimeout(_srInit, 50);
  });
})();



/* ═══ SEGMENTS MULTI-SELECT DROPDOWN ═══ */
window.segToggleDD = function() {
  var panel = document.getElementById('segDropPanel');
  if (!panel) return;
  var isOpen = panel.style.display !== 'none';
  panel.style.display = isOpen ? 'none' : 'block';
  var btn = document.getElementById('segDropBtn');
  if (btn) btn.classList.toggle('open', !isOpen);
  if (!isOpen) { var s = document.getElementById('segSearch'); if(s) s.focus(); }
};

window.segUpdateLabel = function() {
  var checked = document.querySelectorAll('#segOptionsList .seg-opt input:checked');
  var label   = document.getElementById('segDropLabel');
  var chips   = document.getElementById('segChips');
  var names   = Array.from(checked).map(function(cb) {
    return cb.closest('label').textContent.trim();
  });

  if (label) {
    label.textContent = names.length === 0 ? 'No segments selected'
      : names.length === 1 ? names[0]
      : names.length + ' segments selected';
  }

  if (chips) {
    chips.innerHTML = names.map(function(n) {
      return '<span class="ds-chip">'
        + n
        + '<button class="ds-chip-close" data-seg-name="' + n + '" onclick="segRemoveChip(this)" title="Remove">&times;</button>'
        + '</span>';
    }).join('');
  }
};

window.segRemoveChip = function(btn) {
  var name = btn.getAttribute('data-seg-name');
  document.querySelectorAll('#segOptionsList .seg-opt').forEach(function(lbl) {
    if (lbl.textContent.trim() === name) lbl.querySelector('input').checked = false;
  });
  segUpdateLabel();
};

window.segFilter = function(val) {
  var q = val.toLowerCase();
  document.querySelectorAll('#segOptionsList .seg-opt').forEach(function(lbl) {
    lbl.style.display = lbl.textContent.trim().toLowerCase().includes(q) ? '' : 'none';
  });
};

window.segSelectAll = function(select) {
  document.querySelectorAll('#segOptionsList .seg-opt input').forEach(function(cb) { cb.checked = select; });
  segUpdateLabel();
};

// Keep segment dropdown open when clicking items inside it
(function() {
  var panel = document.getElementById('segDropPanel');
  if (panel) panel.addEventListener('click', function(e) { e.stopPropagation(); });
})();

// Close on outside click
document.addEventListener('click', function(e) {
  var wrap = document.getElementById('segDropWrap');
  if (wrap && !wrap.contains(e.target)) {
    var panel = document.getElementById('segDropPanel');
    if (panel) panel.style.display = 'none';
  }
});

// Option hover styles + init
setTimeout(function() {
  document.querySelectorAll('#segOptionsList .seg-opt').forEach(function(lbl) {
    lbl.addEventListener('mouseenter', function() { this.style.background = 'var(--surface-2)'; });
    lbl.addEventListener('mouseleave', function() { this.style.background = ''; });
  });
  segUpdateLabel();
}, 300);





/* ═══ CALENDAR MONTH PICKER — two-panel ═══ */
(function() {
  var MONTH_ABBR    = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  var drLeftYear    = 2026; // left panel year; right panel = drLeftYear+1

  /*
   * Three-phase click cycle:
   *   phase 0 — nothing selected (drSelStartIdx = null)
   *   phase 1 — start chosen, waiting for end (drSelEndIdx = null)
   *   phase 2 — full range selected
   * Clicking any month in phase 2 clears back to phase 0.
   */
  var drSelStartIdx = null; // null = no selection
  var drSelEndIdx   = null;
  var drPhase       = 0;

  /* End fallback: if no explicit end, auto = start + 11 (capped) */
  function getEndIdx(startIdx) {
    return Math.min(startIdx + 11, ALL_MONTHS.length - 1);
  }

  /* Build the 4×3 month grid for a given year into containerId */
  function renderGrid(containerId, year) {
    var el = document.getElementById(containerId);
    if (!el) return;
    /* Determine effective range for highlights */
    var startIdx = drSelStartIdx;
    var endIdx   = (drPhase === 2 && drSelEndIdx !== null) ? drSelEndIdx
                 : (drPhase === 1 && startIdx !== null)    ? startIdx   // start only, no strip
                 : null;

    el.innerHTML = MONTH_ABBR.map(function(name, mi) {
      var col = mi % 4;
      /* Find index in ALL_MONTHS for (year, month) */
      var idx = -1;
      for (var i = 0; i < ALL_MONTHS.length; i++) {
        if (ALL_MONTHS[i].year === year && ALL_MONTHS[i].month === (mi + 1)) { idx = i; break; }
      }
      var inData  = idx >= 0;
      var isStart = inData && startIdx !== null && idx === startIdx;
      /* In phase 1 no range strip — only the start circle */
      var isEnd   = drPhase === 2 && inData && endIdx !== null && idx === endIdx;
      var isMid   = drPhase === 2 && inData && startIdx !== null && endIdx !== null
                    && idx > startIdx && idx < endIdx;

      var cls = 'caldr-cell c-col' + col;
      if (!inData)           cls += ' c-disabled';
      if (isStart && isEnd)  cls += ' c-start c-end';
      else if (isStart)      cls += ' c-start';
      else if (isEnd)        cls += ' c-end';
      else if (isMid)        cls += ' c-mid';

      var onclick = inData ? ' onclick="calDRMonthClick(' + idx + ')"' : '';
      return '<div class="' + cls + '"' + onclick + '>'
           + '<span class="caldr-cell-bg"></span>'
           + '<span class="caldr-cell-lbl">' + name + '</span>'
           + '</div>';
    }).join('');
  }

  /* Render both panels + footer */
  function renderBothGrids() {
    var leftYearEl  = document.getElementById('calDRLeftYear');
    var rightYearEl = document.getElementById('calDRRightYear');
    if (leftYearEl)  leftYearEl.textContent  = drLeftYear;
    if (rightYearEl) rightYearEl.textContent = drLeftYear + 1;
    renderGrid('calDRLeftGrid',  drLeftYear);
    renderGrid('calDRRightGrid', drLeftYear + 1);
    var foot = document.getElementById('calDRFooterLabel');
    if (foot) {
      if (drPhase === 0 || drSelStartIdx === null) {
        foot.textContent = 'Select a start month';
      } else if (drPhase === 1) {
        var startM = ALL_MONTHS[drSelStartIdx];
        foot.textContent = (startM ? startM.name : '') + ' \u2013 ?  (select end month)';
      } else {
        var startM = ALL_MONTHS[drSelStartIdx];
        var endM   = drSelEndIdx !== null ? ALL_MONTHS[drSelEndIdx] : null;
        foot.textContent = (startM ? startM.name : '') + ' \u2013 ' + (endM ? endM.name : '');
      }
    }
    /* Dim Apply when range not complete */
    var applyBtn = document.querySelector('#calDRPanel .caldr-btn-apply');
    if (applyBtn) applyBtn.style.opacity = drPhase === 2 ? '1' : '0.4';
  }

  /* ── Month click ── */
  window.calDRMonthClick = function(idx) {
    if (drPhase === 1) {
      /* Second click: set end; swap if needed so start ≤ end */
      drSelEndIdx = idx;
      if (drSelEndIdx < drSelStartIdx) {
        var tmp = drSelStartIdx; drSelStartIdx = drSelEndIdx; drSelEndIdx = tmp;
      }
      drPhase = 2;
    } else {
      /* Phase 0 or 2: any click starts a fresh selection */
      drSelStartIdx = idx;
      drSelEndIdx   = null;
      drPhase       = 1;
    }
    renderBothGrids();
  };

  /* ── Navigate years (both panels move together) ── */
  window.calDRNav = function(delta) {
    drLeftYear += delta;
    renderBothGrids();
  };

  /* ── Open / close ── */
  window.calDRToggle = function() {
    var panel   = document.getElementById('calDRPanel');
    var trigger = document.getElementById('calDRTrigger');
    if (!panel) return;
    if (panel.style.display !== 'none') { panel.style.display = 'none'; return; }
    /* Sync to current calendar state — open with a confirmed range showing */
    drSelStartIdx = calStartIdx;
    drSelEndIdx   = Math.min(calStartIdx + calDisplayView - 1, ALL_MONTHS.length - 1);
    drPhase       = 2;
    drLeftYear    = ALL_MONTHS[calStartIdx] ? ALL_MONTHS[calStartIdx].year : 2026;
    /* Position below trigger, keep in viewport */
    var rect     = trigger.getBoundingClientRect();
    var panelW   = 443; /* ~210+210+1+padding */
    var left     = rect.left;
    if (left + panelW > window.innerWidth - 8) left = Math.max(8, window.innerWidth - panelW - 8);
    panel.style.left    = left + 'px';
    panel.style.top     = (rect.bottom + 6) + 'px';
    panel.style.display = 'block';
    renderBothGrids();
  };

  window.calDRCancel = function() {
    document.getElementById('calDRPanel').style.display = 'none';
  };

  window.calDRApply = function() {
    if (drPhase !== 2 || drSelStartIdx === null || drSelEndIdx === null) return; // range not complete
    document.getElementById('calDRPanel').style.display = 'none';
    var startM = ALL_MONTHS[drSelStartIdx];
    var endM   = ALL_MONTHS[drSelEndIdx];
    var lbl = document.getElementById('calDRLabel');
    if (lbl) lbl.textContent = (startM ? startM.name : '') + ' \u2013 ' + (endM ? endM.name : '');
    calStartIdx = drSelStartIdx;
    /* Derive view length from selected range */
    var viewLen = drSelEndIdx - drSelStartIdx + 1;
    calView = viewLen; calDisplayView = viewLen;
    if (typeof calSetDisplayView === 'function') calSetDisplayView(viewLen);
    else renderCalendar();
    renderCalMonthlySummary();
  };

  /* ── Close on outside click ── */
  document.addEventListener('click', function(e) {
    var panel = document.getElementById('calDRPanel');
    var wrap  = document.getElementById('calDRWrap');
    if (!panel || panel.style.display === 'none') return;
    if (wrap && wrap.contains(e.target)) return;
    panel.style.display = 'none';
  }, true);

  /* ── Compatibility no-ops ── */
  window.applyOutOfRange     = function() {};
  window.applyCalDisplayRange = function() {};

  /* ── Init: 2 months from Jan 2026 on first load ── */
  setTimeout(function() {
    calStartIdx   = 0;
    drSelStartIdx = 0;
    drSelEndIdx   = Math.min(1, ALL_MONTHS.length - 1);
    drPhase       = 2;
    var lbl = document.getElementById('calDRLabel');
    if (lbl) lbl.textContent = ALL_MONTHS[0].name + ' \u2013 ' + ALL_MONTHS[Math.min(1, ALL_MONTHS.length-1)].name;
    if (typeof calSetDisplayView === 'function') calSetDisplayView(2);
    else { calView = 2; calDisplayView = 2; renderCalendar(); }
    renderCalMonthlySummary();
  }, 400);

})();


/* ═══ CALENDAR EVENT TOOLTIPS ═══ */
(function() {
  var tip = null;

  function getTip() {
    if (!tip) {
      tip = document.createElement('div');
      tip.className = 'cal-event-tooltip';
      tip.style.display = 'none';
      document.body.appendChild(tip);
    }
    return tip;
  }

  window.calShowEventTip = function(e, key) {
    var events = (typeof CAL_EVENTS !== 'undefined' && key) ? CAL_EVENTS[key] : null;
    var closures = (typeof PARTIAL_CLOSURES !== 'undefined' && key) ? PARTIAL_CLOSURES[key] : null;
    var hasEvents = events && events.length > 0;
    var hasClosures = closures && Array.isArray(closures) && closures.length > 0;

    // Hide rooms/capacity tooltip first
    window.calHideCapTip();

    var t = getTip();
    var calSvg = '<span class="material-icons" style="font-size:18px;color:#006461;vertical-align:middle;margin-right:2px">today</span>';
    var html = '';

    if (!hasEvents && !hasClosures) {
      // No events — show "No events" placeholder
      html = '<div style="color:#6b7280;font-size:12px;padding:2px 0;display:flex;align-items:center;gap:6px">'
           + calSvg + '<span>No events</span></div>';
    } else {
      if (hasEvents) {
        html += '<div class="cal-event-tooltip-title">' + calSvg + ' Events</div>'
          + events.map(function(ev) {
              return '<div style="margin-bottom:' + (events.length > 1 ? '6px' : '0') + '">'
                + '<div class="cal-event-tooltip-name">• ' + ev.name + '</div>'
                + '<div class="cal-event-tooltip-meta">| ' + ev.type + '<br>' + ev.date + '</div>'
                + '</div>';
            }).join('');
      }
      if (hasClosures) {
        if (hasEvents) html += '<div style="border-top:1px solid #e5e7eb;margin:8px 0 6px"></div>';
        html += '<div class="cal-event-tooltip-title"><span class="material-icons" style="font-size:18px;color:#fbbf24;vertical-align:middle;margin-right:2px">lock_open</span> Closures</div>'
          + closures.map(function(cl) {
              var parts = [];
              if (cl.tos && cl.tos.length) parts.push(cl.tos.join(', '));
              if (cl.roomTypes && cl.roomTypes.length) parts.push(cl.roomTypes.join(', '));
              if (cl.boards && cl.boards.length) parts.push(cl.boards.map(function(b){return b.toUpperCase();}).join(', '));
              return '<div style="margin-bottom:4px"><div class="cal-event-tooltip-name">• ' + (parts.join(' · ') || 'All') + '</div>'
                + '<div class="cal-event-tooltip-meta">| ' + (cl.appliedBy || '') + '</div></div>';
            }).join('');
      }
    }

    t.innerHTML = html;

    // Position below the icon — support all icon class types
    var _evEl = e.target.closest('.wv-event-cal-icon')
             || e.target.closest('.cell-event-icon')
             || e.target.closest('.cell-event-ico');
    if (!_evEl) return;
    var rect = _evEl.getBoundingClientRect();
    t.style.display = 'block';
    var tW = t.offsetWidth || 180;
    var tH = t.offsetHeight || 80;

    var left = rect.left;
    var top  = rect.bottom + 7;

    // Keep within viewport
    if (left + tW > window.innerWidth - 8) left = window.innerWidth - tW - 8;
    if (left < 8) left = 8;
    if (top + tH > window.innerHeight - 8) top = rect.top - tH - 7;

    t.style.left = left + 'px';
    t.style.top  = top  + 'px';
  };

  window.calHideEventTip = function() {
    var t = getTip();
    t.style.display = 'none';
  };

  // Also hide on scroll
  window.addEventListener('scroll', window.calHideEventTip, true);

window.calShowCapTip = function(e, hotel, hotelRooms, to, toRooms, avail, month, day) {
  // Hide events tooltip first
  window.calHideEventTip();

  var tip = document.getElementById('calCapTip');
  if (!tip) return;

  // Check if room type filter is active
  var _fCal = typeof filterState !== 'undefined' ? filterState.cal : {};
  var _rtFilt = _fCal.calFiltRoom || 'all';
  var _rtShares = {standard:0.34,superior:0.24,deluxe:0.18,suite:0.08,'jr. suite':0.10,family:0.06};
  var filteredCap = HOTEL_CAPACITY;
  var rtLabel = '';
  var isFiltered = _rtFilt !== 'all';

  if (isFiltered) {
    var _rtParts = _rtFilt.split(',');
    var _rtMult = _rtParts.reduce(function(a,b){ return a + (_rtShares[b.trim().toLowerCase()] || 0.15); }, 0);
    _rtMult = Math.min(1, _rtMult);
    filteredCap = Math.round(HOTEL_CAPACITY * _rtMult);
    rtLabel = _rtParts.map(function(s){ return s.trim().charAt(0).toUpperCase() + s.trim().slice(1); }).join(', ');
  }

  var filtHotelRooms = Math.round(filteredCap * hotel / 100);
  var filtToRooms = Math.round(filteredCap * to / 100);
  var filtAvail = Math.max(0, filteredCap - filtHotelRooms);

  var infoIco = '<span class="material-icons" style="font-size:20px;color:#00298C;flex-shrink:0">info</span>';
  var html = '';
  if (isFiltered) {
    html += '<div style="display:flex;align-items:center;gap:4px;border:1px solid #00298C;border-radius:4px;padding:6px 4px;margin-bottom:8px">'
      + infoIco
      + '<span style="font-size:14px;font-family:Lato,sans-serif;color:#00298C;line-height:1.2">Filtered: ' + rtLabel + ' (' + filteredCap + ' rooms)</span>'
      + '</div>';
  }
  html += '<div style="font-size:14px;font-weight:700;color:#1c1c1c;margin-bottom:3px;font-family:Lato,sans-serif">Hotel: ' + hotel + '% (' + filtHotelRooms + ' rooms)</div>'
    + '<div style="font-size:14px;color:#8C7843;margin-bottom:3px;font-family:Lato,sans-serif">&nbsp;&nbsp;TO: ' + to + '% (' + filtToRooms + ' rooms)</div>'
    + '<div style="font-size:14px;font-weight:700;color:' + (filtAvail < 10 ? '#dc2626' : '#16a34a') + ';font-family:Lato,sans-serif">'
    + filtAvail + ' rooms available' + (isFiltered ? ' (' + rtLabel + ')' : '') + '</div>';

  tip.innerHTML = html;
  tip.style.display = 'block';
  var x = e.clientX + 14, y = e.clientY - 10;
  if (x + 240 > window.innerWidth) x = e.clientX - 250;
  if (y + 100 > window.innerHeight) y = e.clientY - 110;
  tip.style.left = x + 'px';
  tip.style.top  = y + 'px';
};
window.calHideCapTip = function() {
  var tip = document.getElementById('calCapTip');
  if (tip) tip.style.display = 'none';
};
})();


/* ═══ MONTHLY CALENDAR MULTI-SELECT FILTERS ═══ */
(function() {
  // Track selected values per filter group
  var calFilterState = { to: ['all'], rt: ['all'], mp: ['all'], origin: ['all'], pickup: 365 };

  // Labels for display
  var LABELS = {
    to:     { all: 'Operator', map: { sunwing:'Sunwing', tui:'TUI', 'thomas-cook':'Thomas Cook', 'club-med':'Club Med', jet2:'Jet2' } },
    rt:     { all: 'Room Type',     map: { standard:'Standard', superior:'Superior', deluxe:'Deluxe', suite:'Suite' } },
    mp:     { all: 'Meal Plan',     map: { ai:'All Incl.', hb:'Half Board', bb:'B&B', ro:'Room Only' } },
    origin: { all: 'Source Geo',        map: { UK:'UK', SP:'Spain', US:'US', MX:'Mexico' } },
  };

  function updateLabel(filt) {
    var state = calFilterState[filt];
    var lbl   = document.getElementById('calFilt' + filt.charAt(0).toUpperCase() + filt.slice(1) + 'Label');
    if (!lbl) return;
    if (!state || state.includes('all') || state.length === 0) {
      lbl.textContent = LABELS[filt].all;
    } else if (state.length === 1) {
      lbl.textContent = LABELS[filt].map[state[0]] || state[0];
    } else {
      lbl.textContent = state.length + ' selected';
    }
  }

  window.calMsChange = function(filt, cb) {
    var val = cb.value;
    var allCbs = document.querySelectorAll('.cal-ms[data-filt="' + filt + '"]');

    if (val === 'all') {
      // Check/uncheck all
      allCbs.forEach(function(c) { c.checked = cb.checked; });
      calFilterState[filt] = cb.checked ? ['all'] : [];
    } else {
      var allCb = document.querySelector('.cal-ms[data-filt="' + filt + '"][value="all"]');
      var selected = Array.from(allCbs)
        .filter(function(c) { return c.value !== 'all' && c.checked; })
        .map(function(c) { return c.value; });

      if (selected.length === 0) {
        // Nothing selected — revert to all
        if (allCb) allCb.checked = true;
        calFilterState[filt] = ['all'];
      } else {
        if (allCb) allCb.checked = false;
        calFilterState[filt] = selected;
      }
    }
    updateLabel(filt);
  };

  window.calPickupNumUpdate = function(val) {
    val = val ? parseInt(val) : null;
    calFilterState.pickup = val;
    var lbl = document.getElementById('calFiltPickupLabel');
    if (lbl) lbl.textContent = val ? 'Pickup: ' + val + 'd' : 'Pickup Window';
    var hidden = document.getElementById('calFiltPickup');
    if (hidden) hidden.value = val || '';
  };

  window.calApplyFilters = function() {
    // Sync calFiltTO from filterState for rendering
    if (typeof filterState !== 'undefined' && typeof calFiltTO !== 'undefined') {
      calFiltTO = filterState.cal.calFiltTO || 'all';
    }
    // Close consolidated dropdown
    var dd = document.getElementById('calFiltersDropdown');
    if (dd) dd.style.display = 'none';
    var btn = document.getElementById('calFiltBtn');
    if (btn) btn.classList.remove('active');
    renderCalendar();
  };

  window.calPickupSliderUpdate = function(val) {
    window.calPickupNumUpdate(val);
    var disp = document.getElementById('calFiltPickupVal');
    if (disp) disp.textContent = val ? val + 'd' : '365d';
  };

  window.calToggleFilters = function(btn) {
    var dd = document.getElementById('calFiltersDropdown');
    if (!dd) return;
    var open = dd.style.display !== 'none';
    dd.style.display = open ? 'none' : 'block';
    if (btn) btn.classList.toggle('active', !open);
  };

  window.calResetFilters = function() {
    // Reset unified filterState.cal
    if (typeof filterState !== 'undefined') {
      Object.keys(filterState.cal).forEach(function(k) {
        filterState.cal[k] = 'all';
      });
      if (typeof calFiltTO !== 'undefined') calFiltTO = 'all';
    }
    // Reset pickup buttons
    if (typeof pickupBtnReset === 'function') pickupBtnReset('cal');
    // Refresh checkbox UI
    if (typeof applyFilterUI === 'function') applyFilterUI('calFiltersDropdown');
    // Close dropdown
    var dd = document.getElementById('calFiltersDropdown');
    if (dd) dd.style.display = 'none';
    var btn = document.getElementById('calFiltBtn');
    if (btn) btn.classList.remove('active');
    renderCalendar();
  };

  // Also update calToggleMFilt to include pickup panel
  var _origToggle = window.calToggleMFilt;
  window.calToggleMFilt = function(panelId, btn) {
    var allPanels = ['calFiltTOPanel','calFiltRTPanel','calFiltMPPanel','calFiltOriginPanel','calFiltPickupPanel'];
    var allBtns   = ['calFiltTOBtn','calFiltRTBtn','calFiltMPBtn','calFiltOriginBtn','calFiltPickupBtn'];
    allPanels.forEach(function(pid, i) {
      var p = document.getElementById(pid);
      if (!p) return;
      if (pid === panelId) {
        var isOpen = p.style.display !== 'none';
        p.style.display = isOpen ? 'none' : 'block';
        var b = document.getElementById(allBtns[i]);
        if (b) b.classList.toggle('active', !isOpen);
      } else {
        p.style.display = 'none';
        var b2 = document.getElementById(allBtns[i]);
        if (b2) b2.classList.remove('active');
      }
    });
  };

  // Close consolidated filters dropdown on outside click
  document.addEventListener('click', function(e) {
    if (!e.target.closest('#calFiltersWrap')) {
      var dd = document.getElementById('calFiltersDropdown');
      if (dd) dd.style.display = 'none';
      var btn = document.getElementById('calFiltBtn');
      if (btn) btn.classList.remove('active');
    }
  });
})();


/* ═══ CELL METRICS v2 — Segment mode + LY/STLY/Fcst for H & T ═══ */
(function() {

  // ── State ──────────────────────────────────────────────────────
  var cmMode      = 'individual'; // always individual (segments always shown)
  var cmSegs      = ['fit','dynamic','series']; // all selected by default
  var cmMetrics   = ['hocc','tocc','availRooms'];  // selected metric keys (Hotel Occ + T Occ + Avail Rooms by default)
  var cmHotel     = true;         // include Hotel row

  var SEG_COLORS = { fit: '#0891b2', dynamic: '#7c3aed', series: '#f59e0b' };
  var SEG_LABELS = { fit: 'FIT', dynamic: 'Dyn', series: 'Ser' };
  var SEG_FULL   = { fit: 'Static FIT', dynamic: 'TO Dynamic', series: 'Tour Series' };

  var METRIC_DEFS = {
    occ:         { label:'Occ',       color:'#5883ed', fmt: function(v){ return v+'%'; },        maxVal:100   },
    adr:         { label:'ADR',       color:'#7c3aed', fmt: function(v){ return '$'+v; },         maxVal:400   },
    rev:         { label:'Rev',       color:'#ea580c', fmt: function(v){ return '$'+Math.round(v/1000)+'k'; }, maxVal:50000 },
    pickup:      { label:'Pkp',       color:'#16a34a', fmt: function(v){ return (v>=0?'+':'')+v; }, maxVal:30 },
    rn:          { label:'RN',        color:'#2e65e8', fmt: function(v){ return String(v); },     maxVal:210   },
    revpar:      { label:'RVP',       color:'#9333ea', fmt: function(v){ return '$'+v; },         maxVal:500   },
    ly:          { label:'LY',        color:'#93c5fd', fmt: function(v){ return v; },             maxVal:100   },
    stly:        { label:'STLY',      color:'#6ee7b7', fmt: function(v){ return v; },             maxVal:100   },
    fcst:        { label:'Fcst',      color:'#fbbf24', fmt: function(v){ return v; },             maxVal:100   },
    avgLos:      { label:'LOS',       color:'#0891b2', fmt: function(v){ return v.toFixed(1)+'n';}, maxVal:14 },
    avgLeadTime: { label:'Lead',      color:'#6366f1', fmt: function(v){ return v+'d'; },         maxVal:365   },
    avgAdults:   { label:'AdA',       color:'#2e65e8', fmt: function(v){ return v.toFixed(1); },  maxVal:4     },
    avgChildren: { label:'AdC',       color:'#d33030', fmt: function(v){ return v.toFixed(1); },  maxVal:2     },
    availRooms:  { label:'AvR',       color:'#16a34a', fmt: function(v){ return String(v); },     maxVal:210   },
    availGuar:   { label:'AvG',       color:'#ea580c', fmt: function(v){ return String(v); },     maxVal:30    },
    bizMixTO:    { label:'TO%',       color:'#006461', fmt: function(v){ return v+'%'; },         maxVal:100   },
    bizMixDirect:{ label:'Dir%',      color:'#0284c7', fmt: function(v){ return v+'%'; },         maxVal:100   },
    bizMixOTA:   { label:'OTA%',      color:'#D97706', fmt: function(v){ return v+'%'; },         maxVal:100   },
    rateTO:      { label:'TO-R',      color:'#0f766e', fmt: function(v){ return '$'+v; },         maxVal:500   },
    ratePromo:   { label:'Prmo%',     color:'#d97706', fmt: function(v){ return v+'%'; },         maxVal:50    },
    rateBase:    { label:'Base',      color:'#9333ea', fmt: function(v){ return '$'+v; },         maxVal:500   },
  };

  // ── Compute how many rows the current config produces ──────────
  function countRows() {
    // In new model each checked item = 1 row
    // In individual mode, each metric × each segment = rows
    if (cmMode === 'individual') {
      var segCount = cmSegs.length || 1;
      return Math.min(cmMetrics.length * segCount, 99);
    }
    return cmMetrics.length;
  }

  // isAdding: true when user just checked a metric — show over-limit warning
  //           false when user just unchecked — only disable Apply, no warning text
  function updateHint(isAdding) {
    var rows = countRows();
    var hint = document.getElementById('calCmHint');
    if (hint) {
      hint.textContent = rows + ' / 4 rows';
      hint.style.color = rows > 4 ? '#dc2626' : rows === 4 ? '#f59e0b' : '#6b7280';
    }
    var applyBtn = document.getElementById('cmApplyBtn');
    var limitHint = document.getElementById('cmSelectUpTo4');
    if (rows > 4) {
      // Disable Apply in all over-limit cases
      if (applyBtn) { applyBtn.disabled = true; applyBtn.style.background = '#9ca3af'; applyBtn.style.cursor = 'not-allowed'; applyBtn.style.opacity = '0.7'; }
      // Show "Select up to 4" warning only when actively adding beyond limit
      if (limitHint) limitHint.style.display = (isAdding !== false) ? '' : 'none';
    } else {
      if (applyBtn) { applyBtn.disabled = false; applyBtn.style.background = '#006461'; applyBtn.style.cursor = 'pointer'; applyBtn.style.opacity = '1'; }
      if (limitHint) limitHint.style.display = 'none';
    }
  }

  // ── Mode toggle ────────────────────────────────────────────────
  window.cmSetMode = function() {}; // toggle removed — mode is always individual

  // ── Segment toggle (min 1 must always stay selected) ───────────
  window.cmToggleSeg = function(key, cb) {
    var willUncheck = cb.classList.contains('checked');
    if (willUncheck && cmSegs.length <= 1) return; // block deselecting last one
    cb.classList.toggle('checked');
    var idx = cmSegs.indexOf(key);
    if (cb.classList.contains('checked')) {
      if (idx < 0) cmSegs.push(key);
    } else {
      if (idx >= 0) cmSegs.splice(idx, 1);
    }
    _syncSegAllCb();
    updateHint();
    renderCalendar();
  };

  // ── Toggle all segments (clicking All selects all; if all selected, no-op since min 1) ──
  window.cmToggleAllSegs = function(cb) {
    var allSelected = cmSegs.length === 3;
    if (allSelected) return; // already all selected — nothing to do
    cmSegs = ['fit','dynamic','series'];
    document.querySelectorAll('#cmSegSection .cal-md-cb[data-seg-key]').forEach(function(el) {
      el.classList.add('checked');
    });
    cb.classList.add('checked');
    updateHint();
    renderCalendar();
  };

  function _syncSegAllCb() {
    var allCb = document.getElementById('cmSegAllCb');
    if (allCb) allCb.classList.toggle('checked', cmSegs.length === 3);
  }

  // ── Metric toggle (updates pending state only — Apply to commit) ─
  // Keys starting with 't' = Combined column; 'h' = Individual (Hotel) column.
  // Selecting from one group clears all selections from the other.
  // ── Metric toggle — Hotel (h) and T (t) columns can both be selected freely.
  // Only rule: max 4 total (enforced by updateHint / Apply button disable).
  window.cmToggleMetric = function(key, cb) {
    if (cb.classList.contains('checked')) {
      // Removing — don't show over-limit warning
      cb.classList.remove('checked');
      var idx = cmMetrics.indexOf(key);
      if (idx >= 0) cmMetrics.splice(idx, 1);
      updateHint(false);
    } else {
      // Adding — show warning if over limit
      cb.classList.add('checked');
      if (cmMetrics.indexOf(key) < 0) cmMetrics.push(key);
      updateHint(true);
    }
    // Don't call renderCalendar() here — wait for Apply
  };

  // ── Apply: commit metric selections and re-render ──────────────
  window.cmApplyMetrics = function() {
    if (countRows() > 4) return; // safety guard — button should already be disabled
    var dd = document.getElementById('calMetricsDropdown');
    if (dd) dd.style.display = 'none';
    renderCalendar();
  };

  // ── Reset: uncheck all metrics ──────────────────────────────────
  window.cmResetMetrics = function() {
    cmMetrics = [];
    document.querySelectorAll('#calMetricsDropdown .cal-md-cb').forEach(function(cb) {
      cb.classList.remove('checked');
    });
    updateHint();
  };

  // ── Hotel toggle ───────────────────────────────────────────────
  var hotelToggle = document.getElementById('cmHotelToggle');
  if (hotelToggle) {
    hotelToggle.addEventListener('change', function() {
      cmHotel = this.checked;
      var track = document.getElementById('cmHotelTrack');
      var thumb = document.getElementById('cmHotelThumb');
      if (track) track.style.background = cmHotel ? '#006461' : '#d1d5db';
      if (thumb) thumb.style.left = cmHotel ? '16px' : '2px';
      updateHint();
      renderCalendar();
    });
  }

  window.cmUpdateHint = function() { updateHint(); };


  // ── Build metric rows for a cell ──────────────────────────────
  // Exposed globally so renderCalendar can call it
  window.cmBuildRows = function(cellVals) {
    // New granular key map
    var KEY_MAP = {
      hocc:'hotelOcc', tocc:'toOcc', hadr:'hotelAdr', tadr:'toAdr',
      hrev:'hotelRev', trev:'toRev', hpickup:'hotelPickup', tpickup:'toPickup',
      hrn:'hotelRn', trn:'toRn', hrevpar:'hotelTrev', trevpar:'toTrev',
      havgAdults:'avgAdults', tavgAdults:'avgAdults',
      havgChildren:'avgChildren', tavgChildren:'avgChildren',
      havgLos:'avgLos', tavgLos:'avgLos',
      havgLeadTime:'avgLeadTime', tavgLeadTime:'avgLeadTime',
      htotalGuests:'totalGuests', ttotalGuests:'totalGuests',
      // Per-metric LY (maps to base metric, scaled)
      hlyOcc:'hotelOcc', tlyOcc:'toOcc',
      hlyAdr:'hotelAdr', tlyAdr:'toAdr',
      hlyRev:'hotelRev', tlyRev:'toRev',
      hlyRn:'hotelRn',   tlyRn:'toRn',
      hlyRevpar:'hotelTrev', tlyRevpar:'toTrev',
      hlyLos:'avgLos',   tlyLos:'avgLos',
      // STLY
      hstlyOcc:'hotelOcc', tstlyOcc:'toOcc',
      hstlyAdr:'hotelAdr', tstlyAdr:'toAdr',
      hstlyRev:'hotelRev', tstlyRev:'toRev',
      hstlyRn:'hotelRn',   tstlyRn:'toRn',
      hstlyRevpar:'hotelTrev', tstlyRevpar:'toTrev',
      hstlyLos:'avgLos',   tstlyLos:'avgLos',
      // Fcst
      hfcstOcc:'hotelOcc', tfcstOcc:'toOcc',
      hfcstAdr:'hotelAdr', tfcstAdr:'toAdr',
      hfcstRev:'hotelRev', tfcstRev:'toRev',
      hfcstRn:'hotelRn',   tfcstRn:'toRn',
      hfcstRevpar:'hotelTrev', tfcstRevpar:'toTrev',
      hfcstLos:'avgLos',   tfcstLos:'avgLos',
      // Old compat
      hly:'hotelOcc', tly:'toOcc', hstly:'hotelOcc', tstly:'toOcc',
      hfcst:'hotelOcc', tfcst:'toOcc',
    };
    var KEY_LABELS = {
      hocc:'H-Occ', tocc:'TO-Occ', hadr:'H-ADR', tadr:'TO-ADR',
      hrev:'H-Rev', trev:'TO-Rev', hpickup:'H-Pkp', tpickup:'TO-Pkp',
      hrn:'H-RN', trn:'TO-RN', hrevpar:'H-RVP', trevpar:'TO-RVP',
      havgAdults:'H-AdA', tavgAdults:'TO-AdA',
      havgChildren:'H-AdC', tavgChildren:'TO-AdC',
      havgLos:'H-LOS', tavgLos:'TO-LOS',
      havgLeadTime:'H-Lead', tavgLeadTime:'TO-Lead',
      htotalGuests:'H-Gst', ttotalGuests:'TO-Gst',
      // Per-metric LY colors (blue family)
      hlyOcc:'#93c5fd', tlyOcc:'#6ee7b7',
      hstlyOcc:'#bfdbfe', tstlyOcc:'#a7f3d0',
      hfcstOcc:'#fde68a', tfcstOcc:'#fef3c7',
      hlyAdr:'#c4b5fd', tlyAdr:'#a5b4fc',
      hstlyAdr:'#ddd6fe', tstlyAdr:'#c7d2fe',
      hfcstAdr:'#fde68a', tfcstAdr:'#fef08a',
      hlyRev:'#fdba74', tlyRev:'#fcd34d',
      hstlyRev:'#fed7aa', tstlyRev:'#fef3c7',
      hfcstRev:'#fcd34d', tfcstRev:'#fef9c3',
      hlyRn:'#93c5fd', tlyRn:'#7dd3fc',
      hstlyRn:'#bfdbfe', tstlyRn:'#bae6fd',
      hfcstRn:'#fbbf24', tfcstRn:'#fde68a',
      hlyRevpar:'#d8b4fe', tlyRevpar:'#c4b5fd',
      hstlyRevpar:'#ede9fe', tstlyRevpar:'#e0e7ff',
      hfcstRevpar:'#fef08a', tfcstRevpar:'#fefce8',
      hlyLos:'#a5f3fc', tlyLos:'#67e8f9',
      hstlyLos:'#cffafe', tstlyLos:'#e0f2fe',
      hfcstLos:'#fde68a', tfcstLos:'#fef9c3',
      // Per-metric LY/STLY/Fcst
      hlyOcc:'H-LY-Occ', tlyOcc:'TO-LY-Occ',
      hlyAdr:'H-LY-ADR', tlyAdr:'TO-LY-ADR',
      hlyRev:'H-LY-Rev', tlyRev:'TO-LY-Rev',
      hlyRn:'H-LY-RN', tlyRn:'TO-LY-RN',
      hlyRevpar:'H-LY-RVP', tlyRevpar:'TO-LY-RVP',
      hlyLos:'H-LY-LOS', tlyLos:'TO-LY-LOS',
      hstlyOcc:'H-STLY-Occ', tstlyOcc:'TO-STLY-Occ',
      hstlyAdr:'H-STLY-ADR', tstlyAdr:'TO-STLY-ADR',
      hstlyRev:'H-STLY-Rev', tstlyRev:'TO-STLY-Rev',
      hstlyRn:'H-STLY-RN', tstlyRn:'TO-STLY-RN',
      hstlyRevpar:'H-STLY-RVP', tstlyRevpar:'TO-STLY-RVP',
      hstlyLos:'H-STLY-LOS', tstlyLos:'TO-STLY-LOS',
      hfcstOcc:'H-Fcst-Occ', tfcstOcc:'TO-Fcst-Occ',
      hfcstAdr:'H-Fcst-ADR', tfcstAdr:'TO-Fcst-ADR',
      hfcstRev:'H-Fcst-Rev', tfcstRev:'TO-Fcst-Rev',
      hfcstRn:'H-Fcst-RN', tfcstRn:'TO-Fcst-RN',
      hfcstRevpar:'H-Fcst-RVP', tfcstRevpar:'TO-Fcst-RVP',
      hfcstLos:'H-Fcst-LOS', tfcstLos:'TO-Fcst-LOS',
      hly:'H-LY', tly:'TO-LY', hstly:'H-STLY', tstly:'TO-STLY', hfcst:'H-Fcst', tfcst:'TO-Fcst',
      availRooms:'AvR', availGuar:'T-AvG',
      bizMixTO:'TO%', bizMixDirect:'Dir%', bizMixOTA:'OTA%',
      rateTO:'TO-R', ratePromo:'Prmo%', rateBase:'Base',
      htotalGuests:'#0369a1', ttotalGuests:'#0ea5e9',
    };
    var KEY_COLORS = {
      hocc:'#5883ed', tocc:'#006461', hadr:'#7c3aed', tadr:'#4f46e5',
      hrev:'#ea580c', trev:'#b45309', hpickup:'#16a34a', tpickup:'#0d9488',
      hrn:'#2e65e8',  trn:'#0284c7',  hrevpar:'#9333ea', trevpar:'#7c3aed',
      havgAdults:'#2e65e8', tavgAdults:'#60a5fa',
      havgChildren:'#d33030', tavgChildren:'#f87171',
      havgLos:'#0891b2', tavgLos:'#22d3ee',
      havgLeadTime:'#6366f1', tavgLeadTime:'#a5b4fc',
      hly:'#93c5fd', tly:'#6ee7b7', hstly:'#bfdbfe', tstly:'#a7f3d0',
      hfcst:'#fbbf24', tfcst:'#fde68a',
      availRooms:'#16a34a', availGuar:'#ea580c',
      bizMixTO:'#006461', bizMixDirect:'#0284c7', bizMixOTA:'#D97706',
      rateTO:'#0f766e', ratePromo:'#d97706', rateBase:'#9333ea',
    };
    var COMP_MULTS = {
      hly:.88, tly:.88, hstly:.83, tstly:.83, hfcst:1.04, tfcst:1.04,
      hlyOcc:.88, tlyOcc:.88, hstlyOcc:.83, tstlyOcc:.83, hfcstOcc:1.04, tfcstOcc:1.04,
      hlyAdr:.91, tlyAdr:.91, hstlyAdr:.87, tstlyAdr:.87, hfcstAdr:1.03, tfcstAdr:1.03,
      hlyRev:.89, tlyRev:.89, hstlyRev:.85, tstlyRev:.85, hfcstRev:1.05, tfcstRev:1.05,
      hlyRn:.88, tlyRn:.88, hstlyRn:.83, tstlyRn:.83, hfcstRn:1.04, tfcstRn:1.04,
      hlyRevpar:.90, tlyRevpar:.90, hstlyRevpar:.86, tstlyRevpar:.86, hfcstRevpar:1.04, tfcstRevpar:1.04,
      hlyLos:.95, tlyLos:.95, hstlyLos:.92, tstlyLos:.92, hfcstLos:1.02, tfcstLos:1.02,
    };
    var COMP_KEYS = [
      'hly','tly','hstly','tstly','hfcst','tfcst',
      'hlyOcc','tlyOcc','hstlyOcc','tstlyOcc','hfcstOcc','tfcstOcc',
      'hlyAdr','tlyAdr','hstlyAdr','tstlyAdr','hfcstAdr','tfcstAdr',
      'hlyRev','tlyRev','hstlyRev','tstlyRev','hfcstRev','tfcstRev',
      'hlyRn','tlyRn','hstlyRn','tstlyRn','hfcstRn','tfcstRn',
      'hlyRevpar','tlyRevpar','hstlyRevpar','tstlyRevpar','hfcstRevpar','tfcstRevpar',
      'hlyLos','tlyLos','hstlyLos','tstlyLos','hfcstLos','tfcstLos',
    ];
    var SRC_KEYS  = ['hocc','tocc','hadr','tadr','hrev','trev','hpickup','tpickup',
                     'hrn','trn','hrevpar','trevpar','havgAdults','tavgAdults',
                     'havgChildren','tavgChildren','havgLos','tavgLos','havgLeadTime','tavgLeadTime'];
    var SINGLE_KEYS = ['availRooms','availGuar','bizMixTO','bizMixDirect','bizMixOTA',
                       'rateTO','ratePromo','rateBase'];

    var rows = [];

    cmMetrics.forEach(function(key) {
      if (rows.length >= 4) return;

      // Pickup keys: render compact 3-window grid row
      if (key === 'hpickup' || key === 'tpickup') {
        var isH = key === 'hpickup';
        var pfx = isH ? 'hotelPickup' : 'toPickup';
        var clr = isH ? '#16a34a' : '#0d9488';
        var cells = '';
        pickupDayValues.forEach(function(dv, i) {
          var pv = cellVals[pfx + '_' + i];
          if (pv === undefined) {
            var sc = dv<=1?0.3:dv<=3?0.6:dv<=7?1:Math.min(2,dv/7);
            pv = Math.max(0, Math.round((cellVals[pfx]||0) * sc));
          }
          cells += '<div class="cal-pu-cell">'
            + '<span class="cal-pu-win">' + dv + '</span>'
            + '<span class="cal-pu-val" style="color:' + clr + '">+' + pv + '</span>'
            + '</div>';
        });
        rows.push({ _html: '<div class="cal-pu-grid">' + cells + '</div>', label:'Pkp', color: clr, value:'', raw:0 });
        return;
      }

      // Segment-filtered metric: one row, value scaled by selected segments
      if (cmMode === 'individual' && SRC_KEYS.indexOf(key) >= 0 && COMP_KEYS.indexOf(key) < 0) {
        var baseVal = cellVals[KEY_MAP[key]] || 0;
        var segMults = { fit:0.62, dynamic:0.53, series:0.31 };
        var totalMult = 0.62 + 0.53 + 0.31;
        var selMult = 0;
        cmSegs.forEach(function(s){ selMult += (segMults[s] || 0); });
        var scale = totalMult > 0 ? selMult / totalMult : 1;
        var v = Math.round(baseVal * scale);
        var metricLbl = KEY_LABELS[key] ? KEY_LABELS[key].replace(/^[HT]-/,'') : key;
        var clr = KEY_COLORS[key] || '#006461';
        rows.push({ label: metricLbl, color: clr, value: String(v), raw: v });
        return;
      }

      // Comparison keys (LY/STLY/Fcst per source)
      if (COMP_KEYS.indexOf(key) >= 0) {
        var isH   = key.charAt(0) === 'h';
        var mult  = COMP_MULTS[key] || 1;
        var lbl   = KEY_LABELS[key] || key;
        var clr   = KEY_COLORS[key] || (isH ? '#93c5fd' : '#6ee7b7');
        // Determine base value from the metric suffix
        var baseVal = 0;
        var fmtFn = function(v){ return v+'%'; };
        if (key.indexOf('Occ') >= 0 || key === 'hly' || key === 'tly' || key === 'hstly' || key === 'tstly' || key === 'hfcst' || key === 'tfcst') {
          baseVal = cellVals[isH ? 'hotelOcc' : 'toOcc'] || 0;
          fmtFn = function(v){ return v+'%'; };
        } else if (key.indexOf('Adr') >= 0) {
          baseVal = cellVals[isH ? 'hotelAdr' : 'toAdr'] || 0;
          fmtFn = function(v){ return '$'+v; };
        } else if (key.indexOf('Rev') >= 0 && key.indexOf('Revpar') < 0) {
          baseVal = cellVals[isH ? 'hotelRev' : 'toRev'] || 0;
          fmtFn = function(v){ return '$'+Math.round(v/1000)+'k'; };
        } else if (key.indexOf('Rn') >= 0) {
          baseVal = cellVals[isH ? 'hotelRn' : 'toRn'] || 0;
          fmtFn = function(v){ return String(v); };
        } else if (key.indexOf('Revpar') >= 0) {
          baseVal = cellVals[isH ? 'hotelTrev' : 'toTrev'] || 0;
          fmtFn = function(v){ return '$'+v; };
        } else if (key.indexOf('Los') >= 0) {
          baseVal = cellVals['avgLos'] || 0;
          fmtFn = function(v){ return v.toFixed ? v.toFixed(1)+'n' : v+'n'; };
        }
        var v = typeof baseVal === 'number' ? Math.round(baseVal * mult * 10) / 10 : baseVal;
        rows.push({ label: lbl, color: clr, value: fmtFn(v), raw: v });
        return;
      }

      // Granular H/T keys — explicit format by key suffix
      if (SRC_KEYS.indexOf(key) >= 0) {
        var rawKey = KEY_MAP[key];
        var v = cellVals[rawKey] || 0;
        var lbl = KEY_LABELS[key] || key;
        var clr = KEY_COLORS[key] || '#6b7280';
        var formatted;
        var k = key.toLowerCase();
        if (k.indexOf('revpar') >= 0 || k.indexOf('adr') >= 0) {
          // Dollar per-room metric: $NNN
          formatted = '$' + Math.round(v);
        } else if (k.indexOf('rev') >= 0) {
          // Total revenue: $NNk
          formatted = '$' + Math.round(v / 1000) + 'k';
        } else if (k.indexOf('occ') >= 0) {
          formatted = Math.round(v) + '%';
        } else if (k.indexOf('pickup') >= 0) {
          formatted = (v >= 0 ? '+' : '') + Math.round(v);
        } else if (k.indexOf('rn') >= 0) {
          formatted = Math.round(v) + ' rn';
        } else if (k.indexOf('avglos') >= 0 || k.indexOf('los') >= 0) {
          formatted = (parseFloat(v)||0).toFixed(1) + 'n';
        } else if (k.indexOf('lead') >= 0) {
          formatted = Math.round(v) + 'd';
        } else if (k.indexOf('adults') >= 0 || k.indexOf('children') >= 0 || k.indexOf('guests') >= 0) {
          formatted = String(Math.round(v));
        } else {
          formatted = String(Math.round(v));
        }
        rows.push({ label: lbl, color: clr, value: formatted, raw: v });
        return;
      }

      // Single/shared keys (Business Mix, Selling Rates, Avail)
      if (SINGLE_KEYS.indexOf(key) >= 0) {
        var v = cellVals[key] || 0;
        var lbl2 = KEY_LABELS[key] || key;
        var clr2 = KEY_COLORS[key] || '#6b7280';
        var fmt2;
        if (key === 'availRooms' || key === 'availGuar') fmt2 = Math.round(v) + ' rm';
        else if (key.indexOf('Mix') >= 0) fmt2 = Math.round(v) + '%';
        else if (key.indexOf('rate') >= 0 || key.indexOf('Rate') >= 0 || key === 'rateTO' || key === 'rateBase') fmt2 = '$' + Math.round(v);
        else if (key === 'ratePromo') fmt2 = Math.round(v) + '%';
        else fmt2 = String(Math.round(v));
        rows.push({ label: lbl2, color: clr2, value: fmt2, raw: v });
      }
    });

    return rows.slice(0, 4);
  };
  function adjustColor(hex) {
    // Slightly darken/saturate for T vs Hotel
    return hex;
  }

  // Init: pre-check default metrics in the UI and update hint
  setTimeout(function() {
    cmMetrics.forEach(function(key) {
      var cb = document.querySelector('#calMetricsDropdown .cal-md-cb[data-cm-key="' + key + '"]');
      if (cb) cb.classList.add('checked');
    });
    updateHint();
  }, 300);

})();


/* ═══ HEATMAP CONFIGURATOR ═══════════════════════════════════════ */
(function() {

  // ── State ──────────────────────────────────────────────────────
  var hmState = {
    type: '',        // active heatmap type key
    grey:  { threshold: 85, params: {} },
    green: { threshold: 60, params: {} },
    blue:  { params: {} },
    enabled: false,
    condition: { enabled: false, metric: 'hotel', op: '>', value: 50 },
    stopSalesRoomTypes: [],  // [] = all room types, or array of selected room type names
    colors: {}  // custom color overrides, e.g. { grey: '#D33030', blue: '#FDCF61', green: '#CEF2D1' }
  };

  // ── Type definitions ───────────────────────────────────────────
  // Each type defines what grey/green/blue mean and what inputs to show
  var HM_TYPES = {
    stopsales: {
      label: 'Stop Sales',
      icon: '🔒',
      svgPath: 'M18 8h-1V6c0-2.76-2.24-5-5-5S7 3.24 7 6v2H6c-1.1 0-2 .9-2 2v10c0 1.1.9 2 2 2h12c1.1 0 2-.9 2-2V10c0-1.1-.9-2-2-2zM12 17c-1.1 0-2-.9-2-2s.9-2 2-2 2 .9 2 2-.9 2-2 2zM15.1 8H8.9V6c0-1.71 1.39-3.1 3.1-3.1s3.1 1.39 3.1 3.1v2z',
      grey:  { desc: 'Full close out day',              input: null },
      green: { desc: 'No stop sale',                    input: null },
      blue:  { desc: 'At least 1 partial close out',    input: null }
    },
    hotelocc: {
      label: 'Hotel Occupancy',
      icon: '🏨',
      svgPath: 'M7 13c1.66 0 3-1.34 3-3S8.66 7 7 7s-3 1.34-3 3 1.34 3 3 3zm12-6h-8v7H3V5H1v15h2v-3h18v3h2v-9c0-2.21-1.79-4-4-4z',
      grey:  { desc: 'Occupancy above (%)',  input: { param: 'greyT',  def: 85, unit: '%' } },
      green: { desc: 'Occupancy below (%)',  input: { param: 'greenT', def: 60, unit: '%' } },
      blue:  { desc: 'Between Grey & Green thresholds', input: null }
    },
    remaining: {
      label: 'Remaining Rooms',
      icon: '🛏',
      svgPath: 'M19 7h-8v7H3V5H1v15h2v-3h18v3h2V11c0-2.21-1.79-4-4-4z',
      grey:  { desc: 'Remaining rooms less than',  input: { param: 'greyT',  def: 10, unit: 'rms', allowUnitToggle: true } },
      green: { desc: 'Remaining rooms more than',  input: { param: 'greenT', def: 50, unit: 'rms', allowUnitToggle: true } },
      blue:  { desc: 'Between Grey & Green thresholds', input: null }
    },
    mealplan: {
      label: 'Meal Plan Guests',
      icon: '🍽',
      svgPath: 'M8.1 13.34l2.83-2.83L3.91 3.5c-1.56 1.56-1.56 4.09 0 5.66l4.19 4.18zm6.78-1.81c1.53.71 3.68.21 5.27-1.38 1.91-1.91 2.28-4.65.81-6.12-1.46-1.46-4.2-1.1-6.12.81-1.59 1.59-2.09 3.74-1.38 5.27L3.7 19.87l1.41 1.41L12 14.41l6.88 6.88 1.41-1.41L13.41 13l1.47-1.47z',
      grey:  { desc: 'Total guests above', input: { param: 'greyT',  def: 200, unit: 'guests' } },
      green: { desc: 'Total guests below', input: { param: 'greenT', def: 100, unit: 'guests' } },
      blue:  { desc: 'Between Grey & Green thresholds', input: null }
    },
    toforecast: {
      label: 'TO Forecast',
      icon: '📊',
      svgPath: 'M3.5 18.49l6-6.01 4 4L22 6.92l-1.41-1.41-7.09 7.97-4-4L2 16.99z',
      grey:  { desc: 'OTB exceeds the forecast by', input: { param: 'greyT',  def: 20, unit: 'rms', allowUnitToggle: true } },
      green: { desc: 'OTB is below the forecast by', input: { param: 'greenT', def: 20, unit: 'rms', allowUnitToggle: true } },
      blue:  { desc: 'OTB within forecast variance', input: null }
    }
  };

  // ── Open / Close ───────────────────────────────────────────────
  window.hmToggle = function() {
    var m = document.getElementById('hmModal');
    if (!m) return;
    var open = m.style.display === 'flex';
    m.style.display = open ? 'none' : 'flex';
    var btn = document.getElementById('hmBtn');
    if (btn) btn.classList.toggle('active', !open);
    if (!open && hmState.type) hmRenderColours(hmState.type);
  };

  window.hmModalBg = function(e) {
    if (e.target.id === 'hmModal') hmToggle();
  };

  // ── Select heatmap type ────────────────────────────────────────
  window.hmSelectType = function(type) {
    hmState.type = type;
    // Update card active state
    document.querySelectorAll('.hm-type-option').forEach(function(c) {
      c.classList.toggle('active', c.dataset.hmtype === type);
    });
    hmRenderColours(type);
  };

  // ── Render colour rows for selected type ──────────────────────
  function hmRenderColours(type) {
    var def = HM_TYPES[type];
    if (!def) return;
    var section = document.getElementById('hmColourSection');
    var rows    = document.getElementById('hmColourRows');
    if (!section || !rows) return;
    section.style.display = '';

    // Show condition section only for Stop Sales
    var condSection = document.getElementById('hmConditionSection');
    if (condSection) condSection.style.display = type === 'stopsales' ? 'block' : 'none';
    if (type === 'stopsales') {
      var condCb = document.getElementById('hmCondEnabled');
      if (condCb) condCb.checked = hmState.condition.enabled;
      var condCtrls = document.getElementById('hmCondControls');
      if (condCtrls) condCtrls.style.display = hmState.condition.enabled ? 'block' : 'none';
      var condMetric = document.getElementById('hmCondMetric');
      if (condMetric) condMetric.value = hmState.condition.metric;
      var condOp = document.getElementById('hmCondOp');
      if (condOp) condOp.value = hmState.condition.op;
      var condVal = document.getElementById('hmCondValue');
      if (condVal) condVal.value = hmState.condition.value;
    }

    // Figma 2026 swatches
    var isStopSales  = type === 'stopsales';
    var isToForecast = type === 'toforecast';
    var colours = [
      { key: 'grey',  swatch: isStopSales ? '#D33030' : isToForecast ? '#F97316' : '#D9D9D9', label: isStopSales ? 'Closed'  : isToForecast ? 'Above Forecast' : 'Grey',  cfg: def.grey  },
      { key: 'blue',  swatch: isStopSales ? '#FDCF61' : isToForecast ? '#D7F7ED' : '#D7F7ED', label: isStopSales ? 'Partial' : isToForecast ? 'Within Range'   : 'Blue',  cfg: def.blue  },
      { key: 'green', swatch: isStopSales ? '#CEF2D1' : '#388C3F',                            label: isStopSales ? 'Open'    : isToForecast ? 'Below Forecast'  : 'Green', cfg: def.green }
    ];

    rows.innerHTML = colours.map(function(c) {
      var currentClr = hmState.colors[c.key] || c.swatch;
      var bodyHtml = '';
      if (c.cfg.input) {
        var val = hmState[c.key][c.cfg.input.param] !== undefined
          ? hmState[c.key][c.cfg.input.param]
          : c.cfg.input.def;
        var curUnit = (c.cfg.input.allowUnitToggle && hmState[c.key].unitType) ? hmState[c.key].unitType : c.cfg.input.unit;
        var unitToggleHtml = c.cfg.input.allowUnitToggle
          ? '<select class="hm-unit-select" onchange="hmUnitChange(\'' + c.key + '\',this.value)">'
            + '<option value="rms"' + (curUnit === 'rms' ? ' selected' : '') + '>rms</option>'
            + '<option value="%"'   + (curUnit === '%'   ? ' selected' : '') + '>%</option>'
            + '</select>'
          : '<span class="hm-unit-label">' + c.cfg.input.unit + '</span>';
        bodyHtml = '<div class="hm-threshold-body">'
          + '<div class="hm-threshold-name">' + c.label + '</div>'
          + '<div class="hm-threshold-field">'
          + '<div class="hm-threshold-field-label">' + c.cfg.desc + '</div>'
          + '<div style="display:flex;align-items:center;gap:6px">'
          + '<input type="number" class="hm-input" min="0" max="9999" value="' + val + '"'
          + ' data-hm-color="' + c.key + '" data-hm-param="' + c.cfg.input.param + '"'
          + ' onchange="hmParamChange(this)" placeholder="' + c.cfg.input.def + '">'
          + unitToggleHtml
          + '</div>'
          + '</div></div>';
      } else {
        bodyHtml = '<div class="hm-threshold-body">'
          + '<div class="hm-threshold-name">' + c.label + '</div>'
          + '<div class="hm-threshold-between">' + c.cfg.desc + '</div>'
          + '</div>';
      }
      var rowCls = 'hm-threshold-row' + (!c.cfg.input ? ' hm-threshold-no-input' : '');
      return '<div class="' + rowCls + '">'
        + '<div class="hm-threshold-swatch hm-swatch-pick" style="background:' + currentClr + ';cursor:pointer;position:relative" title="Click to change colour">'
        + '<input type="color" value="' + currentClr + '" data-hm-swatch="' + c.key + '"'
        + ' onchange="hmSwatchChange(this)" oninput="hmSwatchChange(this)"'
        + ' style="position:absolute;inset:0;width:100%;height:100%;opacity:0;cursor:pointer">'
        + '</div>'
        + bodyHtml
        + '</div>';
    }).join('');

    // Stop Sales room type selector + chips
    var rtSection = document.getElementById('hmRtSection');
    if (rtSection) {
      if (isStopSales) {
        rtSection.style.display = '';
        var rtOpts = ['Standard','Superior','Deluxe','Suite','Jr. Suite','Family'];
        var ddList = document.getElementById('hmRtDDList');
        if (ddList) {
          var sel = hmState.stopSalesRoomTypes || [];
          ddList.innerHTML = rtOpts.map(function(rt) {
            var isOn = sel.indexOf(rt) >= 0;
            return '<label class="hm-rt-dd-item"><input type="checkbox" value="' + rt + '"' + (isOn ? ' checked' : '') + ' onchange="hmRtItemChange()">' + rt + '</label>';
          }).join('');
        }
        hmRtUpdateTrigger();
        hmRtRenderChips();
      } else {
        rtSection.style.display = 'none';
      }
    }
    // Hide old rtFilter if it exists
    var rtFilterEl = document.getElementById('hmRtFilter');
    if (rtFilterEl) { rtFilterEl.innerHTML = ''; rtFilterEl.style.display = 'none'; }
  }

  // ── Param change ───────────────────────────────────────────────
  window.hmParamChange = function(el) {
    var color = el.dataset.hmColor;
    var param = el.dataset.hmParam;
    if (!hmState[color]) hmState[color] = { params: {} };
    hmState[color][param] = parseFloat(el.value) || 0;
  };

  window.hmUnitChange = function(colorKey, unit) {
    if (!hmState[colorKey]) hmState[colorKey] = { params: {} };
    hmState[colorKey].unitType = unit;
    hmRenderColours(hmState.type);
  };

  // ── Swatch colour change ────────────────────────────────────────
  window.hmSwatchChange = function(el) {
    var key = el.dataset.hmSwatch;
    hmState.colors[key] = el.value;
    // Update swatch preview immediately
    var swatch = el.parentElement;
    if (swatch) swatch.style.background = el.value;
  };

  // ── Room type selector (Figma style) ──────────────────────────
  window.hmRtToggleDD = function() {
    var box = document.getElementById('hmRtSelectBox');
    var list = document.getElementById('hmRtDDList');
    if (!box || !list) return;
    var isOpen = list.classList.contains('open');
    if (isOpen) {
      list.classList.remove('open');
      box.classList.remove('open');
    } else {
      list.classList.add('open');
      box.classList.add('open');
    }
  };

  // Keep heatmap room-type dropdown open when clicking inside it
  (function() {
    var list = document.getElementById('hmRtDDList');
    if (list) list.addEventListener('click', function(e) { e.stopPropagation(); });
  })();

  // Close dropdown when clicking outside
  document.addEventListener('click', function(e) {
    var wrap = document.querySelector('.hm-rt-selector-wrap');
    if (wrap && !wrap.contains(e.target)) {
      var list = document.getElementById('hmRtDDList');
      var box = document.getElementById('hmRtSelectBox');
      if (list) list.classList.remove('open');
      if (box) box.classList.remove('open');
    }
  });

  window.hmRtItemChange = function() {
    hmRtReadSelect();
    hmRtUpdateTrigger();
    hmRtRenderChips();
  };

  function hmRtUpdateTrigger() {
    var sel = hmState.stopSalesRoomTypes || [];
    var txt = document.getElementById('hmRtSelectText');
    if (!txt) return;
    txt.textContent = sel.length === 0 ? 'All' : sel.length + ' selected';
  }

  function hmRtRenderChips() {
    var el = document.getElementById('hmRtChips');
    if (!el) return;
    var sel = hmState.stopSalesRoomTypes || [];
    if (sel.length === 0) { el.innerHTML = ''; return; }
    el.innerHTML = sel.map(function(rt) {
      return '<span class="hm-rt-chip">' + rt + '</span>';
    }).join('');
  }

  window.hmRtReadSelect = function() {
    var ddList = document.getElementById('hmRtDDList');
    if (!ddList) return;
    var arr = [];
    ddList.querySelectorAll('input[type="checkbox"]').forEach(function(cb) {
      if (cb.checked) arr.push(cb.value);
    });
    hmState.stopSalesRoomTypes = arr;
  };

  // ── Condition toggle ───────────────────────────────────────────
  window.hmCondToggle = function(cb) {
    var ctrls = document.getElementById('hmCondControls');
    if (ctrls) ctrls.style.display = cb.checked ? 'block' : 'none';
  };

  // ── Apply ──────────────────────────────────────────────────────
  window.hmApply = function() {
    // Read any number inputs
    document.querySelectorAll('#hmModal .hm-input').forEach(function(el) {
      var c = el.dataset.hmColor, p = el.dataset.hmParam;
      if (c && p) hmState[c][p] = parseFloat(el.value) || 0;
    });
    hmState.enabled = !!hmState.type;
    // Read room types from select
    hmRtReadSelect();
    // Read condition
    var condCb     = document.getElementById('hmCondEnabled');
    var condMetric = document.getElementById('hmCondMetric');
    var condOp     = document.getElementById('hmCondOp');
    var condValEl  = document.getElementById('hmCondValue');
    hmState.condition = {
      enabled: !!(condCb && condCb.checked),
      metric:  condMetric ? condMetric.value : 'hotel',
      op:      condOp     ? condOp.value     : '>',
      value:   condValEl  ? (parseFloat(condValEl.value) || 0) : 50
    };
    // Apply custom heatmap colours as CSS variables
    hmApplyColors();
    // Update button to reflect active heatmap type
    hmUpdateBtn();
    hmToggle();
    renderCalendar();
  };

  function hmApplyColors() {
    var grid = document.getElementById('calMonths');
    if (!grid) return;
    var isStopSales = hmState.type === 'stopsales';
    var defaults = isStopSales
      ? { grey: '#D33030', blue: '#FDCF61', green: '#CEF2D1' }
      : { grey: '#D9D9D9', blue: '#D7F7ED', green: '#388C3F' };
    var gc = hmState.colors.grey  || defaults.grey;
    var bc = hmState.colors.blue  || defaults.blue;
    var gnc = hmState.colors.green || defaults.green;
    grid.style.setProperty('--hm-grey-bg',  gc + '30');
    grid.style.setProperty('--hm-grey-bdr', gc);
    grid.style.setProperty('--hm-blue-bg',  bc + '30');
    grid.style.setProperty('--hm-blue-bdr', bc);
    grid.style.setProperty('--hm-green-bg', gnc + '30');
    grid.style.setProperty('--hm-green-bdr',gnc);
    // Stop sales specific aliases
    grid.style.setProperty('--hm-closed-bg',  gc + '40');
    grid.style.setProperty('--hm-closed-bdr', gc);
    grid.style.setProperty('--hm-partial-bg', bc + '30');
    grid.style.setProperty('--hm-partial-bdr',bc);
    grid.style.setProperty('--hm-open-bg',    gnc + '30');
    grid.style.setProperty('--hm-open-bdr',   gnc);
  }

  function hmUpdateBtn() {
    var iconEl   = document.getElementById('hmBtnIcon');
    var defaultIconEl = document.getElementById('hmBtnDefaultIcon');
    var labelEl  = document.getElementById('hmBtnLabel');
    var btn      = document.getElementById('hmBtn');
    if (!labelEl) return;
    if (hmState.enabled && hmState.type && HM_TYPES[hmState.type]) {
      var def = HM_TYPES[hmState.type];
      // Show SVG icon inside a hollow chip
      if (iconEl) {
        iconEl.innerHTML = '<span class="hm-btn-chip"><svg viewBox="0 0 24 24" fill="#006461" width="14" height="14"><path d="' + def.svgPath + '"/></svg></span>';
        iconEl.style.display = 'inline-flex';
      }
      if (defaultIconEl) defaultIconEl.style.display = 'none';
      labelEl.textContent = 'Heatmap';
      if (btn) btn.classList.add('active');
    } else {
      // Restore defaults
      if (iconEl)        { iconEl.innerHTML = ''; iconEl.style.display = 'none'; }
      if (defaultIconEl) defaultIconEl.style.display = '';
      labelEl.textContent = 'Heatmap';
      if (btn) btn.classList.remove('active');
    }
  }

  // ── Reset ──────────────────────────────────────────────────────
  window.hmReset = function() {
    hmState = { type: '', grey: { params:{} }, green: { params:{} }, blue: { params:{} }, enabled: false,
                condition: { enabled: false, metric: 'hotel', op: '>', value: 50 }, stopSalesRoomTypes: [], colors: {} };
    // Clear custom color CSS variables
    var grid = document.getElementById('calMonths');
    if (grid) ['--hm-grey-bg','--hm-grey-bdr','--hm-blue-bg','--hm-blue-bdr','--hm-green-bg','--hm-green-bdr',
               '--hm-closed-bg','--hm-closed-bdr','--hm-partial-bg','--hm-partial-bdr','--hm-open-bg','--hm-open-bdr']
      .forEach(function(v){ grid.style.removeProperty(v); });
    document.querySelectorAll('.hm-type-option').forEach(function(c) { c.classList.remove('active'); });
    var section = document.getElementById('hmColourSection');
    if (section) section.style.display = 'none';
    var rows = document.getElementById('hmColourRows');
    if (rows) rows.innerHTML = '';
    var condSection = document.getElementById('hmConditionSection');
    if (condSection) condSection.style.display = 'none';
    var condCb = document.getElementById('hmCondEnabled');
    if (condCb) condCb.checked = false;
    var condCtrls = document.getElementById('hmCondControls');
    if (condCtrls) condCtrls.style.display = 'none';
    var rtFilter = document.getElementById('hmRtFilter');
    if (rtFilter) { rtFilter.innerHTML = ''; rtFilter.style.display = 'none'; }
    var rtSect = document.getElementById('hmRtSection');
    if (rtSect) rtSect.style.display = 'none';
    var rtDDList = document.getElementById('hmRtDDList');
    if (rtDDList) rtDDList.innerHTML = '';
    var rtBox = document.getElementById('hmRtSelectBox');
    if (rtBox) rtBox.classList.remove('open');
    var rtSelectText = document.getElementById('hmRtSelectText');
    if (rtSelectText) rtSelectText.textContent = 'All';
    var rtChips = document.getElementById('hmRtChips');
    if (rtChips) rtChips.innerHTML = '';
    hmUpdateBtn();
    renderCalendar();
  };

  // ── Get cell colour class ──────────────────────────────────────
  window.hmIsStopSales = function() {
    return hmState.enabled && hmState.type === 'stopsales';
  };

  window.hmGetCellClass = function(dayData) {
    if (!hmState.enabled || !hmState.type) return '';

    // Condition gate — skip colouring if condition not met
    if (hmState.condition.enabled) {
      var cond = hmState.condition;
      var mval;
      switch (cond.metric) {
        case 'hotel':       mval = dayData.hotel;       break;
        case 'remainRooms': mval = dayData.remainRooms; break;
        case 'totalGuests': mval = dayData.totalGuests; break;
        case 'toOtb':       mval = dayData.toOtb;       break;
        default:            mval = 0;
      }
      var pass = false;
      switch (cond.op) {
        case '>':  pass = mval >  cond.value; break;
        case '>=': pass = mval >= cond.value; break;
        case '<':  pass = mval <  cond.value; break;
        case '<=': pass = mval <= cond.value; break;
      }
      if (!pass) return '';
    }

    var type = hmState.type;
    var gT  = parseFloat(hmState.grey.greyT)   || 0;
    var gnT = parseFloat(hmState.green.greenT) || 0;

    // Stop sales room type helpers (multiselect)
    function ssRtClosed() {
      var rts = hmState.stopSalesRoomTypes;
      if (!rts || rts.length === 0) return false; // no room type filter (All)
      var rules = dayData.closureRules || [];
      // Check if ALL selected room types are closed
      for (var ri = 0; ri < rts.length; ri++) {
        var found = false;
        for (var i = 0; i < rules.length; i++) {
          var r = rules[i];
          if (r.roomTypes.length === 0 || r.roomTypes.indexOf(rts[ri]) >= 0) { found = true; break; }
        }
        if (!found) return false; // at least one selected RT is not closed
      }
      return true; // all selected RTs are closed
    }

    function testGrey() {
      if (type === 'stopsales') {
        if (dayData.isFullClose) return true;
        if (hmState.stopSalesRoomTypes && hmState.stopSalesRoomTypes.length) return ssRtClosed();
        return false;
      }
      if (type === 'hotelocc')   return dayData.hotel >= gT;
      if (type === 'remaining')  return dayData.remainRooms < gT;
      if (type === 'mealplan')   return dayData.totalGuests >= gT;
      if (type === 'toforecast') return (dayData.toOtb - dayData.toFcst) >= gT;
      return false;
    }
    function testGreen() {
      if (type === 'stopsales') {
        if (dayData.isFullClose) return false;
        if (hmState.stopSalesRoomTypes && hmState.stopSalesRoomTypes.length) return !ssRtClosed();
        return !dayData.hasPartialClose;
      }
      if (type === 'hotelocc')   return dayData.hotel < gnT;
      if (type === 'remaining')  return dayData.remainRooms > gnT;
      if (type === 'mealplan')   return dayData.totalGuests < gnT;
      if (type === 'toforecast') return (dayData.toFcst - dayData.toOtb) >= gnT;
      return false;
    }
    function testBlue() {
      if (type === 'stopsales') {
        if (dayData.isFullClose) return false;
        if (hmState.stopSalesRoomTypes && hmState.stopSalesRoomTypes.length) return false; // RT filter → binary closed/open only
        return dayData.hasPartialClose;
      }
      // Between grey and green for threshold types
      var lo = Math.min(gT, gnT), hi = Math.max(gT, gnT);
      if (type === 'hotelocc')   return dayData.hotel >= lo && dayData.hotel <= hi;
      if (type === 'remaining')  return dayData.remainRooms >= lo && dayData.remainRooms <= hi;
      if (type === 'mealplan')   return dayData.totalGuests >= lo && dayData.totalGuests <= hi;
      if (type === 'toforecast') { var diff = Math.abs(dayData.toOtb - dayData.toFcst); return diff >= lo && diff <= hi; }
      return false;
    }

    // Stop Sales uses semantic colour classes; other types use generic ones
    if (type === 'stopsales') {
      if (testGrey())  return 'hm-closed';
      if (testBlue())  return 'hm-partial';
      if (testGreen()) return 'hm-open';
      return '';
    }
    if (testGrey())  return 'hm-grey';
    if (testGreen()) return 'hm-green';
    if (testBlue())  return 'hm-blue';
    return '';
  };

})();

/* ═══════════════════════════════════════════
   DS SNACKBAR
   ═══════════════════════════════════════════ */
var _snackbarTimer = null;
window.dsSnackbarShow = function(msg) {
  var el = document.getElementById('dsSnackbar');
  if (!el) return;
  var msgEl = el.querySelector('.ds-snackbar-msg');
  if (msgEl) msgEl.textContent = msg || 'Saved successfully';
  el.classList.add('show');
  clearTimeout(_snackbarTimer);
  _snackbarTimer = setTimeout(function() { el.classList.remove('show'); }, 3500);
};
window.dsSnackbarHide = function() {
  var el = document.getElementById('dsSnackbar');
  if (el) el.classList.remove('show');
  clearTimeout(_snackbarTimer);
};

/* ═══════════════════════════════════════════
   CONFIG DIRTY TRACKING
   ═══════════════════════════════════════════ */
var _configDirty = false;

window._markConfigDirty = function() {
  if (_configDirty) return;
  _configDirty = true;
  var footer = document.getElementById('configSaveFooter');
  if (footer) footer.classList.add('show');
  ['footerSaveBtn','btSaveBtn'].forEach(function(id) {
    var b = document.getElementById(id); if (b) b.disabled = false;
  });
};

window._configSetClean = function() {
  _configDirty = false;
  var footer = document.getElementById('configSaveFooter');
  if (footer) footer.classList.remove('show');
  ['footerSaveBtn','btSaveBtn'].forEach(function(id) {
    var b = document.getElementById(id); if (b) b.disabled = true;
  });
};

window._configDiscard = function() {
  _configSetClean();
};

/* ── Page-level save ── */
window.configPageSave = function() {
  // Stop any active AG Grid cell edit across all grids
  if (typeof _btGridApi !== 'undefined' && _btGridApi) _btGridApi.stopEditing();
  if (typeof window._srGridApi === 'function' && window._srGridApi()) window._srGridApi().stopEditing();
  dsSnackbarShow('Saved successfully');
  _configSetClean();
};

// Wire up dirty listeners — covers every input/select/checkbox/radio
// across ALL config panels in one sweep
setTimeout(function() {
  document.querySelectorAll(
    '#settingsPage input[type="radio"], ' +
    '#settingsPage input[type="checkbox"], ' +
    '#settingsPage select'
  ).forEach(function(el) {
    el.addEventListener('change', _markConfigDirty);
  });
}, 600);

