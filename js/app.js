// AutoDashboard — App-Logik v0.2

// ─── Helpers ──────────────────────────────────────────────────────────────────

const q    = s => document.querySelector(s);
const qAll = s => document.querySelectorAll(s);

// ─── State ────────────────────────────────────────────────────────────────────

const state = {
  tank:     [],
  kosten:   [],
  kmstand:  [],
  fahrzeug: {},
  view:     'dashboard',
  year:     String(new Date().getFullYear()),
  charts:   {}
};

// ─── Formatierung ─────────────────────────────────────────────────────────────

const eur   = v  => Number(v || 0).toLocaleString('de-DE', { style: 'currency', currency: 'EUR' });
const num   = (v, d=2) => Number(v || 0).toLocaleString('de-DE', { minimumFractionDigits: d, maximumFractionDigits: d });
const dat   = v  => v ? new Date(v + 'T12:00:00').toLocaleDateString('de-DE') : '—';
const today = () => new Date().toISOString().split('T')[0];
const COLORS = ['#3b82f6','#22c55e','#a855f7','#f59e0b','#94a3b8','#ef4444','#06b6d4'];
const KAT_ICONS = { 'KFZ-Steuer':'📋', Versicherung:'🛡', Inspektion:'🔧', Reparatur:'⚙️', Sonstiges:'📌' };

// ─── API ──────────────────────────────────────────────────────────────────────

async function apiCall(params) {
  const url = new URL(API_URL);
  for (const [k, v] of Object.entries(params))
    url.searchParams.set(k, typeof v === 'object' ? JSON.stringify(v) : String(v));
  const res = await fetch(url.toString());
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  const json = await res.json();
  if (json.error) throw new Error(json.error);
  return json;
}

const api = {
  read:      sheet              => apiCall({ action: 'read',      sheet }),
  add:       (sheet, data)      => apiCall({ action: 'add',       sheet, data }),
  update:    (sheet, row, data) => apiCall({ action: 'update',    sheet, row, data }),
  deleteAll: sheet              => apiCall({ action: 'deleteAll', sheet }),
};

// ─── Export & Delete ─────────────────────────────────────────────────────────

function exportToExcel() {
  const wb = XLSX.utils.book_new();

  const tankRows = state.tank.map(e => ({
    Datum:            e.datum        || '',
    Tankstelle:       e.tankstelle   || '',
    Kraftstoff:       e.kraftstoff   || '',
    'Liter':          Number(e.liter  || 0),
    'Kosten (€)':     Number(e.kosten || 0),
    '€/Liter':        Number(e.preis_pro_liter || 0),
    KM_Stand:         e.km_stand     || '',
    Konto:            e.konto        || '',
    Karte:            e.karte        || '',
    Beleg:            e.beleg        || '',
    Hinweis:          e.hinweis      || '',
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(tankRows),  'Tanken');

  const kostenRows = state.kosten.map(e => ({
    Datum:               e.datum                || '',
    Kategorie:           e.kategorie            || '',
    Beschreibung:        e.beschreibung         || '',
    'Betrag (€)':        Number(e.betrag        || 0),
    Intervall:           e.intervall            || '',
    Naechste_Faelligkeit: e.naechste_faelligkeit || '',
    Konto:               e.konto               || '',
    Beleg:               e.beleg               || '',
    Hinweis:             e.hinweis             || '',
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(kostenRows), 'Ausgaben');

  const filename = `AutoDashboard_Export_${today()}.xlsx`;
  XLSX.writeFile(wb, filename);
}

function confirmDeleteAll(sheet, label) {
  q('#confirm-msg').textContent   = `Alle ${label}-Einträge unwiderruflich löschen?`;
  q('#confirm-sheet').value       = sheet;
  q('#confirm-label').textContent = label;
  q('#modal-confirm').classList.remove('hidden');
}

// ─── Fälligkeits-Helfer ──────────────────────────────────────────────────────

const INTERVALL_MONTHS = { monatlich: 1, quartalsweise: 3, halbjährlich: 6, jährlich: 12, '2-jährlich': 24 };

function nextDueDate(baseDateStr, intervall) {
  if (!baseDateStr || !intervall || intervall === 'einmalig') return '';
  const m = INTERVALL_MONTHS[intervall];
  if (!m) return '';
  const d = new Date(baseDateStr + 'T12:00:00');
  d.setMonth(d.getMonth() + m);
  return d.toISOString().split('T')[0];
}

async function markTerminErledigt(entry) {
  const btn = q(`[data-erledigt="${entry._row}"]`);
  if (btn) { btn.disabled = true; btn.textContent = '…'; }
  try {
    const newFaelligkeit = nextDueDate(entry.naechste_faelligkeit, entry.intervall);
    await api.update('Kosten', entry._row, { ...entry, naechste_faelligkeit: newFaelligkeit });
    await loadData();
  } catch (err) {
    showError('Konnte Termin nicht aktualisieren: ' + err.message);
    if (btn) { btn.disabled = false; btn.textContent = 'Erledigt'; }
  }
}

// ─── Daten laden ──────────────────────────────────────────────────────────────

async function loadData() {
  setLoading(true);
  clearError();
  try {
    const [t, k, f, km] = await Promise.all([
      api.read('Tanken'),
      api.read('Kosten'),
      api.read('Fahrzeug'),
      api.read('KMStand').catch(() => ({ data: [] }))
    ]);

    state.tank    = (t.data  || []).sort((a, b) => String(b.datum).localeCompare(String(a.datum)));
    state.kosten  = (k.data  || []).sort((a, b) => String(b.datum).localeCompare(String(a.datum)));
    state.kmstand = (km.data || []).sort((a, b) => String(b.datum).localeCompare(String(a.datum)));

    // Fahrzeug: key-value Tabellenblatt → Objekt
    state.fahrzeug = {};
    (f.data || []).forEach(row => {
      const key = Object.values(row)[1]; // erste Spalte nach _row
      const val = Object.values(row)[2];
      if (key) state.fahrzeug[key] = val;
    });

    buildYearFilter();
    setLoading(false);
    renderAll();
  } catch (err) {
    setLoading(false);
    showError('Verbindung fehlgeschlagen: ' + err.message);
  }
}

// ─── Jahresfilter ─────────────────────────────────────────────────────────────

function buildYearFilter() {
  const years = new Set();
  [...state.tank, ...state.kosten].forEach(e => {
    const y = String(e.datum || '').slice(0, 4);
    if (y && y.length === 4) years.add(y);
  });

  const selectors = [q('#year-filter'), q('#year-filter-mobile')].filter(Boolean);
  const current   = selectors[0]?.value || '';
  selectors.forEach(sel => {
    sel.innerHTML = '<option value="">Alle Jahre</option>';
    [...years].sort().reverse().forEach(y => {
      const opt = document.createElement('option');
      opt.value = y;
      opt.textContent = y;
      if (y === current || (!current && y === String(new Date().getFullYear()))) opt.selected = true;
      sel.appendChild(opt);
    });
  });
  state.year = selectors[0]?.value || '';
}

// ─── Filter-Hilfsfunktion ─────────────────────────────────────────────────────

function filtered(arr, dateField = 'datum') {
  if (!state.year) return arr;
  return arr.filter(e => String(e[dateField] || '').startsWith(state.year));
}

// ─── Stats berechnen ──────────────────────────────────────────────────────────

function calcStats() {
  const tank   = filtered(state.tank);
  const kosten = filtered(state.kosten);

  const totalTank   = tank.reduce((s, e) => s + Number(e.kosten || 0), 0);
  const totalLiter  = tank.reduce((s, e) => s + Number(e.liter  || 0), 0);
  const totalKosten = kosten.reduce((s, e) => s + Number(e.betrag || 0), 0);
  const totalAll    = totalTank + totalKosten;
  const avgPreis    = totalLiter > 0 ? totalTank / totalLiter : 0;

  // Letzter KM-Stand (aus dediziertem KMStand-Sheet)
  const kmSorted = [...state.kmstand].sort((a, b) => String(b.datum).localeCompare(String(a.datum)));
  const lastKm   = kmSorted.length ? Number(kmSorted[0].km_stand) : 0;
  const lastKmDat = kmSorted.length ? kmSorted[0].datum : null;

  // KM diesen Monat (aus KMStand)
  const monatStr = today().slice(0, 7);
  const kmMonat = (() => {
    const thisMonth = state.kmstand.filter(e => String(e.datum).startsWith(monatStr));
    const before    = state.kmstand.filter(e => !String(e.datum).startsWith(monatStr))
      .sort((a, b) => String(b.datum).localeCompare(String(a.datum)));
    if (!thisMonth.length && !before.length) return 0;
    const maxKm = thisMonth.length ? Math.max(...thisMonth.map(e => Number(e.km_stand || 0))) : 0;
    const minKm = before.length
      ? Number(before[0].km_stand || 0)
      : Math.min(...thisMonth.map(e => Number(e.km_stand || 0)));
    return Math.max(0, maxKm - minKm);
  })();

  // Monatliche Tankkosten (letzte 6 Monate)
  const monthly = {};
  state.tank.forEach(e => {
    const m = String(e.datum).slice(0, 7);
    if (m) monthly[m] = (monthly[m] || 0) + Number(e.kosten || 0);
  });
  const last6 = Array.from({ length: 6 }, (_, i) => {
    const d = new Date(); d.setMonth(d.getMonth() - (5 - i));
    return d.toISOString().slice(0, 7);
  });

  // Kosten nach Kategorie (für Donut)
  const katMap = { Tanken: totalTank };
  kosten.forEach(e => {
    const k = e.kategorie || 'Sonstiges';
    katMap[k] = (katMap[k] || 0) + Number(e.betrag || 0);
  });

  // Kraftstoff-Verteilung
  const kraftstoffMap = {};
  tank.forEach(e => {
    const k = e.kraftstoff || 'Unbekannt';
    kraftstoffMap[k] = (kraftstoffMap[k] || 0) + Number(e.liter || 0);
  });

  // Jahres-Tankkosten für Statistik
  const yearMap = {};
  state.tank.forEach(e => {
    const y = String(e.datum).slice(0, 4);
    if (y) yearMap[y] = (yearMap[y] || 0) + Number(e.kosten || 0);
  });

  return { tank, kosten, totalTank, totalLiter, totalKosten, totalAll,
           avgPreis, lastKm, lastKmDat, kmMonat, monthly, last6, katMap, kraftstoffMap, yearMap };
}

// ─── Render Übersicht ─────────────────────────────────────────────────────────

const KAT_SVG = {
  'Tanken':       `<svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M19.428 15.428a2 2 0 00-1.022-.547l-2.387-.477a6 6 0 00-3.86.517l-.318.158a6 6 0 01-3.86.517L6.05 15.21a2 2 0 00-1.806.547M8 4h8l-1 1v5.172a2 2 0 00.586 1.414l5 5c1.26 1.26.367 3.414-1.415 3.414H4.828c-1.782 0-2.674-2.154-1.414-3.414l5-5A2 2 0 009 10.172V5L8 4z"/></svg>`,
  'Versicherung': `<svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M9 12l2 2 4-4m5.618-4.016A11.955 11.955 0 0112 2.944a11.955 11.955 0 01-8.618 3.04A12.02 12.02 0 003 9c0 5.591 3.824 10.29 9 11.622 5.176-1.332 9-6.03 9-11.622 0-1.042-.133-2.052-.382-3.016z"/></svg>`,
  'Inspektion':   `<svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M10.325 4.317c.426-1.756 2.924-1.756 3.35 0a1.724 1.724 0 002.573 1.066c1.543-.94 3.31.826 2.37 2.37a1.724 1.724 0 001.065 2.572c1.756.426 1.756 2.924 0 3.35a1.724 1.724 0 00-1.066 2.573c.94 1.543-.826 3.31-2.37 2.37a1.724 1.724 0 00-2.572 1.065c-.426 1.756-2.924 1.756-3.35 0a1.724 1.724 0 00-2.573-1.066c-1.543.94-3.31-.826-2.37-2.37a1.724 1.724 0 00-1.065-2.572c-1.756-.426-1.756-2.924 0-3.35a1.724 1.724 0 001.066-2.573c-.94-1.543.826-3.31 2.37-2.37.996.608 2.296.07 2.572-1.065z"/><circle cx="12" cy="12" r="3"/></svg>`,
  'Reparatur':    `<svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z"/></svg>`,
  'KFZ-Steuer':   `<svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M9 7h6m0 10v-3m-3 3h.01M9 17h.01M9 14h.01M12 14h.01M15 11h.01M12 11h.01M9 11h.01M7 21h10a2 2 0 002-2V5a2 2 0 00-2-2H7a2 2 0 00-2 2v14a2 2 0 002 2z"/></svg>`,
  'Parken':       `<svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M8 14v3m4-3v3m4-3v3M3 21h18M3 10h18M3 7l9-4 9 4M4 10h16v11H4V10z"/></svg>`,
  'Waschanlage':  `<svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M19.428 15.428a2 2 0 00-1.022-.547l-2.387-.477a6 6 0 00-3.86.517l-.318.158a6 6 0 01-3.86.517L6.05 15.21a2 2 0 00-1.806.547"/></svg>`,
};
const KAT_COLOR = {
  'Tanken':'text-blue-500 bg-blue-50','Versicherung':'text-emerald-500 bg-emerald-50',
  'Inspektion':'text-orange-500 bg-orange-50','Reparatur':'text-red-500 bg-red-50',
  'KFZ-Steuer':'text-violet-500 bg-violet-50','Parken':'text-gray-500 bg-gray-100',
  'Waschanlage':'text-cyan-500 bg-cyan-50',
};

function renderDashboard() {
  const s   = calcStats();
  const cur = state.year || String(new Date().getFullYear());
  const prv = String(Number(cur) - 1);

  // ── KPI: Kilometerstand ──
  q('#kpi-km').textContent     = s.lastKm ? s.lastKm.toLocaleString('de-DE') + ' km' : '—';
  q('#kpi-km-sub').textContent = s.lastKm ? `Zuletzt: ${dat(s.lastKmDat)}` : 'KM-Stand noch nicht erfasst';

  // ── KPI: Gesamtausgaben + YoY ──
  q('#kpi-gesamt').textContent = eur(s.totalAll);
  const prvAll = [...state.tank, ...state.kosten]
    .filter(e => String(e.datum || '').startsWith(prv))
    .reduce((acc, e) => acc + Number(e.kosten || e.betrag || 0), 0);
  if (prvAll > 0 && s.totalAll > 0) {
    const pct  = Math.round((s.totalAll - prvAll) / prvAll * 100);
    const up   = pct >= 0;
    q('#kpi-gesamt-trend-icon').innerHTML = up
      ? `<svg class="w-3 h-3 text-red-500" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M5 15l7-7 7 7"/></svg>`
      : `<svg class="w-3 h-3 text-green-500" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M19 9l-7 7-7-7"/></svg>`;
    q('#kpi-gesamt-trend').textContent  = `${up ? '+' : ''}${pct}% vs. ${prv}`;
    q('#kpi-gesamt-trend').className    = `text-xs mt-1.5 ${up ? 'text-red-500' : 'text-green-500'}`;
  } else {
    q('#kpi-gesamt-trend').textContent  = `${s.tank.length} Tankungen + ${s.kosten.length} Ausgaben`;
    q('#kpi-gesamt-trend').className    = 'text-xs text-gray-400 mt-1.5';
  }

  // ── KPI: Kosten pro km & Ø Verbrauch (aus KMStand-Sheet) ──
  const kmYear = state.kmstand
    .filter(e => String(e.datum || '').startsWith(cur))
    .sort((a, b) => String(a.datum).localeCompare(String(b.datum)));
  const kmYearEntries = kmYear.length >= 2 ? kmYear : (() => {
    // Fallback: letzten Eintrag vor dem Jahr als Start nehmen
    const before = state.kmstand.filter(e => String(e.datum || '') < cur)
      .sort((a, b) => String(b.datum).localeCompare(String(a.datum)));
    return before.length ? [before[0], ...kmYear] : kmYear;
  })();
  if (kmYearEntries.length >= 2) {
    const kmFirst = Number(kmYearEntries[0].km_stand);
    const kmLast  = Number(kmYearEntries[kmYearEntries.length - 1].km_stand);
    const kmDelta = kmLast - kmFirst;
    if (kmDelta > 0) {
      q('#kpi-kost-km').textContent       = num(s.totalAll / kmDelta, 2) + ' €/km';
      q('#kpi-kost-km-sub').textContent   = `${kmDelta.toLocaleString('de-DE')} km in ${cur}`;
      q('#kpi-verbrauch').textContent     = num(s.totalLiter / kmDelta * 100, 1) + ' l';
      q('#kpi-verbrauch-sub').textContent = 'l / 100 km';
    }
  }

  // ── Letztes Tanken (letzter Eintrag mit tatsächlichen Daten bevorzugt) ──
  const allTankSorted = [...state.tank].sort((a,b) => String(b.datum).localeCompare(String(a.datum)));
  const ltWithData    = allTankSorted.find(e => Number(e.liter) > 0 && Number(e.kosten) > 0) || allTankSorted[0];
  if (allTankSorted.length > 0) {
    const lt    = ltWithData;
    const liter = Number(lt.liter) || 0;
    const kost  = Number(lt.kosten) || 0;
    const preis = liter > 0 ? kost / liter : 0;
    q('#lt-liter').textContent  = liter  > 0 ? num(liter)  + ' l'    : '—';
    q('#lt-preis').textContent  = preis  > 0 ? num(preis, 3) + ' €/l' : '—';
    q('#lt-kosten').textContent = kost   > 0 ? eur(kost)               : '—';
    q('#lt-datum').textContent  = dat(lt.datum);
  }

  // ── KM diesen Monat ──
  if (s.kmMonat > 0) {
    q('#km-monat-bar').classList.remove('hidden');
    q('#km-monat-val').textContent = s.kmMonat.toLocaleString('de-DE');
  } else {
    q('#km-monat-bar').classList.add('hidden');
  }

  renderBarChart(s);
  renderDonutChart(s);
  renderUpcomingList();
  renderAusgabenUeberblick(s);
  renderBottomKpis(s);
  renderVehicleSidebar();
}

// ─── Bar Chart ────────────────────────────────────────────────────────────────

function renderBarChart(s) {
  const labels = s.last6.map(m => new Date(m + '-01T12:00:00').toLocaleDateString('de-DE', { month: 'short' }));
  const data   = s.last6.map(m => Math.round((s.monthly[m] || 0) * 100) / 100);

  if (state.charts.bar) state.charts.bar.destroy();
  state.charts.bar = new Chart(q('#barChart'), {
    type: 'bar',
    data: { labels, datasets: [{ data, backgroundColor: '#3b82f6', borderRadius: 5, borderSkipped: false }] },
    options: {
      responsive: true, maintainAspectRatio: false,
      layout: { padding: { top: 18 } },
      plugins: { legend: { display: false }, tooltip: { callbacks: { label: ctx => ' ' + eur(ctx.parsed.y) } } },
      scales: {
        x: { grid: { display: false }, border: { display: false }, ticks: { color: '#94a3b8', font: { size: 10 } } },
        y: { display: false }
      }
    },
    plugins: [{
      id: 'barLabels',
      afterDatasetsDraw(chart) {
        const { ctx, data: d } = chart;
        d.datasets.forEach((ds, i) => {
          chart.getDatasetMeta(i).data.forEach((bar, idx) => {
            const v = ds.data[idx];
            if (!v) return;
            ctx.save();
            ctx.font = '600 10px system-ui';
            ctx.fillStyle = '#64748b';
            ctx.textAlign = 'center';
            ctx.textBaseline = 'bottom';
            ctx.fillText(Math.round(v) + ' €', bar.x, bar.y - 2);
            ctx.restore();
          });
        });
      }
    }]
  });
}

// ─── Donut Chart ──────────────────────────────────────────────────────────────

function renderDonutChart(s) {
  const entries = Object.entries(s.katMap).filter(([, v]) => v > 0).sort((a, b) => b[1] - a[1]);
  const total   = entries.reduce((sum, [, v]) => sum + v, 0);

  if (state.charts.donut) state.charts.donut.destroy();
  state.charts.donut = new Chart(q('#donutChart'), {
    type: 'doughnut',
    data: { labels: entries.map(([k]) => k), datasets: [{ data: entries.map(([, v]) => v), backgroundColor: COLORS, borderWidth: 0, hoverOffset: 4 }] },
    options: {
      responsive: true, maintainAspectRatio: false, cutout: '72%',
      plugins: { legend: { display: false }, tooltip: { callbacks: { label: ctx => ` ${ctx.label}: ${eur(ctx.parsed)}` } } }
    },
    plugins: [{
      id: 'centerText',
      afterDraw(chart) {
        const { ctx, chartArea: { top, left, width, height } } = chart;
        ctx.save();
        ctx.font = 'bold 14px system-ui'; ctx.fillStyle = '#1e293b';
        ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
        ctx.fillText(eur(total), left + width / 2, top + height / 2 - 8);
        ctx.font = '11px system-ui'; ctx.fillStyle = '#94a3b8';
        ctx.fillText('Gesamt', left + width / 2, top + height / 2 + 10);
        ctx.restore();
      }
    }]
  });

  q('#donut-legend').innerHTML = entries.map(([k, v], i) => `
    <div class="flex items-center justify-between">
      <div class="flex items-center gap-2">
        <span class="w-2.5 h-2.5 rounded-full shrink-0" style="background:${COLORS[i]}"></span>
        <span class="text-xs text-gray-600">${k}</span>
      </div>
      <div class="text-right">
        <span class="text-xs font-semibold text-gray-700">${eur(v)}</span>
        <span class="text-xs text-gray-400 ml-1">${total > 0 ? Math.round(v/total*100) : 0}%</span>
      </div>
    </div>`).join('');
}

// ─── Upcoming Termine ─────────────────────────────────────────────────────────

function renderUpcomingList() {
  const upcoming = state.kosten
    .filter(e => e.naechste_faelligkeit)
    .sort((a, b) => String(a.naechste_faelligkeit).localeCompare(String(b.naechste_faelligkeit)))
    .slice(0, 5);

  const el = q('#upcoming-list');
  if (!upcoming.length) {
    el.innerHTML = '<p class="text-sm text-gray-400 text-center py-6">Keine Fälligkeiten eingetragen</p>';
    return;
  }
  el.innerHTML = upcoming.map(e => {
    const days    = Math.ceil((new Date(e.naechste_faelligkeit + 'T12:00:00') - new Date()) / 86400000);
    const overdue = days < 0;
    const soon    = !overdue && days <= 60;
    const badge   = overdue
      ? `<span class="text-xs px-2.5 py-1 rounded-full bg-red-100 text-red-600 font-semibold whitespace-nowrap">Überfällig</span>`
      : soon
        ? `<span class="text-xs px-2.5 py-1 rounded-full bg-orange-100 text-orange-600 font-semibold whitespace-nowrap">fällig bald</span>`
        : `<span class="text-xs px-2.5 py-1 rounded-full bg-blue-50 text-blue-600 font-semibold whitespace-nowrap">in ${days} Tagen</span>`;
    const col  = KAT_COLOR[e.kategorie] || 'text-gray-500 bg-gray-100';
    const icon = KAT_SVG[e.kategorie]   || KAT_SVG['Inspektion'];
    return `
      <div class="flex items-center gap-3 py-2 border-b border-slate-50 last:border-0">
        <div class="w-8 h-8 rounded-xl flex items-center justify-center shrink-0 ${col}">${icon}</div>
        <div class="flex-1 min-w-0">
          <p class="text-sm font-semibold text-gray-800 truncate">${e.beschreibung || e.kategorie}</p>
          <p class="text-xs text-gray-400">${dat(e.naechste_faelligkeit)}</p>
        </div>
        <div class="flex items-center gap-2 shrink-0">
          ${badge}
        </div>
      </div>`;
  }).join('');
}

// ─── Ausgaben im Überblick ────────────────────────────────────────────────────

function renderAusgabenUeberblick(s) {
  const all = [
    ...s.tank.slice(0, 10).map(e => ({ datum: e.datum, kat: 'Tanken', label: e.tankstelle || 'Tankstelle', betrag: Number(e.kosten || 0) })),
    ...s.kosten.slice(0, 10).map(e => ({ datum: e.datum, kat: e.kategorie || 'Sonstiges', label: e.beschreibung || e.kategorie, betrag: Number(e.betrag || 0) }))
  ].sort((a, b) => String(b.datum).localeCompare(String(a.datum))).slice(0, 7);

  q('#ausgaben-uberblick').innerHTML = all.length ? all.map(e => {
    const col  = KAT_COLOR[e.kat] || 'text-gray-500 bg-gray-100';
    const icon = KAT_SVG[e.kat]   || KAT_SVG['Inspektion'];
    return `
      <div class="grid grid-cols-3 items-center py-2 gap-1">
        <span class="text-xs text-gray-400">${dat(e.datum)}</span>
        <div class="flex items-center gap-1.5 min-w-0">
          <span class="w-5 h-5 rounded-lg flex items-center justify-center shrink-0 ${col}">${icon}</span>
          <span class="text-xs text-gray-700 truncate">${e.kat}</span>
        </div>
        <span class="text-xs font-semibold text-gray-600 text-right">${e.betrag > 0 ? eur(e.betrag) : '—'}</span>
      </div>`;
  }).join('')
    : '<p class="text-xs text-gray-400 text-center py-4">Noch keine Einträge</p>';
}

// ─── Untere KPI-Karten ────────────────────────────────────────────────────────

function renderBottomKpis(s) {
  const curYear = state.year || String(new Date().getFullYear());

  // Versicherung
  const versSum = state.kosten
    .filter(e => (e.kategorie || '').toLowerCase().includes('versicherung'))
    .reduce((acc, e) => acc + Number(e.betrag || 0), 0);
  const versFaell = state.kosten
    .filter(e => (e.kategorie || '').toLowerCase().includes('versicherung') && e.naechste_faelligkeit)
    .sort((a, b) => String(a.naechste_faelligkeit).localeCompare(String(b.naechste_faelligkeit)))[0];
  q('#kpi-versicherung').textContent     = versSum > 0 ? eur(versSum) : '—';
  q('#kpi-versicherung-sub').textContent = versFaell ? 'Nächste Zahlung ' + dat(versFaell.naechste_faelligkeit) : 'Keine Einträge';

  // Sonstige Ausgaben
  const hauptkat = new Set(['Tanken','Versicherung','Inspektion','Reparatur','KFZ-Steuer','Wartung']);
  const sonstigeSum = s.kosten.filter(e => !hauptkat.has(e.kategorie)).reduce((acc, e) => acc + Number(e.betrag || 0), 0);
  const sonstigePct = s.totalAll > 0 ? Math.round(sonstigeSum / s.totalAll * 100) : 0;
  q('#kpi-sonstige').textContent     = sonstigeSum > 0 ? eur(sonstigeSum) : '—';
  q('#kpi-sonstige-sub').textContent = s.totalAll > 0 ? `${sonstigePct}% der Gesamtausgaben` : 'Keine Ausgaben';

  // Wartungskosten YTD
  const wartKat = new Set(['Inspektion','Reparatur','Wartung','Service','HU/TÜV','Reifenwechsel']);
  const wartSum = state.kosten
    .filter(e => wartKat.has(e.kategorie) && String(e.datum || '').startsWith(curYear))
    .reduce((acc, e) => acc + Number(e.betrag || 0), 0);
  const wartPct = s.totalAll > 0 ? Math.round(wartSum / s.totalAll * 100) : 0;
  q('#kpi-wartung-ytd').textContent     = wartSum > 0 ? eur(wartSum) : '—';
  q('#kpi-wartung-ytd-sub').textContent = s.totalAll > 0 && wartSum > 0 ? `${wartPct}% der Gesamtausgaben` : 'Keine Wartungskosten';

  // Letzter Service
  const lastService = state.kosten
    .filter(e => wartKat.has(e.kategorie))
    .sort((a, b) => String(b.datum).localeCompare(String(a.datum)))[0];
  if (lastService) {
    const daysAgo = Math.floor((new Date() - new Date(lastService.datum + 'T12:00:00')) / 86400000);
    q('#kpi-letzter-service').textContent     = dat(lastService.datum);
    q('#kpi-letzter-service-sub').textContent = `vor ${daysAgo} Tagen`;
  } else {
    q('#kpi-letzter-service').textContent     = '—';
    q('#kpi-letzter-service-sub').textContent = 'Noch kein Service';
  }
}

// ─── Fahrzeug Sidebar ─────────────────────────────────────────────────────────

function renderVehicleSidebar() {
  const f = state.fahrzeug;

  // Fahrzeug-Sheet liest key-value Paare — wir suchen nach bekannten Keys
  // Die Sheet-Struktur ist: Spalte A = key, Spalte B = value
  // state.fahrzeug wird durch readSheet befüllt; Keys sind Spaltenheader, daher
  // suchen wir in den Row-Objekten nach den richtigen Feldern
  const rows = state.fahrzeug._raw || [];

  q('#v-modell').textContent      = f.modell      || f.Modell      || 'Seat Leon';
  q('#v-zulassung').textContent   = f.zulassung_datum || f['Zulassung Datum'] || '2018';
  q('#v-kraftstoff').textContent  = f.kraftstoff_standard || f.Kraftstoff || 'Benzin';
  q('#v-kennzeichen').textContent = f.kennzeichen || f.Kennzeichen || '—';

  const sub = q('#header-sub');
  const modell = f.modell || f.Modell;
  if (modell) sub.textContent = modell;
}

// ─── Liste: Tankvorgänge ──────────────────────────────────────────────────────

function renderTankList() {
  const tank = filtered(state.tank);
  q('#tank-count').textContent = `${tank.length} Einträge`;

  q('#list-tank').innerHTML = tank.length ? `
    <table class="w-full text-sm">
      <thead>
        <tr class="border-b border-slate-100">
          <th class="text-left px-5 py-3 text-xs font-semibold text-gray-400 uppercase tracking-wide">Datum</th>
          <th class="text-left px-5 py-3 text-xs font-semibold text-gray-400 uppercase tracking-wide hidden md:table-cell">Tankstelle</th>
          <th class="text-left px-5 py-3 text-xs font-semibold text-gray-400 uppercase tracking-wide hidden sm:table-cell">Kraftstoff</th>
          <th class="text-right px-5 py-3 text-xs font-semibold text-gray-400 uppercase tracking-wide">Liter</th>
          <th class="text-right px-5 py-3 text-xs font-semibold text-gray-400 uppercase tracking-wide">€/L</th>
          <th class="text-right px-5 py-3 text-xs font-semibold text-gray-400 uppercase tracking-wide">Kosten</th>
        </tr>
      </thead>
      <tbody>
        ${tank.map(e => {
          const preisL = Number(e.liter) > 0 ? (Number(e.kosten) / Number(e.liter)) : 0;
          return `<tr class="border-b border-slate-50 hover:bg-slate-50 transition-colors">
            <td class="px-5 py-3 text-gray-600 whitespace-nowrap">${dat(e.datum)}</td>
            <td class="px-5 py-3 text-gray-700 font-medium hidden md:table-cell max-w-xs truncate">${e.tankstelle || '—'}</td>
            <td class="px-5 py-3 hidden sm:table-cell">
              <span class="px-2 py-0.5 bg-blue-50 text-blue-600 rounded-full text-xs font-medium">${e.kraftstoff || '—'}</span>
            </td>
            <td class="px-5 py-3 text-right text-gray-600">${num(e.liter)}</td>
            <td class="px-5 py-3 text-right text-gray-500 text-xs">${preisL ? num(preisL, 3) : '—'}</td>
            <td class="px-5 py-3 text-right font-semibold text-gray-800">${eur(e.kosten)}</td>
          </tr>`;
        }).join('')}
      </tbody>
    </table>` : '<p class="text-sm text-gray-400 text-center py-12">Noch keine Tankeinträge — ersten Eintrag erstellen!</p>';
}

// ─── Liste: Kosten ────────────────────────────────────────────────────────────

function renderKostenList() {
  const kosten = filtered(state.kosten);
  q('#kosten-count').textContent = `${kosten.length} Einträge`;

  q('#list-kosten').innerHTML = kosten.length ? `
    <table class="w-full text-sm">
      <thead>
        <tr class="border-b border-slate-100">
          <th class="text-left px-5 py-3 text-xs font-semibold text-gray-400 uppercase tracking-wide">Datum</th>
          <th class="text-left px-5 py-3 text-xs font-semibold text-gray-400 uppercase tracking-wide">Beschreibung</th>
          <th class="text-left px-5 py-3 text-xs font-semibold text-gray-400 uppercase tracking-wide hidden sm:table-cell">Kategorie</th>
          <th class="text-left px-5 py-3 text-xs font-semibold text-gray-400 uppercase tracking-wide hidden md:table-cell">Fälligkeit</th>
          <th class="text-right px-5 py-3 text-xs font-semibold text-gray-400 uppercase tracking-wide">Betrag</th>
        </tr>
      </thead>
      <tbody>
        ${kosten.map(e => `
          <tr class="border-b border-slate-50 hover:bg-slate-50 transition-colors">
            <td class="px-5 py-3 text-gray-600 whitespace-nowrap">${dat(e.datum)}</td>
            <td class="px-5 py-3 text-gray-700 font-medium">${KAT_ICONS[e.kategorie] || '📌'} ${e.beschreibung || '—'}</td>
            <td class="px-5 py-3 hidden sm:table-cell">
              <span class="px-2 py-0.5 bg-slate-100 text-gray-600 rounded-full text-xs font-medium">${e.kategorie || '—'}</span>
            </td>
            <td class="px-5 py-3 text-gray-500 text-xs hidden md:table-cell">${e.naechste_faelligkeit ? dat(e.naechste_faelligkeit) : '—'}</td>
            <td class="px-5 py-3 text-right font-semibold text-gray-800">${eur(e.betrag)}</td>
          </tr>`).join('')}
      </tbody>
    </table>` : '<p class="text-sm text-gray-400 text-center py-12">Noch keine Ausgaben — ersten Eintrag erstellen!</p>';
}

// ─── Liste: Termine ───────────────────────────────────────────────────────────

function renderTermineList() {
  const container = q('#list-termine');
  const termine = state.kosten
    .filter(e => e.naechste_faelligkeit)
    .sort((a, b) => String(a.naechste_faelligkeit).localeCompare(String(b.naechste_faelligkeit)));

  if (!termine.length) {
    container.innerHTML = '<p class="text-sm text-gray-400 text-center py-12">Keine Fälligkeiten eingetragen — erst eine Ausgabe mit Fälligkeit erfassen.</p>';
    return;
  }

  // Event-Delegation: einmal registrieren, wird bei jedem render ersetzt
  container._termineData = termine;
  container.innerHTML = `
    <table class="w-full text-sm">
      <thead>
        <tr class="border-b border-slate-100">
          <th class="text-left px-5 py-3 text-xs font-semibold text-gray-400 uppercase tracking-wide">Beschreibung</th>
          <th class="text-left px-5 py-3 text-xs font-semibold text-gray-400 uppercase tracking-wide hidden sm:table-cell">Kategorie</th>
          <th class="text-left px-5 py-3 text-xs font-semibold text-gray-400 uppercase tracking-wide hidden sm:table-cell">Intervall</th>
          <th class="text-left px-5 py-3 text-xs font-semibold text-gray-400 uppercase tracking-wide">Fällig</th>
          <th class="text-right px-5 py-3 text-xs font-semibold text-gray-400 uppercase tracking-wide">Betrag</th>
          <th class="text-left px-5 py-3 text-xs font-semibold text-gray-400 uppercase tracking-wide">Status</th>
          <th class="px-5 py-3"></th>
        </tr>
      </thead>
      <tbody>
        ${termine.map(e => {
          const days  = Math.ceil((new Date(e.naechste_faelligkeit + 'T12:00:00') - new Date()) / 86400000);
          const badge = days < 0
            ? '<span class="text-xs px-2.5 py-1 rounded-full bg-red-100 text-red-600 font-semibold">Überfällig</span>'
            : days <= 30
              ? `<span class="text-xs px-2.5 py-1 rounded-full bg-orange-100 text-orange-600 font-semibold">in ${days}d</span>`
              : days <= 90
                ? `<span class="text-xs px-2.5 py-1 rounded-full bg-yellow-100 text-yellow-700 font-semibold">in ${days}d</span>`
                : `<span class="text-xs px-2.5 py-1 rounded-full bg-green-50 text-green-600 font-semibold">in ${days}d</span>`;
          const col  = KAT_COLOR[e.kategorie] || 'text-gray-500 bg-gray-100';
          const icon = KAT_SVG[e.kategorie]   || KAT_SVG['Inspektion'];
          const canAdvance = e.intervall && e.intervall !== 'einmalig' && INTERVALL_MONTHS[e.intervall];
          const btnLabel   = canAdvance
            ? `Erledigt → ${dat(nextDueDate(e.naechste_faelligkeit, e.intervall))}`
            : 'Erledigt';
          return `
            <tr class="border-b border-slate-50 hover:bg-slate-50 transition-colors">
              <td class="px-5 py-3">
                <div class="flex items-center gap-2.5">
                  <span class="w-7 h-7 rounded-lg flex items-center justify-center shrink-0 ${col}">${icon}</span>
                  <span class="font-medium text-gray-700">${e.beschreibung || e.kategorie}</span>
                </div>
              </td>
              <td class="px-5 py-3 hidden sm:table-cell">
                <span class="px-2 py-0.5 bg-slate-100 text-gray-600 rounded-full text-xs">${e.kategorie}</span>
              </td>
              <td class="px-5 py-3 hidden sm:table-cell text-xs text-gray-500">${e.intervall || '—'}</td>
              <td class="px-5 py-3 text-gray-600 whitespace-nowrap">${dat(e.naechste_faelligkeit)}</td>
              <td class="px-5 py-3 text-right font-semibold text-gray-800">${eur(e.betrag)}</td>
              <td class="px-5 py-3">${badge}</td>
              <td class="px-5 py-3">
                <button data-erledigt="${e._row}"
                  class="text-xs px-3 py-1.5 rounded-xl bg-green-50 text-green-700 hover:bg-green-100 font-medium transition-colors whitespace-nowrap">
                  ${btnLabel}
                </button>
              </td>
            </tr>`;
        }).join('')}
      </tbody>
    </table>`;

  // Event-Delegation auf Container
  container.onclick = ev => {
    const btn = ev.target.closest('[data-erledigt]');
    if (!btn) return;
    const row  = Number(btn.dataset.erledigt);
    const entry = container._termineData.find(e => e._row === row);
    if (entry) markTerminErledigt(entry);
  };
}

// ─── Statistiken ──────────────────────────────────────────────────────────────

function renderStatistiken() {
  const s = calcStats();

  // Jahres-Chart
  const years = Object.keys(s.yearMap).sort();
  if (state.charts.year) state.charts.year.destroy();
  state.charts.year = new Chart(q('#yearChart'), {
    type: 'bar',
    data: {
      labels: years,
      datasets: [{ data: years.map(y => s.yearMap[y]),
        backgroundColor: '#3b82f6', borderRadius: 8, borderSkipped: false }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false },
        tooltip: { callbacks: { label: ctx => ' ' + eur(ctx.parsed.y) } } },
      scales: {
        x: { grid: { display: false }, border: { display: false }, ticks: { color: '#94a3b8', font: { size: 11 } } },
        y: { grid: { color: '#f1f5f9' }, border: { display: false }, ticks: { color: '#94a3b8', font: { size: 11 }, callback: v => eur(v) } }
      }
    }
  });

  // Kraftstoff Donut
  const kEntries = Object.entries(s.kraftstoffMap).filter(([, v]) => v > 0);
  const kTotal   = kEntries.reduce((s, [, v]) => s + v, 0);
  if (state.charts.kraftstoff) state.charts.kraftstoff.destroy();
  state.charts.kraftstoff = new Chart(q('#kraftstoffChart'), {
    type: 'doughnut',
    data: {
      labels: kEntries.map(([k]) => k),
      datasets: [{ data: kEntries.map(([, v]) => v), backgroundColor: COLORS, borderWidth: 0 }]
    },
    options: {
      responsive: true, maintainAspectRatio: false, cutout: '65%',
      plugins: { legend: { display: false },
        tooltip: { callbacks: { label: ctx => ` ${ctx.label}: ${num(ctx.parsed)} L` } } }
    }
  });
  q('#kraftstoff-legend').innerHTML = kEntries.map(([k, v], i) => `
    <div class="flex items-center justify-between">
      <div class="flex items-center gap-2">
        <span class="w-2.5 h-2.5 rounded-full" style="background:${COLORS[i]}"></span>
        <span class="text-xs text-gray-600">${k}</span>
      </div>
      <span class="text-xs font-medium text-gray-700">${num(v)} L · ${kTotal > 0 ? Math.round(v/kTotal*100) : 0}%</span>
    </div>`).join('');
}

// ─── Render All ───────────────────────────────────────────────────────────────

function renderAll() {
  renderDashboard();
  renderTankList();
  renderKostenList();
  renderTermineList();
  renderStatistiken();
}

// ─── Navigation ───────────────────────────────────────────────────────────────

function switchView(view) {
  state.view = view;
  qAll('.view').forEach(el => el.classList.remove('active'));
  const el = q(`#view-${view}`);
  if (el) el.classList.add('active');

  // Sidebar
  qAll('.nav-link').forEach(btn => {
    const active = btn.dataset.view === view;
    btn.classList.toggle('active', active);
  });

  // Mobile nav
  qAll('.mobile-nav').forEach(btn => {
    const active = btn.dataset.view === view;
    btn.classList.toggle('text-blue-600', active);
    btn.classList.toggle('text-gray-400', !active);
  });
}

// ─── Modals ───────────────────────────────────────────────────────────────────

const openModal  = id => q(`#${id}`).classList.remove('hidden');
const closeAll   = ()  => qAll('[id^="modal-"]').forEach(m => m.classList.add('hidden'));

// Form-Listener werden in DOMContentLoaded registriert (siehe unten)

async function submitForm(formId, modalId, sheet, data, successMsg) {
  const btn = q(`#${formId} button[type="submit"]`);
  btn.disabled = true; btn.textContent = 'Speichern…';
  try {
    await api.add(sheet, data);
    closeAll();
    q(`#${formId}`).reset();
    setDefaultDates(formId);
    toast(successMsg);
    await loadData();
  } catch (err) {
    toast('❌ Fehler: ' + err.message, 4000);
  } finally {
    btn.disabled = false; btn.textContent = 'Speichern';
  }
}

// ─── Toast ────────────────────────────────────────────────────────────────────

function toast(msg, ms = 2500) {
  const el = q('#toast');
  el.textContent = msg;
  el.classList.remove('hidden');
  clearTimeout(el._t);
  el._t = setTimeout(() => el.classList.add('hidden'), ms);
}

function setLoading(show) {
  q('#loading').classList.toggle('hidden', !show);
  q('#main').classList.toggle('hidden', show);
}
function showError(msg) { const el = q('#error'); el.textContent = msg; el.classList.remove('hidden'); }
function clearError()   { q('#error').classList.add('hidden'); }
function showToast(msg, ms = 3000) {
  const el = q('#toast');
  el.textContent = msg;
  el.classList.remove('hidden');
  clearTimeout(el._t);
  el._t = setTimeout(() => el.classList.add('hidden'), ms);
}

function setDefaultDates(formId) {
  qAll(`#${formId} input[type="date"]`).forEach(el => {
    if (!el.value && el.name !== 'naechste_faelligkeit') el.value = today();
  });
}

// ─── Init ─────────────────────────────────────────────────────────────────────

document.addEventListener('DOMContentLoaded', () => {
  setDefaultDates('form-tank');
  setDefaultDates('form-kosten');
  setDefaultDates('form-km');

  // Sidebar + mobile nav
  qAll('.nav-link[data-view], .mobile-nav[data-view]').forEach(btn => {
    btn.addEventListener('click', () => switchView(btn.dataset.view));
  });

  // Refresh (Desktop + Mobile)
  q('#btn-refresh').addEventListener('click', loadData);
  const btnRefreshMobile = q('#btn-refresh-mobile');
  if (btnRefreshMobile) btnRefreshMobile.addEventListener('click', loadData);

  // Jahr-Filter (Desktop + Mobile)
  const syncYear = val => {
    state.year = val;
    const f  = q('#year-filter');
    const fm = q('#year-filter-mobile');
    if (f  && f.value  !== val) f.value  = val;
    if (fm && fm.value !== val) fm.value = val;
    renderAll();
  };
  q('#year-filter').addEventListener('change', e => syncYear(e.target.value));
  const yfm = q('#year-filter-mobile');
  if (yfm) yfm.addEventListener('change', e => syncYear(e.target.value));

  // Fälligkeit auto-berechnen: Datum oder Intervall ändern → Fälligkeitsfeld aktualisieren
  const autoFaelligkeit = () => {
    const d = q('#kosten-datum').value;
    const i = q('#kosten-intervall').value;
    const f = q('#kosten-faelligkeit');
    if (!f.value) f.value = nextDueDate(d, i) || '';
  };
  q('#kosten-datum').addEventListener('change', autoFaelligkeit);
  q('#kosten-intervall').addEventListener('change', () => {
    const d = q('#kosten-datum').value;
    const i = q('#kosten-intervall').value;
    q('#kosten-faelligkeit').value = nextDueDate(d, i) || '';
  });
  q('#btn-calc-faelligkeit').addEventListener('click', () => {
    const d = q('#kosten-datum').value;
    const i = q('#kosten-intervall').value;
    q('#kosten-faelligkeit').value = nextDueDate(d, i) || '';
  });

  // Export
  q('#btn-export-tank').addEventListener('click',   exportToExcel);
  q('#btn-export-kosten').addEventListener('click', exportToExcel);

  // Löschen (Bestätigungs-Modal öffnen)
  q('#btn-delete-tank').addEventListener('click',   () => confirmDeleteAll('Tanken',  'Tankvorgänge'));
  q('#btn-delete-kosten').addEventListener('click', () => confirmDeleteAll('Kosten',  'Ausgaben'));

  // Bestätigungs-Modal Buttons
  q('#confirm-cancel').addEventListener('click', () => q('#modal-confirm').classList.add('hidden'));
  q('#modal-confirm').addEventListener('click', e => { if (e.target === q('#modal-confirm')) q('#modal-confirm').classList.add('hidden'); });
  q('#confirm-ok').addEventListener('click', async () => {
    const sheet = q('#confirm-sheet').value;
    const label = q('#confirm-label').textContent;
    const btn   = q('#confirm-ok');
    btn.disabled    = true;
    btn.textContent = 'Löschen…';
    try {
      await api.deleteAll(sheet);
      q('#modal-confirm').classList.add('hidden');
      showToast(`${label} gelöscht`);
      await loadData();
    } catch (err) {
      showError('Fehler beim Löschen: ' + err.message);
    } finally {
      btn.disabled    = false;
      btn.textContent = 'Löschen';
    }
  });

  // Modals öffnen
  q('#btn-add-km').addEventListener('click', () => { setDefaultDates('form-km'); openModal('modal-km'); });
  q('#btn-add-tank').addEventListener('click',   () => { setDefaultDates('form-tank');   openModal('modal-tank'); });
  const openKostenModal = () => {
    setDefaultDates('form-kosten');
    // Fälligkeit sofort aus Datum + Intervall berechnen (nicht einfach heute)
    q('#kosten-faelligkeit').value = nextDueDate(q('#kosten-datum').value, q('#kosten-intervall').value) || '';
    openModal('modal-kosten');
  };
  q('#btn-add-kosten').addEventListener('click', openKostenModal);
  q('#btn-add-termin').addEventListener('click', openKostenModal);

  // Modals schließen
  qAll('.modal-close').forEach(btn => btn.addEventListener('click', closeAll));
  qAll('[id^="modal-"]').forEach(m => m.addEventListener('click', e => { if (e.target === m) closeAll(); }));

  // Form: Tankvorgang
  q('#form-tank').addEventListener('submit', async e => {
    e.preventDefault();
    const fd = new FormData(e.target);
    const l  = parseFloat(fd.get('liter'))  || 0;
    const k  = parseFloat(fd.get('kosten')) || 0;
    await submitForm('form-tank', 'modal-tank', 'Tanken', {
      datum:           fd.get('datum'),
      tankstelle:      fd.get('tankstelle'),
      kraftstoff:      fd.get('kraftstoff'),
      liter:           l,
      kosten:          k,
      preis_pro_liter: l > 0 ? parseFloat((k / l).toFixed(4)) : '',
      km_stand:        fd.get('km_stand') || '',
      konto:           fd.get('konto'),
      karte:           fd.get('karte') || '',
      beleg:           fd.get('beleg') ? 'ja' : 'nein',
      hinweis:         fd.get('hinweis') || ''
    }, '⛽ Tankvorgang gespeichert');
  });

  // Form: Kosten
  q('#form-kosten').addEventListener('submit', async e => {
    e.preventDefault();
    const fd = new FormData(e.target);
    await submitForm('form-kosten', 'modal-kosten', 'Kosten', {
      datum:                fd.get('datum'),
      kategorie:            fd.get('kategorie'),
      beschreibung:         fd.get('beschreibung'),
      betrag:               parseFloat(fd.get('betrag')) || 0,
      intervall:            fd.get('intervall'),
      naechste_faelligkeit: fd.get('naechste_faelligkeit') || '',
      konto:                fd.get('konto'),
      beleg:                fd.get('beleg') ? 'ja' : 'nein',
      hinweis:              fd.get('hinweis') || ''
    }, '✅ Ausgabe gespeichert');
  });

  // Form: KM-Stand
  q('#form-km').addEventListener('submit', async e => {
    e.preventDefault();
    const fd = new FormData(e.target);
    await submitForm('form-km', 'modal-km', 'KMStand', {
      datum:    fd.get('datum'),
      km_stand: parseInt(fd.get('km_stand')) || 0,
      notiz:    fd.get('notiz') || ''
    }, '📍 KM-Stand gespeichert');
  });

  // ── Excel Import ──────────────────────────────────────────────────────────
  let importEntries = [];

  function excelSerialToISO(serial) {
    const d = XLSX.SSF.parse_date_code(serial);
    if (!d) return '';
    return `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`;
  }

  function parseExcel(workbook) {
    const ws    = workbook.Sheets[workbook.SheetNames[0]];
    const rows  = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    const results = [];

    for (const row of rows) {
      // Datenzeilen erkennen: Spalte C (index 2) enthält laufende Nummer
      const no = row[2];
      if (!no || isNaN(Number(no))) continue;

      const datumSerial = row[3];
      if (!datumSerial || isNaN(Number(datumSerial))) continue;
      const datum = excelSerialToISO(Number(datumSerial));
      if (!datum) continue;

      const liter  = parseFloat(row[6]) || 0;
      const kosten = parseFloat(row[7]) || 0;

      results.push({
        datum,
        tankstelle:      String(row[4] || '').trim(),
        kraftstoff:      String(row[5] || '').trim(),
        liter:           Math.round(liter  * 1000) / 1000,
        kosten:          Math.round(kosten * 100)  / 100,
        preis_pro_liter: liter > 0 ? Math.round(kosten / liter * 10000) / 10000 : '',
        km_stand:        '',
        konto:           String(row[9] || '').trim(),
        karte:           String(row[10] || '').trim(),
        beleg:           String(row[1] || '') === 'o' ? 'ja' : 'nein',
        hinweis:         String(row[11] || '').trim(),
      });
    }
    return results.sort((a, b) => a.datum.localeCompare(b.datum));
  }

  function filterNew(entries) {
    const existing = new Set(
      state.tank.map(e => `${e.datum}|${String(e.tankstelle).trim().toLowerCase()}`)
    );
    return entries.filter(e =>
      !existing.has(`${e.datum}|${e.tankstelle.toLowerCase()}`)
    );
  }

  async function runImport(entries) {
    const CHUNK = 15;
    const chunks = [];
    for (let i = 0; i < entries.length; i += CHUNK) chunks.push(entries.slice(i, i + CHUNK));

    q('#import-actions').classList.add('hidden');
    q('#import-progress').classList.remove('hidden');

    let done = 0;
    for (let i = 0; i < chunks.length; i++) {
      q('#import-progress-text').textContent = `Batch ${i+1}/${chunks.length} — ${done}/${entries.length} importiert…`;
      q('#import-bar').style.width = `${Math.round(done / entries.length * 100)}%`;
      try {
        await apiCall({ action: 'batchAdd', sheet: 'Tanken', data: JSON.stringify(chunks[i]) });
        done += chunks[i].length;
      } catch (e) {
        showError('Import-Fehler: ' + e.message);
        return;
      }
    }

    q('#import-bar').style.width = '100%';
    q('#import-progress-text').textContent = `${done} Einträge importiert.`;
    showToast(`✅ ${done} Einträge importiert`);
    setTimeout(() => {
      q('#import-card').classList.add('hidden');
      q('#import-file').value = '';
      importEntries = [];
      loadData();
    }, 1500);
  }

  function showImportPreview(entries, total) {
    const newCount  = entries.length;
    const skipCount = total - newCount;
    q('#import-title').textContent = `${newCount} neue Einträge gefunden`;
    q('#import-info').textContent  = skipCount > 0
      ? `${skipCount} bereits vorhanden (werden übersprungen)`
      : 'Keine Duplikate erkannt';

    q('#import-sample').innerHTML = entries.slice(0, 8).map(e =>
      `<div class="flex gap-2"><span class="text-gray-400 w-24 flex-shrink-0">${e.datum}</span><span class="truncate">${e.tankstelle}</span><span class="ml-auto text-gray-500 flex-shrink-0">${e.kosten > 0 ? e.kosten.toFixed(2)+' €' : '—'}</span></div>`
    ).join('') + (entries.length > 8 ? `<div class="text-gray-400 pt-1">… und ${entries.length - 8} weitere</div>` : '');

    if (newCount > 0) {
      q('#import-actions').classList.remove('hidden');
      q('#import-actions').style.display = 'flex';
    } else {
      q('#import-info').textContent = 'Alle Einträge bereits vorhanden — nichts zu importieren.';
    }
    q('#import-progress').classList.add('hidden');
    q('#import-card').classList.remove('hidden');
  }

  q('#import-file').addEventListener('change', async e => {
    const file = e.target.files[0];
    if (!file) return;
    q('#import-card').classList.remove('hidden');
    q('#import-title').textContent = 'Datei wird gelesen…';
    q('#import-info').textContent  = file.name;
    q('#import-sample').innerHTML  = '';
    q('#import-actions').classList.add('hidden');
    q('#import-progress').classList.add('hidden');

    try {
      const buffer   = await file.arrayBuffer();
      const wb       = XLSX.read(buffer, { type: 'array' });
      const all      = parseExcel(wb);
      importEntries  = filterNew(all);
      showImportPreview(importEntries, all.length);
    } catch (err) {
      q('#import-title').textContent = 'Fehler beim Lesen';
      q('#import-info').textContent  = err.message;
    }
  });

  q('#import-confirm').addEventListener('click', () => runImport(importEntries));

  const cancelImport = () => {
    q('#import-card').classList.add('hidden');
    q('#import-file').value = '';
    importEntries = [];
  };
  q('#import-cancel').addEventListener('click',  cancelImport);
  q('#import-cancel2').addEventListener('click', cancelImport);

  loadData();
});
