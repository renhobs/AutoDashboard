// AutoDashboard — App-Logik v0.2

// ─── State ────────────────────────────────────────────────────────────────────

const state = {
  tank:    [],
  kosten:  [],
  fahrzeug: {},
  view:    'dashboard',
  year:    String(new Date().getFullYear()),
  charts:  {}
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
  read: sheet         => apiCall({ action: 'read', sheet }),
  add:  (sheet, data) => apiCall({ action: 'add',  sheet, data }),
};

// ─── Daten laden ──────────────────────────────────────────────────────────────

async function loadData() {
  setLoading(true);
  clearError();
  try {
    const [t, k, f] = await Promise.all([
      api.read('Tanken'),
      api.read('Kosten'),
      api.read('Fahrzeug')
    ]);

    state.tank   = (t.data || []).sort((a, b) => String(b.datum).localeCompare(String(a.datum)));
    state.kosten = (k.data || []).sort((a, b) => String(b.datum).localeCompare(String(a.datum)));

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

  const sel = q('#year-filter');
  const current = sel.value;
  sel.innerHTML = '<option value="">Alle Jahre</option>';
  [...years].sort().reverse().forEach(y => {
    const opt = document.createElement('option');
    opt.value = y;
    opt.textContent = y;
    if (y === current || (!current && y === String(new Date().getFullYear()))) opt.selected = true;
    sel.appendChild(opt);
  });
  state.year = sel.value;
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

  // Letzter KM-Stand
  const mitKm = [...state.tank].filter(e => e.km_stand && Number(e.km_stand) > 0)
    .sort((a, b) => String(b.datum).localeCompare(String(a.datum)));
  const lastKm = mitKm.length ? Number(mitKm[0].km_stand) : 0;

  // KM diesen Monat
  const monatStr = today().slice(0, 7);
  const kmMonat = (() => {
    const thisMonth = state.tank.filter(e => String(e.datum).startsWith(monatStr) && e.km_stand);
    const lastMonth = state.tank.filter(e => !String(e.datum).startsWith(monatStr) && e.km_stand)
      .sort((a, b) => String(b.datum).localeCompare(String(a.datum)));
    if (thisMonth.length < 2 && !lastMonth.length) return 0;
    const maxKm = Math.max(...thisMonth.map(e => Number(e.km_stand || 0)));
    const minKm = lastMonth.length
      ? Number(lastMonth[0].km_stand || 0)
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
           avgPreis, lastKm, kmMonat, monthly, last6, katMap, kraftstoffMap, yearMap };
}

// ─── Render Übersicht ─────────────────────────────────────────────────────────

function renderDashboard() {
  const s = calcStats();

  // KPI
  q('#kpi-gesamt').textContent      = eur(s.totalAll);
  q('#kpi-gesamt-sub').textContent  = `${s.tank.length} Tankungen + ${s.kosten.length} Ausgaben`;
  q('#kpi-tank-gesamt').textContent = eur(s.totalTank);
  q('#kpi-tank-sub').textContent    = `${s.tank.length} Tankungen · ${num(s.totalLiter)} L`;
  q('#kpi-preis').textContent       = s.avgPreis ? num(s.avgPreis, 3) + ' €/L' : '—';
  q('#kpi-preis-sub').textContent   = `${s.tank.length > 0 ? num(s.totalLiter) + ' L gesamt' : 'Keine Daten'}`;
  q('#kpi-km').textContent          = s.lastKm ? s.lastKm.toLocaleString('de-DE') + ' km' : '—';
  q('#kpi-km-sub').textContent      = s.lastKm ? 'Zuletzt erfasst' : 'KM-Stand noch nicht erfasst';

  // Letztes Tanken
  if (s.tank.length > 0) {
    const lt = s.tank[0];
    const preisL = Number(lt.liter) > 0 ? (Number(lt.kosten) / Number(lt.liter)).toFixed(3) : '—';
    q('#last-tank-info').classList.remove('hidden');
    q('#last-tank-date').textContent    = dat(lt.datum);
    q('#last-tank-details').textContent = `${num(lt.liter)} L · ${preisL} €/L · ${eur(lt.kosten)}`;
  }

  // KM diesen Monat
  if (s.kmMonat > 0) {
    q('#km-monat-bar').classList.remove('hidden');
    q('#km-monat-val').textContent = s.kmMonat.toLocaleString('de-DE');
  }

  // Bar Chart
  renderBarChart(s);

  // Donut Chart
  renderDonutChart(s);

  // Nächste Termine
  renderUpcomingList();

  // Letzte Ausgaben
  renderRecentAusgaben(s);

  // Fahrzeug Sidebar
  renderVehicleSidebar();
}

// ─── Bar Chart ────────────────────────────────────────────────────────────────

function renderBarChart(s) {
  const labels = s.last6.map(m => {
    const d = new Date(m + '-01T12:00:00');
    return d.toLocaleDateString('de-DE', { month: 'short' });
  });
  const data = s.last6.map(m => Math.round((s.monthly[m] || 0) * 100) / 100);

  if (state.charts.bar) state.charts.bar.destroy();
  state.charts.bar = new Chart(q('#barChart'), {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        data,
        backgroundColor: '#3b82f6',
        borderRadius: 6,
        borderSkipped: false,
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false }, tooltip: {
        callbacks: { label: ctx => ' ' + eur(ctx.parsed.y) }
      }},
      scales: {
        x: { grid: { display: false }, border: { display: false },
             ticks: { color: '#94a3b8', font: { size: 11 } } },
        y: { grid: { color: '#f1f5f9' }, border: { display: false },
             ticks: { color: '#94a3b8', font: { size: 11 },
               callback: v => v === 0 ? '' : eur(v) } }
      }
    }
  });
}

// ─── Donut Chart ──────────────────────────────────────────────────────────────

function renderDonutChart(s) {
  const entries = Object.entries(s.katMap).filter(([, v]) => v > 0)
    .sort((a, b) => b[1] - a[1]);
  const total = entries.reduce((sum, [, v]) => sum + v, 0);

  if (state.charts.donut) state.charts.donut.destroy();
  state.charts.donut = new Chart(q('#donutChart'), {
    type: 'doughnut',
    data: {
      labels: entries.map(([k]) => k),
      datasets: [{ data: entries.map(([, v]) => v),
        backgroundColor: COLORS, borderWidth: 0, hoverOffset: 4 }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: '72%',
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: ctx => ` ${ctx.label}: ${eur(ctx.parsed)}` } }
      }
    },
    plugins: [{
      id: 'centerText',
      afterDraw(chart) {
        const { ctx, chartArea: { top, left, width, height } } = chart;
        ctx.save();
        ctx.font = 'bold 15px system-ui';
        ctx.fillStyle = '#1e293b';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText(eur(total), left + width / 2, top + height / 2 - 8);
        ctx.font = '11px system-ui';
        ctx.fillStyle = '#94a3b8';
        ctx.fillText('Gesamt', left + width / 2, top + height / 2 + 10);
        ctx.restore();
      }
    }]
  });

  // Legende
  q('#donut-legend').innerHTML = entries.map(([k, v], i) => `
    <div class="flex items-center justify-between">
      <div class="flex items-center gap-2">
        <span class="w-2.5 h-2.5 rounded-full shrink-0" style="background:${COLORS[i]}"></span>
        <span class="text-xs text-gray-600">${k}</span>
      </div>
      <div class="text-right">
        <span class="text-xs font-medium text-gray-700">${eur(v)}</span>
        <span class="text-xs text-gray-400 ml-1">${total > 0 ? Math.round(v/total*100) : 0}%</span>
      </div>
    </div>`).join('');
}

// ─── Upcoming Termine ─────────────────────────────────────────────────────────

function renderUpcomingList() {
  const now = today();
  const upcoming = state.kosten
    .filter(e => e.naechste_faelligkeit)
    .sort((a, b) => String(a.naechste_faelligkeit).localeCompare(String(b.naechste_faelligkeit)))
    .slice(0, 6);

  const el = q('#upcoming-list');
  if (!upcoming.length) {
    el.innerHTML = '<p class="text-sm text-gray-400 text-center py-4">Keine Fälligkeiten eingetragen</p>';
    return;
  }
  el.innerHTML = upcoming.map(e => {
    const days = Math.ceil((new Date(e.naechste_faelligkeit) - new Date()) / 86400000);
    const overdue = days < 0;
    const soon    = days >= 0 && days <= 60;
    const badge   = overdue
      ? '<span class="text-xs px-2 py-0.5 rounded-full bg-red-100 text-red-600 font-medium">Überfällig</span>'
      : soon
        ? `<span class="text-xs px-2 py-0.5 rounded-full bg-orange-100 text-orange-600 font-medium">in ${days} Tagen</span>`
        : `<span class="text-xs px-2 py-0.5 rounded-full bg-green-50 text-green-600 font-medium">in ${days} Tagen</span>`;
    return `
      <div class="flex items-center justify-between py-2 border-b border-slate-50 last:border-0">
        <div class="flex items-center gap-3">
          <span class="text-base">${KAT_ICONS[e.kategorie] || '📌'}</span>
          <div>
            <p class="text-sm font-medium text-gray-700">${e.beschreibung || e.kategorie}</p>
            <p class="text-xs text-gray-400">${dat(e.naechste_faelligkeit)} · ${e.kategorie}</p>
          </div>
        </div>
        <div class="flex items-center gap-3 shrink-0">
          <span class="text-sm font-semibold text-gray-700">${eur(e.betrag)}</span>
          ${badge}
        </div>
      </div>`;
  }).join('');
}

// ─── Letzte Ausgaben ──────────────────────────────────────────────────────────

function renderRecentAusgaben(s) {
  const all = [
    ...s.tank.slice(0, 8).map(e => ({ datum: e.datum, label: e.tankstelle || '—', sub: `⛽ ${num(e.liter)} L`, betrag: Number(e.kosten || 0), typ: 'tank' })),
    ...s.kosten.slice(0, 8).map(e => ({ datum: e.datum, label: e.beschreibung || e.kategorie, sub: e.kategorie, betrag: Number(e.betrag || 0), typ: 'kosten' }))
  ].sort((a, b) => String(b.datum).localeCompare(String(a.datum))).slice(0, 7);

  q('#recent-ausgaben').innerHTML = all.length ? all.map(e => `
    <div class="flex items-center justify-between py-1.5">
      <div class="flex-1 min-w-0 pr-2">
        <p class="text-sm font-medium text-gray-700 truncate">${e.label}</p>
        <p class="text-xs text-gray-400">${dat(e.datum)} · ${e.sub}</p>
      </div>
      <span class="text-sm font-semibold text-gray-700 shrink-0">${eur(e.betrag)}</span>
    </div>`).join('')
    : '<p class="text-sm text-gray-400 text-center py-4">Noch keine Einträge</p>';
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
  const now = today();
  const termine = state.kosten
    .filter(e => e.naechste_faelligkeit)
    .sort((a, b) => String(a.naechste_faelligkeit).localeCompare(String(b.naechste_faelligkeit)));

  q('#list-termine').innerHTML = termine.length ? `
    <table class="w-full text-sm">
      <thead>
        <tr class="border-b border-slate-100">
          <th class="text-left px-5 py-3 text-xs font-semibold text-gray-400 uppercase tracking-wide">Beschreibung</th>
          <th class="text-left px-5 py-3 text-xs font-semibold text-gray-400 uppercase tracking-wide hidden sm:table-cell">Kategorie</th>
          <th class="text-left px-5 py-3 text-xs font-semibold text-gray-400 uppercase tracking-wide">Fällig</th>
          <th class="text-right px-5 py-3 text-xs font-semibold text-gray-400 uppercase tracking-wide">Betrag</th>
          <th class="text-left px-5 py-3 text-xs font-semibold text-gray-400 uppercase tracking-wide">Status</th>
        </tr>
      </thead>
      <tbody>
        ${termine.map(e => {
          const days = Math.ceil((new Date(e.naechste_faelligkeit) - new Date()) / 86400000);
          const badge = days < 0
            ? '<span class="text-xs px-2 py-0.5 rounded-full bg-red-100 text-red-600 font-medium">Überfällig</span>'
            : days <= 30
              ? `<span class="text-xs px-2 py-0.5 rounded-full bg-orange-100 text-orange-600 font-medium">in ${days}d</span>`
              : days <= 90
                ? `<span class="text-xs px-2 py-0.5 rounded-full bg-yellow-100 text-yellow-700 font-medium">in ${days}d</span>`
                : `<span class="text-xs px-2 py-0.5 rounded-full bg-green-50 text-green-600 font-medium">in ${days}d</span>`;
          return `<tr class="border-b border-slate-50 hover:bg-slate-50 transition-colors">
            <td class="px-5 py-3 text-gray-700 font-medium">${KAT_ICONS[e.kategorie] || '📌'} ${e.beschreibung || '—'}</td>
            <td class="px-5 py-3 hidden sm:table-cell"><span class="px-2 py-0.5 bg-slate-100 text-gray-600 rounded-full text-xs">${e.kategorie}</span></td>
            <td class="px-5 py-3 text-gray-600 whitespace-nowrap">${dat(e.naechste_faelligkeit)}</td>
            <td class="px-5 py-3 text-right font-semibold text-gray-800">${eur(e.betrag)}</td>
            <td class="px-5 py-3">${badge}</td>
          </tr>`;
        }).join('')}
      </tbody>
    </table>` : '<p class="text-sm text-gray-400 text-center py-12">Keine Fälligkeiten eingetragen</p>';
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

// ─── Helpers ──────────────────────────────────────────────────────────────────

const q    = s => document.querySelector(s);
const qAll = s => document.querySelectorAll(s);

function setLoading(show) {
  q('#loading').classList.toggle('hidden', !show);
  q('#main').classList.toggle('hidden', show);
}
function showError(msg) { const el = q('#error'); el.textContent = msg; el.classList.remove('hidden'); }
function clearError()   { q('#error').classList.add('hidden'); }
function setDefaultDates(formId) {
  qAll(`#${formId} input[type="date"]`).forEach(el => { if (!el.value) el.value = today(); });
}

// ─── Init ─────────────────────────────────────────────────────────────────────

document.addEventListener('DOMContentLoaded', () => {
  setDefaultDates('form-tank');
  setDefaultDates('form-kosten');

  // Sidebar + mobile nav
  qAll('.nav-link[data-view], .mobile-nav[data-view]').forEach(btn => {
    btn.addEventListener('click', () => switchView(btn.dataset.view));
  });

  // Refresh
  q('#btn-refresh').addEventListener('click', loadData);

  // Jahr-Filter
  q('#year-filter').addEventListener('change', e => {
    state.year = e.target.value;
    renderAll();
  });

  // Modals öffnen
  q('#btn-add-tank').addEventListener('click',   () => { setDefaultDates('form-tank');   openModal('modal-tank'); });
  q('#btn-add-kosten').addEventListener('click', () => { setDefaultDates('form-kosten'); openModal('modal-kosten'); });
  q('#btn-add-termin').addEventListener('click', () => { setDefaultDates('form-kosten'); openModal('modal-kosten'); });

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

  loadData();
});
