// AutoDashboard — App-Logik v0.2

// ─── Helpers ──────────────────────────────────────────────────────────────────

const q    = s => document.querySelector(s);
const qAll = s => document.querySelectorAll(s);

// ─── State ────────────────────────────────────────────────────────────────────

const state = {
  tank:              [],
  kosten:            [],
  kmstand:           [],
  fahrzeug:          {},
  zahlungsmethoden:  [],
  view:              'dashboard',
  year:              String(new Date().getFullYear()),
  charts:            {}
};

let _editKostenRow = null;

// ─── Formatierung ─────────────────────────────────────────────────────────────

const eur   = v  => Number(v || 0).toLocaleString('de-DE', { style: 'currency', currency: 'EUR' });
const num   = (v, d=2) => Number(v || 0).toLocaleString('de-DE', { minimumFractionDigits: d, maximumFractionDigits: d });
const dat   = v  => v ? new Date(v + 'T12:00:00').toLocaleDateString('de-DE') : '—';
const today = () => new Date().toISOString().split('T')[0];
const COLORS = ['#3b82f6','#22c55e','#a855f7','#f59e0b','#94a3b8','#ef4444','#06b6d4'];
const KAT_ICONS = { 'KFZ-Steuer':'📋', Versicherung:'🛡', Inspektion:'🔧', Reparatur:'⚙️', Sonstiges:'📌' };

// ─── API ──────────────────────────────────────────────────────────────────────

async function apiCall(params, timeoutMs = 15000) {
  const url = new URL(API_URL);
  for (const [k, v] of Object.entries(params))
    url.searchParams.set(k, typeof v === 'object' ? JSON.stringify(v) : String(v));
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const res = await fetch(url.toString(), { signal: controller.signal });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const json = await res.json();
    if (json.error) throw new Error(json.error);
    return json;
  } catch (err) {
    if (err.name === 'AbortError') throw new Error('Zeitüberschreitung — Server antwortet nicht (>15s)');
    throw err;
  } finally {
    clearTimeout(timer);
  }
}

const api = {
  read:      sheet              => apiCall({ action: 'read',      sheet }),
  add:       (sheet, data)      => apiCall({ action: 'add',       sheet, data }),
  update:    (sheet, row, data) => apiCall({ action: 'update',    sheet, row, data }),
  delete:    (sheet, row)       => apiCall({ action: 'delete',    sheet, row }),
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

function downloadImportTemplate() {
  const wb = XLSX.utils.book_new();
  const header  = ['No.', 'Datum', 'Tankstelle', 'PLZ', 'Ort', 'Land', 'Kraftstoff', 'Liter', 'Kosten (€)', 'KM-Stand', 'Konto', 'Karten-ID', 'Beleg', 'Hinweis'];
  const example = [1, '07.06.2026', 'ARAL', '69190', 'Walldorf', 'DE', 'Super E10', 45.00, 78.50, 75000, 'AMEX', '#0218', 'ja', 'Beispieleintrag – bitte löschen'];
  const ws = XLSX.utils.aoa_to_sheet([header, example]);
  ws['!cols'] = [
    {wch:5},{wch:12},{wch:22},{wch:6},{wch:18},{wch:6},
    {wch:12},{wch:8},{wch:11},{wch:11},{wch:14},{wch:11},{wch:8},{wch:30},
  ];
  XLSX.utils.book_append_sheet(wb, ws, 'Tankvorgänge');
  XLSX.writeFile(wb, 'AutoDashboard_Tanken_Vorlage.xlsx');
}

function parseTankstelleField(raw) {
  raw = String(raw || '').trim();
  if (!raw || raw === '- unbekannt -') return [raw, '', '', ''];
  let land = '';
  const ccM = raw.match(/\s*\(([A-Z]{2})\)\s*$/);
  if (ccM) { land = ccM[1]; raw = raw.slice(0, raw.length - ccM[0].length).trim(); }
  if (raw.includes('/')) {
    const [b, loc] = raw.split('/', 2).map(s => s.trim());
    const pm = loc.match(/^(\d{4,5})\s+(.+)$/);
    if (pm) return [b, pm[1], pm[2].replace(/\.$/, ''), land || 'DE'];
    return [b, '', loc, land || 'DE'];
  }
  const pm = raw.match(/\b(\d{4,5})\b\s+(\S.*)$/);
  if (pm) return [raw.slice(0, pm.index).trim(), pm[1], pm[2].trim(), land || 'DE'];
  const tokens = raw.split(' ');
  if (tokens.length >= 2 && !raw.startsWith('-'))
    return [tokens.slice(0, -1).join(' '), '', tokens[tokens.length - 1], land || 'DE'];
  return [raw, '', '', land];
}

async function deleteTankEntry(row) {
  if (!confirm('Diesen Tankeintrag wirklich löschen?')) return;
  try {
    await api.delete('Tanken', row);
    await loadData();
    renderTankList();
    showToast('Eintrag gelöscht');
  } catch (err) {
    showToast('Fehler: ' + err.message);
  }
}

function confirmDeleteAll(sheet) {
  const sel = q('#confirm-sheet-select');
  if (sel && sheet) sel.value = sheet;
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
    const [t, k, f, km, zm] = await Promise.all([
      api.read('Tanken'),
      api.read('Kosten'),
      api.read('Fahrzeug'),
      api.read('KMStand').catch(() => ({ data: [] })),
      api.read('Zahlungsmethoden').catch(() => ({ data: [] }))
    ]);

    state.tank              = (t.data  || []).sort((a, b) => String(b.datum).localeCompare(String(a.datum)));
    state.kosten            = (k.data  || []).sort((a, b) => String(b.datum).localeCompare(String(a.datum)));
    state.kmstand           = (km.data || []).sort((a, b) => String(b.datum).localeCompare(String(a.datum)));
    state.zahlungsmethoden  = (zm.data || []);

    // Fahrzeug: erste Datenzeile direkt als Objekt
    state.fahrzeug = (f.data || [])[0] || {};

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

  const sortedYears = [...years].sort().reverse();
  const selectors   = [q('#year-filter'), q('#year-filter-mobile')].filter(Boolean);
  const current     = state.year || selectors[0]?.value || '';

  selectors.forEach(sel => {
    sel.innerHTML = `
      <optgroup label="Zeitraum">
        <option value="week">Diese Woche</option>
        <option value="month">Dieser Monat</option>
        <option value="quarter">Dieses Quartal</option>
      </optgroup>
      <optgroup label="Jahr">
        <option value="">Alle Jahre</option>
        ${sortedYears.map(y => `<option value="${y}">${y}</option>`).join('')}
      </optgroup>`;
    const defaultY = String(new Date().getFullYear());
    sel.value = current || (sortedYears.includes(defaultY) ? defaultY : sortedYears[0] || '');
  });
  state.year = selectors[0]?.value || '';
}

function currentYear() {
  const special = ['week', 'month', 'quarter'];
  return special.includes(state.year)
    ? String(new Date().getFullYear())
    : (state.year || String(new Date().getFullYear()));
}

// ─── Filter-Hilfsfunktion ─────────────────────────────────────────────────────

function filtered(arr, dateField = 'datum') {
  if (!state.year) return arr;
  if (state.year === 'week') {
    const now = new Date(); const dow = now.getDay() || 7;
    const mon = new Date(now); mon.setDate(now.getDate() - dow + 1); mon.setHours(0,0,0,0);
    const sun = new Date(mon); sun.setDate(mon.getDate() + 6); sun.setHours(23,59,59,999);
    return arr.filter(e => { const d = new Date(String(e[dateField]||'')+'T12:00:00'); return d >= mon && d <= sun; });
  }
  if (state.year === 'month') {
    const prefix = today().slice(0, 7);
    return arr.filter(e => String(e[dateField] || '').startsWith(prefix));
  }
  if (state.year === 'quarter') {
    const now = new Date(); const q0 = Math.floor(now.getMonth() / 3) * 3;
    const qStart = new Date(now.getFullYear(), q0, 1);
    const qEnd   = new Date(now.getFullYear(), q0 + 3, 0, 23, 59, 59);
    return arr.filter(e => { const d = new Date(String(e[dateField]||'')+'T12:00:00'); return d >= qStart && d <= qEnd; });
  }
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

  // KM im aktuellen Filterzeitraum (aus KMStand)
  const kmPeriod = (() => {
    const filteredKm = filtered(state.kmstand)
      .sort((a, b) => String(a.datum).localeCompare(String(b.datum)));
    if (!filteredKm.length) return 0;
    const before = state.kmstand
      .filter(e => !filteredKm.includes(e))
      .sort((a, b) => String(b.datum).localeCompare(String(a.datum)));
    const startKm = before.length
      ? Number(before[0].km_stand || 0)
      : Number(filteredKm[0].km_stand || 0);
    const endKm = Number(filteredKm[filteredKm.length - 1].km_stand || 0);
    return Math.max(0, endKm - startKm);
  })();

  // Monatliche Tankkosten (letzte 6 Monate)
  const monthly = {};
  state.tank.forEach(e => {
    const m = String(e.datum).slice(0, 7);
    const v = parseFloat(String(e.kosten || '0').replace(/[€\s ]/g, '').replace(',', '.')) || 0;
    if (m) monthly[m] = (monthly[m] || 0) + v;
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
           avgPreis, lastKm, lastKmDat, kmPeriod, monthly, last6, katMap, kraftstoffMap, yearMap };
}

// ─── Render Übersicht ─────────────────────────────────────────────────────────

const KAT_SVG = {
  'Tanken':         `<svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M19.428 15.428a2 2 0 00-1.022-.547l-2.387-.477a6 6 0 00-3.86.517l-.318.158a6 6 0 01-3.86.517L6.05 15.21a2 2 0 00-1.806.547M8 4h8l-1 1v5.172a2 2 0 00.586 1.414l5 5c1.26 1.26.367 3.414-1.415 3.414H4.828c-1.782 0-2.674-2.154-1.414-3.414l5-5A2 2 0 009 10.172V5L8 4z"/></svg>`,
  'Versicherung':   `<svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M9 12l2 2 4-4m5.618-4.016A11.955 11.955 0 0112 2.944a11.955 11.955 0 01-8.618 3.04A12.02 12.02 0 003 9c0 5.591 3.824 10.29 9 11.622 5.176-1.332 9-6.03 9-11.622 0-1.042-.133-2.052-.382-3.016z"/></svg>`,
  'Inspektion':     `<svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M10.325 4.317c.426-1.756 2.924-1.756 3.35 0a1.724 1.724 0 002.573 1.066c1.543-.94 3.31.826 2.37 2.37a1.724 1.724 0 001.065 2.572c1.756.426 1.756 2.924 0 3.35a1.724 1.724 0 00-1.066 2.573c.94 1.543-.826 3.31-2.37 2.37a1.724 1.724 0 00-2.572 1.065c-.426 1.756-2.924 1.756-3.35 0a1.724 1.724 0 00-2.573-1.066c-1.543.94-3.31-.826-2.37-2.37a1.724 1.724 0 00-1.065-2.572c-1.756-.426-1.756-2.924 0-3.35a1.724 1.724 0 001.066-2.573c-.94-1.543.826-3.31 2.37-2.37.996.608 2.296.07 2.572-1.065z"/><circle cx="12" cy="12" r="3"/></svg>`,
  'Reparatur':      `<svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z"/></svg>`,
  'KFZ-Steuer':     `<svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M9 7h6m0 10v-3m-3 3h.01M9 17h.01M9 14h.01M12 14h.01M15 11h.01M12 11h.01M9 11h.01M7 21h10a2 2 0 002-2V5a2 2 0 00-2-2H7a2 2 0 00-2 2v14a2 2 0 002 2z"/></svg>`,
  'Parken':         `<svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M8 14v3m4-3v3m4-3v3M3 21h18M3 10h18M3 7l9-4 9 4M4 10h16v11H4V10z"/></svg>`,
  'Bewohnerparken': `<svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M8 14v3m4-3v3m4-3v3M3 21h18M3 10h18M3 7l9-4 9 4M4 10h16v11H4V10z"/></svg>`,
  'Autowäsche':     `<svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm0 18c-4.41 0-8-3.59-8-8s3.59-8 8-8 8 3.59 8 8-3.59 8-8 8z"/><path stroke-linecap="round" stroke-linejoin="round" d="M8 12h8M8 8l4 4-4 4"/></svg>`,
  'Waschanlage':    `<svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2z"/></svg>`,
  'TÜV':            `<svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M9 12l2 2 4-4M7.835 4.697a3.42 3.42 0 001.946-.806 3.42 3.42 0 014.438 0 3.42 3.42 0 001.946.806 3.42 3.42 0 013.138 3.138 3.42 3.42 0 00.806 1.946 3.42 3.42 0 010 4.438 3.42 3.42 0 00-.806 1.946 3.42 3.42 0 01-3.138 3.138 3.42 3.42 0 00-1.946.806 3.42 3.42 0 01-4.438 0 3.42 3.42 0 00-1.946-.806 3.42 3.42 0 01-3.138-3.138 3.42 3.42 0 00-.806-1.946 3.42 3.42 0 010-4.438 3.42 3.42 0 00.806-1.946 3.42 3.42 0 013.138-3.138z"/></svg>`,
  'Motoröl':        `<svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M19.428 15.428a2 2 0 00-1.022-.547l-2.387-.477a6 6 0 00-3.86.517l-.318.158a6 6 0 01-3.86.517L6.05 15.21a2 2 0 00-1.806.547M8 4h8l-1 1v5.172a2 2 0 00.586 1.414l5 5c1.26 1.26.367 3.414-1.415 3.414H4.828c-1.782 0-2.674-2.154-1.414-3.414l5-5A2 2 0 009 10.172V5L8 4z"/></svg>`,
  'Ersatzteile':    `<svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M10.325 4.317c.426-1.756 2.924-1.756 3.35 0a1.724 1.724 0 002.573 1.066c1.543-.94 3.31.826 2.37 2.37a1.724 1.724 0 001.065 2.572c1.756.426 1.756 2.924 0 3.35a1.724 1.724 0 00-1.066 2.573c.94 1.543-.826 3.31-2.37 2.37a1.724 1.724 0 00-2.572 1.065c-.426 1.756-2.924 1.756-3.35 0a1.724 1.724 0 00-2.573-1.066c-1.543.94-3.31-.826-2.37-2.37a1.724 1.724 0 00-1.065-2.572c-1.756-.426-1.756-2.924 0-3.35a1.724 1.724 0 001.066-2.573c-.94-1.543.826-3.31 2.37-2.37.996.608 2.296.07 2.572-1.065z"/><circle cx="12" cy="12" r="3"/></svg>`,
  'An-/Ummeldung':  `<svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/></svg>`,
  'Sonstiges':      `<svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M5 12h.01M12 12h.01M19 12h.01"/></svg>`,
};
const KAT_COLOR = {
  'Tanken':         'text-blue-500 bg-blue-50',
  'Versicherung':   'text-emerald-500 bg-emerald-50',
  'Inspektion':     'text-orange-500 bg-orange-50',
  'Reparatur':      'text-red-500 bg-red-50',
  'KFZ-Steuer':     'text-violet-500 bg-violet-50',
  'Parken':         'text-gray-500 bg-gray-100',
  'Bewohnerparken': 'text-gray-500 bg-gray-100',
  'Autowäsche':     'text-cyan-500 bg-cyan-50',
  'Waschanlage':    'text-cyan-500 bg-cyan-50',
  'TÜV':            'text-green-500 bg-green-50',
  'Motoröl':        'text-yellow-600 bg-yellow-50',
  'Ersatzteile':    'text-orange-500 bg-orange-50',
  'An-/Ummeldung':  'text-violet-500 bg-violet-50',
  'Sonstiges':      'text-gray-500 bg-gray-100',
};

function renderDashboard() {
  const s   = calcStats();
  const cur = currentYear();
  const prv = String(Number(cur) - 1);

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

  // ── KPI: Kosten pro km & Ø Verbrauch (filter-aware) ──
  const kmFiltered = filtered(state.kmstand)
    .sort((a, b) => String(a.datum).localeCompare(String(b.datum)));
  const kmPeriodEntries = (() => {
    if (kmFiltered.length >= 2) return kmFiltered;
    const earliestDate = kmFiltered.length ? String(kmFiltered[0].datum) : null;
    const before = state.kmstand
      .filter(e => !earliestDate || String(e.datum) < earliestDate)
      .sort((a, b) => String(b.datum).localeCompare(String(a.datum)));
    return before.length ? [before[0], ...kmFiltered] : kmFiltered;
  })();
  const periodLabel = state.year === 'week'    ? 'diese Woche'
                    : state.year === 'month'   ? 'diesen Monat'
                    : state.year === 'quarter' ? 'dieses Quartal'
                    : `in ${cur}`;
  if (kmPeriodEntries.length >= 2) {
    const kmFirst = Number(kmPeriodEntries[0].km_stand);
    const kmLast  = Number(kmPeriodEntries[kmPeriodEntries.length - 1].km_stand);
    const kmDelta = kmLast - kmFirst;
    if (kmDelta > 0) {
      q('#kpi-kost-km').textContent       = num(s.totalAll / kmDelta, 2) + ' €/km';
      q('#kpi-kost-km-sub').textContent   = `${kmDelta.toLocaleString('de-DE')} km ${periodLabel}`;
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

  // ── KM im Zeitraum ──
  const kmPeriodEl = q('#km-monat-period');
  if (kmPeriodEl) kmPeriodEl.textContent = periodLabel;
  if (s.kmPeriod > 0) {
    q('#km-monat-bar').classList.remove('hidden');
    q('#km-monat-val').textContent = s.kmPeriod.toLocaleString('de-DE');
  } else {
    q('#km-monat-bar').classList.add('hidden');
  }

  renderBarChart(s);
  renderDonutChart(s);
  renderUpcomingList();
  renderAusgabenUeberblick(s);
  renderBottomKpis(s);
  renderVehicleSidebar();
  renderFahrzeugInfoBox();
  renderAusgabenIntervall();
  renderKraftstoffpreise();
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

  const subtitleEl = q('#donut-subtitle');
  if (subtitleEl) {
    const label = state.year === 'week'    ? 'Diese Woche'
                : state.year === 'month'   ? 'Dieser Monat'
                : state.year === 'quarter' ? 'Dieses Quartal'
                : `Jahr ${state.year || new Date().getFullYear()}`;
    subtitleEl.textContent = label;
  }
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
    ...s.tank.filter(e => Number(e.kosten) > 0).slice(0, 10)
      .map(e => ({ datum: e.datum, kat: 'Tanken', label: e.tankstelle || 'Tankstelle', betrag: Number(e.kosten) })),
    ...s.kosten.slice(0, 10)
      .map(e => ({ datum: e.datum, kat: e.kategorie || 'Sonstiges', label: e.beschreibung || e.kategorie, betrag: Number(e.betrag || 0) }))
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

  q('#v-modell').textContent      = f.modell || '—';
  q('#v-zulassung').textContent   = f.zulassung_datum ? String(f.zulassung_datum).slice(0, 4) : '—';
  q('#v-kraftstoff').textContent  = f.kraftsoff_standard || f.kraftstoff_standard || '—';
  q('#v-kennzeichen').textContent = f.kennzeichen || '—';

  const sub = q('#header-sub');
  const modell = f.modell || f.Modell;
  if (modell) sub.textContent = modell;
}

// ─── Liste: Tankvorgänge ──────────────────────────────────────────────────────

function renderTankList() {
  const tank     = filtered(state.tank);
  const total    = state.tank.length;
  const incAll   = state.tank.filter(e => !(Number(e.liter) > 0) || !(Number(e.kosten) > 0));
  const incView  = tank.filter(e => !(Number(e.liter) > 0) || !(Number(e.kosten) > 0));

  let countText = state.year ? `${tank.length} von ${total} Einträgen` : `${total} Einträge`;
  if (incAll.length > 0) countText += ` · ${incAll.length} unvollständig`;
  q('#tank-count').textContent = countText;

  const banner = q('#tank-incomplete-banner');
  if (incView.length > 0) {
    banner.innerHTML = `<svg class="w-4 h-4 text-amber-500 mt-0.5 flex-shrink-0" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
      <path stroke-linecap="round" stroke-linejoin="round" d="M12 9v2m0 4h.01M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/>
    </svg>
    <div>
      <p class="text-sm font-medium text-amber-800">${incView.length} Eintrag${incView.length !== 1 ? 'e fehlen' : ' fehlt'} Liter/Kosten-Daten</p>
      <p class="text-xs text-amber-600 mt-0.5">Beim Import ohne Mengen- und Kostendaten importiert. Eintrag löschen und manuell neu erfassen — oder direkt in Google Sheets korrigieren.</p>
    </div>`;
    banner.classList.remove('hidden');
  } else {
    banner.classList.add('hidden');
  }

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
          <th class="px-3 py-3"></th>
        </tr>
      </thead>
      <tbody>
        ${tank.map(e => {
          const inc   = !(Number(e.liter) > 0) || !(Number(e.kosten) > 0);
          const preisL = Number(e.liter) > 0 ? (Number(e.kosten) / Number(e.liter)) : 0;
          const rowCls = inc
            ? 'border-b border-amber-100 bg-amber-50/50 hover:bg-amber-50 transition-colors'
            : 'border-b border-slate-50 hover:bg-slate-50 transition-colors';
          return `<tr class="${rowCls}">
            <td class="px-5 py-3 text-gray-600 whitespace-nowrap">
              ${inc ? '<span class="inline-block w-2 h-2 rounded-full bg-amber-400 mr-1.5 align-middle"></span>' : ''}${dat(e.datum)}
            </td>
            <td class="px-5 py-3 text-gray-700 font-medium hidden md:table-cell max-w-xs truncate">${e.tankstelle || '—'}</td>
            <td class="px-5 py-3 hidden sm:table-cell">
              <span class="px-2 py-0.5 bg-blue-50 text-blue-600 rounded-full text-xs font-medium">${e.kraftstoff || '—'}</span>
            </td>
            <td class="px-5 py-3 text-right ${inc ? 'text-amber-500 font-medium' : 'text-gray-600'}">${inc ? '—' : num(e.liter)}</td>
            <td class="px-5 py-3 text-right text-gray-500 text-xs">${preisL ? num(preisL, 3) : '—'}</td>
            <td class="px-5 py-3 text-right font-semibold ${inc ? 'text-amber-500' : 'text-gray-800'}">${inc ? '—' : eur(e.kosten)}</td>
            <td class="px-3 py-3 text-right">
              <button onclick="deleteTankEntry(${e._row})" class="text-gray-300 hover:text-red-400 transition-colors" title="Eintrag löschen">
                <svg class="w-4 h-4" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
                  <path stroke-linecap="round" stroke-linejoin="round" d="M6 18L18 6M6 6l12 12"/>
                </svg>
              </button>
            </td>
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
                <div class="flex items-center gap-1.5">
                  <button data-erledigt="${e._row}"
                    class="text-xs px-3 py-1.5 rounded-xl bg-green-50 text-green-700 hover:bg-green-100 font-medium transition-colors whitespace-nowrap">
                    ${btnLabel}
                  </button>
                  <button data-edit="${e._row}"
                    class="text-xs px-2.5 py-1.5 rounded-xl bg-slate-100 text-gray-500 hover:bg-blue-50 hover:text-blue-600 font-medium transition-colors" title="Bearbeiten">
                    <svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
                      <path stroke-linecap="round" stroke-linejoin="round" d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z"/>
                    </svg>
                  </button>
                </div>
              </td>
            </tr>`;
        }).join('')}
      </tbody>
    </table>`;

  // Event-Delegation auf Container
  container.onclick = ev => {
    const btnE = ev.target.closest('[data-erledigt]');
    if (btnE) {
      const entry = container._termineData.find(e => e._row === Number(btnE.dataset.erledigt));
      if (entry) markTerminErledigt(entry);
      return;
    }
    const btnEd = ev.target.closest('[data-edit]');
    if (btnEd) {
      const entry = container._termineData.find(e => e._row === Number(btnEd.dataset.edit));
      if (entry) openEditTermin(entry);
    }
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

function renderFahrzeugAlter() {
  const z = state.fahrzeug.zulassung_datum;
  if (!z) return;
  const start     = new Date(z + 'T12:00:00');
  const totalDays = Math.floor((new Date() - start) / 86400000);
  const years     = Math.floor(totalDays / 365.25);
  const remDays   = totalDays - Math.floor(years * 365.25);
  const months    = Math.floor(remDays / 30.44);
  const weeks     = Math.floor((remDays - Math.floor(months * 30.44)) / 7);
  const parts     = [];
  if (years  > 0) parts.push(`${years} Jahr${years  !== 1 ? 'e'  : ''}`);
  if (months > 0) parts.push(`${months} Monat${months !== 1 ? 'e' : ''}`);
  parts.push(`${weeks} Woche${weeks !== 1 ? 'n' : ''}`);
  q('#fahrzeug-alter').textContent     = parts.join(', ');
  q('#fahrzeug-alter-sub').textContent = `Zugelassen seit ${dat(z)}`;
}

function renderFahrzeugInfoBox() {
  const f = state.fahrzeug;
  const zul = f.zulassung_datum ? String(f.zulassung_datum).slice(0, 10) : null;
  let monthsSinceZul = 0;

  if (zul) {
    const from = new Date(zul + 'T12:00:00');
    const now  = new Date();
    monthsSinceZul = Math.max(1, (now.getFullYear() - from.getFullYear()) * 12 + (now.getMonth() - from.getMonth()));
  }

  // Letzter KM-Stand
  const kmSorted = [...state.kmstand].sort((a, b) => String(b.datum).localeCompare(String(a.datum)));
  const lastKm   = kmSorted.length ? Number(kmSorted[0].km_stand).toLocaleString('de-DE') + ' km' : '—';

  // Fiktive Leasingrate = Gesamtkosten ever / Monate seit Zulassung
  const allTimeTank  = state.tank.reduce((s, e) => s + Number(e.kosten  || 0), 0);
  const allTimeKosten = state.kosten.reduce((s, e) => s + Number(e.betrag || 0), 0);
  const allTimeTotal  = allTimeTank + allTimeKosten;
  const leasingrate   = monthsSinceZul > 0 ? eur(allTimeTotal / monthsSinceZul) : '—';

  let alterText = '—';
  let alterSub  = 'Seit Erstzulassung';
  if (zul) {
    const from2     = new Date(zul + 'T12:00:00');
    const totalDays = Math.floor((new Date() - from2) / 86400000);
    const yy  = Math.floor(totalDays / 365.25);
    const rem = totalDays - Math.floor(yy * 365.25);
    const mm  = Math.floor(rem / 30.44);
    const ww  = Math.floor((rem - Math.floor(mm * 30.44)) / 7);
    const parts = [];
    if (yy > 0) parts.push(`${yy} Jahr${yy !== 1 ? 'e' : ''}`);
    if (mm > 0) parts.push(`${mm} Monat${mm !== 1 ? 'e' : ''}`);
    parts.push(`${ww} Woche${ww !== 1 ? 'n' : ''}`);
    alterText = parts.join(', ');
    alterSub  = `Zugelassen seit ${dat(zul)}`;
  }
  const ibAlter    = q('#fib-alter-val');
  const ibAlterSub = q('#fib-alter-sub');
  const ibKm       = q('#fib-km');
  const ibLeasing  = q('#fib-leasing');
  const ibTotal    = q('#fib-leasing-total');
  if (ibAlter)    ibAlter.textContent    = alterText;
  if (ibAlterSub) ibAlterSub.textContent = alterSub;
  if (ibKm)       ibKm.textContent       = lastKm;
  if (ibLeasing)  ibLeasing.textContent  = leasingrate + '/Monat';
  if (ibTotal)    ibTotal.textContent    = `Gesamtkosten ${eur(allTimeTotal)}`;
}

function renderAusgabenIntervall() {
  const intervall  = q('#intervall-filter')?.value || 'monat';
  const allEntries = [
    ...state.tank.map(e   => ({ datum: e.datum, betrag: Number(e.kosten  || 0) })),
    ...state.kosten.map(e => ({ datum: e.datum, betrag: Number(e.betrag || 0) }))
  ].filter(e => e.betrag > 0 && e.datum);
  if (!allEntries.length) return;

  const getPeriod = datum => {
    const d = new Date(datum + 'T12:00:00');
    if (intervall === 'monat')    return datum.slice(0, 7);
    if (intervall === 'quartal')  return `${d.getFullYear()}-Q${Math.ceil((d.getMonth()+1)/3)}`;
    if (intervall === 'halbjahr') return `${d.getFullYear()}-H${d.getMonth() < 6 ? 1 : 2}`;
    return datum.slice(0, 4);
  };
  const now = new Date();
  const currentPeriod = getPeriod(now.toISOString().slice(0, 10));
  const periodMap = {};
  allEntries.forEach(e => { const p = getPeriod(e.datum); periodMap[p] = (periodMap[p]||0) + e.betrag; });

  const currentTotal = periodMap[currentPeriod] || 0;
  const vals         = Object.values(periodMap);
  const avg          = vals.reduce((s, v) => s + v, 0) / vals.length;
  const labels       = { monat: 'Monat', quartal: 'Quartal', halbjahr: 'Halbjahr', jahr: 'Jahr' };

  q('#kpi-ausgaben-intervall').textContent     = eur(currentTotal);
  q('#kpi-ausgaben-intervall-sub').textContent = `Ø ${eur(Math.round(avg))} / ${labels[intervall]} (historisch)`;
}

function renderKraftstoffpreise() {
  const mitDaten   = state.tank.filter(e => Number(e.liter) > 0 && Number(e.kosten) > 0);
  const totalLiter = mitDaten.reduce((s, e) => s + Number(e.liter),  0);
  const totalKost  = mitDaten.reduce((s, e) => s + Number(e.kosten), 0);

  q('#kst-avg-liter').textContent     = totalLiter > 0 ? num(totalKost / totalLiter, 3) + ' €/l' : '—';
  q('#kst-avg-liter-sub').textContent = `aus ${mitDaten.length} Tankungen`;
  q('#kst-avg-kosten').textContent     = mitDaten.length > 0 ? eur(totalKost / mitDaten.length) : '—';
  q('#kst-avg-kosten-sub').textContent = '€ pro Tankvorgang';

  const allWithCost = state.tank.filter(e => Number(e.kosten) > 0)
    .sort((a, b) => String(a.datum).localeCompare(String(b.datum)));
  if (allWithCost.length >= 2) {
    const ms     = new Date(allWithCost.at(-1).datum+'T12:00:00') - new Date(allWithCost[0].datum+'T12:00:00');
    const months = Math.max(1, ms / (1000*60*60*24*30.44));
    const total  = allWithCost.reduce((s, e) => s + Number(e.kosten), 0);
    q('#kst-avg-monat').textContent     = eur(total / months);
    q('#kst-avg-monat-sub').textContent = `über ${Math.round(months)} Monate`;
    q('#kst-avg-jahr').textContent      = eur(total / (months / 12));
    q('#kst-avg-jahr-sub').textContent  = `über ${num(months/12, 1)} Jahre`;
  }

  const byArt   = {};
  mitDaten.forEach(e => {
    const k = e.kraftstoff || 'Unbekannt';
    if (!byArt[k]) byArt[k] = { liter: 0, kosten: 0, count: 0 };
    byArt[k].liter  += Number(e.liter);
    byArt[k].kosten += Number(e.kosten);
    byArt[k].count  += 1;
  });
  const artCol = { 'Super E10':'text-blue-600 bg-blue-50', 'Super E5':'text-green-600 bg-green-50', 'Diesel':'text-yellow-700 bg-yellow-50' };
  q('#kst-by-art').innerHTML = Object.entries(byArt).sort((a,b) => b[1].kosten - a[1].kosten).map(([art, d]) => {
    const avgL = d.liter > 0 ? d.kosten / d.liter : 0;
    const col  = artCol[art] || 'text-gray-600 bg-gray-100';
    return `<div class="flex items-center justify-between py-2 border-t border-slate-50">
      <div class="flex items-center gap-2">
        <span class="text-xs font-medium px-2 py-0.5 rounded-full ${col}">${art}</span>
        <span class="text-xs text-gray-400">${d.count} Tankungen</span>
      </div>
      <div class="flex gap-5 text-xs">
        <span class="text-gray-500">Ø <span class="font-semibold text-gray-800">${num(avgL,3)} €/l</span></span>
        <span class="text-gray-500">Gesamt <span class="font-semibold text-gray-800">${eur(d.kosten)}</span></span>
      </div>
    </div>`;
  }).join('');
}

function renderAll() {
  syncKategorieDropdown();
  syncZahlungsmethodeDropdown();
  renderDashboard();
  renderTankList();
  renderKostenList();
  renderTermineList();
  renderStatistiken();
}

function openEditTermin(entry) {
  _editKostenRow = entry._row;
  const form = q('#form-kosten');
  form.querySelector('[name="datum"]').value               = entry.datum || '';
  form.querySelector('[name="kategorie"]').value           = entry.kategorie || '';
  form.querySelector('[name="beschreibung"]').value        = entry.beschreibung || '';
  form.querySelector('[name="betrag"]').value              = entry.betrag || '';
  form.querySelector('[name="intervall"]').value           = entry.intervall || 'einmalig';
  form.querySelector('[name="naechste_faelligkeit"]').value = entry.naechste_faelligkeit || '';
  form.querySelector('[name="konto"]').value               = entry.konto || '';
  form.querySelector('[name="beleg"]').checked             = entry.beleg === 'ja';
  form.querySelector('[name="hinweis"]').value             = entry.hinweis || '';
  q('#modal-kosten-title').textContent = '✏️ Ausgabe bearbeiten';
  openModal('modal-kosten');
}

function syncKategorieDropdown() {
  const sel = q('select[name="kategorie"]');
  if (!sel) return;
  const existing = new Set([...sel.options].map(o => o.value));
  const fromData = [...new Set(state.kosten.map(e => e.kategorie).filter(Boolean))];
  fromData.forEach(kat => {
    if (!existing.has(kat)) {
      const opt = document.createElement('option');
      opt.value = opt.textContent = kat;
      sel.appendChild(opt);
    }
  });
  // Alphabetisch sortieren (Leerzeichen-Option oben lassen, falls vorhanden)
  const opts = Array.from(sel.options).filter(o => o.value);
  opts.sort((a, b) => a.text.localeCompare(b.text, 'de'));
  const placeholder = Array.from(sel.options).find(o => !o.value);
  sel.innerHTML = '';
  if (placeholder) sel.appendChild(placeholder);
  opts.forEach(o => sel.appendChild(o));
}

function syncZahlungsmethodeDropdown() {
  const sel = q('#tank-zahlungsmethode');
  if (!sel) return;
  const cur = sel.value;
  sel.innerHTML = '<option value="">— Bitte wählen —</option>';
  state.zahlungsmethoden.forEach(z => {
    const val = JSON.stringify({ konto: z.konto || '', karte: z.endziffern ? '#' + z.endziffern : '' });
    const opt = document.createElement('option');
    opt.value = val;
    opt.textContent = z.name + (z.endziffern ? ' (#' + z.endziffern + ')' : '');
    if (val === cur) opt.selected = true;
    sel.appendChild(opt);
  });
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
  q('#btn-delete-tank').addEventListener('click',   () => confirmDeleteAll('Tanken'));
  q('#btn-delete-kosten').addEventListener('click', () => confirmDeleteAll('Kosten'));

  // Bestätigungs-Modal Buttons
  q('#confirm-cancel').addEventListener('click', () => q('#modal-confirm').classList.add('hidden'));
  q('#modal-confirm').addEventListener('click', e => { if (e.target === q('#modal-confirm')) q('#modal-confirm').classList.add('hidden'); });
  q('#confirm-ok').addEventListener('click', async () => {
    const sel   = q('#confirm-sheet-select');
    const sheet = sel ? sel.value : q('#confirm-sheet').value;
    const label = sel ? sel.options[sel.selectedIndex].text : sheet;
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
  q('#intervall-filter').addEventListener('change', renderAusgabenIntervall);
  q('#btn-add-km').addEventListener('click', () => { setDefaultDates('form-km'); openModal('modal-km'); });
  q('#btn-add-tank').addEventListener('click',   () => { setDefaultDates('form-tank');   openModal('modal-tank'); });
  const openKostenModal = () => {
    _editKostenRow = null;
    const t = q('#modal-kosten-title');
    if (t) t.textContent = '📋 Ausgabe erfassen';
    q('#form-kosten').reset();
    setDefaultDates('form-kosten');
    q('#kosten-faelligkeit').value = nextDueDate(q('#kosten-datum').value, q('#kosten-intervall').value) || '';
    openModal('modal-kosten');
  };
  q('#btn-add-kosten').addEventListener('click', openKostenModal);
  q('#btn-add-termin').addEventListener('click', openKostenModal);

  // Modals schließen — edit-Modus zurücksetzen
  const closeAndReset = () => {
    closeAll();
    _editKostenRow = null;
    const t = q('#modal-kosten-title');
    if (t) t.textContent = '📋 Ausgabe erfassen';
  };
  qAll('.modal-close').forEach(btn => btn.addEventListener('click', closeAndReset));
  qAll('[id^="modal-"]').forEach(m => m.addEventListener('click', e => { if (e.target === m) closeAndReset(); }));

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

  // Form: Kosten (neu + bearbeiten)
  q('#form-kosten').addEventListener('submit', async e => {
    e.preventDefault();
    const fd  = new FormData(e.target);
    const data = {
      datum:                fd.get('datum'),
      kategorie:            fd.get('kategorie'),
      beschreibung:         fd.get('beschreibung'),
      betrag:               parseFloat(fd.get('betrag')) || 0,
      intervall:            fd.get('intervall'),
      naechste_faelligkeit: fd.get('naechste_faelligkeit') || '',
      konto:                fd.get('konto'),
      beleg:                fd.get('beleg') ? 'ja' : 'nein',
      hinweis:              fd.get('hinweis') || ''
    };
    if (_editKostenRow) {
      const btn = q('#form-kosten button[type="submit"]');
      btn.disabled = true; btn.textContent = 'Speichern…';
      try {
        await api.update('Kosten', _editKostenRow, data);
        _editKostenRow = null;
        closeAll();
        const t = q('#modal-kosten-title'); if (t) t.textContent = '📋 Ausgabe erfassen';
        toast('✅ Ausgabe aktualisiert');
        await loadData();
      } catch (err) { toast('❌ Fehler: ' + err.message, 4000); }
      finally { btn.disabled = false; btn.textContent = 'Speichern'; }
    } else {
      await submitForm('form-kosten', 'modal-kosten', 'Kosten', data, '✅ Ausgabe gespeichert');
    }
  });

  // Zahlungsmethode-Dropdown → konto/karte auto-befüllen
  const zmSel = q('#tank-zahlungsmethode');
  if (zmSel) {
    zmSel.addEventListener('change', () => {
      const val = zmSel.value;
      const kh  = q('#tank-konto-hidden');
      const kth = q('#tank-karte-hidden');
      if (!val) { if (kh) kh.value = ''; if (kth) kth.value = ''; return; }
      try {
        const { konto, karte } = JSON.parse(val);
        if (kh)  kh.value  = konto || '';
        if (kth) kth.value = karte || '';
      } catch {}
    });
  }

  // Button: neue Zahlungsmethode öffnen
  const btnZm = q('#btn-add-zahlungsmethode');
  if (btnZm) btnZm.addEventListener('click', () => openModal('modal-zahlungsmethode'));

  // Form: Zahlungsmethode
  const formZm = q('#form-zahlungsmethode');
  if (formZm) {
    formZm.addEventListener('submit', async e => {
      e.preventDefault();
      const fd = new FormData(e.target);
      const btn = formZm.querySelector('button[type="submit"]');
      btn.disabled = true; btn.textContent = 'Speichern…';
      try {
        await api.add('Zahlungsmethoden', {
          name:        fd.get('zm_name'),
          konto:       fd.get('zm_konto'),
          endziffern:  fd.get('zm_endziffern') || '',
          typ:         fd.get('zm_typ') || ''
        });
        closeAll();
        formZm.reset();
        toast('💳 Zahlungsmethode gespeichert');
        await loadData();
      } catch (err) { toast('❌ Fehler: ' + err.message, 4000); }
      finally { btn.disabled = false; btn.textContent = 'Speichern'; }
    });
  }

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

  function parseDateCell(raw) {
    if (!raw && raw !== 0) return '';
    if (typeof raw === 'number') return excelSerialToISO(raw);
    const s = String(raw).trim();
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
    const m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})/);
    if (m) return `${m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`;
    return '';
  }

  function parseExcel(workbook) {
    const ws   = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    if (!rows.length) return [];
    const f0 = String(rows[0][0] || '').toLowerCase().trim();
    const f1 = String(rows[0][1] || '').toLowerCase().trim();
    // Header-Format: erste Zelle = "datum"/"date"/"no."/"no", oder zweite = "datum"
    if (['datum','date','no.','no'].includes(f0) || f1 === 'datum' || f1 === 'date')
      return parseExcelTemplate(rows);
    return parseExcelLegacy(rows);
  }

  function parseExcelTemplate(rows) {
    const headers = rows[0].map(h => String(h || '').toLowerCase().trim());
    const ci  = name => headers.findIndex(h => h.includes(name));
    const iDat   = ci('datum');
    const iTank  = ci('tankstelle');
    const iPlz   = ci('plz');
    const iOrt   = ci('ort');
    const iLand  = ci('land');
    const iKraft = ci('kraftstoff');
    const iLit   = ci('liter');
    const iKost  = ci('kosten');
    const iKm    = ci('km');
    const iKont  = ci('konto');
    const iKar   = headers.findIndex(h => h.includes('karten') || h.includes('karte'));
    const iBel   = ci('beleg');
    const iHin   = ci('hinweis');

    const results = [];
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const datum = parseDateCell(row[iDat]);
      if (!datum) continue;
      const tankRaw = String(row[iTank] || '').trim();
      if (!tankRaw) continue;

      const clean = s => parseFloat(String(s || '').replace(/[€\s ]/g, '').replace(',', '.')) || 0;
      const liter  = clean(row[iLit]);
      const kosten = clean(row[iKost]);

      // PLZ/Ort/Land: separate Spalten nutzen; wenn leer → aus Tankstelle parsen
      let plz  = iPlz  >= 0 ? String(row[iPlz]  || '').trim() : '';
      let ort  = iOrt  >= 0 ? String(row[iOrt]  || '').trim() : '';
      let land = iLand >= 0 ? String(row[iLand] || '').trim() : '';
      let tankstelle = tankRaw;
      if (!plz && !ort) {
        [tankstelle, plz, ort, land] = parseTankstelleField(tankRaw);
      }

      // KM-Stand: Werte < 10 sind Legacy-Preis/Liter-Daten → ignorieren
      let kmStand = iKm >= 0 ? String(row[iKm] || '').trim() : '';
      const kmNum = parseFloat(kmStand.replace(',', '.').replace(/[€\s ]/g, ''));
      if (!isNaN(kmNum) && kmNum > 0 && kmNum < 10) kmStand = '';

      // Karten-ID: "None" normalisieren
      const karte = iKar >= 0 ? String(row[iKar] || '').trim().replace(/^none$/i, '') : '';

      results.push({
        datum, tankstelle, plz, ort, land,
        kraftstoff:      String(row[iKraft] || '').trim(),
        liter:           Math.round(liter  * 1000) / 1000,
        kosten:          Math.round(kosten * 100)  / 100,
        preis_pro_liter: liter > 0 ? Math.round(kosten / liter * 10000) / 10000 : '',
        km_stand:        kmStand,
        konto:           iKont >= 0 ? String(row[iKont] || '').trim() : '',
        karte,
        beleg:  iBel >= 0 && /^ja$/i.test(String(row[iBel] || '').trim()) ? 'ja' : 'nein',
        hinweis: iHin >= 0 ? String(row[iHin] || '').trim() : '',
      });
    }
    return results.sort((a, b) => a.datum.localeCompare(b.datum));
  }

  function parseExcelLegacy(rows) {
    const results = [];
    for (const row of rows) {
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
        konto:           String(row[9]  || '').trim(),
        karte:           String(row[10] || '').trim(),
        beleg:           String(row[1]  || '') === 'o' ? 'ja' : 'nein',
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
