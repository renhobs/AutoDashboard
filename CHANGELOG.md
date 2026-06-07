# Changelog

Alle bedeutsamen Änderungen werden hier dokumentiert.
Format: [Semantic Versioning](https://semver.org/) — `MAJOR.MINOR.PATCH`

---

## [Unreleased]

### Geplant
- [ ] Tankerking API Integration
- [ ] GitHub Pages Deployment (iPhone-Zugriff)

---

## [v0.3.0] — 2026-06-07

### Hinzugefügt
- `img/seat-leon.svg` — SVG-Fahrzeugbild (schwarzer Seat Leon Mk3, Seitenansicht)
- `scripts/import_tanken.py` — Import-Skript für ÜbersichtTanken.xlsx → Google Sheets (7 Batches à 15 Einträge)

### Geändert
- `index.html` — Fahrzeugkarte in Sidebar zeigt jetzt Seat Leon SVG + Stammdaten (2018, Benzin)
- `scripts/Code.gs` — `initHeaders()` (leere Sheets werden automatisch initialisiert) + `batchAdd`-Aktion für Massenimport

### Technische Notizen
- Excel-Seriennummern werden mit `datetime(1899,12,30) + timedelta(days=serial)` korrekt konvertiert
- URL-Parameter-Limit (~8k Zeichen): Import läuft in Chunks von 15 Einträgen
- batchAdd nutzt `setValues()` statt appendRow-Loop — deutlich schneller für Massenimport

---

## [v0.2.0] — 2026-06-07

### Hinzugefügt
- `index.html` — vollständiges Dashboard (Übersicht, Tanken, Kosten)
- `js/app.js` — komplette App-Logik (State, API, Render, Navigation, Modals)
- `js/config.js` — API-URL Konfiguration
- `scripts/Code.gs` — Google Apps Script Backend (GET-only, CORS-sicher)
- `.claude/launch.json` — lokaler Entwicklungsserver

### Technische Notizen
- GET-only Architektur für Apps Script vermeidet Browser-CORS-Probleme
- Google Sheet Tabs: Tanken, Kosten, Fahrzeug
- Verbindung zu Google Sheets erfolgreich getestet

---

## [v0.1.0] — 2026-06-07

### Hinzugefügt
- Initiales Projekt-Setup (CLAUDE.md, README.md, CHANGELOG.md, LEARNINGS.md)
- `.claude/settings.json` mit Basis-Berechtigungen
- Dokumentationsstruktur (docs/, memory/)
- Projektbasis definiert: Tech Stack, Architektur, Konventionen

### Technische Notizen
- Stack: Vanilla JS + Tailwind CSS (CDN) + Google Sheets/Apps Script + GitHub Pages
- Kein Build-Step, kein npm — bewusste Entscheidung für Einfachheit
- Bestehende Tankdaten (ÜbersichtTanken.xlsx) werden in Phase 1 importiert

---

<!--
TEMPLATE FÜR NEUE EINTRÄGE:

## [vX.Y.Z] — YYYY-MM-DD

### Hinzugefügt
- Neue Features beschreiben

### Geändert
- Breaking Changes oder Verhaltenänderungen

### Behoben
- Bugfixes

### Entfernt
- Gelöschte Funktionen

### Technische Notizen
- Lernpunkte, Architekturentscheidungen, Gotchas
-->
