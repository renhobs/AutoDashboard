# Changelog

Alle bedeutsamen Änderungen werden hier dokumentiert.
Format: [Semantic Versioning](https://semver.org/) — `MAJOR.MINOR.PATCH`

---

## [Unreleased]

### Geplant
- [ ] Tankerking API Integration (Kraftstoffpreise in der Nähe)
- [ ] KM-Stand-Erfassung beim Tanken aktivieren

---

## [v0.5.1] — 2026-06-08

### Hinzugefügt
- **Export-Funktion** — alle Daten als `.xlsx`-Datei herunterladen (SheetJS)
  - Button "Export" in Tanken- und Ausgaben-Ansicht
  - Zwei Tabellenblätter: Tanken + Ausgaben, Dateiname mit Datum
- **Alles löschen** — alle Datensätze eines Sheets auf einmal entfernen
  - Bestätigungs-Modal mit Warntext vor dem Löschen
  - Schließbar per Klick außerhalb des Modals oder "Abbrechen"
  - Toast-Meldung nach erfolgreichem Löschen
- `api.deleteAll()` — neuer API-Aufruf für `action=deleteAll`
- `showToast(msg)` — Kurzmeldungs-Einblendung unten rechts
- **Code.gs: `deleteAll`-Aktion** — löscht alle Zeilen ab Zeile 2 (Header bleibt erhalten)

### Geändert
- `index.html` — Bestätigungs-Modal `#modal-confirm` hinzugefügt
- Cache-Buster für `app.js` aktualisiert

### Technische Notizen
- `deleteAll` in GAS: `ws.deleteRows(2, lastRow - 1)` — Header-Zeile wird nie gelöscht
- Code.gs muss in Google Apps Script neu deployt werden damit `deleteAll` wirkt

---

## [v0.5.0] — 2026-06-07

### Hinzugefügt
- **Fälligkeits-Termine Phase 2** — vollständige Termin-Verwaltung mit Wiederholungslogik
  - `nextDueDate()` — berechnet nächste Fälligkeit aus Datum + Intervall
  - "Erledigt"-Button in Termine-Tabelle: markiert als bezahlt und schreibt Fälligkeit automatisch fort
  - Preview des nächsten Datums direkt im Button: "Erledigt → 07.6.2027"
  - SVG-Icons + Kategorie-Farben in der Termine-Liste
  - Intervall-Spalte in der Tabelle (monatlich / quartalsweise / halbjährlich / jährlich / 2-jährlich)
- **Modal-Verbesserungen (Ausgabe erfassen)**:
  - 5 Intervall-Optionen: Einmalig, Monatlich, Quartalsweise, Halbjährlich, Jährlich (default), Alle 2 Jahre
  - Auto-Berechnung der Nächsten Fälligkeit beim Öffnen und bei Änderung von Datum/Intervall
  - "aus Datum berechnen"-Link für manuelle Neuberechnung
  - `naechste_faelligkeit` wird nicht mehr mit "heute" vorausgefüllt
- `api.update()` — REST-Update für bestehende Kosten-Einträge in Google Sheets

### Geändert
- `renderTermineList()` — Event-Delegation statt Inline-Handler, neue Tabellenspalten
- `setDefaultDates()` — befüllt `naechste_faelligkeit` nicht mehr automatisch mit heute

### Technische Notizen
- `INTERVALL_MONTHS` Map für Monat-Berechnung aller Intervall-Typen
- `markTerminErledigt()` schreibt `naechste_faelligkeit` per `api.update()` fort, kein neuer Datensatz

---

## [v0.4.0] — 2026-06-07

### Hinzugefügt
- **Dashboard-Redesign** — komplett neues Layout nach Screenshot-Vorlage
  - Header-Bereich mit Fahrzeugbild, Titel und Jahres-Filter
  - 4 KPI-Karten: Kilometerstand, Gesamtausgaben (inkl. YoY-Trend), Kosten/km, Ø Verbrauch
  - 3-spaltig (Desktop): Tanken & Fahrleistung | Kostenstruktur | Ausgaben im Überblick
  - Balkendiagramm mit Wertebeschriftung über den Balken
  - Donut-Chart mit Mitte-Beschriftung und Legende
  - Tabelle "Ausgaben im Überblick" mit SVG-Icons pro Kategorie
  - Unterer Bereich: Termine + 4 kleine KPI-Karten (Versicherung, Sonstiges, Wartung YTD, Letzter Service)
- Excel-Upload-Funktion im Tanken-View (SheetJS, Duplikat-Erkennung, Batch-Import)
- GitHub Pages Deployment (öffentliches Repo renhobs/AutoDashboard)
- Mobile Header (Jahr-Filter + Refresh synchronisiert mit Desktop-Header)

### Geändert
- `js/app.js` — `q` und `qAll` an den Dateianfang verschoben (TDZ-Fix)
- `js/app.js` — `renderDashboard()` komplett neu geschrieben für neue Element-IDs
- `js/app.js` — `renderBarChart()` mit custom Plugin für Wert-Labels
- `js/app.js` — `renderUpcomingList()` mit SVG-Icons und farbigen Badges
- `js/app.js` — `buildYearFilter()` befüllt jetzt auch Mobile-Select synchron
- `index.html` — `view-dashboard` vollständig ersetzt mit neuem Layout

### Behoben
- Chart.js Canvas-Timing: `setLoading(false)` vor `renderAll()` verhindert 0×0-Canvas

### Technische Notizen
- KAT_SVG/KAT_COLOR für konsistente Icon+Farb-Zuordnung pro Kategorie
- `lastEntryWithData`: letzter Tankeintrag mit liter > 0 statt chronologisch letztem
- YoY-Trend berechnet aus Vorjahresdaten im State (kein separater API-Call)

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
