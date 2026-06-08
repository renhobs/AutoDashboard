# Changelog

Alle bedeutsamen Änderungen werden hier dokumentiert.
Format: [Semantic Versioning](https://semver.org/) — `MAJOR.MINOR.PATCH`

---

## [Unreleased]

### Geplant
- [ ] Tankerking API Integration (Kraftstoffpreise in der Nähe)
- [ ] KM-Stand-Erfassung beim Tanken aktivieren
- [ ] Tankbelege-Upload (benötigt POST-Endpoint / Google Drive Integration)

---

## [v0.8.1] — 2026-06-08

### Hinzugefügt
- **Tankstellen-Struktur erweitert** — neue Felder `plz`, `ort`, `land` pro Tankeintrag
- **Import-Template** aktualisiert: Spalten `No.`, `PLZ`, `Ort`, `Land`, `Karten-ID` (14 Spalten statt 10)
- **`parseTankstelleField()`** — Parser zerlegt alte `"Brand/ PLZ Ort (CC)"` Notation automatisch in Einzelfelder:
  - `"BFT/ 58762 Altena"` → Tankstelle=BFT, PLZ=58762, Ort=Altena, Land=DE
  - `"Eni/ 6175 Kematen (AT)"` → Land=AT
  - `"HEM Rauenberg"` (kein Slash) → Ort=Rauenberg per Last-Token-Heuristik
- **Daten bereinigt** (`AutoDashboard_Tanken_Bereinigt.xlsx`) — alle 99 Einträge mit aufgeteilten PLZ/Ort/Land Spalten

### Geändert
- `parseExcelTemplate()` — liest PLZ/Ort/Land-Spalten direkt; wenn leer → automatisches Parsing aus Tankstelle
- `parseExcelTemplate()` — erkennt `Karten-ID`-Header, normalisiert `"None"` → leer
- `parseExcelTemplate()` — KM-Stand < 10 wird als Legacy-Preis/Liter-Wert ignoriert (nicht als km_stand gespeichert)
- `parseExcel()` — Format-Erkennung erkennt jetzt auch `"No."` in Spalte A als neues Template-Format
- `downloadImportTemplate()` — erzeugt Template mit 14 Spalten im neuen Format

### Technische Notizen
- Neues Google-Sheet-Tanken muss 3 neue Spaltenheader bekommen: `plz`, `ort`, `land`
- Legacy-J-Spalte (preis_pro_liter) wird beim Import automatisch verworfen und neu aus liter/kosten berechnet

---

## [v0.8.0] — 2026-06-08

### Hinzugefügt
- **Filter auf Wochenbasis / Monatsbasis** — Jahresfilter erweitert um "Diese Woche", "Dieser Monat", "Dieses Quartal" als optgroup im Select
  - `filtered()` behandelt Sonderwerte `'week'`, `'month'`, `'quarter'` mit Datumsgrenzen
  - `currentYear()` — Helper für KPI-Berechnungen die immer 4-stellige Jahreszahl brauchen
- **Fahrzeug Info-Box im Dashboard** — kompakte Übersicht mit 3 Kennzahlen:
  - Fahrzeugalter (Jahre + Monate seit Erstzulassung)
  - Letzter Kilometerstand
  - Fiktive Leasingrate (Gesamtkosten all-time / Monate seit Zulassung)
- **Zahlungsmethoden** — neues Sheet `Zahlungsmethoden` in Google Sheets
  - Felder: `name`, `konto`, `endziffern`, `typ`
  - Tankform: Konto/Karte-Felder ersetzt durch Zahlungsmethode-Dropdown mit Auto-Fill
  - "+ Neu"-Button öffnet Zahlungsmethode-Modal direkt aus der Tankform
  - `syncZahlungsmethodeDropdown()` befüllt Dropdown aus `state.zahlungsmethoden`
- **Termine bearbeiten** — Bearbeiten-Button (Stift-Icon) pro Zeile in Termine-Liste
  - Öffnet Kosten-Modal vorausgefüllt mit bestehenden Werten
  - Speichert via `api.update()` statt `api.add()`
  - Modal-Titel wechselt zu "✏️ Ausgabe bearbeiten" im Edit-Modus
- **Kategorien alphabetisch** — Kategorie-Dropdown im Kosten-Formular alphabetisch sortiert
- **Tanken-Icon** — neues Fuel-Pump-Icon in Sidebar und Mobile-Navigation

### Geändert
- `renderDashboard()` verwendet jetzt `currentYear()` statt raw `state.year` für YoY-Vergleich
- `openKostenModal()` setzt `_editKostenRow = null` zurück beim Neu-Erstellen
- Modal-Close-Handler setzt Edit-Modus zurück

### Technische Notizen
- `state.zahlungsmethoden[]` mit `.catch()` Fallback falls Sheet noch nicht existiert
- `_editKostenRow` als Modul-Level-Variable steuert Add vs. Update im Kosten-Formular
- Zahlungsmethode-Wert im Dropdown als JSON: `{"konto":"AMEX","karte":"#0218"}`
- **Hinweis:** `Zahlungsmethoden`-Sheet muss manuell in Google Sheets angelegt werden (Header: `name`, `konto`, `endziffern`, `typ`)
- **Hinweis:** Tankbelege-Upload (Feature-Request) benötigt `doPost` in Code.gs + Storage-Lösung (Google Drive) — komplexer Architektur-Entscheid, zurückgestellt

---

## [v0.7.1] — 2026-06-08

### Behoben
- **Tankvorgänge-Darstellung** — unvollständige Einträge (liter=0/kosten=0) klar sichtbar gemacht
  - Zähler zeigt jetzt `"14 von 99 Einträgen · 15 unvollständig"` statt nur `"14 Einträge"`
  - Amber-Warnbanner erscheint wenn unvollständige Einträge in der aktuellen Ansicht vorhanden sind
  - Unvollständige Zeilen haben amber Hintergrund + orangener Punkt + `—` statt `0,00`
  - `×`-Löschen-Button pro Zeile — direkt aus der Liste einzelne Einträge entfernen

### Hinzugefügt
- `api.delete(sheet, row)` — einzelne Zeile per Row-Nummer löschen
- `deleteTankEntry(row)` — globale Funktion mit Bestätigung + Reload

### Technische Notizen
- Ursache: Excel-Import las `row[6]`/`row[7]` (Spalte G/H) — Zellen waren in diesen Zeilen leer
- Workaround: unvollständige Einträge löschen und manuell neu erfassen (× Button), oder direkt in Google Sheets korrigieren

---

## [v0.7.0] — 2026-06-08

### Hinzugefügt
- **Fahrzeugalter-Karte** — zeigt Alter in Jahren, Monaten und Wochen seit Erstzulassung
- **Ausgaben pro Intervall** — Karte mit wählbarem Filter (diese Woche / diesen Monat / dieses Quartal / dieses Jahr / alles)
  - Zeigt aktuelle Periode und historischen Durchschnitt im Vergleich
- **Kraftstoffpreise-Statistiken** — neue Karte im Dashboard
  - Ø €/L, Ø Kosten/Betankung, Ø Kosten/Monat, Ø Kosten/Jahr
  - Aufschlüsselung nach Kraftstoffart (Super E5 / Super E10)
- **KM-Stand-Tracking** — dediziertes `KMStand`-Sheet in Google Sheets
  - Eingabe über Modal (Datum, KM-Stand, Notiz)
  - KPI-Karten (Kosten/km, Ø Verbrauch) nutzen KMStand-Daten statt Tank-Einträgen
  - Jahresfilter-sensitiv mit Fallback auf letzten bekannten Stand

### Technische Notizen
- `state.kmstand[]` mit `.catch(() => ({data:[]}))` — graceful fallback wenn Sheet fehlt
- `renderFahrzeugAlter()`, `renderAusgabenIntervall()`, `renderKraftstoffpreise()` — neue Render-Funktionen

---

## [v0.6.1] — 2026-06-08

### Behoben
- **0€-Label** — `renderAusgabenUeberblick()` filtert jetzt Einträge mit `kosten=0` heraus
- **Löschen-Modal** — Sheet-Auswahl (Tanken/Ausgaben/KM-Stände) per Dropdown statt fixer Zuweisung
- **Kostenstruktur** — `syncKategorieDropdown()` synchronisiert Kategorie-Dropdown dynamisch mit DB-Werten
- **Fahrzeugdaten** — Sheet im Zeilenformat geparst: `state.fahrzeug = (f.data||[])[0]||{}`; Tippfehler `kraftsoff_standard` abgefangen

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
