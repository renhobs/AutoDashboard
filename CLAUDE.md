# CLAUDE.md — Projekt-Kontext für Claude Code

> Dieses File wird bei jedem Claude Code-Aufruf automatisch geladen.
> Halte es aktuell, präzise und vollständig — es ist Claudes primäre Wissensquelle.

---

---

## Projektübersicht

**Name:** AutoDashboard
**Version:** v0.5.0
**Ziel:** Dashboard zur Übersicht aller laufenden Autokosten. Phase 1: Tankvorgänge + wiederkehrende Ausgaben (KFZ-Steuer, Versicherung, Inspektionen, Reparaturen) erfassen. Phase 2: Fälligkeits-Timestamps, damit bevorstehende Ausgaben antizipiert werden können.
**Status:** 🚧 In Entwicklung

---

## Modus

**Typ:** Solo-Projekt
**Versionierung:** Manuell (Archive/)
**Deadline:** "kein Deadline"
**Fokus:** alles: Speed | Codequalität | Dokumentation | Demo-tauglich

### Bei Versionierung = Git

```bash
# Datei committen
git add [datei] && git commit -m "[was + warum]"

# Feature-Branch anlegen
git checkout -b feature/[name]

# Status prüfen
git status && git log --oneline -5
```

### Bei Versionierung = Manuell (Archive/)

```bash
# Vor jeder Änderung: Sicherungskopie anlegen
cp src/[datei].js Archive/[datei]_v[VERSION].js

# Oder per Slash Command:
# /archive src/[datei].js v[VERSION]
```

---

## Tech Stack

| Layer | Technologie | Version | Bemerkung |
|-------|-------------|---------|-----------|
| Frontend | Vanilla JS + HTML | — | Kein Build-Step, kein npm |
| Styling | Tailwind CSS | CDN | Kein Build-Step |
| Backend | Google Apps Script | — | Läuft als Web App in Google-Infrastruktur |
| Datenbank | Google Sheets | — | Tabellenblatt = Datenbank, Sync über Google |
| Hosting | GitHub Pages | — | Statisch, HTTPS, kostenlos, iPhone erreichbar |
| Kraftstoffpreise | Tankerking API | — | Kostenlos, kein API-Key für Lesezugriff |

---

## Wichtige Befehle

```bash
# App lokal starten (Entwicklung)
open index.html
# oder mit Live-Reload:
python3 -m http.server 8080  # dann http://localhost:8080

# Tests ausführen
# — (keine automatisierten Tests)

# Deployment: Datei auf GitHub pushen → GitHub Pages aktualisiert automatisch
# (Einrichtung: Repository Settings → Pages → Branch: main / root)
```

---

## Regeln & Konventionen

### Allgemein
- Kommentare und Dokumentation auf **Deutsch** wahlweise auch auf **Englisch**, Code-Bezeichner auf **Englisch**
- Keine Pakete/Dependencies ohne explizite Freigabe installieren
- CHANGELOG.md nach jedem bedeutsamen Change aktualisieren

### Code-Stil
- Keine externen Abhängigkeiten ohne Rücksprache
- Alle API-Aufrufe durch den Proxy routen
- Immer Fehlerzustand in der UI anzeigen, nie stillschweigend ignorieren

### Sicherheit
- Keine echten API-Keys in Code committen → `.env` nutzen
- Keine Write-APIs ohne explizite Freigabe
- Kein Versand von Nutzerdaten an Dritte

---

## APIs & Schnittstellen

| Endpunkt | Zweck | Proxy | Auth |
|----------|-------|-------|------|
| Google Apps Script Web App URL | Lesen/Schreiben der Google Sheets Datenbank | ✗ | URL ist das Secret (nicht committen!) |
| Tankerking API (`creativecommons.tankerkoenig.de`) | Aktuelle Kraftstoffpreise in der Nähe | ✗ | API-Key (kostenlos, `.env`) |

**Wichtig:** Die Apps Script Web App URL und der Tankerking API-Key dürfen nie im Code committet werden.

---

## Bekannte Einschränkungen & offene TODOs

- [ ] Google Sheet + Apps Script muss einmalig manuell eingerichtet werden (Setup-Anleitung folgt)
- [ ] Tankerking API-Key beantragen (kostenlos unter creativecommons.tankerkoenig.de)
- [ ] Excel-Import (ÜbersichtTanken.xlsx) → CSV-Konverter ins Dashboard bauen
- [ ] Phase 2: Fälligkeits-Timestamps für wiederkehrende Ausgaben

---

## Niemals tun

- [ ] `.env`-Datei committen
- [ ] API-Keys direkt in Quellcode schreiben
- [ ] `Archive/`-Ordner löschen (bei manuellem Versionieren)
- [ ] Daten die sensibel erscheinen ohne Rücksprache auf einen Server laden

---

## Changelog-Zusammenfassung (letzte 3 Versionen)

> Changelog immer detailliert führen und bei jeder Änderung vermerken. Große Versionssprünge sollen zu einer Änderung von v1 zu v2 führen während inkrementelle Änderungen im Zehntelbereich der Versionierung gezählt werden sollen.

Vollständige Historie: [CHANGELOG.md](CHANGELOG.md)

| Version | Datum | Highlight |
|---------|-------|-----------|
| v0.5.0 | 2026-06-07 | Fälligkeits-Termine Phase 2: Erledigt-Button, Auto-Berechnung, Intervall-Auswahl |
| v0.4.0 | 2026-06-07 | Dashboard-Redesign nach Screenshot-Vorlage, Excel-Upload, GitHub Pages |
| v0.3.0 | 2026-06-07 | SVG-Fahrzeugbild, Python-Import-Skript für ÜbersichtTanken.xlsx |
| v0.2.0 | 2026-06-07 | Initiales Dashboard, App-Logik, Apps Script Backend |
