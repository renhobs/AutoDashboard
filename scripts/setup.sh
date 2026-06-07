#!/bin/bash
# Projekt-Setup-Skript
# Ausführen: bash scripts/setup.sh

set -e  # Abbrechen bei Fehler

echo "=== [Projektname] Setup ==="

# 1. Prüfe Voraussetzungen
echo "[1/4] Prüfe Voraussetzungen..."
command -v python3 >/dev/null 2>&1 || { echo "ERROR: python3 nicht gefunden"; exit 1; }
# command -v node >/dev/null 2>&1 || { echo "ERROR: node nicht gefunden"; exit 1; }

echo "     ✓ python3 $(python3 --version)"

# 2. Umgebungsvariablen
echo "[2/4] Konfiguriere Umgebungsvariablen..."
if [ ! -f .env ]; then
  cp .env.example .env
  echo "     ✓ .env erstellt — bitte mit echten Werten befüllen"
  echo "     → Öffne .env und trage deine API-Keys ein"
else
  echo "     ✓ .env bereits vorhanden"
fi

# 3. Dependencies (falls vorhanden)
echo "[3/4] Installiere Abhängigkeiten..."
# pip3 install -r requirements.txt
# npm install
echo "     ✓ Keine Abhängigkeiten (CDN-basiert)"

# 4. Archive-Ordner
echo "[4/4] Erstelle Verzeichnisstruktur..."
mkdir -p Archive
echo "     ✓ Archive/ erstellt"

echo ""
echo "=== Setup abgeschlossen ==="
echo ""
echo "Nächste Schritte:"
echo "  1. .env mit deinen API-Keys befüllen"
echo "  2. App starten: [Startbefehl]"
echo "  3. Browser öffnen: http://localhost:8080"
