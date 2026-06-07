#!/usr/bin/env python3
"""
Import ÜbersichtTanken.xlsx → Google Sheets (Tanken-Tab)

Verwendung:
  python3 scripts/import_tanken.py            # normaler Import
  python3 scripts/import_tanken.py --dry-run  # Vorschau, kein Upload

Voraussetzung: Code.gs muss neu deployed sein (mit batchAdd-Aktion).
API-URL wird automatisch aus js/config.js gelesen.
"""

import os
import zipfile
import xml.etree.ElementTree as ET
import urllib.request
import urllib.parse
import json
import re
import sys
from datetime import datetime, timedelta
from pathlib import Path

# macOS: Python.org-Installationen kennen kein System-Keychain — Systemzertifikat direkt setzen
if sys.platform == 'darwin' and not os.environ.get('SSL_CERT_FILE'):
    os.environ.setdefault('SSL_CERT_FILE', '/etc/ssl/cert.pem')

XLSX_PATH = Path(__file__).parent.parent / 'ÜbersichtTanken.xlsx'
CONFIG_JS = Path(__file__).parent.parent / 'js' / 'config.js'


def read_api_url():
    text = CONFIG_JS.read_text()
    m = re.search(r"API_URL\s*=\s*'([^']+)'", text)
    if not m:
        sys.exit('API_URL nicht in js/config.js gefunden.')
    return m.group(1)


def excel_serial_to_iso(serial):
    try:
        d = datetime(1899, 12, 30) + timedelta(days=float(serial))
        return d.strftime('%Y-%m-%d')
    except Exception:
        return ''


def read_xlsx(path):
    with zipfile.ZipFile(path) as z:
        ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        strings = [
            t.text or ''
            for t in ET.fromstring(z.read('xl/sharedStrings.xml')).findall('.//s:t', ns)
        ]

        def cell_val(c):
            v = c.find('s:v', ns)
            if v is None:
                return ''
            return strings[int(v.text)] if c.get('t') == 's' else v.text

        def col_letter(c):
            return ''.join(ch for ch in c.get('r', '') if ch.isalpha())

        rows = ET.fromstring(z.read('xl/worksheets/sheet1.xml')).findall('.//s:row', ns)
        entries = []

        for row in rows:
            cells = {col_letter(c): cell_val(c) for c in row.findall('s:c', ns)}

            # Datenspalte C enthält laufende Nummer (Integer) — überspringt Header/Summenzeilen
            no = cells.get('C', '')
            try:
                int(float(no))
            except (ValueError, TypeError):
                continue

            datum = excel_serial_to_iso(cells.get('D', ''))
            if not datum:
                continue

            liter  = cells.get('G', '') or '0'
            kosten = cells.get('H', '') or '0'
            try:
                l = round(float(liter),  3)
                k = round(float(kosten), 2)
                preis = round(k / l, 4) if l > 0 else ''
            except Exception:
                l, k, preis = 0, 0, ''

            entries.append({
                'datum':           datum,
                'tankstelle':      cells.get('E', ''),
                'kraftstoff':      cells.get('F', ''),
                'liter':           l,
                'kosten':          k,
                'preis_pro_liter': preis,
                'km_stand':        '',
                'konto':           cells.get('J', ''),
                'karte':           cells.get('K', ''),
                'beleg':           'ja' if cells.get('B', '') == 'o' else 'nein',
                'hinweis':         cells.get('L', ''),
            })

        return sorted(entries, key=lambda e: e['datum'])


CHUNK_SIZE = 15


def send_chunk(api_url, chunk, chunk_num, total_chunks):
    params = urllib.parse.urlencode({
        'action': 'batchAdd',
        'sheet':  'Tanken',
        'data':   json.dumps(chunk),
    })
    url = f'{api_url}?{params}'
    with urllib.request.urlopen(url, timeout=60) as resp:
        return json.loads(resp.read().decode())


def batch_upload(api_url, entries):
    chunks = [entries[i:i + CHUNK_SIZE] for i in range(0, len(entries), CHUNK_SIZE)]
    total  = len(entries)
    done   = 0

    print(f'Sende {total} Einträge in {len(chunks)} Batches…')
    for i, chunk in enumerate(chunks, 1):
        try:
            result = send_chunk(api_url, chunk, i, len(chunks))
        except urllib.error.HTTPError as e:
            print(f'  Batch {i}: HTTP-Fehler {e.code}')
            continue
        except Exception as e:
            print(f'  Batch {i}: Verbindungsfehler — {e}')
            continue

        if result.get('success'):
            done += result.get('count', len(chunk))
            first = chunk[0]['datum']
            last  = chunk[-1]['datum']
            print(f'  Batch {i}/{len(chunks)}: {first} – {last}  ✓')
        else:
            print(f'  Batch {i}: Server-Fehler — {result.get("error")}')

    print(f'\n{done}/{total} Einträge importiert.')


if __name__ == '__main__':
    dry_run = '--dry-run' in sys.argv

    print(f'Lese {XLSX_PATH.name}…')
    entries = read_xlsx(XLSX_PATH)
    print(f'{len(entries)} Tankeinträge gefunden.')

    if dry_run:
        print('\nVorschau (erste 5 Einträge):')
        for e in entries[:5]:
            print(f'  {e["datum"]}  {e["tankstelle"][:35]:<35}  {e["liter"]} L  {e["kosten"]} €  {e["kraftstoff"]}')
        print(f'  … und {max(0, len(entries)-5)} weitere')
        print('\n(Kein Upload — --dry-run aktiv)')
        sys.exit(0)

    api_url = read_api_url()
    batch_upload(api_url, entries)
