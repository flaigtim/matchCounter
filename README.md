# Match Counter

Dieses Projekt automatisiert den kompletten Saison-Workflow:

1. Daten von fussball.de laden
2. Jahreswerte in die Overall-Excel uebernehmen
3. Urkunden als DOCX erstellen
4. DOCX in PDF konvertieren und DOCX loeschen

## Projektstruktur

- `py_files/counter.py`: Scraping und yearly-Excel Export
- `py_files/merge_excels.py`: Merge yearly-Excel in Overall-Excel
- `py_files/create_certificate.py`: Urkunden als DOCX erstellen
- `py_files/convert_certificates_to_pdf.py`: DOCX -> PDF und DOCX loeschen
- `py_files/run_all.py`: Orchestrator fuer den gesamten Ablauf
- `player_stats/`: Excel- und Zertifikatsdaten

## Voraussetzungen

- Python 3.10+
- Microsoft 365 / Word (fuer PDF-Konvertierung via docx2pdf)
- Internetzugang (fussball.de)

Benötigte Pakete:

- `playwright`
- `openpyxl`
- `docx2pdf`

Installation:

```powershell
pip install playwright openpyxl docx2pdf
python -m playwright install chromium
```

## Zentraler Ablauf (empfohlen)

Konfiguration erfolgt zentral in `py_files/run_all.py`.

Wichtige Config-Bloecke:

1. `COUNTER_CONFIG`
   - `first_team_url`
   - `second_team_url`
   - `own_team_name_prefix`
   - `date_from`, `date_to`
   - `headless`, `wait_ms`, `debug`
2. `MERGE_CONFIG`
   - `yearly`, `overall`, `new_sheet`, `prev_sheet`
3. `CERT_CONFIG`
   - `date`, `sheet`, `overall`, `template`, `cert_outdir`
4. `PDF_CONFIG`
   - `pdf_input_dir`

Ausfuehrung ohne Parameter:

```powershell
python py_files/run_all.py
```

`run_all.py` fuehrt diese Reihenfolge aus:

1. `counter.py`
2. `merge_excels.py`
3. `create_certificate.py`
4. `convert_certificates_to_pdf.py`

Am Ende wird die Anzahl der vorhandenen PDF-Dateien im Zielordner ausgegeben.

## Einzelskripte (optional)

Alle Skripte bleiben auch einzeln nutzbar.

### 1) Daten von fussball.de laden

```powershell
python py_files/counter.py
```

Optionen (Auszug):

- `--first-team-url` (Alias: `--team-url`)
- `--second-team-url`
- `--date-from`, `--date-to`
- `--headless`, `--wait-ms`, `--debug`

### 2) yearly in overall mergen

```powershell
python py_files/merge_excels.py --yearly player_stats/stats_2025_2026.xlsx --overall player_stats/_stats_overall.xlsx --new-sheet 25-26 --prev-sheet 24-25
```

### 3) Zertifikate als DOCX erstellen

```powershell
python py_files/create_certificate.py --date 06.06.2026 --sheet 25-26 --overall player_stats/_stats_overall.xlsx --template template_certificate.docx --outdir player_stats/certificates/25-26
```

### 4) DOCX -> PDF konvertieren

```powershell
python py_files/convert_certificates_to_pdf.py --input-dir player_stats/certificates/25-26
```

Verhalten:

- erzeugt PDF mit gleichem Basisnamen
- loescht die jeweilige DOCX nur bei erfolgreicher Konvertierung

## Jahres-Checkliste

1. In `py_files/run_all.py` Saisonwerte anpassen:
   - `COUNTER_CONFIG.first_team_url`
   - `COUNTER_CONFIG.second_team_url`
   - `COUNTER_CONFIG.date_from`, `COUNTER_CONFIG.date_to`
   - `MERGE_CONFIG.yearly`, `MERGE_CONFIG.new_sheet`, `MERGE_CONFIG.prev_sheet`
   - `CERT_CONFIG.date`, `CERT_CONFIG.sheet`
2. Sicherstellen, dass Overall-Excel das Vorsaison-Blatt und Zielblatt korrekt hat.
3. `python py_files/run_all.py` ausfuehren.
4. Ergebnis pruefen (Excel + PDFs).

## Fehlerbehebung

### Keine PDF-Erstellung

1. Prüfen, ob Microsoft Word installiert und startbar ist.
2. Prüfen, ob `docx2pdf` installiert ist:

```powershell
python -m pip install docx2pdf
```

3. Sicherstellen, dass Dateien nicht in Word geoeffnet sind.

### Playwright-Fehler

```powershell
python -m playwright install chromium
```
