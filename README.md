# Match Counter

Dieses Repository enthält ein Python-Skript, das Spielerdaten aus fussball.de für die 1. und 2. Mannschaft sammelt und als Statistik in eine Excel-Datei exportiert.

## Was das Skript macht

`counter.py` führt automatisiert folgende Schritte aus:

1. Öffnet die Team-Spielpläne auf fussball.de.
2. Setzt einen Datumsbereich (`DATE_FROM` bis `DATE_TO`).
3. Lädt alle verfügbaren Spiele mit Ergebnis.
4. Öffnet jedes Spiel und liest die Aufstellung der eigenen Mannschaft.
5. Erfasst pro Spieler Einsätze nach Spieltyp:
   - Punktspiel
   - Pokalspiel
   - Freundschaftsspiel
6. Führt Daten aus 1. und 2. Mannschaft zusammen.
7. Exportiert die Ergebnisse in eine Excel-Datei (`.xlsx`).

## Projektstruktur

- `counter.py` - Hauptskript für Scraping, Auswertung und Export.
- `player_stats/` - vorhandener Projektordner.
- `yearly_stats/` - wird beim Lauf automatisch erstellt (Exportdateien).

## Voraussetzungen

- Python 3.10+ (empfohlen)
- Windows, Linux oder macOS
- Internetzugang

Benötigte Python-Pakete:

- `playwright`
- `openpyxl`

## Installation

### 1) Virtuelle Umgebung (empfohlen)

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

### 2) Abhängigkeiten installieren

```powershell
pip install playwright openpyxl
python -m playwright install chromium
```

## Konfiguration

Die wichtigsten Einstellungen stehen direkt oben in `counter.py`:

- `TEAM_URL` - URL der 1. Mannschaft
- `SECOND_TEAM_URL` - URL der 2. Mannschaft (leer lassen, um zu deaktivieren)
- `OWN_TEAM_NAME_PREFIX` - Präfix des Teamnamens zur Zuordnung Home/Away
- `DATE_FROM` - Startdatum im Format `TT.MM.JJJJ`
- `DATE_TO` - Enddatum im Format `TT.MM.JJJJ`
- `HEADLESS` - `True` für unsichtbaren Browser, `False` mit Browserfenster
- `WAIT_MS` - Wartezeit zwischen wichtigen UI-Schritten
- `DEBUG` - zusätzliche Debug-Ausgaben in der Konsole

## Ausführung

```powershell
python counter.py
```

## Ausgabe

Nach erfolgreichem Lauf wird eine Datei erzeugt in:

- `yearly_stats/stats_<startjahr>_<endjahr>.xlsx`

Die Excel-Datei enthält zwei Tabellenblätter:

1. `Zusammenfassung`
   - Spieler
   - Gesamt
   - Punktspiel
   - Pokalspiel
   - Freundschaftsspiel
2. `Einzelereignisse`
   - Spieler
   - Datum
   - Typ
   - Mannschaft-Priorität

## Hinweise

- Das Skript ist auf aktuelle Seitenstrukturen von fussball.de ausgelegt. Änderungen am HTML können Anpassungen im Code erforderlich machen.
- Bei Cookie-Banner- oder UI-Abweichungen hilft häufig `HEADLESS = False` und `DEBUG = True`.
- Der Datumsbereich wirkt sich direkt auf Laufzeit und Ergebnisumfang aus.

## Fehlerbehebung

### Playwright-Browser fehlt

Falls ein Fehler zu fehlenden Browser-Binaries erscheint:

```powershell
python -m playwright install chromium
```

### PowerShell blockiert Aktivierung

Falls die Aktivierung der venv nicht erlaubt ist:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

Dann erneut aktivieren:

```powershell
.\.venv\Scripts\Activate.ps1
```

## Lizenz

Aktuell ist keine Lizenzdatei hinterlegt. Ergänze bei Bedarf eine `LICENSE`-Datei.
