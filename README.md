# Match Counter

## Yearly To-Do (zuerst erledigen)

1. In `counter.py` die Saisonparameter anpassen:
   - `TEAM_URL` (1. Mannschaft, neuer Saison-Link)
   - `SECOND_TEAM_URL` (2. Mannschaft, neuer Saison-Link)
   - `DATE_FROM` und `DATE_TO`
   - optional: `OWN_TEAM_NAME_PREFIX`, `HEADLESS`, `WAIT_MS`, `DEBUG`
2. In der Overall-Excel ein neues Tabellenblatt fuer das Jahr/die Saison anlegen (z. B. `26-27`).
3. In diesem neuen Overall-Blatt manuell pflegen:
   - Stand (Datum)
   - Saison
   - Turnier-Spiele
   - Anzahl der Spiele vor der Saison
4. `counter.py` ausfuehren und neue yearly-Excel erzeugen.
5. Jahreswerte mit `merge_excels.py` in das neue Overall-Blatt uebernehmen.
6. Ergebnis pruefen: Namen, nur Punkt-/Pokal-/Testspiele, Sortierung nach Nachname.

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
- `merge_excels.py` - übernimmt Werte aus yearly-Excel in ein gewähltes Overall-Blatt.
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

### 1) Virtuelle Umgebung

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

## Jahresablauf für die Overall-Excel

Für jedes neue Jahr muss in der Overall-Datei ein neues Tabellenblatt angelegt werden.

Wichtig: Die Links zu den Mannschaften in `counter.py` muessen jedes Jahr auf die neue Saison aktualisiert werden (`TEAM_URL` und `SECOND_TEAM_URL`).

Empfohlener Ablauf pro Saison:

1. In der Overall-Excel ein neues Blatt für das Jahr/die Saison erstellen (z. B. `2026_2027`).
2. In diesem neuen Blatt folgende Werte manuell pflegen:
   - Stand (Datum, auf welchem Datenstand die Werte basieren)
   - Saison
   - Turnier-Spiele
   - Anzahl der Spiele vor der Saison
3. Danach die Jahreswerte aus der yearly-Datei in dieses Blatt übernehmen (mit `merge_excels.py`).
4. Abschließend prüfen, ob Namen, Summen und Sortierung nach Nachname korrekt sind.

Hinweis: Das Skript übernimmt nur die drei Spielspalten (Punktspiele, Pokalspiele, Testspiele/Freundschaftsspiele). Die oben genannten Metadaten bleiben bewusst manuell.

Beispiel:

```powershell
python merge_excels.py --yearly-file player_stats/stats_2025_2026.xlsx --overall-file player_stats/_stats_overall.xlsx --overall-sheet 2025_2026
```

## Was jedes Jahr in den Python-Dateien angepasst werden muss

### `counter.py`

Vor jedem neuen Saisonlauf prüfen/ändern:

1. `TEAM_URL` auf die neue Saison der 1. Mannschaft setzen.
2. `SECOND_TEAM_URL` auf die neue Saison der 2. Mannschaft setzen (oder leer lassen, falls nicht benötigt).
3. `DATE_FROM` und `DATE_TO` auf den gewünschten Saison-Zeitraum setzen.
4. `OWN_TEAM_NAME_PREFIX` prüfen (nur ändern, wenn sich der Teamname geändert hat).
5. Optional Laufparameter anpassen:
   - `HEADLESS`
   - `WAIT_MS`
   - `DEBUG`

Typischer Jahreslauf:

1. `counter.py` starten und neue yearly-Excel erzeugen.
2. Ergebnisdatei inhaltlich kurz prüfen (`Zusammenfassung`, `Einzelereignisse`).

### `merge_excels.py`

Am Skript selbst ist jährlich meist keine Codeänderung nötig. Pro Jahr relevant sind vor allem die Eingaben:

1. Neue yearly-Datei (`--yearly-file`) angeben.
2. Overall-Datei (`--overall-file`) angeben.
3. Neues Jahresblatt in der Overall-Datei als Zielblatt (`--overall-sheet`) angeben.
4. Bei abweichendem Blattnamen in der yearly-Datei `--yearly-sheet` setzen.

Nur bei Strukturänderungen in Excel-Dateien (abweichende Spaltenüberschriften) muss der Alias-Abgleich im Skript erweitert werden.

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
