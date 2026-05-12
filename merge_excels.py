from __future__ import annotations

from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.formula.translate import Translator
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font
from openpyxl.styles import PatternFill


# =========================================
# KONFIGURATION (hier jaehrlich anpassen)
# =========================================
YEARLY_FILE = "player_stats/stats_2025_2026.xlsx"
OVERALL_FILE = "player_stats/_stats_overall.xlsx"
NEW_SHEET    = "25-26"   # Name des neuen Tabellenblatts
PREV_SHEET   = "24-25"   # Vorsaison-Blatt, das kopiert wird


def normalize_name(value: str) -> str:
    return " ".join(value.strip().casefold().split())


def to_int(value: object) -> int:
    if value is None:
        return 0
    if isinstance(value, (int, float)):
        return int(value)
    text = str(value).strip().replace(",", ".")
    if not text:
        return 0
    try:
        return int(float(text))
    except ValueError:
        return 0


def read_yearly_players(yearly_path: Path) -> list[tuple[str, int, int, int]]:
    if not yearly_path.exists():
        raise FileNotFoundError(f"Yearly-Datei nicht gefunden: {yearly_path}")

    yearly_wb = openpyxl.load_workbook(yearly_path, data_only=True)
    sheet = yearly_wb["Zusammenfassung"] if "Zusammenfassung" in yearly_wb.sheetnames else yearly_wb.active

    # Bevorzugt explizite Spalten aus der Kopfzeile.
    name_col = None
    punkt_col = None
    pokal_col = None
    test_col = None
    for col in range(1, sheet.max_column + 1):
        value = sheet.cell(row=1, column=col).value
        if not value:
            continue
        header = str(value).strip().lower()
        if header in {"spieler", "name", "spielername"}:
            name_col = col
        if header in {"punktspiel", "punktspiele", "puntkspiel"}:
            punkt_col = col
        if header in {"pokalspiel", "pokalspiele"}:
            pokal_col = col
        if header in {"testspiel", "testspiele", "freundschaftsspiel", "freundschaftsspiele"}:
            test_col = col

    if name_col is None:
        name_col = 1
    if punkt_col is None:
        raise ValueError("In der yearly-Datei wurde keine Spalte 'Punktspiel' gefunden.")
    if pokal_col is None:
        raise ValueError("In der yearly-Datei wurde keine Spalte 'Pokalspiel' gefunden.")
    if test_col is None:
        raise ValueError("In der yearly-Datei wurde keine Spalte 'Testspiel' gefunden.")

    players: list[tuple[str, int, int, int]] = []
    seen: set[str] = set()
    for row in range(2, sheet.max_row + 1):
        name_value = sheet.cell(row=row, column=name_col).value
        if not name_value:
            continue
        name = str(name_value).strip()
        if not name:
            continue
        key = name.casefold()
        if key in seen:
            continue
        seen.add(key)
        punkt = to_int(sheet.cell(row=row, column=punkt_col).value)
        pokal = to_int(sheet.cell(row=row, column=pokal_col).value)
        test = to_int(sheet.cell(row=row, column=test_col).value)
        players.append((name, punkt, pokal, test))

    return players


def read_prev_name_rows(prev_ws: openpyxl.worksheet.worksheet.Worksheet) -> dict[str, int]:
    first_data_row = 8
    name_to_row: dict[str, int] = {}

    for row in range(first_data_row, prev_ws.max_row + 1):
        value = prev_ws.cell(row=row, column=1).value
        if not value:
            continue
        key = normalize_name(str(value))
        if not key or key in name_to_row:
            continue
        name_to_row[key] = row

    return name_to_row


def ensure_row_formulas(ws: openpyxl.worksheet.worksheet.Worksheet, row: int, first_data_row: int) -> None:
    # G (Gesamt)
    g_cell = ws.cell(row=row, column=7)
    if not (isinstance(g_cell.value, str) and g_cell.value.startswith("=")):
        if row > first_data_row:
            src_row = row - 1
            src_formula = ws.cell(row=src_row, column=7).value
            if isinstance(src_formula, str) and src_formula.startswith("="):
                g_cell.value = Translator(src_formula, origin=f"G{src_row}").translate_formula(f"G{row}")
        if not (isinstance(g_cell.value, str) and g_cell.value.startswith("=")):
            g_cell.value = f"=C{row}+D{row}+E{row}+F{row}"

    # H (Anzahl Spiele nach Saison)
    h_cell = ws.cell(row=row, column=8)
    if not (isinstance(h_cell.value, str) and h_cell.value.startswith("=")):
        if row > first_data_row:
            src_row = row - 1
            src_formula = ws.cell(row=src_row, column=8).value
            if isinstance(src_formula, str) and src_formula.startswith("="):
                h_cell.value = Translator(src_formula, origin=f"H{src_row}").translate_formula(f"H{row}")
        if not (isinstance(h_cell.value, str) and h_cell.value.startswith("=")):
            h_cell.value = f"=G{row}+B{row}"


def build_ehrung_formula(row: int) -> str:
    # Liefert den hoechsten Schwellwert aus J3:R3, der zwischen B (vorher) und H (nachher) ueberschritten wurde.
    return f"=IFERROR(LOOKUP(2,1/((B{row}<$J$3:$R$3)*(H{row}>$J$3:$R$3)),$J$3:$R$3),\"\")"


def main() -> None:
    yearly_path = Path(YEARLY_FILE)
    overall_path = Path(OVERALL_FILE)

    if not overall_path.exists():
        raise FileNotFoundError(f"Overall-Datei nicht gefunden: {overall_path}")

    overall_wb = openpyxl.load_workbook(overall_path)

    if PREV_SHEET not in overall_wb.sheetnames:
        raise ValueError(
            f"Vorsaison-Blatt '{PREV_SHEET}' nicht gefunden. "
            f"Vorhanden: {', '.join(overall_wb.sheetnames)}"
        )

    if NEW_SHEET in overall_wb.sheetnames:
        del overall_wb[NEW_SHEET]

    src = overall_wb[PREV_SHEET]
    prev_name_rows = read_prev_name_rows(src)
    ws = overall_wb.copy_worksheet(src)
    ws.title = NEW_SHEET
    ws.freeze_panes = "B8"

    yearly_players = read_yearly_players(yearly_path)

    # B1:D1 mit Saisontext aus dem Blattnamen fuellen (z. B. 25-26 -> Saison 2025/26).
    try:
        season_xx, season_yy = NEW_SHEET.split("-", 1)
    except ValueError as exc:
        raise ValueError(
            f"NEW_SHEET muss das Format 'xx-yy' haben, erhalten: {NEW_SHEET}"
        ) from exc
    ws["B1"] = f"Saison 20{season_xx}/{season_yy}"
    ws["B4"] = f"Stand {datetime.now().strftime('%d.%m.%Y')}"

    # Bereiche A8:F<letzte Zeile> und I8:I<letzte Zeile> leeren.
    first_data_row = 8
    last_row = max(ws.max_row, first_data_row + len(yearly_players) - 1)
    for row in range(first_data_row, last_row + 1):
        for col in range(1, 7):  # A bis F
            ws.cell(row=row, column=col).value = None
        ws.cell(row=row, column=9).value = None  # I

    # Namen aus der yearly-Datei in Spalte A schreiben.
    missing_prev_rows: list[int] = []
    for idx, (name, punktspiel, pokalspiel, testspiel) in enumerate(yearly_players):
        target_row = first_data_row + idx
        ws.cell(row=target_row, column=1).value = name
        ws.cell(row=target_row, column=3).value = punktspiel
        ws.cell(row=target_row, column=4).value = pokalspiel
        ws.cell(row=target_row, column=5).value = testspiel

        prev_row = prev_name_rows.get(normalize_name(name))
        if prev_row is not None:
            ws.cell(row=target_row, column=2).value = f"='{PREV_SHEET}'!H{prev_row}"
        else:
            b_cell = ws.cell(row=target_row, column=2)
            b_cell.value = "=0"
            missing_prev_rows.append(target_row)

        ensure_row_formulas(ws, target_row, first_data_row)
        ws.cell(row=target_row, column=9).value = build_ehrung_formula(target_row)

    # Spalte I nur bei echtem Eintrag (nicht leer) rot/fett formatieren.
    ws.conditional_formatting.add(
        f"I{first_data_row}:I{last_row}",
        FormulaRule(formula=[f'I{first_data_row}<>""'], font=missing_font),
    )

    # Datenbereich ab Zeile 8 abwechselnd einfärben.
    fill_a = PatternFill(fill_type="solid", fgColor="FFFFFF")
    fill_b = PatternFill(fill_type="solid", fgColor="EDEDED")
    missing_font = Font(color="9C0006", bold=True)
    for row in range(first_data_row, last_row + 1):
        fill = fill_a if (row - first_data_row) % 2 == 0 else fill_b
        for col in range(1, ws.max_column + 1):
            ws.cell(row=row, column=col).fill = fill

    # Fehlende Vorjahreswerte in Spalte B sichtbar markieren.
    for row in missing_prev_rows:
        ws.cell(row=row, column=2).font = missing_font

    overall_wb.save(overall_path)
    print(f"Blatt '{PREV_SHEET}' kopiert nach '{NEW_SHEET}'")
    print(f"Namen aus yearly uebernommen: {len(yearly_players)}")
    print(f"Gespeichert: {overall_path}")


if __name__ == "__main__":
    main()
