from __future__ import annotations

import argparse
import re
from pathlib import Path

import openpyxl


def normalize_text(value: str) -> str:
    text = (value or "").strip().lower()
    text = text.replace("ä", "ae").replace("ö", "oe").replace("ü", "ue").replace("ß", "ss")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def normalize_name(value: str) -> str:
    text = (value or "").strip()
    if not text:
        return ""

    if "," in text:
        last, first = text.split(",", 1)
        first_tokens = normalize_text(first).split()
        last_tokens = normalize_text(last).split()
    else:
        tokens = normalize_text(text).split()
        if not tokens:
            return ""
        if len(tokens) == 1:
            return tokens[0]
        first_tokens = tokens[:-1]
        last_tokens = [tokens[-1]]

    return " ".join(last_tokens + first_tokens)


def sort_key_from_name(value: str) -> tuple[str, str]:
    text = (value or "").strip()
    if not text:
        return ("", "")

    if "," in text:
        last, first = text.split(",", 1)
        return (normalize_text(last), normalize_text(first))

    tokens = normalize_text(text).split()
    if len(tokens) <= 1:
        return (tokens[0] if tokens else "", "")

    return (tokens[-1], " ".join(tokens[:-1]))


def to_int(value) -> int:
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


def header_index_by_aliases(sheet, aliases: list[str]) -> int | None:
    alias_set = {normalize_text(a) for a in aliases}
    for col in range(1, sheet.max_column + 1):
        value = sheet.cell(row=1, column=col).value
        if value is None:
            continue
        if normalize_text(str(value)) in alias_set:
            return col
    return None


def ensure_column(sheet, aliases: list[str], fallback_header: str) -> int:
    existing = header_index_by_aliases(sheet, aliases)
    if existing:
        return existing

    new_col = sheet.max_column + 1
    sheet.cell(row=1, column=new_col, value=fallback_header)
    return new_col


def read_yearly_values(yearly_sheet) -> dict[str, dict[str, int | str]]:
    name_col = header_index_by_aliases(yearly_sheet, ["Spieler", "Name", "Spielername"])
    punkt_col = header_index_by_aliases(yearly_sheet, ["Punktspiel", "Punktspiele"])
    pokal_col = header_index_by_aliases(yearly_sheet, ["Pokalspiel", "Pokalspiele"])
    test_col = header_index_by_aliases(
        yearly_sheet,
        ["Freundschaftsspiel", "Freundschaftsspiele", "Testspiel", "Testspiele"],
    )

    missing = []
    if not name_col:
        missing.append("Name/Spieler")
    if not punkt_col:
        missing.append("Punktspiel(e)")
    if not pokal_col:
        missing.append("Pokalspiel(e)")
    if not test_col:
        missing.append("Freundschafts-/Testspiel(e)")

    if missing:
        raise ValueError(f"Jahresdatei: benötigte Spalten fehlen: {', '.join(missing)}")

    data: dict[str, dict[str, int | str]] = {}
    for row in range(2, yearly_sheet.max_row + 1):
        raw_name = yearly_sheet.cell(row=row, column=name_col).value
        if raw_name is None:
            continue

        display_name = str(raw_name).strip()
        if not display_name:
            continue

        key = normalize_name(display_name)
        if not key:
            continue

        data[key] = {
            "name": display_name,
            "punkt": to_int(yearly_sheet.cell(row=row, column=punkt_col).value),
            "pokal": to_int(yearly_sheet.cell(row=row, column=pokal_col).value),
            "test": to_int(yearly_sheet.cell(row=row, column=test_col).value),
        }

    return data


def merge_into_overall(overall_sheet, yearly_data: dict[str, dict[str, int | str]]) -> tuple[int, int]:
    name_col = ensure_column(overall_sheet, ["Spieler", "Name", "Spielername"], "Spieler")
    punkt_col = ensure_column(overall_sheet, ["Punktspiel", "Punktspiele"], "Punktspiele")
    pokal_col = ensure_column(overall_sheet, ["Pokalspiel", "Pokalspiele"], "Pokalspiele")
    test_col = ensure_column(
        overall_sheet,
        ["Testspiel", "Testspiele", "Freundschaftsspiel", "Freundschaftsspiele"],
        "Testspiele",
    )

    existing_by_key: dict[str, int] = {}
    for row in range(2, overall_sheet.max_row + 1):
        raw_name = overall_sheet.cell(row=row, column=name_col).value
        if raw_name is None:
            continue
        key = normalize_name(str(raw_name))
        if key:
            existing_by_key[key] = row

    updated = 0
    appended = 0

    for key, values in yearly_data.items():
        row = existing_by_key.get(key)

        if row is None:
            row = overall_sheet.max_row + 1
            overall_sheet.cell(row=row, column=name_col, value=values["name"])
            existing_by_key[key] = row
            appended += 1
        else:
            updated += 1

        overall_sheet.cell(row=row, column=punkt_col, value=values["punkt"])
        overall_sheet.cell(row=row, column=pokal_col, value=values["pokal"])
        overall_sheet.cell(row=row, column=test_col, value=values["test"])

    return updated, appended


def sort_overall_by_last_name(overall_sheet) -> None:
    name_col = header_index_by_aliases(overall_sheet, ["Spieler", "Name", "Spielername"])
    if not name_col:
        return

    max_col = overall_sheet.max_column
    max_row = overall_sheet.max_row

    rows = []
    for row_idx in range(2, max_row + 1):
        name_value = overall_sheet.cell(row=row_idx, column=name_col).value
        if name_value is None or not str(name_value).strip():
            continue
        row_values = [overall_sheet.cell(row=row_idx, column=col).value for col in range(1, max_col + 1)]
        rows.append(row_values)

    rows.sort(key=lambda values: sort_key_from_name(str(values[name_col - 1] or "")))

    for row_idx in range(2, max_row + 1):
        for col in range(1, max_col + 1):
            overall_sheet.cell(row=row_idx, column=col, value=None)

    for offset, values in enumerate(rows, start=2):
        for col, value in enumerate(values, start=1):
            overall_sheet.cell(row=offset, column=col, value=value)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Liest eine yearly-Excel und eine overall-Excel ein, "
            "übernimmt Punkt-/Pokal-/Testspiele ins angegebene Overall-Blatt, "
            "ergänzt fehlende Namen und sortiert nach Nachname."
        )
    )
    parser.add_argument("--yearly-file", required=True, help="Pfad zur yearly Excel-Datei")
    parser.add_argument("--overall-file", required=True, help="Pfad zur overall Excel-Datei")
    parser.add_argument(
        "--overall-sheet",
        required=True,
        help="Tabellenblatt in der overall-Datei, das abgeglichen werden soll",
    )
    parser.add_argument(
        "--yearly-sheet",
        default="Zusammenfassung",
        help="Tabellenblatt in der yearly-Datei (Standard: Zusammenfassung)",
    )
    parser.add_argument(
        "--output-file",
        default="",
        help="Optionaler Ausgabepfad. Leer = overall-Datei wird überschrieben.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    yearly_path = Path(args.yearly_file)
    overall_path = Path(args.overall_file)

    if not yearly_path.exists():
        raise FileNotFoundError(f"Yearly-Datei nicht gefunden: {yearly_path}")
    if not overall_path.exists():
        raise FileNotFoundError(f"Overall-Datei nicht gefunden: {overall_path}")

    yearly_wb = openpyxl.load_workbook(yearly_path)
    overall_wb = openpyxl.load_workbook(overall_path)

    if args.yearly_sheet not in yearly_wb.sheetnames:
        raise ValueError(
            f"Yearly-Sheet '{args.yearly_sheet}' nicht gefunden. Vorhanden: {', '.join(yearly_wb.sheetnames)}"
        )
    if args.overall_sheet not in overall_wb.sheetnames:
        raise ValueError(
            f"Overall-Sheet '{args.overall_sheet}' nicht gefunden. Vorhanden: {', '.join(overall_wb.sheetnames)}"
        )

    yearly_sheet = yearly_wb[args.yearly_sheet]
    overall_sheet = overall_wb[args.overall_sheet]

    yearly_data = read_yearly_values(yearly_sheet)
    updated, appended = merge_into_overall(overall_sheet, yearly_data)
    sort_overall_by_last_name(overall_sheet)

    output_path = Path(args.output_file) if args.output_file.strip() else overall_path
    overall_wb.save(output_path)

    print(f"Yearly-Einträge gelesen: {len(yearly_data)}")
    print(f"Aktualisierte Namen in overall: {updated}")
    print(f"Neu hinzugefügte Namen in overall: {appended}")
    print(f"Gespeichert: {output_path}")


if __name__ == "__main__":
    main()
