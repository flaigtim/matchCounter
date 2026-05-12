from __future__ import annotations

import argparse
from pathlib import Path

try:
    from docx2pdf import convert as docx2pdf_convert
except Exception:  # pragma: no cover
    docx2pdf_convert = None

DEFAULT_INPUT_DIR = "player_stats/certificates/25-26"


def convert_with_docx2pdf(docx_file: Path, pdf_file: Path) -> bool:
    if docx2pdf_convert is None:
        return False

    try:
        docx2pdf_convert(str(docx_file.resolve()), str(pdf_file.resolve()))
        return True
    except Exception:
        return False


def convert_one_file(docx_file: Path) -> bool:
    pdf_file = docx_file.with_suffix(".pdf")

    if convert_with_docx2pdf(docx_file, pdf_file):
        if pdf_file.exists():
            docx_file.unlink()
            return True
    return False


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Konvertiert DOCX-Urkunden manuell zu PDF und loescht "
            "die DOCX-Dateien nach erfolgreicher Konvertierung."
        )
    )
    parser.add_argument(
        "--input-dir",
        default=DEFAULT_INPUT_DIR,
        help=f"Ordner mit DOCX-Dateien (Standard: {DEFAULT_INPUT_DIR})",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    input_dir = Path(args.input_dir)

    if not input_dir.exists():
        raise FileNotFoundError(f"Ordner nicht gefunden: {input_dir}")

    docx_files = sorted(
        f for f in input_dir.glob("*.docx") if not f.name.startswith("~$")
    )
    if not docx_files:
        print("Keine DOCX-Dateien gefunden.")
        return

    converted = 0
    failed: list[Path] = []

    for docx_file in docx_files:
        ok = convert_one_file(docx_file)
        if ok:
            converted += 1
        else:
            failed.append(docx_file)

    print(f"DOCX gefunden: {len(docx_files)}")
    print(f"PDF erstellt + DOCX geloescht: {converted}")

    if failed:
        print("Fehlgeschlagen:")
        for file in failed:
            print(f"- {file}")
        if docx2pdf_convert is None:
            print("Hinweis: docx2pdf fehlt. Installation: pip install docx2pdf")
        else:
            print("Hinweis: Stelle sicher, dass Microsoft Word installiert und nicht durch Dialoge blockiert ist.")


if __name__ == "__main__":
    main()
