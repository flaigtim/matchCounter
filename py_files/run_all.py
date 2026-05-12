from __future__ import annotations

import subprocess
import sys
from pathlib import Path


SCRIPTS_IN_ORDER = [
    "counter.py",
    "merge_excels.py",
    "create_certificate.py",
    "convert_certificates_to_pdf.py",
]


# =========================================
# ZENTRALE KONFIGURATION (nur hier anpassen)
# =========================================
COUNTER_CONFIG: dict[str, object] = {
    "first_team_url": "https://www.fussball.de/mannschaft/sgm-mariazell-locherhof-stetten-lackendorf-sv-mariazell-wuerttemberg/-/saison/2526/team-id/02TCEJ3RA4000000VS5489BRVTHNGU03#!/",
    "second_team_url": "https://www.fussball.de/mannschaft/sgm-mariazell-locherhof-stetten-lackendorf-ii-sv-mariazell-wuerttemberg/-/saison/2526/team-id/02TCEK4EA0000000VS5489BRVTHNGU03#!/",
    "own_team_name_prefix": "SGM Mariazell/Locherhof/Stetten-Lackendorf",
    "date_from": "24.07.2025",
    "date_to": "11.05.2026",
    "headless": False,
    "wait_ms": 1200,
    "debug": False,
}

MERGE_CONFIG: dict[str, object] = {
    "yearly": "player_stats/stats_2025_2026.xlsx",
    "overall": "player_stats/_stats_overall.xlsx",
    "new_sheet": "25-26",
    "prev_sheet": "24-25",
}

CERT_CONFIG: dict[str, object] = {
    "date": "06.06.2026",
    "sheet": "25-26",
    "overall": "player_stats/_stats_overall.xlsx",
    "template": "template_certificate.docx",
    "cert_outdir": "player_stats/certificates/25-26",
}

PDF_CONFIG: dict[str, object] = {
    "pdf_input_dir": "player_stats/certificates/25-26",
}


def run_step(
    python_exe: str,
    script_path: Path,
    project_root: Path,
    extra_args: list[str] | None = None,
) -> int:
    cmd = [python_exe, str(script_path)]
    if extra_args:
        cmd.extend(extra_args)

    print(f"\n>> Starte: {script_path.name}")
    print(f"Befehl: {' '.join(cmd)}")

    result = subprocess.run(cmd, cwd=str(project_root))
    print(f"<< Ende: {script_path.name} (Exit-Code: {result.returncode})")
    return result.returncode


def main() -> None:
    py_dir = Path(__file__).resolve().parent
    project_root = py_dir.parent

    for script_name in SCRIPTS_IN_ORDER:
        script_path = py_dir / script_name
        if not script_path.exists():
            raise FileNotFoundError(f"Skript nicht gefunden: {script_path}")

        extra_args: list[str] = []

        if script_name == "counter.py":
            if isinstance(COUNTER_CONFIG.get("first_team_url"), str) and COUNTER_CONFIG["first_team_url"]:
                extra_args.extend(["--first-team-url", str(COUNTER_CONFIG["first_team_url"])])
            if COUNTER_CONFIG.get("second_team_url") is not None:
                extra_args.extend(["--second-team-url", str(COUNTER_CONFIG["second_team_url"])])
            if isinstance(COUNTER_CONFIG.get("own_team_name_prefix"), str) and COUNTER_CONFIG["own_team_name_prefix"]:
                extra_args.extend(["--own-team-name-prefix", str(COUNTER_CONFIG["own_team_name_prefix"])])
            if isinstance(COUNTER_CONFIG.get("date_from"), str) and COUNTER_CONFIG["date_from"]:
                extra_args.extend(["--date-from", str(COUNTER_CONFIG["date_from"])])
            if isinstance(COUNTER_CONFIG.get("date_to"), str) and COUNTER_CONFIG["date_to"]:
                extra_args.extend(["--date-to", str(COUNTER_CONFIG["date_to"])])
            if bool(COUNTER_CONFIG.get("headless")):
                extra_args.append("--headless")
            if COUNTER_CONFIG.get("wait_ms") is not None:
                extra_args.extend(["--wait-ms", str(COUNTER_CONFIG["wait_ms"])])
            if bool(COUNTER_CONFIG.get("debug")):
                extra_args.append("--debug")

        if script_name == "merge_excels.py":
            if isinstance(MERGE_CONFIG.get("yearly"), str) and MERGE_CONFIG["yearly"]:
                extra_args.extend(["--yearly", str(MERGE_CONFIG["yearly"])])
            if isinstance(MERGE_CONFIG.get("overall"), str) and MERGE_CONFIG["overall"]:
                extra_args.extend(["--overall", str(MERGE_CONFIG["overall"])])
            if isinstance(MERGE_CONFIG.get("new_sheet"), str) and MERGE_CONFIG["new_sheet"]:
                extra_args.extend(["--new-sheet", str(MERGE_CONFIG["new_sheet"])])
            if isinstance(MERGE_CONFIG.get("prev_sheet"), str) and MERGE_CONFIG["prev_sheet"]:
                extra_args.extend(["--prev-sheet", str(MERGE_CONFIG["prev_sheet"])])

        if script_name == "create_certificate.py":
            if isinstance(CERT_CONFIG.get("date"), str) and CERT_CONFIG["date"]:
                extra_args.extend(["--date", str(CERT_CONFIG["date"])])
            if isinstance(CERT_CONFIG.get("sheet"), str) and CERT_CONFIG["sheet"]:
                extra_args.extend(["--sheet", str(CERT_CONFIG["sheet"])])
            if isinstance(CERT_CONFIG.get("overall"), str) and CERT_CONFIG["overall"]:
                extra_args.extend(["--overall", str(CERT_CONFIG["overall"])])
            if isinstance(CERT_CONFIG.get("template"), str) and CERT_CONFIG["template"]:
                extra_args.extend(["--template", str(CERT_CONFIG["template"])])
            if isinstance(CERT_CONFIG.get("cert_outdir"), str) and CERT_CONFIG["cert_outdir"]:
                extra_args.extend(["--outdir", str(CERT_CONFIG["cert_outdir"])])

        if script_name == "convert_certificates_to_pdf.py":
            if isinstance(PDF_CONFIG.get("pdf_input_dir"), str) and PDF_CONFIG["pdf_input_dir"]:
                extra_args.extend(["--input-dir", str(PDF_CONFIG["pdf_input_dir"])])
            elif isinstance(CERT_CONFIG.get("cert_outdir"), str) and CERT_CONFIG["cert_outdir"]:
                extra_args.extend(["--input-dir", str(CERT_CONFIG["cert_outdir"])])

        exit_code = run_step(
            python_exe=sys.executable,
            script_path=script_path,
            project_root=project_root,
            extra_args=extra_args,
        )
        if exit_code != 0:
            print(f"\nAbbruch wegen Fehler in {script_name}.")
            sys.exit(exit_code)

    pdf_dir_raw = PDF_CONFIG.get("pdf_input_dir")
    if not isinstance(pdf_dir_raw, str) or not pdf_dir_raw:
        pdf_dir_raw = CERT_CONFIG.get("cert_outdir")

    if isinstance(pdf_dir_raw, str) and pdf_dir_raw:
        pdf_dir = project_root / pdf_dir_raw
        if pdf_dir.exists():
            pdf_count = len(list(pdf_dir.glob("*.pdf")))
            print(f"PDF-Dateien im Zielordner: {pdf_count}")
            print(f"PDF-Ordner: {pdf_dir}")
        else:
            print(f"Hinweis: PDF-Ordner nicht gefunden: {pdf_dir}")

    print("\nAlle Schritte erfolgreich abgeschlossen.")


if __name__ == "__main__":
    main()
