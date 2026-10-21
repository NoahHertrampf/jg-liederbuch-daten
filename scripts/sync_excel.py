import os
import json
from io import BytesIO
from pathlib import Path
from urllib.parse import urlparse, parse_qs, urlencode, urlunparse

import requests
from openpyxl import load_workbook


def main():
    share_url = os.environ.get("ONEDRIVE_EXCEL_URL", "").strip()

    if not share_url:
        raise ValueError("Das GitHub Secret ONEDRIVE_EXCEL_URL fehlt.")

    # OneDrive-Downloadlink vorbereiten
    parts = urlparse(share_url)
    params = parse_qs(parts.query)
    params["download"] = ["1"]

    download_url = urlunparse(
        parts._replace(query=urlencode(params, doseq=True))
    )

    # Excel-Datei herunterladen
    response = requests.get(
        download_url,
        timeout=60,
        allow_redirects=True,
    )
    response.raise_for_status()

    if not response.content.startswith(b"PK"):
        raise ValueError(
            "OneDrive hat keine Excel-Datei geliefert. "
            "Bitte den Freigabelink prüfen."
        )

    # Excel-Datei öffnen
    workbook = load_workbook(
        BytesIO(response.content),
        read_only=True,
        data_only=True,
    )

    try:
        if "Lieder" not in workbook.sheetnames:
            raise ValueError(
                f"Tabellenblatt 'Lieder' nicht gefunden. "
                f"Vorhanden: {workbook.sheetnames}"
            )

        worksheet = workbook["Lieder"]
        rows = list(worksheet.iter_rows(values_only=True))

        while rows and all(value is None for value in rows[-1]):
            rows.pop()

        if not rows:
            raise ValueError("Das Tabellenblatt 'Lieder' ist leer.")

        headers = [
            str(value).strip() if value is not None else ""
            for value in rows[0]
        ]

        column_map = {
            name: index
            for index, name in enumerate(headers)
            if name
        }

        required = [
            "Titel",
            "Interpret",
            "Liedtext",
            "Ausgabe 17",
            "Ausgabe 18",
            "Ausgabe 19",
        ]

        missing = [
            name for name in required
            if name not in column_map
        ]

        if missing:
            raise ValueError(
                f"Fehlende Spalten: {missing}. "
                f"Gefundene Spalten: {headers}"
            )

        songs = []

        for row in rows[1:]:
            def get_value(name):
                index = column_map[name]
                return row[index] if index < len(row) else None

            title = get_value("Titel")

            if title is None or str(title).strip() == "":
                continue

            def edition_value(name):
                value = get_value(name)

                if value is None or str(value).strip() == "":
                    return None

                if isinstance(value, (int, float)):
                    if value <= 0:
                        return None
                    return int(value)

                return value

            song = {
                "id": str(len(songs) + 1),
                "titel": str(title).strip(),
                "interpret": (
                    str(get_value("Interpret")).strip()
                    if get_value("Interpret") is not None
                    else ""
                ),
                "liedtext": (
                    str(get_value("Liedtext"))
                    if get_value("Liedtext") is not None
                    else ""
                ),
                "ausgabe17": edition_value("Ausgabe 17"),
                "ausgabe18": edition_value("Ausgabe 18"),
                "ausgabe19": edition_value("Ausgabe 19"),
            }

            songs.append(song)

    finally:
        workbook.close()

    # JSON-Datei inklusive Liedtexte schreiben
    output = Path("lieder.json")
    output.write_text(
        json.dumps(songs, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )

    print(f"{len(songs)} Lieder erfolgreich verarbeitet.")
    print("Liedtexte wurden in lieder.json übernommen.")


if __name__ == "__main__":
    main()
