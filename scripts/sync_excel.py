import json
import os
import tempfile
import urllib.request
from pathlib import Path

from openpyxl import load_workbook

EXCEL_URL = os.environ["ONEDRIVE_EXCEL_URL"]
WORKSHEET_NAME = "Lieder"
OUTPUT_FILE = Path("lieder.json")


def main():
    with tempfile.TemporaryDirectory() as temp_dir:
        excel_path = Path(temp_dir) / "lieder.xlsx"

        # OneDrive-Freigabelink in eine API-Download-Adresse umwandeln
        import base64

        share_token = (
            base64.urlsafe_b64encode(EXCEL_URL.encode("utf-8"))
            .decode("ascii")
            .rstrip("=")
        )
        share_id = "u!" + share_token
        download_url = EXCEL_URL

        request = urllib.request.Request(
            download_url,
            headers={
                "User-Agent": "Mozilla/5.0",
                "Accept": "*/*",
            },
        )

        with urllib.request.urlopen(request, timeout=60) as response:
            file_data = response.read()

        if not file_data.startswith(b"PK"):
            raise ValueError(
                "OneDrive hat keine gültige Excel-Datei geliefert."
            )

        excel_path.write_bytes(file_data)

        workbook = load_workbook(
            excel_path,
            read_only=True,
            data_only=True,
        )

        if WORKSHEET_NAME not in workbook.sheetnames:
            raise ValueError(
                f"Arbeitsblatt '{WORKSHEET_NAME}' nicht gefunden."
            )

        sheet = workbook[WORKSHEET_NAME]
        rows = sheet.iter_rows(values_only=True)
        headers = next(rows, None)

        if not headers:
            raise ValueError("Excel-Datei enthält keine Kopfzeile.")

        headers = [
            str(value).strip() if value is not None else ""
            for value in headers
        ]

        required = [
            "Titel",
            "Interpret",
            "Liedtext",
            "Ausgabe 17",
            "Ausgabe 18",
            "Ausgabe 19",
        ]

        missing = [name for name in required if name not in headers]
        if missing:
            raise ValueError(
                "Fehlende Excel-Spalten: " + ", ".join(missing)
            )

        column = {name: headers.index(name) for name in required}
        songs = []

        for row in rows:
            def value(name):
                index = column[name]
                return row[index] if index < len(row) else None

            title = value("Titel")
            if title is None or not str(title).strip():
                continue

            def edition(name):
                number = value(name)
                if isinstance(number, (int, float)) and number > 0:
                    return int(number)
                return None

            songs.append({
                "id": str(len(songs) + 1),
                "titel": str(title).strip(),
                "interpret": str(value("Interpret") or "").strip(),
                "liedtext": str(value("Liedtext") or ""),
                "ausgabe17": edition("Ausgabe 17"),
                "ausgabe18": edition("Ausgabe 18"),
                "ausgabe19": edition("Ausgabe 19"),
            })

        workbook.close()

        OUTPUT_FILE.write_text(
            json.dumps(songs, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )

        print(f"{len(songs)} Lieder exportiert.")

if __name__ == "__main__":
    main()
