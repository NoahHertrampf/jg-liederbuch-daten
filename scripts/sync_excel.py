import json
import os
import tempfile
import urllib.request
from pathlib import Path

from openpyxl import load_workbook

EXCEL_URL = os.environ["ONEDRIVE_EXCEL_URL"]

WORKSHEET_NAME = "Lieder"
OUTPUT_FILE = Path("lieder.json")


def download_excel(url, destination):
    request = urllib.request.Request(
        url,
        headers={"User-Agent": "JG-Liederbuch-Sync/1.0"},
    )

    with urllib.request.urlopen(request, timeout=60) as response:
        destination.write_bytes(response.read())


def main():
    with tempfile.TemporaryDirectory() as temp_dir:
        excel_path = Path(temp_dir) / "lieder.xlsx"
        download_excel(EXCEL_URL, excel_path)

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
            raise ValueError("Die Excel-Datei enthält keine Kopfzeile.")

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
            raise ValueEr
