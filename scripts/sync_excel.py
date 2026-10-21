import os
import json
from io import BytesIO
from pathlib import Path
from urllib.parse import urlparse, parse_qs, urlencode, urlunparse

import requests
from openpyxl import load_workbook


def main():
    # OneDrive-Freigabelink aus den GitHub Secrets lesen
    share_url = os.environ.get("ONEDRIVE_EXCEL_URL", "").strip()

    if not share_url:
        raise ValueError("Das GitHub Secret ONEDRIVE_EXCEL_URL fehlt.")

# OneDrive-Link für einen Downloadversuch vorbereiten
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

# Prüfen, ob tatsächlich eine Excel-Datei geladen wurde
if not response.content.startswith(b"PK"):
    raise ValueError(
        "OneDrive hat keine Excel-Datei geliefert. "
        "Der Freigabelink muss anders aufgerufen werden."
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
            "Das Tabellenblatt 'Lieder' wurde nicht gefunden."
        )

    worksheet = workbook["Lieder"]

    # Excel-Zeilen einschließlich Überschriften auslesen
    rows = list(worksheet.iter_rows(values_only=True))

    # Leere Zeilen am Ende entfernen
    while rows and all(value is None for value in rows[-1]):
        rows.pop()

    if not rows:
        raise ValueError("Das Tabellenblatt 'Lieder' ist leer.")

    # Spaltenüberschriften aus der ersten Zeile lesen
    headers = [
        str(value).strip() if value is not None else ""
        for value in rows[0]
    ]

    # Doppelte oder leere Überschriften werden nicht benötigt
    column_map = {
        name: index
        for index, name in enumerate(headers)
        if name
    }

    required = [
        "Titel",
        "Interpret",
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
            f"Fehlende Excel-Spalten: {missing}. "
            f"Gefundene Spalten: {headers}"
        )

    # Ausschließlich Liedmetadaten übernehmen.
    # Die Spalte 'Liedtext' wird nicht veröffentlicht.
    songs = []

    for row in rows[1:]:
        def get_value(name):
            index = column_map[name]
            return row[index] if index < len(row) else None

        title = get_value("Titel")

        if title is None or str(title).strip() == "":
            continue

            song = {
                "id": str(len(songs) + 1),
                "titel": str(title).strip(),
                "interpret": (
                    str(get_value("Interpret")).strip()
                    if get_value("Interpret") is not None
                    else ""
                ),
                "ausgabe17": get_value("Ausgabe 17"),
                "ausgabe18": get_value("Ausgabe 18"),
                "ausgabe19": get_value("Ausgabe 19"),
            }
