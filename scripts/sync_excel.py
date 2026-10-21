import os
import json
import base64
from pathlib import Path

import requests
from openpyxl import load_workbook
from io import BytesIO

# OneDrive-Freigabelink aus den GitHub Secrets lesen
share_url = os.environ["ONEDRIVE_EXCEL_URL"].strip()

# Freigabelink für die OneDrive-API kodieren
encoded = base64.urlsafe_b64encode(
    share_url.encode("utf-8")
).decode("ascii").rstrip("=")

from urllib.parse import urlparse, parse_qs, urlencode, urlunparse

# OneDrive-Freigabelink für einen Downloadversuch vorbereiten
parts = urlparse(share_url)
params = parse_qs(parts.query)
params["download"] = ["1"]

download_url = urlunparse(
    parts._replace(query=urlencode(params, doseq=True))
)

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
    data_only=True
)

worksheet = workbook["Lieder"]

# Excel-Zeilen einschließlich Überschriften auslesen
rows = list(worksheet.iter_rows(values_only=True))

# Leere Zeilen am Ende entfernen
while rows and all(value is None for value in rows[-1]):
    rows.pop()

if not rows:
    raise ValueError("Das Tabellenblatt 'Lieder' ist leer.")

headers = [str(value or "").strip() for value in rows[0]]
column_map = {name: index for index, name in enumerate(headers)}

required = [
    "Titel",
    "Interpret",
    "Ausgabe 17",
    "Ausgabe 18",
    "Ausgabe 19",
]

missing = [name for name in required if name not in column_map]

if missing:
    raise ValueError(f"Fehlende Excel-Spalten: {missing}")

songs = []

for row in rows[1:]:
    def get_value(name):
        value = row[column_map[name]]
        return value

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

    # Leere Excel-Zellen als JSON-null speichern
    for key in ["ausgabe17", "ausgabe18", "ausgabe19"]:
        if song[key] == "":
            song[key] = None

    songs.append(song)

workbook.close()

# Ausschließlich Metadaten veröffentlichen – keine Liedtexte!
output = Path("lieder.json")
output.write_text(
    json.dumps(songs, ensure_ascii=False, indent=2) + "\n",
    encoding="utf-8"
)

print(f"{len(songs)} Lieder erfolgreich verarbeitet.")
