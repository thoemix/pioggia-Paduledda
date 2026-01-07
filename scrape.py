import requests
from bs4 import BeautifulSoup
from datetime import date
import pandas as pd
import os

URL = "https://www.sardegnacedoc.it/idrografico/sensore/425100/32967/20207"

headers = {
    "User-Agent": "Mozilla/5.0"
}

response = requests.get(URL, headers=headers, timeout=30)
response.raise_for_status()

soup = BeautifulSoup(response.text, "html.parser")

value = None

# Alle Zeilen der Tabelle durchsuchen
for row in soup.find_all("tr"):
    cells = row.find_all("td")
    if len(cells) >= 2:  # mindestens 2 Spalten
        label = cells[0].get_text(strip=True)
        if "Pioggia cumulata ieri" in label:
            value = cells[1].get_text(strip=True)
            break

if value is None:
    raise Exception("Pioggia-Wert nicht gefunden")

# Datum heute
today = date.today().isoformat()
file = "pioggia.xlsx"

# Excel laden oder neu erstellen
if os.path.exists(file):
    df = pd.read_excel(file)
else:
    df = pd.DataFrame(columns=["Datum", "Pioggia cumulata ieri"])

# Neuen Wert eintragen
df.loc[len(df)] = [today, value]

# Excel speichern
df.to_excel(file, index=False)

print(f"Heutiger Pioggia-Wert: {value}")
