import requests
from bs4 import BeautifulSoup
from datetime import date
import pandas as pd
import os

URL = "https://www.sardegnacedoc.it/idrografico/sensore/425100/32967/20207"

response = requests.get(URL, timeout=30)
soup = BeautifulSoup(response.text, "html.parser")

# Text suchen
text = soup.get_text(separator=" ", strip=True)

value = None
for part in text.split():
    if part.replace(",", ".").replace(".", "", 1).isdigit():
        value = part
        break

today = date.today().isoformat()

file = "pioggia.xlsx"

if os.path.exists(file):
    df = pd.read_excel(file)
else:
    df = pd.DataFrame(columns=["Datum", "Pioggia cumulata ieri"])

df.loc[len(df)] = [today, value]
df.to_excel(file, index=False)
