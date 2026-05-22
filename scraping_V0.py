import requests
import pandas as pd
from bs4 import BeautifulSoup

# 1. Iniciamos una sesión para mantener las cookies
session = requests.Session()

# Headers para simular un navegador real
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36",
    "Referer": "https://calidaddelaire.puebla.gob.mx/views/reporteICA.php"
}

# 2. Visitamos la página principal para que nos asigne una sesión
url_principal = "https://calidaddelaire.puebla.gob.mx/views/reporteICA.php"
session.get(url_principal, headers=headers)

# 3. Ahora pedimos los datos al archivo PHP
url_datos = "https://calidaddelaire.puebla.gob.mx/views/consultar_reporte.php"

# IMPORTANTE: Ajusta la fecha a la de HOY y una hora válida (ej. 08, 09, 10)
payload = {
    'fecha': '2026-01-08', # Formato YYYY-MM-DD
    'hora': '10:10'           # Formato HH
}

response = session.post(url_datos, data=payload, headers=headers)

# 4. Procesamos con BeautifulSoup
soup = BeautifulSoup(response.text, 'lxml')

# Buscamos todas las tablas
tablas = soup.find_all('table')

print(f"Tablas encontradas: {len(tablas)}")

# Si encontró tablas, mostramos la información de la tabla de interés
if len(tablas) > 0:
    for i, tabla in enumerate(tablas):
        # Convertimos la tabla de HTML a un DataFrame de Pandas directamente
        df = pd.read_html(str(tabla))[0]
        print(f"\n--- TABLA {i} ---")
        print(df.head())
else:
    print("No se encontraron tablas. Verifica que la fecha y hora tengan datos en el portal.")