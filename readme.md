# Facebook Marketplace Extractor V2

Extractor de publicaciones propias de Facebook Marketplace.

## Características

- Conexión mediante Chrome CDP.
- No almacena credenciales.
- Extracción de publicaciones propias.
- Extracción de precio.
- Extracción de disponibilidad.
- Extracción de SKU.
- Extracción de listing_id.
- Extracción de antigüedad de publicación.
- Exportación a Excel.
- Conservación de registros anteriores.
- Recuperación mediante checkpoint.

## Requisitos

- Python 3.x
- Google Chrome
- Playwright
- OpenPyXL

## Instalación

```powershell
py -m pip install -r requirements.txt
py -m playwright install chromium



## COMANDOS DE INICIO RAPIDOS -- Una ves instalado todo y siendo rutinario este proceso

  taskkill /F /IM chrome.exe

   & "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\Users\compumarker\Desktop\SCRIPS\EXTRACTOR FB\chrome_debug"

 py Facebook_Marketplace_V2.py