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