# -*- coding: utf-8 -*-
import csv
import json
import re
import time
from pathlib import Path

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

CDP_URL = "http://127.0.0.1:9222"
MARKETPLACE_URL = "https://www.facebook.com/marketplace/you/selling"

BASE = Path(__file__).resolve().parent
CSV_SALIDA = BASE / "Facebook_Marketplace.csv"
JSON_SALIDA = BASE / "Facebook_Marketplace.json"

PAUSA_ENTRE_PRODUCTOS = 0.35
PAUSA_SCROLL = 0.45
MAX_SCROLLS = 80
TIMEOUT_MODAL = 5000

CAMPOS = [
    "nombre",
    "precio",
    "disponibilidad",
    "sku",
    "publicado",
    "listing_id",
]


def limpiar_texto(texto):
    if not texto:
        return ""
    texto = texto.replace("\xa0", " ")
    return "\n".join(
        re.sub(r"\s+", " ", x).strip()
        for x in texto.replace("\r", "\n").split("\n")
        if x.strip()
    )


def precios(texto):
    return [
        x.replace(",", "")
        for x in re.findall(r"S/\s*([\d.,]+)", texto or "", re.I)
    ]


def precio(texto):
    p = precios(texto)
    return p[0] if p else ""


def disponibilidad(texto):
    if re.search(r"\bDisponible\b", texto or "", re.I):
        return "Disponible"
    if re.search(r"\bAgotado\b", texto or "", re.I):
        return "Agotado"
    return ""


def publicado(texto):
    for linea in (texto or "").splitlines():
        linea = linea.strip()
        if re.search(r"Publicado\s+(el|hace|en)", linea, re.I):
            return linea
    return ""


def sku_desde_modal(texto):
    lineas = [x.strip() for x in (texto or "").splitlines() if x.strip()]

    for i, linea in enumerate(lineas):
        if re.fullmatch(r"Publicación\s*:?", linea, re.I) and i > 0:
            anterior = lineas[i - 1]
            if re.fullmatch(r"\d+", anterior):
                return anterior

    m = re.search(r"(\d+)\s*Publicación\s*:", texto or "", re.I)
    return m.group(1) if m else ""


def nombre_desde_texto(texto):
    lineas = [x.strip() for x in (texto or "").splitlines() if x.strip()]

    for linea in lineas:
        if re.fullmatch(r"S/\s*[\d.,]+(?:S/\s*[\d.,]+)*", linea, re.I):
            continue
        if re.fullmatch(r"\d+", linea):
            continue
        if re.search(r"^(Disponible|Agotado)\b", linea, re.I):
            continue
        if re.search(r"Publicado\s+(el|hace|en)", linea, re.I):
            continue
        if re.search(r"Publicado en Marketplace", linea, re.I):
            continue
        if re.search(r"clics?\s+en\s+la publicación", linea, re.I):
            continue
        if len(linea) >= 3:
            return linea

    return ""


def normalizar(nombre):
    return re.sub(r"\s+", " ", nombre or "").strip().lower()


def cargar_registros():
    if not CSV_SALIDA.exists():
        return []

    registros = []
    try:
        with open(CSV_SALIDA, "r", encoding="utf-8-sig", newline="") as f:
            for fila in csv.DictReader(f):
                registros.append({c: fila.get(c, "") for c in CAMPOS})
    except Exception as e:
        print("Aviso leyendo CSV anterior:", e)

    return registros


def guardar(registros):
    # Un registro por SKU cuando existe.
    por_sku = {}
    sin_sku = []

    for r in registros:
        s = str(r.get("sku", "")).strip()
        if s:
            por_sku[s] = r
        else:
            sin_sku.append(r)

    finales = list(por_sku.values()) + sin_sku

    with open(CSV_SALIDA, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=CAMPOS)
        w.writeheader()
        w.writerows(finales)

    with open(JSON_SALIDA, "w", encoding="utf-8") as f:
        json.dump(finales, f, ensure_ascii=False, indent=2)


def detectar_tarjetas(page):
    """
    No usa clases x... de Facebook.
    Busca img[alt] y sube por sus padres hasta hallar
    un bloque pequeño que contenga precio + estado.
    """
    resultado = []
    firmas = set()

    try:
        imagenes = page.locator("img[alt]")
        cantidad = imagenes.count()
        print("Imágenes encontradas:", cantidad)

        for i in range(cantidad):
            try:
                img = imagenes.nth(i)
                alt = (img.get_attribute("alt") or "").strip()

                if not alt:
                    continue

                if any(
                    x in alt.lower()
                    for x in ("facebook", "marketplace", "avatar", "perfil", "logo", "icon")
                ):
                    continue

                candidato = None
                texto = ""

                for nivel in range(2, 10):
                    try:
                        padre = img.locator("xpath=" + "/.." * nivel)
                        if padre.count() == 0:
                            continue

                        t = limpiar_texto(padre.inner_text(timeout=800))

                        if not t or len(t) > 1600:
                            continue

                        if not re.search(r"S/\s*[\d.,]+", t, re.I):
                            continue

                        if not re.search(r"Disponible|Agotado|Publicado", t, re.I):
                            continue

                        n = nombre_desde_texto(t) or alt
                        if not n:
                            continue

                        candidato = padre
                        texto = t
                        break
                    except Exception:
                        pass

                if candidato is None:
                    continue

                nombre = nombre_desde_texto(texto) or alt
                p = precio(texto)
                firma = (normalizar(nombre), p)

                if firma in firmas:
                    continue

                firmas.add(firma)
                resultado.append({
                    "locator": candidato,
                    "nombre": nombre,
                    "precio": p,
                })

            except Exception:
                pass

    except Exception as e:
        print("Error detectando tarjetas:", e)

    return resultado


def cargar_todo(page):
    print("\nCargando publicaciones...\n")

    ultimo = -1
    sin_cambios = 0

    for vuelta in range(1, MAX_SCROLLS + 1):
        try:
            actual = page.locator("img[alt]").count()
        except Exception:
            actual = 0

        print(f"Scroll {vuelta:02d} | imágenes DOM: {actual}")

        if actual == ultimo:
            sin_cambios += 1
        else:
            sin_cambios = 0
            ultimo = actual

        if sin_cambios >= 5:
            break

        try:
            page.mouse.wheel(0, 1300)
        except Exception:
            break

        time.sleep(PAUSA_SCROLL)

    try:
        page.evaluate("window.scrollTo(0,0)")
    except Exception:
        pass

    time.sleep(0.5)


def obtener_modal(page):
    selectores = [
        '[role="dialog"][aria-label="Tu publicación"]',
        '[aria-label="Tu publicación"][role="dialog"]',
    ]

    for selector in selectores:
        try:
            loc = page.locator(selector)
            if loc.count():
                modal = loc.last
                if modal.is_visible(timeout=500):
                    return modal
        except Exception:
            pass

    try:
        loc = page.locator('[role="dialog"]')
        if loc.count():
            return loc.last
    except Exception:
        pass

    return None


def cerrar_modal(page):
    try:
        page.keyboard.press("Escape")
        time.sleep(0.2)
    except Exception:
        pass


def listing_id_modal(modal):
    try:
        links = modal.locator('[href*="listing_id="]')
        for i in range(links.count()):
            href = links.nth(i).get_attribute("href") or ""
            m = re.search(r"listing_id[=%3D]+(\d+)", href, re.I)
            if m:
                return m.group(1)
    except Exception:
        pass

    try:
        html = modal.inner_html()
        m = re.search(r"listing_id.{0,100}?(\d{8,})", html, re.I)
        if m:
            return m.group(1)
    except Exception:
        pass

    return ""


def extraer_publicacion(page, info):
    nombre_tarjeta = info["nombre"]
    precio_tarjeta = info["precio"]

    print("\n--------------------------------------------")
    print("Abriendo:", nombre_tarjeta)

    try:
        info["locator"].scroll_into_view_if_needed(timeout=3000)
        time.sleep(0.1)
        info["locator"].click(timeout=3000)
    except Exception:
        try:
            info["locator"].locator("img[alt]").first.click(timeout=3000)
        except Exception as e:
            print("No se pudo abrir:", e)
            return {
                "nombre": nombre_tarjeta,
                "precio": precio_tarjeta,
                "disponibilidad": "",
                "sku": "",
                "publicado": "",
                "listing_id": "",
            }

    try:
        modal = page.locator(
            '[role="dialog"][aria-label="Tu publicación"]'
        ).last
        modal.wait_for(state="visible", timeout=TIMEOUT_MODAL)
    except Exception:
        modal = obtener_modal(page)

    if not modal:
        print("No apareció el modal.")
        cerrar_modal(page)
        return {
            "nombre": nombre_tarjeta,
            "precio": precio_tarjeta,
            "disponibilidad": "",
            "sku": "",
            "publicado": "",
            "listing_id": "",
        }

    try:
        texto = limpiar_texto(modal.inner_text(timeout=2000))
    except Exception:
        texto = ""

    s = sku_desde_modal(texto)
    ps = precios(texto)

    r = {
        "nombre": nombre_desde_texto(texto) or nombre_tarjeta,
        "precio": ps[0] if ps else precio_tarjeta,
        "disponibilidad": disponibilidad(texto),
        "sku": s,
        "publicado": publicado(texto),
        "listing_id": listing_id_modal(modal),
    }

    print(json.dumps(r, ensure_ascii=False, indent=2))

    cerrar_modal(page)
    time.sleep(PAUSA_ENTRE_PRODUCTOS)

    return r


def main():
    print("\n============================================")
    print(" EXTRACTOR FACEBOOK MARKETPLACE")
    print("============================================\n")

    registros = cargar_registros()
    print("Registros existentes:", len(registros))

    nombres_existentes = {
        normalizar(x.get("nombre", ""))
        for x in registros
        if x.get("nombre")
    }

    with sync_playwright() as p:
        print("Conectando al Chrome existente...")

        try:
            browser = p.chromium.connect_over_cdp(CDP_URL)
        except Exception as e:
            print("\nNO SE PUDO CONECTAR AL CHROME.")
            print("Inicia Chrome con --remote-debugging-port=9222")
            print(e)
            input("\nENTER para salir...")
            return

        print("Chrome conectado correctamente.")

        page = None

        for context in browser.contexts:
            for pestaña in context.pages:
                try:
                    if "facebook.com" in pestaña.url.lower():
                        page = pestaña
                        if "marketplace/you/selling" in pestaña.url.lower():
                            break
                except Exception:
                    pass
            if page and "marketplace/you/selling" in page.url.lower():
                break

        if page is None:
            for context in browser.contexts:
                if context.pages:
                    page = context.pages[0]
                    break

        if page is None:
            print("No hay pestañas disponibles.")
            input("\nENTER para salir...")
            return

        print("Página seleccionada:")
        print(page.url)

        if "marketplace/you/selling" not in page.url.lower():
            print("Abriendo Marketplace...")
            try:
                page.goto(
                    MARKETPLACE_URL,
                    wait_until="domcontentloaded",
                    timeout=30000,
                )
            except Exception as e:
                print("Aviso:", e)

        print("Esperando Marketplace...")
        page.wait_for_timeout(1200)

        print("URL actual:")
        print(page.url)

        cargar_todo(page)

        tarjetas = detectar_tarjetas(page)

        print("\n============================================")
        print("TARJETAS DETECTADAS:", len(tarjetas))
        print("============================================\n")

        if not tarjetas:
            print("No se detectaron tarjetas.")
            print("La página está abierta, pero el DOM no coincide.")
            input("\nENTER para salir...")
            return

        procesadas = 0
        saltadas = 0

        for i, info in enumerate(tarjetas, 1):
            nombre = info["nombre"]

            print(f"\n[{i}/{len(tarjetas)}]")

            if normalizar(nombre) in nombres_existentes:
                print("Ya existe. Se omite para ahorrar recursos.")
                saltadas += 1
                continue

            registro = extraer_publicacion(page, info)
            registros.append(registro)

            nombres_existentes.add(
                normalizar(registro.get("nombre", ""))
            )

            # Guardado incremental: si se corta el script,
            # lo ya extraído queda en el CSV.
            guardar(registros)

            procesadas += 1

        guardar(registros)

        print("\n============================================")
        print(" EXTRACCIÓN TERMINADA")
        print("============================================")
        print("Tarjetas:", len(tarjetas))
        print("Procesadas:", procesadas)
        print("Saltadas:", saltadas)
        print("Total:", len(registros))
        print("\nCSV:")
        print(CSV_SALIDA)
        print("\nJSON:")
        print(JSON_SALIDA)

        input("\nENTER para salir...")


if __name__ == "__main__":
    main()
