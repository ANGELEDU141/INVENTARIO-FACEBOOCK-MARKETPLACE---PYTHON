import csv
import re
import time
from pathlib import Path

from playwright.sync_api import (
    sync_playwright,
    TimeoutError as PlaywrightTimeoutError
)


# ============================================================
# CONFIGURACIÓN
# ============================================================

FACEBOOK_URL = "https://www.facebook.com/marketplace/you/selling"


# Cuánto desplazamos cada vez
SCROLL_PIXELS = 1100

# Tiempo entre scrolls
SCROLL_WAIT = 1.2

# Cantidad de ciclos sin encontrar productos nuevos
# antes de asumir que llegamos al final.
MAX_SIN_NUEVOS = 7


# ============================================================
# CSV
# ============================================================

CAMPOS = [
    "SKU",
    "listing_id",
    "Nombre",
    "Precio",
    "Disponibilidad"
]


def cargar_existentes():

    existentes = set()

    if not CSV_FILE.exists():
        return existentes

    try:

        with open(
            CSV_FILE,
            "r",
            encoding="utf-8-sig",
            newline=""
        ) as archivo:

            lector = csv.DictReader(
                archivo,
                delimiter=";"
            )

            for fila in lector:

                listing_id = (
                    fila.get("listing_id") or ""
                ).strip()

                if listing_id:
                    existentes.add(listing_id)

    except Exception as e:

        print(
            f"[AVISO] No se pudo leer CSV existente: {e}"
        )

    return existentes


def guardar_registro(registro):

    existe = CSV_FILE.exists()

    with open(
        CSV_FILE,
        "a",
        encoding="utf-8-sig",
        newline=""
    ) as archivo:

        escritor = csv.DictWriter(
            archivo,
            fieldnames=CAMPOS,
            delimiter=";"
        )

        if not existe:
            escritor.writeheader()

        escritor.writerow(registro)

        archivo.flush()


# ============================================================
# EXTRACCIÓN
# ============================================================

def extraer_precio(texto):

    if not texto:
        return ""

    texto = texto.replace("\xa0", " ")

    # Ejemplos:
    # S/693
    # S/ 693
    # S/1,299
    # S/ 1,299.50

    match = re.search(
        r"S\/\s*([\d.,]+)",
        texto
    )

    if not match:
        return ""

    return match.group(1).strip()


def extraer_sku(texto):

    if not texto:
        return ""

    texto = texto.replace("\xa0", " ")

    # Estructura observada:
    #
    # 3023
    # Publicación:
    #
    match = re.search(
        r"(?:^|\n)\s*(\d+)\s*\n\s*Publicación\s*:",
        texto,
        re.IGNORECASE
    )

    if match:
        return match.group(1)

    return ""


def extraer_listing_id(modal):

    try:

        enlaces = modal.locator(
            'a[href*="listing_id="]'
        )

        cantidad = enlaces.count()

        for i in range(cantidad):

            try:

                href = enlaces.nth(i).get_attribute(
                    "href"
                )

                if not href:
                    continue

                match = re.search(
                    r"listing_id=(\d+)",
                    href
                )

                if match:
                    return match.group(1)

            except Exception:
                continue

    except Exception:
        pass

    return ""


def extraer_nombre(modal, nombre_tarjeta):

    # El nombre de la tarjeta ya es confiable.
    # Lo usamos primero porque Facebook lo entrega
    # en aria-label.

    if nombre_tarjeta:
        return nombre_tarjeta.strip()

    try:

        texto = modal.inner_text()

        lineas = [
            x.strip()
            for x in texto.splitlines()
            if x.strip()
        ]

        for linea in lineas:

            if linea == "Tu publicación":
                continue

            if linea.startswith("S/"):
                continue

            if linea == "Disponible":
                continue

            if linea == "Agotado":
                continue

            if "Publicación:" in linea:
                continue

            return linea

    except Exception:
        pass

    return ""


def extraer_disponibilidad(texto):

    if not texto:
        return ""

    texto = texto.replace("\xa0", " ")

    if re.search(
        r"\bDisponible\b",
        texto,
        re.IGNORECASE
    ):
        return "Disponible"

    if re.search(
        r"\bAgotado\b",
        texto,
        re.IGNORECASE
    ):
        return "Agotado"

    # En tu modal aparece:
    #
    # Marcar como agotado
    #
    # Esto significa que actualmente está disponible.

    if re.search(
        r"Marcar como agotado",
        texto,
        re.IGNORECASE
    ):
        return "Disponible"

    return ""


# ============================================================
# MODAL
# ============================================================

def obtener_modal(page):

    modal = page.locator(
        '[aria-label="Tu publicación"]'
    )

    try:

        modal.wait_for(
            state="visible",
            timeout=7000
        )

        return modal

    except PlaywrightTimeoutError:

        return None


def cerrar_modal(page):

    # ESC es la forma más rápida.
    try:

        page.keyboard.press("Escape")

        page.wait_for_timeout(300)

    except Exception:
        pass

    # Comprobar que realmente desapareció.
    try:

        modal = page.locator(
            '[aria-label="Tu publicación"]'
        )

        if modal.count() > 0:

            try:

                modal.wait_for(
                    state="hidden",
                    timeout=1500
                )

            except Exception:

                # Intentar botón cerrar
                botones = page.locator(
                    '[aria-label="Tu publicación"] '
                    'button[aria-label="Cerrar"]'
                )

                if botones.count() > 0:

                    try:
                        botones.first.click(
                            timeout=1000
                        )
                    except Exception:
                        pass

    except Exception:
        pass

    page.wait_for_timeout(250)


# ============================================================
# TARJETAS
# ============================================================

def obtener_tarjetas(page):

    """
    Busca las tarjetas de productos.

    La estructura que comprobamos en tu Marketplace es:

        div[role="button"][aria-label]

    y el aria-label contiene el nombre del producto.

    Excluimos elementos relacionados con el modal.
    """

    try:

        elementos = page.locator(
            'div[role="button"][aria-label]'
        )

        cantidad = elementos.count()

        tarjetas = []

        for i in range(cantidad):

            try:

                elemento = elementos.nth(i)

                aria = (
                    elemento.get_attribute(
                        "aria-label"
                    ) or ""
                ).strip()

                if not aria:
                    continue

                # No queremos botones del modal.
                if aria in [
                    "Tu publicación",
                    "Cerrar"
                ]:
                    continue

                # Las tarjetas reales contienen una imagen
                # con alt igual al nombre.
                imagenes = elemento.locator(
                    "img[alt]"
                )

                if imagenes.count() == 0:
                    continue

                # La tarjeta también debe contener un precio.
                texto = elemento.inner_text()

                if not re.search(
                    r"S\/\s*[\d.,]+",
                    texto
                ):
                    continue

                tarjetas.append(elemento)

            except Exception:
                continue

        return tarjetas

    except Exception:

        return []


# ============================================================
# PROCESAR TARJETA
# ============================================================

def procesar_tarjeta(
    page,
    tarjeta,
    nombre_tarjeta
):

    print()
    print("--------------------------------------------")
    print("PRODUCTO")
    print(nombre_tarjeta)

    try:

        # Click en la tarjeta
        tarjeta.click(
            timeout=5000
        )

    except Exception as e:

        print(
            "[ERROR] No se pudo abrir la tarjeta:",
            e
        )

        return None

    # Esperar modal real
    modal = obtener_modal(page)

    if modal is None:

        print(
            "[ERROR] No apareció 'Tu publicación'."
        )

        cerrar_modal(page)

        return None

    try:

        # Dar un pequeño margen al renderizado interno.
        page.wait_for_timeout(250)

        texto = modal.inner_text()

        # --------------------------------------------
        # DATOS
        # --------------------------------------------

        sku = extraer_sku(texto)

        precio = extraer_precio(texto)

        listing_id = extraer_listing_id(
            modal
        )

        nombre = extraer_nombre(
            modal,
            nombre_tarjeta
        )

        disponibilidad = (
            extraer_disponibilidad(texto)
        )

        resultado = {
            "SKU": sku,
            "listing_id": listing_id,
            "Nombre": nombre,
            "Precio": precio,
            "Disponibilidad": disponibilidad
        }

        print()
        print("SKU:", sku)
        print("Listing ID:", listing_id)
        print("Precio:", precio)
        print("Disponibilidad:", disponibilidad)

        if not sku:
            print(
                "[AVISO] No se pudo extraer SKU."
            )

        if not listing_id:
            print(
                "[AVISO] No se pudo extraer listing_id."
            )

        return resultado

    except Exception as e:

        print(
            "[ERROR] Extrayendo datos:",
            e
        )

        return None

    finally:

        cerrar_modal(page)


# ============================================================
# MAIN
# ============================================================

def main():

    print()
    print("============================================")
    print(" EXTRACTOR FACEBOOK MARKETPLACE")
    print("============================================")
    print()

    print("Perfil de extracción:")
    print(CHROME_USER_DATA)
    # --------------------------------------------
    # Cargar registros anteriores
    # --------------------------------------------

    listing_ids_existentes = (
        cargar_existentes()
    )

    print()

    if listing_ids_existentes:

        print(
            "Registros existentes:",
            len(listing_ids_existentes)
        )

    else:

        print(
            "No hay registros anteriores."
        )

    # --------------------------------------------
    # Playwright
    # --------------------------------------------

    with sync_playwright() as p:

        print()
        print("Iniciando Chrome...")

        try:

                context = p.chromium.launch_persistent_context(
                user_data_dir=CHROME_USER_DATA,
                headless=False,
                args=[
                    "--start-maximized",
                    "--no-first-run",
                    "--no-default-browser-check"
                ],
                viewport=None,
                record_video_dir=None
                  )

        except Exception as e:

            print()
            print("============================================")
            print("ERROR ABRIENDO EL PERFIL DE CHROME")
            print("============================================")
            print()
            print(e)
            print()
            print(
                "Asegúrate de que Chrome esté completamente cerrado."
            )

            input(
                "\nPresiona ENTER para salir..."
            )

            return

     # --------------------------------------------
        # Página
        # --------------------------------------------

        if context.pages:

            page = context.pages[0]

        else:

            page = context.new_page()

        print()
        print("Página inicial:")
        print(page.url)


        # --------------------------------------------
        # Reducir recursos innecesarios
        # --------------------------------------------

        def controlar_recursos(route):

            tipo = route.request.resource_type

            # Bloqueamos únicamente recursos pesados
            # que no necesitamos para extraer los datos.
            #
            # NO bloqueamos imágenes porque las tarjetas
            # utilizan img[alt] para identificarse.

            if tipo in [
                "media",
                "font"
            ]:

                route.abort()

            else:

                route.continue_()


        page.route(
            "**/*",
            controlar_recursos
        )

        # --------------------------------------------
        # Facebook
        # --------------------------------------------

        print()
        print("Abriendo Facebook...")

        try:

            page.goto(
                "https://www.facebook.com/",
                wait_until="domcontentloaded",
                timeout=60000
            )

        except Exception as e:

            print()
            print(
                "[AVISO] Facebook no terminó la navegación:"
            )

            print(e)

        page.wait_for_timeout(3000)

        print()
        print("URL actual:")
        print(page.url)

        # --------------------------------------------
        # LOGIN / SESIÓN
        # --------------------------------------------

        print()
        print("--------------------------------------------")
        print("VERIFICACIÓN DE SESIÓN")
        print("--------------------------------------------")

        print(
            "Si Facebook muestra alguna pantalla de"
        )
        print(
            "inicio de sesión, verificación o checkpoint,"
        )
        print(
            "resuélvela manualmente."
        )

        print()
        print(
            "Cuando Facebook esté completamente abierto,"
        )
        print(
            "presiona ENTER aquí."
        )

        input()

        # --------------------------------------------
        # Marketplace
        # --------------------------------------------

        print()
        print("Abriendo tus publicaciones...")

        try:

            page.goto(
                FACEBOOK_URL,
                wait_until="domcontentloaded",
                timeout=60000
            )

        except Exception as e:

            print()
            print(
                "[AVISO] Error de navegación de Marketplace:"
            )

            print(e)

        page.wait_for_timeout(4000)

        print()
        print("URL:")
        print(page.url)

        # --------------------------------------------
        # Comprobar que Facebook está realmente cargado
        # --------------------------------------------

        try:

            tarjetas_iniciales = obtener_tarjetas(
                page
            )

        except Exception:

            tarjetas_iniciales = []

        print()
        print(
            "Tarjetas encontradas inicialmente:",
            len(tarjetas_iniciales)
        )

        if not tarjetas_iniciales:

            print()
            print("============================================")
            print("NO SE ENCONTRARON PUBLICACIONES")
            print("============================================")
            print()
            print(
                "Verifica que Facebook esté mostrando:"
            )
            print(
                "Marketplace > Tus publicaciones"
            )
            print()

            input(
                "Presiona ENTER para cerrar..."
            )

            context.close()

            return

        # --------------------------------------------
        # CONTROL
        # --------------------------------------------

        # Firmas de tarjetas ya examinadas.
        tarjetas_vistas = set()

        # IDs encontrados.
        listing_ids = set(
            listing_ids_existentes
        )

        sin_nuevos = 0

        total_guardados = len(
            listing_ids_existentes
        )

        # --------------------------------------------
        # BUCLE
        # --------------------------------------------

        while True:

            tarjetas = obtener_tarjetas(page)

            print()
            print("============================================")
            print(
                "Tarjetas en DOM:",
                len(tarjetas)
            )
            print(
                "Productos guardados:",
                total_guardados
            )
            print("============================================")

            nuevos = 0

            # ----------------------------------------
            # Procesar tarjetas
            # ----------------------------------------

            for tarjeta in tarjetas:

                try:

                    nombre = (
                        tarjeta.get_attribute(
                            "aria-label"
                        ) or ""
                    ).strip()

                    if not nombre:
                        continue

                    # --------------------------------
                    # Firma de tarjeta
                    # --------------------------------

                    try:

                        texto_tarjeta = (
                            tarjeta.inner_text()
                        )

                    except Exception:

                        texto_tarjeta = ""

                    firma = (
                        nombre,
                        texto_tarjeta
                    )

                    if firma in tarjetas_vistas:

                        continue

                    tarjetas_vistas.add(
                        firma
                    )

                    # --------------------------------
                    # Procesar
                    # --------------------------------

                    resultado = procesar_tarjeta(
                        page,
                        tarjeta,
                        nombre
                    )

                    if not resultado:

                        continue

                    listing_id = (
                        resultado["listing_id"]
                    )

                    # --------------------------------
                    # Si ya existe
                    # --------------------------------

                    if listing_id:

                        if listing_id in listing_ids:

                            print(
                                "[YA EXISTE] "
                                "No se vuelve a guardar."
                            )

                            continue

                        listing_ids.add(
                            listing_id
                        )

                    # --------------------------------
                    # Guardar
                    # --------------------------------

                    guardar_registro(
                        resultado
                    )

                    total_guardados += 1
                    nuevos += 1

                    print(
                        "[GUARDADO]",
                        total_guardados
                    )

                except Exception as e:

                    print()
                    print(
                        "[ERROR TARJETA]",
                        e
                    )

                    cerrar_modal(page)

                    continue

            # ----------------------------------------
            # Control de final
            # ----------------------------------------

            print()
            print(
                "Nuevos en esta vuelta:",
                nuevos
            )

            if nuevos == 0:

                sin_nuevos += 1

            else:

                sin_nuevos = 0

            print(
                "Vueltas sin nuevos:",
                sin_nuevos,
                "/",
                MAX_SIN_NUEVOS
            )

            if sin_nuevos >= MAX_SIN_NUEVOS:

                print()
                print(
                    "No se detectan publicaciones nuevas."
                )

                break

            # ----------------------------------------
            # Scroll
            # ----------------------------------------

            print()
            print(
                "Haciendo scroll..."
            )

            try:

                page.mouse.wheel(
                    0,
                    SCROLL_PIXELS
                )

            except Exception as e:

                print(
                    "[ERROR SCROLL]",
                    e
                )

            page.wait_for_timeout(
                int(
                    SCROLL_WAIT * 1000
                )
            )

        # --------------------------------------------
        # FINAL
        # --------------------------------------------

        print()
        print("============================================")
        print(" EXTRACCIÓN TERMINADA")
        print("============================================")
        print()
        print(
            "Total guardados:",
            total_guardados
        )
        print()
        print(
            "CSV:"
        )
        print(
            CSV_FILE.resolve()
        )

        print()
        print(
            "El archivo está listo para Excel."
        )

        input(
            "\nPresiona ENTER para cerrar Chrome..."
        )

        context.close()


# ============================================================
# EJECUCIÓN
# ============================================================

if __name__ == "__main__":
    main()