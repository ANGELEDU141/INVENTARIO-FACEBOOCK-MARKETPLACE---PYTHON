from playwright.sync_api import sync_playwright


with sync_playwright() as p:

    print("Conectando al Chrome existente...")

    browser = p.chromium.connect_over_cdp(
        "http://127.0.0.1:9222"
    )

    print("Chrome conectado correctamente.")

    if not browser.contexts:
        print("No hay contextos.")
        input("ENTER para salir...")
        raise SystemExit

    for i, context in enumerate(browser.contexts):

        print()
        print("=" * 50)
        print("CONTEXTO:", i)
        print("=" * 50)

        for j, page in enumerate(context.pages):

            print(
                f"Pestaña {j}: {page.url}"
            )

    input("\nENTER para terminar...")