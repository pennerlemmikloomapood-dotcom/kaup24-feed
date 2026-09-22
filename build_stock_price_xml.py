"""
Erply -> Kaup24 Stock & Prices XML generaator
================================================

Genereerib "Stock and prices import" faili Kaup24 ametliku malli järgi.
Kasutab TÄPSELT samu tooteid, mis on ka Products XML-is (erply_to_kaup24.py) -
sama kvalifitseerumisloogika (kategooria, EAN, kaal, pilt, kirjeldus).

KASUTAMINE: täpselt sama moodi nagu erply_to_kaup24.py (samad
keskkonnamuutujad ERPLY_CLIENT_CODE / ERPLY_USERNAME / ERPLY_PASSWORD).

    python build_stock_price_xml.py

Väljund: kaup24_stock_prices.xml
"""

from erply_to_kaup24 import (
    authenticate, get_all_products, get_stock_map, iter_qualifying_products,
)

OUTPUT_FILE = "kaup24_stock_prices.xml"

# Kasutame sama hinda kõigis neljas riigis (LT/LV/EE/FI) - kinnitatud.
# Komplekteerimisaeg: Eestis 24-48h, mujal 72h.
COLLECTION_HOURS_EE = "48"
COLLECTION_HOURS_OTHER = "72"


def build_stock_price_xml(products, stock_map, out_path):
    lines = ['<?xml version="1.0" encoding="utf-8"?>', "<products>"]
    included_count = 0

    for p, cat_id, cat_name, ean, weight, images, longdesc, stats in iter_qualifying_products(products):
        included_count += 1

        price = p.get("priceWithVAT") or p.get("price") or 0
        # Praegu allahindlust ei rakendata - hind ja allahindlusjärgne hind
        # on samad. Kui hiljem lisandub Erplys sooduskampaania, saab siia
        # eraldi allahinnatud hinna välja lisada.
        price_after_discount = price
        stock = int(stock_map.get(p["productID"], 0))

        supplier_code = p.get("code", "")

        lines.append("  <product>")
        lines.append(f"    <sku>{supplier_code}</sku>")
        lines.append(f"    <ean>{ean}</ean>")

        for country, hours in [
            ("lt", COLLECTION_HOURS_OTHER),
            ("lv", COLLECTION_HOURS_OTHER),
            ("ee", COLLECTION_HOURS_EE),
            ("fi", COLLECTION_HOURS_OTHER),
        ]:
            lines.append(f"    <price_{country}>{price}</price_{country}>")
            lines.append(f"    <price_after_discount_{country}>{price_after_discount}</price_after_discount_{country}>")

        lines.append(f"    <stock>{stock}</stock>")

        for country, hours in [
            ("lt", COLLECTION_HOURS_OTHER),
            ("lv", COLLECTION_HOURS_OTHER),
            ("ee", COLLECTION_HOURS_EE),
            ("fi", COLLECTION_HOURS_OTHER),
        ]:
            lines.append(f"    <collection_hours_{country}>{hours}</collection_hours_{country}>")

        lines.append("  </product>")

    lines.append("</products>")

    with open(out_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    print(f"\nValmis! {included_count} toodet kirjutatud faili: {out_path}")


def main():
    print("Autendin Erply API-sse...")
    session_key = authenticate()
    print("Autentimine õnnestus.\n")

    print("Laen tooteid Erplyst...")
    products = get_all_products(session_key)
    print(f"Leitud {len(products)} toodet kokku Erplys.")

    print("\nLaen laoseisu...")
    stock_map = get_stock_map(session_key)
    print(f"Laoseisu andmed leitud {len(stock_map)} toote kohta.")

    build_stock_price_xml(products, stock_map, OUTPUT_FILE)


if __name__ == "__main__":
    main()
