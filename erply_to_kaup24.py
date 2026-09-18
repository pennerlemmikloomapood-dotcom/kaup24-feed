#!/usr/bin/env python3
"""
Erply -> Kaup24 XML generaator
================================

Mida see teeb:
1. Autendib Erply API-sse (verifyUser)
2. Laeb kõik tooted (getProducts, lehitsedes läbi kõik leheküljed)
3. Filtreerib: ainult active=1 JA displayedInWebshop=1
4. Määrab Kaup24 kategooria toote nime järgi (praegu ainult koer/kass -
   laienda CATEGORY_RULES muutujat, kui lisandub uusi tootegruppe)
5. Kirjutab Kaup24 nõutud XML struktuuri faili kaup24_products.xml

KASUTAMINE:
-----------
1. Paigalda vajalik teek (ühekordselt):
     pip install requests

2. Sea keskkonnamuutujad (ÄRA kirjuta paroole otse skripti!):

   Mac / Linux (terminalis):
     export ERPLY_CLIENT_CODE="37488"
     export ERPLY_USERNAME="apikaup24"
     export ERPLY_PASSWORD="teie_parool"

   Windows (PowerShell):
     $env:ERPLY_CLIENT_CODE="37488"
     $env:ERPLY_USERNAME="apikaup24"
     $env:ERPLY_PASSWORD="teie_parool"

3. Käivita:
     python erply_to_kaup24.py

   Kui tahad kõigepealt vaadata, milline üks toode API-st täpselt
   tagasi tuleb (nt et kontrollida pildi/laoseisu väljade nimesid),
   käivita diagnostika-režiimis:

     python erply_to_kaup24.py --test

   See prindib esimese toote TÄIELIKU JSON-i, ilma XML-i genereerimata.
   Kui pildi- või laoseisu väljad on teistsuguse nimega kui allolevas
   koodis eeldatud, saada see väljund Claude'ile - parandame koodi koos.

TÄHTIS:
-------
Selles skriptis on paar kohta, mida ma EI SAANUD ise Erply API vastu
testida (mul pole otsest võrgujuurdepääsu erply.com domeenile). Need
on märgitud "TODO" kommentaariga. Kõige tõenäolisemalt vajavad
kontrollimist:
  - pildi URL-ide väli (get_image_urls funktsioon)
  - laoseisu väli (get_free_stock funktsioon)
Kui skript jookseb, aga pildid/laoseis tulevad valed või tühjad,
käivita --test režiimis ja saada väljund, et need täpselt ära parandada.
"""

import os
import sys
import json
import time
import requests

# --------------------------------------------------------------------------
# SEADED
# --------------------------------------------------------------------------

CLIENT_CODE = os.environ.get("ERPLY_CLIENT_CODE")
USERNAME = os.environ.get("ERPLY_USERNAME")
PASSWORD = os.environ.get("ERPLY_PASSWORD")

if not CLIENT_CODE:
    print("VIGA: keskkonnamuutuja ERPLY_CLIENT_CODE puudub.")
    sys.exit(1)

API_URL = f"https://{CLIENT_CODE}.erply.com/api/"
OUTPUT_FILE = "kaup24_products.xml"
RECORDS_PER_PAGE = 100  # Erply lubab tavaliselt kuni 100 rea lehekülje kohta

# Kaup24 kategooria vastavustabel Erply tootegruppide (groupID) järgi.
# Ehitatud erply_groups.csv (78 gruppi, 1851 toodet) põhjal, kasutades
# Kaup24/PHH kategooriate faili (Categories_fields_and_values.xlsx).
# Formaat: groupID: (kaup24_category_id, category_name_en)
# None väärtusega grupid on liiga segased (sisaldavad nii koera- kui
# kassitooteid vms) ja jäävad XML-ist hetkel välja - vaata skripti lõpus
# olevat hoiatust ja tee otsus käsitsi.
GROUP_CATEGORY_MAP = {
    156: ("10394", "Toys for dogs"),
    52:  ("10388", "Dog leashes, collars, harnesses"),
    51:  ("10388", "Dog leashes, collars, harnesses"),
    72:  ("10451", "Toys for cats"),
    122: ("10376", "Dog snacks"),
    154: ("10394", "Toys for dogs"),
    76:  ("21813", "Leashes for dogs"),
    53:  ("21813", "Leashes for dogs"),
    9:   ("10415", "Dog clothes"),
    146: ("10394", "Toys for dogs"),
    55:  ("21813", "Leashes for dogs"),
    177: ("10394", "Toys for dogs"),
    48:  ("21558", "Care products for animals"),
    171: ("10442", "Cat leashes, collars, harnesses"),
    116: ("10382", "Bearings, cushions"),
    164: ("10391", "Dog bowls, pet food containers"),
    188: ("10370", "Dry dog food"),          # Go! Solutions (koera liin)
    24:  ("21553", "Cosmetics for animals"),
    172: ("10400", "Dog training products"),
    46:  ("10427", "Cat scratching posts"),
    212: ("10373", "Wet/canned dog food"),   # LandFleisch
    35:  ("21558", "Care products for animals"),
    151: ("10388", "Dog leashes, collars, harnesses"),
    166: ("10424", "Cat litter boxes"),
    69:  ("10412", "Cat snacks"),
    40:  ("10400", "Dog training products"),
    202: ("10490", "Wet/canned cat food"),   # Vibrisse
    161: ("10385", "Dog transport cages, bags"),
    167: ("10403", "Travel accessories for dog"),
    179: ("21813", "Leashes for dogs"),
    112: ("10391", "Dog bowls, pet food containers"),
    178: ("10400", "Dog training products"),
    114: ("21558", "Care products for animals"),  # Muud koeratarbed - segane, üldkategooria
    197: ("10493", "Dry cat food"),          # Go! Solutions (kassi liin)
    108: ("10400", "Dog training products"),
    190: ("10373", "Wet/canned dog food"),   # Alpha Spirit (Poolmärg)
    169: ("10424", "Cat litter boxes"),
    163: ("10403", "Travel accessories for dog"),
    148: ("10394", "Toys for dogs"),
    194: ("10373", "Wet/canned dog food"),   # Amora (koera liin)
    150: ("10400", "Dog training products"),
    149: ("10400", "Dog training products"),
    60:  ("21558", "Care products for animals"),
    170: ("10424", "Cat litter boxes"),
    201: ("10490", "Wet/canned cat food"),   # Amora (kassi liin)
    196: ("10373", "Wet/canned dog food"),   # Alpha Spirit (koera konserv)
    209: ("10370", "Dry dog food"),          # Tales & Tails
    165: ("10391", "Dog bowls, pet food containers"),
    162: ("22346", "Food for farm animals"),
    91:  ("21558", "Care products for animals"),
    11:  ("10421", "Cat litter"),
    191: ("10370", "Dry dog food"),          # Enova
    123: ("10400", "Dog training products"),
    153: ("10400", "Dog training products"),
    103: ("10370", "Dry dog food"),
    113: ("10382", "Bearings, cushions"),
    59:  ("22346", "Food for farm animals"),
    26:  ("21558", "Care products for animals"),
    157: ("21558", "Care products for animals"),
    208: ("10400", "Dog training products"),
    200: ("10493", "Dry cat food"),          # Charm
    27:  ("21558", "Care products for animals"),
    74:  ("10484", "Rodent cages, accessories"),
    168: ("10391", "Dog bowls, pet food containers"),
    160: ("10385", "Dog transport cages, bags"),
    58:  ("10472", "Rodent food"),
    115: ("21558", "Care products for animals"),  # Muud kassitarbed - segane, üldkategooria
    213: ("10478", "Rodent litter, hay"),
    25:  ("21558", "Care products for animals"),  # Parasiiditõrje - segamini koer/kass, üldkategooria
    175: ("10400", "Dog training products"),
    106: ("10493", "Dry cat food"),
    203: ("10490", "Wet/canned cat food"),   # Alpha Spirit (kassi konserv)
    189: ("10373", "Wet/canned dog food"),   # Toortoit - lähim vaste
    158: ("21558", "Care products for animals"),
    159: ("21558", "Care products for animals"),
    199: ("10490", "Wet/canned cat food"),   # Primal Spirit
    206: ("10370", "Dry dog food"),          # Buddy
    10:  ("10385", "Dog transport cages, bags"),
}


# --------------------------------------------------------------------------
# ERPLY API ABIFUNKTSIOONID
# --------------------------------------------------------------------------

def authenticate():
    """Logib Erply API-sse sisse ja tagastab sessionKey."""
    resp = requests.post(
        API_URL,
        data={
            "request": "verifyUser",
            "clientCode": CLIENT_CODE,
            "username": USERNAME,
            "password": PASSWORD,
            "sendContentType": "1",
        },
        timeout=30,
    )
    resp.raise_for_status()
    data = resp.json()
    status = data.get("status", {})
    if status.get("responseStatus") != "ok":
        raise RuntimeError(f"Autentimine ebaõnnestus: {status}")
    return data["records"][0]["sessionKey"]


def api_call(session_key, request_name, **params):
    """Üldine abifunktsioon Erply API kutsumiseks, koos automaatse uuesti-proovimisega."""
    payload = {
        "request": request_name,
        "clientCode": CLIENT_CODE,
        "sessionKey": session_key,
    }
    payload.update(params)

    max_attempts = 4
    for attempt in range(1, max_attempts + 1):
        try:
            resp = requests.post(API_URL, data=payload, timeout=120)
            resp.raise_for_status()
            data = resp.json()
            status = data.get("status", {})
            if status.get("responseStatus") != "ok":
                raise RuntimeError(f"{request_name} API viga: {status}")
            return data.get("records", [])
        except (requests.exceptions.RequestException, RuntimeError) as e:
            if attempt == max_attempts:
                raise
            wait = 5 * attempt
            print(f"  ({request_name} ebaõnnestus, proovin uuesti {wait}s pärast... "
                  f"katse {attempt}/{max_attempts}: {e})")
            time.sleep(wait)


def get_all_products(session_key):
    """Laeb KÕIK tooted Erplyst, lehitsedes läbi kõik leheküljed."""
    all_products = []
    page = 1
    while True:
        print(f"  Laen lehekülge {page}...")
        records = api_call(
            session_key,
            "getProducts",
            recordsOnPage=RECORDS_PER_PAGE,
            pageNo=page,
            active=1,  # küsi Erplylt kohe ainult aktiivseid tooteid
        )
        if not records:
            break
        all_products.extend(records)
        if len(records) < RECORDS_PER_PAGE:
            break
        page += 1
        time.sleep(0.2)  # väike paus, et mitte API limiiti tabada
    return all_products


# --------------------------------------------------------------------------
# ANDMETE TÖÖTLEMINE
# --------------------------------------------------------------------------

def is_active_and_visible(p):
    """Ainult aktiivsed JA e-poes nähtavad tooted."""
    return str(p.get("active")) == "1" and str(p.get("displayedInWebshop")) == "1"


def map_category(p):
    """Tuvastab Kaup24 kategooria Erply groupID järgi."""
    group_id = p.get("groupID")
    try:
        group_id = int(group_id)
    except (TypeError, ValueError):
        return None, None
    result = GROUP_CATEGORY_MAP.get(group_id)
    if result is None:
        return None, None
    return result


def get_image_urls(p):
    """
    Tagastab nimekirja pildi URL-idest.
    TODO: KONTROLLIMATA väli - vaata --test väljundist, kuidas pildid
    tegelikult struktureeritud on, ja paranda seda funktsiooni vastavalt.
    Erply API-s on pildid tavaliselt "images" massiivis, iga pildi
    juures väli nagu "fullURL", "url" või "largeURL".
    """
    urls = []
    for img in p.get("images", []) or []:
        url = img.get("fullURL") or img.get("largeURL") or img.get("url")
        if url:
            urls.append(url)
    return urls


ACTIVE_WAREHOUSE_IDS = [1, 3]  # Aardla Lemmikloomapood, Tartu ladu (teised laod pole kasutusel)


def get_stock_map(session_key):
    """
    Tõmbab KOGU laoseisu aktiivsete ladude jaoks ÜHE päringuga lao kohta
    (mitte ühe päringuga toote kohta - see on Erply API juures palju
    kiirem ja ei tabata päringulimiiti).
    Tagastab dict: {productID: kokku_laoseis_kahes_laos}
    """
    stock_map = {}
    for warehouse_id in ACTIVE_WAREHOUSE_IDS:
        print(f"  Laen laoseisu laost ID {warehouse_id}...")
        records = api_call(session_key, "getProductStock", warehouseID=warehouse_id)
        for rec in records:
            pid = rec["productID"]
            amount = float(rec.get("amountInStock", 0) or 0)
            stock_map[pid] = stock_map.get(pid, 0) + amount
    return stock_map


def cdata(text):
    if text is None:
        text = ""
    text = str(text)
    # Kaup24 juhend keelab HTML erimärgid (&nbsp; jms) - need kuvataks
    # klientidele toore tekstina, mitte tühiku/märgina. Puhastame need.
    text = text.replace("&nbsp;", " ")
    text = text.replace("&sdot;", "·")
    text = text.replace("&rsquo;", "'")
    text = text.replace("&lsquo;", "'")
    text = text.replace("&rdquo;", '"')
    text = text.replace("&ldquo;", '"')
    text = text.replace("]]>", "]]]]><![CDATA[>")
    return f"<![CDATA[ {text} ]]>"


# --------------------------------------------------------------------------
# XML GENEREERIMINE
# --------------------------------------------------------------------------

def build_xml(products, stock_map, out_path):
    lines = ['<?xml version="1.0" encoding="UTF-8"?>', "<products>"]
    skipped_no_category = []
    included_count = 0

    for p in products:
        if not is_active_and_visible(p):
            continue

        name = p.get("name", "")
        cat_id, cat_name = map_category(p)
        if not cat_id:
            skipped_no_category.append((p.get("groupName", ""), name))
            continue

        included_count += 1

        lines.append("  <product>")
        lines.append(f"    <category-id>{cat_id}</category-id>")
        lines.append(f"    <category-name>{cdata(cat_name)}</category-name>")
        lines.append(f"    <title>{cdata(name)}</title>")
        lines.append(f"    <title-ee>{cdata(name)}</title-ee>")

        longdesc = p.get("longdesc", "")
        lines.append(f"    <long-description>{cdata(longdesc)}</long-description>")
        lines.append(f"    <long-description-ee>{cdata(longdesc)}</long-description-ee>")

        lines.append("    <colours>")
        lines.append("      <colour>")
        lines.append("        <images>")
        for url in get_image_urls(p):
            lines.append(f"          <image><url>{url}</url></image>")
        lines.append("        </images>")
        lines.append("        <modifications>")
        lines.append("          <modification>")

        weight = p.get("netWeight") or p.get("weight") or 0
        lines.append(f"            <weight>{weight}</weight>")
        # MÄRKUS: Erplys pole tegelikke pikkus/laius/kõrgus andmeid (kõigil
        # oli sama mõttetu 1/0/0/0 väärtus), aga Kaup24 nõuab neid välju
        # KOHUSTUSLIKUNA. Kasutame ajutist hinnangulist vaikeväärtust (10cm),
        # kuni saate Erplysse sisestada tegelikud pakendi mõõdud.
        lines.append("            <length>0.1</length>")
        lines.append("            <height>0.1</height>")
        lines.append("            <width>0.1</width>")

        lines.append("            <attributes>")
        lines.append("              <barcodes>")
        ean = p.get("code2", "")
        if ean:
            lines.append(f"                <barcode>{cdata(ean)}</barcode>")
        lines.append("              </barcodes>")
        lines.append(f"              <supplier-code>{cdata(p.get('code', ''))}</supplier-code>")
        lines.append("            </attributes>")
        # MÄRKUS: hind ja laoseis EI KUULU siia - Kaup24 ametlik XML-mall ei
        # sisalda <price>/<quantity> elemente. Need lähevad eraldi "Stock
        # and prices import" mehhanismi kaudu (vaata Kaup24 Import menüüst).

        lines.append("          </modification>")
        lines.append("        </modifications>")
        lines.append("      </colour>")
        lines.append("    </colours>")
        lines.append("  </product>")

    lines.append("</products>")

    with open(out_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    print(f"\nValmis! {included_count} toodet kirjutatud faili: {out_path}")

    if skipped_no_category:
        unique_groups = sorted(set(g for g, n in skipped_no_category))
        print(
            f"\nHOIATUS: {len(skipped_no_category)} aktiivset/nähtavat toodet "
            f"jäid kategooria puudumise tõttu XML-ist välja."
        )
        print("Puudutatud tootegrupid:")
        for g in unique_groups:
            count = sum(1 for gg, n in skipped_no_category if gg == g)
            print(f"  - {g} ({count} toodet)")
        print(
            "\nNeed grupid olid teadlikult jäetud märgituna None GROUP_CATEGORY_MAP "
            "tabelis, kuna nad sisaldavad segamini eri liiki tooteid. Vaata need "
            "käsitsi üle ja otsusta õige Kaup24 kategooria, või jäta need "
            "hetkel esimesest saadetisest välja."
        )


# --------------------------------------------------------------------------
# DIAGNOSTIKA (--test režiim)
# --------------------------------------------------------------------------

def run_diagnostic(session_key):
    """Laeb ühe toote ja prindib selle täieliku JSON struktuuri."""
    print("Laen 1 toote diagnostikaks...")
    records = api_call(
        session_key,
        "getProducts",
        recordsOnPage=1,
        pageNo=1,
    )
    if not records:
        print("Ei leidnud ühtegi toodet.")
        return
    print("\n" + "=" * 70)
    print("ESIMESE TOOTE TÄIS JSON (saada see Claude'ile, kui pildid/laoseis")
    print("ei tööta oodatult):")
    print("=" * 70)
    print(json.dumps(records[0], indent=2, ensure_ascii=False))
    print("=" * 70)


# --------------------------------------------------------------------------
# PEAFUNKTSIOON
# --------------------------------------------------------------------------

def main():
    if not all([CLIENT_CODE, USERNAME, PASSWORD]):
        print(
            "VIGA: seadke keskkonnamuutujad ERPLY_CLIENT_CODE, "
            "ERPLY_USERNAME, ERPLY_PASSWORD enne käivitamist."
        )
        sys.exit(1)

    print("Autendin Erply API-sse...")
    session_key = authenticate()
    print("Autentimine õnnestus.\n")

    if "--test" in sys.argv:
        run_diagnostic(session_key)
        return

    print("Laen tooteid Erplyst (see võib mõne minuti aega võtta)...")
    products = get_all_products(session_key)
    print(f"\nLeitud {len(products)} toodet kokku Erplys (kõik staatused).")

    print("\nLaen laoseisu (kiire, ainult mõni päring lao kohta)...")
    stock_map = get_stock_map(session_key)
    print(f"Laoseisu andmed leitud {len(stock_map)} toote kohta.")

    build_xml(products, stock_map, OUTPUT_FILE)


if __name__ == "__main__":
    main()
