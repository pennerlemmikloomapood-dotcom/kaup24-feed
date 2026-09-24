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
# PIM API - eraldi, uuem Erply liides, mis annab LT/LV tõlkeid, mida
# klassikaline API (ülal) ei paku. Aadress leiti getServiceEndpoints kaudu -
# kui see kunagi muutub, tuleb see uuesti kontrollida.
PIM_BASE_URL = "https://api-pim-eu.erply.com/"
RECORDS_PER_PAGE = 100  # Erply lubab tavaliselt kuni 100 rea lehekülje kohta

# Kaup24 kategooria vastavustabel Erply tootegruppide (groupID) järgi.
# Ehitatud erply_groups.csv (78 gruppi, 1851 toodet) põhjal, kasutades
# Kaup24/PHH kategooriate faili (Categories_fields_and_values.xlsx).
# Formaat: groupID: (kaup24_category_id, category_name_en)
# None väärtusega grupid on liiga segased (sisaldavad nii koera- kui
# kassitooteid vms) ja jäävad XML-ist hetkel välja - vaata skripti lõpus
# olevat hoiatust ja tee otsus käsitsi.
# Brändid, mida EI TOHI Kaup24-sse saata (kirjutage täpselt nii, nagu
# Erplys "Kaubamärk" väljal kirjas, suur/väiketäht ei loe).
# Näide: EXCLUDED_BRANDS = {"GO! Solutions", "Alpha Spirit"}
EXCLUDED_BRANDS = set()

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
    Tagastab nimekirja pildi URL-idest. Kaup24 lubab ainult JPG/JPEG/PNG
    pilte - filtreerime välja teised formaadid (nt .webp, .gif).
    """
    urls = []
    for img in p.get("images", []) or []:
        url = img.get("fullURL") or img.get("largeURL") or img.get("url")
        if url and url.lower().split("?")[0].endswith((".jpg", ".jpeg", ".png")):
            urls.append(url)
    return urls


ACTIVE_WAREHOUSE_IDS = [1, 3]  # Aardla Lemmikloomapood, Tartu ladu (teised laod pole kasutusel)


def get_pim_translation(session_key, product_id):
    """
    Küsib PIM API-lt ühe toote LT/LV (ja teiste keelte) nime/kirjeldust.
    Tagastab dict {"name": {...}, "description": {...}} või None, kui
    päring ebaõnnestub (nt toodet pole PIM-is, ajutine viga vms) - sel
    juhul kasutab helistaja lihtsalt olemasolevat eestikeelset varianti.
    """
    url = f"{PIM_BASE_URL}v1/product/{product_id}"
    headers = {"clientCode": CLIENT_CODE, "sessionKey": session_key}
    try:
        resp = requests.get(url, headers=headers, timeout=20)
        if resp.status_code != 200:
            return None
        data = resp.json()
        if not data:
            return None
        product = data[0] if isinstance(data, list) else data
        return {
            "name": product.get("name", {}) or {},
            "description": product.get("description", {}) or {},
        }
    except (requests.exceptions.RequestException, ValueError, KeyError, IndexError):
        return None


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
    text = str(text).strip()
    # Kaup24 juhend keelab HTML erimärgid (&nbsp; jms) - need kuvataks
    # klientidele toore tekstina, mitte tühiku/märgina. Puhastame need.
    text = text.replace("&nbsp;", " ")
    text = text.replace("&sdot;", "·")
    text = text.replace("&rsquo;", "'")
    text = text.replace("&lsquo;", "'")
    text = text.replace("&rdquo;", '"')
    text = text.replace("&ldquo;", '"')
    text = text.replace("&amp;", "&")
    text = text.replace("&quot;", '"')
    text = text.replace("]]>", "]]]]><![CDATA[>")
    return f"<![CDATA[ {text} ]]>"


import re

_VOLUME_PATTERN = re.compile(
    r"(\d+(?:[.,]\d+)?)\s*(ml|l)\b", re.IGNORECASE
)


def extract_volume_liters(*texts):
    """
    Otsib tekstist (nt tootenimi, kirjeldus) mahtu liitrites või ml-des
    ja tagastab selle liitrites (nt "0.5"). Otsib antud tekstide
    järjekorras, tagastab esimese leitud vaste.
    """
    for text in texts:
        if not text:
            continue
        match = _VOLUME_PATTERN.search(str(text))
        if match:
            value = float(match.group(1).replace(",", "."))
            unit = match.group(2).lower()
            if unit == "ml":
                value = value / 1000
            # Ümardame mõistlikult, eemaldades tarbetu .0
            return f"{value:g}"
    return None


_COMPOSITION_PATTERN = re.compile(
    r"koostis\s*:\s*(?:</?\w+[^>]*>\s*)*(.*?)<br", re.IGNORECASE | re.DOTALL
)


def extract_composition(html_text):
    """
    Otsib kirjeldusest "Koostis:" järgset teksti kuni järgmise <br>-ni,
    puhastab selle HTML-tagidest ja tagastab lihttekstina. None, kui
    "Koostis:" ei leitud.
    """
    if not html_text:
        return None
    match = _COMPOSITION_PATTERN.search(str(html_text))
    if not match:
        return None
    text = match.group(1)
    text = re.sub(r"<[^>]+>", "", text)
    text = text.replace("&nbsp;", " ").strip()
    return text or None


_MANUFACTURER_PATTERN = re.compile(
    r"tootja\s*:\s*(?:</?\w+[^>]*>\s*)*(.*?)<br", re.IGNORECASE | re.DOTALL
)


def extract_manufacturer_info(html_text):
    """
    Otsib kirjeldusest "Tootja:" järgset teksti kuni järgmise <br>-ni
    (nt "Croci S.P.A; Via Sant'Alessandro 8, Castronno (Va), Itaalia").
    Puhastab HTML-tagidest ja tagastab lihttekstina. None, kui ei leitud.
    """
    if not html_text:
        return None
    match = _MANUFACTURER_PATTERN.search(str(html_text))
    if not match:
        return None
    text = match.group(1)
    text = re.sub(r"<[^>]+>", "", text)
    text = text.replace("&nbsp;", " ").strip()
    return text or None


def split_manufacturer_info(raw_text):
    """
    Erplys on "Tootja" väli vahel täidetud kombineeritud kujul, nt:
    "Flamingo Pet Products, Lammerdries - Winkelstraat 25, 2250 Olen,
    Belgia, info@flamingo.be" - nimi, aadress ja e-post koos, komadega
    eraldatud. Lahutab need kolmeks: (name, address, email).
    Kui e-posti ei leitud, eeldab kogu teksti olevat lihtsalt nimi
    (nt "PPN Limited Partnership" ilma aadressita).
    """
    if not raw_text:
        return "", "", ""
    raw_text = str(raw_text).strip()
    email_match = re.search(r"[\w.+-]+@[\w-]+\.[\w.-]+", raw_text)
    if not email_match:
        return raw_text, "", ""
    email = email_match.group(0)
    remainder = raw_text.replace(email, "").strip().rstrip(",").strip()
    parts = [p.strip() for p in remainder.split(",") if p.strip()]
    name = parts[0] if parts else ""
    address = ", ".join(parts[1:]) if len(parts) > 1 else ""
    return name, address, email


_DISALLOWED_REMOVE_WITH_CONTENT = ["table", "iframe", "script"]
_DISALLOWED_STRIP_TAG_ONLY = ["a", "font", "h1", "h4", "h5", "h6", "u"]


def strip_disallowed_html(html_text):
    """
    Kaup24 juhend (p.7) lubab ainult: p, br, b, strong, i, em, ul, ol, li,
    div, span, h2, h3. Eemaldab keelatud elemendid kas koos sisuga
    (table/iframe/script) või jätab ainult sisu alles (a/font/h1 jne).
    """
    if not html_text:
        return html_text
    text = str(html_text)
    for tag in _DISALLOWED_REMOVE_WITH_CONTENT:
        text = re.sub(rf"<{tag}\b[^>]*>.*?</{tag}>", "", text, flags=re.IGNORECASE | re.DOTALL)
    for tag in _DISALLOWED_STRIP_TAG_ONLY:
        text = re.sub(rf"</?{tag}\b[^>]*>", "", text, flags=re.IGNORECASE)
    return text


# --------------------------------------------------------------------------
# XML GENEREERIMINE
# --------------------------------------------------------------------------

def clean_ean(raw_ean):
    """
    Puhastab EAN-i, eemaldades lõpust tähed (partii/kuupäeva tähis, nt
    "4260591822177A" -> "4260591822177"). Tagastab (puhas_ean, täht) või
    (None, None), kui tulemus pole kehtiv EAN (11-13 numbrit).
    """
    raw_ean = str(raw_ean or "").strip()
    suffix_letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    cleaned = raw_ean.rstrip("abcdefghijklmnopqrstuvwxyz").rstrip(suffix_letters)
    letter = raw_ean[len(cleaned):].upper() if len(cleaned) < len(raw_ean) else ""
    if cleaned.isdigit() and 11 <= len(cleaned) <= 13:
        return cleaned, letter
    return None, None


def find_ean_winners(products):
    """
    Mõnel tootel (nt Go! Solutions partiid) on EAN-i lõpus täht, mis
    tähistab partiid/kuupäeva (A = vanim, aegub esimesena). Kui mitu
    toodet annavad sama puhastatud EAN-i, valime AINULT kõige suurema
    tähega (värskeima partii) - vanemad partiid jäävad XML-ist välja.
    Tagastab: {productID: True} võitjate kohta, ja nimekirja väljajäetutest.
    """
    by_ean = {}  # cleaned_ean -> list of (letter, productID, name, raw_ean)
    for p in products:
        if not is_active_and_visible(p):
            continue
        cleaned, letter = clean_ean(p.get("code2", ""))
        if not cleaned:
            continue
        by_ean.setdefault(cleaned, []).append((letter, p["productID"], p.get("name", ""), p.get("code2", "")))

    winners = set()
    losers = []  # (name, raw_ean) - uuemad partiid, mis jäävad hetkel välja
    for cleaned, entries in by_ean.items():
        entries.sort(key=lambda e: e[0])  # täheta ("") < "A" < "B" < "C" ...
        winner = entries[0]  # vanim/esimene partii (aegub esimesena, müügis kõige enne - FIFO)
        winners.add(winner[1])
        for letter, pid, name, raw_ean in entries[1:]:
            losers.append((name, raw_ean))
    return winners, losers


def iter_qualifying_products(products):
    """
    Ühine "kvalifitseerumise" loogika: käib läbi kõik Erply tooted ja
    tagastab (generaatorina) ainult need, mis vastavad KÕIGILE Kaup24
    nõuetele - aktiivne+nähtav, kategoriseeritud, kehtiv EAN (uusima
    partii omast), kaal, pilt ja kirjeldus olemas.

    Kasutatakse nii Products XML kui Stock/Price XML genereerimisel,
    et mõlemad failid sisaldaksid täpselt samu tooteid.

    Annab iga kvalifitseeruva toote kohta: (p, ean, weight, images, longdesc)
    """
    ean_winners, skipped_old_batch = find_ean_winners(products)
    stats = {
        "skipped_no_category": [], "skipped_no_ean": [],
        "zero_weight_count": 0, "no_image_count": 0, "empty_desc_count": 0,
        "skipped_old_batch": skipped_old_batch,
    }

    for p in products:
        if not is_active_and_visible(p):
            continue

        brand_check = (p.get("brandName") or "").strip().lower()
        if brand_check and brand_check in {b.lower() for b in EXCLUDED_BRANDS}:
            continue

        name = p.get("name", "")
        cat_id, cat_name = map_category(p)
        if not cat_id:
            stats["skipped_no_category"].append((p.get("groupName", ""), name))
            continue

        cleaned_ean, letter = clean_ean(p.get("code2", ""))
        if not cleaned_ean:
            stats["skipped_no_ean"].append(name)
            continue
        if p["productID"] not in ean_winners:
            continue
        ean = cleaned_ean

        weight = p.get("grossWeight") or p.get("netWeight") or p.get("weight") or 0
        images = get_image_urls(p)
        longdesc = strip_disallowed_html(p.get("longdesc", "")).strip()

        if not weight or float(weight) == 0:
            stats["zero_weight_count"] += 1
            continue
        if not images:
            stats["no_image_count"] += 1
            continue
        if not longdesc:
            stats["empty_desc_count"] += 1
            continue

        yield p, cat_id, cat_name, ean, weight, images, longdesc, stats


# Lemmikloomatoidu kategooriad, mis vajavad "Looma vanus" ja "Eriomadus" välju
_PET_FOOD_CATEGORIES = {
    "10370": "koer",   # Dry dog food
    "10373": "koer",   # Wet dog food
    "10376": "koer",   # Dog snacks
    "10493": "kass",   # Dry cat food
    "10490": "kass",   # Wet cat food
    "10412": "kass",   # Cat snacks
}


def derive_animal_age(name):
    """Tuletab 'Looma vanus' väärtuse tootenimest, vaikimisi üldine."""
    name_lower = (name or "").lower()
    if any(kw in name_lower for kw in ["kutsika", "kassipoja", "kassipojale"]):
        return "Kutsikas/kassipoeg"
    if any(kw in name_lower for kw in ["seenior", "vanematele", "senior"]):
        return "Seenior"
    return "Erinevatele vanustele"


def derive_special_feature(name, species):
    """Tuletab 'Eriomadus' väärtuse - täpsem tootenimest, muidu üldine liigi järgi."""
    name_lower = (name or "").lower()
    if "steriliseeri" in name_lower:
        return f"Steriliseeritud {'kassidele' if species == 'kass' else 'koertele'}"
    if "tundlik" in name_lower:
        return "Tundliku seedimisega loomadele"
    if "ülekaalul" in name_lower:
        return "Ülekaalulistele loomadele"
    return f"Kõigile {'kassidele' if species == 'kass' else 'koertele'}"


def build_xml(products, stock_map, out_path, session_key=None):
    lines = ['<?xml version="1.0" encoding="UTF-8"?>', "<products>"]
    included_count = 0
    final_stats = {}

    for p, cat_id, cat_name, ean, weight, images, longdesc, stats in iter_qualifying_products(products):
        final_stats = stats
        included_count += 1
        name = p.get("name", "")

        # Proovime PIM API-st LT/LV tõlkeid (need puuduvad klassikalisest
        # API-st). Kui PIM päring ebaõnnestub või tõlget pole, kasutame
        # eestikeelset teksti varulahendusena <title> (LT-slot) jaoks.
        pim_data = get_pim_translation(session_key, p["productID"]) if session_key else None
        if session_key:
            time.sleep(0.05)  # väike paus, et mitte PIM API limiiti tabada
        pim_name = (pim_data or {}).get("name", {})
        pim_desc = (pim_data or {}).get("description", {})

        title_lt = pim_name.get("lt") or name  # LT puudumisel: eesti tekst
        title_lv = pim_name.get("lv") or ""    # LV puudumisel: jätame välja

        lines.append("  <product>")
        lines.append(f"    <category-id>{cat_id}</category-id>")
        lines.append(f"    <category-name>{cdata(cat_name)}</category-name>")
        lines.append(f"    <title>{cdata(title_lt)}</title>")
        if title_lv:
            lines.append(f"    <title-lv>{cdata(title_lv)}</title-lv>")
        lines.append(f"    <title-ee>{cdata(name)}</title-ee>")

        # LT ja LV kirjeldused PIM-ist, kui olemas
        desc_lt = (pim_desc.get("lt") or {}).get("plain_text") or (pim_desc.get("lt") or {}).get("html") or ""
        desc_lt = strip_disallowed_html(desc_lt).strip()
        long_description_base = desc_lt or longdesc  # LT puudumisel: eesti tekst varulahendusena

        lines.append(f"    <long-description>{cdata(long_description_base)}</long-description>")
        lines.append(f"    <long-description-ee>{cdata(longdesc)}</long-description-ee>")

        desc_lv = (pim_desc.get("lv") or {}).get("plain_text") or (pim_desc.get("lv") or {}).get("html") or ""
        desc_lv = strip_disallowed_html(desc_lv).strip()
        if desc_lv:
            lines.append(f"    <long-description-lv>{cdata(desc_lv)}</long-description-lv>")

        # Vene ja soome kirjeldused - lisame AINULT kui Erplys olemas
        # (paljudel toodetel puuduvad, see on täiesti OK - EE juba katab
        # kohustusliku "vähemalt 1 keel" nõude).
        longdesc_ru = strip_disallowed_html(p.get("longdescRUS", "")).strip()
        if longdesc_ru:
            lines.append(f"    <long-description-ru>{cdata(longdesc_ru)}</long-description-ru>")
        longdesc_fi = strip_disallowed_html(p.get("longdescFIN", "")).strip()
        if longdesc_fi:
            lines.append(f"    <long-description-fi>{cdata(longdesc_fi)}</long-description-fi>")
        longdesc_en = strip_disallowed_html(p.get("longdescENG", "")).strip()
        if longdesc_en:
            lines.append(f"    <long-description-en>{cdata(longdesc_en)}</long-description-en>")


        # Koostis (kohustuslik lemmikloomatoidu aktiveerimiseks): tõmmatud
        # automaatselt kirjeldusest, kuna "Koostis:" on stabiilselt olemas.
        composition = extract_composition(longdesc)
        if composition:
            lines.append(f"    <composition>{cdata(composition)}</composition>")
            lines.append(f"    <composition-ee>{cdata(composition)}</composition-ee>")

        # Tootja nimi: kasutame Erply "manufacturerName" (Tootja) välja, kui
        # täidetud; muidu proovime kirjeldusest "Tootja:" info automaatselt
        # välja tõmmata (stabiilselt olemas enamikul toodetel).
        manufacturer_raw = p.get("manufacturerName") or extract_manufacturer_info(longdesc) or ""
        man_name, man_address, man_email = split_manufacturer_info(manufacturer_raw)
        if man_name:
            lines.append(f"    <manufacturer-name>{cdata(man_name)}</manufacturer-name>")
        if man_address:
            lines.append(f"    <manufacturer-address>{cdata(man_address)}</manufacturer-address>")
        if man_email:
            lines.append(f"    <manufacturer-email>{cdata(man_email)}</manufacturer-email>")

        # Properties: bränd, paki kaal, ja "Tüüp" (Erply "Seeria" väljalt).
        # ID-d on vabas vormis (nemad linkivad need hiljem oma süsteemis).
        brand = p.get("brandName") or ""
        series = p.get("seriesName") or ""
        props = []
        if brand:
            props.append(("Kaubamärk", brand))
        if series:
            props.append(("Tüüp", series))
        if cat_id == "10391":  # Söögi-/joogikausid - vajavad mahtu liitrites
            volume = extract_volume_liters(name, longdesc)
            if volume:
                props.append(("Maht", f"{volume} l"))
        if weight:
            props.append(("Paki kaal", str(weight)))
        species = _PET_FOOD_CATEGORIES.get(cat_id)
        if species:
            props.append(("Looma vanus", derive_animal_age(name)))
            props.append(("Eriomadus", derive_special_feature(name, species)))
        if props:
            lines.append("    <properties>")
            for prop_id, prop_value in props:
                lines.append("      <property>")
                lines.append(f"        <id>{cdata(prop_id)}</id>")
                lines.append("        <values>")
                lines.append(f"          <value>{cdata(prop_value)}</value>")
                lines.append("        </values>")
                lines.append("      </property>")
            lines.append("    </properties>")

        lines.append("    <colours>")
        lines.append("      <colour>")
        lines.append("        <images>")
        for url in images:
            lines.append(f"          <image><url>{url}</url></image>")
        lines.append("        </images>")
        lines.append("        <modifications>")
        lines.append("          <modification>")

        lines.append(f"            <weight>{weight}</weight>")
        # Pikkus/laius/kõrgus: Erplys on need väljad millimeetrites (nt
        # "Pikkus mm"), Kaup24 nõuab meetrites - teisendame /1000.
        # Kui väärtus puudub/on 0 (pole veel sisestatud), kasutame ajutist
        # hinnangulist vaikeväärtust (10cm), kuni Erplysse on sisestatud.
        def _dim_meters(field_name):
            try:
                val_mm = float(p.get(field_name) or 0)
            except (TypeError, ValueError):
                val_mm = 0
            return val_mm / 1000 if val_mm > 0 else 0.1

        lines.append(f"            <length>{_dim_meters('length')}</length>")
        lines.append(f"            <height>{_dim_meters('height')}</height>")
        lines.append(f"            <width>{_dim_meters('width')}</width>")

        lines.append("            <attributes>")
        lines.append("              <barcodes>")
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
    print(f"\nVälja jäetud puuduliku andmestiku tõttu (range täielikkuse nõue):")
    print(f"  Kaal puudub/0: {final_stats.get('zero_weight_count', 0)}")
    print(f"  Sobiv pilt (JPG/PNG) puudub: {final_stats.get('no_image_count', 0)}")
    print(f"  Kirjeldus tühi: {final_stats.get('empty_desc_count', 0)}")

    skipped_old_batch = final_stats.get("skipped_old_batch", [])
    if skipped_old_batch:
        print(
            f"\nINFO: {len(skipped_old_batch)} toodet jäeti XML-ist välja, kuna "
            f"tegemist on uuema partiiga (sama EAN, suurema tähega) - "
            f"vanim partii (mis aegub esimesena, FIFO) on juba kaasas."
        )
        for n, raw_ean in sorted(set(skipped_old_batch))[:15]:
            print(f"  - {n} (EAN: {raw_ean})")

    skipped_no_ean = final_stats.get("skipped_no_ean", [])
    if skipped_no_ean:
        print(
            f"\nHOIATUS: {len(skipped_no_ean)} aktiivset/nähtavat toodet jäid "
            f"XML-ist välja, kuna neil puudub kehtiv EAN-kood (11-13 numbrit) Erplys."
        )
        print("Näited (kuni 15):")
        for n in sorted(set(skipped_no_ean))[:15]:
            print(f"  - {n}")
        print(
            "Lisage neile Erplys korrektne EAN (\"EAN kood\" väli), et need "
            "järgmisel käivitamisel automaatselt kaasa läheksid."
        )

    skipped_no_category = final_stats.get("skipped_no_category", [])
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

    build_xml(products, stock_map, OUTPUT_FILE, session_key=session_key)


if __name__ == "__main__":
    main()
