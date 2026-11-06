#!/usr/bin/env python3
import argparse
from collections import Counter, defaultdict, deque
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from pathlib import Path
from urllib.parse import urljoin, urlsplit
from zoneinfo import ZoneInfo
import fcntl
import json
import os
import re
import signal
import subprocess
import threading
import time
import unicodedata
from xml.etree import ElementTree as ET

from lxml import etree, html
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / "dont touch" / "duzenlenmis_veri.xlsx"
OUTPUT = ROOT / "defacto_tam_cekim"
BASE = "https://www.defacto.com.tr"
HOST = urlsplit(BASE).hostname
SKU_FORMAT = re.compile(r"([A-Z0-9-]+)\.([A-Z0-9]+)")
FIT = {"Küçük Geldi": "kucuk_geldi", "Tam Oldu": "tam_oldu", "Büyük Geldi": "buyuk_geldi"}
FIELDS = {"Modelin Ölçüleri": "manken_olculeri", "Kumaş Kalınlık Bilgisi": "kumas_kalinligi",
          "Paça Detay": "paca_detayi", "Menşei": "mensei"}
WILDCARD_RULES = []
request_lock = threading.Lock()
stop = threading.Event()
next_request = 0.0


class SitePaused(Exception):
    pass


def now():
    return datetime.now(ZoneInfo("Europe/Istanbul")).isoformat(timespec="seconds")


def read_json(path):
    return json.loads(path.read_bytes())


def save_json(path, obj):
    temp = path.with_suffix(".tmp")
    try:
        temp.write_text(json.dumps(obj, ensure_ascii=False, indent=2, allow_nan=False) + "\n", encoding="utf-8")
        temp.replace(path)
    finally:
        temp.unlink(missing_ok=True)


def host_path(value):
    if not isinstance(value, str):
        return None
    url = urlsplit(value)
    if url.scheme not in {"http", "https"} or url.hostname not in {
        "defacto.com.tr", "www.defacto.com.tr", "defacto.com", "www.defacto.com",
        "defacto.com.eg", "www.defacto.com.eg",
    }:
        return None
    slug = url.path.rstrip("/").rsplit("/", 1)[-1]
    return BASE + "/" + slug if re.search(r"-\d+$", slug) else None


def page_id(url):
    match = re.search(r"-(\d+)$", urlsplit(url).path.rstrip("/"))
    return match[1] if match else None


def permitted(url):
    parsed = urlsplit(url)
    if parsed.scheme != "https" or parsed.hostname != HOST or parsed.query:
        return False
    hits = [(length, allow) for pattern, length, allow in WILDCARD_RULES if pattern.match(parsed.path)]
    return max(hits, default=(0, True))[1]


def fetch(url, delay, robots=False):
    global next_request
    for _ in range(5):
        if not robots and not permitted(url):
            return 0, url, b""
        with request_lock:
            wait = max(0.0, next_request - time.monotonic())
            if stop.wait(wait):
                raise SitePaused("Durduruldu.")
            next_request = time.monotonic() + delay
        r = subprocess.run(
            ["/usr/bin/curl", "-4", "-sS", "--compressed", "--max-time", "25",
             "--max-filesize", "90000000",
             "-w", "\n%{http_code}\t%{redirect_url}", url],
            capture_output=True,
        )
        if r.returncode or b"\n" not in r.stdout:
            return 0, url, b""
        body, trailer = r.stdout.rsplit(b"\n", 1)
        status_text, location = trailer.decode().split("\t", 1)
        status = int(status_text)
        if status in (403, 429):
            stop.set()
            raise SitePaused(f"{url} HTTP {status} döndürdü. Site istekleri durduruldu.")
        if status not in (301, 302, 303, 307, 308) or not location:
            return status, url, body
        url = urljoin(url, location)
        robots = False
    return 0, url, b""


def load_robots(delay):
    status, _, body = fetch(BASE + "/robots.txt", delay, robots=True)
    if status != 200:
        raise SitePaused(f"robots.txt okunamadı (HTTP {status}).")
    active = False
    for line in body.decode(errors="replace").splitlines():
        field, sep, value = line.partition("#")[0].partition(":")
        if not sep:
            continue
        field, value = field.strip().lower(), value.strip()
        if field == "user-agent":
            active = value == "*"
        elif active and field in {"allow", "disallow"} and value:
            suffix = "$" if value.endswith("$") else ""
            rule = value.removesuffix("$")
            regex = re.escape(rule).replace(r"\*", ".*") + suffix
            WILDCARD_RULES.append((re.compile(regex), len(rule.replace("*", "")), field == "allow"))


def load_excel():
    book = load_workbook(SOURCE, read_only=True, data_only=True)
    try:
        rows = book["Sayfa1"].iter_rows(min_row=15, max_col=18, values_only=True)
        headers = next(rows)
        assert headers[0] == "Product Option Code" and headers[16:] == ("Product Page URL", "Model Page URL")
        result = defaultdict(lambda: defaultdict(list))
        for row_number, cells in enumerate(rows, 16):
            if not any(x is not None for x in cells):
                continue
            match = SKU_FORMAT.fullmatch(str(cells[0]).strip().upper())
            if not match:
                raise ValueError(f"Excel {row_number}. satırdaki ürün kodu okunamadı.")
            model, color = match.groups()
            record = {key: value.isoformat() if isinstance(value, datetime) else value for key, value in zip(headers, cells)}
            result[model][color].append({"excel_satiri": row_number, **record})
        return result
    finally:
        book.close()


def sitemap_candidates(models, delay):
    cache = OUTPUT / "site_haritasi.json"
    if cache.exists():
        existing = read_json(cache)
        if existing["tarih"] == now()[:10]:
            return existing
    listing = BASE + "/sharedxml/tr-tr/sitemapindex.xml"
    status, _, body = fetch(listing, delay)
    locations = []
    if status == 200:
        locations = [node.text.strip() for node in ET.fromstring(body).findall(".//{*}loc")
                     if node.text and re.search(r"/product(?:-\d+)?\.xml$", node.text.strip())]
    candidates = defaultdict(dict)
    failures = []
    if not locations:
        failures.append({"url": listing, "http": status})
    for location in locations:
        code, _, data = fetch(location, delay)
        if code != 200:
            failures.append({"url": location, "http": code})
            continue
        for entry in ET.fromstring(data):
            loc = entry.find("{*}loc")
            page = host_path(loc.text.strip()) if loc is not None and loc.text else None
            if not page:
                continue
            identities = []
            for image in entry.findall("{*}image/{*}loc"):
                parts = Path(urlsplit(image.text or "").path).name.split("_")
                if len(parts) >= 4 and parts[0] in models and parts[2].isalnum():
                    identity = (parts[0], parts[2])
                    if identity not in identities:
                        identities.append(identity)
            if len(identities) == 1:
                model, color = identities[0]
                candidates[model].setdefault(page, model + color)
    result = {"tarih": now()[:10], "dosyalar": locations, "hatalar": failures, "modeller": candidates}
    save_json(cache, result)
    return result


def css_class(value):
    return "contains(concat(' ',normalize-space(@class),' '),' " + value + " ')"


def first(root, selector):
    return next(iter(root.xpath(selector)), None)


def value(node):
    return " ".join(" ".join(node.itertext()).split()) if node is not None else ""


def own(node):
    return " ".join(" ".join(node.xpath("./text()")).split())


def truncate(text, words):
    tokens = text.split()
    return " ".join(tokens[:words]) + ("…" if len(tokens) > words else "")


def clean_label(label):
    if label in FIELDS:
        return FIELDS[label]
    ascii_text = unicodedata.normalize("NFKD", label.lower().replace("ı", "i")).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]+", "_", ascii_text).strip("_")


def first_class(root, name):
    return first(root, f".//*[{css_class(name)}]")


def product_structured(page):
    for script in page.xpath('//script[@type="application/ld+json"]/text()'):
        try:
            obj = json.loads(script)
        except ValueError:
            continue
        if isinstance(obj, dict) and obj.get("@type") == "Product":
            return obj
    return {}


def numeric(value_):
    try:
        return float(str(value_).replace(",", "."))
    except (ValueError, TypeError):
        return None


def count(value_):
    number = numeric(value_)
    return int(number) if number is not None and number.is_integer() else None


def review_date(raw):
    if isinstance(raw, str):
        for format_ in ("%Y-%m-%d", "%d.%m.%Y"):
            try:
                return datetime.strptime(raw, format_).date().isoformat()
            except ValueError:
                continue
    return None


def read_summary(page, structured):
    agg = structured.get("aggregateRating") or {}
    if not isinstance(agg, dict):
        agg = {}
    distribution = {}
    for pattern in page.xpath(f'//*[@id="section__reviews"]//*[{css_class("body-size__pattern")}]'):
        label = value(first_class(pattern, "body-size__pattern--text"))
        ratio = value(first_class(pattern, "body-size__pattern--count")).lstrip("% ")
        if label in FIT and (number := numeric(ratio)) is not None:
            distribution[FIT[label]] = number
    reviews = structured.get("review") or []
    if isinstance(reviews, dict):
        reviews = [reviews]
    samples = []
    for review in reviews if isinstance(reviews, list) else ():
        if not isinstance(review, dict):
            continue
        message = review.get("reviewBody") or ""
        if not isinstance(message, str) or not message.strip():
            continue
        score = review.get("reviewRating") or {}
        samples.append({"tarih": review_date(review.get("datePublished")),
                        "puan": numeric(score.get("ratingValue")) if isinstance(score, dict) else None,
                        "metin": truncate(message, 20)})
    return {"ortalama_puan": numeric(agg.get("ratingValue")), "degerlendirme_sayisi": count(agg.get("ratingCount")),
            "yorum_sayisi": count(agg.get("reviewCount")),
            "kalip_dagilimi_yuzde": distribution if any(distribution.values()) else None,
            "yorum_ornekleri": samples[:10]}


def read_price(page, checked):
    panel = first_class(page, "product-detail__price")
    if panel is None:
        return None
    result = {}
    for tag in panel.xpath(f'.//*[{css_class("base-price")}]'):
        text = value(tag)
        found = re.search(r"([0-9][0-9.,]*)\s*TL\b", text)
        if not found:
            continue
        amount = found[1]
        if "," in amount and "." in amount:
            divider = max(amount.rfind(","), amount.rfind("."))
            amount = re.sub(r"[.,]", "", amount[:divider]) + "." + amount[divider+1:]
        price = numeric(amount)
        if price is None:
            continue
        name = ("uzeri_cizili_fiyat" if "lined-base-price" in tag.get("class", "").split()
                else "sepette_fiyat" if "sepette" in text.lower() else "gosterilen_fiyat")
        result.setdefault(name, price)
    if not result:
        return None
    if label := value(first_class(panel, "product-detail-view__price-campaing")):
        result["etiket"] = label
    return {"para_birimi": "TRY", "kontrol_tarihi": checked[:10], **result}


def picture(node, url):
    raw = node.get("data-src") or node.get("src")
    if not raw or "placeholder" in raw.lower():
        return None
    link = urljoin(url, raw)
    parsed = urlsplit(link)
    return link if parsed.scheme == "https" and parsed.hostname == "dfcdn.defacto.com.tr" else None


def read_page(url, model, checked, delay):
    status, final, body = fetch(url, delay)
    if status != 200 or not page_id(final):
        return None, {}, None, status, "Ürün sayfası açılmadı"
    soup = html.fromstring(body, parser=html.HTMLParser(encoding="utf-8"))
    match = re.fullmatch(r"(.+?)\s*/\s*Ürün Kodu\s*:\s*([A-Z0-9-]+)", value(first_class(soup, "product-detail__sku")))
    if not match or not match[2].startswith(model) or match[2] == model:
        return None, {}, None, status, "Sayfada bu modelin ürün kodu bulunamadı"
    marker = re.search(rb"\bwindow\.VIEWED_PRODUCT\s*=\s*", body)
    if marker:
        try:
            viewed, _ = json.JSONDecoder().raw_decode(body[marker.end():].decode(errors="replace"))
        except ValueError:
            viewed = None
        if isinstance(viewed, dict) and (published_model := viewed.get("ProductCode")) and published_model != model:
            return None, {}, None, status, "Sayfada farklı bir model kodu yayımlandı"
    long_sku = match[2]
    color = long_sku[len(model):]
    canonical = first(soup, '//link[@rel="canonical"]')
    canonical = host_path(canonical.get("href")) if canonical is not None else None
    if canonical and page_id(canonical) == page_id(final):
        final = canonical
    structured = product_structured(soup)
    refs = structured.get("isRelatedTo") or []
    by_id = {page_id(v): host_path(v) for v in refs if isinstance(v, str) and urlsplit(v).hostname == HOST}
    selected = None
    siblings = {}
    for swatch in soup.xpath(f'//*[{css_class("product-detail__color-swatch")}]'):
        index = swatch.get("data-variantindex", "")
        if index == page_id(final):
            selected = swatch
        elif index.isdigit():
            siblings[index] = by_id.get(index) or final.rsplit("-", 1)[0] + "-" + index
    hex_match = re.search(r"background-color\s*:\s*(#[\da-f]{6})(?:;|$)", selected.get("style", ""), re.I) if selected is not None else None
    gallery = []
    query = " | ".join(f'//*[{css_class(cls)}]//img' for cls in (
        "product-first-media", "product-detail__secondary-media", "product-detail__extra-image-wrapper"))
    for node in soup.xpath(query):
        link = picture(node, final)
        if not link:
            continue
        parts = Path(urlsplit(link).path).name.split("_")
        if len(parts) >= 4 and (parts[0], parts[2]) == (model, color) and link not in gallery:
            gallery.append(link)
    info, about = {}, []
    for box in soup.xpath(f'//*[@id="section__productinfo"]//*[{css_class("sideMenu__box")}]'):
        is_about = value(first_class(box, "sideMenu__subtitle")) == "Ürün Hakkında"
        for entry in box.xpath(".//li"):
            line = own(entry)
            if not line:
                continue
            if is_about:
                if not line.startswith("Bu ürünü evde deneyin"):
                    about.append(truncate(line, 20))
                continue
            if line.startswith("Modelin Ölçüleri"):
                info["Modelin Ölçüleri"] = line.removeprefix("Modelin Ölçüleri").lstrip(" :")
            else:
                label, sep, detail = line.partition(":")
                if sep and detail.lstrip(" :"):
                    info[label.strip()] = detail.lstrip(" :")
    if info.get("Ürün Kodu", long_sku) != long_sku:
        return None, {}, None, status, "Ürün bilgileri farklı koda ait"
    sizes = structured.get("size") or []
    if isinstance(sizes, str):
        sizes = [sizes]
    if not isinstance(sizes, list):
        sizes = []
    rules = []
    for li in soup.xpath('//*[@id="section__productwashing"]//li'):
        label = own(li)
        if label:
            icon = first(li, "./img")
            address = picture(icon, final) if icon is not None else None
            rules.append({"talimat": label, **({"simge_url": address} if address else {})})
    record = {
        "urun_renk_kodu": long_sku, "renk_kodu": color, "renk_adi": match[1], "renk_hex": hex_match[1].upper() if hex_match else None,
        "urun_adi": value(first_class(soup, "product-detail__main-title")) or value(first_class(soup, "product-detail__info-header-title")),
        "kisa_aciklama": truncate(value(first_class(soup, "product-detail__description-text")), 20),
        "urun_sayfasi_url": final, "galeri_url_listesi": gallery,
        "bedenler": {"web": list(dict.fromkeys(str(x) for x in sizes if isinstance(x, (str, int)))), "excel": []},
        "web_beden_kaynagi": "web_urun_sayfasi",
        "urun_bilgileri": {clean_label(k): v for k, v in info.items() if k not in {"Ürün Kodu", "Ürün Grubu", "Renk"}},
        "urun_grubu": info.get("Ürün Grubu"), "urun_hakkinda": about,
        "bakim_talimati": {"talimatlar": rules} if rules else None, "fiyat": read_price(soup, checked),
    }
    return record, siblings, read_summary(soup, structured), status, None


def image_urls(rows, model, color):
    result = []
    for row in rows:
        address = row.get("WebImageUrl")
        if not isinstance(address, str):
            continue
        url = urlsplit(address)
        parts = Path(url.path).name.split("_")
        if url.scheme == "https" and len(parts) >= 4 and (parts[0], parts[2]) == (model, color) and address not in result:
            result.append(address)
    return result


def seed_urls(model, rows_by_color, sitemap, pilot):
    result = {}
    if pilot:
        for row in pilot.get("renkler", []):
            if url := host_path(row.get("urun_sayfasi_url")):
                result.setdefault(page_id(url), url)
    for row_set in rows_by_color.values():
        for row in row_set:
            if url := host_path(row.get("Product Page URL")):
                result.setdefault(page_id(url), url)
    for row_set in rows_by_color.values():
        for row in row_set:
            if url := host_path(row.get("Model Page URL")):
                result.setdefault(page_id(url), url)
    for url in sitemap.get(model, {}):
        result.setdefault(page_id(url), url)
    return result


def attach_excel(variants, model, colors):
    for color, rows in colors.items():
        first_row = rows[0]
        item = variants.setdefault(color, {
            "urun_renk_kodu": model + color, "renk_kodu": color, "renk_adi": first_row.get("Product Color Name"),
            "renk_hex": None, "urun_sayfasi_url": None, "galeri_url_listesi": [], "bedenler": {"web": [], "excel": []},
        })
        item["bedenler"]["excel"] = list(dict.fromkeys(str(r["SizeDefinition"]) for r in rows if r["SizeDefinition"] is not None))
        item["excel_satirlari"] = rows
        item["excel_gorsel_url_listesi"] = image_urls(rows, model, color)
    for item in variants.values():
        item.setdefault("excel_satirlari", [])
        item.setdefault("excel_gorsel_url_listesi", [])


def prepare_baseline(target, models, pilots):
    for model, colors in models.items():
        file = target / (model + ".json")
        if file.exists():
            continue
        pilot = pilots.get(model, {})
        variants = {item["renk_kodu"]: item for item in pilot.get("renkler", [])}
        attach_excel(variants, model, colors)
        save_json(file, {
            "model_kodu": model, "urun_grubu": pilot.get("urun_grubu"),
            "degerlendirmeler": pilot.get("degerlendirmeler"),
            "renkler": list(variants.values()), "excel_kaynagi": {"dosya": SOURCE.name, "sayfa": "Sayfa1"},
            "web_kontrol": {"durum": "bekliyor", "tarih": None, "denenen_sayfa": 0,
                            "dogrulanan_renk": sum(bool(item["urun_sayfasi_url"]) for item in variants.values())},
        })


def collect(model, colors, sitemap, pilot, delay):
    checked = now()
    queue = seed_urls(model, colors, sitemap, pilot)
    known = set()
    variants = {}
    notes = []
    shared = None
    primary_group = pilot.get("urun_grubu") if pilot else None
    while queue:
        index = next(iter(queue))
        url = queue.pop(index)
        if index in known:
            continue
        known.add(index)
        try:
            item, siblings, summary, status, reason = read_page(url, model, checked, delay)
        except (ValueError, TypeError, KeyError, etree.ParserError) as error:
            notes.append({"url": url, "http": None, "durum": str(error)})
            continue
        if not item:
            notes.append({"url": url, "http": status or None, "durum": reason})
            continue
        color = item["renk_kodu"]
        group = item.pop("urun_grubu")
        if not primary_group:
            primary_group = group
        elif group and group != primary_group:
            item["urun_grubu"] = group
        if color in variants:
            previous = variants[color]
            previous["galeri_url_listesi"] = list(dict.fromkeys(previous["galeri_url_listesi"] + item["galeri_url_listesi"]))
            if item["urun_sayfasi_url"] != previous["urun_sayfasi_url"]:
                previous.setdefault("diger_urun_sayfalari", []).append(item["urun_sayfasi_url"])
        else:
            variants[color] = item
        if not shared or (shared["ortalama_puan"] is None and summary["ortalama_puan"] is not None):
            shared = summary
        for sibling, address in siblings.items():
            if sibling not in known:
                queue.setdefault(sibling, address)
    if pilot:
        pilot_summary = pilot.get("degerlendirmeler") or {}
        if not shared:
            shared = pilot_summary.copy()
        elif len(pilot_summary.get("yorum_ornekleri", [])) > len(shared["yorum_ornekleri"]):
            shared["yorum_ornekleri"] = pilot_summary["yorum_ornekleri"][:10]
    if pilot:
        for item in pilot.get("renkler", []):
            variants.setdefault(item["renk_kodu"], item)
    attach_excel(variants, model, colors)
    web = sum(bool(item["urun_sayfasi_url"]) for item in variants.values())
    status = "web_dogrulandi" if web else "web_sayfasi_bulunamadi" if known else "aday_url_yok"
    if any(note["http"] is None or note["http"] >= 500 for note in notes):
        status = "tekrar_kontrol"
    result = {
        "model_kodu": model, "urun_grubu": primary_group, "degerlendirmeler": shared,
        "renkler": list(variants.values()), "excel_kaynagi": {"dosya": SOURCE.name, "sayfa": "Sayfa1"},
        "web_kontrol": {"durum": status, "tarih": checked, "denenen_sayfa": len(known), "dogrulanan_renk": web},
    }
    if notes:
        result["web_sayfasi_kontrol_notlari"] = notes
    return result


def run(delay, retry, limit, excel_only):
    target = OUTPUT / "urunler"
    target.mkdir(exist_ok=True)
    models = load_excel()
    pilots = {p.stem: read_json(p) for p in (ROOT / "defacto_urunler").glob("*.json")}
    prepare_baseline(target, models, pilots)
    progress = {
        "durum": "hazirlaniyor", "toplam_model": len(models),
        "excel_urun_renk_kodu": sum(map(len, models.values())),
        "excel_satiri": sum(len(rows) for model in models.values() for rows in model.values()),
        "json_dosyasi": len(models),
    }
    completed = {}
    for p in target.glob("*.json"):
        doc = read_json(p)
        if (p.stem in models and doc["web_kontrol"]["durum"] != "bekliyor"
                and (not retry or doc["web_kontrol"]["durum"] == "web_dogrulandi")):
            completed[p.stem] = doc["web_kontrol"]
    counters = Counter()
    for item in completed.values():
        counters[item["durum"]] += 1

    def update(state):
        progress.update({"durum": state, "guncellendi": now(), "web_kontrol_edilen_model": len(completed),
                         "web_renk_kaydi": sum(r["dogrulanan_renk"] for r in completed.values()),
                         "model_sayimlari": dict(counters)})
        save_json(OUTPUT / "durum.json", progress)

    if excel_only:
        update("excel_hazir")
        return
    update("hazirlaniyor")
    try:
        load_robots(delay)
        mapping = sitemap_candidates(models, delay)
        sitemap = mapping["modeller"]
        progress["site_haritasi_modeli"] = len(sitemap)
        progress["site_haritasi_hatalari"] = mapping["hatalar"]
        order = sorted(models, key=lambda model: (model not in pilots, model not in sitemap,
                       not bool(seed_urls(model, models[model], sitemap, pilots.get(model))), model))
        remaining = [model for model in order if model not in completed]
        remaining = deque(remaining[:limit] if limit else remaining)
        update("calisiyor")
        with ThreadPoolExecutor(max_workers=2) as pool:
            # Keep only a small set of futures queued so a pause stops requests promptly.
            queued = {}
            while remaining or queued:
                while remaining and len(queued) < 2 and not stop.is_set():
                    model = remaining.popleft()
                    queued[model] = pool.submit(collect, model, models[model], sitemap, pilots.get(model), delay)
                if not queued:
                    break
                model = next(iter(queued))
                try:
                    doc = queued.pop(model).result()
                except SitePaused:
                    stop.set()
                    raise
                save_json(target / (model + ".json"), doc)
                check = doc["web_kontrol"]
                completed[model] = check
                counters[check["durum"]] += 1
                update("calisiyor")
                print(now(), model, check["durum"], check["dogrulanan_renk"],
                      f"{len(completed)}/{len(models)}", flush=True)
    except SitePaused as error:
        progress["bekleme_nedeni"] = str(error)
        update("durdu")
        print(str(error), flush=True)
        return
    except Exception as error:
        stop.set()
        progress["bekleme_nedeni"] = str(error)
        update("hata")
        raise
    update("bitti" if len(completed) == len(models) else "ornek_tamamlandi")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--aralik", type=float, default=0.7, help="Ürün sayfası istekleri arasında saniye")
    parser.add_argument("--yeniden-dene", action="store_true", help="Web eşleşmesi eksik dosyaları tekrar sorgula")
    parser.add_argument("--limit", type=int, default=0, help="Bu çalıştırmada işlenecek model sayısı")
    parser.add_argument("--excel-hazirla", action="store_true", help="Sadece Excel tabanlı model dosyalarını hazırla")
    args = parser.parse_args()
    OUTPUT.mkdir(exist_ok=True)
    signal.signal(signal.SIGTERM, lambda *_: stop.set())
    signal.signal(signal.SIGINT, lambda *_: stop.set())
    with (OUTPUT / "cekim.lock").open("a") as lock:
        fcntl.flock(lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
        pid_file = OUTPUT / "cekim.pid"
        pid_file.write_text(f"{os.getpid()}\n", encoding="utf-8")
        try:
            run(max(0.5, args.aralik), args.yeniden_dene, args.limit, args.excel_hazirla)
        finally:
            pid_file.unlink(missing_ok=True)


if __name__ == "__main__":
    main()
