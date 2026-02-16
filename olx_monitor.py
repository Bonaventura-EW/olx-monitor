#!/usr/bin/env python3
"""
OLX Monitor — monitoring ogłoszeń z datą publikacji, wiekiem i historią cen.

Generuje:
  olx_monitoring.xlsx     — pełna tabela z każdym skanem
  price_history.json      — historia cen dla dashboardu HTML

Kolumny Excela:
  Data skanu | Profil | Tytuł | Cena (zł) | Data publikacji | Dni od publikacji | URL | ID
"""

import requests, re, json, os, time
from bs4 import BeautifulSoup
from datetime import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

CONFIG_FILE        = "config.json"
EXCEL_FILE         = "olx_monitoring.xlsx"
PRICE_HISTORY_FILE = "price_history.json"

HEADERS = {
    "User-Agent":      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36",
    "Accept-Language": "pl-PL,pl;q=0.9",
    "Accept":          "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}

def parse_created(html):
    idx = html.find("createdTime")
    if idx < 0:
        return None, None
    snippet = html[idx:idx + 80]
    m = re.search(r"(\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}[+-]\d{2}:\d{2})", snippet)
    if not m:
        return None, None
    dt_str = m.group(1)
    try:
        dt  = datetime.fromisoformat(dt_str)
        now = datetime.now(tz=dt.tzinfo)
        days = max(0, (now - dt).days)
        return dt.strftime("%d.%m.%Y"), days
    except Exception:
        return None, None

def today_label():
    months = ["sty","lut","mar","kwi","maj","cze","lip","sie","wrz","paź","lis","gru"]
    n = datetime.now()
    return f"{n.day} {months[n.month - 1]}"

def scrape_profile(profile_name, profile_url):
    print(f"  [{profile_name}] {profile_url}")
    try:
        r = requests.get(profile_url, headers=HEADERS, timeout=15)
        r.raise_for_status()
    except Exception as e:
        print(f"    ⚠ Błąd pobierania profilu: {e}")
        return []
    soup = BeautifulSoup(r.text, "html.parser")
    listings, seen = [], set()
    for a in soup.find_all("a", href=lambda h: h and "/d/oferta/" in h):
        parent = a.parent
        if not parent:
            continue
        if "css-1pktvhb" not in " ".join(parent.get("class", [])):
            continue
        href = re.sub(r"\?.*", "", a.get("href", ""))
        if href in seen:
            continue
        seen.add(href)
        title = a.get_text(strip=True)
        if not title or len(title) < 5:
            continue
        card_text = parent.get_text(" ", strip=True)
        price_m   = re.search(r"([\d\s]{2,8})zł", card_text)
        price     = int(re.sub(r"[^\d]", "", price_m.group(1))) if price_m and re.sub(r"[^\d]", "", price_m.group(1)) else 0
        full_url  = ("https://www.olx.pl" + href) if href.startswith("/") else href
        id_m      = re.search(r"/d/oferta/([^/?]+)", href)
        listing_id = id_m.group(1) if id_m else href.replace("/", "_")
        listings.append({"id": listing_id, "profile": profile_name, "title": title, "price": price, "url": full_url, "created": None, "days_old": None})
    print(f"    → {len(listings)} ogłoszeń")
    return listings

def fetch_dates(listings, delay=1.2):
    print(f"\n  Pobieranie dat publikacji ({len(listings)} ogłoszeń, ~{len(listings)*delay:.0f}s)...")
    for i, l in enumerate(listings, 1):
        try:
            r = requests.get(l["url"], headers=HEADERS, timeout=12)
            created, days = parse_created(r.text)
            l["created"]  = created
            l["days_old"] = days
            status = f"{created}  ({days} dni)" if created else "brak daty"
        except Exception as e:
            l["created"]  = None
            l["days_old"] = None
            status = f"błąd: {e}"
        print(f"    [{i:2}/{len(listings)}] {l['title'][:50]:<50} {status}")
        time.sleep(delay)
    return listings

def update_price_history(listings):
    today = today_label()
    history = {}
    if os.path.exists(PRICE_HISTORY_FILE):
        try:
            with open(PRICE_HISTORY_FILE, "r", encoding="utf-8") as f:
                history = json.load(f)
        except Exception:
            pass
    for l in listings:
        lid = l["id"]
        if lid not in history:
            history[lid] = {"title": l["title"], "profile": l["profile"], "created": l["created"] or "", "prices": []}
        else:
            if not history[lid].get("created") and l.get("created"):
                history[lid]["created"] = l["created"]
        prices = history[lid]["prices"]
        entry  = next((e for e in prices if e["date"] == today), None)
        if entry:
            entry["price"] = l["price"]
        elif l["price"] > 0:
            prices.append({"date": today, "price": l["price"]})
    with open(PRICE_HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(history, f, ensure_ascii=False, indent=2)
    print(f"  → {PRICE_HISTORY_FILE}: {len(history)} ogłoszeń")

COLUMNS = [
    ("Data skanu", 20), ("Profil", 16), ("Tytuł", 54), ("Cena (zł)", 12),
    ("Data publikacji", 16), ("Dni od publikacji", 18), ("URL", 60), ("ID ogłoszenia", 44),
]

def save_to_excel(listings):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    thin = Side(style="thin", color="2a2a38")
    border = Border(bottom=thin)
    if os.path.exists(EXCEL_FILE):
        wb = load_workbook(EXCEL_FILE)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Historia"
        ws.append([col for col, _ in COLUMNS])
        hfill  = PatternFill("solid", fgColor="1a1a2e")
        hfont  = Font(color="e8ff47", bold=True, size=10)
        halign = Alignment(horizontal="center", vertical="center", wrap_text=False)
        for col_idx, cell in enumerate(ws[1], 1):
            cell.fill = hfill
            cell.font = hfont
            cell.alignment = halign
            cell.border = border
            ws.column_dimensions[get_column_letter(col_idx)].width = COLUMNS[col_idx-1][1]
        ws.freeze_panes = "A2"
        ws.row_dimensions[1].height = 20
    for l in listings:
        days = l.get("days_old")
        ws.append([now, l["profile"], l["title"], l["price"], l["created"] or "", days if days is not None else "", l["url"], l["id"]])
        row = ws.max_row
        ws.cell(row, 1).alignment = Alignment(horizontal="left")
        ws.cell(row, 3).alignment = Alignment(horizontal="left")
        ws.cell(row, 4).alignment = Alignment(horizontal="center")
        ws.cell(row, 5).alignment = Alignment(horizontal="center")
        ws.cell(row, 6).alignment = Alignment(horizontal="center")
        if days is not None:
            cell = ws.cell(row, 6)
            if days <= 3:    cell.font = Font(color="47ffa0", bold=True, size=10)
            elif days <= 14: cell.font = Font(color="e8ff47", bold=False, size=10)
            elif days > 60:  cell.font = Font(color="ff6b6b", bold=False, size=10)
    wb.save(EXCEL_FILE)
    print(f"  → {EXCEL_FILE}: +{len(listings)} wierszy (łącznie {ws.max_row - 1})")

def main():
    print("=" * 60)
    print(f"OLX Monitor — {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print("=" * 60)
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        config = json.load(f)
    print("\n[1/3] Scraping profili OLX...")
    all_listings = []
    for p in config.get("profiles", []):
        all_listings.extend(scrape_profile(p["name"], p["url"]))
    if not all_listings:
        print("⚠ Brak ogłoszeń. Koniec.")
        return
    print(f"\nRazem: {len(all_listings)} ogłoszeń")
    print("\n[2/3] Pobieranie dat publikacji z OLX...")
    all_listings = fetch_dates(all_listings, delay=1.2)
    print("\n[3/3] Zapisywanie...")
    update_price_history(all_listings)
    save_to_excel(all_listings)
    with_date = [l for l in all_listings if l["created"]]
    no_date   = [l for l in all_listings if not l["created"]]
    print(f"\n✓ Gotowe!")
    print(f"  Daty publikacji znalezione: {len(with_date)}/{len(all_listings)}")
    if no_date:
        print(f"  Bez daty: {[l['title'][:40] for l in no_date]}")
    print("=" * 60)

if __name__ == "__main__":
    main()
