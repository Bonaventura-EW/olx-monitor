import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime, date
import os
import re
import time
import json

# Opcjonalna synchronizacja z Google Sheets
ENABLE_SHEETS_SYNC = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "") != ""

# ─── KONFIGURACJA ────────────────────────────────────────────────────────────

PROFILES = [
    {
        "name": "wszystkie-lublin",
        "url": "https://www.olx.pl/nieruchomosci/stancje-pokoje/lublin/",
        "type": "category"
    },
    {
        "name": "artymiuk",
        "url": "https://www.olx.pl/oferty/uzytkownik/BAm3j/",
        "type": "user"
    },
    {
        "name": "poqui",
        "url": "https://www.olx.pl/oferty/uzytkownik/p8eWV/",
        "type": "user"
    },
    {
        "name": "stylowepokoje",
        "url": "https://www.olx.pl/oferty/uzytkownik/3cxbz/",
        "type": "user"
    },
    {
        "name": "villahome",
        "url": "https://www.olx.pl/oferty/uzytkownik/1n7fOJ/",
        "type": "user"
    },
]

EXCEL_FILE = "data/olx_monitoring.xlsx"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept-Language": "pl-PL,pl;q=0.9,en-US;q=0.8",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}

# ─── KOLORY ──────────────────────────────────────────────────────────────────

COLOR_HEADER_BG   = "2C5F8A"   # ciemny niebieski – nagłówki
COLOR_HEADER_FONT = "FFFFFF"   # biały
COLOR_NEW_BG      = "C6EFCE"   # zielony – nowe ogłoszenia
COLOR_DEL_BG      = "FFC7CE"   # czerwony – usunięte
COLOR_DATE_BG     = "D9E1F2"   # jasny niebieski – wiersz daty
COLOR_ODD_ROW     = "F2F7FB"   # bardzo jasny niebieski – naprzemienne wiersze
COLOR_SUMMARY_BG  = "FFF2CC"   # żółty – zakładka PODSUMOWANIE

# ─── SCRAPING ────────────────────────────────────────────────────────────────

def get_ad_count(url: str) -> int | None:
    """Pobiera liczbę ogłoszeń z podanego URL."""
    try:
        resp = requests.get(url, headers=HEADERS, timeout=15)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.content, "html.parser")

        # Metoda 1: atrybut data-testid="total-count"
        el = soup.find(attrs={"data-testid": "total-count"})
        if el:
            digits = re.sub(r"\D", "", el.get_text())
            if digits:
                return int(digits)

        # Metoda 2: zlicz karty ogłoszeń
        cards = soup.find_all("div", {"data-cy": "l-card"})
        if cards:
            return len(cards)

        # Metoda 3: regex w h1
        h1 = soup.find("h1")
        if h1:
            m = re.search(r"(\d[\d\s]*)\s*ogłoszeń", h1.get_text())
            if m:
                return int(re.sub(r"\D", "", m.group(1)))

        # Metoda 4: meta description
        meta = soup.find("meta", {"name": "description"})
        if meta:
            m = re.search(r"(\d+)\s*ogłoszeń", meta.get("content", ""))
            if m:
                return int(m.group(1))

        return 0

    except Exception as e:
        print(f"  ⚠  Błąd przy {url}: {e}")
        return None


def get_individual_ads(url: str) -> list[dict]:
    """
    Pobiera listę ogłoszeń (id + tytuł) ze strony profilu użytkownika.
    Używana do śledzenia konkretnych ogłoszeń (nowe / usunięte).
    """
    ads = []
    try:
        resp = requests.get(url, headers=HEADERS, timeout=15)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.content, "html.parser")

        for card in soup.find_all("div", {"data-cy": "l-card"}):
            link = card.find("a", href=True)
            if not link:
                continue
            href = link["href"]
            # wyciągnij ID ogłoszenia z URL-a
            m = re.search(r"ID(\w+)\.html", href)
            ad_id = m.group(1) if m else href.split("/")[-1]
            title_el = card.find(["h3", "h4", "h6"])
            title = title_el.get_text(strip=True) if title_el else "Brak tytułu"
            ads.append({"id": ad_id, "title": title, "url": href})

    except Exception as e:
        print(f"  ⚠  Błąd przy pobieraniu ogłoszeń z {url}: {e}")

    return ads


# ─── EXCEL: POMOCNICZE ───────────────────────────────────────────────────────

def thin_border() -> Border:
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)


def header_cell(ws, row: int, col: int, value: str, width: int | None = None):
    c = ws.cell(row=row, column=col, value=value)
    c.font      = Font(bold=True, color=COLOR_HEADER_FONT, name="Arial", size=10)
    c.fill      = PatternFill("solid", start_color=COLOR_HEADER_BG)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border    = thin_border()
    if width:
        ws.column_dimensions[get_column_letter(col)].width = width


def data_cell(ws, row: int, col: int, value, bg: str | None = None,
              bold: bool = False, align: str = "center"):
    c = ws.cell(row=row, column=col, value=value)
    c.font      = Font(bold=bold, name="Arial", size=10)
    c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
    c.border    = thin_border()
    if bg:
        c.fill = PatternFill("solid", start_color=bg)
    return c


# ─── EXCEL: INICJALIZACJA / ŁADOWANIE ────────────────────────────────────────

def load_or_create_workbook() -> openpyxl.Workbook:
    if os.path.exists(EXCEL_FILE):
        return openpyxl.load_workbook(EXCEL_FILE)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)         # usuń domyślny pusty arkusz
    _create_summary_sheet(wb)
    for p in PROFILES:
        _create_profile_sheet(wb, p["name"])
    return wb


def _create_summary_sheet(wb: openpyxl.Workbook):
    ws = wb.create_sheet("PODSUMOWANIE", 0)
    ws.sheet_properties.tabColor = "2C5F8A"
    ws.row_dimensions[1].height = 30

    cols = ["Profil", "Data ostatniego sprawdzenia", "Liczba ogłoszeń",
            "Nowe (+)", "Usunięte (−)", "Status"]
    widths = [22, 26, 20, 12, 14, 14]
    for i, (col, w) in enumerate(zip(cols, widths), 1):
        header_cell(ws, 1, i, col, w)

    for i, p in enumerate(PROFILES, 2):
        ws.cell(row=i, column=1, value=p["name"])
        ws.row_dimensions[i].height = 18

    ws.freeze_panes = "A2"


def _create_profile_sheet(wb: openpyxl.Workbook, name: str):
    ws = wb.create_sheet(name)
    ws.row_dimensions[1].height = 30

    cols = ["Data", "Łączna liczba ogłoszeń", "Nowe ogłoszenia (+)",
            "Usunięte ogłoszenia (−)", "Zmiana netto", "Szczegóły nowych",
            "Szczegóły usuniętych", "Status"]
    widths = [20, 24, 20, 22, 14, 40, 40, 14]
    for i, (col, w) in enumerate(zip(cols, widths), 1):
        header_cell(ws, 1, i, col, w)

    ws.freeze_panes = "A2"


# ─── EXCEL: ZAPIS DANYCH ─────────────────────────────────────────────────────

def _get_previous_data(ws) -> dict:
    """Zwraca dane z ostatniego wiersza arkusza profilu."""
    max_row = ws.max_row
    if max_row < 2:
        return {"count": None, "ads": []}

    # Szukamy ostatniego wiersza z danymi (od końca)
    for r in range(max_row, 1, -1):
        val = ws.cell(row=r, column=2).value
        if val is not None:
            prev_count = int(val) if str(val).isdigit() else None
            # Odczytujemy zapisane ID ogłoszeń z kolumny 9 (ukryta)
            raw = ws.cell(row=r, column=9).value or ""
            prev_ids = set(raw.split("|")) if raw else set()
            return {"count": prev_count, "ids": prev_ids, "row": r}

    return {"count": None, "ids": set(), "row": 1}


def update_profile_sheet(ws, profile: dict, today_str: str,
                          total: int | None, today_ads: list[dict]):
    prev = _get_previous_data(ws)
    prev_count = prev.get("count")
    prev_ids   = prev.get("ids", set())

    today_ids  = {a["id"] for a in today_ads}
    new_ids    = today_ids - prev_ids
    del_ids    = prev_ids - today_ids

    # Nowe ogłoszenia
    new_ads = [a for a in today_ads if a["id"] in new_ids]
    new_count = len(new_ads) if today_ids else (
        max(0, (total or 0) - (prev_count or 0)) if prev_count is not None else 0
    )

    # Usunięte
    del_count = len(del_ids)

    net_change = (total or 0) - (prev_count or 0) if prev_count is not None else 0

    new_row = ws.max_row + 1
    row_bg = COLOR_ODD_ROW if new_row % 2 == 0 else None

    status = "OK" if total is not None else "BŁĄD"

    def bg(special=None):
        return special or row_bg

    data_cell(ws, new_row, 1, today_str,  bg=COLOR_DATE_BG, bold=True)
    data_cell(ws, new_row, 2, total,       bg=bg())
    data_cell(ws, new_row, 3, new_count,   bg=bg(COLOR_NEW_BG if new_count > 0 else None), bold=(new_count > 0))
    data_cell(ws, new_row, 4, del_count,   bg=bg(COLOR_DEL_BG if del_count > 0 else None), bold=(del_count > 0))
    data_cell(ws, new_row, 5, net_change,  bg=bg(), bold=True)

    new_details = "; ".join([f"{a['title'][:50]}" for a in new_ads[:10]]) or "—"
    del_details = "; ".join(list(del_ids)[:10]) or "—"

    data_cell(ws, new_row, 6, new_details, bg=bg(COLOR_NEW_BG if new_count > 0 else None), align="left")
    data_cell(ws, new_row, 7, del_details, bg=bg(COLOR_DEL_BG if del_count > 0 else None), align="left")
    data_cell(ws, new_row, 8, status,      bg=bg("C6EFCE" if status == "OK" else "FFC7CE"))

    # Kolumna 9 – ukryta, przechowuje ID ogłoszeń do porównania następnego dnia
    ws.cell(row=new_row, column=9, value="|".join(today_ids))
    ws.column_dimensions["I"].hidden = True
    ws.row_dimensions[new_row].height = 18

    return {"new": new_count, "deleted": del_count, "total": total, "status": status}


def update_summary_sheet(wb: openpyxl.Workbook, today_str: str, results: dict):
    ws = wb["PODSUMOWANIE"]
    for i, p in enumerate(PROFILES, 2):
        r = results.get(p["name"], {})
        status_bg = "C6EFCE" if r.get("status") == "OK" else "FFC7CE"
        data_cell(ws, i, 1, p["name"],         bold=True,  align="left")
        data_cell(ws, i, 2, today_str)
        data_cell(ws, i, 3, r.get("total"))
        data_cell(ws, i, 4, r.get("new"),      bg="C6EFCE" if (r.get("new") or 0) > 0 else None, bold=True)
        data_cell(ws, i, 5, r.get("deleted"),  bg="FFC7CE" if (r.get("deleted") or 0) > 0 else None, bold=True)
        data_cell(ws, i, 6, r.get("status"),   bg=status_bg)


# ─── GŁÓWNA LOGIKA ───────────────────────────────────────────────────────────

def run():
    today_str = datetime.now().strftime("%Y-%m-%d %H:%M")
    print(f"\n{'='*55}")
    print(f"  OLX Monitor  |  {today_str}")
    print(f"{'='*55}")

    os.makedirs("data", exist_ok=True)
    wb = load_or_create_workbook()
    results  = {}
    all_ads  = {}   # przechowuje listę ogłoszeń dla Sheets sync

    for profile in PROFILES:
        name = profile["name"]
        url  = profile["url"]
        print(f"\n▶  {name}")
        print(f"   URL: {url}")

        total     = get_ad_count(url)
        today_ads = []

        if profile["type"] == "user":
            today_ads = get_individual_ads(url)
            print(f"   Ogłoszeń (scraping listy): {len(today_ads)}")
            if total is None:
                total = len(today_ads)

        print(f"   Łącznie: {total}")
        time.sleep(2)

        if name not in wb.sheetnames:
            _create_profile_sheet(wb, name)

        ws = wb[name]
        r  = update_profile_sheet(ws, profile, today_str, total, today_ads)
        results[name]  = r
        all_ads[name]  = today_ads
        print(f"   ✅  Nowe: +{r['new']}  |  Usunięte: -{r['deleted']}  |  Status: {r['status']}")

    update_summary_sheet(wb, today_str, results)
    wb.save(EXCEL_FILE)
    print(f"\n✔  Dane zapisane → {EXCEL_FILE}")

    # Synchronizacja z Google Sheets (jeśli skonfigurowana)
    if ENABLE_SHEETS_SYNC:
        from sheets_sync import sync_to_sheets
        sync_to_sheets(today_str, results, all_ads)
    else:
        print("\n☁️  Google Sheets sync pominięty (brak GOOGLE_SERVICE_ACCOUNT_JSON)")

    # Log JSON dla GitHub Actions summary
    log = {"date": today_str, "results": results}
    with open("data/last_run.json", "w", encoding="utf-8") as f:
        json.dump(log, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    run()
