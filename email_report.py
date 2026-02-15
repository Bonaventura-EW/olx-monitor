"""
Tygodniowy raport e-mail z analizą AI.
Wysyłany w każdy poniedziałek – zbiera dane z ostatnich 7 dni z pliku Excel
i wysyła podsumowanie przez Gmail SMTP.
"""

import smtplib
import json
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime, timedelta
import openpyxl
import requests

# ─── KONFIGURACJA ────────────────────────────────────────────────────────────

SENDER_EMAIL    = "slowholidays00@gmail.com"
RECIPIENT_EMAIL = "malczarski@gmail.com"
EXCEL_FILE      = "data/olx_monitoring.xlsx"

PROFILES = [
    "wszystkie-lublin",
    "artymiuk",
    "poqui",
    "stylowepokoje",
    "villahome",
]

# ─── ZBIERANIE DANYCH Z EXCELA ───────────────────────────────────────────────

def get_weekly_data() -> dict:
    """Odczytuje dane z ostatnich 7 dni z każdej zakładki Excela."""
    if not os.path.exists(EXCEL_FILE):
        return {}

    wb   = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    week_ago = datetime.now() - timedelta(days=7)
    data = {}

    for profile in PROFILES:
        if profile not in wb.sheetnames:
            continue

        ws   = wb[profile]
        rows = []

        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row[0]:
                continue
            try:
                # Kolumna A: data jako string "2026-02-15 09:00"
                row_date = datetime.strptime(str(row[0])[:16], "%Y-%m-%d %H:%M")
            except Exception:
                continue

            if row_date >= week_ago:
                rows.append({
                    "date":    str(row[0])[:10],
                    "total":   row[1] or 0,
                    "new":     row[2] or 0,
                    "deleted": row[3] or 0,
                    "net":     row[4] or 0,
                    "status":  row[7] or "?",
                })

        if rows:
            data[profile] = rows

    return data


def compute_summary(weekly_data: dict) -> dict:
    """Oblicza sumaryczne statystyki tygodniowe dla każdego profilu."""
    summary = {}
    for profile, rows in weekly_data.items():
        total_new     = sum(r["new"]     for r in rows)
        total_deleted = sum(r["deleted"] for r in rows)
        last_total    = rows[-1]["total"] if rows else 0
        first_total   = rows[0]["total"]  if rows else 0
        errors        = sum(1 for r in rows if r["status"] != "OK")

        summary[profile] = {
            "days_tracked": len(rows),
            "total_new":     total_new,
            "total_deleted": total_deleted,
            "net_week":      total_new - total_deleted,
            "last_count":    last_total,
            "first_count":   first_total,
            "errors":        errors,
            "rows":          rows,
        }
    return summary


# ─── TABELA TEKSTOWA ─────────────────────────────────────────────────────────

def build_text_table(summary: dict) -> str:
    """Buduje czytelną tabelę ASCII do wklejenia w e-mail."""
    header = (
        f"{'Profil':<22} {'Dni':>4} {'Stan':>6} {'Nowe':>6} "
        f"{'Usun.':>6} {'Netto':>6} {'Błędy':>6}"
    )
    sep = "─" * len(header)
    lines = [sep, header, sep]

    for profile, s in summary.items():
        trend = "↑" if s["net_week"] > 0 else ("↓" if s["net_week"] < 0 else "→")
        line = (
            f"{profile:<22} {s['days_tracked']:>4} {s['last_count']:>6} "
            f"{s['total_new']:>+6} {s['total_deleted']:>6} "
            f"{s['net_week']:>+5}{trend} {s['errors']:>5}"
        )
        lines.append(line)

    lines.append(sep)
    return "\n".join(lines)


def build_daily_breakdown(weekly_data: dict) -> str:
    """Buduje dzienne zestawienie dla każdego profilu."""
    sections = []
    for profile, rows in weekly_data.items():
        lines = [f"\n  {profile.upper()}"]
        lines.append(f"  {'Data':<12} {'Ogłoszeń':>10} {'Nowe':>7} {'Usun.':>7} {'Netto':>7}")
        lines.append("  " + "─" * 45)
        for r in rows:
            net_str = f"{r['net']:+d}" if r['net'] != 0 else "  —"
            lines.append(
                f"  {r['date']:<12} {r['total']:>10} "
                f"{r['new']:>+7} {r['deleted']:>7} {net_str:>7}"
            )
        sections.append("\n".join(lines))
    return "\n".join(sections)


# ─── ANALIZA AI (Claude API) ──────────────────────────────────────────────────

def generate_ai_analysis(summary: dict, weekly_data: dict) -> str:
    """Wysyła dane do Claude API i zwraca analizę tekstową (5-10 zdań)."""
    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        return "⚠  Analiza AI niedostępna – brak klucza ANTHROPIC_API_KEY."

    # Przygotuj dane dla modelu
    data_for_ai = {}
    for profile, s in summary.items():
        data_for_ai[profile] = {
            "stan_na_koniec_tygodnia":     s["last_count"],
            "stan_na_poczatek_tygodnia":   s["first_count"],
            "laczna_liczba_nowych":        s["total_new"],
            "laczna_liczba_usunietych":    s["total_deleted"],
            "zmiana_netto":                s["net_week"],
            "dni_monitorowania":           s["days_tracked"],
        }

    prompt = f"""Jesteś analitykiem rynku nieruchomości. 
Poniżej masz tygodniowe dane z monitoringu ogłoszeń na OLX.pl (stancje i pokoje w Lublinie).

Dane z ostatnich 7 dni:
{json.dumps(data_for_ai, ensure_ascii=False, indent=2)}

Napisz zwięzłą analizę (5-10 zdań) po polsku. Uwzględnij:
- Ogólny trend na rynku pokoi w Lublinie (profil wszystkie-lublin)
- Aktywność poszczególnych wynajmujących (artymiuk, poqui, stylowepokoje, villahome)
- Czy rynek jest aktywny czy spokojny w tym tygodniu
- Które profile są najbardziej aktywne i co to może oznaczać
- Krótką rekomendację lub obserwację dla obserwującego rynek

Pisz naturalnie, bez wypunktowań, jako spójny tekst analityczny."""

    try:
        resp = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={
                "x-api-key":         api_key,
                "anthropic-version": "2023-06-01",
                "content-type":      "application/json",
            },
            json={
                "model":      "claude-opus-4-5-20251101",
                "max_tokens": 600,
                "messages":   [{"role": "user", "content": prompt}],
            },
            timeout=30,
        )
        resp.raise_for_status()
        return resp.json()["content"][0]["text"].strip()
    except Exception as e:
        return f"⚠  Błąd generowania analizy AI: {e}"


# ─── BUDOWANIE E-MAILA ───────────────────────────────────────────────────────

def build_email_body(summary: dict, weekly_data: dict, analysis: str) -> str:
    today      = datetime.now().strftime("%d.%m.%Y")
    week_start = (datetime.now() - timedelta(days=6)).strftime("%d.%m.%Y")
    table      = build_text_table(summary)
    breakdown  = build_daily_breakdown(weekly_data)

    return f"""OLX MONITOR – TYGODNIOWY RAPORT
Okres: {week_start} – {today}
{"═" * 55}

📊 PODSUMOWANIE TYGODNIA
{table}

Kolumny: Stan = aktualna liczba ogłoszeń | Nowe = przybyło w tygodniu
         Usun. = usunięto | Netto = zmiana netto | Błędy = dni z błędem odczytu

{"═" * 55}

🤖 ANALIZA
{analysis}

{"═" * 55}

📅 ZESTAWIENIE DZIENNE
{breakdown}

{"═" * 55}
Raport wygenerowany automatycznie przez OLX Monitor
GitHub Actions | {datetime.now().strftime("%Y-%m-%d %H:%M")}
"""


# ─── WYSYŁANIE E-MAILA ───────────────────────────────────────────────────────

def send_email(subject: str, body: str):
    """Wysyła e-mail przez Gmail SMTP używając App Password ze zmiennych środowiskowych."""
    gmail_password = os.environ.get("GMAIL_APP_PASSWORD", "")
    if not gmail_password:
        print("⚠  Brak GMAIL_APP_PASSWORD – e-mail nie zostanie wysłany.")
        return False

    msg = MIMEMultipart("mixed")
    msg["Subject"] = subject
    msg["From"]    = SENDER_EMAIL
    msg["To"]      = RECIPIENT_EMAIL
    msg.attach(MIMEText(body, "plain", "utf-8"))

    # Załącz plik Excel jeśli istnieje
    if os.path.exists(EXCEL_FILE):
        today      = datetime.now().strftime("%Y-%m-%d")
        attachment_name = f"OLX_Monitor_{today}.xlsx"
        with open(EXCEL_FILE, "rb") as f:
            part = MIMEBase("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", "attachment", filename=attachment_name)
        msg.attach(part)
        print(f"  📎 Załączono: {attachment_name}")
    else:
        print("  ⚠  Plik Excel nie znaleziony – wysyłam bez załącznika.")

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(SENDER_EMAIL, gmail_password)
            server.sendmail(SENDER_EMAIL, RECIPIENT_EMAIL, msg.as_string())
        print(f"✅  E-mail wysłany → {RECIPIENT_EMAIL}")
        return True
    except Exception as e:
        print(f"❌  Błąd wysyłania e-maila: {e}")
        return False


# ─── GŁÓWNA FUNKCJA ──────────────────────────────────────────────────────────

def send_weekly_report():
    print("\n📧  Generowanie tygodniowego raportu e-mail...")

    weekly_data = get_weekly_data()
    if not weekly_data:
        print("⚠  Brak danych z ostatnich 7 dni – raport nie zostanie wysłany.")
        return

    summary  = compute_summary(weekly_data)
    analysis = generate_ai_analysis(summary, weekly_data)

    today   = datetime.now().strftime("%d.%m.%Y")
    subject = f"📊 OLX Monitor – raport tygodniowy {today}"
    body    = build_email_body(summary, weekly_data, analysis)

    print("\n" + "─" * 55)
    print(body[:500] + "...")   # podgląd w logach
    print("─" * 55)

    send_email(subject, body)


if __name__ == "__main__":
    send_weekly_report()
