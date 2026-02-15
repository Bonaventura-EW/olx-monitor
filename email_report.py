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


# ─── HTML E-MAIL ─────────────────────────────────────────────────────────────

def build_html_email(summary: dict, weekly_data: dict, analysis: str) -> str:
    today      = datetime.now().strftime("%d.%m.%Y")
    week_start = (datetime.now() - timedelta(days=6)).strftime("%d.%m.%Y")

    # ── Tabela podsumowania tygodnia ──
    summary_rows = ""
    for profile, s in summary.items():
        trend       = "↑" if s["net_week"] > 0 else ("↓" if s["net_week"] < 0 else "→")
        new_style   = "color:#1a7a3c;font-weight:bold;" if s["total_new"] > 0 else ""
        del_style   = "color:#c0392b;font-weight:bold;" if s["total_deleted"] > 0 else ""
        net_color   = "#1a7a3c" if s["net_week"] > 0 else ("#c0392b" if s["net_week"] < 0 else "#555")
        err_style   = "color:#c0392b;font-weight:bold;" if s["errors"] > 0 else "color:#888;"
        net_str     = f"{s['net_week']:+d}{trend}"

        summary_rows += f"""
        <tr>
          <td style="padding:10px 14px;border-bottom:1px solid #eee;font-weight:600;">{profile}</td>
          <td style="padding:10px 14px;border-bottom:1px solid #eee;text-align:center;">{s['days_tracked']}</td>
          <td style="padding:10px 14px;border-bottom:1px solid #eee;text-align:center;font-weight:600;">{s['last_count']}</td>
          <td style="padding:10px 14px;border-bottom:1px solid #eee;text-align:center;{new_style}">{s['total_new']:+d}</td>
          <td style="padding:10px 14px;border-bottom:1px solid #eee;text-align:center;{del_style}">{s['total_deleted']}</td>
          <td style="padding:10px 14px;border-bottom:1px solid #eee;text-align:center;color:{net_color};font-weight:bold;">{net_str}</td>
          <td style="padding:10px 14px;border-bottom:1px solid #eee;text-align:center;{err_style}">{s['errors']}</td>
        </tr>"""

    # ── Zestawienie dzienne ──
    daily_sections = ""
    for profile, rows in weekly_data.items():
        daily_rows = ""
        for i, r in enumerate(rows):
            bg       = "#f9f9f9" if i % 2 == 0 else "#ffffff"
            net_str  = f"{r['net']:+d}" if r['net'] != 0 else "—"
            net_col  = "#1a7a3c" if r['net'] > 0 else ("#c0392b" if r['net'] < 0 else "#888")
            new_col  = "#1a7a3c" if r['new'] > 0 else "#333"
            del_col  = "#c0392b" if r['deleted'] > 0 else "#333"
            daily_rows += f"""
            <tr style="background:{bg};">
              <td style="padding:8px 12px;border-bottom:1px solid #eee;">{r['date']}</td>
              <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:center;font-weight:600;">{r['total']}</td>
              <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:center;color:{new_col};font-weight:{'bold' if r['new']>0 else 'normal'};">{r['new']:+d}</td>
              <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:center;color:{del_col};">{r['deleted']}</td>
              <td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:center;color:{net_col};font-weight:bold;">{net_str}</td>
            </tr>"""

        daily_sections += f"""
        <div style="margin-bottom:24px;">
          <h3 style="margin:0 0 8px 0;font-size:13px;text-transform:uppercase;
                     letter-spacing:1px;color:#2c5f8a;">{profile}</h3>
          <table width="100%" cellpadding="0" cellspacing="0"
                 style="border-collapse:collapse;font-size:13px;">
            <thead>
              <tr style="background:#2c5f8a;color:#fff;">
                <th style="padding:8px 12px;text-align:left;">Data</th>
                <th style="padding:8px 12px;text-align:center;">Ogłoszeń</th>
                <th style="padding:8px 12px;text-align:center;">Nowe</th>
                <th style="padding:8px 12px;text-align:center;">Usunięte</th>
                <th style="padding:8px 12px;text-align:center;">Netto</th>
              </tr>
            </thead>
            <tbody>{daily_rows}</tbody>
          </table>
        </div>"""

    return f"""<!DOCTYPE html>
<html lang="pl">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:#f0f4f8;font-family:Arial,sans-serif;">
<div style="max-width:680px;margin:32px auto;background:#fff;border-radius:10px;
            overflow:hidden;box-shadow:0 2px 12px rgba(0,0,0,.08);">

  <!-- NAGŁÓWEK -->
  <div style="background:#2c5f8a;padding:28px 32px;">
    <h1 style="margin:0;color:#fff;font-size:20px;font-weight:700;">📊 OLX Monitor</h1>
    <p style="margin:6px 0 0;color:#a8c8e8;font-size:13px;">
      Raport tygodniowy &nbsp;·&nbsp; {week_start} – {today}
    </p>
  </div>

  <div style="padding:28px 32px;">

    <!-- PODSUMOWANIE TYGODNIA -->
    <h2 style="margin:0 0 16px;font-size:15px;color:#2c5f8a;text-transform:uppercase;
               letter-spacing:.5px;border-bottom:2px solid #2c5f8a;padding-bottom:8px;">
      Podsumowanie tygodnia
    </h2>
    <table width="100%" cellpadding="0" cellspacing="0"
           style="border-collapse:collapse;font-size:13px;margin-bottom:8px;">
      <thead>
        <tr style="background:#2c5f8a;color:#fff;">
          <th style="padding:10px 14px;text-align:left;">Profil</th>
          <th style="padding:10px 14px;text-align:center;">Dni</th>
          <th style="padding:10px 14px;text-align:center;">Stan</th>
          <th style="padding:10px 14px;text-align:center;">Nowe</th>
          <th style="padding:10px 14px;text-align:center;">Usun.</th>
          <th style="padding:10px 14px;text-align:center;">Netto</th>
          <th style="padding:10px 14px;text-align:center;">Błędy</th>
        </tr>
      </thead>
      <tbody>{summary_rows}</tbody>
    </table>
    <p style="margin:4px 0 24px;font-size:11px;color:#888;">
      Stan = aktualna liczba ogłoszeń &nbsp;|&nbsp; Nowe = przybyło w tygodniu &nbsp;|&nbsp;
      Usun. = usunięto &nbsp;|&nbsp; Netto = zmiana netto &nbsp;|&nbsp; Błędy = dni z błędem odczytu
    </p>

    <!-- ANALIZA AI -->
    <h2 style="margin:0 0 12px;font-size:15px;color:#2c5f8a;text-transform:uppercase;
               letter-spacing:.5px;border-bottom:2px solid #2c5f8a;padding-bottom:8px;">
      🤖 Analiza
    </h2>
    <div style="background:#f0f4f8;border-left:4px solid #2c5f8a;padding:16px 20px;
                border-radius:0 6px 6px 0;margin-bottom:28px;font-size:14px;
                line-height:1.7;color:#333;">
      {analysis.replace(chr(10), '<br>')}
    </div>

    <!-- ZESTAWIENIE DZIENNE -->
    <h2 style="margin:0 0 16px;font-size:15px;color:#2c5f8a;text-transform:uppercase;
               letter-spacing:.5px;border-bottom:2px solid #2c5f8a;padding-bottom:8px;">
      📅 Zestawienie dzienne
    </h2>
    {daily_sections}

  </div>

  <!-- STOPKA -->
  <div style="background:#f0f4f8;padding:16px 32px;text-align:center;
              font-size:11px;color:#888;border-top:1px solid #e0e8f0;">
    Raport wygenerowany automatycznie przez OLX Monitor &nbsp;·&nbsp;
    GitHub Actions &nbsp;·&nbsp; {datetime.now().strftime("%Y-%m-%d %H:%M")}
  </div>

</div>
</body>
</html>"""


# ─── ANALIZA AI (Google Gemini API) ──────────────────────────────────────────

def generate_ai_analysis(summary: dict, weekly_data: dict) -> str:
    """Wysyła dane do Google Gemini API i zwraca analizę tekstową (5-10 zdań)."""
    api_key = os.environ.get("GEMINI_API_KEY", "")
    if not api_key:
        return "⚠  Analiza AI niedostępna – brak klucza GEMINI_API_KEY."

    data_for_ai = {}
    for profile, s in summary.items():
        data_for_ai[profile] = {
            "stan_na_koniec_tygodnia":   s["last_count"],
            "stan_na_poczatek_tygodnia": s["first_count"],
            "laczna_liczba_nowych":      s["total_new"],
            "laczna_liczba_usunietych":  s["total_deleted"],
            "zmiana_netto":              s["net_week"],
            "dni_monitorowania":         s["days_tracked"],
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
        url  = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={api_key}"
        resp = requests.post(url, json={
            "contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {"maxOutputTokens": 600, "temperature": 0.7},
        }, timeout=30)

        if not resp.ok:
            print(f"  ⚠  Gemini API error {resp.status_code}: {resp.text}")
            return f"⚠ Błąd API Gemini ({resp.status_code}): {resp.text[:200]}"

        return resp.json()["candidates"][0]["content"]["parts"][0]["text"].strip()

    except Exception as e:
        return f"⚠  Błąd generowania analizy AI: {e}"


# ─── WYSYŁANIE E-MAILA ───────────────────────────────────────────────────────

def send_email(subject: str, html_body: str):
    """Wysyła e-mail HTML przez Gmail SMTP."""
    gmail_password = os.environ.get("GMAIL_APP_PASSWORD", "")
    if not gmail_password:
        print("⚠  Brak GMAIL_APP_PASSWORD – e-mail nie zostanie wysłany.")
        return False

    msg = MIMEMultipart("mixed")
    msg["Subject"] = subject
    msg["From"]    = SENDER_EMAIL
    msg["To"]      = RECIPIENT_EMAIL

    # Część HTML
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    # Załącz plik Excel jeśli istnieje
    if os.path.exists(EXCEL_FILE):
        today           = datetime.now().strftime("%Y-%m-%d")
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
    html    = build_html_email(summary, weekly_data, analysis)

    print("  ✉  Treść HTML wygenerowana, wysyłam...")
    send_email(subject, html)


if __name__ == "__main__":
    send_weekly_report()
