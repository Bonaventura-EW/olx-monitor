#!/usr/bin/env python3
"""
Wstrzyknij dane z JSON do olx_dashboard.html
Wywoływane z GitHub Actions workflow
"""

import json
import re
import os
import sys

def main():
    try:
        # Wczytaj price_history
        with open("data/price_history.json", "r", encoding="utf-8") as f:
            history = json.load(f)
        print(f"✓ Wczytano price_history.json: {len(history)} ogłoszeń")
    except Exception as e:
        print(f"❌ Błąd wczytywania price_history.json: {e}")
        sys.exit(1)

    try:
        # Wczytaj market_total z last_run.json
        market_total = None
        if os.path.exists("data/last_run.json"):
            with open("data/last_run.json", "r", encoding="utf-8") as f:
                last_run = json.load(f)
            market_total = last_run.get("market_total")
        print(f"✓ Market total: {market_total}")
    except Exception as e:
        print(f"⚠  Błąd wczytywania last_run.json: {e}")
        market_total = None

    try:
        with open("olx_dashboard.html", "r", encoding="utf-8") as f:
            html = f.read()
        print(f"✓ Wczytano olx_dashboard.html: {len(html)} bajtów")
    except Exception as e:
        print(f"❌ Błąd wczytywania olx_dashboard.html: {e}")
        sys.exit(1)

    # ──────────────────────────────────────────────────────
    # Wstrzyknij PRICE_HISTORY
    # ──────────────────────────────────────────────────────
    inject_ph = f"window.__PRICE_HISTORY__ = {json.dumps(history, ensure_ascii=False)};"

    if "window.__PRICE_HISTORY__" in html:
        # Istniejący marker — zastąp
        pattern = r"window\.__PRICE_HISTORY__\s*=\s*(?:\{.*?\}|null);?"
        html_new = re.sub(pattern, inject_ph, html, count=1, flags=re.DOTALL)
        if html_new != html:
            html = html_new
            print("✓ Price history zaktualizowana (istniejący marker)")
        else:
            print("⚠  Nie udało się zaktualizować istniejącego markera — dodaję nowy")
            # Fallback: dodaj przed </script>
            if "</script>" in html:
                html = html.replace("</script>", f"\n{inject_ph}\n</script>", 1)
                print("✓ Price history dodana (fallback)")
    else:
        # Nowy marker — dodaj
        if "</script>" in html:
            html = html.replace("</script>", f"\n{inject_ph}\n</script>", 1)
            print("✓ Price history dodana (nowy marker)")

    # ──────────────────────────────────────────────────────
    # Wstrzyknij MARKET_TOTAL
    # ──────────────────────────────────────────────────────
    if market_total is not None:
        inject_mt = f"window.__MARKET_TOTAL__ = {market_total};"
        
        if "window.__MARKET_TOTAL__" in html:
            html = re.sub(
                r"window\.__MARKET_TOTAL__\s*=\s*\d+;?",
                inject_mt,
                html,
                count=1
            )
            print("✓ Market total zaktualizowany")
        else:
            if "</script>" in html:
                html = html.replace("</script>", f"\n{inject_mt}\n</script>", 1)
                print("✓ Market total dodany")

    # ──────────────────────────────────────────────────────
    # Wstrzyknij LAST_RUN (czas ostatniego skanu)
    # ──────────────────────────────────────────────────────
    try:
        if os.path.exists("data/last_run.json"):
            with open("data/last_run.json", "r", encoding="utf-8") as f_lr:
                last_run_data = json.load(f_lr)
            run_at = last_run_data.get("run_at", "")
            if run_at:
                inject_lr = f'window.__LAST_RUN__ = "{run_at}";'
                if "window.__LAST_RUN__" in html:
                    html = re.sub(
                        r'window\.__LAST_RUN__\s*=\s*"[^"]*";?',
                        inject_lr,
                        html,
                        count=1
                    )
                    print("✓ Last run zaktualizowany")
                else:
                    if "</script>" in html:
                        html = html.replace("</script>", f"\n{inject_lr}\n</script>", 1)
                        print("✓ Last run dodany")
    except Exception as e:
        print(f"⚠  Błąd wstrzykiwania LAST_RUN: {e}")

    # ──────────────────────────────────────────────────────
    # Wstrzyknij PROFILES_DATA (obsługa wieloliniowego JSON)
    # ──────────────────────────────────────────────────────
    try:
        if os.path.exists("data/profiles_state.json"):
            with open("data/profiles_state.json", "r", encoding="utf-8") as f_ps:
                profiles_state = json.load(f_ps)
            inject_pd = f"window.__PROFILES_DATA__ = {json.dumps(profiles_state, ensure_ascii=False)};"
            
            if "window.__PROFILES_DATA__" in html:
                # Najlepiej: znajdź dokładne granice zmiennej
                start = html.find("window.__PROFILES_DATA__ = ")
                if start != -1:
                    # Szukaj końca (";")
                    search_from = start + len("window.__PROFILES_DATA__ = ")
                    end = html.find(";", search_from)
                    if end != -1:
                        # Zamień zawartość między "= " a ";"
                        html = html[:start] + f"window.__PROFILES_DATA__ = {json.dumps(profiles_state, ensure_ascii=False)};" + html[end+1:]
                        print("✓ Profiles data zaktualizowana (precyzyjnie)")
                    else:
                        print("⚠  Nie znaleziono końca PROFILES_DATA")
                else:
                    print("⚠  Nie znaleziono początku PROFILES_DATA")
            else:
                # Nowy marker
                if "</script>" in html:
                    html = html.replace("</script>", f"\n{inject_pd}\n</script>", 1)
                    print("✓ Profiles data dodana (nowy marker)")
    except Exception as e:
        print(f"⚠  Błąd wstrzykiwania PROFILES_DATA: {e}")

    # ──────────────────────────────────────────────────────
    # Zapisz zaktualizowany HTML
    # ──────────────────────────────────────────────────────
    try:
        with open("olx_dashboard.html", "w", encoding="utf-8") as f:
            f.write(html)
        print(f"✅ Dashboard zaktualizowany – {len(history)} ogłoszeń, rynek: {market_total}")
    except Exception as e:
        print(f"❌ Błąd zapisu dashboard: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
