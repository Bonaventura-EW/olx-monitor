name: OLX Monitor – codzienny raport

on:
  schedule:
    # Codziennie o 09:00 CET (07:00 UTC zima)
    - cron: "0 7 * * *"

  # Możliwość ręcznego uruchomienia z panelu GitHub
  workflow_dispatch:
    inputs:
      force_email:
        description: 'Wymuś wysyłkę e-maila (niezależnie od dnia tygodnia)'
        required: false
        default: 'false'
        type: boolean

jobs:
  monitor:
    runs-on: ubuntu-latest
    permissions:
      contents: write

    steps:
      - name: Checkout repozytorium
        uses: actions/checkout@v4
        with:
          fetch-depth: 0

      - name: Ustaw Python 3.11
        uses: actions/setup-python@v5
        with:
          python-version: "3.11"
          cache: "pip"

      - name: Zainstaluj zależności
        run: pip install -r requirements.txt

      # ── Codzienny scraping i zapis do Excela ─────────────────────────────
      - name: Uruchom OLX Monitor (scraping)
        run: python olx_monitor.py

      # ── Tygodniowy raport e-mail (tylko w poniedziałek lub ręcznie) ──────
      - name: Wyślij tygodniowy raport e-mail
        run: |
          DAY=$(date +%u)   # 1=poniedziałek, 7=niedziela
          FORCE="${{ github.event.inputs.force_email }}"
          if [ "$DAY" = "1" ] || [ "$FORCE" = "true" ]; then
            echo "📧 Wysyłam tygodniowy raport e-mail..."
            python email_report.py
          else
            echo "⏭  Nie poniedziałek (dzień $DAY) – pomijam e-mail."
          fi
        env:
          GMAIL_APP_PASSWORD: ${{ secrets.GMAIL_APP_PASSWORD }}
          GEMINI_API_KEY: ${{ secrets.GEMINI_API_KEY }}

      # ── Podsumowanie w logach GitHub ─────────────────────────────────────
      - name: Pokaż wyniki (summary)
        run: |
          echo "## 📊 OLX Monitor – wyniki $(date +'%Y-%m-%d')" >> $GITHUB_STEP_SUMMARY
          echo "" >> $GITHUB_STEP_SUMMARY
          if [ -f data/last_run.json ]; then
            echo '```json' >> $GITHUB_STEP_SUMMARY
            cat data/last_run.json >> $GITHUB_STEP_SUMMARY
            echo '```' >> $GITHUB_STEP_SUMMARY
          fi

      # ── Commit zaktualizowanego pliku Excel ──────────────────────────────
      - name: Zapisz plik Excel do repozytorium
        run: |
          git config user.name  "OLX Monitor Bot"
          git config user.email "bot@github-actions"
          git add data/olx_monitoring.xlsx data/last_run.json
          git diff --cached --quiet && echo "Brak zmian" && exit 0
          git commit -m "📊 OLX Monitor $(date +'%Y-%m-%d %H:%M')"
          git push origin HEAD:main --force-with-lease
