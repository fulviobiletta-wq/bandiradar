"""
notify.py — Email settimanale con riepilogo bandi
Usa Resend (gratuito: 3.000 email/mese, piano free permanente)
Setup: https://resend.com → API Key gratuita → aggiungi RESEND_API_KEY nei GitHub Secrets
"""

import json
import os
import sys
from datetime import date
from pathlib import Path

import httpx

RESEND_API_KEY  = os.environ.get("RESEND_API_KEY", "")
EMAIL_TO        = os.environ.get("EMAIL_TO", "")          # es. commerciale@azienda.it
EMAIL_FROM      = os.environ.get("EMAIL_FROM", "BandiRadar <noreply@tuodominio.com>")
REPO_OWNER      = os.environ.get("GITHUB_REPOSITORY_OWNER", "TUO-USERNAME")
REPO_NAME       = os.environ.get("GITHUB_REPOSITORY", "TUO-USERNAME/bandi-radar").split("/")[-1]

JSON_FILE = Path(__file__).parent / "output" / "bandi.json"
TODAY     = date.today()


def carica_dati():
    if not JSON_FILE.exists():
        print("bandi.json non trovato — skip email")
        return None
    with open(JSON_FILE, encoding="utf-8") as f:
        return json.load(f)


def build_email(dati):
    bandi       = dati["bandi"]
    aggiornato  = dati.get("aggiornato", TODAY.isoformat())
    link_xlsx   = f"https://{REPO_OWNER}.github.io/{REPO_NAME}/BandiRadar.xlsx"
    link_web    = f"https://{REPO_OWNER}.github.io/{REPO_NAME}/"

    urgenti = [(b, (date.fromisoformat(b["scadenza"]) - TODAY).days)
               for b in bandi
               if b.get("scadenza") and 0 <= (date.fromisoformat(b["scadenza"]) - TODAY).days <= 30]
    urgenti.sort(key=lambda x: x[1])

    nuovi_aperti = [b for b in bandi if b.get("status") == "APERTO"][:5]

    # ── HTML email ────────────────────────────────────────
    urgenti_html = ""
    if urgenti:
        rows = ""
        for b, dl in urgenti:
            rows += f"""
            <tr>
              <td style="padding:8px 12px;border-bottom:1px solid #fde68a;font-weight:600">{b['titolo']}</td>
              <td style="padding:8px 12px;border-bottom:1px solid #fde68a;text-align:center;color:#92400e;font-weight:700">{dl} giorni</td>
              <td style="padding:8px 12px;border-bottom:1px solid #fde68a;text-align:center">
                <a href="{b['link']}" style="color:#2563eb">Apri</a>
              </td>
            </tr>"""
        urgenti_html = f"""
        <div style="background:#fffbeb;border:2px solid #f59e0b;border-radius:10px;margin:20px 0;padding:20px">
          <h3 style="color:#92400e;margin:0 0 12px">⚠️ {len(urgenti)} bandi in scadenza entro 30 giorni</h3>
          <table style="width:100%;border-collapse:collapse">
            <tr style="background:#fef3c7">
              <th style="padding:8px 12px;text-align:left">Bando</th>
              <th style="padding:8px 12px;text-align:center">Giorni rimasti</th>
              <th style="padding:8px 12px;text-align:center">Link</th>
            </tr>{rows}
          </table>
        </div>"""
    else:
        urgenti_html = """
        <div style="background:#f0fdf4;border:1px solid #86efac;border-radius:10px;margin:20px 0;padding:16px;color:#166534">
          ✅ Nessun bando in scadenza imminente questa settimana.
        </div>"""

    bandi_rows = ""
    for b in nuovi_aperti:
        bandi_rows += f"""
        <tr style="border-bottom:1px solid #e5e7eb">
          <td style="padding:8px 12px;font-weight:600">{b['titolo']}</td>
          <td style="padding:8px 12px;color:#6b7280">{b['regione']}</td>
          <td style="padding:8px 12px;color:#1e3a5f;font-weight:700">{b['contributo']}</td>
          <td style="padding:8px 12px">
            <a href="{b['link']}" style="color:#2563eb;font-size:12px">Apri →</a>
          </td>
        </tr>"""

    html = f"""
<!DOCTYPE html>
<html>
<head><meta charset="utf-8"></head>
<body style="font-family:Arial,sans-serif;max-width:700px;margin:0 auto;color:#1a1a2e">

  <div style="background:linear-gradient(135deg,#0f1c35,#1e3a5f);padding:30px;border-radius:12px 12px 0 0;text-align:center">
    <h1 style="color:#fff;margin:0;font-size:24px">📋 BandiRadar</h1>
    <p style="color:rgba(255,255,255,0.6);margin:6px 0 0;font-size:13px">Riepilogo settimanale — {TODAY.strftime('%d/%m/%Y')}</p>
  </div>

  <div style="background:#fff;padding:28px;border:1px solid #e5e7eb;border-top:none">

    <div style="display:flex;gap:16px;margin-bottom:24px">
      {''.join(f"""
      <div style="flex:1;background:#f8fafc;border-radius:8px;padding:14px;text-align:center;border:1px solid #e5e7eb">
        <div style="font-size:22px;font-weight:800;color:{col}">{val}</div>
        <div style="font-size:11px;color:#6b7280;margin-top:3px">{label}</div>
      </div>""" for val, label, col in [
        (dati['totale'], 'Bandi totali', '#1e3a5f'),
        (dati['aperti'], 'Aperti', '#059669'),
        (dati['in_scadenza'], 'In scadenza', '#d97706'),
        (dati['verificati'], 'Verificati', '#7c3aed'),
      ])}
    </div>

    {urgenti_html}

    <h3 style="color:#1e3a5f;margin:24px 0 12px">Bandi aperti — panoramica</h3>
    <table style="width:100%;border-collapse:collapse;font-size:13px">
      <tr style="background:#f8fafc">
        <th style="padding:8px 12px;text-align:left">Bando</th>
        <th style="padding:8px 12px;text-align:left">Regione</th>
        <th style="padding:8px 12px;text-align:left">Contributo</th>
        <th style="padding:8px 12px"></th>
      </tr>
      {bandi_rows}
    </table>

    <div style="margin-top:24px;text-align:center">
      <a href="{link_xlsx}" style="display:inline-block;background:linear-gradient(135deg,#f59e0b,#ef4444);color:#fff;padding:12px 28px;border-radius:8px;font-weight:700;text-decoration:none;margin-right:10px">
        📥 Scarica Excel aggiornato
      </a>
      <a href="{link_web}" style="display:inline-block;background:#1e3a5f;color:#fff;padding:12px 28px;border-radius:8px;font-weight:700;text-decoration:none">
        🌐 Vedi tutti i bandi
      </a>
    </div>

    <p style="margin-top:20px;font-size:11px;color:#9ca3af;text-align:center">
      Bandi verificati manualmente aggiornati al {aggiornato}. Verificare sempre le fonti ufficiali prima di presentare ai clienti.<br>
      BandiRadar — aggiornamento automatico settimanale via GitHub Actions
    </p>
  </div>

</body>
</html>"""

    subject = f"📋 BandiRadar {TODAY.strftime('%d/%m/%Y')} — {dati['aperti']} bandi aperti"
    if urgenti:
        subject = f"⚠️ {len(urgenti)} bandi in scadenza | " + subject

    return subject, html


def invia_email(subject, html):
    if not RESEND_API_KEY or not EMAIL_TO:
        print("RESEND_API_KEY o EMAIL_TO non impostati — skip email")
        return False

    try:
        r = httpx.post(
            "https://api.resend.com/emails",
            headers={"Authorization": f"Bearer {RESEND_API_KEY}", "Content-Type": "application/json"},
            json={"from": EMAIL_FROM, "to": [EMAIL_TO], "subject": subject, "html": html},
            timeout=30,
        )
        if r.status_code in (200, 201):
            print(f"Email inviata a {EMAIL_TO}")
            return True
        else:
            print(f"Errore email: {r.status_code} — {r.text}")
            return False
    except Exception as e:
        print(f"Errore invio email: {e}")
        return False


def main():
    print(f"\n--- Notifica email {TODAY.strftime('%d/%m/%Y')} ---")
    dati = carica_dati()
    if not dati:
        return 0
    subject, html = build_email(dati)
    print(f"Subject: {subject}")
    invia_email(subject, html)
    return 0

if __name__ == "__main__":
    sys.exit(main())
