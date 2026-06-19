"""
notify.py — Email settimanale con riepilogo bandi
Usa Resend (gratuito: 3.000 email/mese)
"""

import json
import os
import sys
from datetime import date
from pathlib import Path

import httpx

RESEND_API_KEY = os.environ.get("RESEND_API_KEY", "")
EMAIL_TO       = os.environ.get("EMAIL_TO", "")
EMAIL_FROM     = os.environ.get("EMAIL_FROM", "BandiRadar <onboarding@resend.dev>")
REPO_OWNER     = os.environ.get("GITHUB_REPOSITORY_OWNER", "TUO-USERNAME")
REPO_NAME      = os.environ.get("GITHUB_REPOSITORY", "bandi-radar").split("/")[-1]

JSON_FILE = Path(__file__).parent / "output" / "bandi.json"
TODAY     = date.today()


def carica_dati():
    if not JSON_FILE.exists():
        print("bandi.json non trovato — skip email")
        return None
    with open(JSON_FILE, encoding="utf-8") as f:
        return json.load(f)


def days_left(scadenza):
    try:
        return (date.fromisoformat(scadenza) - TODAY).days
    except Exception:
        return 999


def format_date(iso):
    try:
        return date.fromisoformat(iso).strftime("%d/%m/%Y")
    except Exception:
        return iso


def build_email(dati):
    bandi      = dati["bandi"]
    aggiornato = dati.get("aggiornato", TODAY.isoformat())
    link_xlsx  = "https://{}.github.io/{}/BandiRadar.xlsx".format(REPO_OWNER, REPO_NAME)
    link_web   = "https://{}.github.io/{}/".format(REPO_OWNER, REPO_NAME)

    urgenti = [(b, days_left(b["scadenza"]))
               for b in bandi
               if b.get("scadenza") and 0 <= days_left(b["scadenza"]) <= 30]
    urgenti.sort(key=lambda x: x[1])

    # ── Sezione urgenti ───────────────────────────────────
    if urgenti:
        righe_urgenti = ""
        for b, dl in urgenti:
            righe_urgenti += (
                "<tr>"
                "<td style='padding:8px 12px;border-bottom:1px solid #fde68a;font-weight:600'>{}</td>"
                "<td style='padding:8px 12px;border-bottom:1px solid #fde68a;text-align:center;"
                "color:#92400e;font-weight:700'>{} giorni</td>"
                "<td style='padding:8px 12px;border-bottom:1px solid #fde68a;text-align:center'>"
                "<a href='{}' style='color:#2563eb'>Apri</a></td>"
                "</tr>"
            ).format(b["titolo"], dl, b.get("link", "#"))

        box_urgenti = (
            "<div style='background:#fffbeb;border:2px solid #f59e0b;border-radius:10px;"
            "margin:20px 0;padding:20px'>"
            "<h3 style='color:#92400e;margin:0 0 12px'>⚠️ {} bandi in scadenza entro 30 giorni</h3>"
            "<table style='width:100%;border-collapse:collapse'>"
            "<tr style='background:#fef3c7'>"
            "<th style='padding:8px 12px;text-align:left'>Bando</th>"
            "<th style='padding:8px 12px;text-align:center'>Giorni rimasti</th>"
            "<th style='padding:8px 12px;text-align:center'>Link</th>"
            "</tr>{}</table></div>"
        ).format(len(urgenti), righe_urgenti)
    else:
        box_urgenti = (
            "<div style='background:#f0fdf4;border:1px solid #86efac;border-radius:10px;"
            "margin:20px 0;padding:16px;color:#166534'>"
            "Nessun bando in scadenza imminente questa settimana."
            "</div>"
        )

    # ── Tabella bandi aperti ──────────────────────────────
    aperti = [b for b in bandi if b.get("status") == "APERTO"][:6]
    righe_bandi = ""
    for b in aperti:
        righe_bandi += (
            "<tr style='border-bottom:1px solid #e5e7eb'>"
            "<td style='padding:8px 12px;font-weight:600;font-size:13px'>{}</td>"
            "<td style='padding:8px 12px;color:#6b7280;font-size:12px'>{}</td>"
            "<td style='padding:8px 12px;color:#1e3a5f;font-weight:700;font-size:13px'>{}</td>"
            "<td style='padding:8px 12px'>"
            "<a href='{}' style='color:#2563eb;font-size:12px'>Apri</a>"
            "</td></tr>"
        ).format(b["titolo"], b["regione"], b["contributo"], b.get("link", "#"))

    # ── Statistiche ───────────────────────────────────────
    stats_html = ""
    for val, label, col in [
        (dati["totale"], "Bandi totali", "#1e3a5f"),
        (dati["aperti"], "Aperti", "#059669"),
        (dati["in_scadenza"], "In scadenza", "#d97706"),
        (dati["verificati"], "Verificati", "#7c3aed"),
    ]:
        stats_html += (
            "<div style='flex:1;background:#f8fafc;border-radius:8px;padding:14px;"
            "text-align:center;border:1px solid #e5e7eb;margin:0 4px'>"
            "<div style='font-size:22px;font-weight:800;color:{}'>{}</div>"
            "<div style='font-size:11px;color:#6b7280;margin-top:3px'>{}</div>"
            "</div>"
        ).format(col, val, label)

    # ── HTML finale ───────────────────────────────────────
    html = """<!DOCTYPE html>
<html>
<head><meta charset="utf-8"></head>
<body style="font-family:Arial,sans-serif;max-width:700px;margin:0 auto;color:#1a1a2e">

<div style="background:linear-gradient(135deg,#0f1c35,#1e3a5f);padding:30px;
border-radius:12px 12px 0 0;text-align:center">
  <h1 style="color:#fff;margin:0;font-size:24px">BandiRadar</h1>
  <p style="color:rgba(255,255,255,0.6);margin:6px 0 0;font-size:13px">
    Riepilogo settimanale — {data}
  </p>
</div>

<div style="background:#fff;padding:28px;border:1px solid #e5e7eb;border-top:none">
  <div style="display:flex;gap:8px;margin-bottom:24px">{stats}</div>

  {urgenti}

  <h3 style="color:#1e3a5f;margin:24px 0 12px">Bandi aperti</h3>
  <table style="width:100%;border-collapse:collapse">
    <tr style="background:#f8fafc">
      <th style="padding:8px 12px;text-align:left;font-size:12px">Bando</th>
      <th style="padding:8px 12px;text-align:left;font-size:12px">Regione</th>
      <th style="padding:8px 12px;text-align:left;font-size:12px">Contributo</th>
      <th></th>
    </tr>
    {righe}
  </table>

  <div style="margin-top:24px;text-align:center">
    <a href="{xlsx}" style="display:inline-block;background:linear-gradient(135deg,#f59e0b,#ef4444);
    color:#fff;padding:12px 28px;border-radius:8px;font-weight:700;text-decoration:none;margin-right:10px">
      Scarica Excel
    </a>
    <a href="{web}" style="display:inline-block;background:#1e3a5f;color:#fff;padding:12px 28px;
    border-radius:8px;font-weight:700;text-decoration:none">
      Vedi tutti i bandi
    </a>
  </div>

  <p style="margin-top:20px;font-size:11px;color:#9ca3af;text-align:center">
    Dataset aggiornato al {aggiornato}. Verificare le fonti ufficiali prima di presentare ai clienti.
  </p>
</div>
</body>
</html>""".format(
        data=TODAY.strftime("%d/%m/%Y"),
        stats=stats_html,
        urgenti=box_urgenti,
        righe=righe_bandi,
        xlsx=link_xlsx,
        web=link_web,
        aggiornato=aggiornato,
    )

    subject = "BandiRadar {} — {} bandi aperti".format(
        TODAY.strftime("%d/%m/%Y"), dati["aperti"]
    )
    if urgenti:
        subject = "⚠️ {} in scadenza | {}".format(len(urgenti), subject)

    return subject, html


def invia_email(subject, html):
    if not RESEND_API_KEY or not EMAIL_TO:
        print("RESEND_API_KEY o EMAIL_TO non impostati — skip email")
        return False
    try:
        r = httpx.post(
            "https://api.resend.com/emails",
            headers={
                "Authorization": "Bearer {}".format(RESEND_API_KEY),
                "Content-Type": "application/json",
            },
            json={"from": EMAIL_FROM, "to": [EMAIL_TO], "subject": subject, "html": html},
            timeout=30,
        )
        if r.status_code in (200, 201):
            print("Email inviata a {}".format(EMAIL_TO))
            return True
        else:
            print("Errore email: {} — {}".format(r.status_code, r.text))
            return False
    except Exception as e:
        print("Errore invio email: {}".format(e))
        return False


def main():
    print("--- Notifica email {} ---".format(TODAY.strftime("%d/%m/%Y")))
    dati = carica_dati()
    if not dati:
        return 0
    subject, html = build_email(dati)
    print("Subject: {}".format(subject))
    invia_email(subject, html)
    return 0


if __name__ == "__main__":
    sys.exit(main())
