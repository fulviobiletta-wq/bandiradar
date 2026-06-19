# BandiRadar v2 📋

Monitor automatico di bandi e contributi pubblici per PMI italiane.
**Zero costi. Aggiornamento automatico ogni lunedì.**

---

## Come funziona

```
Ogni lunedì alle 8:00 (automatico)
→ GitHub Actions scarica bandi dai siti ufficiali
→ Aggiorna stato (aperto/in scadenza/chiuso) da data reale
→ Genera Excel + JSON aggiornati
→ Pubblica su GitHub Pages (link fisso pubblico)
→ Invia email con riepilogo settimanale
```

**Link pubblico fisso:** `https://TUO-USERNAME.github.io/bandi-radar/`
**Excel scaricabile:** `https://TUO-USERNAME.github.io/bandi-radar/BandiRadar.xlsx`

---

## Setup (15 minuti, una sola volta)

### 1. Crea il repository GitHub
```bash
git init
git add .
git commit -m "BandiRadar v2 setup"
# Crea repo su github.com (pubblico o privato)
git remote add origin https://github.com/TUO-USERNAME/bandi-radar.git
git push -u origin main
```

### 2. Abilita GitHub Pages
- Vai su **Settings → Pages**
- Source: **GitHub Actions**
- Salva

### 3. Aggiungi i Secret (Settings → Secrets → Actions)

| Secret | Valore | Obbligatorio |
|--------|--------|--------------|
| `RESEND_API_KEY` | Chiave da resend.com (gratis) | Per email |
| `EMAIL_TO` | es. commerciale@azienda.it | Per email |
| `EMAIL_FROM` | es. BandiRadar <noreply@tuodominio.com> | Per email |

**Per l'email:** vai su [resend.com](https://resend.com) → Sign up gratis → API Keys → crea chiave.
Il piano free include 3.000 email/mese (4 email/mese per questo bot = sempre gratis).

### 4. Esegui il primo aggiornamento manuale
- Vai su **Actions → BandiRadar — Aggiornamento Settimanale**
- Clicca **Run workflow**
- Attendi ~2 minuti
- Il link pubblico è attivo!

---

## Struttura

```
bandi-radar/
├── aggiorna_bandi.py     # Scraper + generatore Excel/JSON
├── notify.py             # Email settimanale
├── requirements.txt
├── data/
│   └── bandi_base.json   # Bandi verificati manualmente ← aggiorna qui
├── output/               # Excel + JSON generati (auto)
├── docs/                 # GitHub Pages (auto)
└── web/
    └── index.html        # Pagina pubblica
```

---

## Aggiungere/aggiornare bandi manualmente

Modifica `data/bandi_base.json` e fai `git push`.
Il formato è:
```json
{
  "titolo": "Nome del bando",
  "ente": "Ente erogatore",
  "regione": "Regione Lombardia",
  "settore": "Digitalizzazione / Transizione 4.0",
  "contributo": "Fino a 50.000 euro",
  "intensita": "50% fondo perduto",
  "scadenza": "2026-12-31",
  "descrizione": "Breve descrizione",
  "link": "https://...",
  "note": "Note operative",
  "fonte": "VERIFICATO"
}
```

---

## Costi

| Voce | Costo |
|------|-------|
| GitHub (repo + Actions + Pages) | **€0** |
| Resend email (4/mese) | **€0** (free fino a 3.000/mese) |
| Server/hosting | **€0** |
| **TOTALE** | **€0/mese** |
