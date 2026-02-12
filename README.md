# Mail Merge App - Excel & Word Sender

## Applicazione per inviare email personalizzate da Excel usando template Word per oggetto e corpo. Include creazione rapida di template e modalità test.

## Caratteristiche

- Interfaccia grafica **dark theme** con CustomTkinter.
- Caricamento **Excel** come sorgente dati.
- Caricamento **Word** come template (oggetto + corpo).
- Anteprima prima email e log invii.
- Modalità TEST: invia email solo a te.
- Creazione template Word/Excel (`Matrice_mail`) in un clic.
- Placeholder `{{NomeColonna}}` mappati alle colonne Excel.

---

## Requisiti

```bash
pip install customtkinter pandas python-dotenv openpyxl python-docx
```

- File `.env` con credenziali email:

```env
SMTP_SERVER=smtp.gmail.com
SMTP_PORT=587
SMTP_USER=tua_email@gmail.com
SMTP_PASSWORD=la_tua_password_app
USE_TLS=True
```

> Nota: Gmail richiede password per app.

---

## Inizializzazione UV

```bash
uv init .
uvi add EmailApp.py
uvi add .env
uvi add requirements.txt
uv commit -m "Prima versione Mail Merge App"
```

- `uv init .` → inizializza il progetto UV.
- `uvi add ...` → aggiunge file al controllo versione.
- `uv commit` → salva le modifiche nel repository locale.

---

## Uso

1. Avvia l’app:

```bash
python EmailApp.py
```

2. Carica Excel e Word (o crea template).
3. Verifica anteprima e campi disponibili.
4. Abilita modalità TEST se necessario.
5. Premi **Invia Email**.
6. Controlla log in `email_log.txt`.

---

## Suggerimenti

- Assicurati che i placeholder Word corrispondano ai nomi colonne Excel.
- Per grandi invii, valuta batch o limiti SMTP del provider.
