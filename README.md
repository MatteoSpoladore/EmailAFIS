# Mail Merge App - Excel & Word Sender

Applicazione per inviare email personalizzate a partire da un file Excel, utilizzando template Word per oggetto e corpo email. Include funzionalità di creazione rapida di template, anteprima messaggi, modalità test e log invii. Supporta placeholder `{{NomeColonna}}` mappati alle colonne Excel.

---

## Caratteristiche

- Interfaccia grafica moderna con **CustomTkinter** e tema scuro.
- Caricamento **Excel** come sorgente dati.
- Caricamento **Word** come template (oggetto + corpo).
- Possibilità di creare rapidamente template Word ed Excel predefiniti.
- Supporto per email **HTML** con grassetto, corsivo, sottolineato e link.
- Anteprima della prima email e visualizzazione dei campi disponibili.
- Modalità TEST: invio solo all’utente per verifica.
- Log dettagliato degli invii in `email_log.txt`.
- Placeholder `{{NomeColonna}}` sostituiti con i valori delle colonne Excel.

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

## Funzionamento dell’app

1. Avvia l’app con:

```bash
python EmailApp.py
```

2. **Carica un file Excel** contenente le email e i dati dei destinatari.
3. **Carica un template Word** già pronto o creane uno nuovo tramite l’app.
   - Il primo paragrafo del Word viene usato come oggetto dell’email.
   - Il resto come corpo.
   - Puoi usare i placeholder `{{NomeColonna}}` per sostituire i dati personalizzati.

4. **Visualizza campi disponibili** con il pulsante “Mostra Campi”.
5. **Anteprima della prima email** per verificare il contenuto e i placeholder.
6. Abilita la **modalità TEST** per inviare solo a te.
7. Premi **Invia Email** per iniziare l’invio.
8. Controlla il log in `email_log.txt` per verificare invii riusciti e errori.

---

## Suggerimenti

- Verifica che i placeholder nel Word corrispondano esattamente ai nomi delle colonne Excel.
- Per email con formattazione (grassetto, corsivo, link), il corpo deve essere scritto in **HTML** nel Word.
- Per grandi invii, considera i limiti del provider SMTP o l’invio a batch.
- Mantieni aggiornato il file `.env` con credenziali sicure e proteggi le password.

---

## Template di esempio

Oggetto:

```
Promemoria pagamento {{AnnoCorso}}
```

Corpo HTML:

```html
<p>Gentili Genitori,</p>

<p>
  Con la presente desideriamo ricordarvi il pagamento della retta relativa al
  secondo trimestre del corso di musica {{AnnoCorso}} frequentato da vostro/a
  figlio/a.
</p>

<p>
  <b>L’importo dovuto è pari a {{Prezzo}} €</b> e può essere versato tramite
  bonifico bancario:
</p>
<ul>
  <li><b>IBAN: IT25L0863165011066000001528</b></li>
  <li><b>Beneficiario: Associazione Filarmonica Sanvitese</b></li>
</ul>

<p>
  Cogliamo inoltre l’occasione per ricordare che è necessario rinnovare la quota
  associativa e assicurativa, per un importo complessivo di
  <u>{{QuotaAssociativa}}</u> €.
</p>

<p>
  Tali quote dovranno essere versate presso la segreteria nei giorni di
  mercoledì e venerdì, dalle ore 16.30 alle ore 18.30, con pagamento in contanti
  entro il mese di {{MesePagamento}} {{AnnoPagamento}}.
</p>

<hr />

<p>
  <b>Cordiali saluti,</b><br />
  <b>Segreteria AFIS</b><br />
  <i>E. Pitton</i>
</p>

<p>_____________________________</p>

<p>
  <b>Filarmonica Sanvitese APS</b><br />
  Piazzale Hermann Zotti, 1<br />
  33078 San Vito Al Tagliamento (PN)<br />
  Tel.: 3396771611<br />
  <a href="https://www.afisanvitese.it">www.afisanvitese.it</a><br />
  <a href="mailto:scuola@afisanvitese.it">scuola@afisanvitese.it</a>
</p>
```
