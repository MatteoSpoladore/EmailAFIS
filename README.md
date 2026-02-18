# EmailApp – Mail Merge Excel con Invio SMTP

EmailApp è un’applicazione desktop con interfaccia grafica (basata su `customtkinter`) che consente l’invio massivo di email personalizzate (mail merge) utilizzando un file Excel come sorgente dati.

È pensata per contesti amministrativi, scolastici, associativi o aziendali in cui sia necessario inviare comunicazioni strutturate e personalizzate a più destinatari, con eventuali allegati specifici per ciascun destinatario.

---

## Funzionalità principali

- Mail merge da file Excel (.xlsx)
- Placeholder dinamici nel formato `{{NomeColonna}}`
- Supporto corpo email in HTML o testo semplice
- Allegati personalizzati per ogni riga Excel
- Modalità TEST (invio singolo all’utente SMTP)
- Anteprima HTML della prima email
- Invio in thread separato (interfaccia non bloccante)
- Logging dettagliato su file locale
- Tema chiaro/scuro selezionabile

---

## Requisiti

- Python 3.9+
- Dipendenze:

```bash
pip install customtkinter pandas python-dotenv python-docx openpyxl
```

---

## Configurazione SMTP

Le credenziali SMTP vengono caricate tramite variabili d’ambiente (consigliato uso di file `.env`):

```
SMTP_SERVER=smtp.example.com
SMTP_PORT=587
SMTP_USER=tuamail@email.com
SMTP_PASSWORD=tuapassword
USE_TLS=True
```

Note:

- `SMTP_USER` viene utilizzato anche come indirizzo mittente.
- `USE_TLS=True` abilita STARTTLS.
- La configurazione viene validata prima dell’invio.

---

## Struttura del file Excel

Il file deve essere in formato `.xlsx`.

Struttura esempio:

| Email                                   | Nome  | Corso | Prezzo | ALLEGATO |
| --------------------------------------- | ----- | ----- | ------ | -------- |
| [test@email.com](mailto:test@email.com) | Marco | Piano | 120    | file.pdf |

Regole:

- Prima colonna → email destinatario.
- Colonne successive → campi dinamici utilizzabili nei placeholder.
- Colonna opzionale `ALLEGATO` → solo nome file (non percorso completo).
- Ogni riga corrisponde a una email.

Se una cella è vuota (anche ALLEGATO), viene trattata come stringa vuota.

---

## Sistema di Placeholder

I placeholder devono usare doppie parentesi graffe:

```
{{NomeColonna}}
```

Esempio:

Oggetto:

```
Promemoria pagamento corso {{Corso}}
```

Corpo:

```
Gentile {{Nome}},
l’importo di {{Prezzo}} € è in scadenza.
```

Prima dell’invio:

1. I placeholder vengono estratti.
2. Viene verificata la corrispondenza con le colonne Excel.
3. In caso di campo mancante l’invio viene bloccato.
4. I valori vengono sostituiti riga per riga.

Se una cella contiene NaN o è vuota, viene convertita in stringa vuota.

---

## Logica di Composizione Email

Ogni email viene costruita come:

- `multipart/mixed` (contenitore esterno)
  - `multipart/alternative`
    - versione `text/plain`
    - versione `text/html`

  - allegato (se presente)

Sanitizzazione oggetto:

- I caratteri di ritorno a capo (`\r`) e nuova riga (`\n`) vengono rimossi.
- Questo impedisce vulnerabilità di header injection e garantisce che l’oggetto sia una singola riga valida.

Se nel corpo sono presenti tag HTML (es. `<p>`, `<b>`, `<a>`), viene inviato come HTML.
In ogni caso viene sempre generata anche una versione plain text.

---

## Gestione Allegati

Gli allegati sono opzionali e richiedono:

- Checkbox “Abilita allegati”
- Presenza della colonna `ALLEGATO`
- Selezione della cartella contenente i file

Comportamento per ogni riga:

- Se la cella ALLEGATO è vuota → email inviata senza allegato.
- Se è presente un nome file:
  - Il file viene cercato nella cartella selezionata.
  - Se trovato → allegato correttamente.
  - Se non trovato → errore registrato nel log, ma l’email viene comunque inviata.

È supportato un solo allegato per riga.

---

## Modalità TEST

Se attivata:

- Viene processata solo la prima riga del file.
- Il destinatario viene forzato a `SMTP_USER`.
- Utile per verificare formattazione e allegati prima dell’invio massivo.

---

## Anteprima

La funzione “Anteprima Prima Email”:

1. Usa la prima riga del file Excel.
2. Sostituisce i placeholder.
3. Genera un file HTML temporaneo.
4. Lo apre nel browser predefinito.
5. Lo elimina dopo un intervallo di tempo.

Se l’app viene chiusa bruscamente, alcuni file temporanei potrebbero rimanere nel sistema.

---

## Logging

Tutte le operazioni vengono registrate in:

```
%LOCALAPPDATA%/EmailApp/email_log.txt
```

Vengono tracciati:

- Errori di connessione SMTP
- Email non valide
- File allegati mancanti
- Errori di invio
- Invii corretti

---

## Gestione Errori

L’applicazione:

- Valida la configurazione SMTP
- Verifica i placeholder
- Valida il formato email
- Gestisce le eccezioni SMTP
- Continua l’elaborazione anche se una riga fallisce

Al termine viene mostrato un riepilogo:

- Email inviate
- Errori riscontrati

---

## Limitazioni

- Supporta solo file `.xlsx`
- Un solo allegato per riga
- Nessun meccanismo automatico di retry
- Nessun rate limiting
- Nessuna autenticazione OAuth

---

# Esempio di Workflow

### Scenario:

Una scuola di musica deve inviare promemoria di pagamento con PDF personalizzato a ogni genitore.

---

### 1 – Preparazione Excel

File `studenti.xlsx`:

| Email                                   | Nome  | Corso   | Prezzo | Mese  | ALLEGATO  |
| --------------------------------------- | ----- | ------- | ------ | ----- | --------- |
| [gen1@email.com](mailto:gen1@email.com) | Marco | Piano   | 150    | Marzo | marco.pdf |
| [gen2@email.com](mailto:gen2@email.com) | Anna  | Violino | 130    | Marzo | anna.pdf  |

Cartella allegati:

```
/allegati/
    marco.pdf
    anna.pdf
```

---

### 2 – Scrittura Template

Oggetto:

```
Promemoria pagamento – corso {{Corso}}
```

Corpo HTML:

```
<p>Gentile {{Nome}},</p>

<p>Le ricordiamo che il pagamento di <b>{{Prezzo}} €</b> relativo al corso di {{Corso}} è in scadenza nel mese di {{Mese}}.</p>

<p>In allegato trova il documento dettagliato.</p>

<p>Cordiali saluti,<br>
Segreteria</p>
```

---

### 3 – Attivazione Allegati

- Spuntare “Abilita allegati”
- Selezionare la cartella `/allegati/`

---

### 4 – Modalità Test

- Attivare TEST
- Inviare
- Verificare corretta ricezione

---

### 5 – Invio Massivo

- Disattivare TEST
- Cliccare “Invia Email”
- Monitorare barra di avanzamento
- Verificare riepilogo finale
- Controllare il file di log in caso di errori
