# Collegare la scheda a Google Sheets

## 1. Crea il foglio Google Sheet
Crea un nuovo Google Sheet con **4 fogli (tab)** con questi nomi esatti:

### Tab "Giorni"
| Giorno | Etichetta | Sottotitolo | Tag |
|---|---|---|---|
| 1 | Giorno 1 | Gambe/Glutei A | legs |
| 2 | Giorno 2 | Spalle/Braccia | braccia |
| 3 | Giorno 3 | Gambe/Glutei B | legs |

(Tag può essere: `legs`, `braccia`, `addome` — determina solo il colore dell'etichetta)

### Tab "Esercizi"
| Giorno | Ordine | Nome | Serie | Target | Video |
|---|---|---|---|---|---|
| 1 | 1 | Squat con bilanciere | 4 | 10-12 rip | https://www.youtube.com/watch?v=ThD3d6kNjuA |
| 1 | 2 | Hip Thrust con bilanciere | 4 | 12-15 rip | https://www.youtube.com/watch?v=IHk9Qn8ttX8 |
| 1 | 3 | Leg Press 45° | 3 | 12-15 rip | https://www.youtube.com/watch?v=LMTyPl_oo38 |
| 1 | 4 | Affondi in camminata con manubri | 3 | 10 rip/gamba | https://www.youtube.com/watch?v=5IzLQWqiggg |
| 1 | 5 | Leg Curl (sdraiato/seduto) | 3 | 12-15 rip | https://www.youtube.com/watch?v=1zevKZn_n1E |
| 2 | 1 | Military Press manubri (seduta) | 4 | 10-12 rip | https://www.youtube.com/watch?v=n5emn5Pee5o |
| 2 | 2 | Alzate laterali con manubri | 3 | 12-15 rip | https://www.youtube.com/watch?v=6sT8LVeGVoc |
| 2 | 3 | Curl bicipiti con manubri | 3 | 12 rip | https://www.youtube.com/watch?v=QaXCjUMCp90 |
| 2 | 4 | Curl Hammer con manubri | 3 | 12 rip | https://www.youtube.com/watch?v=R_Afr1mMRi8 |
| 2 | 5 | French Press manubri su panca | 3 | 12-15 rip | https://www.youtube.com/watch?v=RLGF2C58pPI |
| 3 | 1 | Hip Thrust con bilanciere | 4 | 10-12 rip (carico alto) | https://www.youtube.com/watch?v=IHk9Qn8ttX8 |
| 3 | 2 | Leg Curl (sdraiato/seduto) | 4 | 12-15 rip | https://www.youtube.com/watch?v=1zevKZn_n1E |
| 3 | 3 | Squat con bilanciere | 3 | 12-15 rip | https://www.youtube.com/watch?v=ThD3d6kNjuA |
| 3 | 4 | Leg Press 45° | 3 | 12-15 rip | https://www.youtube.com/watch?v=LMTyPl_oo38 |
| 3 | 5 | Affondi in camminata con manubri | 3 | 12 rip/gamba | https://www.youtube.com/watch?v=5IzLQWqiggg |

**Per modificare la scheda in futuro**: aggiungi/rimuovi/modifica righe qui. `Giorno` deve corrispondere a un valore presente nel tab "Giorni". `Ordine` decide la posizione nella lista.

### Tab "Circuito"
| Ordine | Nome | Target | Video | TimerSecondi |
|---|---|---|---|---|
| 1 | Plank | 30-45 sec | https://www.youtube.com/watch?v=Is-7PPaBcsM | 40 |
| 2 | Crunch classico | 15-20 rip | https://www.youtube.com/watch?v=Iep9ffXKQHw | |
| 3 | Crunch a bicicletta | 15-20 rip/lato | https://www.youtube.com/watch?v=RC4wOdLdZr0 | |

Questi esercizi si ripetono automaticamente per 3 giri in ogni giorno. `TimerSecondi` vuoto = esercizio a ripetizioni (checkbox); valorizzato = timer automatico (come per il plank).

### Tab "Log"
Lascialo vuoto, solo intestazione sulla prima riga:
| Timestamp | Data | Giorno | Esercizio | Valore1 | Valore2 | Valore3 |
|---|---|---|---|---|---|---|

L'app scriverà qui i risultati ogni volta che premi "Sincronizza".

---

## 2. Aggiungi lo script (Apps Script)
Nel foglio: **Estensioni → Apps Script**. Cancella il contenuto di default e incolla questo codice:

```javascript
function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  function readSheet(name) {
    const sheet = ss.getSheetByName(name);
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    return data
      .filter(row => row[0] !== "" && row[0] !== null)
      .map(row => {
        const obj = {};
        headers.forEach((h, i) => obj[h] = row[i]);
        return obj;
      });
  }

  const payload = {
    giorni: readSheet('Giorni'),
    esercizi: readSheet('Esercizi'),
    circuito: readSheet('Circuito')
  };

  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('Log');
  const payload = JSON.parse(e.postData.contents);
  const ts = new Date();

  (payload.entries || []).forEach(function(entry) {
    logSheet.appendRow([
      ts, payload.date, payload.day,
      entry.exercise, entry.setsDone + '/' + entry.setsTotal + ' serie', entry.weight || '', ''
    ]);
  });

  (payload.circuito || []).forEach(function(c) {
    const dettagli = Object.keys(c).filter(k => k !== 'round').map(k => k + ': ' + (c[k] ? 'fatto' : '-')).join(', ');
    logSheet.appendRow([ts, payload.date, payload.day, 'Circuito giro ' + c.round, dettagli, '', '']);
  });

  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}
```

Salva il progetto (icona a forma di floppy disk, dai un nome tipo "API Scheda").

## 3. Pubblica come Applicazione Web
1. In alto a destra: **Deploy → Nuovo deployment**.
2. Icona a ingranaggio accanto a "Seleziona tipo" → **Applicazione web**.
3. Configura così:
   - **Esegui come**: Me (il tuo account)
   - **Chi ha accesso**: Chiunque
4. **Deploy**. La prima volta Google chiederà di autorizzare: clicca "Autorizza accesso" → scegli il tuo account → se compare "Google non ha verificato questa app", clicca "Avanzate" → "Vai al progetto (non sicuro)" → Consenti. (È il tuo script personale, è normale che compaia questo avviso.)
5. Copia l'**URL dell'app web** che termina con `/exec`.

## 4. Collega l'app HTML
Apri `index.html`, trova questa riga vicino all'inizio dello script:

```javascript
const SHEET_API_URL = "";
```

e incolla il tuo URL tra le virgolette:

```javascript
const SHEET_API_URL = "https://script.google.com/macros/s/XXXXXXXXX/exec";
```

Salva, carica su GitHub (sovrascrivendo il file), e ricarica la pagina/app sul telefono.

---

## Note pratiche
- Se in futuro modifichi lo script, devi rifare **Deploy → Gestisci deployment → modifica (icona matita) → Nuova versione → Deploy**, altrimenti l'URL pubblicato resta collegato alla versione vecchia.
- I dati di oggi (serie spuntate, pesi) restano salvati **sul telefono** anche senza connessione. Il tasto "Sincronizza" invece richiede internet e scrive lo storico nel foglio "Log".
- Se vuoi rivedere i tuoi progressi nel tempo, apri semplicemente il foglio "Log" su Google Sheets — è un elenco cronologico di tutte le sessioni sincronizzate.
