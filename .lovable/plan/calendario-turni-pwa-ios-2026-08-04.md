# Calendario Turni — PWA iOS

App a pagina singola (route `/`) in stile Apple Health / iOS Calendar: sfondo grigio chiaro, card bianche arrotondate, tipografia di sistema. Tutti i dati restano nel browser (nessun backend).

## Barra strumenti (in alto, sticky)

- **Carica CSV**: pulsante sempre visibile. Il file (separatore `;`) viene letto nel browser e sostituisce completamente il precedente, resettando la vista.
- **Data domenica di riferimento**: date picker in stile iOS. Da questa data l'app deriva i 7 giorni.
- **Menu a tendina voci**: default `Seleziona una voce...` (valore vuoto) all'avvio e a ogni nuovo CSV; nessun dato mostrato finché non si sceglie.
- **Pulsante “+ / Gestisci voci”**: apre un modal per aggiungere o eliminare manualmente le voci del menu. La lista è salvata in `localStorage` e sopravvive al riavvio.

## Estrazione dati

Filtra le righe dove la Colonna 2 è esattamente uguale alla voce scelta, poi mappa i blocchi da 4 colonne (1 colonna saltata tra un blocco e l'altro):

```text
Lunedì    -> col 8, 9, 10, 11
Martedì   -> col 13, 14, 15, 16
Mercoledì -> col 18, 19, 20, 21
Giovedì   -> col 23, 24, 25, 26
Venerdì   -> col 28, 29, 30, 31
Sabato    -> col 33, 34, 35, 36
Domenica  -> col 3, 4, 5, 6
```

Le date: la domenica scelta con il picker è la domenica della settimana; Lunedì = domenica −6g … oppure, se preferisci, +1g. **Assunzione adottata**: la settimana visualizzata va da Lunedì a Domenica e la domenica scelta è l'ultimo giorno (Lunedì = domenica −6 giorni), coerente con l'ordine “una settimana per riga da lunedì”.

Per ogni blocco non vuoto:
1. Ora Inizio = primo valore presente, Ora Fine = ultimo valore presente; mostrate in verticale (inizio sopra, fine sotto).
2. Cella “Totale ore” sotto gli orari: differenza inizio–fine, meno 1 ora se il turno supera le 8 ore (es. 8–18 = 10h → 9h).
3. Blocco completamente vuoto: la card mostra solo intestazione giorno + data, nessun testo, badge o totale.

## Layout

- Griglia di card verticali affiancate, 7 per riga (una settimana per riga) su schermo largo; su iPhone la riga si adatta con scroll/colonne compatte mantenendo l'ordine Lun→Dom.
- Card bianche, angoli arrotondati, ombre morbide, separatori sottili in stile iOS.

## PWA

- `public/manifest.webmanifest` con nome, display `standalone`, theme/background color e icone.
- Tag `<link rel="manifest">`, `theme-color` e `<link rel="apple-touch-icon">` nel root route, così “Aggiungi alla schermata Home” su Safari usa l'icona corretta.
- Solo installabilità (nessun service worker / offline, non richiesto).

## Icona

Icona in stile iOS: gradiente blu scuro → azzurro, al centro simbolo minimale “liquid glass” di calendario con orologio. Esportata in PNG 180×180 (apple-touch-icon) e 192/512 per il manifest; usata anche come favicon.

## Note tecniche

- Parsing CSV lato client (split su `;`, gestione righe/quote basilare), nessun upload al server.
- Stato in React: righe CSV, voce selezionata, data di riferimento; voci del menu persistite in `localStorage`.
- Colori e stile definiti come token in `src/styles.css`.
