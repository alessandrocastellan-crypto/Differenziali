# Web app schede differenziali

Questa è una prima versione operativa della web app per:
- caricare uno ZIP con schede `.xls` / `.xlsx`
- estrarre Blocco, Piano, Nome Quadro, Reparto
- estrarre gli interruttori con Circuito, Tipo differenziale e Dati nominali
- modificare manualmente i dati da browser
- esportare Excel
- generare PDF stampabile
- salvare/importare il database JSON

## Avvio locale

Serve Python installato sul PC o su un server.

```bash
pip install -r requirements.txt
streamlit run app.py
```

Poi apri il link che compare nel browser.

## Uso senza installare programmi sul PC

Carica questa cartella su un servizio tipo:
- Streamlit Community Cloud
- Render
- Hugging Face Spaces
- server aziendale interno

Il file principale è `app.py`.

## Note importanti

Questa versione contiene un estrattore tollerante ma generico. 
Se alcune schede hanno layout particolare, puoi correggere i dati direttamente nell'editor.
Dopo una prova reale, si può migliorare l'estrattore in base ai casi non letti correttamente.
