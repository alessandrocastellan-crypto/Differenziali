# Gestionale schede differenziali

App web per gestire il database delle schede differenziali.

## Funzioni

- Carica database Excel `.xlsx`
- Cerca per sede, blocco, quadro, reparto
- Modifica intestazione scheda
- Modifica righe interruttori
- Crea nuove schede
- Duplica/elimina schede
- Scarica database Excel aggiornato
- Genera PDF stampabile di una scheda o di tutte le schede

## Pubblicazione su Streamlit Cloud

Carica su GitHub questi file:
- `app.py`
- `requirements.txt`
- `README.md`

Poi su Streamlit Cloud:
- Repository: il tuo repository GitHub
- Branch: `main`
- Main file path: `app.py`

## Uso

1. Apri la web app
2. Carica il database Excel
3. Modifica o crea schede
4. Premi `Salva modifiche`
5. Scarica il database aggiornato
6. Scarica il PDF per stampare

## Nota

Questa app non conserva automaticamente i dati online.
Per non perdere modifiche, scarica sempre il database aggiornato dopo il lavoro.
