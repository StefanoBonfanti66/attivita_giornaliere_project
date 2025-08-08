# Attività Giornaliere Project by Stefano Bonfanti

Questo progetto fornisce una dashboard web per visualizzare e filtrare le attività giornaliere degli operatori, aggregando dati da file Excel.

## Struttura del Progetto

*   `main.py`: Il server FastAPI che serve la dashboard web e l'API per i dati.
*   `aggregator.py`: Script Python per l'aggregazione dei dati dai file Excel in un singolo file CSV.
*   `index.html`: La pagina web frontend per la visualizzazione dei dati.
*   `requirements.txt`: Elenco delle dipendenze Python.
*   `render.yaml`: File di configurazione per il deploy automatico su Render.com.
*   `OpzioniEsportazione*.xlsx`: I file Excel originali contenenti i dati delle attività.

## Esecuzione in Locale

Per eseguire l'applicazione sul tuo computer:

1.  **Clona il repository:**
    ```bash
    git clone https://github.com/StefanoBonfanti66/attivita_giornaliere_project.git
    cd attivita_giornaliere_project
    ```
2.  **Installa le dipendenze:**
    ```bash
    pip install -r requirements.txt
    ```
3.  **Assicurati di avere i file Excel:** Posiziona i file `OpzioniEsportazione*.xlsx` nella stessa directory del progetto.
4.  **Avvia il server:**
    ```bash
    python main.py
    ```
    L'applicazione sarà accessibile all'indirizzo `http://127.0.0.1:8000`.

## Deploy su Render.com

Questa applicazione è configurata per il deploy automatico su [Render.com](https://render.com/).

1.  **Connetti il tuo account GitHub a Render.**
2.  **Crea un nuovo "Web Service"** su Render, selezionando questo repository.
3.  Render utilizzerà il file `render.yaml` per configurare automaticamente il processo di build e deploy.
4.  L'applicazione sarà disponibile all'URL fornito da Render (es. `https://attivita-giornaliere-project.onrender.com`).

## Aggiornamento dei Dati

I dati vengono aggregati dai file Excel `OpzioniEsportazione*.xlsx` presenti nella directory del progetto.

1.  **Esegui lo script di automazione locale:** Continua a utilizzare il tuo script `attivita_giornaliere.py` sul tuo PC per generare i nuovi file Excel. Questo script aggiornerà anche il file `aggregated_data.csv` in locale.
2.  **Sincronizza i dati con GitHub:**
    *   Dopo aver generato i nuovi file Excel, assicurati che siano nella directory del progetto.
    *   Apri GitHub Desktop.
    *   Vedrai i nuovi file Excel e il file `aggregated_data.csv` modificati.
    *   Fai un **commit** e un **push** su GitHub.
3.  **Render.com aggiornerà automaticamente:** Ogni volta che fai un push su GitHub, Render rileverà le modifiche, rieseguirà il processo di build (che include l'aggregazione dei dati) e riavvierà l'applicazione con i dati aggiornati.
