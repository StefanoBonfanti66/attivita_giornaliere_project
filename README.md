# Attività Giornaliere Project by Stefano Bonfanti

Questo progetto fornisce una dashboard web per visualizzare e filtrare le attività giornaliere degli operatori, aggregando dati da file Excel.

## Struttura del Progetto

*   `main.py`: Il server FastAPI che serve la dashboard web e l'API per i dati.
*   `aggregator.py`: Script Python per l'aggregazione dei dati dai file Excel in un singolo file CSV.
*   `attivita_giornaliere.py`: Script Python locale per l'automazione dell'estrazione dati da un'applicazione Windows e la pre-elaborazione dei file Excel.
*   `index.html`: La pagina web frontend per la visualizzazione dei dati.
*   `requirements.txt`: Elenco delle dipendenze Python.
*   `render.yaml`: File di configurazione per il deploy automatico su Render.com.
*   `OpzioniEsportazione*.xlsx`: I file Excel originali contenenti i dati delle attività.
*   `aggregated_data.csv`: Il file CSV aggregato contenente i dati elaborati, utilizzato dalla dashboard.

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
4.  **Avvia il server (per la dashboard web):**
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

## Aggiornamento dei Dati e Funzionamento

Il flusso di aggiornamento dei dati e il funzionamento generale del progetto sono i seguenti:

1.  **Generazione Dati Locali (`attivita_giornaliere.py`):**
    *   Esegui lo script `attivita_giornaliere.py` sul tuo PC Windows.
    *   Questo script automatizza l'estrazione dei dati da un'applicazione esterna, genera i file `OpzioniEsportazione*.xlsx` e li pre-elabora (rimuovendo colonne non necessarie e formattando i fogli).
    *   **Nota:** I percorsi delle immagini (`logo.png`, `operatrice.png`) all'interno di questo script sono stati aggiornati per riflettere la nuova posizione del progetto.

2.  **Aggregazione e Sincronizzazione Git (`aggregator.py`):**
    *   Dopo la generazione dei file Excel, lo script `aggregator.py` viene eseguito (automaticamente da `attivita_giornaliere.py` o manualmente).
    *   `aggregator.py` legge tutti i fogli elaborati dai file `OpzioniEsportazione*.xlsx`, aggrega i dati in un singolo `aggregated_data.csv` e include correttamente le colonne `Operatore` e `Categoria` estratte dai nomi dei fogli.
    *   Lo script `aggregator.py` è ora più robusto e gestisce automaticamente il commit e il push di `aggregated_data.csv` su GitHub, anche se Git non rileva differenze di contenuto (forzando un commit vuoto se necessario).

    Nota utile: se vuoi che venga creata automaticamente una bozza di email Outlook con il file `aggregated_data.csv` allegato (solo su Windows con Outlook e `pywin32`), puoi eseguire `aggregator.py` con l'opzione `--email`. Esempio:

    ```bash
    python aggregator.py --email --email-to stefano@example.com --email-subject "Report giornaliero"
    ```

    Questo aprirà una bozza di Outlook allegando `aggregated_data.csv` (o stamperà un messaggio descrittivo se Outlook/pywin32 non sono disponibili).

3.  **Deploy Automatico su Render.com:**
    *   Ogni volta che `aggregated_data.csv` viene pushato su GitHub, Render rileva le modifiche e avvia un nuovo processo di deploy.
    *   **Importante:** La dashboard su Render ora serve direttamente il file `aggregated_data.csv` presente nel repository, senza tentare di ri-eseguire l'aggregazione all'avvio del server (questo previene la sovrascrittura del file con dati vuoti).

## Funzionalità della Dashboard Web (`index.html`)

La dashboard web è stata significativamente migliorata:

*   **Filtri Avanzati:**
    *   Sono stati aggiunti filtri per `Operatore`, `Categoria`, `Data Inizio` e `Data Fine`.
    *   Il layout dei filtri è stato ottimizzato per una migliore visualizzazione.
*   **Visualizzazione Controllata:**
    *   La dashboard non mostra più tutte le attività di default. Richiede la selezione di un periodo tramite i filtri per visualizzare i dati, migliorando le performance per grandi set di dati.
*   **Ordinamento Colonne Personalizzato:**
    *   Le colonne della tabella sono ora visualizzate in un ordine specifico per una migliore leggibilità: `Data`, `Soggetto`, `Ragione sociale`, `Contatto`, `Note interne 1`, `Operatore`, `Categoria`.
*   **Personalizzazione Grafica:**
    *   È stato aggiunto il logo aziendale nella barra di navigazione.
    *   È stata inserita un'immagine di marketing come sfondo nella sezione principale della dashboard, con il testo reso bianco per una migliore visibilità.

## Risoluzione Problemi Comuni

*   **Problemi di Proprietà Git (`dubious ownership`):** Se riscontri errori di proprietà Git, esegui il seguente comando nel tuo terminale nella directory del progetto:
    ```bash
    git config --global --add safe.directory C:/progetti_stefano/automations/attivita_giornaliere_project
    ```
*   **Immagini non visualizzate:** Assicurati che i file immagine (`logo.png`, `operatrice.png`) siano presenti nella directory radice del progetto e siano stati aggiunti e committati su GitHub.

    ## Preparare una email in Outlook (Windows)

    Se vuoi preparare rapidamente una bozza di email in Outlook con il file generato (es. `aggregated_data.csv`), ho aggiunto uno script utility `outlook_email.py` che crea una bozza e allega il file.

    Prerequisiti:

    * Windows + Outlook installato
    * `pywin32` (aggiunto a `requirements.txt`)

    Esempi:

    Aprire una bozza con allegato:

    ```bash
    python outlook_email.py --file aggregated_data.csv --subject "Report giornaliero" --body "In allegato il report." --to stefano@example.com
    ```

    Inviare direttamente (attenzione: invia l'email):

    ```bash
    python outlook_email.py --file aggregated_data.csv --subject "Report giornaliero" --body "In allegato il report." --to stefano@example.com --send
    ```

    Se non hai `pywin32` installalo con:

    ```bash
    pip install pywin32
    ```

    Nota: lo script mostra un messaggio d'errore utile se pywin32/Outlook non sono disponibili.

### Troubleshooting: errore COM / "Esecuzione del server non riuscito"

Se, come nel tuo output, vedi un errore tipo:

```
Impossibile creare la bozza Outlook: (-2146959355, 'Esecuzione del server non riuscito.', None, None)
```

Prova i seguenti passaggi:

1. Assicurati che Outlook sia installato e aperto con lo stesso utente con cui esegui lo script.
2. Avvia Outlook manualmente (apri l'app) e rilancia lo script; spesso il problema è che il server COM non è pronto.
3. Verifica che `pywin32` sia installato nel tuo ambiente Python (esegui `pip show pywin32`).
4. Evita di eseguire lo script come utente diverso (ad es. amministratore) rispetto a Outlook — i COM server possono rifiutare l'esecuzione tra utenti/elevazioni diverse.
5. Se il problema persiste, prova ad aggiornare lo script `outlook_email.py` (già incluso) che tenta più strategie: EnsureDispatch, DispatchEx e lancia Outlook se necessario. Se anche così fallisce, copia qui l'errore completo e lo esamino.

Questi consigli risolvono la maggior parte dei casi.
