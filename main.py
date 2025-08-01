from fastapi import FastAPI, APIRouter
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
import pandas as pd
import aggregator
import os

# --- Configurazione CORS ---
# Definiamo le origini (i siti web) che possono fare richieste al nostro server.
# È una misura di sicurezza fondamentale.
origins = [
    "https://stefanobonfanti66.github.io", # Il tuo frontend su GitHub Pages
    "http://localhost",
    "http://localhost:8000",
    "http://127.0.0.1",
    "http://127.0.0.1:8000",
]
# --- Fine Configurazione ---

app = FastAPI()

# Aggiungiamo il middleware CORS all'applicazione.
# Questo permette al frontend su GitHub di comunicare con il backend sulla VPS.
app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Eseguiamo l'aggregazione dei dati all'avvio del server.
aggregator.aggregate_data()

# Creiamo un router per separare le logiche dell'API.
api_router = APIRouter()

@api_router.get("/data")
def get_data():
    try:
        # Usiamo un percorso assoluto per il file CSV per evitare problemi sulla VPS.
        base_dir = os.path.dirname(os.path.abspath(__file__))
        csv_path = os.path.join(base_dir, "aggregated_data.csv")
        
        df = pd.read_csv(csv_path)
        if df.empty:
            return []
        df = df.fillna('')
        return df.to_dict(orient="records")
    except FileNotFoundError:
        return []

# Includiamo le rotte dell'API nell'app principale, con il prefisso /api.
app.include_router(api_router, prefix="/api")

# Montiamo i file statici per ultimo.
# Questo risolve il conflitto che causava l'errore 404.
static_files_path = os.path.dirname(os.path.abspath(__file__))
app.mount("/", StaticFiles(directory=static_files_path, html=True), name="static")


if __name__ == "__main__":
    import uvicorn
    # Usiamo la porta 8000 come da configurazione finale.
    uvicorn.run(app, host="0.0.0.0", port=8000)