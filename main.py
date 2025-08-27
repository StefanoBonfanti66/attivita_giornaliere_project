
import os
import pandas as pd
from fastapi import FastAPI, APIRouter
from fastapi.staticfiles import StaticFiles
# import aggregator # No longer needed if not calling aggregate_data()

# --- App Principale ---
app = FastAPI()

# --- Aggregazione Dati all'Avvio ---
# Eseguiamo l'aggregazione dei dati all'avvio del server.
# Questo assicura che i dati siano sempre aggiornati al deploy.
# REMOVED: aggregator.aggregate_data()

# --- API Router ---
# Definiamo le rotte per i dati in un router separato per pulizia.
api_router = APIRouter()

@api_router.get("/data")
def get_data():
    try:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        csv_path = os.path.join(base_dir, "aggregated_data.csv")
        
        df = pd.read_csv(csv_path)
        if df.empty:
            return []
        df = df.fillna('')
        return df.to_dict(orient="records")
    except FileNotFoundError:
        return []

# Includiamo il router dell'API nell'app principale.
app.include_router(api_router, prefix="/api")

# --- Montaggio File Statici ---
# Montiamo la directory corrente per servire i file statici (index.html, etc.).
# Questo deve essere l'ultimo montaggio per non interferire con le rotte API.
app.mount("/", StaticFiles(directory=".", html=True), name="static")

# --- Avvio Server (per sviluppo locale) ---
if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
