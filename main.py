
from fastapi import FastAPI, APIRouter
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
import pandas as pd
import aggregator

app = FastAPI()
api_router = APIRouter()

# Esegui l'aggregazione dei dati all'avvio dell'applicazione
aggregator.aggregate_data()

@api_router.get("/data")
def get_data():
    try:
        df = pd.read_csv("aggregated_data.csv")
        if df.empty:
            return []
        # Sostituisci NaN con stringhe vuote per la compatibilità JSON
        df = df.fillna('')
        return df.to_dict(orient="records")
    except FileNotFoundError:
        return []  # Ritorna una lista vuota se il file non esiste

app.include_router(api_router, prefix="/api")

# Monta la directory corrente come directory per i file statici
app.mount("/", StaticFiles(directory=".", html=True), name="static")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
