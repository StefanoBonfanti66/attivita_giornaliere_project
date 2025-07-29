
from fastapi import FastAPI
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
import pandas as pd

app = FastAPI()

# Monta la directory corrente come directory per i file statici
app.mount("/", StaticFiles(directory=".", html=True), name="static")

# Le route specifiche per i file principali (index.html, manifest.json, service-worker.js)
# non sono più strettamente necessarie se html=True è impostato su StaticFiles,
# ma le mantengo per chiarezza o per override specifici.

@app.get("/")
def read_root():
    return FileResponse('index.html')

@app.get("/api/data")
def get_data():
    df = pd.read_csv("aggregated_data.csv")
    # Replace NaN with empty strings for JSON compatibility
    df = df.fillna('')
    return df.to_dict(orient="records")

@app.get("/manifest.json")
def get_manifest():
    return FileResponse('manifest.json')

@app.get("/service-worker.js")
def get_sw():
    return FileResponse('service-worker.js')
