from fastapi import FastAPI
import ExportarDados

app = FastAPI()

@app.get("/exportar")
def exportar():

    resultado = ExportarDados.executar()

    return {
        "status": "ok",
        "resultado": resultado
    }
