from fastapi import FastAPI
import ExportarDados

app = FastAPI()

@app.get("/exportar")
def exportar(SELECT,WHERE,FROM,GROUP,ORDER:str):

    resultado = ExportarDados.executar(SELECT,WHERE,FROM,GROUP,ORDER:str)

    return {
        "status": "ok",
        "resultado": resultado
    }
