from fastapi import FastAPI
from fastapi.responses import FileResponse
import ExportarDados

app = FastAPI()

@app.get("/exportar")
def exportar(SELECT: str, WHERE: str, FROM: str, GROUPBY: str, ORDERBY: str):

    arquivo = ExportarDados.executar(SELECT, WHERE, FROM, GROUPBT, ORDERBY)

    return FileResponse(
        path=arquivo,
        filename=arquivo,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
