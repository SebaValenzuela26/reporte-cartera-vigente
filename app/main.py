from fastapi import FastAPI, UploadFile, File
from fastapi.responses import HTMLResponse
from fastapi.templating import Jinja2Templates
from fastapi.requests import Request
from fastapi.responses import StreamingResponse
from io import BytesIO
from app.ppt_generator import generar_ppt

app = FastAPI()

# Templates
templates = Jinja2Templates(directory="app/templates")

@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse(
        "index.html",
        {"request": request}
    )

@app.post("/generar-ppt")
async def generar_ppt_endpoint(file: UploadFile = File(...)):
    excel_bytes = await file.read()
    ppt_bytes = generar_ppt(excel_bytes)

    return StreamingResponse(
        BytesIO(ppt_bytes),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={
            "Content-Disposition": "attachment; filename=reporte_cartera_vigente.pptx"
        },
    )