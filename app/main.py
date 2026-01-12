from fastapi import FastAPI, UploadFile, File
from fastapi.responses import HTMLResponse
from fastapi.templating import Jinja2Templates
from fastapi.requests import Request
from fastapi.responses import FileResponse
from app.ppt_generator import generar_ppt, pptx_a_pdf

app = FastAPI()

templates = Jinja2Templates(directory="app/templates")

@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse(
        "index.html",
        {"request": request}
    )

@app.post("/generar-pdf")
async def generar_pdf_endpoint(file: UploadFile = File(...)):
    excel_bytes = await file.read()
    ppt_bytes = generar_ppt(excel_bytes)

    pdf_path = "reporte_cartera_vigente.pdf"
    pptx_a_pdf(ppt_bytes, pdf_path)

    return FileResponse(
        pdf_path,
        media_type="application/pdf",
        filename="reporte_cartera_vigente.pdf"
    )