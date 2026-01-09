from io import BytesIO
import math
import pandas as pd
import datetime
from pptx import Presentation
from pptx.util import Inches, Pt


COLUMNAS_REPORTE = [
    "RUT Cliente",
    "Cliente",
    "Ejecutivo",
    "ID Deudor",
    "Deudor",
    "Fecha Otorgamiento",
    "Tipo Documento",
    "N°Documento",
    "Fecha Vencimiento",
    "Días Mora",
    "Monto Documento",
    "Monto Recaudado",
    "Capital Amortizado",
    "Monto Saldo",
]

FILAS_POR_SLIDE = 10  # se mantiene igual
SLIDE_WIDTH = Inches(13.33)
SLIDE_HEIGHT = Inches(7.5)

TABLE_WIDTH = Inches(12.3)
TABLE_HEIGHT = Inches(5.2)


def generar_ppt(excel_bytes: bytes) -> bytes:
    # 1️⃣ Leer Excel
    df = pd.read_excel(BytesIO(excel_bytes))

    columnas_faltantes = [c for c in COLUMNAS_REPORTE if c not in df.columns]
    if columnas_faltantes:
        raise ValueError(f"Faltan columnas: {', '.join(columnas_faltantes)}")

    df = df[COLUMNAS_REPORTE]

    # 2️⃣ Abrir PPT (con portada)
    prs = Presentation("app/template.pptx")

    # 3️⃣ Completar info del cliente en la portada
    cliente = str(df.iloc[0]["Cliente"])
    rut_cliente = str(df.iloc[0]["RUT Cliente"])

    portada = prs.slides[0]

    for shape in portada.shapes:
        if shape.name == "info_cliente" and shape.has_text_frame:
            tf = shape.text_frame
            tf.clear()

            p1 = tf.paragraphs[0]
            p1.text = f"{cliente}\n{rut_cliente}"
            p1.font.bold = True
            p1.font.size = Pt(20)

    # 4️⃣ Crear slides con tablas
    total_slides = math.ceil(len(df) / FILAS_POR_SLIDE)

    for slide_idx in range(total_slides):
        inicio = slide_idx * FILAS_POR_SLIDE
        fin = inicio + FILAS_POR_SLIDE
        df_slice = df.iloc[inicio:fin]

        slide = prs.slides.add_slide(prs.slide_layouts[5])  # blank

        filas = len(df_slice) + 1
        columnas = len(COLUMNAS_REPORTE)

        left = (SLIDE_WIDTH - TABLE_WIDTH) // 2
        top = (SLIDE_HEIGHT - TABLE_HEIGHT) // 2

        table = slide.shapes.add_table(
            filas,
            columnas,
            left,
            top,
            TABLE_WIDTH,
            TABLE_HEIGHT,
        ).table

        # Encabezados
        for col_idx, col_name in enumerate(COLUMNAS_REPORTE):
            cell = table.cell(0, col_idx)
            cell.text = col_name

            p = cell.text_frame.paragraphs[0]
            p.font.bold = True
            p.font.size = Pt(10)

        # Datos
        for row_idx, row in enumerate(df_slice.itertuples(index=False), start=1):
            for col_idx, value in enumerate(row):
                cell = table.cell(row_idx, col_idx)

                if pd.isna(value):
                    texto = ""
                elif COLUMNAS_REPORTE[col_idx] == "Cliente" or COLUMNAS_REPORTE[col_idx] == "Deudor":
                    texto = str(value).title()
                elif COLUMNAS_REPORTE[col_idx] in ["Fecha Otorgamiento", "Fecha Vencimiento"]:
                    if isinstance(value, (pd.Timestamp, datetime)):
                        texto = value.strftime("%d-%m-%Y")
                    else:
                        texto = str(value)
                else:
                    texto = str(value)

                cell.text = texto
                cell.text_frame.paragraphs[0].font.size = Pt(9)

    # 5️⃣ Guardar en memoria
    output = BytesIO()
    prs.save(output)
    output.seek(0)

    return output.read()
