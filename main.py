import base64
import os
import requests
import json
from io import BytesIO
from typing import List, Any
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

app = FastAPI(title="API Lector Facturas PY")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

api_key_servidor = os.getenv("GEMINI_API_KEY")


@app.get("/")
async def inicio():
    return {"estado": "activo"}


@app.post("/procesar")
async def procesar_factura(file: UploadFile = File(...)):
    if not api_key_servidor:
        raise HTTPException(status_code=500, detail="API Key de Gemini no configurada en el servidor.")

    try:
        contents = await file.read()
        img_b64 = base64.b64encode(contents).decode("utf-8")

        url_api = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-3.6-flash:generateContent?key={api_key_servidor}"

        prompt = """Extrae la información de esta factura física de Paraguay en formato JSON exacto:
        {
            "ruc": "sin dv",
            "emisor": "razon social",
            "fecha": "YYYY-MM-DD",
            "timbrado": "numero",
            "nro_factura": "000-000-0000000",
            "items": [
                {
                    "cantidad": 1,
                    "descripcion": "detalle del producto o servicio",
                    "unidad_medida": "litro, kg, unidad, etc.",
                    "precio_unitario": 0,
                    "total_item": 0,
                    "iva": "10%, 5% o EXENTA"
                }
            ],
            "total": 0,
            "condicion": "CONTADO o CREDITO"
        }
        Responde únicamente el JSON puro, sin texto adicional ni formateo markdown."""

        payload = {
            "contents": [{
                "parts": [
                    {"text": prompt},
                    {"inline_data": {"mime_type": file.content_type or "image/jpeg", "data": img_b64}}
                ]
            }]
        }

        response = requests.post(url_api, json=payload, timeout=60)
        res_json = response.json()

        if "candidates" in res_json:
            texto = res_json["candidates"][0]["content"]["parts"][0]["text"]
            texto_limpio = texto.replace("```json", "").replace("```", "").strip()
            return json.loads(texto_limpio)
        else:
            error_msg = res_json.get("error", {}).get("message", "Sin respuesta de Gemini")
            raise HTTPException(status_code=400, detail=error_msg)

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/exportar-excel")
async def exportar_excel(datos_facturas: List[Any]):
    try:
        wb = Workbook()
        ws = wb.active
        ws.title = "Facturas Detalladas"

        columnas = [
            "Fecha", "Factura", "Proveedor", "Descripción Original (Factura)",
            "Cantidad", "Unidad de Medida", "Subtotal (Gs.) c/IVA",
            "Tasa IVA (%)", "IVA (Gs.)", "Subtotal sin IVA (Gs.)"
        ]
        ws.append(columnas)

        font_header = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        fill_header = PatternFill(start_color="004a99", end_color="004a99", fill_type="solid")
        alignment_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        thin_side = Side(border_style="thin", color="D3D3D3")
        border_cell = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

        for col_num, _ in enumerate(columnas, 1):
            cell = ws.cell(row=1, column=col_num)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = alignment_center

        for factura in datos_facturas:
            fecha = factura.get("fecha", "")
            nro_factura = factura.get("nro_factura", "")
            proveedor = factura.get("emisor", "")

            for item in factura.get("items", []):
                cantidad = float(item.get("cantidad", 1) or 1)
                descripcion = item.get("descripcion", "")
                unidad = item.get("unidad_medida", "unidad")
                subtotal_c_iva = float(item.get("total_item", 0) or 0)

                tasa_str = str(item.get("iva", "10")).replace("%", "").strip().upper()
                if "EXENTA" in tasa_str or tasa_str == "0":
                    tasa = 0
                    iva = 0.0
                    subtotal_s_iva = subtotal_c_iva
                elif tasa_str == "5":
                    tasa = 5
                    iva = subtotal_c_iva / 21.0
                    subtotal_s_iva = subtotal_c_iva - iva
                else:
                    tasa = 10
                    iva = subtotal_c_iva / 11.0
                    subtotal_s_iva = subtotal_c_iva - iva

                ws.append([
                    fecha, nro_factura, proveedor, descripcion,
                    cantidad, unidad, subtotal_c_iva, tasa, round(iva, 2), round(subtotal_s_iva, 2)
                ])

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=10):
            for cell in row:
                cell.font = Font(name="Arial", size=10)
                cell.border = border_cell
                if cell.column in (1, 2, 5, 6, 8):
                    cell.alignment = Alignment(horizontal="center")
                if cell.column in (7, 9, 10):
                    cell.number_format = "#,##0"
                    cell.alignment = Alignment(horizontal="right")

        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            ws.column_dimensions[col[0].column_letter].width = max(max_len + 3, 12)

        stream = BytesIO()
        wb.save(stream)
        stream.seek(0)

        return StreamingResponse(
            stream,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=facturas_detalladas_py.xlsx"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
