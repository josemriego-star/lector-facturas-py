import base64
import os
import re
import requests
import json
from io import BytesIO
from typing import List, Any
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from openpyxl import Workbook, load_workbook
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

# Planilla Excel original del "Registro de producción diaria". Debe estar en la misma
# carpeta que main.py (en el repositorio). El Excel exportado se arma copiando esta planilla.
PLANTILLA_FORMULARIO = os.path.join(os.path.dirname(os.path.abspath(__file__)), "plantilla_formulario.xlsx")


def a_numero(valor, defecto=0.0):
    """Convierte a número sin romper si viene vacío o con texto."""
    try:
        if valor is None or valor == "":
            return defecto
        return float(valor)
    except (ValueError, TypeError):
        return defecto


# ---------------------------------------------------------------------------
# Prompts
# ---------------------------------------------------------------------------

JSON_FACTURA = """{
    "es_factura": true,
    "tipo_comprobante": "FACTURA o TICKET",
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
}"""

JSON_NO_FACTURA = """{
    "es_factura": false,
    "tipo_documento": "que tipo de documento o imagen parece ser",
    "texto_detectado": "transcripcion completa y ordenada de todo el texto legible, seguida de una breve descripcion de lo que se ve"
}"""

JSON_FORMULARIO = """{
    "es_factura": false,
    "es_formulario": true,
    "fecha": "YYYY-MM-DD",
    "responsable": "nombre de la cocinera / responsable",
    "lugar": "Brasur, Coresa, AS, Cantina u Otro",
    "lugar_otro": "texto escrito junto a Otro (vacio si no aplica)",
    "hoja": "numero de hoja",
    "hoja_total": "total de hojas",
    "preparaciones": [
        {
            "numero": 1,
            "nombre": "nombre de la preparacion",
            "porciones_producidas": 0,
            "porciones_vendidas": 0,
            "precio_venta": 0,
            "ingredientes": [
                {
                    "codigo": "",
                    "ingrediente": "nombre del ingrediente",
                    "cantidad": 0,
                    "unidad": "gr, kg, lt, ml, un o mazo",
                    "observacion": ""
                }
            ],
            "hueso_g": 0,
            "grasa_g": 0,
            "sobro": "texto de sobrantes"
        }
    ]
}"""

REGLAS_COMPROBANTE = """
REGLAS IMPORTANTES (valen para facturas y también para TICKETS de compra):
- Un TICKET de compra (de balanza, caja registradora o comercio chico, a veces sin RUC, sin timbrado y sin número de factura) también se considera comprobante de compra: es_factura = true y tipo_comprobante = "TICKET". Si es una factura formal, tipo_comprobante = "FACTURA".
- emisor: el nombre del comercio que figura en el encabezado (ej. "FIAMBRERIA BRIGIDO").
- ruc y timbrado: si NO figuran en el comprobante, dejalos como "" (vacío). No inventes datos.
- Si en la foto se ven bordes o pedazos de OTRO papel/ticket distinto (por ejemplo un ticket pegado debajo o arriba con otro nombre o RUC), IGNORALOS: leé solamente el comprobante principal.
- nro_factura: si no hay número de factura pero hay número de operación (ej. "OP:00001"), usá "TICKET-00001". Si no hay ninguno, "".
- Ítems por peso (ej. "0,690 kg   47500 GS/kg   32775"): cantidad = 0.69, unidad_medida = "kg", precio_unitario = 47500, total_item = 32775.
- Ítems por unidad (ej. "1 un x 83000   83000"): cantidad = 1, unidad_medida = "unidad", precio_unitario = 83000, total_item = 83000.
- Transcribí la descripción tal como está impresa, sin corregir abreviaturas.
- Números: devolvelos como número, sin separador de miles (ej. 218775). En Paraguay la coma es separador decimal y el punto separa miles.
- iva: si el comprobante no indica la tasa de IVA, usá "10%".
- fecha: convertí cualquier formato a YYYY-MM-DD (ej. "02/OCT/26" -> "2026-10-02").
- condicion: "CONTADO" si no indica otra cosa.
- total: el TOTAL impreso en el comprobante."""

REGLAS_FORMULARIO = """
Es la foto de la planilla impresa "REGISTRO DE PRODUCCIÓN DIARIA", completada A MANO. Tiene esta estructura:
- Encabezado: FECHA (día / mes / año), COCINERA / RESPONSABLE, LUGAR / CLIENTE (casillas: Brasur, Coresa, AS, Cantina, Otro: ____), HOJA N° ("__ de __").
- Hasta 2 bloques "PREPARACIÓN N° 1" y "PREPARACIÓN N° 2". Cada bloque tiene: nombre de la preparación, porciones producidas, porciones vendidas, precio de venta por porción (opcional).
- Debajo, una tabla de hasta 9 ingredientes: N°, CÓD. (opcional), INGREDIENTE, CANTIDAD, UNIDAD (una X marcada bajo una de estas columnas: gr, kg, lt, ml, un, mazo) y OBSERVACIÓN.
- Al final de cada bloque: "SUBPRODUCTOS / SOBRANTES: Hueso __ g, Grasa __ g, Sobró: __".

REGLAS:
- Transcribí lo escrito a mano lo más fielmente posible. No inventes datos: si algo está vacío, dejalo vacío ("" o null).
- Omití las filas de ingredientes que estén totalmente vacías.
- "unidad" debe ser exactamente una de: gr, kg, lt, ml, un, mazo, según la columna donde está la X. Si no hay X pero la unidad está escrita al lado de la cantidad, usá esa.
- "lugar" es la casilla marcada. Si marcaron "Otro", poné "Otro" y escribí el texto en lugar_otro.
- Cantidades y porciones como número con punto decimal (ej. 1.5), sin separador de miles.
- fecha en formato YYYY-MM-DD. Si el año no está escrito, asumí 2026.
- Si hay una sola preparación completada, devolvé solo esa en la lista."""


def construir_prompt(modo: str, multipagina: bool) -> str:
    intro = "las siguientes imágenes, que son MÚLTIPLES PARTES/PÁGINAS de UN MISMO documento" if multipagina else "esta imagen"
    combinar = ("Combina la información de todas las imágenes en un único resultado (por ejemplo, si el encabezado está en una foto "
                "y el detalle continúa en otra, unificá todo en una sola lista y un solo total).") if multipagina else ""

    if modo == "otro":
        return (f"Analiza {intro} y responde SOLO con un JSON puro (sin texto adicional ni markdown), con este formato exacto:\n"
                f"{JSON_NO_FACTURA}\n"
                "No analices ítems de factura, RUC, IVA ni totales de comprobante: es solo transcripción y descripción del documento.\n"
                f"{combinar}")

    if modo == "formulario":
        return (f"Analiza {intro}. {REGLAS_FORMULARIO}\n\n"
                f"Responde SOLO con un JSON puro (sin texto adicional ni markdown), con este formato exacto:\n{JSON_FORMULARIO}\n"
                f"{combinar}")

    if modo == "factura":
        return (f"Analiza {intro}, que es una FACTURA o TICKET de compra de Paraguay, y responde SOLO con un JSON puro "
                f"(sin texto adicional ni markdown), con este formato exacto:\n{JSON_FACTURA}\n"
                f"{REGLAS_COMPROBANTE}\n{combinar}")

    # modo "auto"
    return (f"Analiza {intro} y responde SOLO con un JSON puro (sin texto adicional ni markdown).\n{combinar}\n\n"
            "PASO 1: Determina qué es: (A) factura o ticket de compra de Paraguay, (B) el formulario impreso "
            "\"REGISTRO DE PRODUCCIÓN DIARIA\" completado a mano, o (C) cualquier otro documento.\n\n"
            f"CASO A - FACTURA O TICKET, responde con este formato exacto:\n{JSON_FACTURA}\n{REGLAS_COMPROBANTE}\n\n"
            f"CASO B - FORMULARIO DE PRODUCCIÓN.{REGLAS_FORMULARIO}\nResponde con este formato exacto:\n{JSON_FORMULARIO}\n\n"
            f"CASO C - OTRO DOCUMENTO, responde con este formato:\n{JSON_NO_FACTURA}")


def extraer_json(texto: str):
    t = texto.replace("```json", "").replace("```", "").strip()
    i, j = t.find("{"), t.rfind("}")
    if i != -1 and j > i:
        t = t[i:j + 1]
    return json.loads(t)


@app.get("/")
async def inicio():
    return {"estado": "activo"}


@app.post("/procesar")
async def procesar_factura(files: List[UploadFile] = File(...), modo: str = Form("auto")):
    """
    Acepta una o varias imágenes (páginas) de UN MISMO documento y las analiza juntas
    en una sola llamada a Gemini.

    "modo":
      - "factura": comprobante de compra (factura o ticket)
      - "formulario": planilla "Registro de producción diaria" completada a mano
      - "otro": transcripción simple de cualquier otro documento
      - "auto" (por defecto): Gemini decide
    """
    if not api_key_servidor:
        raise HTTPException(status_code=500, detail="API Key de Gemini no configurada en el servidor.")

    if not files:
        raise HTTPException(status_code=400, detail="No se recibió ningún archivo.")

    try:
        url_api = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-3.6-flash:generateContent?key={api_key_servidor}"

        prompt = construir_prompt(modo, len(files) > 1)

        parts = [{"text": prompt}]
        for f in files:
            contenido = await f.read()
            img_b64 = base64.b64encode(contenido).decode("utf-8")
            parts.append({"inline_data": {"mime_type": f.content_type or "image/jpeg", "data": img_b64}})

        payload = {
            "contents": [{"parts": parts}],
            "generationConfig": {"responseMimeType": "application/json"},
        }

        response = requests.post(url_api, json=payload, timeout=90)
        res_json = response.json()

        if "candidates" in res_json:
            partes = res_json["candidates"][0]["content"]["parts"]
            texto = "".join(p.get("text", "") for p in partes if not p.get("thought"))
            return extraer_json(texto)
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
            if not isinstance(factura, dict):
                continue
            # Los archivos que no eran facturas (ni formularios) no van a este Excel
            if factura.get("es_factura") is False:
                continue

            fecha = factura.get("fecha", "")
            nro_factura = factura.get("nro_factura", "")
            proveedor = factura.get("emisor", "")

            for item in factura.get("items", []) or []:
                cantidad = a_numero(item.get("cantidad"), 1.0)
                descripcion = item.get("descripcion", "")
                unidad = item.get("unidad_medida", "unidad")
                subtotal_c_iva = a_numero(item.get("total_item"), 0.0)

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


# ---------------------------------------------------------------------------
# Exportar formulario de producción (mismo formato que la planilla original)
# ---------------------------------------------------------------------------

COLUMNA_UNIDAD = {"gr": "E", "kg": "F", "lt": "G", "ml": "H", "un": "I", "mazo": "J"}
ALIAS_UNIDAD = {
    "g": "gr", "gr": "gr", "grs": "gr", "gramo": "gr", "gramos": "gr",
    "kg": "kg", "kgs": "kg", "kilo": "kg", "kilos": "kg",
    "l": "lt", "lt": "lt", "lts": "lt", "litro": "lt", "litros": "lt",
    "ml": "ml", "cc": "ml",
    "un": "un", "u": "un", "unid": "un", "unidad": "un", "unidades": "un",
    "mazo": "mazo", "mazos": "mazo",
}
# Celdas de cada bloque de preparación en la planilla (1° y 2°)
BLOQUES = [
    {"nombre": "A8", "prod": "D8", "vend": "G8", "precio": "J8", "fila0": 11, "sub": "A20"},
    {"nombre": "A24", "prod": "D24", "vend": "G24", "precio": "J24", "fila0": 27, "sub": "A36"},
]
LUGARES = ["Brasur", "Coresa", "AS", "Cantina"]


def num_o_texto(v):
    """Número si se puede (entero sin .0), si no el texto; None si está vacío."""
    if v is None or v == "":
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return int(v) if float(v).is_integer() else v
    s = str(v).strip().replace(",", ".")
    try:
        f = float(s)
        return int(f) if f.is_integer() else f
    except ValueError:
        return str(v).strip()


def fecha_dmy(f):
    s = str(f or "").strip()
    m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})$", s)
    return f"{int(m.group(3)):02d}/{int(m.group(2)):02d}/{m.group(1)}" if m else s


def texto(v):
    return "" if v is None else str(v).strip()


def llenar_formulario(ws, d: dict):
    """Completa una hoja copiada de la planilla con los datos leídos."""
    fecha = fecha_dmy(d.get("fecha"))
    if fecha:
        ws["A4"] = fecha
    if texto(d.get("responsable")):
        ws["C4"] = texto(d.get("responsable"))

    # Lugar / cliente: marca la casilla correspondiente
    lugar = texto(d.get("lugar"))
    otro = texto(d.get("lugar_otro"))
    sel = next((l for l in LUGARES if l.lower() == lugar.lower()), None)
    if sel is None and lugar and lugar.lower() != "otro":
        otro = otro or lugar  # un lugar que no está en la lista va como "Otro"
    es_otro = sel is None and bool(lugar or otro)
    if lugar or otro:
        marca = lambda activo: "☑" if activo else "☐"
        partes = [f"{marca(sel == l)} {l}" for l in LUGARES]
        partes.append(f"{marca(es_otro)} Otro: {otro if otro else '________'}")
        ws["D4"] = "   ".join(partes)

    hoja, total = texto(d.get("hoja")), texto(d.get("hoja_total"))
    if hoja or total:
        ws["K4"] = f"{hoja or '_____'} de {total or '_____'}"

    for bloque, prep in zip(BLOQUES, d.get("preparaciones") or []):
        if not isinstance(prep, dict):
            continue
        for clave, celda in (("nombre", "nombre"), ("porciones_producidas", "prod"),
                             ("porciones_vendidas", "vend"), ("precio_venta", "precio")):
            v = prep.get(clave)
            v = texto(v) if clave == "nombre" else num_o_texto(v)
            if v not in (None, ""):
                ws[bloque[celda]] = v

        for i, ing in enumerate((prep.get("ingredientes") or [])[:9]):
            if not isinstance(ing, dict):
                continue
            fila = bloque["fila0"] + i
            if texto(ing.get("codigo")):
                ws[f"B{fila}"] = texto(ing.get("codigo"))
            ws[f"C{fila}"] = texto(ing.get("ingrediente"))
            cant = num_o_texto(ing.get("cantidad"))
            if cant is not None:
                ws[f"D{fila}"] = cant
            obs = texto(ing.get("observacion"))
            unidad_raw = texto(ing.get("unidad"))
            unidad = ALIAS_UNIDAD.get(unidad_raw.lower().rstrip("."))
            if unidad:
                celda_x = ws[f"{COLUMNA_UNIDAD[unidad]}{fila}"]
                celda_x.value = "X"
                celda_x.alignment = Alignment(horizontal="center", vertical="center")
            elif unidad_raw:
                obs = f"{obs} (unidad: {unidad_raw})".strip()
            if obs:
                ws[f"K{fila}"] = obs
            # Mismo estilo que el resto de la planilla (Arial 9)
            for col in "BCDEFGHIJK":
                c = ws[f"{col}{fila}"]
                c.font = Font(name="Arial", size=9)
                c.alignment = Alignment(horizontal="left" if col in "CK" else "center", vertical="center")

        hueso, grasa, sobro = num_o_texto(prep.get("hueso_g")), num_o_texto(prep.get("grasa_g")), texto(prep.get("sobro"))
        if hueso is not None or grasa is not None or sobro:
            ws[bloque["sub"]] = (
                f"SUBPRODUCTOS / SOBRANTES:   Hueso {hueso if hueso is not None else '________'} g    "
                f"Grasa {grasa if grasa is not None else '________'} g    "
                f"Sobró: {sobro or '____________________'}"
            )


@app.post("/exportar-formulario")
async def exportar_formulario(datos_formularios: List[Any]):
    """Recibe los formularios leídos y devuelve un Excel con una hoja por formulario,
    copiando la planilla original (mismo diseño, celdas combinadas y formato)."""
    try:
        formularios = [f for f in datos_formularios if isinstance(f, dict) and f.get("es_formulario")]
        if not formularios:
            raise HTTPException(status_code=400, detail="No hay formularios para exportar.")
        if not os.path.exists(PLANTILLA_FORMULARIO):
            raise HTTPException(status_code=500, detail="Falta el archivo plantilla_formulario.xlsx en el servidor.")

        wb = load_workbook(PLANTILLA_FORMULARIO)
        base = wb.active
        # Primero se copian las hojas vacías y recién después se llenan
        hojas = [base] + [wb.copy_worksheet(base) for _ in formularios[1:]]
        for i, (ws, datos) in enumerate(zip(hojas, formularios), 1):
            ws.title = f"Hoja {i}"
            llenar_formulario(ws, datos)

        stream = BytesIO()
        wb.save(stream)
        stream.seek(0)
        return StreamingResponse(
            stream,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=formulario_produccion.xlsx"}
        )
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
