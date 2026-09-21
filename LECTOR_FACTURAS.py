import streamlit as st
import json
import requests
import base64
import os
import time
import pandas as pd
from PIL import Image

# --- 1. CONFIGURACIÓN DE SEGURIDAD Y PÁGINA ---
api_key_servidor = os.getenv("GEMINI_API_KEY")

# Carga segura de la imagen local para la pestaña
logo_path = "logo.png"
if os.path.exists(logo_path):
    icon_image = Image.open(logo_path)
else:
    icon_image = "🇵🇾"

# --- CONFIGURACIÓN DE PÁGINA E ICONO PWA PARA IPHONE ---
st.set_page_config(
    page_title="Lector Facturas PY",
    page_icon="logo.png",
    layout="wide"
)

# URL del icono en GitHub con parametro para romper la caché de Safari
apple_icon_url = "https://raw.githubusercontent.com/josemriego-star/lector-facturas-py/main/apple-touch-icon.png?v=2"

st.markdown(
    f"""
    <head>
        <!-- Icono para iPhone y iPad -->
        <link rel="apple-touch-icon" href="{apple_icon_url}">
        <link rel="apple-touch-icon-precomposed" href="{apple_icon_url}">
        
        <!-- Icono para navegador y Android -->
        <link rel="shortcut icon" href="{apple_icon_url}">
        <link rel="icon" type="image/png" href="{apple_icon_url}">
    </head>
    """,
    unsafe_allow_html=True
)

if "lista_resultados" not in st.session_state:
    st.session_state.lista_resultados = []

st.markdown("""
    <style>
    .stButton>button {
        width: 100%;
        border-radius: 5px;
        height: 3em;
        background-color: #004a99;
        color: white;
        font-weight: bold;
    }
    .invoice-card {
        background-color: white;
        padding: 20px;
        border-radius: 10px;
        border-left: 5px solid #004a99;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        margin-bottom: 15px;
    }
    </style>
    """, unsafe_allow_html=True)

st.title("Lector de Facturas e Ítems Detallados")
st.caption("Extracción de cabecera y desglose de mercaderías/servicios para contabilidad")

# --- 2. BARRA LATERAL ---
with st.sidebar:
    st.header("⚙️ Configuración")
    if not api_key_servidor:
        api_key = st.text_input("Ingresar API Key manualmente", type="password")
    else:
        api_key = api_key_servidor
        st.success("✅ Servidor conectado de forma segura")

    if st.button("🗑️ Limpiar Memoria"):
        st.session_state.lista_resultados = []
        st.rerun()

# --- 3. PROCESAMIENTO CON DETALLE DE ÍTEMS ---
if api_key:
    uploaded_files = st.file_uploader("Subir o capturar facturas", type=['jpg', 'jpeg', 'png'], accept_multiple_files=True)

    if uploaded_files:
        if st.button("🚀 PROCESAR FACTURAS"):
            st.session_state.lista_resultados = []
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            for idx, file in enumerate(uploaded_files):
                try:
                    status_text.text(f"Analizando {file.name} ({idx+1}/{len(uploaded_files)})...")
                    url_api = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-3.6-flash:generateContent?key={api_key_servidor}"
                    img_b64 = base64.b64encode(file.getvalue()).decode('utf-8')
                    
                    prompt = """Extrae la información de esta factura física de Paraguay en formato JSON exacto:
                    {
                        "ruc": "sin dv",
                        "emisor": "razon social",
                        "fecha": "DD/MM/YYYY",
                        "timbrado": "numero",
                        "nro_factura": "000-000-0000000",
                        "items": [
                            {
                                "cantidad": 1,
                                "descripcion": "detalle del producto o servicio",
                                "precio_unitario": 0,
                                "total_item": 0,
                                "iva": "10%, 5% o EXENTA"
                            }
                        ],
                        "gravada_10": 0,
                        "gravada_5": 0,
                        "exenta": 0,
                        "total": 0,
                        "condicion": "CONTADO o CREDITO"
                    }
                    Responde únicamente el JSON puro, sin texto adicional ni formateo markdown."""
                    
                    payload = {
                        "contents": [{
                            "parts": [
                                {"text": prompt},
                                {"inline_data": {"mime_type": "image/jpeg", "data": img_b64}}
                            ]
                        }]
                    }
                    
                    response = requests.post(url_api, json=payload)
                    res_json = response.json()

                    if 'candidates' in res_json:
                        texto = res_json['candidates'][0]['content']['parts'][0]['text']
                        texto_limpio = texto.replace('```json', '').replace('```', '').strip()
                        st.session_state.lista_resultados.append(json.loads(texto_limpio))
                    else:
                        st.error(f"Error en {file.name}: {res_json.get('error', {}).get('message', 'Sin respuesta')}")
                    
                    time.sleep(1)
                    progress_bar.progress((idx + 1) / len(uploaded_files))
                
                except Exception as e:
                    st.error(f"Fallo en {file.name}: {e}")
            
            status_text.text("✅ Procesamiento finalizado.")

    # --- 4. VISUALIZACIÓN Y EXPORTACIÓN ---
    if st.session_state.lista_resultados:
        res = st.session_state.lista_resultados
        
        st.subheader("📜 Facturas Procesadas")
        for r in res:
            with st.container():
                st.markdown(f"""
                <div class="invoice-card">
                    <h4>{r.get('emisor', 'S/N')} | Fecha: {r.get('fecha', 'S/F')}</h4>
                    <p><b>RUC:</b> {r.get('ruc')} | <b>Factura:</b> {r.get('nro_factura')} | <b>Condición:</b> {r.get('condicion')}</p>
                    <p><b>Total General:</b> {float(r.get('total', 0)):,.0f} ₲</p>
                </div>
                """, unsafe_allow_html=True)
                
                items = r.get('items', [])
                if items:
                    st.write("**Detalle de Ítems / Productos:**")
                    df_items = pd.DataFrame(items)
                    st.dataframe(df_items, use_container_width=True)
                st.divider()

        filas_csv = []
        for r in res:
            for item in r.get('items', [{}]):
                filas_csv.append({
                    "RUC": r.get('ruc'),
                    "EMISOR": r.get('emisor'),
                    "FECHA": r.get('fecha'),
                    "TIMBRADO": r.get('timbrado'),
                    "FACTURA": r.get('nro_factura'),
                    "CONDICION": r.get('condicion'),
                    "CANTIDAD": item.get('cantidad', 1),
                    "DESCRIPCION": item.get('descripcion', ''),
                    "PRECIO_UNITARIO": item.get('precio_unitario', 0),
                    "TOTAL_ITEM": item.get('total_item', 0),
                    "IVA_ITEM": item.get('iva', ''),
                    "TOTAL_FACTURA": r.get('total')
                })
        
        df_export = pd.DataFrame(filas_csv)
        csv_bytes = df_export.to_csv(index=False).encode('utf-8-sig')
        
        st.download_button(
            "📥 Descargar Excel con Ítems Detallados (.csv)", 
            csv_bytes, 
            "facturas_detalladas_py.csv", 
            "text/csv"
        )
else:
    st.info("👈 Requiere API Key conectada.")