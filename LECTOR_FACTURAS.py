import streamlit as st
import streamlit.components.v1 as components
import json
import requests
import base64
import os
import time
import pandas as pd
from PIL import Image

# --- 1. CONFIGURACIÓN DE SEGURIDAD Y PÁGINA NATIVA ---
api_key_servidor = os.getenv("GEMINI_API_KEY")

# Carga segura de la imagen local para la pestaña
logo_path = "logo.png"
if os.path.exists(logo_path):
    icon_image = Image.open(logo_path)
else:
    icon_image = "🇵🇾"

st.set_page_config(
    page_title="Lector Facturas PY",
    page_icon=icon_image,
    layout="wide"
)

# --- ENLACES DE RECURSOS (CORREGIDO) ---
raw_logo_url = "https://githubusercontent.com"

# --- INYECCIÓN DE MANIFEST Y ESTILOS AVANZADOS PARA EL IPHONE ---
st.markdown(
    f"""
    <link rel="manifest" href="https://githubusercontent.com">
    <link rel="apple-touch-icon" sizes="180x180" href="{raw_logo_url}">
    <meta name="apple-mobile-web-app-capable" content="yes">
    <meta name="apple-mobile-web-app-status-bar-style" content="#004a99">
    <meta name="apple-mobile-web-app-title" content="Facturas PY">
    <meta name="apple-mobile-web-app-image" content="{raw_logo_url}">
    
    <style>
        /* Rediseño de fondo de la aplicación */
        .stApp {{
            background-color: #f8f9fa;
        }}
        /* Botones principales modernos */
        .stButton>button {{
            width: 100%;
            border-radius: 8px;
            height: 3.2em;
            background-color: #004a99;
            color: white;
            font-weight: bold;
            border: none;
            box-shadow: 0 4px 6px rgba(0,74,153,0.15);
            transition: all 0.3s ease;
        }}
        .stButton>button:hover {{
            background-color: #003366;
            transform: translateY(-1px);
        }}
        /* Tarjetas de facturas pulidas */
        .invoice-card {{
            background-color: white;
            padding: 22px;
            border-radius: 12px;
            border-left: 6px solid #004a99;
            box-shadow: 0 4px 12px rgba(0,0,0,0.04);
            margin-bottom: 20px;
        }}
        /* Contenedores modernos del área principal */
        .main-hero-box {{
            background: linear-gradient(135deg, #004a99 0%, #002244 100%);
            color: white;
            padding: 30px;
            border-radius: 16px;
            margin-bottom: 25px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.08);
        }}
    </style>
    """,
    unsafe_allow_html=True
)

if "lista_resultados" not in st.session_state:
    st.session_state.lista_resultados = []

# --- 2. ÁREA PRINCIPAL REDISEÑADA (ESTÉTICA) ---
st.markdown(
    """
    <div class="main-hero-box">
        <h1 style='margin:0; font-size: 2.2rem; color: white;'>Lector de Facturas e Ítems Detallados</h1>
        <p style='margin:8px 0 0 0; opacity: 0.85; font-size: 1.05rem;'>Extracción automatizada de cabeceras y desglose de mercaderías para contabilidad paraguaya</p>
    </div>
    """, 
    unsafe_allow_html=True
)

# --- 3. BARRA LATERAL CON LOGOTIPO VISUAL ---
with st.sidebar:
    # Agrega el logotipo arriba en el menú lateral para mejorar el aspecto visual
    if os.path.exists(logo_path):
        st.image(icon_image, width=120)
        
    st.markdown("<h2 style='margin-top:10px;'>⚙️ Configuración</h2>", unsafe_allow_html=True)
    if not api_key_servidor:
        api_key = st.text_input("Ingresar API Key manualmente", type="password")
    else:
        api_key = api_key_servidor
        st.success("✅ Servidor conectado de forma segura")

    st.markdown("---")
    if st.button("🗑️ Limpiar Memoria"):
        st.session_state.lista_resultados = []
        st.rerun()

# --- 4. PROCESAMIENTO CON DETALLE DE ÍTEMS ---
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
                    url_api = f"https://googleapis.com{api_key_servidor}"
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

    # --- 5. VISUALIZACIÓN Y EXPORTACIÓN ---
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