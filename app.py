import streamlit as st
import pandas as pd
import pypdf
import pdfplumber
import logging
import datetime
import re
import os
import io
import base64
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as OpenpyxlImage
from google.oauth2 import service_account
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer, Image as ReportLabImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from supabase import create_client, Client

# ---------------------------------------------------------
# Configuración inicial
# ---------------------------------------------------------
st.set_page_config(
    page_title="Asistente_web", 
    page_icon="📦", 
    layout="wide"
)

st.markdown("""
    <html lang="es" class="notranslate" translate="no">
    <head><meta name="google" content="notranslate" /></head>
    </html>
    <style>
        .stDataFrame, .stTable { width: 100% !important; overflow-x: auto !important; }
        div[data-testid="stDataFrame"] div, td, th { white-space: normal !important; word-wrap: break-word !important; }
        div.stFormSubmitButton > button {
            background-color: #1F3864 !important; color: white !important; font-size: 18px !important;
            font-weight: bold !important; padding: 1rem 2rem !important; border-radius: 10px !important;
            border: 3px solid #FFD700 !important; box-shadow: 0px 6px 12px rgba(0, 0, 0, 0.4) !important;
            width: 100% !important; cursor: pointer !important;
        }
        div.stFormSubmitButton > button:hover { background-color: #2c4d8c !important; }
    </style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# Parámetros y Constantes
# ---------------------------------------------------------
PRODUCT_SPREADSHEET_ID = "1aTlmA6JBldTX3zN-djDjWA5HEAExTPcdNhJsPJL9Kgo"
SPREADSHEET_ID = "1HatcJlMpdxk4Z92sFMjU_MPEwjqJT5oww171jPC2Gnw"
DRIVE_FOLDER_ID = "1Amwy8_uQgo6X0VS2DXH028Ep80BMi4rP"
CEDULA_DEV_CORRECTA = "1073513861"
MASTER_CSV_PATH = "registro_entregas.csv"
LOGO_PATH = "logo.png" if os.path.exists("logo.png") else ("download.png" if os.path.exists("download.png") else None)

# --- Logging ---
LOG_FILENAME = "log_ejecucion.txt"
logging.basicConfig(filename=LOG_FILENAME, level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s", force=True)

def registrar_log(mensaje, tipo="INFO"):
    print(f"[{tipo}] {mensaje}")
    if tipo == "INFO": logging.info(mensaje)
    elif tipo == "WARNING": logging.warning(mensaje)
    elif tipo == "ERROR": logging.error(mensaje)
    for handler in logging.root.handlers: handler.flush()

registrar_log("--- Sesión iniciada ---")

def get_client_ip():
    try:
        if hasattr(st, "context") and hasattr(st.context, "headers"):
            headers = st.context.headers
            ip = headers.get("X-Forwarded-For", headers.get("X-Real-IP", ""))
            if ip:
                if "," in ip: ip = ip.split(",")[0].strip()
                return ip
    except Exception: pass
    if "browser_session_id" not in st.session_state:
        import uuid
        st.session_state["browser_session_id"] = str(uuid.uuid4())
    return st.session_state["browser_session_id"]

def get_google_creds():
    scope = ["https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/spreadsheets"]
    if "google" in st.secrets:
        try:
            creds_dict = dict(st.secrets["google"])
            if "private_key" in creds_dict:
                pk = creds_dict["private_key"].strip().strip('"').strip("'")
                if "\\n" in pk: pk = pk.replace("\\n", "\n")
                creds_dict["private_key"] = pk
            return service_account.Credentials.from_service_account_info(creds_dict, scopes=scope)
        except Exception as e:
            registrar_log(f"Error cargando credenciales: {e}", "WARNING")
    if os.path.exists('credentials.json'):
        return ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
    return None


def descargar_pdf_desde_drive(folder_id, numero_entrega):
    try:
        creds = get_google_creds()
        if not creds:
            registrar_log("No se encontraron credenciales de Google Drive.", "ERROR")
            return None

        service = build('drive', 'v3', credentials=creds)
        query = f"('{folder_id}' in parents or name contains '{numero_entrega}') and mimeType='application/pdf' and trashed=false"
        results = service.files().list(
            q=query,
            pageSize=5,
            fields="files(id, name)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        ).execute()
        files = results.get('files', [])

        matched_file = None
        for f in files:
            if str(numero_entrega) in f['name']:
                matched_file = f
                break

        if not matched_file and files:
            matched_file = files[0]

        if not matched_file:
            registrar_log(f"No se encontró PDF en Drive para la entrega: {numero_entrega}", "WARNING")
            return None

        file_id = matched_file['id']
        file_name = matched_file['name']
        request = service.files().get_media(fileId=file_id)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()

        fh.seek(0)
        registrar_log(f"PDF descargado exitosamente desde Drive: {file_name}")
        return fh

    except Exception as e:
        registrar_log(f"Error al descargar PDF de Drive para {numero_entrega}: {e}", "ERROR")
        return None


def extract_pdf_full_text(pdf_file):
    try:
        with pdfplumber.open(pdf_file) as pdf:
            text = ""
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    text += t + "\n"
            return text
    except Exception:
        return ""


def extract_metadata_general(text):
    meta = {
        "No. Remisión": "No detectado",
        "Fecha Emisión": "No detectado",
        "Cliente / Empresa": "No detectado",
        "Ciudad Destino": "No detectado"
    }

    if (rem_m := re.search(r"(?:No\.|N°|Remisión|Entrega)\s*[:#]?\s*(\d{6,10})", text, re.IGNORECASE)):
        meta["No. Remisión"] = rem_m.group(1).strip()

    if (f_m := re.search(r"Fecha\s*[:]??\s*(\d{2}/\d{2}/\d{4})", text, re.IGNORECASE)):
        meta["Fecha Emisión"] = f_m.group(1).strip()

    banned_headers = ["ENTREGA", "MERCANCIA", "REMISIÓN", "ORDEN", "PEDIDOS", "DIRECCIÓN"]
    for line in text.split('\n'):
        if "CLIENTE" in line.upper() and ":" in line:
            val = line.split(":", 1)[1].strip()
            if len(val) > 3 and not any(b in val.upper() for b in banned_headers):
                meta["Cliente / Empresa"] = val
                break
    if meta["Cliente / Empresa"] == "No detectado":
        if (cli_m := re.search(r"CLIENTE\s*[:]??\s*([^\n]+)", text, re.IGNORECASE)):
            val = cli_m.group(1).strip()
            if not any(b in val.upper() for b in banned_headers):
                meta["Cliente / Empresa"] = val

    if (ciu_m := re.search(r"Ciudad\s*[:]??\s*([^\n]+)", text, re.IGNORECASE)):
        meta["Ciudad Destino"] = ciu_m.group(1).strip()

    return meta


def extract_observation_text(pdf_file):
    try:
        with pdfplumber.open(pdf_file) as pdf:
            first_page = pdf.pages[0]
            text = first_page.extract_text()

        if not text:
            return None

        pattern = r"Observación:\s*(.*?)(?=\n[A-Z]|$)"
        match = re.search(pattern, text, re.DOTALL | re.IGNORECASE)

        if match:
            obs_raw = match.group(1).replace("\n", " ").strip()
            return re.split(r'Basado en Pedidos', obs_raw, flags=re.IGNORECASE)[0].strip()

        return text[:300].replace("\n", " ")
    except Exception:
        return None


def parse_observation_data(text):
    data = {
        "Dirección": "No detectado",
        "Contacto": "No detectado",
        "Celular": "No detectado",
        "OC": "No detectado"
    }

    text_clean = text.replace("\n", " ").strip() if text else ""

    cel_match = re.search(r'\b(3\d{2}[\s\-]?\d{3}[\s\-]?\d{4}|\b3\d{9}\b|\b\d{10}\b)', text_clean)
    if cel_match:
        data["Celular"] = cel_match.group(1).strip()

    cont_match = re.search(r"CONTACTO[:]??\s*([A-ZÁÉÍÓÚÑ\s]+?)(?=\s*-\s*TELEFONO|\s*-\s*TEL|\s*CEL\.|\s*OC|\s*//|$)", text_clean, re.IGNORECASE)
    if cont_match:
        nombre = cont_match.group(1).strip()
        if len(nombre) > 2:
            data["Contacto"] = nombre

    dir_match = re.search(r"(?:ENTREGAR EN|DIRECCION DE ENVIO|DIRECCIÓN DE ENVÍO|DIRECCIÓN|DIRECCION)[:]??\s*(.*?)(?=\s*//|\s*CONTACTO|\s*HORARIO|\s*OC|$)", text_clean, re.IGNORECASE)
    if dir_match:
        dir_val = dir_match.group(1).strip()
        data["Dirección"] = re.split(r'Basado en Pedidos', dir_val, flags=re.IGNORECASE)[0].strip()
    else:
        data["Dirección"] = text_clean

    if data["Contacto"] != "No detectado":
        data["Contacto"] = data["Contacto"].replace("PRINCIPAL", "").replace("ALMACEN", "").strip()

    return data


def procesar_ubicacion(direccion):
    dir_upper = direccion.upper()
    mun = ["PUERTO LOPEZ", "PUERTO GAITAN", "BARRANQUILLA", "MEDELLIN", "ITAGUI", "COTA", "FUNZA", "MOSQUERA", "MADRID", "FACATATIVA", "CHIA", "CAJICA", "ZIPAQUIRA", "YUMBO", "CALI", "BOGOTA", "TOCANCIPA", "SOACHA", "PUENTE ARANDA"]
    muni_det = "NO ESPECIFICADO"
    for m in mun:
        if m in dir_upper:
            muni_det = m
            break
    if muni_det in ["FUNZA", "MOSQUERA", "MADRID", "FACATATIVA"]:
        reg = "SABANA OCCIDENTE"
    elif muni_det in ["COTA", "CHIA", "CAJICA", "ZIPAQUIRA", "TOCANCIPA"]:
        reg = "SABANA NORTE"
    elif muni_det in ["BOGOTA", "PUENTE ARANDA"]:
        reg = "NORTE" if any(z in dir_upper for z in ["SUBA", "USAQUEN", "CALLE 170"]) else "SUR"
    elif muni_det in ["YUMBO", "CALI", "BARRANQUILLA", "MEDELLIN", "ITAGUI"]:
        reg = "SUR"
    else:
        reg = "POR DEFINIR"
    return muni_det, reg


def generar_link_maps(direccion):
    if not direccion or direccion == "No detectado":
        return "#"
    return "https://www.google.com/maps/search/?api=1&query=" + direccion.strip().replace(" ", "+")


# Supabase Auth
@st.cache_resource
def init_supabase():
    url = st.secrets["supabase"]["url"]
    key = st.secrets["supabase"]["key"]
    return create_client(url, key)

supabase: Client = init_supabase()

if "autenticado" not in st.session_state: st.session_state.autenticado = False
if "usuario_actual" not in st.session_state: st.session_state.usuario_actual = None
if "empresa_actual" not in st.session_state: st.session_state.empresa_actual = None

def validar_en_supabase(correo, password):
    try:
        response = supabase.table("usuarios_licencias").select("*").eq("correo", correo.strip()).eq("password", password.strip()).execute()
        data = response.data
        if data and len(data) > 0 and data[0].get("activo") == True:
            return data[0]
        return None
    except Exception as e:
        st.error(f"Error de conexión: {e}")
        return None

def pantalla_login():
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.subheader("🚀 Optimiza Tus Procesos")
        st.caption("Ingrese sus credenciales autorizadas para acceder al sistema.")
        with st.form("form_login"):
            correo = st.text_input("Correo electrónico")
            password = st.text_input("Contraseña", type="password")
            if st.form_submit_button("Iniciar Sesión", use_container_width=True):
                user_data = validar_en_supabase(correo, password)
                if user_data:
                    client_token = get_client_ip()
                    try:
                        supabase.table("usuarios_licencias").update({"current_ip": client_token}).eq("correo", correo.strip()).execute()
                    except Exception: pass
                    st.session_state.autenticado = True
                    st.session_state.usuario_actual = user_data["correo"]
                    st.session_state.empresa_actual = user_data["empresa"]
                    st.session_state.login_token = client_token
                    st.success("¡Licencia verificada con éxito!")
                    st.rerun()
                else:
                    st.error("❌ Credenciales incorrectas o licencia inactiva.")

if not st.session_state.autenticado:
    pantalla_login()
    st.stop()

# ---------------------------------------------------------
# Capa de Datos: Google Sheets API con Respaldo Universal Local
# ---------------------------------------------------------
def limpiar_peso(val):
    if pd.isna(val): return 0.0
    val_str = str(val).strip().replace(',', '.').replace('KG', '').replace('kg', '').strip()
    try: return float(val_str)
    except ValueError: return 0.0

@st.cache_data(ttl=1800, show_spinner=False)
def fetch_google_sheet_database():
    try:
        creds = get_google_creds()
        if not creds: return None
        service = build('sheets', 'v4', credentials=creds)
        
        sheet_metadata = service.spreadsheets().get(spreadsheetId=PRODUCT_SPREADSHEET_ID).execute()
        sheets = sheet_metadata.get('sheets', [])
        
        list_dfs = []
        for s in sheets:
            sheet_name = s.get('properties', {}).get('title')
            try:
                result = service.spreadsheets().values().get(
                    spreadsheetId=PRODUCT_SPREADSHEET_ID,
                    range=f"{sheet_name}!A1:Z5000"
                ).execute()
                rows = result.get('values', [])
                if not rows or len(rows) < 2: continue
                
                df_h = pd.DataFrame(rows)
                df_h.columns = df_h.iloc[0].astype(str).str.strip()
                df_h = df_h.iloc[1:].copy()
                cols = list(df_h.columns)
                
                col_c = next((c for c in cols if c.lower() in ['codigo', 'código']), cols[1] if len(cols) > 1 else None)
                col_d = next((c for c in cols if 'desc' in c.lower()), cols[2] if len(cols) > 2 else None)
                col_p = next((c for c in cols if c.lower() in ['peso_kg', 'peso', 'kg']), cols[4] if len(cols) > 4 else (cols[-1] if len(cols) > 3 else None))

                if col_c:
                    temp_df = pd.DataFrame()
                    temp_df['Codigo'] = df_h[col_c]
                    temp_df['Descripcion'] = df_h[col_d] if col_d else "Sin descripción"
                    temp_df['Peso_KG'] = df_h[col_p] if col_p else 0.0
                    temp_df = temp_df.dropna(subset=['Codigo'])
                    list_dfs.append(temp_df)
            except Exception:
                continue
                
        if list_dfs:
            df_final = pd.concat(list_dfs, ignore_index=True)
            df_final['Codigo'] = df_final['Codigo'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
            df_final['Peso_KG'] = df_final['Peso_KG'].apply(limpiar_peso)
            return df_final
        return None
    except Exception as e:
        return None

def cargar_bd_local_universal():
    for filename in ["BD_Pesos.xlsx", "bd_pesos.csv"]:
        if os.path.exists(filename):
            try:
                if filename.endswith(".csv"):
                    df = pd.read_csv(filename, dtype=str)
                    df.columns = df.columns.astype(str).str.strip()
                    col_c = next((c for c in df.columns if c.lower() in ['codigo', 'código']), df.columns[1] if len(df.columns) > 1 else df.columns[0])
                    col_d = next((c for c in df.columns if 'desc' in c.lower()), df.columns[2] if len(df.columns) > 2 else None)
                    col_p = next((c for c in df.columns if c.lower() in ['peso_kg', 'peso', 'kg']), df.columns[4] if len(df.columns) > 4 else df.columns[-1])
                    
                    temp_df = pd.DataFrame()
                    temp_df['Codigo'] = df[col_c].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
                    temp_df['Descripcion'] = df[col_d].astype(str).str.strip() if col_d else "Sin descripción"
                    temp_df['Peso_KG'] = df[col_p].apply(limpiar_peso) if col_p else 0.0
                    return temp_df.dropna(subset=['Codigo'])
                else:
                    hojas_dict = pd.read_excel(filename, sheet_name=None, engine='openpyxl')
                    list_dfs = []
                    for nombre_hoja, df_hoja in hojas_dict.items():
                        if not df_hoja.empty:
                            df_h = df_hoja.copy()
                            df_h.columns = df_h.columns.astype(str).str.strip()
                            cols = list(df_h.columns)
                            col_c = next((c for c in cols if c.lower() in ['codigo', 'código']), cols[1] if len(cols) > 1 else None)
                            col_d = next((c for c in cols if 'desc' in c.lower()), cols[2] if len(cols) > 2 else None)
                            col_p = next((c for c in cols if c.lower() in ['peso_kg', 'peso', 'kg']), cols[4] if len(cols) > 4 else (cols[-1] if len(cols) > 3 else None))
                            if col_c:
                                temp_df = pd.DataFrame()
                                temp_df['Codigo'] = df_h[col_c]
                                temp_df['Descripcion'] = df_h[col_d] if col_d else "Sin descripción"
                                temp_df['Peso_KG'] = df_h[col_p] if col_p else 0.0
                                list_dfs.append(temp_df.dropna(subset=['Codigo']))
                    if list_dfs:
                        df_final = pd.concat(list_dfs, ignore_index=True)
                        df_final['Codigo'] = df_final['Codigo'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
                        df_final['Peso_KG'] = df_final['Peso_KG'].apply(limpiar_peso)
                        return df_final
            except Exception: pass
    return None

def get_product_data_from_source(codigo_input, df_bd):
    clean_input = str(codigo_input).strip().upper()
    if df_bd is not None:
        match_exact = df_bd[df_bd['Codigo'].astype(str).str.strip().str.upper() == clean_input]
        if not match_exact.empty: return match_exact.iloc[0]
        match_parcial = df_bd[df_bd['Codigo'].astype(str).str.contains(clean_input, case=False, na=False)]
        if not match_parcial.empty: return match_parcial.iloc[0]
    return None

def extraer_tabla_materiales(pdf_source, nombre_doc="Desconocido"):
    texto_completo = ""
    try:
        if hasattr(pdf_source, 'seek'): pdf_source.seek(0)
        with pdfplumber.open(pdf_source) as pdf:
            for page in pdf.pages:
                txt = page.extract_text()
                if txt: texto_completo += txt + "\n"
    except Exception: return []

    lineas = [l.strip() for l in texto_completo.split('\n') if l.strip()]
    items = []; excluidos = ['806014553', '806014553-6', '6658461', '8237779', '10109294']
    for linea_str in lineas:
        if any(term in linea_str.upper() for term in ['DECLARAMOS', 'GARANTÍA', 'RESPONSABLE', 'MERCANCÍA', 'IMPRESO', 'PÁGINA', 'CONDICIÓN', 'DIRECCIÓN', 'TELÉFONOS', 'ÍTEM', 'CÓDIGO', 'CANTIDAD', 'DESCRIPCIÓN']): continue
        if linea_str.startswith('#'): continue

        codigo = None; cant_val = None; desc_bruta = None
        match_b = re.search(r'^(?:\d+\s+)?([A-Z0-9\-]{7,12})\s+(\d+(?:[\.,]\d+)?)\s+(?:UND|MTS|MT|Und|Unidad)\b\s*(.+)', linea_str, re.IGNORECASE)
        if match_b:
            codigo = match_b.group(1).upper(); cant_val = float(match_b.group(2).replace(',', '.')); desc_bruta = match_b.group(3).strip()
        else:
            match_a = re.search(r'^(?:\d+\s+)?([A-Z0-9\-]{7,12})\s+(\d+(?:[\.,]\d+)?)\s*(?:UND|MTS|MT|Und|Unidad)?\s*(.+)$', linea_str, re.IGNORECASE)
            if match_a:
                codigo = match_a.group(1).upper(); cant_val = float(match_a.group(2).replace(',', '.')); desc_bruta = match_a.group(3).strip()

        if codigo and cant_val is not None and desc_bruta:
            if any(exc in codigo for exc in excluidos): continue
            desc_limpia = re.sub(r'(?:UND|MTS|MT|UNDS)\s*$', '', desc_bruta, flags=re.IGNORECASE).strip()
            if not any(item['Código'] == codigo for item in items):
                items.append({"Entrega": nombre_doc, "Código": codigo, "Descripción": desc_limpia, "Cantidad": cant_val})
    return items

def render_consulta_despacho(lista_fuentes):
    if not lista_fuentes: return
    todos_los_items = []
    for pdf_source, tag_entrega in lista_fuentes:
        items_doc = extraer_tabla_materiales(pdf_source, nombre_doc=tag_entrega)
        todos_los_items.extend(items_doc)

    if todos_los_items:
        df_resumen = pd.DataFrame(todos_los_items)
        pesos_u = []; pesos_t = []
        for _, row in df_resumen.iterrows():
            item_match = get_product_data_from_source(str(row['Código']).strip(), df_bd)
            p_val = float(str(item_match.get('Peso_KG', 0.0)).replace(',', '.')) if item_match is not None else 0.0
            cant_val = float(row['Cantidad'])
            pesos_u.append(p_val); pesos_t.append(round(p_val * cant_val, 2))

        df_resumen["Peso Unit. (KG)"] = pesos_u; df_resumen["Peso Total (KG)"] = pesos_t
        df_resumen["No."] = [str(i + 1) for i in range(len(df_resumen))]
        df_resumen = df_resumen[["No.", "Entrega", "Código", "Descripción", "Cantidad", "Peso Unit. (KG)", "Peso Total (KG)"]]

        st.markdown(f"### 📋 Tabla Consolidada de Materiales (Consulta)")
        st.dataframe(df_resumen, use_container_width=True)
        total_kg = pd.to_numeric(df_resumen["Peso Total (KG)"], errors='coerce').sum()
        total_docs = len(set([tag for _, tag in lista_fuentes]))
        m1, m2 = st.columns(2)
        m1.metric("📄 Total Documentos", f"{total_docs}")
        m2.metric("📦 Peso Total (KG)", f"{total_kg:,.2f} KG")

# Cargar base de datos
with st.spinner("🔄 Sincronizando catálogo..."):
    df_bd = fetch_google_sheet_database()

if df_bd is None or df_bd.empty:
    df_bd = cargar_bd_local_universal()

# --- Panel Lateral ---
st.sidebar.markdown(f"👤 **Usuario:** {st.session_state.usuario_actual}")
if st.sidebar.button("Cerrar Sesión", use_container_width=True):
    st.session_state.autenticado = False
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.header("🔐 Modo de Operación")
modo_app = st.sidebar.radio("Seleccione la interfaz:", ["Modo Usuario", "Modo Destroller"])

es_dev_autenticado = False
if modo_app == "Modo Destroller":
    cedula_input = st.sidebar.text_input("Ingrese Contraseña de Destroller:", type="password")
    if cedula_input.strip() == CEDULA_DEV_CORRECTA:
        st.sidebar.success("🔓 Acceso Destroller Autorizado")
        es_dev_autenticado = True

if st.sidebar.button("🔄 Sincronizar Datos"):
    st.cache_data.clear()
    st.rerun()

# --- TÍTULO PRINCIPAL ---
st.title("👍 Facilitador De Procesos Administrativos")

# --- Control de Pestañas según el Modo ---
if es_dev_autenticado:
    tab1, tab2, tab3, tab_rutas, tab_extractor, tab_maestro, tab4 = st.tabs([
        "🔍 Búsqueda por Código", "📄 Procesar Remisión / PDF", "📤 Relación de Envío",
        "Gestión de Rutas", "Clasificador de Rutas", "🗂️ Registro Maestro", "📜 Logs"
    ])
else:
    tab1, tab2, tab3 = st.tabs([
        "🔍 Búsqueda por Código", "📄 Procesar Remisión / PDF", "📤 Relación de Envío"
    ])

# --- PESTAÑA 1: Búsqueda por Código ---
with tab1:
    st.subheader("Búsqueda por Código")
    codigo_input = st.text_input("Ingrese Código o Descripción del Artículo", value="", placeholder="Ej. 108001051")
    cant_input = st.number_input("Cantidad a despachar", min_value=1.0, value=1.0, step=1.0)
    
    if codigo_input.strip() != "":
        item = get_product_data_from_source(codigo_input, df_bd)
        if item is not None:
            peso_unit = float(str(item.get('Peso_KG', 0.0)).replace(',', '.'))
            
            st.markdown(f"""
                <div style="background-color: #f0f2f6; padding: 20px; border-radius: 8px; border-left: 6px solid #1F3864; margin-bottom: 20px; word-wrap: break-word; overflow-wrap: break-word;">
                    <p style="margin: 0 0 8px 0; font-size: 18px; color: #1F3864;"><b>📦 Código:</b> {item['Codigo']}</p>
                    <p style="margin: 0; font-size: 17px; color: #000000; line-height: 1.5;"><b>📝 Descripción:</b> {item['Descripcion']}</p>
                </div>
            """, unsafe_allow_html=True)

            res_col1, res_col2 = st.columns(2)
            res_col1.metric("⚖️ Peso Unitario", f"{peso_unit:.2f} KG")
            res_col2.metric("📊 Peso Total", f"{peso_unit * cant_input:.2f} KG")
        else:
            st.warning("⚠️ No se encontraron coincidencias.")

# --- PESTAÑA 2: Procesar Remisión / PDF (Solo Consulta) ---
with tab2:
    st.subheader("Búsqueda por PDF o Entrega (Consulta de Pesos)")
    st.info("ℹ️ Esta sección es exclusivamente de consulta rápida de pesos mediante archivos PDF locales.")
    
    uploaded_files_tab2 = st.file_uploader("Cargar PDFs de Remisión para Consulta", type=["pdf"], accept_multiple_files=True, key="up_t2")
    lista_fuentes_local = []
    if uploaded_files_tab2:
        for f in uploaded_files_tab2:
            m = re.search(r'\d+', f.name)
            tag = m.group(0) if m else f.name
            lista_fuentes_local.append((f, f"Entrega_{tag}"))
    render_consulta_despacho(lista_fuentes_local)

# --- PESTAÑA 3: Relación de Envío (Generación Oficial) ---
with tab3:
    st.subheader("📤 Generación de Relación de Envío (Formato Oficial)")
    uploaded_files = st.file_uploader("Cargar PDFs para Formato Oficial", type=["pdf"], accept_multiple_files=True, key="up_t3")
    if uploaded_files:
        lista_fuentes = [(f, f"Entrega_{re.search(r'\\d+', f.name).group(0) if re.search(r'\\d+', f.name) else f.name}") for f in uploaded_files]
        render_consulta_despacho(lista_fuentes)

# --- PESTAÑAS MODO DESTROLLER ---
if es_dev_autenticado:
    with tab_rutas:
        st.subheader("Clasificador de Rutas")
        drive_busqueda = st.text_input("🔍 Buscar por No. Entrega en Google Drive:", placeholder="Ej: 20006895")
        pdf_fuente = descargar_pdf_desde_drive(DRIVE_FOLDER_ID, drive_busqueda.strip()) if drive_busqueda.strip() else None
        if pdf_fuente:
            raw_obs = extract_observation_text(pdf_fuente)
            if raw_obs:
                structured = parse_observation_data(raw_obs)
                muni, reg = procesar_ubicacion(structured["Dirección"])
                st.info(f"Dirección: {structured['Dirección']} | Región: {reg}")

    with tab_extractor:
        st.subheader("📑 Clasificador Individual Avanzado")
        uploaded_file_ext = st.file_uploader("Cargar PDF para auditoría", type="pdf", key="uploader_dev")
        if uploaded_file_ext:
            txt = extract_pdf_full_text(uploaded_file_ext)
            st.code(txt[:1000], language="text")

    with tab_maestro:
        st.subheader("🗂️ Registro Maestro de Entregas")
        if os.path.exists(MASTER_CSV_PATH):
            df_master = pd.read_csv(MASTER_CSV_PATH)
            st.dataframe(df_master, use_container_width=True)
            if st.button("🗑️ Limpiar Registro"):
                os.remove(MASTER_CSV_PATH)
                st.success("Registro reiniciado.")
                st.rerun()
        else:
            st.warning("⚠️ No hay registros guardados.")

    with tab4:
        st.subheader("📜 Log Auditoría de Ejecución")
        if os.path.exists(LOG_FILENAME):
            try:
                with open(LOG_FILENAME, "r", encoding="utf-8") as f:
                    log_content = f.read()
                if log_content.strip():
                    st.download_button("💾 Descargar log", log_content, LOG_FILENAME, "text/plain")
                    st.text_area("Contenido del Log:", value=log_content, height=400)
                else:
                    st.info("ℹ️ El archivo de log está vacío por el momento.")
            except Exception as e:
                st.error(f"Error leyendo el archivo de log: {e}")
        else:
            st.warning("⚠️ El archivo de log aún no ha sido creado.")
