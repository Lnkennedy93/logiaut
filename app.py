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
import csv
import requests
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as OpenpyxlImage
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer, Image as ReportLabImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from supabase import create_client, Client

# ---------------------------------------------------------
# Configuración inicial (Adaptativa y Pestaña)
# ---------------------------------------------------------
st.set_page_config(
    page_title="Asistente_web", 
    page_icon="📦", 
    layout="wide"
)

# Estilos CSS globales para adaptabilidad y botón de guardar grande, azul y muy visible
st.markdown("""
    <html lang="es" class="notranslate" translate="no">
    <head><meta name="google" content="notranslate" /></head>
    </html>
    <style>
        .stDataFrame, .stTable {
            width: 100% !important;
            overflow-x: auto !important;
        }
        div[data-testid="stDataFrame"] div, td, th {
            white-space: normal !important;
            word-wrap: break-word !important;
            text-overflow: clip !important;
        }
        div.stFormSubmitButton > button {
            background-color: #1F3864 !important;
            color: white !important;
            font-size: 18px !important;
            font-weight: bold !important;
            padding: 1rem 2rem !important;
            border-radius: 10px !important;
            border: 3px solid #FFD700 !important;
            box-shadow: 0px 6px 12px rgba(0, 0, 0, 0.4) !important;
            width: 100% !important;
            cursor: pointer !important;
        }
        div.stFormSubmitButton > button:hover {
            background-color: #2c4d8c !important;
            border-color: #ffffff !important;
        }
        @media (max-width: 768px) {
            .main .block-container {
                padding: 1rem 0.5rem;
            }
        }
    </style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# Parámetros y Constantes
# ---------------------------------------------------------
DRIVE_FOLDER_ID = "1Amwy8_uQgo6X0VS2DXH028Ep80BMi4rP"
GOOGLE_SHEET_ID = "1aTlmA6JBldTX3zN-djDjWA5HEAExTPcdNhJsPJL9Kgo"
GOOGLE_SHEET_GID = "1928951055"
GOOGLE_SHEET_CSV_URL = (
    f"https://docs.google.com/spreadsheets/d/{GOOGLE_SHEET_ID}/export"
    f"?format=csv&gid={GOOGLE_SHEET_GID}"
)
CEDULA_DEV_CORRECTA = "1073513861"
BD_LOCAL_PATH = "BD_Pesos.xlsx"
MASTER_CSV_PATH = "registro_entregas.csv"
LOGO_PATH = "logo.png" if os.path.exists("logo.png") else ("download.png" if os.path.exists("download.png") else None)

# --- Logging ---
LOG_FILENAME = "log_ejecucion.txt"
logging.basicConfig(
    filename=LOG_FILENAME,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    force=True
)

def registrar_log(mensaje, tipo="INFO"):
    print(f"[{tipo}] {mensaje}")
    if tipo == "INFO": logging.info(mensaje)
    elif tipo == "WARNING": logging.warning(mensaje)
    elif tipo == "ERROR": logging.error(mensaje)
    for handler in logging.root.handlers:
        handler.flush()

registrar_log("--- Sesión iniciada ---")

# ---------------------------------------------------------
# Detector de Sesión / IP del Cliente Único
# ---------------------------------------------------------
def get_client_ip():
    try:
        if hasattr(st, "context") and hasattr(st.context, "headers"):
            headers = st.context.headers
            ip = headers.get("X-Forwarded-For", headers.get("X-Real-IP", ""))
            if ip:
                if "," in ip: ip = ip.split(",")[0].strip()
                return ip
    except Exception as e:
        registrar_log(f"Error obteniendo IP: {e}", "WARNING")
    if "browser_session_id" not in st.session_state:
        import uuid
        st.session_state["browser_session_id"] = str(uuid.uuid4())
    return st.session_state["browser_session_id"]

# ---------------------------------------------------------
# Conexión Google Drive
# ---------------------------------------------------------
def get_google_creds():
    scope = [
        "https://www.googleapis.com/auth/drive",
        "https://www.googleapis.com/auth/spreadsheets.readonly",
    ]
    if "google" in st.secrets:
        creds_dict = dict(st.secrets["google"])
        if "private_key" in creds_dict:
            creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        try:
            return service_account.Credentials.from_service_account_info(creds_dict, scopes=scope)
        except Exception as error:
            registrar_log(
                f"La clave de Google en st.secrets no es válida: {error}. "
                "Se intentará credentials.json.",
                "WARNING"
            )
    elif os.path.exists('credentials.json'):
        return service_account.Credentials.from_service_account_file(
            'credentials.json', scopes=scope
        )
    if os.path.exists('credentials.json'):
        return service_account.Credentials.from_service_account_file(
            'credentials.json', scopes=scope
        )
    return None

def descargar_pdf_desde_drive(folder_id, numero_entrega):
    try:
        creds = get_google_creds()
        if not creds: return None
        service = build('drive', 'v3', credentials=creds)
        query = f"mimeType='application/pdf' and name contains '{numero_entrega}' and trashed=false"
        results = service.files().list(
            q=query, pageSize=1, fields="files(id, name)",
            supportsAllDrives=True, includeItemsFromAllDrives=True
        ).execute()
        files = results.get('files', [])
        if not files: return None
        file_id = files[0]['id']
        request = service.files().get_media(fileId=file_id)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()
        fh.seek(0)
        return fh
    except Exception as e:
        registrar_log(f"Error Drive {numero_entrega}: {e}", "ERROR")
        return None

# ---------------------------------------------------------
# Supabase y Autenticación
# ---------------------------------------------------------
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
        registrar_log(f"Error de conexión Supabase login: {e}", "ERROR")
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
                    except Exception as ex:
                        registrar_log(f"Error actualizando IP login: {ex}", "ERROR")
                    
                    st.session_state.autenticado = True
                    st.session_state.usuario_actual = user_data["correo"]
                    st.session_state.empresa_actual = user_data["empresa"]
                    st.session_state.login_token = client_token
                    st.success("¡Licencia verificada con éxito!")
                    st.rerun()
                else:
                    st.error("❌ Credenciales incorrectas o licencia inactiva.")
        
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.markdown(
            """
            <div style="text-align: center; color: gray; font-size: 14px;">
                <p>Dificil Mientras Llega A Nosotros</p>
                <p><b>Elaborado por Liontech</b></p>
            </div>
            """,
            unsafe_allow_html=True
        )

if not st.session_state.autenticado:
    pantalla_login()
    st.stop()
else:
    try:
        resp_token = supabase.table("usuarios_licencias").select("current_ip").eq("correo", st.session_state.usuario_actual).execute()
        if resp_token.data:
            db_token = resp_token.data[0].get("current_ip")
            if db_token and db_token != st.session_state.get("login_token", ""):
                st.session_state.autenticado = False
                st.session_state.usuario_actual = None
                st.session_state.empresa_actual = None
                st.warning("⚠️ Tu sesión ha sido cerrada porque se inició sesión desde otro dispositivo.")
                st.stop()
    except Exception as e:
        registrar_log(f"Error validando IP activa: {e}", "WARNING")

# ---------------------------------------------------------
# Capa de Datos Confiable (Base Local Excel)
# ---------------------------------------------------------
def limpiar_peso(val):
    if pd.isna(val): return 0.0
    val_str = str(val).strip().replace(',', '.').replace('KG', '').replace('kg', '').strip()
    try:
        return float(val_str)
    except ValueError:
        return 0.0

def get_product_data_from_source(codigo_input, df_bd):
    env_mode = st.session_state.get("env_mode", "DEV (Google)")
    clean_input = str(codigo_input).strip().upper()
    
    if env_mode == "PROD (SAP)":
        if df_bd is not None:
            match_exact = df_bd[df_bd['Codigo'].astype(str).str.strip().str.upper() == clean_input]
            if not match_exact.empty:
                item_sap = match_exact.iloc[0].copy()
                item_sap['Descripcion'] = f"[SAP S/4HANA] {item_sap['Descripcion']}"
                return item_sap
            match_parcial = df_bd[df_bd['Codigo'].astype(str).str.contains(clean_input, case=False, na=False)]
            if not match_parcial.empty:
                item_sap = match_parcial.iloc[0].copy()
                item_sap['Descripcion'] = f"[SAP S/4HANA] {item_sap['Descripcion']}"
                return item_sap
        return pd.Series({"Codigo": clean_input, "Descripcion": f"[SAP S/4HANA] NO ENCONTRADO ({codigo_input})", "Peso_KG": 0.0})
    else:
        if df_bd is not None:
            match_exact = df_bd[df_bd['Codigo'].astype(str).str.strip().str.upper() == clean_input]
            if not match_exact.empty: return match_exact.iloc[0]
            match_parcial = df_bd[df_bd['Codigo'].astype(str).str.contains(clean_input, case=False, na=False)]
            if not match_parcial.empty: return match_parcial.iloc[0]
        return None

def parse_google_sheet_csv(csv_text):
    """Convierte la pestaña CSV en el esquema de productos de la aplicación."""
    rows = list(csv.reader(io.StringIO(csv_text)))
    if len(rows) < 2:
        return None

    headers = [str(value).strip().upper().replace('"', '') for value in rows[0]]

    def find_column(*terms):
        return next((index for index, header in enumerate(headers)
                     if any(term in header for term in terms)), None)

    code_idx = find_column("CODIGO", "CÓDIGO", "MATERIAL", "SAP")
    desc_idx = find_column("DESC", "DESCRIPCION", "DESCRIPCIÓN", "PRODUCTO", "NOMBRE")
    peso_idx = find_column("PESO", "KG", "KILOS", "UNIT")
    if code_idx is None:
        code_idx = 0
    if desc_idx is None:
        desc_idx = 1
    if peso_idx is None:
        peso_idx = 2

    products = []
    for row in rows[1:]:
        if not row:
            continue

        def value_at(index):
            return str(row[index]).strip() if index < len(row) else ""

        clean_code = re.sub(r"\.0$", "", value_at(code_idx))
        description = value_at(desc_idx)
        raw_weight = re.sub(r"[^\d.,-]", "", value_at(peso_idx))
        if "," in raw_weight and "." in raw_weight:
            raw_weight = raw_weight.replace(".", "").replace(",", ".")
        else:
            raw_weight = raw_weight.replace(",", ".")

        try:
            weight = float(raw_weight) if raw_weight else 0.0
        except ValueError:
            weight = 0.0

        if clean_code and description:
            products.append({
                "Codigo": clean_code,
                "Descripcion": description,
                "Peso_KG": weight,
            })

    return pd.DataFrame(products) if products else None

def fetch_google_sheet_database_api():
    """Lee y unifica todas las pestañas del catálogo mediante la API."""
    creds = get_google_creds()
    if not creds:
        return None

    service = build("sheets", "v4", credentials=creds)
    metadata = service.spreadsheets().get(
        spreadsheetId=GOOGLE_SHEET_ID,
        fields="sheets.properties(sheetId,title,hidden)",
    ).execute()
    sheets = [sheet["properties"] for sheet in metadata.get("sheets", [])]
    if not sheets:
        return None

    ranges = [
        f"'{sheet['title'].replace(chr(39), chr(39) * 2)}'!A:Z"
        for sheet in sheets
    ]
    result = service.spreadsheets().values().batchGet(
        spreadsheetId=GOOGLE_SHEET_ID,
        ranges=ranges,
    ).execute()
    data_frames = []
    sheets_with_data = 0

    for sheet, value_range in zip(sheets, result.get("valueRanges", [])):
        rows = value_range.get("values", [])
        if len(rows) < 2:
            continue

        csv_buffer = io.StringIO()
        csv.writer(csv_buffer).writerows(rows)
        df_sheet = parse_google_sheet_csv(csv_buffer.getvalue())
        if df_sheet is not None and not df_sheet.empty:
            df_sheet["PESTAÑA_BD"] = sheet["title"]
            data_frames.append(df_sheet)
            sheets_with_data += 1

    if not data_frames:
        return None

    df_all = pd.concat(data_frames, ignore_index=True)
    df_all = df_all.drop_duplicates(subset=["Codigo"], keep="first")
    registrar_log(
        f"Catálogo sincronizado desde {sheets_with_data}/{len(sheets)} pestañas: "
        f"{len(df_all)} productos válidos."
    )
    return df_all

@st.cache_data(ttl=1800, show_spinner=False)
def fetch_google_sheet_database():
    """Sincroniza todas las pestañas y conserva Excel como respaldo local."""
    try:
        registrar_log(
            f"Conectando con todas las pestañas de Google Sheets ({GOOGLE_SHEET_ID})..."
        )
        df_google = fetch_google_sheet_database_api()
        if df_google is not None and not df_google.empty:
            return df_google
        registrar_log("Google Sheets no contiene filas válidas; probando CSV.", "WARNING")
    except Exception as error:
        registrar_log(
            f"Error leyendo todas las pestañas por API: {error}.",
            "WARNING"
        )

    try:
        response = requests.get(GOOGLE_SHEET_CSV_URL, timeout=20)
        response.raise_for_status()
        df_google = parse_google_sheet_csv(response.text)
        if df_google is not None and not df_google.empty:
            registrar_log(
                f"CSV de Google Sheets sincronizado: {len(df_google)} productos."
            )
            return df_google
    except Exception as error:
        registrar_log(f"CSV de Google Sheets no disponible: {error}.", "WARNING")

    registrar_log("Usando base local de productos como respaldo.", "WARNING")
    return cargar_bd_local(BD_LOCAL_PATH)

@st.cache_data(ttl=3600, show_spinner=False)
def cargar_bd_local(ruta_archivo):
    if not os.path.exists(ruta_archivo): 
        registrar_log(f"No se encontró el archivo local de pesos en {ruta_archivo}", "ERROR")
        return None
    try:
        hojas_dict = pd.read_excel(ruta_archivo, sheet_name=None, engine='openpyxl')
        list_dfs = []
        for nombre_hoja, df_hoja in hojas_dict.items():
            if not df_hoja.empty:
                df_h = df_hoja.copy()
                df_h.columns = df_h.columns.astype(str).str.strip()
                
                col_c = next((c for c in df_h.columns if c.lower() in ['codigo', 'código']), None)
                col_d = next((c for c in df_h.columns if 'desc' in c.lower()), None)
                col_p = next((c for c in df_h.columns if c.lower() in ['peso_kg', 'peso', 'kg']), None)
                
                if not col_c and len(df_h.columns) > 1: col_c = df_h.columns[1]
                if not col_d and len(df_h.columns) > 2: col_d = df_h.columns[2]
                if not col_p and len(df_h.columns) > 4: col_p = df_h.columns[4]

                if col_c:
                    temp_df = pd.DataFrame()
                    temp_df['Codigo'] = df_h[col_c]
                    temp_df['Descripcion'] = df_h[col_d] if col_d else "Sin descripción"
                    temp_df['Peso_KG'] = df_h[col_p] if col_p else 0.0
                    
                    temp_df = temp_df.dropna(subset=['Codigo'])
                    list_dfs.append(temp_df)
                    
        if list_dfs:
            df_final = pd.concat(list_dfs, ignore_index=True)
            df_final['Codigo'] = df_final['Codigo'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
            df_final['Peso_KG'] = df_final['Peso_KG'].apply(limpiar_peso)
            registrar_log(f"Base de datos local estandarizada y cargada con {len(df_final)} registros.", "INFO")
            return df_final
        return None
    except Exception as e:
        registrar_log(f"Error cargando base local Excel: {e}", "ERROR")
        return None

# ---------------------------------------------------------
# Generadores Excel y PDF (GID-F-010)
# ---------------------------------------------------------
def generar_excel_bytes(df_resumen, total_docs, total_kg, total_ton, driver_info=None, dest_info=None, elaborado_info=None, empaques_info=None):
    if driver_info is None: driver_info = {}
    if dest_info is None: dest_info = {}
    if elaborado_info is None: elaborado_info = {}
    if empaques_info is None: empaques_info = {}

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "RELACION DE ENVIO"
    ws.views.sheetView[0].showGridLines = True

    font_brand_bold_10 = Font(name="Arial", size=10, bold=True, color="1F3864")
    font_brand_bold_14 = Font(name="Arial", size=14, bold=True, color="1F3864")
    font_arial_bold_10 = Font(name="Arial", size=10, bold=True)
    font_arial_bold_11 = Font(name="Arial", size=11, bold=True)
    font_calibri_9_bold = Font(name="Calibri", size=9, bold=True)

    thin_border = Border(left=Side(style='thin', color='000000'), right=Side(style='thin', color='000000'), top=Side(style='thin', color='000000'), bottom=Side(style='thin', color='000000'))
    fill_header_gray = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    fill_section_gray = PatternFill(start_color="BFBFBF", end_color="BFBFBF", fill_type="solid")

    ws.merge_cells("A1:B5")
    ws["A1"] = "Sistema de Gestión Integral\nTUVACOL S.A."
    ws["A1"].font = font_brand_bold_10
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    if LOGO_PATH and os.path.exists(LOGO_PATH):
        try:
            img = OpenpyxlImage(LOGO_PATH); img.width = 110; img.height = 55; ws.add_image(img, "A1")
        except Exception: pass

    ws.merge_cells("C1:E5"); ws["C1"] = "RELACION DE ENVIO DE MERCANCIAS"
    ws["C1"].font = font_brand_bold_14; ws["C1"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("F1:G1"); ws["F1"] = "Versión: 05"
    ws.merge_cells("F2:G2"); ws["F2"] = "Vigente desde: 12-02-2021"
    ws.merge_cells("F3:G3"); ws["F3"] = "Codigo: GID-F-010"
    ws.merge_cells("F4:G4"); ws["F4"] = "Elaborado por: Comité SGI"
    ws.merge_cells("F5:G5"); ws["F5"] = "Revisado y Aprobado por: Comité SGI"

    for r in range(1, 6):
        for c in range(1, 8): ws.cell(r, c).border = thin_border

    ws["A7"] = "Responsable / Cargo:"; ws["A7"].font = font_arial_bold_10
    ws.merge_cells("C7:D7"); ws["C7"] = "LOGÍSTICA SEDE FUNZA."; ws["C7"].font = font_arial_bold_10
    
    num_relacion = f"No.{datetime.datetime.now().strftime('%m%d%H%M')}"
    ws["E7"] = num_relacion; ws["E7"].font = font_arial_bold_11
    ws["E7"].alignment = Alignment(horizontal="center", vertical="center")
    
    ws["F7"] = "Fecha:"; ws["F7"].font = font_arial_bold_10
    ws["G7"] = datetime.datetime.now().strftime("%d-%m-%Y"); ws["G7"].font = font_arial_bold_10

    headers = ["Unidades enviadas", "Descripción del empaque", "Descripción de la mercancía", "", "Código de la mercancía", "Peso (Kilos)", "Documentos de referencia"]
    ws.merge_cells("C9:D9")
    for col_idx, text in enumerate(headers, 1):
        if col_idx == 4: continue
        cell = ws.cell(9, col_idx, text)
        cell.font = font_calibri_9_bold; cell.fill = fill_header_gray; cell.border = thin_border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    current_row = 10; tot_und = 0.0; tot_mts = 0.0; docs_unicos = []
    for idx, row in df_resumen.iterrows():
        cant_val = float(row["Cantidad"])
        desc_str = str(row["Descripción"])
        doc_num = re.sub(r'^[^\d]*', '', str(row["Entrega"]))
        if doc_num and doc_num not in docs_unicos: docs_unicos.append(doc_num)

        if "TUBERIA" in desc_str.upper():
            empaque_tag = "TUBERIA"; unidad_str = f"{cant_val:,.0f} MTS"; tot_mts += cant_val
        else:
            empaque_tag = "ACCESORIOS"; unidad_str = f"{cant_val:,.0f} UND"; tot_und += cant_val

        ws.cell(current_row, 1, unidad_str).alignment = Alignment(horizontal="center")
        ws.cell(current_row, 2, empaque_tag).alignment = Alignment(horizontal="center")
        ws.merge_cells(start_row=current_row, start_column=3, end_row=current_row, end_column=4)
        ws.cell(current_row, 3, desc_str).alignment = Alignment(horizontal="left", wrap_text=True)
        ws.cell(current_row, 5, str(row["Código"])).alignment = Alignment(horizontal="center")
        p_tot = float(row.get("Peso Total (KG)", 0.0))
        ws.cell(current_row, 6, p_tot).alignment = Alignment(horizontal="right"); ws.cell(current_row, 6).number_format = "#,##0.00"
        ws.cell(current_row, 7, doc_num).alignment = Alignment(horizontal="center", vertical="center")

        for col in range(1, 8): ws.cell(current_row, col).border = thin_border
        current_row += 1

    emp_str = f"Empaques - Guacales: {empaques_info.get('Guacales',0)}, Estibas: {empaques_info.get('Estibas',0)}, Cajas: {empaques_info.get('Cajas',0)}, Paquetes: {empaques_info.get('Paquetes',0)}, Sobres: {empaques_info.get('Sobres',0)}, Tubos: {empaques_info.get('Tubos',0)}"
    obs_texto = f"Observaciones: TOTAL: {tot_und:,.0f} UND, {tot_mts:,.0f} MTS | {emp_str} | DOCS: {', '.join(docs_unicos)}"
    
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=4)
    ws.cell(current_row, 1, obs_texto).font = font_calibri_9_bold
    ws.cell(current_row, 5, "Peso Total:").fill = fill_header_gray
    ws.merge_cells(start_row=current_row, start_column=6, end_row=current_row, end_column=7)
    cell_tot = ws.cell(current_row, 6, f"=SUM(F10:F{current_row-1})")
    cell_tot.font = font_arial_bold_11; cell_tot.number_format = "#,##0.00"

    for col in range(1, 8): ws.cell(current_row, col).border = thin_border
    current_row += 2

    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=3); ws.cell(current_row, 1, "Datos del Remitente").fill = fill_header_gray
    ws.merge_cells(start_row=current_row, start_column=4, end_row=current_row, end_column=7); ws.cell(current_row, 4, "Datos del Destinatario").fill = fill_header_gray
    for col in range(1, 8): ws.cell(current_row, col).border = thin_border
    current_row += 1

    ws.cell(current_row, 1, "Nombre:"); ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=3); ws.cell(current_row, 2, "TUVACOL SEDE FUNZA")
    ws.cell(current_row, 4, "Nombre:"); ws.merge_cells(start_row=current_row, start_column=5, end_row=current_row, end_column=7); ws.cell(current_row, 5, dest_info.get("nombre", ""))
    for col in range(1, 8): ws.cell(current_row, col).border = thin_border
    current_row += 1

    ws.cell(current_row, 1, "Dirección y Teléfono:"); ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=3); ws.cell(current_row, 2, "KM 3,5 VIA FUNZA SIBERIA")
    ws.cell(current_row, 4, "Dirección y Teléfono:"); ws.merge_cells(start_row=current_row, start_column=5, end_row=current_row, end_column=7); ws.cell(current_row, 5, dest_info.get("direccion", ""))
    for col in range(1, 8): ws.cell(current_row, col).border = thin_border
    current_row += 1

    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7); ws.cell(current_row, 1, "Datos del Conductor y Vehiculo").fill = fill_section_gray
    for col in range(1, 8): ws.cell(current_row, col).border = thin_border
    current_row += 1

    ws.cell(current_row, 1, "Nombre:"); ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=3); ws.cell(current_row, 2, driver_info.get("nombre", ""))
    ws.cell(current_row, 4, "N. de Placa:"); ws.cell(current_row, 5, driver_info.get("placa", ""))
    ws.cell(current_row, 6, "Modelo/Marca:"); ws.cell(current_row, 7, driver_info.get("marca", ""))
    for col in range(1, 8): ws.cell(current_row, col).border = thin_border
    current_row += 1

    ws.cell(current_row, 1, "Cedula No.:"); ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=3); ws.cell(current_row, 2, driver_info.get("cedula", ""))
    ws.cell(current_row, 4, "Transportadora:"); ws.merge_cells(start_row=current_row, start_column=5, end_row=current_row, end_column=7); ws.cell(current_row, 5, driver_info.get("transportadora", ""))
    for col in range(1, 8): ws.cell(current_row, col).border = thin_border
    current_row += 1

    ws.cell(current_row, 1, "N. de Celular:"); ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=3); ws.cell(current_row, 2, driver_info.get("celular", ""))
    for col in range(1, 8): ws.cell(current_row, col).border = thin_border
    current_row += 2

    output = io.BytesIO(); wb.save(output); output.seek(0)
    return output.getvalue()

def generar_pdf_bytes(df_resumen, total_docs, total_kg, total_ton, driver_info=None, dest_info=None, elaborado_info=None, empaques_info=None):
    if driver_info is None: driver_info = {}
    if dest_info is None: dest_info = {}
    if elaborado_info is None: elaborado_info = {}
    if empaques_info is None: empaques_info = {}

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
    elements = []
    styles = getSampleStyleSheet()
    normal_style = ParagraphStyle('Normal8', parent=styles['Normal'], fontSize=7.5, leading=9)
    bold_style = ParagraphStyle('Bold8', parent=normal_style, fontName='Helvetica-Bold')
    brand_title_style = ParagraphStyle('BrandTitle10', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, leading=12, alignment=1, textColor=colors.HexColor('#1F3864'))

    sgi_cell = Paragraph("<b>Sistema de Gestión Integral</b><br/><font size=9><b>TUVACOL S.A.</b></font>", brand_title_style)
    num_relacion = f"No.{datetime.datetime.now().strftime('%m%d%H%M')}"
    meta_cell = Paragraph(f"Versión: 05<br/><b>{num_relacion}</b><br/><b>Codigo: GID-F-010</b>", normal_style)
    h_data = [[sgi_cell, Paragraph("<b>RELACION DE ENVIO DE MERCANCIAS</b>", brand_title_style), meta_cell]]
    
    t_header = Table(h_data, colWidths=[150, 270, 150])
    t_header.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 1, colors.black), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
    elements.append(t_header); elements.append(Spacer(1, 4))

    d_table_data = [[Paragraph("<b>Unidades enviadas</b>", normal_style), Paragraph("<b>Empaque</b>", normal_style), Paragraph("<b>Descripción de la mercancía</b>", normal_style), Paragraph("<b>Código</b>", normal_style), Paragraph("<b>Peso (KG)</b>", normal_style), Paragraph("<b>Doc. Ref.</b>", normal_style)]]
    tot_und = 0.0; tot_mts = 0.0; docs_unicos = []

    for _, row in df_resumen.iterrows():
        p_tot = float(row.get("Peso Total (KG)", 0.0))
        doc_num = re.sub(r'^[^\d]*', '', str(row["Entrega"]))
        if doc_num and doc_num not in docs_unicos: docs_unicos.append(doc_num)
        cant_val = float(row["Cantidad"]); desc_str = str(row["Descripción"])
        if "TUBERIA" in desc_str.upper(): unidad_str = f"{cant_val:,.0f} MTS"; tot_mts += cant_val
        else: unidad_str = f"{cant_val:,.0f} UND"; tot_und += cant_val

        d_table_data.append([Paragraph(unidad_str, normal_style), Paragraph("TUBERIA" if "TUBERIA" in desc_str.upper() else "ACCESORIO", normal_style), Paragraph(desc_str, normal_style), Paragraph(str(row['Código']), normal_style), Paragraph(f"{p_tot:,.2f}", normal_style), Paragraph(doc_num, normal_style)])

    emp_str = f"Guacales: {empaques_info.get('Guacales',0)}, Estibas: {empaques_info.get('Estibas',0)}, Cajas: {empaques_info.get('Cajas',0)}, Paquetes: {empaques_info.get('Paquetes',0)}, Sobres: {empaques_info.get('Sobres',0)}, Tubos: {empaques_info.get('Tubos',0)}"
    d_table_data.append([Paragraph(f"<b>TOTAL: {tot_und:,.0f} UND, {tot_mts:,.0f} MTS | {emp_str}</b>", normal_style), "", "", Paragraph("<b>Peso Total:</b>", bold_style), Paragraph(f"<b>{total_kg:,.2f}</b>", bold_style), Paragraph(f"<b>{total_ton:,.3f} T</b>", bold_style)])
    t_data = Table(d_table_data, colWidths=[65, 65, 245, 75, 62, 60])
    t_data.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 1, colors.black), ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#D9D9D9')), ('SPAN', (0, -1), (2, -1))]))
    elements.append(t_data); elements.append(Spacer(1, 4))

    doc.build(elements)
    buffer.seek(0)
    return buffer.getvalue()

# ---------------------------------------------------------
# Extractor Inteligente de Materiales
# ---------------------------------------------------------
def extraer_tabla_materiales(pdf_source, nombre_doc="Desconocido"):
    texto_completo = ""
    try:
        if hasattr(pdf_source, 'seek'): pdf_source.seek(0)
        with pdfplumber.open(pdf_source) as pdf:
            for page in pdf.pages:
                txt = page.extract_text()
                if txt: texto_completo += txt + "\n"
    except Exception as e:
        registrar_log(f"Error extrayendo texto del PDF {nombre_doc}: {e}", "ERROR")
        return []

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

def extract_pdf_full_text(pdf_file):
    try:
        with pdfplumber.open(pdf_file) as pdf:
            return "".join([page.extract_text() + "\n" for page in pdf.pages if page.extract_text()])
    except Exception: return ""

def extract_observation_text(pdf_file):
    try:
        with pdfplumber.open(pdf_file) as pdf:
            text = pdf.pages[0].extract_text()
        match = re.search(r"Observación:\s*(.*?)(?=\n[A-Z]|$)", text, re.DOTALL | re.IGNORECASE)
        return match.group(1).replace("\n", " ").strip() if match else text[:300].replace("\n", " ")
    except Exception: return None

def parse_observation_data(text):
    data = {"Dirección": "No detectado", "Contacto": "No detectado", "Celular": "No detectado", "OC": "No detectado"}
    text_clean = text.replace("\n", " ").strip()
    if (cel_m := re.search(r'\b(3\d{2}[\s\-]?\d{3}[\s\-]?\d{4}|\b3\d{9}\b|\b\d{10}\b)', text_clean)): data["Celular"] = cel_m.group(1).strip()
    if (dir_m := re.search(r"(?:ENTREGAR EN|DIRECCION DE ENVIO|DIRECCIÓN)[:]??\s*(.*?)(?=\s*//|\s*CONTACTO|$)", text_clean, re.IGNORECASE)): data["Dirección"] = dir_m.group(1).strip()
    else: data["Dirección"] = text_clean
    return data

def procesar_ubicacion(direccion):
    dir_upper = direccion.upper()
    mun = ["BOGOTA", "FUNZA", "MOSQUERA", "MADRID", "FACATATIVA", "CHIA", "CAJICA", "ZIPAQUIRA", "YUMBO", "CALI", "BARRANQUILLA", "MEDELLIN", "ITAGUI", "COTA", "TOCANCIPA"]
    muni_det = next((m for m in mun if m in dir_upper), "NO ESPECIFICADO")
    reg = "SABANA OCCIDENTE" if muni_det in ["FUNZA", "MOSQUERA", "MADRID", "FACATATIVA"] else "SUR"
    return muni_det, reg

# ---------------------------------------------------------
# Renderizador de Consulta Pestaña 2
# ---------------------------------------------------------
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
        total_ton = total_kg / 1000.0
        total_docs = len(set([tag for _, tag in lista_fuentes]))

        m1, m2 = st.columns(2)
        m1.metric("📄 Total Documentos", f"{total_docs}")
        m2.metric("📦 Peso Total (KG)", f"{total_kg:,.2f} KG")
        st.info("ℹ️ Esta pestaña es exclusivamente de consulta de pesos y materiales. Para generar el formato oficial, use la pestaña **Relación de Envío**.")

# ---------------------------------------------------------
# Renderizador de Despacho Completo (Pestaña 3)
# ---------------------------------------------------------
def render_procesamiento_despacho(lista_fuentes, tab_key):
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

        st.markdown(f"### 📋 Tabla Consolidada de Materiales")
        st.dataframe(df_resumen, use_container_width=True)

        total_kg = pd.to_numeric(df_resumen["Peso Total (KG)"], errors='coerce').sum()
        total_ton = total_kg / 1000.0
        total_docs = len(set([tag for _, tag in lista_fuentes]))

        m1, m2 = st.columns(2)
        m1.metric("📄 Total Documentos", f"{total_docs}")
        m2.metric("📦 Peso Total (KG)", f"{total_kg:,.2f} KG")

        st.markdown("---")
        st.markdown("### 📝 Datos Obligatorios para Generar el Formato Oficial")
        st.info("⚠️ **Aviso:** Todos los campos marcados con asterisco (*) son obligatorios. Rellene los datos y presione 'Guardar Cambios' para habilitar las descargas.")

        form_key = f"form_despacho_{tab_key}"
        
        if st.session_state.get(f"limpiar_form_{tab_key}", False):
            st.session_state[f"datos_guardados_{tab_key}"] = False
            if f"saved_data_{tab_key}" in st.session_state: del st.session_state[f"saved_data_{tab_key}"]
            if f"saved_empaques_{tab_key}" in st.session_state: del st.session_state[f"saved_empaques_{tab_key}"]
            st.session_state[f"limpiar_form_{tab_key}"] = False

        saved = st.session_state.get(f"saved_data_{tab_key}", {
            "dest_name": "", "dest_address": "", "d_nombre": "", "d_placa": "",
            "d_cedula": "", "d_marca": "", "d_celular": "", "d_transp": "", "elab_nombre": ""
        })
        
        saved_emp = st.session_state.get(f"saved_empaques_{tab_key}", {
            "Guacales": 0, "Estibas": 0, "Cajas": 0, "Paquetes": 0, "Sobres": 0, "Tubos": 0
        })

        with st.form(key=form_key):
            col_d1, col_d2 = st.columns(2)
            with col_d1:
                dest_name = st.text_input("Nombre del Destinatario *", value=saved["dest_name"], placeholder="Obligatorio")
            with col_d2:
                dest_address = st.text_input("Dirección / Teléfono *", value=saved["dest_address"], placeholder="Obligatorio")

            col_c1, col_c2, col_c3 = st.columns(3)
            with col_c1:
                d_nombre = st.text_input("Nombre Conductor *", value=saved["d_nombre"], placeholder="Obligatorio")
                d_placa = st.text_input("Placa Vehículo *", value=saved["d_placa"], placeholder="Obligatorio")
            with col_c2:
                d_cedula = st.text_input("Cédula No. *", value=saved["d_cedula"], placeholder="Obligatorio")
                d_marca = st.text_input("Marca / Modelo *", value=saved["d_marca"], placeholder="Obligatorio")
            with col_c3:
                d_celular = st.text_input("Celular *", value=saved["d_celular"], placeholder="Obligatorio")
                d_transp = st.text_input("Empresa Transportadora *", value=saved["d_transp"], placeholder="Obligatorio")

            elab_nombre = st.text_input("Elaborado por (Nombre y Cargo) *", value=saved["elab_nombre"], placeholder="Obligatorio")

            st.markdown("---")
            st.markdown("#### 📦 Cantidades de Empaques Requeridos (Al menos uno mayor a 0)")
            e_col1, e_col2, e_col3 = st.columns(3)
            with e_col1:
                guacales = st.number_input("1. Guacales *", min_value=0, step=1, value=int(saved_emp["Guacales"]))
                estibas = st.number_input("2. Estibas *", min_value=0, step=1, value=int(saved_emp["Estibas"]))
            with e_col2:
                cajas = st.number_input("3. Cajas *", min_value=0, step=1, value=int(saved_emp["Cajas"]))
                paquetes = st.number_input("4. Paquetes *", min_value=0, step=1, value=int(saved_emp["Paquetes"]))
            with e_col3:
                sobres = st.number_input("5. Sobres *", min_value=0, step=1, value=int(saved_emp["Sobres"]))
                tubos = st.number_input("6. Tubos *", min_value=0, step=1, value=int(saved_emp["Tubos"]))

            submitted = st.form_submit_button(label="💾 GUARDAR CAMBIOS Y HABILITAR DESCARGAS")

            if submitted:
                total_empaques = guacales + estibas + cajas + paquetes + sobres + tubos
                if not dest_name.strip() or not dest_address.strip() or not d_nombre.strip() or not d_placa.strip() or not d_cedula.strip() or not d_marca.strip() or not d_celular.strip() or not d_transp.strip() or not elab_nombre.strip():
                    st.error("❌ Todos los campos de texto son obligatorios.")
                    st.session_state[f"datos_guardados_{tab_key}"] = False
                elif total_empaques <= 0:
                    st.error("❌ Debes indicar la cantidad de al menos un tipo de empaque.")
                    st.session_state[f"datos_guardados_{tab_key}"] = False
                else:
                    st.session_state[f"datos_guardados_{tab_key}"] = True
                    st.session_state[f"saved_data_{tab_key}"] = {
                        "dest_name": dest_name.strip(), "dest_address": dest_address.strip(),
                        "d_nombre": d_nombre.strip(), "d_placa": d_placa.strip(), "d_cedula": d_cedula.strip(),
                        "d_marca": d_marca.strip(), "d_celular": d_celular.strip(), "d_transp": d_transp.strip(),
                        "elab_nombre": elab_nombre.strip()
                    }
                    st.session_state[f"saved_empaques_{tab_key}"] = {
                        "Guacales": guacales, "Estibas": estibas, "Cajas": cajas,
                        "Paquetes": paquetes, "Sobres": sobres, "Tubos": tubos
                    }
                    st.success("✅ ¡Datos validados y guardados correctamente! Ya puedes descargar los formatos oficiales.")

        if st.session_state.get(f"datos_guardados_{tab_key}", False):
            saved_active = st.session_state[f"saved_data_{tab_key}"]
            saved_emp_active = st.session_state[f"saved_empaques_{tab_key}"]
            dest_info = {"nombre": saved_active["dest_name"], "direccion": saved_active["dest_address"]}
            driver_info = {"nombre": saved_active["d_nombre"], "cedula": saved_active["d_cedula"], "celular": saved_active["d_celular"], "placa": saved_active["d_placa"], "marca": saved_active["d_marca"], "transportadora": saved_active["d_transp"]}
            elaborado_info = {"nombre": saved_active["elab_nombre"]}

            excel_bytes = generar_excel_bytes(df_resumen, total_docs, total_kg, total_ton, driver_info, dest_info, elaborado_info, saved_emp_active)
            pdf_bytes = generar_pdf_bytes(df_resumen, total_docs, total_kg, total_ton, driver_info, dest_info, elaborado_info, saved_emp_active)

            st.markdown("---")
            st.markdown("### 📥 Descargar Formato Oficial TUVACOL (GID-F-010)")
            col_exp1, col_exp2 = st.columns(2)
            with col_exp1:
                st.download_button("📊 Descargar Formato Excel (.xlsx)", data=excel_bytes, file_name=f"Relacion_Envio_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key=f"dl_excel_{tab_key}")
            with col_exp2:
                st.download_button("🖨️ Descargar Formato PDF (Imprimible)", data=pdf_bytes, file_name=f"Relacion_Envio_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.pdf", mime="application/pdf", use_container_width=True, key=f"dl_pdf_{tab_key}")
        else:
            st.warning("🔒 Descargas bloqueadas. Rellena todos los campos, especifica los empaques y presiona el botón **GUARDAR CAMBIOS Y HABILITAR DESCARGAS**.")

# ---------------------------------------------------------
# Interfaz Principal y Pestañas por Rol
# ---------------------------------------------------------
st.sidebar.markdown(f"👤 **Usuario:** {st.session_state.usuario_actual}")
st.sidebar.markdown(f"🏢 **Cliente:** {st.session_state.empresa_actual}")

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

if es_dev_autenticado:
    env_mode = st.sidebar.radio("Modo Conexión de Datos:", ["DEV (Google)", "PROD (SAP)"], key="env_mode")

# Sincronización en vivo con respaldo local
with st.spinner("🔄 Sincronizando base de datos de pesos..."):
    df_bd = fetch_google_sheet_database()

if st.sidebar.button("🔄 Sincronizar Datos"):
    st.cache_data.clear()
    st.rerun()

# --- TÍTULO PRINCIPAL ---
st.title("👍 Facilitador De Procesos Administrativos")

if es_dev_autenticado:
    tab1, tab2, tab3, tab_rutas, tab_extractor, tab_maestro, tab4 = st.tabs([
        "🔍 Búsqueda por Código", "📄 Procesar Remisión / PDF", "📤 Relación de Envío",
        "Gestión de Rutas", "Clasificador de Rutas", "🗂️ Registro Maestro", "📜 Logs"
    ])
else:
    tab1, tab2, tab3 = st.tabs([
        "🔍 Búsqueda por Código", "📄 Procesar Remisión / PDF", "📤 Relación de Envío"
    ])

# --- PESTAÑA 1 ---
with tab1:
    st.subheader("Consulta Dinámica de Producto")
    codigo_input = st.text_input("Ingrese Código o Descripción del Artículo", value="", placeholder="Ej. 108001051")
    cant_input = st.number_input("Cantidad a despachar", min_value=1.0, value=1.0, step=1.0)
    if codigo_input.strip() != "":
        item = get_product_data_from_source(codigo_input, df_bd)
        if item is not None:
            peso_unit = float(str(item.get('Peso_KG', 0.0)).replace(',', '.'))
            
            st.markdown(f"""
                <div style="background-color: #f0f2f6; padding: 15px; border-radius: 8px; border-left: 5px solid #1F3864; margin-bottom: 15px;">
                    <p style="margin: 0; font-size: 16px; color: #1F3864;"><b>📦 Código:</b> {item['Codigo']}</p>
                    <p style="margin: 5px 0 0 0; font-size: 16px; color: #000000;"><b>📝 Descripción Completa:</b> {item['Descripcion']}</p>
                </div>
            """, unsafe_allow_html=True)

            res_col1, res_col2 = st.columns(2)
            res_col1.metric("⚖️ Peso Unitario", f"{peso_unit:.2f} KG")
            res_col2.metric("📊 Peso Total", f"{peso_unit * cant_input:.2f} KG")
        else:
            st.warning("⚠️ No se encontraron coincidencias.")

# --- PESTAÑA 2 ---
with tab2:
    st.subheader("Búsqueda por Número de Entrega o Carga de PDF (Consulta)")
    
    if st.button("🗑️ Limpiar PDFs Subidos / Consultas", key="btn_clear_tab2", use_container_width=True):
        if "uploader_tab2_key" not in st.session_state:
            st.session_state["uploader_tab2_key"] = 0
        st.session_state["uploader_tab2_key"] += 1
        st.success("✅ ¡PDFs y consultas limpiados exitosamente!")
        st.rerun()

    modo_procesar = st.radio("Seleccione el método de entrada:", ["Buscar por Número de Entrega (Google Drive)", "Subir Archivo PDF Local"], horizontal=True)
    
    if modo_procesar == "Buscar por Número de Entrega (Google Drive)":
        num_entregas_input = st.text_input("Números de Entrega:", placeholder="Ej: 20005021, 3171")
        todos_los_items = []
        encontrados = []
        no_encontrados = []
        if num_entregas_input.strip():
            for num in [n.strip() for n in re.split(r'[\s,]+', num_entregas_input) if n.strip()][:50]:
                pdf_buf = descargar_pdf_desde_drive(DRIVE_FOLDER_ID, num)
                if pdf_buf:
                    items = extraer_tabla_materiales(pdf_buf, nombre_doc=f"Entrega_{num}")
                    if items:
                        todos_los_items.extend(items)
                        encontrados.append(num)
                    else: no_encontrados.append(num)
                else: no_encontrados.append(num)
            if encontrados: st.success(f"✅ Procesados: {', '.join(encontrados)}")
            if no_encontrados: st.error(f"❌ No encontrados: {', '.join(no_encontrados)}")
        
        if todos_los_items:
            fuentes_unicas = [(None, f"Entrega_{num}") for num in encontrados]
            render_consulta_despacho(fuentes_unicas)
            
    else:
        file_key_tab2 = st.session_state.get("uploader_tab2_key", 0)
        uploaded_files_tab2 = st.file_uploader("Cargar PDFs de Remisión", type=["pdf"], accept_multiple_files=True, key=f"uploader_tab2_local_{file_key_tab2}")
        lista_fuentes_local = []
        if uploaded_files_tab2:
            for f in uploaded_files_tab2:
                m = re.search(r'\d+', f.name)
                tag = m.group(0) if m else f.name
                lista_fuentes_local.append((f, f"Entrega_{tag}"))
        render_consulta_despacho(lista_fuentes_local)

# --- PESTAÑA 3 ---
with tab3:
    st.subheader("📤 Generador de Relación de Envío (Subir Archivos PDF)")
    
    if st.button("🗑️ Limpiar Formulario y PDFs Subidos", key="btn_clear_tab3", use_container_width=True):
        if "uploader_file_key" not in st.session_state:
            st.session_state["uploader_file_key"] = 0
        st.session_state["uploader_file_key"] += 1
        st.session_state["datos_guardados_tab3"] = False
        if "saved_data_tab3" in st.session_state: del st.session_state["saved_data_tab3"]
        if "saved_empaques_tab3" in st.session_state: del st.session_state["saved_empaques_tab3"]
        st.success("✅ ¡Formulario y archivos PDF limpiados exitosamente!")
        st.rerun()

    file_uploader_key = st.session_state.get("uploader_file_key", 0)
    uploaded_files = st.file_uploader("Cargar PDFs de Entrega", type=["pdf"], accept_multiple_files=True, key=f"uploader_tab3_{file_uploader_key}")
    
    lista_fuentes = []
    if uploaded_files:
        for f in uploaded_files:
            m = re.search(r'\d+', f.name)
            tag = m.group(0) if m else f.name
            lista_fuentes.append((f, f"Entrega_{tag}"))
    render_procesamiento_despacho(lista_fuentes, "tab3")

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
            if st.button("🗑️ Limpiar Registro Maestro"):
                os.remove(MASTER_CSV_PATH)
                st.success("Registro reiniciado.")
                st.rerun()
        else:
            st.warning("⚠️ No hay registros guardados.")

    with tab4:
        st.subheader("📜 Log Auditoría de Ejecución")
        if os.path.exists(LOG_FILENAME):
            with open(LOG_FILENAME, "r", encoding="utf-8") as f:
                log_content = f.read()
            st.download_button("💾 Descargar log", log_content, LOG_FILENAME, "text/plain")
            st.text_area("Contenido:", value=log_content, height=400)
