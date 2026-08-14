import requests
import streamlit as st
import logging

@st.cache_data(ttl=3600)  # Caché del token por 1 hora para optimizar peticiones
def obtener_token_sap():
    """
    Solicita un Access Token OAuth 2.0 a SAP BTP utilizando 
    las credenciales almacenadas en st.secrets.
    """
    try:
        sap_config = st.secrets["sap_s4hana"]
        token_url = sap_config["token_url"]
        client_id = sap_config["client_id"]
        client_secret = sap_config["client_secret"]
        
        payload = {
            "grant_type": "client_credentials"
        }
        
        # Petición POST autenticada por Basic Auth con el Client ID y Client Secret
        response = requests.post(
            token_url,
            data=payload,
            auth=(client_id, client_secret),
            timeout=10
        )
        
        if response.status_code == 200:
            token_data = response.json()
            access_token = token_data.get("access_token")
            logging.info("[SAP AUTH] Token OAuth 2.0 obtenido exitosamente.")
            return access_token
        else:
            logging.error(f"[SAP AUTH Error] Código HTTP {response.status_code}: {response.text}")
            return None
            
    except Exception as e:
        logging.error(f"[SAP AUTH Excepción] No se pudo conectar al servidor de autenticación: {e}")
        return None

def hacer_petecion_odata(endpoint_path):
    """
    Realiza una petición segura GET a una API OData de SAP S/4HANA Cloud 
    inyectando el Token de Acceso corporativo.
    """
    token = obtener_token_sap()
    if not token:
        st.error("❌ Error de autenticación con SAP BTP. Verifique los logs.")
        return None
        
    sap_config = st.secrets["sap_s4hana"]
    base_url = sap_config["service_root_url"]
    url_completa = f"{base_url.rstrip('/')}/{endpoint_path.lstrip('/')}"
    
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json"
    }
    
    try:
        response = requests.get(url_completa, headers=headers, timeout=15)
        if response.status_code == 200:
            return response.json()
        else:
            logging.error(f"[SAP OData Error] {response.status_code} en {url_completa}")
            return None
    except Exception as e:
        logging.error(f"[SAP OData Excepción] {e}")
        return None