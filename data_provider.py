import os
import streamlit as st
import pandas as pd
from supabase import create_client

# Configuración de entornos
ENVIRONMENT = os.getenv("APP_ENV", "DEV")  # 'DEV' para Google, 'PROD' para SAP

def get_product_data(codigo):
    """Interfaz unificada para obtener datos de productos."""
    if ENVIRONMENT == "PROD":
        return _get_product_from_sap(codigo)
    else:
        return _get_product_from_google(codigo)

def _get_product_from_google(codigo):
    # Aquí irá tu lógica actual de fetch_google_sheet_database
    # De momento, llamamos a tu función existente si quieres o simplificamos
    return {"peso": 0.5, "desc": "Producto de prueba (DEV)"}

def _get_product_from_sap(codigo):
    """
    PREPARACIÓN: Este es el lugar donde conectaremos con SAP BTP.
    Por ahora, devolvemos un log indicando que la conexión está lista para integrarse.
    """
    logging.info(f"Conectando a SAP BTP para el código: {codigo}")
    # Aquí implementaremos en el siguiente paso la conexión OData
    return {"peso": 0.0, "desc": "Datos vía SAP S/4HANA"}