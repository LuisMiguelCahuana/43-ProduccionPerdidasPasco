import io
import re
import calendar
import unicodedata
from datetime import datetime

import pandas as pd
import streamlit as st
import gspread

from google.oauth2.service_account import Credentials


# ============================================================
# CONFIGURACIÓN GENERAL
# ============================================================

st.set_page_config(
    page_title="Consolidado de Producción y Costos",
    page_icon="📊",
    layout="wide"
)

SPREADSHEET_ID = "1gNq2nuNGfogcgcq7gsIqmQ5BZmMuWiSgAbDLpytJTIY"
HOJA_DESTINO = "CONSOLIDADO_GENERAL"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

PROCESOS = {
    "Actividad-CP": "11-12-13-14 Actividad CP",
    "Lecturas": "14 Lecturas",
    "Averías": "15. Averias",
    "Notificaciones": "16. Notific",
    "Itinerante": "17 Itinerant",
}

COLUMNAS_FINALES = [
    "periodo",
    "fecha",
    "suministro",
    "cuadrilla",
    "zona",
    "codigo",
    "cantidad",
    "actividad",
    "costo_unitario",
    "costo_total",
    "proceso",
    "hoja_origen",
    "archivo_origen",
]

# ------------------------------------------------------------
# ANCHOS FIJOS DE COLUMNAS PARA LA VISTA PREVIA
# Evita que las columnas se deformen al hacer clic sobre
# una fila o celda del st.dataframe.
# ------------------------------------------------------------

CONFIG_COLUMNAS = {
    "periodo": st.column_config.Column(
        "Periodo", width="small", pinned=True
    ),
    "fecha": st.column_config.Column("Fecha", width="small"),
    "suministro": st.column_config.Column("Suministro", width="small"),
    "cuadrilla": st.column_config.Column("Cuadrilla", width="medium"),
    "zona": st.column_config.Column("Zona", width="small"),
    "codigo": st.column_config.Column("Código", width="small"),
    "cantidad": st.column_config.NumberColumn(
        "Cantidad", width="small", format="%.2f"
    ),
    "actividad": st.column_config.Column(
        "Actividad", width="large"
    ),
    "costo_unitario": st.column_config.NumberColumn(
        "Costo unitario", width="small", format="%.2f"
    ),
    "costo_total": st.column_config.NumberColumn(
        "Costo total", width="small", format="%.2f"
    ),
    "proceso": st.column_config.Column("Proceso", width="medium"),
    "hoja_origen": st.column_config.Column(
        "Hoja origen", width="medium"
    ),
    "archivo_origen": st.column_config.Column(
        "Archivo origen", width="medium"
    ),
}


# ============================================================
# FUNCIONES GENERALES
# ============================================================
