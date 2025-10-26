# upload_to_sheets.py
# --------------------------------------------------------
# Sube un DataFrame a Google Sheets usando la cuenta
# de servicio que pegaste en Streamlit Secrets.
# --------------------------------------------------------

import streamlit as st
import pandas as pd

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2 import service_account


def _get_gcp_credentials(scopes=None):
    creds = service_account.Credentials.from_service_account_info(
        st.secrets["gcp_service_account"]
    )
    if scopes:
        creds = creds.with_scopes(scopes)
    return creds


def _ensure_sheet_exists(service, spreadsheet_id: str, sheet_name: str):
    meta = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    titles = [s["properties"]["title"] for s in meta.get("sheets", [])]
    if sheet_name not in titles:
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": [{"addSheet": {"properties": {"title": sheet_name}}}]}
        ).execute()


def upload_dataframe_to_sheet(df: pd.DataFrame, spreadsheet_id: str = None, sheet_name: str = None) -> bool:
    """
    Sube df al Google Sheet indicado.
    Si no pasas IDs, toma los de [sheets] en secrets.
    Requiere que hayas compartido el Sheet con la service account como Editor.
    """
    if df is None or df.empty:
        st.error("El DataFrame está vacío, no hay nada para subir.")
        return False

    if spreadsheet_id is None:
        spreadsheet_id = st.secrets["sheets"]["spreadsheet_id"]
    if sheet_name is None:
        sheet_name = st.secrets["sheets"]["sheet_name"]

    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = _get_gcp_credentials(scopes=scopes)
    service = build("sheets", "v4", credentials=creds)

    try:
        # Crea la pestaña si no existe
        _ensure_sheet_exists(service, spreadsheet_id, sheet_name)

        # Limpia contenido previo
        service.spreadsheets().values().clear(
            spreadsheetId=spreadsheet_id,
            range=f"'{sheet_name}'!A:ZZ"
        ).execute()

        # Escribe encabezados + filas
        values = [list(df.columns)] + df.astype(str).values.tolist()
        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"'{sheet_name}'!A1",
            valueInputOption="RAW",
            body={"values": values},
        ).execute()

        return True

    except HttpError as e:
        st.error(f"Error subiendo a Sheets: {e}")
        return False
