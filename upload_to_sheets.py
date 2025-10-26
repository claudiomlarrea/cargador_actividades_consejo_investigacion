# upload_to_sheets.py
from typing import List
import os
import pandas as pd

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


# Ruta del JSON de la cuenta de servicio.
# Si ya definiste la variable de entorno, la tomamos de ahí.
SA_PATH = os.getenv("GOOGLE_APPLICATION_CREDENTIALS", "service_account.json")


def _get_sheets_service():
    if not os.path.exists(SA_PATH):
        raise FileNotFoundError(
            f"No encuentro el JSON de la cuenta de servicio: {SA_PATH}\n"
            "Colócalo en la raíz o define GOOGLE_APPLICATION_CREDENTIALS."
        )
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = service_account.Credentials.from_service_account_file(SA_PATH, scopes=scopes)
    return build("sheets", "v4", credentials=creds)


def _ensure_sheet_exists(service, spreadsheet_id: str, sheet_name: str):
    """Crea la pestaña si no existe."""
    meta = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = meta.get("sheets", [])
    titles = [s["properties"]["title"] for s in sheets]

    if sheet_name not in titles:
        requests = [{
            "addSheet": {
                "properties": {"title": sheet_name}
            }
        }]
        body = {"requests": requests}
        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id, body=body
        ).execute()


def _clear_sheet(service, spreadsheet_id: str, sheet_name: str):
    service.spreadsheets().values().clear(
        spreadsheetId=spreadsheet_id, range=f"'{sheet_name}'!A:ZZ"
    ).execute()


def _df_to_values(df: pd.DataFrame) -> List[List[str]]:
    values = [list(df.columns)]
    for _, row in df.iterrows():
        values.append([None if pd.isna(v) else str(v) for v in row.tolist()])
    return values


def upload_dataframe_to_sheet(spreadsheet_id: str, sheet_name: str, df: pd.DataFrame, *, overwrite: bool = True):
    """
    Sube un DataFrame a Google Sheets.
    - Crea la pestaña si no existe.
    - Si overwrite=True, limpia el contenido antes de cargar.
    - Escribe encabezados y filas.
    """
    if df is None or df.empty:
        raise ValueError("El DataFrame está vacío; no hay datos para subir.")

    service = _get_sheets_service()
    _ensure_sheet_exists(service, spreadsheet_id, sheet_name)

    if overwrite:
        _clear_sheet(service, spreadsheet_id, sheet_name)

    values = _df_to_values(df)
    body = {"values": values}

    # Escribimos desde A1
    try:
        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"'{sheet_name}'!A1",
            valueInputOption="RAW",
            body=body
        ).execute()
    except HttpError as e:
        raise RuntimeError(f"Error subiendo datos a Sheets: {e}")


if __name__ == "__main__":
    # Ejemplo rápido (solo para probar):
    demo = pd.DataFrame({
        "Etiqueta": ["Proyecto", "Director", "Fecha"],
        "Valor": ["Comercio Justo...", "DIAZ BAY Javier", "19/11/2024"],
        "Confianza": [0.97, 0.98, 0.99],
    })
    # Reemplaza por tu ID de Google Sheets y nombre de pestaña
    upload_dataframe_to_sheet("TU_SPREADSHEET_ID", "Actas", demo)
    print("OK subido.")
