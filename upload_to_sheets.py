# -*- coding: utf-8 -*-
"""
Sube el CSV generado por extract_actas.py a Google Sheets.
Crea la hoja si no existe y reemplaza el contenido completo.

Variables de entorno:
- OUTPUT_CSV (default: actas_extraccion.csv)
- SHEET_NAME (default: Base Consejo de Investigación)
- WORKSHEET_NAME (default: Actas)
- GOOGLE_APPLICATION_CREDENTIALS (ruta a credenciales.json de Service Account)
"""
import os
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

OUTPUT_CSV = os.environ.get("OUTPUT_CSV", "actas_extraccion.csv")
SHEET_NAME = os.environ.get("SHEET_NAME", "Base Consejo de Investigación")
WORKSHEET_NAME = os.environ.get("WORKSHEET_NAME", "Actas")

SCOPE = ["https://www.googleapis.com/auth/spreadsheets"]

def get_client():
    cred_path = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "credenciales.json")
    creds = Credentials.from_service_account_file(cred_path, scopes=SCOPE)
    return gspread.authorize(creds)

def ensure_worksheet(sh, title):
    try:
        ws = sh.worksheet(title)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=title, rows=1000, cols=20)
    return ws

def main():
    df = pd.read_csv(OUTPUT_CSV)
    client = get_client()

    try:
        sh = client.open(SHEET_NAME)
    except gspread.exceptions.SpreadsheetNotFound:
        sh = client.create(SHEET_NAME)

    ws = ensure_worksheet(sh, WORKSHEET_NAME)
    ws.clear()
    ws.update([df.columns.tolist()] + df.astype(str).values.tolist())
    print(f"✅ Google Sheets actualizado: {SHEET_NAME} / {WORKSHEET_NAME} ({len(df)} filas)")

if __name__ == "__main__":
    main()
