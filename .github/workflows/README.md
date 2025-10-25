# Pipeline Actas → Google Sheets → Looker Studio

## 1) Requisitos
- Cuenta de Google con acceso a Google Sheets.
- Service Account (Proyecto Google Cloud) con rol de **Editor** en Sheets.
- Compartir el Spreadsheet con el email de la Service Account.

## 2) Instalación local
```bash
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
export GOOGLE_APPLICATION_CREDENTIALS=credenciales.json
