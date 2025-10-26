# 🏛️ Extractor de Actas del Consejo de Investigación – UCCuyo

Esta aplicación automatiza la lectura, clasificación y carga de información proveniente de las **Actas del Consejo de Investigación** de la Universidad Católica de Cuyo.  
Fue desarrollada con **Streamlit**, **Google Cloud Document AI** y conexión directa a **Google Sheets** mediante cuentas de servicio seguras.

---

## 🚀 Funcionalidad

El sistema permite:

- Procesar archivos **PDF, DOCX o TXT** de órdenes del día o actas.
- Identificar automáticamente campos como:
  - Año / Fecha  
  - Título o denominación del proyecto  
  - Director / Codirector / Equipo  
  - Facultad o Unidad Académica  
  - Estado o tipo de informe (avance, final, categorización, etc.)
- Exportar automáticamente los resultados a:
  - **Excel (.xlsx)**  
  - **CSV (.csv)**  
  - **Google Sheets** (conectado a la cuenta institucional)
- Integrarse con tableros de control institucionales (Looker Studio, BigQuery, etc.)

---

## 🧠 Arquitectura

- **Frontend:** Streamlit  
- **Backend:** Python 3.11  
- **Procesamiento:** Google Cloud Document AI (Custom Extractor)  
- **Almacenamiento:** Google Drive / Google Sheets  
- **Exportación:** Excel y CSV  
- **Infraestructura:** Streamlit Cloud + GitHub

---

## ⚙️ Configuración de entorno

### 1️⃣ Requisitos de instalación
El archivo `requirements.txt` ya incluye todas las dependencias necesarias:

```bash
pip install -r requirements.txt
