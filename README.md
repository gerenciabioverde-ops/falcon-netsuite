# Falcon · Generador de CSV para NetSuite

App independiente (FastAPI + Jinja2 + HTMX). Recibe un Excel/CSV con la información,
la valida contra `data/TEMPLATE_CSV_FALCON.xlsx` y entrega los CSV de importación de NetSuite.

## Uso
En Windows: doble clic en `iniciar.bat` (crea el entorno, instala dependencias, arranca y abre el navegador).

Manual:
```
pip install -r requirements.txt
uvicorn app.main:app --reload      # http://localhost:8000
```
1. Descargar la plantilla desde la pantalla principal y llenarla.
2. Subir el `.xlsx` (procesa todas las hojas) o un `.csv` (indicando la plantilla).
3. Corregir errores y descargar cada CSV o el `.zip`.

## Qué valida
- Encabezados idénticos a la plantilla.
- Campos obligatorios, fechas, números y formato de periodo (`Oct 2026`).
- Listas desplegables y catálogos de `DATA MAESTRA` (se leen del propio Excel).
- Asientos cuadrados (Débito = Crédito por EXTERNAL ID).
- Fórmulas simples (`=Y2*AA2`): usa el valor guardado por Excel o las evalúa.
- Líneas con el mismo EXTERNAL ID forman un documento; avisa si la cabecera difiere.

## Ajustes
- Obligatoriedad por plantilla: `REGLAS` en `app/esquema.py` (son una primera propuesta; ajustar a lo que exija NetSuite).
- Cambiar listas/columnas: editar `data/TEMPLATE_CSV_FALCON.xlsx`; el esquema se lee de ahí al arrancar.
- Formato de fecha del CSV: se elige en pantalla (debe coincidir con la preferencia de fecha de NetSuite).
- CSV de salida: UTF-8 con BOM, separador coma.
