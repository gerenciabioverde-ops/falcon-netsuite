# Falcon · Generador de CSV para NetSuite

App independiente (FastAPI + Jinja2 + HTMX). Recibe un Excel/CSV con la información,
la valida contra `data/TEMPLATE_CSV_FALCON.xlsx` y entrega los CSV de importación de NetSuite.

## Uso
En Windows: doble clic en `iniciar.bat` (crea el entorno, instala dependencias, arranca y abre el navegador). Usa el primer puerto libre desde el 8002, saltando los reservados por otras apps Falcon (8000 Contable, 8001 Tesorería; editar `RESERVADOS` en `iniciar.bat` para agregar más, p. ej. Falcon PDB).

El botón «⟳ Actualizar» (arriba a la derecha) trae la última versión de `main` desde GitHub y reinicia la app sola (requiere haber abierto con `iniciar.bat`).

Manual:
```
pip install -r requirements.txt
uvicorn app.main:app --port 8002 --reload      # http://localhost:8002
```
1. Escribir o pegar (Ctrl+V desde Excel) los datos en la pestaña de cada plantilla.
2. Pulsar «Validar y generar», corregir errores y descargar cada CSV o el `.zip`.

Las pestañas guardan un borrador en el navegador. «Artículos: Importe» se calcula (cantidad × tarifa) si se deja vacío.

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
