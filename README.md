# Falcon · Generador de CSV para NetSuite

App independiente (FastAPI + Jinja2 + HTMX). Recibe un Excel/CSV con la información,
la valida contra `data/TEMPLATE_CSV_FALCON.xlsx` y entrega los CSV de importación de NetSuite.

## Uso
En Windows: doble clic en `iniciar.bat` (crea el entorno, instala dependencias, arranca y abre el navegador). Usa el primer puerto libre desde el 8002, saltando los reservados por otras apps Falcon (8000 Contable, 8001 Tesorería, 8003 Falcon PDB; editar `RESERVADOS` en `iniciar.bat` para agregar más). Mapa de puertos: Contable 8000 · Tesorería 8001 · NetSuite 8002 · PDB 8003.

`actualizar.bat` baja la última versión de `main` y actualiza dependencias; `detener.bat` cierra el servidor de esta carpeta. El botón «⟳ Actualizar» (arriba a la derecha) trae la última versión de `main` desde GitHub y reinicia la app sola (requiere haber abierto con `iniciar.bat`).

Manual:
```
pip install -r requirements.txt
uvicorn app.main:app --port 8002 --reload      # http://localhost:8002
```
1. Escribir o pegar (Ctrl+V desde Excel) los datos en la pestaña de cada plantilla.
2. Pulsar «Validar y generar», corregir errores y descargar cada CSV o el `.zip`.

El **EXTERNAL ID se genera solo** si se deja vacío, y es siempre el mismo para el mismo comprobante (así NetSuite no lo duplica al recargarlo):

| Plantilla | Formato | Ejemplo |
|---|---|---|
| Factura de Proveedor | `VENDB_{Proveedor}_{Serie}_{Número}` | `VENDB_296720_F001_12` |
| Pago Proveedor | `BP_PE_{Beneficiario}_{BILL}_{AAAAMMDD}` | `BP_PE_296720_25114824_20261009` |
| Factura de Cliente | `FACT_VENTA_{Cliente}_{Serie}_{Número}` | |
| Nota de crédito Cliente | `NotaCredito_{Cliente}_{Serie}_{Número}` | |
| Pago de factura de cliente | `Pago_FCliente_{Beneficiario}_{Aplica}_{AAAAMMDD}` | |
| Asiento | `Asiento_{AAAAMMDD}_{huella}`: líneas consecutivas sin ID hasta que Débito = Crédito | `Asiento_20261013_17b853ed` |

Si se escribe un EXTERNAL ID a mano, se respeta tal cual.

**Editar fila:** cada fila de las grillas tiene un botón «✎ Editar» que abre un formulario con todos los campos (con sus listas desplegables): permite «Autocompletar» los campos vacíos que la app calcula (ID, ciudad, ZIP, persona…), «Validar fila», moverse con «Anterior / Siguiente», «Guardar», «Cancelar» y «Eliminar fila». Esc cierra sin guardar.

**Cargar Excel / CSV:** cada pestaña con grilla tiene el botón «⬆ Cargar Excel / CSV». Lee el archivo en el servidor y llena la(s) grilla(s): del Excel de plantillas carga las 6 hojas, del libro de proveedores toma la hoja `Vendor` (encabezados en la fila 3; ignora las filas de ayuda y de ejemplo) y un CSV va a la pestaña desde la que se carga. Las columnas se buscan por nombre. Hasta 2000 filas por hoja y 40 MB; después se revisa y se pulsa «Validar y generar».

**Proveedores:** réplica de la hoja `Vendor` del libro *Entity - Falcon Coffees - Perú - Legacy* (22 columnas, mismas listas desplegables). Por proveedor basta con **Nombre, Dirección 1, Ubigeo, Tipo de documento y Número**; lo demás se completa solo, igual que en el libro: `External ID` = `VEND_{número}`, Ciudad y Estado desde el ubigeo, Código ZIP (primera coincidencia de la clave `DEPARTAMENTO-PROVINCIA-DISTRITO` en la tabla de códigos postales), País/Subsidiaria/Entidad = `Peru`, Tipo de Persona (DNI → Natural; RUC 20 → Jurídico; RUC 10 → Natural), Dígito Verificador (último dígito del RUC) y las casillas T/F en `F`. Valida RUC (11 dígitos + dígito verificador), DNI (8 dígitos), listas y proveedores repetidos. Las listas (países, ubigeos, tipos de documento y códigos postales) están en `data/proveedores_listas.json`; para regenerarlas desde el libro: `python scripts/extraer_listas_proveedores.py ruta/al/libro.xlsx` (solo extrae listas, no datos de proveedores).

**Cuentas homologadas:** la pestaña del mismo nombre recibe un Excel/CSV con las columnas `Cuenta Falcon` e `ID interno NetSuite` (más `Nombre Falcon` / `Nombre NetSuite` opcionales; hay plantilla descargable). También acepta directamente el Excel del COA de NetSuite con la hoja de equivalencia PCGE (columnas `COA - NetSuite` y `Peru Tax Code`): cruza PCGE → número de cuenta NetSuite → ID interno, guarda también las «Cuentas Destino» del PCGE (asiento de destino: cuenta 94/95/97 y contrapartida) solo para consulta —no se envían a NetSuite— y avisa de cuentas inexistentes, inactivas o de resumen. La tabla se guarda en `datos_usuario/` (fuera de git, no la pisa «Actualizar»). Al generar los CSV, las columnas de cuenta (`Cuenta`, `Gastos: Cuenta`, `Account`, `A/P ACCOUNT`, `ACCOUNT`) que traigan un código de Falcon se reemplazan por su ID interno de NetSuite, y cada traducción aparece como aviso en el resultado; lo que no está en la tabla se deja igual.

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
