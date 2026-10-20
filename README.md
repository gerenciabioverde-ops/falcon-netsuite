# Falcon · Generador de CSV para NetSuite

App independiente (FastAPI + Jinja2 + HTMX). Recibe un Excel/CSV con la información,
la valida contra `data/TEMPLATE_CSV_FALCON.xlsx` y entrega los CSV de importación de NetSuite.

## Uso
En Windows: doble clic en `iniciar.bat` (crea el entorno, instala dependencias, arranca y abre el navegador). Usa el primer puerto libre desde el 8002, saltando los reservados por otras apps Falcon (8000 Contable, 8001 Tesorería, 8003 Falcon PDB; editar `RESERVADOS` en `iniciar.bat` para agregar más). Mapa de puertos: Contable 8000 · Tesorería 8001 · NetSuite 8002 · PDB 8003.

`actualizar.bat` baja la última versión de `main` y actualiza dependencias; `detener.bat` cierra el servidor de esta carpeta. Para actualizar la app se usa `actualizar.bat` (ya no hay botón «Actualizar» en la pantalla).

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
| Asiento | `FCP{AAAAMM}_{0001}`: líneas consecutivas sin ID hasta que Débito = Crédito forman un asiento; el correlativo es por mes de la fecha y continúa tras los IDs ya escritos | `FCP202610_0001` |

Si se escribe un EXTERNAL ID a mano, se respeta tal cual.

**Filas bloqueadas:** las filas que se cargan desde un Excel/CSV quedan de solo lectura (fondo gris); para modificar un registro se usa «✎ Editar» (o doble clic en la celda). Las filas que se escriben o pegan a mano siguen siendo editables en la grilla.

**Editar fila:** cada fila de las grillas tiene un botón «✎ Editar» que abre un formulario con todos los campos (con sus listas desplegables): permite «Autocompletar» los campos vacíos que la app calcula (ID, ciudad, ZIP, persona…), «Validar fila», moverse con «Anterior / Siguiente», «Guardar», «Cancelar» y «Eliminar fila». Esc cierra sin guardar.

**Gastos:** pestaña para el reporte de gastos (`reporte_gastos_AAAA-MM-DD.xls`, que el sistema exporta como tabla HTML con extensión `.xls`; también acepta `.xlsx` y `.csv`). 27 columnas (External ID automático = `VENDB_RUC_Serie-Número`, luego Fecha Factura … Caja y **Met. Pago**: «Transferencia de Fondos» si Caja trae un número de cuenta, «Efectivo» si dice Caja Chica; **Detracción**: columnas «Sujeto a Detracción» (Sí/No), «Concepto Detracción» (tabla de la hoja DATA MAESTRA de la plantilla, con su porcentaje) y «% Detracción»; el «Monto Detr.» = Total × % (2 decimales, en la moneda del comprobante). Si el reporte ya trae Monto Detr. mayor que 0 se marca «Sí» sola; «💡 Proponer…» sugiere el concepto (transporte de bienes 27, movimiento de carga 21, otros servicios empresariales 22, arrendamiento 19, etc.). Valida: concepto obligatorio si es «Sí», avisa si el monto difiere del cálculo, si el IGV es 0 o si hay Monto Detr. sin marcar «Sí». Pasa a `Latam - Concepto Detracción` de Factura de Proveedor («SIN DETRACCIÓN» si es «No»). Luego **Departamento** (lista: Operaciones / Sostenibilidad / Finanzas / Administrativos, constante `DEPARTAMENTOS` en `app/esquema.py`; «💡 Proponer cuenta y departamento» sugiere según el Concepto, reglas en `app/gastos_cuentas.py`) y **Clase** (se escribe; al ponerlos en un gasto se ofrece copiarlos a los demás gastos del mismo concepto que los tienen vacíos, y pasan a `Gastos: Departamento` / `Gastos: Clase` de Factura de Proveedor); se completa sola y «Pasar a Factura de Proveedor» la lleva a `Latam - Método de Pago`), se ignora la fila de Totales. Valida RUC (11 dígitos con dígito verificador) o DNI (8), fechas, números con separador de miles (`1,308.00`), moneda (PEN/USD), que Sub-Total + IGV = Total y que no se repita RUC + Serie-Número; avisa si falta la fecha de factura o de pago y si el Monto Pagado supera el Total. Un `.xls` binario antiguo debe guardarse antes como `.xlsx`. El botón «➜ Pasar a Factura de Proveedor» crea una fila por gasto en esa pestaña con el mismo External ID (si la factura ya existe con ese ID, solo le completa las celdas vacías con lo de Gastos; para traer lo que cambió, usa «🔄 Actualizar desde Gastos» en la pestaña Factura de Proveedor) (Proveedor = `VEND_<RUC>`, moneda PEN→SOL, serie y número separados, Sub-Total como importe de gasto); quedan por completar Cuenta y Código de impuesto. Las filas que vienen de Gastos quedan **bloqueadas** (solo lectura): se modifican únicamente con «✎ Editar» o doble clic. **🔄 Actualizar desde Gastos** (en Factura de Proveedor) compara todas las facturas con Gastos por External ID, muestra cuántas cambiarían (con ejemplos) y, al confirmar, reemplaza solo los campos que vienen de Gastos (proveedor, moneda, fecha, tipo, serie/número, método de pago, cuenta, importe, nota) y agrega las facturas nuevas; cuenta por pagar, impuestos, departamento y clase no se tocan. La **Subsidiaria** (43) y la **Ubicación** (`FCP SOLES`) —como en las filas de ejemplo de la plantilla; constantes `SUBSIDIARIA` y `UBICACION` en `app/esquema.py`— se completan solas si las dejas vacías: Subsidiaria en las facturas, notas de crédito, pagos y asientos; Ubicación (y la de línea, si la línea tiene importe) en las facturas, notas de crédito y pagos (`LOCATION`). Departamento y Clase no se completan. Esos valores se ven ya en la grilla, y las fechas se muestran como en Excel (`9/10/2026`, D/M/AAAA) aunque el reporte las traiga como `2026-10-09`. Gastos tiene las columnas **Cuenta PCGE** (se elige; lista de las cuentas 6xxxxx de «Cuentas homologadas») y **Cuenta NetSuite** (automática: el número de la cuenta en el COA de NetSuite según «Cuentas homologadas», p. ej. PCGE 609219 → `9011.06`; al pasar a Factura de Proveedor se envía su ID interno en *Gastos: Cuenta*). «💡 Proponer cuentas» sugiere la cuenta por Concepto (`app/gastos_cuentas.py`) en amarillo para revisar; al elegir una cuenta se ofrece aplicarla a los demás gastos del mismo concepto.

**Cargar Excel / CSV:** cada pestaña con grilla tiene el botón «⬆ Cargar Excel / CSV». Lee el archivo en el servidor y llena la(s) grilla(s): del Excel de plantillas carga las 6 hojas, del libro de proveedores toma la hoja `Vendor` (encabezados en la fila 3; ignora las filas de ayuda y de ejemplo) y un CSV va a la pestaña desde la que se carga. Las columnas se buscan por nombre. Hasta 2000 filas por hoja y 40 MB; después se revisa y se pulsa «Validar y generar».

**Períodos:** arriba hay un selector de mes (◀ / calendario / ▶ y la lista «Períodos con datos…»). Cada período guarda por separado todas sus pestañas (en `datos_usuario/periodos/AAAA-MM.json`, en el equipo donde corre la app; fuera de git y sin que «Actualizar» lo toque). Se guarda solo unos instantes después de cada cambio y al volver a un período se muestran sus datos; la app recuerda el último período abierto. El período también se usa al validar: rellena `POSTING PERIOD` (formato `Oct 2026`) si está vacío, avisa si el período contable o la fecha del documento no corresponden al mes seleccionado, y nombra el zip `netsuite_csv_AAAA-MM.zip`. «Vaciar período» borra solo los datos del mes elegido. El borrador anterior del navegador se pasa automáticamente al primer período que se abra.

**Proveedores:** réplica de la hoja `Vendor` del libro *Entity - Falcon Coffees - Perú - Legacy* (22 columnas, mismas listas desplegables). Por proveedor basta con **Nombre, Dirección 1, Ubigeo, Tipo de documento y Número**; lo demás se completa solo, igual que en el libro: `External ID` = `VEND_{número}`, Ciudad y Estado desde el ubigeo, Código ZIP (primera coincidencia de la clave `DEPARTAMENTO-PROVINCIA-DISTRITO` en la tabla de códigos postales), País/Subsidiaria/Entidad = `Peru`, Tipo de Persona (DNI → Natural; RUC 20 → Jurídico; RUC 10 → Natural), Dígito Verificador (último dígito del RUC) y las casillas T/F en `F`. Valida RUC (11 dígitos + dígito verificador), DNI (8 dígitos), listas y proveedores repetidos. Las listas (países, ubigeos, tipos de documento y códigos postales) están en `data/proveedores_listas.json`; para regenerarlas desde el libro: `python scripts/extraer_listas_proveedores.py ruta/al/libro.xlsx` (solo extrae listas, no datos de proveedores).

**Cuentas homologadas:** la pestaña del mismo nombre recibe un Excel/CSV con las columnas `Cuenta Falcon` e `ID interno NetSuite` (más `Nombre Falcon` / `Nombre NetSuite` opcionales; hay plantilla descargable). También acepta directamente el Excel del COA de NetSuite con la hoja de equivalencia PCGE (columnas `COA - NetSuite` y `Peru Tax Code`): cruza PCGE → número de cuenta NetSuite → ID interno, guarda también las «Cuentas Destino» del PCGE (asiento de destino: cuenta 94/95/97 y contrapartida) solo para consulta —no se envían a NetSuite— y avisa de cuentas inexistentes, inactivas o de resumen. Los nombres de las cuentas PCGE se muestran en español (`data/nombres_pcge_es.json`; las que no estén ahí se ven como vienen en la tabla), y el nombre de la cuenta de NetSuite se deja tal cual está en tu COA. La tabla se guarda en `datos_usuario/` (fuera de git, no la pisa «Actualizar»). Al generar los CSV, las columnas de cuenta (`Cuenta`, `Gastos: Cuenta`, `Account`, `A/P ACCOUNT`, `ACCOUNT`) que traigan un código de Falcon se reemplazan por su ID interno de NetSuite, y cada traducción aparece como aviso en el resultado; lo que no está en la tabla se deja igual.

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
