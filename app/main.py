"""Aplicativo para generar los CSV de carga de NetSuite a partir de un Excel/CSV."""
from __future__ import annotations

import os
import re
import secrets
import subprocess
import time
from pathlib import Path

from fastapi import Body, FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, Response
from fastapi.templating import Jinja2Templates

from starlette.concurrency import run_in_threadpool

from . import carga, gastos_cuentas, homologacion, padron, periodos, proveedores
from .esquema import CUENTA_POR_PAGAR, CUENTA_POR_PAGAR_RHE, DETRACCIONES, PAGAR_POR_TIPO, RUTA_PLANTILLA, SUBSIDIARIA, UBICACION, cargar_plantilla, norm
from .procesador import FORMATOS_FECHA, ResultadoHoja, empaquetar_zip, leer_csv, leer_xlsx, procesar, reporte_errores_xlsx, validar_hoja

BASE = Path(__file__).parent
RAIZ = BASE.parent
app = FastAPI(title="Falcon NetSuite CSV")
templates = Jinja2Templates(directory=str(BASE / "templates"))
PLANTILLA = cargar_plantilla()

# Resultados en memoria (un servidor, uso interno): token -> (timestamp, resultados)
_RESULTADOS: dict[str, tuple[float, list[ResultadoHoja]]] = {}
_TTL = 3600
_MAX_BYTES = 10 * 1024 * 1024
_MAX_CARGA = 40 * 1024 * 1024  # el libro de carga completo (con la tabla de códigos postales) pesa ~8 MB


def _guardar(resultados: list[ResultadoHoja], periodo: str = "") -> str:
    ahora = time.time()
    for k in [k for k, v in _RESULTADOS.items() if ahora - v[0] > _TTL]:
        del _RESULTADOS[k]
    token = secrets.token_urlsafe(12)
    _RESULTADOS[token] = (ahora, resultados, periodo)
    return token


def _obtener(token: str) -> list[ResultadoHoja]:
    if token not in _RESULTADOS:
        raise HTTPException(404, "El resultado expiró; vuelva a procesar el archivo.")
    return _RESULTADOS[token][1]


def esquema_detracciones() -> dict[str, float]:
    return dict(DETRACCIONES)


def _pestanas() -> list[dict]:
    """Descripción de cada plantilla para dibujar su pestaña con grilla."""
    salida = []
    for clave, h in PLANTILLA.hojas.items():
        req = {norm(c) for c in h.reglas.get("requeridos", [])}
        fechas = {norm(c) for c in h.reglas.get("fechas", [])}
        numeros = {norm(c) for c in h.reglas.get("numericos", [])} | ({norm("Neto a Pagar")} if clave == "gastos" else set())
        salida.append({
            "clave": clave,
            "nombre": h.nombre.strip(),
            "nota": h.reglas.get("nota", ""),
            "columnas": [{
                "nombre": e.strip(),
                "auto": i == 0 and bool(h.reglas.get("id_auto")),
                "requerida": norm(e) in req and not (i == 0 and h.reglas.get("id_auto")),
                "opciones": h.listas[i].opciones if i in h.listas else [],
                "fecha": norm(e) in fechas,
                "numero": norm(e) in numeros,
            } for i, e in enumerate(h.encabezados)],
            "detracciones": esquema_detracciones() if clave == "gastos" else {},
            "por_defecto": {c: list(v) if isinstance(v, tuple) else [v, ""] for c, v in h.reglas.get("por_defecto", {}).items()},
        })
    return salida


def _git(*args: str) -> subprocess.CompletedProcess:
    return subprocess.run(["git", *args], cwd=RAIZ, capture_output=True, text=True, timeout=90,
                          encoding="utf-8", errors="replace",
                          env={**os.environ, "GIT_TERMINAL_PROMPT": "0"})


def _version() -> str:
    try:
        r = _git("log", "-1", "--format=%h|%cd", "--date=format:%d/%m/%Y %H:%M")
        return r.stdout.strip().replace("|", " - ") if r.returncode == 0 else ""
    except Exception:
        return ""


@app.get("/", response_class=HTMLResponse)
def inicio(request: Request):
    pestanas = _pestanas()
    return templates.TemplateResponse(request, "index.html", {
        "version": _version(),
        "hojas": [p["nombre"] for p in pestanas],
        "pestanas": pestanas,
        "formatos": list(FORMATOS_FECHA),
        "subsidiaria": SUBSIDIARIA,
        "ubicacion": UBICACION,
        "pagar_por_tipo": PAGAR_POR_TIPO,
        "cuenta_por_pagar": CUENTA_POR_PAGAR,
        "cuenta_por_pagar_rhe": CUENTA_POR_PAGAR_RHE,
    })


@app.get("/plantilla")
def descargar_plantilla():
    return FileResponse(RUTA_PLANTILLA, filename="TEMPLATE_CSV_FALCON.xlsx")


def _solo_local(request: Request) -> None:
    if not request.client or request.client.host not in ("127.0.0.1", "::1"):
        raise HTTPException(403, "Solo desde el equipo donde corre la app.")


@app.post("/gastos/proponer")
async def proponer_cuentas(request: Request):
    """{'conceptos': [...]} -> {'propuestas': {concepto: código de cuenta o ''}, 'departamentos': {concepto: departamento o ''}}."""
    cuerpo = await request.json()
    conceptos = [str(c) for c in (cuerpo.get("conceptos") or [])][:500]
    return {"propuestas": {c: gastos_cuentas.proponer(c) for c in conceptos},
            "departamentos": {c: gastos_cuentas.proponer_departamento(c) for c in conceptos},
            "detracciones": {c: gastos_cuentas.proponer_detraccion(c) for c in conceptos},
            "clasificaciones": {c: gastos_cuentas.proponer_clasificacion(c) for c in conceptos}}


@app.get("/homologacion")
def ver_homologacion():
    d = homologacion.cargar()
    es = homologacion.nombres_es()
    filas = [{**f, "nombre_es": f.get("nombre_es") or es.get(f["falcon"]) or f.get("nombre_falcon", "")} for f in d["filas"]]
    return {"archivo": d["archivo"], "actualizado": d["actualizado"], "cantidad": len(d["filas"]), "filas": filas}


@app.post("/homologacion")
async def subir_homologacion(request: Request, archivo: UploadFile = File(...)):
    _solo_local(request)
    contenido = await archivo.read(_MAX_BYTES + 1)
    if len(contenido) > _MAX_BYTES:
        raise HTTPException(413, "El archivo supera 10 MB")
    try:
        filas, errores, avisos = homologacion.procesar_archivo(archivo.filename or "", contenido)
    except ValueError as e:
        return JSONResponse({"ok": False, "mensaje": str(e), "errores": [], "avisos": []})
    except Exception:
        return JSONResponse({"ok": False, "mensaje": "No se pudo leer el archivo; verifique que no esté dañado.", "errores": [], "avisos": []})
    if errores:  # no se guarda una tabla ambigua o incompleta
        return JSONResponse({"ok": False, "mensaje": f"Corrija {len(errores)} problema(s) y vuelva a subir.",
                             "errores": errores[:50], "avisos": avisos[:50]})
    homologacion.guardar(filas, archivo.filename or "")
    return {"ok": True, "mensaje": f"{len(filas)} cuentas cargadas.", "errores": [], "avisos": avisos[:50]}


@app.put("/homologacion/cuenta")
async def guardar_cuenta(request: Request):
    """Crea o modifica una cuenta: {falcon, nombre_es, coa, nombre_ns, netsuite (ID interno), destino, contrapartida, original}."""
    _solo_local(request)
    c = await request.json()
    falcon, idn = str(c.get("falcon", "")).strip(), str(c.get("netsuite", "")).strip()
    if not falcon or not idn:
        raise HTTPException(400, "Falta la cuenta PCGE o el ID interno de NetSuite.")
    if not re.fullmatch(r"\d+", idn):
        raise HTTPException(400, "El ID interno de NetSuite debe ser un número.")
    coa = str(c.get("coa", "")).strip()
    nombre = str(c.get("nombre_es", "")).strip()
    otra = next((f for f in homologacion.cargar()["filas"] if homologacion.normalizar(f["falcon"]) == homologacion.normalizar(falcon)
                 and homologacion.normalizar(f["falcon"]) != homologacion.normalizar(str(c.get("original", "")))), None)
    if otra:
        raise HTTPException(400, f"La cuenta {falcon} ya existe: edítela en lugar de crearla.")
    homologacion.guardar_cuenta({
        "falcon": falcon, "netsuite": idn, "nombre_falcon": nombre, "nombre_es": nombre,
        "nombre_netsuite": (coa + " " + str(c.get("nombre_ns", "")).strip()).strip(),
        "destino": str(c.get("destino", "")).strip(), "contrapartida": str(c.get("contrapartida", "")).strip(),
    }, str(c.get("original", "")))
    return {"ok": True}


@app.delete("/homologacion/cuenta/{falcon}")
def eliminar_cuenta(falcon: str, request: Request):
    _solo_local(request)
    if not homologacion.eliminar_cuenta(falcon):
        raise HTTPException(404, "Esa cuenta no está en la tabla.")
    return {"ok": True}


@app.delete("/homologacion")
def borrar_homologacion(request: Request):
    _solo_local(request)
    homologacion.borrar()
    return {"ok": True}


@app.get("/homologacion/plantilla")
def plantilla_homologacion():
    return Response(homologacion.plantilla_xlsx(),
                    media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    headers={"Content-Disposition": 'attachment; filename="cuentas_homologadas_plantilla.xlsx"'})


@app.get("/periodos")
def ver_periodos():
    return {"periodos": periodos.listar(), "nombres": {k: h.nombre.strip() for k, h in PLANTILLA.hojas.items()}}


def _periodo_o_400(periodo: str) -> str:
    if not periodos.es_valido(periodo):
        raise HTTPException(400, "Período no válido (use AAAA-MM)")
    return periodo


@app.get("/periodo/{periodo}")
def ver_periodo(periodo: str):
    return periodos.leer(_periodo_o_400(periodo))


@app.put("/periodo/{periodo}")
def guardar_periodo(periodo: str, request: Request, datos: dict = Body(...)):
    """Guarda las grillas de un período: {hojas: {clave: {filas: [[...]], bloqueadas: [i, ...]}}}."""
    _solo_local(request)
    _periodo_o_400(periodo)
    hojas = datos.get("hojas")
    if not isinstance(hojas, dict):
        raise HTTPException(400, "Datos no válidos")
    limpias: dict = {}
    total = 0
    for clave, h in hojas.items():
        if clave not in PLANTILLA.hojas or not isinstance(h, dict):
            continue
        filas = [[str(c) for c in f] for f in (h.get("filas") or []) if isinstance(f, list)]
        total += len(filas)
        limpias[clave] = {"filas": filas, "bloqueadas": [int(i) for i in (h.get("bloqueadas") or []) if isinstance(i, int)]}
        if clave == "proveedores":       # diseño 2: Nombres y Apellidos junto a «Nombre de la empresa»; una página abierta con el diseño anterior se convierte al guardar
            d = h.get("disenio") or 1
            if d < 2:
                limpias[clave]["filas"] = periodos.proveedores_a_diseno2(filas)
            if d < 3:
                limpias[clave]["filas"] = periodos.proveedores_a_diseno3(limpias[clave]["filas"])
            if d < 4:
                limpias[clave]["filas"] = periodos.proveedores_a_diseno4(limpias[clave]["filas"])
            if d < 5:
                limpias[clave]["filas"] = periodos.proveedores_a_diseno5(limpias[clave]["filas"])
            limpias[clave]["disenio"] = 5
    if total > 50000:
        raise HTTPException(413, "Demasiadas filas")
    return {"ok": True, "actualizado": periodos.guardar(periodo, limpias)}


@app.delete("/periodo/{periodo}")
def borrar_periodo(periodo: str, request: Request):
    _solo_local(request)
    periodos.borrar(_periodo_o_400(periodo))
    return {"ok": True}


@app.post("/fila")
def validar_una_fila(datos: dict = Body(...)):
    """Valida y completa una sola fila (para el formulario de edición): {hoja, fila, formato_fecha}."""
    hoja = PLANTILLA.hojas.get(datos.get("hoja"))
    fila = datos.get("fila")
    fmt = datos.get("formato_fecha", "DD/MM/YYYY")
    if hoja is None or not isinstance(fila, list) or fmt not in FORMATOS_FECHA:
        raise HTTPException(400, "Datos no válidos")
    n = len(hoja.encabezados)
    celdas = [(c.strip() if isinstance(c, str) and c.strip() else None) for c in (list(fila) + [""] * n)[:n]]
    periodo = datos.get("periodo") if periodos.es_valido(datos.get("periodo") or "") else None
    res = validar_hoja(hoja, [(1, list(hoja.encabezados)), (2, celdas)], fmt, periodo)
    return {"fila": res.filas[0] if res.filas else [],
            "problemas": [{"nivel": p.nivel, "columna": p.columna, "mensaje": p.mensaje} for p in res.problemas]}


@app.post("/cargar")
async def cargar_a_grillas(archivo: UploadFile = File(...), hoja: str = Form("")):
    """Lee un Excel/CSV y devuelve las filas de cada pestaña reconocida para llenar las grillas."""
    contenido = await archivo.read(_MAX_CARGA + 1)
    if len(contenido) > _MAX_CARGA:
        return JSONResponse({"ok": False, "mensaje": "El archivo supera 40 MB."})
    try:
        datos, avisos = await run_in_threadpool(carga.leer, archivo.filename or "", contenido, PLANTILLA, hoja)
    except ValueError as e:
        return JSONResponse({"ok": False, "mensaje": str(e)})
    except Exception:
        return JSONResponse({"ok": False, "mensaje": "No se pudo leer el archivo; verifique que no esté dañado."})
    nombres = {k: h.nombre.strip() for k, h in PLANTILLA.hojas.items()}
    return {"ok": True, "hojas": datos, "avisos": avisos, "nombres": {k: nombres[k] for k in datos}}


@app.post("/procesar", response_class=HTMLResponse)
async def procesar_archivo(request: Request, archivo: UploadFile = File(...),
                           hoja: str = Form(""), formato_fecha: str = Form("DD/MM/YYYY")):
    if formato_fecha not in FORMATOS_FECHA:
        raise HTTPException(400, "Formato de fecha no válido")
    contenido = await archivo.read(_MAX_BYTES + 1)
    if len(contenido) > _MAX_BYTES:
        raise HTTPException(413, "El archivo supera 10 MB")
    nombre = (archivo.filename or "").lower()
    avisos: list[str] = []
    try:
        if nombre.endswith(".csv"):
            clave = norm(hoja)
            if clave not in PLANTILLA.hojas:
                raise ValueError("Para un CSV indique a qué plantilla corresponde.")
            datos, avisos = leer_csv(contenido, clave, PLANTILLA)
        elif nombre.endswith((".xlsx", ".xlsm")):
            datos, avisos = leer_xlsx(contenido, PLANTILLA)
        else:
            raise ValueError("Suba un archivo .xlsx o .csv")
        if not datos:
            raise ValueError("No se encontró ninguna hoja con las plantillas esperadas.")
    except ValueError as e:
        return templates.TemplateResponse(request, "_resultado.html", {"error": str(e)})
    except Exception:
        return templates.TemplateResponse(request, "_resultado.html",
                                          {"error": "No se pudo leer el archivo; verifique que no esté dañado."})
    resultados = procesar(datos, PLANTILLA, formato_fecha)
    return templates.TemplateResponse(request, "_resultado.html", {
        "token": _guardar(resultados), "resultados": resultados, "avisos": avisos, "error": None,
    })


@app.post("/procesar_grilla", response_class=HTMLResponse)
def procesar_grilla(request: Request, formato_fecha: str = "DD/MM/YYYY", periodo: str = "",
                    datos: dict[str, list[list[str]]] = Body(...)):
    """Recibe las filas escritas en las pestañas: {clave_hoja: [[celda, ...], ...]}."""
    if formato_fecha not in FORMATOS_FECHA:
        raise HTTPException(400, "Formato de fecha no válido")
    if sum(len(f) for f in datos.values()) > 20000:
        raise HTTPException(413, "Demasiadas filas")
    # el número de fila coincide con la posición en la pestaña (la primera fila de datos es la 2)
    armados = {}
    for clave, filas in datos.items():
        hoja = PLANTILLA.hojas.get(clave)
        if hoja is None:
            continue
        n = len(hoja.encabezados)
        armado = [(1, list(hoja.encabezados))]
        for pos, fila in enumerate(filas, start=2):
            celdas = [(c.strip() if isinstance(c, str) and c.strip() else None) for c in (fila + [""] * n)[:n]]
            if any(c is not None for c in celdas):
                armado.append((pos, celdas))
        armados[clave] = armado
    armados = {k: v for k, v in armados.items() if len(v) > 1}
    if not armados:
        return templates.TemplateResponse(request, "_resultado.html", {"error": "No hay filas con datos en las pestañas."})
    if periodo and not periodos.es_valido(periodo):
        raise HTTPException(400, "Período no válido (use AAAA-MM)")
    resultados = procesar(armados, PLANTILLA, formato_fecha, periodo or None)
    return templates.TemplateResponse(request, "_resultado.html", {
        "token": _guardar(resultados, periodo), "resultados": resultados, "avisos": [], "error": None,
        "periodo_etiqueta": periodos.etiqueta(periodo) if periodo else "",
    })


@app.get("/padron/estado")
def padron_estado():
    return padron.estado()


@app.post("/padron/actualizar")
async def padron_actualizar(request: Request, archivo: UploadFile | None = File(None)):
    """Descarga el padrón reducido de SUNAT (o usa el ZIP/TXT subido) y guarda los RUC que usa el sistema."""
    _solo_local(request)
    ruta = None
    if archivo is not None and archivo.filename:
        import tempfile
        fd, ruta = tempfile.mkstemp(suffix=".zip", prefix="padron_subido_")
        with os.fdopen(fd, "wb") as f:
            while True:
                trozo = await archivo.read(1024 * 1024)
                if not trozo:
                    break
                f.write(trozo)
    if not padron.iniciar_actualizacion(ruta):
        return JSONResponse({"ok": False, "mensaje": "Ya hay una actualización del padrón en curso."})
    return {"ok": True}


@app.get("/padron/ruc/{ruc}")
def padron_ruc(ruc: str):
    """Datos del RUC según el padrón de SUNAT, con los valores de las columnas de Proveedores que se pueden completar."""
    if not padron.es_ruc(ruc):
        return {"encontrado": False, "mensaje": "Escriba un RUC de 11 dígitos para consultar SUNAT."}
    d = padron.buscar(ruc)
    if d is None:
        n = padron.estado()["cargados"]
        return {"encontrado": False, "cargado": n,
                "mensaje": ("El padrón de SUNAT aún no está cargado en este equipo: pulse «🏛️ Actualizar padrón SUNAT» en Proveedores y vuelva a escribir el RUC."
                            if not n else "Ese RUC no está en el padrón cargado (queda anotado): pulse «🏛️ Actualizar padrón SUNAT» para traerlo y vuelva a escribir el RUC.")}
    campos = {proveedores.C_NOMBRE: d["razon"], proveedores.C_DIR1: d["direccion"], proveedores.C_UBIGEO: d["ubigeo"],
              proveedores.C_TIPO_DOC: "REGISTRO ÚNICO DE CONTRIBUYENTES", proveedores.C_NUMERO: d["ruc"]}
    avisos = []
    if d["estado"] and d["estado"].upper() != "ACTIVO":
        avisos.append(f"SUNAT: el RUC está {d['estado']}")
    if d["condicion"] and d["condicion"].upper() != "HABIDO":
        avisos.append(f"SUNAT: condición de domicilio {d['condicion']}")
    if d["ubigeo_codigo"] and not d["ubigeo"]:
        avisos.append(f"El ubigeo {d['ubigeo_codigo']} de SUNAT no está en la lista de la plantilla: elija el ubigeo")
    return {"encontrado": True, "datos": d, "campos": {k: v for k, v in campos.items() if v}, "avisos": avisos}


@app.post("/proveedores/csv")
def proveedores_csv(formato_fecha: str = "DD/MM/YYYY", periodo: str = "", filas: list[list[str]] = Body(...)):
    """Vista previa del CSV de proveedores tal como se enviará a NetSuite (ya completado y validado) + token para descargarlo."""
    if formato_fecha not in FORMATOS_FECHA:
        raise HTTPException(400, "Formato de fecha no válido")
    if len(filas) > 20000:
        raise HTTPException(413, "Demasiadas filas")
    hoja = PLANTILLA.hojas["proveedores"]
    n = len(hoja.encabezados)
    armado = [(1, list(hoja.encabezados))]
    for pos, fila in enumerate(filas, start=2):
        celdas = [(c.strip() if isinstance(c, str) and c.strip() else None) for c in (fila + [""] * n)[:n]]
        if any(c is not None for c in celdas):
            armado.append((pos, celdas))
    if len(armado) <= 1:
        res = ResultadoHoja(nombre=hoja.nombre, encabezados=hoja.encabezados)
    else:
        res = procesar({"proveedores": armado}, PLANTILLA, formato_fecha, periodo or None)[0]
    por_fila: dict[int, list[dict]] = {}
    for p in res.problemas:
        por_fila.setdefault(p.fila or 0, []).append({"nivel": p.nivel, "columna": p.columna, "mensaje": p.mensaje})
    return {"token": _guardar([res], periodo), "encabezados": res.encabezados, "filas": res.filas, "numeros": res.numeros_fila,
            "problemas": por_fila, "errores": res.errores, "avisos": res.avisos,
            "con_error": sorted({p.fila for p in res.problemas if p.nivel == "error" and p.fila})}


@app.get("/descargar/{token}/{archivo}")
def descargar(token: str, archivo: str):
    resultados = _obtener(token)
    if archivo == "todos.zip":
        if not any(r.exportable for r in resultados):
            raise HTTPException(409, "No hay hojas sin errores para exportar.")
        periodo = _RESULTADOS[token][2]
        nombre_zip = f"netsuite_csv_{periodo}.zip" if periodo else "netsuite_csv.zip"
        return Response(empaquetar_zip(resultados), media_type="application/zip",
                        headers={"Content-Disposition": f'attachment; filename="{nombre_zip}"'})
    if archivo == "errores.xlsx":
        periodo = _RESULTADOS[token][2]
        return Response(reporte_errores_xlsx(resultados, periodos.etiqueta(periodo) if periodo else ""),
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        headers={"Content-Disposition": f'attachment; filename="errores_{periodo or "validacion"}.xlsx"'})
    for sufijo in ("_sin_errores.csv", "_nuevos.csv", "_actualizar.csv"):
        if not archivo.endswith(sufijo):
            continue
        # solo las filas sin errores; «_nuevos» = sin ID de NetSuite (para crearlos), «_actualizar» = con ID (para actualizarlos en NetSuite)
        for r in resultados:
            if r.archivo == archivo.replace(sufijo, ".csv"):
                malas = {p.fila for p in r.problemas if p.nivel == "error"}
                filas = [f for n, f in zip(r.numeros_fila, r.filas) if n not in malas]
                if sufijo != "_sin_errores.csv":
                    i = r.encabezados.index("External ID")
                    filas = [f for f in filas if bool(str(f[i]).strip()) == (sufijo == "_actualizar.csv")]
                if not filas:
                    raise HTTPException(409, "Ninguna fila cumple (sin errores" + ("" if sufijo == "_sin_errores.csv" else ", " + ("con" if sufijo == "_actualizar.csv" else "sin") + " ID de NetSuite") + ").")
                sub = ResultadoHoja(nombre=r.nombre, encabezados=r.encabezados, filas=filas)
                return Response(sub.csv_bytes(), media_type="text/csv; charset=utf-8",
                                headers={"Content-Disposition": f'attachment; filename="{archivo}"'})
    for r in resultados:
        if r.archivo_xlsx == archivo:
            if not r.exportable:
                raise HTTPException(409, "La hoja tiene errores; corríjalos antes de exportar.")
            return Response(r.xlsx_bytes(), media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            headers={"Content-Disposition": f'attachment; filename="{r.archivo_xlsx}"'})
    for r in resultados:
        if r.archivo == archivo:
            if not r.exportable:
                raise HTTPException(409, "La hoja tiene errores; corríjalos antes de exportar.")
            return Response(r.csv_bytes(), media_type="text/csv; charset=utf-8",
                            headers={"Content-Disposition": f'attachment; filename="{r.archivo}"'})
    raise HTTPException(404, "Archivo no encontrado")
