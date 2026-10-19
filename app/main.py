"""Aplicativo para generar los CSV de carga de NetSuite a partir de un Excel/CSV."""
from __future__ import annotations

import os
import secrets
import subprocess
import threading
import time
from pathlib import Path

from fastapi import Body, FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, Response
from fastapi.templating import Jinja2Templates

from . import homologacion
from .esquema import RUTA_PLANTILLA, cargar_plantilla, norm
from .procesador import FORMATOS_FECHA, ResultadoHoja, empaquetar_zip, leer_csv, leer_xlsx, procesar

BASE = Path(__file__).parent
RAIZ = BASE.parent
app = FastAPI(title="Falcon NetSuite CSV")
templates = Jinja2Templates(directory=str(BASE / "templates"))
PLANTILLA = cargar_plantilla()

# Resultados en memoria (un servidor, uso interno): token -> (timestamp, resultados)
_RESULTADOS: dict[str, tuple[float, list[ResultadoHoja]]] = {}
_TTL = 3600
_MAX_BYTES = 10 * 1024 * 1024


def _guardar(resultados: list[ResultadoHoja]) -> str:
    ahora = time.time()
    for k in [k for k, (t, _) in _RESULTADOS.items() if ahora - t > _TTL]:
        del _RESULTADOS[k]
    token = secrets.token_urlsafe(12)
    _RESULTADOS[token] = (ahora, resultados)
    return token


def _obtener(token: str) -> list[ResultadoHoja]:
    if token not in _RESULTADOS:
        raise HTTPException(404, "El resultado expiró; vuelva a procesar el archivo.")
    return _RESULTADOS[token][1]


def _pestanas() -> list[dict]:
    """Descripción de cada plantilla para dibujar su pestaña con grilla."""
    salida = []
    for clave, h in PLANTILLA.hojas.items():
        req = {norm(c) for c in h.reglas.get("requeridos", [])}
        salida.append({
            "clave": clave,
            "nombre": h.nombre.strip(),
            "columnas": [{
                "nombre": e.strip(),
                "auto": i == 0 and bool(h.reglas.get("id_auto")),
                "requerida": norm(e) in req and not (i == 0 and h.reglas.get("id_auto")),
                "opciones": h.listas[i].opciones if i in h.listas else [],
            } for i, e in enumerate(h.encabezados)],
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


@app.post("/actualizar")
def actualizar(request: Request):
    """Trae la última versión de `main` desde GitHub. Solo desde el propio equipo."""
    if not request.client or request.client.host not in ("127.0.0.1", "::1"):
        raise HTTPException(403, "Solo se puede actualizar desde el equipo donde corre la app.")
    if not (RAIZ / ".git").exists():
        return JSONResponse({"ok": False, "mensaje": "Esta copia no es un repositorio git; descargue de nuevo desde GitHub."})
    try:
        antes = _git("rev-parse", "HEAD").stdout.strip()
        rama = _git("rev-parse", "--abbrev-ref", "HEAD").stdout.strip()
        if rama != "main":
            r = _git("checkout", "main")
            if r.returncode != 0:
                return JSONResponse({"ok": False, "mensaje": "No se pudo cambiar a main: " + (r.stderr or r.stdout).strip()[:300]})
        r = _git("pull", "--ff-only", "origin", "main")
        if r.returncode != 0:
            return JSONResponse({"ok": False, "mensaje": "No se pudo actualizar: " + (r.stderr or r.stdout).strip()[:300]})
        despues = _git("rev-parse", "HEAD").stdout.strip()
    except (subprocess.TimeoutExpired, FileNotFoundError) as e:
        return JSONResponse({"ok": False, "mensaje": "No se pudo ejecutar git (¿está instalado y hay internet?): " + str(e)[:120]})
    if antes == despues and rama == "main":
        return JSONResponse({"ok": True, "cambios": False, "mensaje": "Ya tiene la última versión."})
    # iniciar.bat vuelve a lanzar el servidor cuando termina con código 75
    reinicia = os.environ.get("FALCON_REINICIAR") == "1"
    if reinicia:
        threading.Timer(1.5, lambda: os._exit(75)).start()
    return JSONResponse({"ok": True, "cambios": True, "reiniciando": reinicia, "mensaje":
                         "Actualizado. Reiniciando…" if reinicia else "Actualizado. Cierre y vuelva a abrir iniciar.bat para aplicar los cambios."})


@app.get("/", response_class=HTMLResponse)
def inicio(request: Request):
    pestanas = _pestanas()
    return templates.TemplateResponse(request, "index.html", {
        "version": _version(),
        "hojas": [p["nombre"] for p in pestanas],
        "pestanas": pestanas,
        "formatos": list(FORMATOS_FECHA),
    })


@app.get("/plantilla")
def descargar_plantilla():
    return FileResponse(RUTA_PLANTILLA, filename="TEMPLATE_CSV_FALCON.xlsx")


def _solo_local(request: Request) -> None:
    if not request.client or request.client.host not in ("127.0.0.1", "::1"):
        raise HTTPException(403, "Solo desde el equipo donde corre la app.")


@app.get("/homologacion")
def ver_homologacion():
    d = homologacion.cargar()
    return {"archivo": d["archivo"], "actualizado": d["actualizado"], "cantidad": len(d["filas"]), "filas": d["filas"]}


@app.post("/homologacion")
async def subir_homologacion(request: Request, archivo: UploadFile = File(...)):
    _solo_local(request)
    contenido = await archivo.read(_MAX_BYTES + 1)
    if len(contenido) > _MAX_BYTES:
        raise HTTPException(413, "El archivo supera 10 MB")
    try:
        filas, errores = homologacion.procesar_archivo(archivo.filename or "", contenido)
    except ValueError as e:
        return JSONResponse({"ok": False, "mensaje": str(e), "errores": []})
    except Exception:
        return JSONResponse({"ok": False, "mensaje": "No se pudo leer el archivo; verifique que no esté dañado.", "errores": []})
    if errores:  # no se guarda una tabla ambigua o incompleta
        return JSONResponse({"ok": False, "mensaje": f"Corrija {len(errores)} problema(s) y vuelva a subir.", "errores": errores[:50]})
    homologacion.guardar(filas, archivo.filename or "")
    return {"ok": True, "mensaje": f"{len(filas)} cuentas cargadas.", "errores": []}


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
def procesar_grilla(request: Request, formato_fecha: str = "DD/MM/YYYY",
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
    resultados = procesar(armados, PLANTILLA, formato_fecha)
    return templates.TemplateResponse(request, "_resultado.html", {
        "token": _guardar(resultados), "resultados": resultados, "avisos": [], "error": None,
    })


@app.get("/descargar/{token}/{archivo}")
def descargar(token: str, archivo: str):
    resultados = _obtener(token)
    if archivo == "todos.zip":
        if not any(r.exportable for r in resultados):
            raise HTTPException(409, "No hay hojas sin errores para exportar.")
        return Response(empaquetar_zip(resultados), media_type="application/zip",
                        headers={"Content-Disposition": 'attachment; filename="netsuite_csv.zip"'})
    for r in resultados:
        if r.archivo == archivo:
            if not r.exportable:
                raise HTTPException(409, "La hoja tiene errores; corríjalos antes de exportar.")
            return Response(r.csv_bytes(), media_type="text/csv; charset=utf-8",
                            headers={"Content-Disposition": f'attachment; filename="{r.archivo}"'})
    raise HTTPException(404, "Archivo no encontrado")
