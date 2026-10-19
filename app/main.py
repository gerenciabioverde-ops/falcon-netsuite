"""Aplicativo para generar los CSV de carga de NetSuite a partir de un Excel/CSV."""
from __future__ import annotations

import secrets
import time
from pathlib import Path

from fastapi import Body, FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, Response
from fastapi.templating import Jinja2Templates

from .esquema import RUTA_PLANTILLA, cargar_plantilla, norm
from .procesador import FORMATOS_FECHA, ResultadoHoja, empaquetar_zip, leer_csv, leer_xlsx, procesar

BASE = Path(__file__).parent
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
                "requerida": norm(e) in req,
                "opciones": h.listas[i].opciones if i in h.listas else [],
            } for i, e in enumerate(h.encabezados)],
        })
    return salida


@app.get("/", response_class=HTMLResponse)
def inicio(request: Request):
    pestanas = _pestanas()
    return templates.TemplateResponse(request, "index.html", {
        "hojas": [p["nombre"] for p in pestanas],
        "pestanas": pestanas,
        "formatos": list(FORMATOS_FECHA),
    })


@app.get("/plantilla")
def descargar_plantilla():
    return FileResponse(RUTA_PLANTILLA, filename="TEMPLATE_CSV_FALCON.xlsx")


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
