"""Cuentas homologadas: equivalencia entre la cuenta contable de Falcon y la cuenta de NetSuite.

La tabla se sube desde la pantalla, se guarda en `datos_usuario/` (fuera de git, para que
«Actualizar» no la pise) y se aplica al generar los CSV en las columnas de cuenta.
"""
from __future__ import annotations

import csv
import io
import json
import re
import time
import warnings
from pathlib import Path

from openpyxl import Workbook, load_workbook

RUTA = Path(__file__).resolve().parent.parent / "datos_usuario" / "cuentas_homologadas.json"
MAX_FILAS = 20000

# Encabezados reconocidos (se comparan sin tildes ni mayúsculas)
_ALIAS_FALCON = ("falcon", "codigo", "pcge", "cuenta contable", "cuenta origen")
_ALIAS_NS = ("netsuite", "id interno", "internal", "cuenta destino")
_ALIAS_NOMBRE = ("nombre", "descripcion", "denominacion")


def _sin_tildes(t: str) -> str:
    return (t.lower().replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u"))


def normalizar(valor) -> str:
    """Forma comparable de una cuenta: sin espacios y sin «.0» de Excel."""
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    return re.sub(r"\s+", "", str(valor)).strip()


def _texto(valor) -> str:
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    return str(valor).strip()


def _filas_de_archivo(nombre: str, contenido: bytes) -> list[list]:
    nombre = nombre.lower()
    if nombre.endswith(".csv"):
        texto = contenido.decode("utf-8-sig", errors="replace")
        primera = texto.split("\n", 1)[0]
        delim = ";" if primera.count(";") > primera.count(",") else ","
        return [f for f in csv.reader(io.StringIO(texto), delimiter=delim)]
    if nombre.endswith((".xlsx", ".xlsm")):
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            wb = load_workbook(io.BytesIO(contenido), data_only=True, read_only=True)
        ws = wb[wb.sheetnames[0]]
        return [list(f) for f in ws.iter_rows(values_only=True)]
    raise ValueError("Suba un archivo .xlsx o .csv")


def _columnas(encabezado: list) -> tuple[int, int, int | None, int | None]:
    """Índices de: cuenta Falcon, cuenta NetSuite, nombre Falcon, nombre NetSuite."""
    h = [_sin_tildes(_texto(c)) for c in encabezado]
    i_ns = next((i for i, t in enumerate(h) if any(a in t for a in _ALIAS_NS) and not t.startswith("nombre")), None)
    i_fa = next((i for i, t in enumerate(h) if any(a in t for a in _ALIAS_FALCON)
                 and "netsuite" not in t and not t.startswith("nombre")), None)
    if i_fa is None or i_ns is None:
        raise ValueError("No encontré las columnas. La primera fila debe tener «Cuenta Falcon» y «Cuenta NetSuite» "
                         "(opcionalmente «Nombre Falcon» y «Nombre NetSuite»). Descargue la plantilla.")
    n_fa = next((i for i, t in enumerate(h) if any(a in t for a in _ALIAS_NOMBRE) and "falcon" in t), None)
    n_ns = next((i for i, t in enumerate(h) if any(a in t for a in _ALIAS_NOMBRE) and "netsuite" in t), None)
    return i_fa, i_ns, n_fa, n_ns


def procesar_archivo(nombre: str, contenido: bytes) -> tuple[list[dict], list[str]]:
    """Devuelve (filas válidas, errores). Lanza ValueError si el archivo no sirve."""
    filas = _filas_de_archivo(nombre, contenido)
    if not filas:
        raise ValueError("El archivo está vacío.")
    i_fa, i_ns, n_fa, n_ns = _columnas(filas[0])
    cuentas: dict[str, dict] = {}
    errores: list[str] = []
    for nro, f in enumerate(filas[1:], start=2):
        f = list(f) + [None] * 8
        falcon, ns = _texto(f[i_fa]), _texto(f[i_ns])
        if not falcon and not ns:
            continue
        if not falcon or not ns:
            errores.append(f"Fila {nro}: falta {'la cuenta Falcon' if not falcon else 'la cuenta NetSuite'}.")
            continue
        clave = normalizar(falcon)
        previa = cuentas.get(clave)
        if previa and previa["netsuite"] != ns:
            errores.append(f"Fila {nro}: la cuenta Falcon {falcon} ya estaba asignada a {previa['netsuite']} y ahora a {ns}.")
            continue
        cuentas[clave] = {
            "falcon": falcon, "netsuite": ns,
            "nombre_falcon": _texto(f[n_fa]) if n_fa is not None else "",
            "nombre_netsuite": _texto(f[n_ns]) if n_ns is not None else "",
        }
        if len(cuentas) > MAX_FILAS:
            raise ValueError(f"Demasiadas cuentas (máximo {MAX_FILAS}).")
    if not cuentas:
        raise ValueError("No se encontró ninguna cuenta con datos.")
    return list(cuentas.values()), errores


def guardar(filas: list[dict], archivo: str) -> None:
    RUTA.parent.mkdir(parents=True, exist_ok=True)
    RUTA.write_text(json.dumps({"archivo": archivo, "actualizado": time.strftime("%d/%m/%Y %H:%M"), "filas": filas},
                               ensure_ascii=False), encoding="utf-8")
    _cache.clear()


def borrar() -> None:
    RUTA.unlink(missing_ok=True)
    _cache.clear()


_cache: dict = {}


def cargar() -> dict:
    """{'archivo','actualizado','filas':[...]} (vacío si no hay tabla)."""
    try:
        mtime = RUTA.stat().st_mtime
    except FileNotFoundError:
        return {"archivo": "", "actualizado": "", "filas": []}
    if _cache.get("mtime") != mtime:
        try:
            _cache["datos"] = json.loads(RUTA.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            _cache["datos"] = {"archivo": "", "actualizado": "", "filas": []}
        _cache["mtime"] = mtime
        _cache["mapa"] = {normalizar(f["falcon"]): f["netsuite"] for f in _cache["datos"]["filas"]}
    return _cache["datos"]


def mapa() -> dict[str, str]:
    cargar()
    return _cache.get("mapa", {})


def plantilla_xlsx() -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Cuentas homologadas"
    ws.append(["Cuenta Falcon", "Nombre Falcon", "Cuenta NetSuite", "Nombre NetSuite"])
    for col, ancho in zip("ABCD", (18, 38, 18, 38)):
        ws.column_dimensions[col].width = ancho
    ws2 = wb.create_sheet("Instrucciones")
    for linea in (
        "Una fila por cuenta. Borre esta hoja o déjela: solo se lee la primera hoja.",
        "Cuenta Falcon: código contable que usa en Falcon Contable (por ejemplo 1041).",
        "Cuenta NetSuite: lo que debe escribirse en el CSV: el ID interno de la cuenta en NetSuite o su nombre completo.",
        "Los nombres son opcionales y solo sirven para que usted revise la tabla.",
        "Una misma cuenta Falcon no puede tener dos cuentas NetSuite distintas.",
    ):
        ws2.append([linea])
    ws2.column_dimensions["A"].width = 110
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
