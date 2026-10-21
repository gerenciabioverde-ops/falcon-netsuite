"""Datos por período contable (mes): cada período guarda sus propias grillas en datos_usuario/periodos/AAAA-MM.json."""
from __future__ import annotations

import json
import os
import re
import time
from pathlib import Path

DIRECTORIO = Path(__file__).resolve().parent.parent / "datos_usuario" / "periodos"
_VALIDO = re.compile(r"^\d{4}-(0[1-9]|1[0-2])$")
MESES_NETSUITE = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def es_valido(periodo: str) -> bool:
    return bool(periodo and _VALIDO.match(periodo))


def etiqueta(periodo: str) -> str:
    """'2026-10' -> 'Oct 2026' (formato del campo POSTING PERIOD de NetSuite)."""
    anio, mes = periodo.split("-")
    return f"{MESES_NETSUITE[int(mes) - 1]} {anio}"


def _ruta(periodo: str) -> Path:
    if not es_valido(periodo):
        raise ValueError("Período no válido")
    return DIRECTORIO / f"{periodo}.json"


def _fila_gastos(f: list) -> list:
    """Lleva una fila de Gastos al diseño actual (34 columnas). Se reconoce el diseño por el contenido: en el diseño
    de 19 columnas la columna 10 es la Moneda (PEN/USD); en el actual es la Cuenta NetSuite (un número). Así también se
    arreglan filas de 19 columnas que se guardaron rellenadas con celdas vacías hasta 21."""
    f = list(f)
    if len(f) == 18:
        f = ["", *f]
    if len(f) >= 19 and str(f[10]).strip().upper() in ("PEN", "USD", "SOL"):
        f = [*f[:9], "", "", *f[9:19]]
    if len(f) >= 34:
        return f[:34]
    f = (f + [""] * 32)[:32]
    return [*f[:17], "", "", *f[17:]]          # N° Constancia y N° Operación Detr. van después de Fecha Detr.


# Nombres anteriores de los departamentos de Gastos -> nombres actuales
_DEPARTAMENTOS_ANTES = {"gerencia general": "Administrativos", "gerencia finanzas": "Finanzas",
                        "gerencia sostenibilidad": "Sostenibilidad", "gerencia operaciones": "Operaciones"}


def _migrar(datos: dict) -> dict:
    """Gastos guardados con un diseño anterior: 18 columnas (sin External ID) o 19 (sin Cuenta PCGE / Cuenta NetSuite)."""
    g = (datos.get("hojas") or {}).get("gastos")
    if g and g.get("filas"):
        g["filas"] = [_fila_gastos(f) if 18 <= len(f) <= 32 else f for f in g["filas"]]
        for f in g["filas"]:
            if len(f) == 34 and str(f[24]).strip().lower() in _DEPARTAMENTOS_ANTES:
                f[24] = _DEPARTAMENTOS_ANTES[str(f[24]).strip().lower()]
    return datos


_ID_ANTIGUO = re.compile(r"^VENDB_([A-Za-z0-9-]+)_([A-Za-z0-9-]+)$")


def _id_nuevo(valor) -> str:
    """VENDB_RUC_Serie-Número (formato anterior) -> VENDB_RUC+SerieNúmero sin separadores (VENDB_20513497769FA0100018047)."""
    m = _ID_ANTIGUO.match(str(valor or "").strip())
    return "VENDB_" + m.group(1).replace("-", "") + m.group(2).replace("-", "") if m else valor


def _migrar_ids(datos: dict) -> dict:
    """IDs de Gastos / Factura de Proveedor: formato sin separadores y de largo fijo (VENDB_RUC + serie de 4 + número de 8)."""
    from . import esquema, ids
    hojas = datos.get("hojas") or {}
    for clave in ("gastos", "factura de proveedor"):
        for f in (hojas.get(clave) or {}).get("filas") or []:
            if f:
                f[0] = _id_nuevo(f[0])
    nuevos: dict[str, str] = {}
    iR, iS = esquema.GASTOS_ENCABEZADOS.index("RUC"), esquema.GASTOS_ENCABEZADOS.index("Serie-Numero")
    for f in (hojas.get("gastos") or {}).get("filas") or []:
        if f and str(f[0]).startswith("VENDB_") and len(f) > max(iR, iS):
            nuevo = ids.id_comprobante("VENDB", f[iR], f[iS])
            if nuevo and nuevo != f[0]:
                nuevos[f[0]], f[0] = nuevo, nuevo
    for f in (hojas.get("factura de proveedor") or {}).get("filas") or []:
        if f and f[0] in nuevos:
            f[0] = nuevos[f[0]]
    return datos


def leer(periodo: str) -> dict:
    try:
        return _migrar_ids(_migrar(json.loads(_ruta(periodo).read_text(encoding="utf-8"))))
    except (FileNotFoundError, ValueError):
        return {"actualizado": "", "hojas": {}}


def _contar(hojas: dict) -> dict[str, int]:
    return {k: sum(1 for f in (h.get("filas") or []) if any(str(c).strip() for c in f)) for k, h in hojas.items()}


def guardar(periodo: str, hojas: dict) -> str:
    """Guarda las grillas del período; si no queda ninguna fila con datos, elimina el archivo."""
    ruta = _ruta(periodo)
    if not any(_contar(hojas).values()):
        ruta.unlink(missing_ok=True)
        return ""
    DIRECTORIO.mkdir(parents=True, exist_ok=True)
    marca = time.strftime("%d/%m/%Y %H:%M")
    tmp = ruta.with_suffix(".tmp")
    tmp.write_text(json.dumps({"actualizado": marca, "hojas": hojas}, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    os.replace(tmp, ruta)
    return marca


def borrar(periodo: str) -> None:
    _ruta(periodo).unlink(missing_ok=True)


def listar() -> list[dict]:
    salida = []
    if DIRECTORIO.exists():
        for f in sorted(DIRECTORIO.glob("*.json"), reverse=True):
            if not es_valido(f.stem):
                continue
            try:
                d = json.loads(f.read_text(encoding="utf-8"))
            except ValueError:
                continue
            conteo = {k: n for k, n in _contar(d.get("hojas", {})).items() if n}
            if conteo:
                salida.append({"periodo": f.stem, "actualizado": d.get("actualizado", ""), "filas": conteo})
    return salida
