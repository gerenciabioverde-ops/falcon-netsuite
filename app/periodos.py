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
    """Lleva una fila de Gastos al diseño actual (27 columnas). Se reconoce el diseño por el contenido: en el diseño
    de 19 columnas la columna 10 es la Moneda (PEN/USD); en el actual es la Cuenta NetSuite (un número). Así también se
    arreglan filas de 19 columnas que se guardaron rellenadas con celdas vacías hasta 21."""
    f = list(f)
    if len(f) == 18:
        f = ["", *f]
    if len(f) >= 19 and str(f[10]).strip().upper() in ("PEN", "USD", "SOL"):
        f = [*f[:9], "", "", *f[9:19]]
    return (f + [""] * 27)[:27]


# Nombres anteriores de los departamentos de Gastos -> nombres actuales
_DEPARTAMENTOS_ANTES = {"gerencia general": "Administrativos", "gerencia finanzas": "Finanzas",
                        "gerencia sostenibilidad": "Sostenibilidad", "gerencia operaciones": "Operaciones"}


def _migrar(datos: dict) -> dict:
    """Gastos guardados con un diseño anterior: 18 columnas (sin External ID) o 19 (sin Cuenta PCGE / Cuenta NetSuite)."""
    g = (datos.get("hojas") or {}).get("gastos")
    if g and g.get("filas"):
        g["filas"] = [_fila_gastos(f) if len(f) in (18, 19, 21, 22, 24) else f for f in g["filas"]]
        for f in g["filas"]:
            if len(f) == 27 and str(f[22]).strip().lower() in _DEPARTAMENTOS_ANTES:
                f[22] = _DEPARTAMENTOS_ANTES[str(f[22]).strip().lower()]
    return datos


def leer(periodo: str) -> dict:
    try:
        return _migrar(json.loads(_ruta(periodo).read_text(encoding="utf-8")))
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
