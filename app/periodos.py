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


def _migrar(datos: dict) -> dict:
    """Gastos guardados antes de existir la columna External ID (18 columnas): se les antepone la celda vacía."""
    g = (datos.get("hojas") or {}).get("gastos")
    if g and g.get("filas") and all(len(f) == 18 for f in g["filas"]):
        g["filas"] = [["", *f] for f in g["filas"]]
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
