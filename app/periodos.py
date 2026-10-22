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
    """Lleva una fila de Gastos al diseño actual (38 columnas: «N° Operación» va después de Cod.Pago; las 3 últimas —% Retención, Monto Retención y Código de impuesto— se agregaron al final). Se reconoce el diseño por el contenido: en el diseño
    de 19 columnas la columna 10 es la Moneda (PEN/USD); en el actual es la Cuenta NetSuite (un número). Así también se
    arreglan filas de 19 columnas que se guardaron rellenadas con celdas vacías hasta 21."""
    f = list(f)
    if len(f) == 18:
        f = ["", *f]
    if len(f) >= 19 and str(f[10]).strip().upper() in ("PEN", "USD", "SOL"):
        f = [*f[:9], "", "", *f[9:19]]
    if len(f) >= 38:
        return f[:38]
    if len(f) >= 34:
        f = (f + [""] * 3)[:37]
        return [*f[:4], "", *f[4:]]
    f = (f + [""] * 32)[:32]
    return [*f[:4], "", *f[4:17], "", "", *f[17:], "", "", ""]          # N° Constancia y N° Operación Detr. van después de Fecha Detr.


# Nombres anteriores de los departamentos de Gastos -> nombres actuales
_DEPARTAMENTOS_ANTES = {"gerencia general": "Administrativos", "gerencia finanzas": "Finanzas",
                        "gerencia sostenibilidad": "Sostenibilidad", "gerencia operaciones": "Operaciones"}


def proveedores_a_diseno2(filas: list) -> list:
    """Proveedores: Nombres y Apellidos pasaron de ser las 2 últimas columnas a ir justo después de «Nombre de la empresa» (diseño 2)."""
    return [f[:2] + f[-2:] + f[2:-2] if len(f) == 23 else f for f in filas]


def proveedores_a_diseno3(filas: list) -> list:
    """Diseño 3: en Proveedores la dirección por defecto de Jaén ya no se guarda (el CSV la sigue poniendo al generarlo): se vacían Dirección 1,
    Ubigeo, Ciudad, Estado y ZIP de las filas que tenían exactamente ese valor por defecto."""
    from . import proveedores as pv
    e = pv.ENCABEZADOS
    iD, iU = e.index(pv.C_DIR1), e.index(pv.C_UBIGEO)
    otros = [e.index(c) for c in (pv.C_CIUDAD, pv.C_ESTADO, pv.C_ZIP)]
    for f in filas:
        if len(f) == len(e) and str(f[iD]).strip() == pv.DIRECCION_DEFECTO and str(f[iU]).strip() == pv.UBIGEO_DEFECTO:
            for i in (iD, iU, *otros):
                f[i] = ""
    return filas


def _migrar(datos: dict) -> dict:
    """Gastos guardados con un diseño anterior: 18 columnas (sin External ID) o 19 (sin Cuenta PCGE / Cuenta NetSuite)."""
    pr = (datos.get("hojas") or {}).get("proveedores")
    if pr and pr.get("filas"):
        for f in pr["filas"]:       # el External ID de un proveedor es el ID de NetSuite: el VEND_<número> que se generaba antes se quita
            if f and str(f[0]).upper().startswith("VEND_"):
                f[0] = ""       # Proveedores tenía 26 columnas: se quitaron Etiqueta (2), Dirección 2 (4) y Default Tax Reg. (10)
        pr["filas"] = [[c for i, c in enumerate(f) if i not in (2, 4, 10)] if len(f) == 26 else f for f in pr["filas"]]
    if pr and pr.get("filas") and (pr.get("disenio") or 1) < 2:
        pr["filas"], pr["disenio"] = proveedores_a_diseno2(pr["filas"]), 2
    if pr and pr.get("filas") and (pr.get("disenio") or 1) < 3:
        pr["filas"], pr["disenio"] = proveedores_a_diseno3(pr["filas"]), 3
    g = (datos.get("hojas") or {}).get("gastos")
    if g and g.get("filas"):
        g["filas"] = [_fila_gastos(f) if (18 <= len(f) <= 32 or 34 <= len(f) < 38) else f for f in g["filas"]]
        for f in g["filas"]:
            if len(f) >= 38 and str(f[25]).strip().lower() in _DEPARTAMENTOS_ANTES:
                f[25] = _DEPARTAMENTOS_ANTES[str(f[25]).strip().lower()]
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
