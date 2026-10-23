"""Datos por período contable (mes): cada período guarda sus propias grillas en datos_usuario/periodos/AAAA-MM.json."""
from __future__ import annotations

import functools
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
    e = pv.ENCABEZADOS[:1] + pv.ENCABEZADOS[3:15] + pv.ENCABEZADOS[1:3] + pv.ENCABEZADOS[15:]       # orden de columnas del diseño 3 (antes de mover el tipo y número de documento)
    iD, iU = e.index(pv.C_DIR1), e.index(pv.C_UBIGEO)
    otros = [e.index(c) for c in (pv.C_CIUDAD, pv.C_ESTADO, pv.C_ZIP)]
    for f in filas:
        if len(f) == len(e) and str(f[iD]).strip() == pv.DIRECCION_DEFECTO and str(f[iU]).strip() == pv.UBIGEO_DEFECTO:
            for i in (iD, iU, *otros):
                f[i] = ""
    return filas


def proveedores_a_diseno4(filas: list) -> list:
    """Proveedores: «Latam - Tipo de doc. de identidad» y «Latam - Tax Reg Number» pasaron a ir justo después de External ID (diseño 4)."""
    return [f[:1] + f[13:15] + f[1:13] + f[15:] if len(f) == 23 else f for f in filas]


def proveedores_a_diseno5(filas: list) -> list:
    """Diseño 5: Nombre de la empresa, Nombres y Apellidos escritos todo en MAYÚSCULAS pasan a inicial mayúscula (Torres Nuñez Sadith);
    las siglas de una sola letra (E.I.R.L., S.A.C.) se conservan y lo que ya tiene minúsculas no se toca."""
    from . import proveedores as pv
    e = pv.ENCABEZADOS
    idx = [e.index(c) for c in (pv.C_NOMBRE, pv.C_NOMBRES, pv.C_APELLIDOS)]
    for f in filas:
        if len(f) == len(e):
            for i in idx:
                if str(f[i]).isupper():
                    f[i] = re.sub(r"[A-ZÁÉÍÓÚÑÜ]+", lambda m: m.group(0).capitalize(), f[i])
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
    if pr and pr.get("filas") and (pr.get("disenio") or 1) < 4:
        pr["filas"], pr["disenio"] = proveedores_a_diseno4(pr["filas"]), 4
    if pr and pr.get("filas") and (pr.get("disenio") or 1) < 5:
        pr["filas"], pr["disenio"] = proveedores_a_diseno5(pr["filas"]), 5
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


@functools.lru_cache(maxsize=None)
def _col_hoja(hoja: str, nombre: str) -> int:
    from . import esquema
    return [e.strip() for e in esquema.cargar_plantilla().hojas[hoja].encabezados].index(nombre)


def _subsidiaria_factura(datos: dict) -> dict:
    """Factura de Proveedor: la subsidiaria de versiones anteriores («43» o «Peru») pasa a la actual («FCP»)."""
    from . import esquema
    fp = (datos.get("hojas") or {}).get("factura de proveedor")
    if fp and fp.get("filas"):
        i = _col_hoja("factura de proveedor", "Subsidiaria")
        for f in fp["filas"]:
            if len(f) > i and str(f[i]).strip().upper() in esquema.SUBSIDIARIAS_ANTERIORES_FACTURA:
                f[i] = esquema.SUBSIDIARIA_FACTURA
    return datos


def _cuenta_pagar_factura(datos: dict) -> dict:
    """Factura de Proveedor (cuenta de la cabecera) y Pago Proveedor (A/P ACCOUNT, la misma cuenta): NetSuite solo admite cuentas de tipo
    «Cuentas por pagar» ahí. Las que la app ponía para honorarios (2400.08) y otros pasivos (2420) pasan a la cuenta por pagar de las
    facturas (42100); una cuenta puesta a mano no se toca."""
    hojas = datos.get("hojas") or {}
    if not any((hojas.get(k) or {}).get("filas") for k in ("factura de proveedor", "pago proveedor")):
        return datos
    from . import esquema, homologacion
    coa, mapa = homologacion.coa(), homologacion.mapa()
    n = homologacion.normalizar
    destino = coa.get(n("421201"))
    if not destino:
        return datos
    antiguas = {esquema.CUENTA_POR_PAGAR_RHE} | {v for k in ("424001", "469901") for v in (coa.get(n(k)), mapa.get(n(k))) if v}
    antiguas.discard(destino)
    for clave, col in (("factura de proveedor", "Cuenta"), ("pago proveedor", "A/P ACCOUNT")):
        h = hojas.get(clave)
        if not (h and h.get("filas")):
            continue
        i = _col_hoja(clave, col)
        for f in h["filas"]:
            if len(f) > i and str(f[i]).strip() in antiguas:
                f[i] = destino
    return datos


def valores_ns_departamento_clase(hojas: dict) -> None:
    """Departamento y clase de las líneas en NetSuite: los de Falcon (Operaciones, Finanzas…, o vacío) pasan a «Accounting & Finance» / «BIB»;
    un valor puesto a mano que no sea de Falcon no se toca. Factura de Proveedor (Gastos: Departamento / Gastos: Clase, con importe) y
    Pago Proveedor (DEPARTMENT / CLASS, con pago). Modifica las filas en el lugar."""
    from . import esquema
    propios = {n.lower() for n in esquema.DEPARTAMENTOS} | set(_DEPARTAMENTOS_ANTES) | {n.lower() for n in _DEPARTAMENTOS_ANTES.values()}
    for clave, c_dep, c_cla, c_imp in (("factura de proveedor", "Gastos: Departamento", "Gastos: Clase", "Gastos: Importe"),
                                       ("pago proveedor", "DEPARTMENT", "CLASS", "PAYMENT")):
        h = hojas.get(clave)
        if not (h and h.get("filas")):
            continue
        iD, iC, iI = _col_hoja(clave, c_dep), _col_hoja(clave, c_cla), _col_hoja(clave, c_imp)
        vacio = lambda v: v is None or not str(v).strip()  # noqa: E731
        for f in h["filas"]:
            if len(f) <= max(iD, iC, iI) or vacio(f[iI]):
                continue
            if vacio(f[iD]) or str(f[iD]).strip().lower() in propios or str(f[iD]).strip() == esquema.DEPARTAMENTO_NS.replace("&", "&amp;"):       # «&amp;»: error de una versión anterior
                f[iD] = esquema.DEPARTAMENTO_NS
            if vacio(f[iC]):
                f[iC] = esquema.CLASE_NS


def _departamento_clase_ns(datos: dict) -> dict:
    valores_ns_departamento_clase(datos.get("hojas") or {})
    return datos


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
        return _departamento_clase_ns(_cuenta_pagar_factura(_subsidiaria_factura(_migrar_ids(_migrar(json.loads(_ruta(periodo).read_text(encoding="utf-8")))))))
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
