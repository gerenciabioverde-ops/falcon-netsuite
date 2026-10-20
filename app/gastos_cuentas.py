"""Propuesta de cuenta contable (código de Falcon/PCGE) para un gasto según su Concepto.

Es solo una sugerencia para revisar: se propone únicamente si la cuenta existe en «Cuentas homologadas».
"""
from __future__ import annotations

import re

from . import homologacion
from .esquema import DEPARTAMENTOS

# (palabras que aparecen en el concepto, cuenta propuesta); gana la primera regla que coincida
REGLAS: list[tuple[tuple[str, ...], str]] = [
    (("seguro", "poliza", "robo"), "651101"),
    (("courier", "courrier", "muestras"), "631201"),
    (("viatico", "pasaje", "boleto", "movilidad"), "631001"),
    (("estiba", "logistic", "thc", "balanza", "marcado", "transporte", "puerto", "flete", "carga"), "631101"),
    (("envase", "empaque"), "656104"),
    (("limpieza", "vigilancia", "seguridad"), "656103"),
    (("util", "oficina"), "656101"),
    (("material", "suministro", "repuesto", "equipo", "combustible"), "656102"),
    (("asesor", "consultor", "legal"), "632001"),
    (("contab", "planilla", "emision de factura"), "632901"),
    (("basico",), "636001"),
    (("alquiler",), "635001"),
    (("senasa", "analisis", "proyecto", "trilla", "procesamiento", "coi/tc", "operativo"), "639901"),
    (("atencion", "representacion", "funeral"), "659101"),
]


def _plano(t: str) -> str:
    t = (t or "").lower()
    for a, b in zip("áéíóúñ", "aeioun"):
        t = t.replace(a, b)
    return re.sub(r"\s+", " ", t).strip()


def proponer(concepto: str) -> str:
    """Código de cuenta sugerido para el concepto, o '' si no hay sugerencia (o la cuenta no está homologada)."""
    c = _plano(concepto)
    existentes = homologacion.mapa()
    for palabras, cuenta in REGLAS:
        if any(p in c for p in palabras):
            return cuenta if cuenta in existentes else ""
    return ""


def metodo_pago(caja: str) -> str:
    """Método de pago según la columna Caja: con número de cuenta = transferencia; «Caja Chica» = efectivo; si no, ''."""
    c = _plano(caja)
    if not c:
        return ""
    if "caja chica" in c:
        return "Efectivo"
    return "Transferencia de Fondos" if any(ch.isdigit() for ch in c) else ""


# (palabras del concepto, departamento propuesto); gana la primera regla que coincida
REGLAS_DEPARTAMENTO: list[tuple[tuple[str, ...], str]] = [
    (("senasa", "analisis", "proyecto", "coi/tc", "certific", "area tecnica"), "Sostenibilidad"),
    (("contab", "planilla", "emision de factura", "seguro", "poliza", "robo", "banc", "financ"), "Finanzas"),
    (("estiba", "logistic", "thc", "balanza", "marcado", "transporte", "puerto", "flete", "carga", "envase", "empaque", "material",
      "suministro", "repuesto", "equipo", "combustible", "trilla", "procesamiento", "operativo", "courier", "courrier", "muestras",
      "viatico", "pasaje", "boleto"), "Operaciones"),
    (("movilidad", "atencion", "representacion", "funeral", "alquiler", "basico", "limpieza", "vigilancia", "util", "oficina", "legal",
      "asesor", "consultor"), "Administrativos"),
]


# (palabras del concepto, código del concepto de detracción); las reglas se contrastaron con los Monto Detr. que ya trae el reporte
REGLAS_DETRACCION: list[tuple[tuple[str, ...], str]] = [
    (("cafe",), "35"),       # compra de café (verde, pergamino, oro): bien exonerado del IGV, 1.5%
    (("pasaje", "boleto"), "26"),
    (("transporte", "puerto", "flete"), "27"),
    (("estiba", "thc", "movimiento de carga"), "21"),
    (("alquiler",), "19"),
    (("mantenimiento", "reparacion"), "20"),
    (("asesor", "legal", "contab", "planilla", "analisis", "coi/tc", "consultor", "logistic", "operativo", "marcado", "trilla",
      "procesamiento", "senasa", "courier", "courrier", "limpieza", "vigilancia", "emision de factura"), "22"),
]


def proponer_detraccion(concepto: str) -> str:
    """Nombre del concepto de detracción (tabla de la plantilla) para el concepto del gasto, o '' si no aplica."""
    from . import esquema
    c = _plano(concepto)
    for palabras, codigo in REGLAS_DETRACCION:
        if any(p in c for p in palabras):
            return next((n for n in esquema.DETRACCIONES if n.split(" ", 1)[0] == codigo), "")
    return ""


# (palabras del concepto, inicio de la opción de «Clasificación de bienes y servicios» de SUNAT); lo demás = «OTROS GASTOS…»
REGLAS_CLASIFICACION: list[tuple[tuple[str, ...], str]] = [
    (("equipo",), "ACTIVO FIJO"),
    (("viatico", "pasaje", "boleto", "movilidad", "atencion", "representacion", "funeral", "combustible"), "GASTOS DE EDUCAC"),
    (("envase", "empaque", "material", "suministro", "repuesto", "util", "compra de cafe"), "MERCADERIA"),
]


def proponer_clasificacion(concepto: str) -> str:
    """Opción de la tabla de clasificación de bienes y servicios de SUNAT (la de la plantilla) para el concepto del gasto."""
    from . import esquema
    opciones = esquema.LISTAS_EXTRA["Gastos"].get("Clasificación Bienes y Servicios", [])
    c = _plano(concepto)
    inicio = next((ini for palabras, ini in REGLAS_CLASIFICACION if any(p in c for p in palabras)), "OTROS GASTOS NO INCLUIDOS")
    return next((o for o in opciones if o.upper().startswith(inicio)), "")


def proponer_departamento(concepto: str) -> str:
    """Departamento sugerido para el concepto, o '' si no hay sugerencia."""
    c = _plano(concepto)
    for palabras, depto in REGLAS_DEPARTAMENTO:
        if any(p in c for p in palabras):
            return depto if depto in DEPARTAMENTOS else ""
    return ""
