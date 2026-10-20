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


def proponer_departamento(concepto: str) -> str:
    """Departamento sugerido para el concepto, o '' si no hay sugerencia."""
    c = _plano(concepto)
    for palabras, depto in REGLAS_DEPARTAMENTO:
        if any(p in c for p in palabras):
            return depto if depto in DEPARTAMENTOS else ""
    return ""
