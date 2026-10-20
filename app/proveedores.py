"""Plantilla de Proveedores (hoja «Vendor» del libro Entity - Falcon Coffees - Perú - Legacy).

Con solo Nombre, Dirección 1, Ubigeo, Tipo de documento y Número, el resto se completa igual que en el libro:
Ciudad/Estado desde el ubigeo, código ZIP desde la tabla de códigos postales (primera coincidencia de la
clave DEPARTAMENTO-PROVINCIA-DISTRITO), país Perú, tipo de persona según el documento, dígito verificador
del RUC y las casillas T/F en F.
"""
from __future__ import annotations

import json
import re
from functools import lru_cache
from pathlib import Path

RUTA = Path(__file__).resolve().parent.parent / "data" / "proveedores_listas.json"

ENCABEZADOS = [
    "External ID", "Nombre de la empresa", "Etiqueta", "Dirección 1", "Dirección 2", "Ciudad", "Estado", "País",
    "Código ZIP", "Latam - PE Ubigeo", "Default Tax Reg.", "Latam - País de la subsidiaria",
    "Latam - País de la entidad", "Latam - Tipo de Persona", "Latam - Tipo de doc. de identidad",
    "Latam - Tax Reg Number", "Latam - Dígito Verificador", "Latam - Es buen contribuyente?",
    "Latam - Es agente retención?", "Latam - Es agente percepción?", "Latam - Es una Entidad Extranjera?",
    "LATAM PE - CtaCte Detracción",
]
C_ID, C_NOMBRE, C_DIR1 = "External ID", "Nombre de la empresa", "Dirección 1"
C_CIUDAD, C_ESTADO, C_PAIS, C_ZIP, C_UBIGEO = "Ciudad", "Estado", "País", "Código ZIP", "Latam - PE Ubigeo"
C_PAIS_SUB, C_PAIS_ENT = "Latam - País de la subsidiaria", "Latam - País de la entidad"
C_PERSONA, C_TIPO_DOC, C_NUMERO, C_DIGITO = ("Latam - Tipo de Persona", "Latam - Tipo de doc. de identidad",
                                              "Latam - Tax Reg Number", "Latam - Dígito Verificador")
C_CHECKS = ["Latam - Es buen contribuyente?", "Latam - Es agente retención?",
            "Latam - Es agente percepción?", "Latam - Es una Entidad Extranjera?"]

REQUERIDOS = [C_ID, C_NOMBRE, C_DIR1, C_CIUDAD, C_ESTADO, C_PAIS, C_UBIGEO, C_PAIS_SUB, C_PAIS_ENT,
              C_PERSONA, C_TIPO_DOC, C_NUMERO]

NOTA = ("Mínimo por proveedor: Nombre, Dirección 1, Ubigeo, Tipo de doc. de identidad y Número. Se completan solos si "
        "los dejas vacíos: External ID (VEND_ + número), Ciudad y Estado (desde el ubigeo), Código ZIP, País = Peru, "
        "Tipo de Persona, Dígito Verificador (último dígito del RUC) y las casillas T/F en F.")


@lru_cache(maxsize=1)
def datos() -> dict:
    return json.loads(RUTA.read_text(encoding="utf-8"))


def listas() -> dict[str, list[str]]:
    """Opciones de las columnas con lista desplegable en la hoja Vendor."""
    d = datos()
    return {
        C_PAIS: d["paises"], C_PAIS_SUB: d["paises"], C_PAIS_ENT: d["paises"],
        C_PERSONA: d["personas"], C_TIPO_DOC: d["tipos_doc"], C_UBIGEO: d["ubigeos"],
        **{c: ["T", "F"] for c in C_CHECKS},
    }


def _t(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def es_ruc(tipo: str) -> bool:
    t = tipo.upper()
    return "REGISTRO" in t or t == "RUC"


def es_dni(tipo: str) -> bool:
    t = tipo.upper()
    return "DNI" in t or "NACIONAL DE IDENTIDAD" in t


def completar(fila: list, idx) -> list[str]:
    """Rellena (en una copia) lo que se deja vacío; devuelve avisos para quien llama."""
    avisos: list[str] = []
    g = lambda n: _t(fila[idx(n)])  # noqa: E731

    def poner(n: str, valor: str) -> None:
        if valor and not g(n):
            fila[idx(n)] = valor

    ubigeo = g(C_UBIGEO)
    dep = prov = dist = ""
    if ubigeo.count("-") >= 2:
        dep, prov, dist = (x.strip() for x in ubigeo.split("-", 2))
        poner(C_ESTADO, dep.title())
        poner(C_CIUDAD, dist.title())
    if not g(C_ZIP) and ubigeo and g(C_ESTADO) and g(C_CIUDAD) and prov:
        clave = f"{g(C_ESTADO)}-{prov}-{g(C_CIUDAD)}".upper()
        codigo = datos()["zip"].get(clave)
        if codigo:
            fila[idx(C_ZIP)] = codigo
        else:
            avisos.append(f"No se encontró código postal para {clave}")
    for c in (C_PAIS, C_PAIS_SUB, C_PAIS_ENT):
        poner(c, "Peru")
    for c in C_CHECKS:
        poner(c, "F")

    tipo, numero = g(C_TIPO_DOC), g(C_NUMERO)
    if numero and es_ruc(tipo):
        poner(C_PERSONA, {"20": "Jurídico", "10": "Natural"}.get(numero[:2], ""))
        if numero.isdigit() and len(numero) == 11:
            poner(C_DIGITO, numero[-1])
    elif numero and es_dni(tipo):
        poner(C_PERSONA, "Natural")
    return avisos
