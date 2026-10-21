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
    "LATAM PE - CtaCte Detracción", "Tipo", "Subsidiaria", "Nombres", "Apellidos",
]
C_ID, C_NOMBRE, C_DIR1 = "External ID", "Nombre de la empresa", "Dirección 1"
C_CIUDAD, C_ESTADO, C_PAIS, C_ZIP, C_UBIGEO = "Ciudad", "Estado", "País", "Código ZIP", "Latam - PE Ubigeo"
C_PAIS_SUB, C_PAIS_ENT = "Latam - País de la subsidiaria", "Latam - País de la entidad"
C_PERSONA, C_TIPO_DOC, C_NUMERO, C_DIGITO = ("Latam - Tipo de Persona", "Latam - Tipo de doc. de identidad",
                                              "Latam - Tax Reg Number", "Latam - Dígito Verificador")
C_TIPO, C_SUBS, C_NOMBRES, C_APELLIDOS = "Tipo", "Subsidiaria", "Nombres", "Apellidos"
C_CHECKS = ["Latam - Es buen contribuyente?", "Latam - Es agente retención?",
            "Latam - Es agente percepción?", "Latam - Es una Entidad Extranjera?"]

# Datos por defecto para los proveedores que no tienen dirección en el reporte de SUNAT (domicilio de Falcon: Jaén, Cajamarca)
UBIGEO_DEFECTO = "CAJAMARCA-JAEN-JAEN"
DIRECCION_DEFECTO = "JAEN - CAJAMARCA"

# Subsidiaria principal del proveedor (ID interno de Falcon Perú en NetSuite, la misma de las demás hojas)
SUBSIDIARIA = "43"
SUBSIDIARIA_NOMBRE = "FCP"      # nombre de la subsidiaria en el formulario del proveedor (en el CSV)

REQUERIDOS = [C_ID, C_NOMBRE, C_DIR1, C_CIUDAD, C_ESTADO, C_PAIS, C_UBIGEO, C_PAIS_SUB, C_PAIS_ENT,
              C_PERSONA, C_TIPO_DOC, C_NUMERO, C_SUBS]

NOTA = ("Mínimo por proveedor: Nombre, Dirección 1, Ubigeo, Tipo de doc. de identidad y Número. Se completan solos si "
        "los dejas vacíos: External ID (VEND_ + número), Ciudad y Estado (desde el ubigeo), Código ZIP, País = Peru, "
        "Tipo de Persona, Dígito Verificador (último dígito del RUC), las casillas T/F en F, Tipo (8 dígitos = Individual, "
        "11 dígitos = Empresa) y Subsidiaria (43).")


@lru_cache(maxsize=1)
def datos() -> dict:
    return json.loads(RUTA.read_text(encoding="utf-8"))


def listas() -> dict[str, list[str]]:
    """Opciones de las columnas con lista desplegable en la hoja Vendor."""
    d = datos()
    return {
        C_PAIS: d["paises"], C_PAIS_SUB: d["paises"], C_PAIS_ENT: d["paises"],
        C_TIPO: ["Empresa", "Individual"], C_PERSONA: d["personas"], C_TIPO_DOC: d["tipos_doc"], C_UBIGEO: d["ubigeos"],
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


_PARTICULAS = {"de", "del", "la", "las", "los", "y", "san", "santa"}


def dividir_nombre(nombre: str) -> tuple[str, str]:
    """Separa el nombre completo de una persona en (nombres, apellidos) para el formulario de NetSuite (Individual).
    En mayúsculas (como viene de SUNAT para los RUC 10) el orden es «APELLIDO PATERNO MATERNO NOMBRES»; en otro caso
    (reporte de gastos) es «Nombres Apellido Apellido». Las partículas («de la», «del»…) se pegan a la palabra siguiente."""
    palabras = _t(nombre).split()
    fichas: list[str] = []
    pegar = ""
    for w in palabras:
        if w.lower() in _PARTICULAS:
            pegar += w + " "
            continue
        fichas.append(pegar + w)
        pegar = ""
    if pegar:
        fichas.append(pegar.strip())
    if len(fichas) < 2:
        return (fichas[0] if fichas else "", "")
    mayusculas = _t(nombre).isupper()
    if mayusculas:      # APELLIDO APELLIDO NOMBRE(S)
        k = min(2, len(fichas) - 1)
        return " ".join(fichas[k:]), " ".join(fichas[:k])
    k = min(2, len(fichas) - 1)      # Nombre(s) Apellido Apellido
    return " ".join(fichas[:-k]), " ".join(fichas[-k:])


def tipo_entidad(numero: str) -> str:
    """Tipo de proveedor en NetSuite según el Tax Reg Number: 8 dígitos (DNI) = Individual, 11 dígitos (RUC) = Empresa."""
    n = _t(numero)
    return "Individual" if n.isdigit() and len(n) == 8 else "Empresa" if n.isdigit() and len(n) == 11 else ""


def completar(fila: list, idx) -> list:
    """Rellena (en una copia) lo que se deja vacío; devuelve avisos para quien llama (texto, o (columna, texto))."""
    avisos: list[str] = []
    g = lambda n: _t(fila[idx(n)])  # noqa: E731

    def poner(n: str, valor: str) -> None:
        if valor and not g(n):
            fila[idx(n)] = valor

    if not g(C_UBIGEO):       # sin ubigeo (el reporte de SUNAT no trae dirección): se usa el de Jaén para que el proveedor se pueda cargar
        fila[idx(C_UBIGEO)] = UBIGEO_DEFECTO
        avisos.append((C_UBIGEO, f"Sin ubigeo en SUNAT: se usó el de Jaén ({UBIGEO_DEFECTO}); corríjalo si conoce el real"))
    if not g(C_DIR1):
        fila[idx(C_DIR1)] = DIRECCION_DEFECTO
        avisos.append((C_DIR1, f"Sin dirección en SUNAT: se usó «{DIRECCION_DEFECTO}»; corríjala si conoce la real"))
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
    poner(C_TIPO, tipo_entidad(numero))
    poner(C_SUBS, SUBSIDIARIA)
    if g(C_TIPO) == "Individual":
        nombres, apellidos = dividir_nombre(g(C_NOMBRE))
        poner(C_NOMBRES, nombres)
        poner(C_APELLIDOS, apellidos)
        if not g(C_APELLIDOS):
            avisos.append((C_APELLIDOS, "No se pudo separar el nombre de los apellidos; complételos (NetSuite los exige para un proveedor Individual)"))
    if numero and es_ruc(tipo):
        poner(C_PERSONA, {"20": "Jurídico", "10": "Natural", "15": "Natural", "17": "Natural"}.get(numero[:2], ""))
        if numero.isdigit() and len(numero) == 11:
            poner(C_DIGITO, numero[-1])
    elif numero and es_dni(tipo):
        poner(C_PERSONA, "Natural")
    return avisos


# ---------- CSV para NetSuite (formulario «Falcon Vendor Form - Latam Peru») ----------
# Columnas del CSV de Proveedores, en el orden del formulario de NetSuite: información principal, subsidiaria,
# sublista Dirección (ventana de edición de la dirección: País, Dirección 1, Ciudad = distrito, County = departamento,
# Estado = país como en los proveedores ya creados, ZIP) y sección LatamReady. Salen de las columnas de la pestaña Proveedores. External ID sale vacío: lo genera
# NetSuite al registrar el proveedor.
FORMULARIO = "Falcon Vendor Form - Latam Peru"
SALIDA = [
    "Formulario personalizado", "External ID", "Internal ID", "Individual", "Nombre de la empresa", "Nombre", "Apellidos",
    "Buyer ID", "Comentarios", "Subsidiaria principal",
    "Dirección: Envío predeterminado", "Dirección: Facturación predeterminada", "Dirección: Etiqueta", "Dirección: País",
    "Dirección: Atención", "Dirección: Destinatario", "Dirección: Teléfono", "Dirección: Dirección 1", "Dirección: Dirección 2",
    "Dirección: Ciudad", "Dirección: County", "Dirección: Estado", "Dirección: Código ZIP", "Dirección: Legacy ID",
    "Dirección: EDI Location ID #", "Dirección: Latam - PE Ubigeo",
    "Latam Entidad - SUNAT País", "Latam Entidad - Código SUNAT del País",
    "Latam - Tipo de documento de identidad", "Latam - Código del tipo de documento de identidad",
    "Latam - Nro de registro de contribuyente", "Latam - ¿Es buen contribuyente?", "Latam - ¿Agente de retención?",
    "Latam - ¿Agente de percepción?", "Latam Entity - PE CtaCte Detracción", "Latam - Vínculo contrib. y res. extranj.",
    "Latam - Vínculo contributivo y código de residencia extranjera", "Auto-Approve Bills", "Auto-Approve PO",
]
CODIGO_SUNAT_PERU = "9589"
# Catálogo SUNAT N.° 6 (tipo de documento de identidad)
CODIGO_TIPO_DOC = {"DOCUMENTO NACIONAL DE IDENTIDAD (DNI)": "1", "CARNET DE EXTRANJERIA": "4", "REGISTRO ÚNICO DE CONTRIBUYENTES": "6",
                   "PASAPORTE": "7", "CÉDULA DIPLOMÁTICA DE IDENTIDAD": "A", "OTROS TIPOS DE DOCUMENTOS": "0"}


def _t_f(tipo: str) -> str:
    """Casilla «Individual» de NetSuite: T = persona, F = empresa (vacío si no se conoce el tipo)."""
    return "T" if tipo == "Individual" else "F" if tipo == "Empresa" else ""


def _si_no(v: str) -> str:
    return "Sí" if _t(v).upper() in ("T", "TRUE", "SI", "SÍ", "S", "1") else "No"


def tipo_doc_por_numero(numero: str, actual: str = "") -> str:
    """Tipo de documento según el número: 8 dígitos = DNI, 11 dígitos = RUC (si no, se deja el que tenga la fila)."""
    n = _t(numero)
    if n.isdigit() and len(n) == 8:
        return "DOCUMENTO NACIONAL DE IDENTIDAD (DNI)"
    if n.isdigit() and len(n) == 11:
        return "REGISTRO ÚNICO DE CONTRIBUYENTES"
    return actual


def a_salida(encabezados: list[str], filas: list[list]) -> list[list[str]]:
    """Convierte filas (ya completadas, con las columnas de la pestaña Proveedores) a las columnas del CSV para NetSuite."""
    pos = {c: i for i, c in enumerate(encabezados)}
    salida = []
    for f in filas:
        g = lambda n: _t(f[pos[n]]) if n in pos and pos[n] < len(f) else ""  # noqa: E731
        pais = g(C_PAIS)
        salida.append([
            FORMULARIO, "", "", _t_f(g(C_TIPO)), "" if g(C_TIPO) == "Individual" else g(C_NOMBRE), g(C_NOMBRES), g(C_APELLIDOS),
            g(C_NUMERO), "", SUBSIDIARIA_NOMBRE,       # Buyer ID = RUC (o DNI); Comentarios vacío
            "T", "T", "", pais,       # Envío y Facturación predeterminada (casillas: T/F)
            "", "", "", g(C_DIR1), "",
            g(C_CIUDAD), g(C_ESTADO), pais, g(C_ZIP), "",
            "", g(C_UBIGEO),
            "PERU" if pais.lower() in ("peru", "perú") else pais.upper(), CODIGO_SUNAT_PERU if pais.lower() in ("peru", "perú") else "",
            tipo_doc_por_numero(g(C_NUMERO), g(C_TIPO_DOC)), CODIGO_TIPO_DOC.get(tipo_doc_por_numero(g(C_NUMERO), g(C_TIPO_DOC)).upper(), ""),
            g(C_NUMERO), _si_no(g(C_CHECKS[0])), _si_no(g(C_CHECKS[1])),
            _si_no(g(C_CHECKS[2])), g("LATAM PE - CtaCte Detracción"), "", "", "", "",       # Auto-Approve Bills / PO vacías
        ])
    return salida

