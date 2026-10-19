"""Esquema de las plantillas NetSuite: encabezados, listas desplegables y reglas.

Los encabezados y las listas (validaciones de datos + DATA MAESTRA) se leen
directamente de `data/TEMPLATE_CSV_FALCON.xlsx`; las reglas de obligatoriedad
están en REGLAS y se ajustan aquí si NetSuite pide otros campos.
"""
from __future__ import annotations

import re
import warnings
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils.cell import column_index_from_string

RUTA_PLANTILLA = Path(__file__).resolve().parent.parent / "data" / "TEMPLATE_CSV_FALCON.xlsx"
HOJA_MAESTRA = "DATA MAESTRA"

REGLAS: dict[str, dict] = {
    "Factura de Proveedor": {
        "requeridos": ["EXTERNAL ID", "Proveedor", "Cuenta", "Moneda", "Fecha", "Subsidiaria",
                       "Latam - Tipo de Transacción", "Latam - Tipo de Documento Fiscal",
                       "Latam - Serie CxP", "Latam - Número Preimpreso"],
        "fechas": ["Fecha", "Latam - Fecha del documento Ref"],
        "numericos": ["Artículos: Cantidad", "Artículos: Tarifa", "Artículos: Importe", "Gastos: Importe"],
        "alguno_de": [["Artículos: Artículo", "Gastos: Cuenta"]],
        "dependientes": [
            ("Artículos: Artículo", ["Artículos: Cantidad", "Artículos: Importe"]),
            ("Gastos: Cuenta", ["Gastos: Importe"]),
        ],
        "lineas_desde": "Artículos: Artículo",
        "autocalcular": [("Artículos: Importe", "Artículos: Cantidad", "Artículos: Tarifa")],
        "cuentas": ["Cuenta", "Gastos: Cuenta"],
        "id_auto": {"prefijo": "VENDB", "campos": ["Proveedor", "Latam - Serie CxP", "Latam - Número Preimpreso"]},
    },
    "Pago Proveedor": {
        "requeridos": ["ExternalID", "A/P ACCOUNT", "Beneficiario", "ACCOUNT", "CURRENCY", "EXCHANGE RATE",
                       "DATE", "POSTING PERIOD", "SUBSIDIARY", "BILL", "PAYMENT"],
        "fechas": ["DATE"],
        "numericos": ["EXCHANGE RATE", "PAYMENT"],
        "periodo": "POSTING PERIOD",
        "lineas_desde": "BILL",
        "cuentas": ["A/P ACCOUNT", "ACCOUNT"],
        "id_auto": {"prefijo": "BP_PE", "campos": ["Beneficiario", "BILL", "DATE"]},
    },
    "Asiento": {
        "requeridos": ["EXTERNAL ID", "Date", "POSTING PERIOD", "Subsidiary", "Currency", "Exchange Rate", "Account"],
        "fechas": ["Date"],
        "numericos": ["Exchange Rate", "Debit", "Credit"],
        "periodo": "POSTING PERIOD",
        "lineas_desde": "Account",
        "cuadre": ("Debit", "Credit"),
        "cuentas": ["Account"],
        "id_auto": {"prefijo": "Asiento", "modo": "asiento"},
    },
    "Factura de Cliente": {
        "requeridos": ["EXTERNAL ID", "Cliente", "Cuenta", "Moneda", "Fecha", "Subsidiaria",
                       "Latam - Tipo de Documento Fiscal", "Latam - Serie CxC", "Latam - Número Preimpreso",
                       "Artículos: Artículo", "Artículos: Cantidad", "Artículos: Tarifa", "Artículos: Tax Code"],
        "fechas": ["Fecha"],
        "numericos": ["Artículos: Cantidad", "Artículos: Tarifa", "Artículos: Importe"],
        "lineas_desde": "Artículos: Artículo",
        "autocalcular": [("Artículos: Importe", "Artículos: Cantidad", "Artículos: Tarifa")],
        "cuentas": ["Cuenta"],
        "id_auto": {"prefijo": "FACT_VENTA", "campos": ["Cliente", "Latam - Serie CxC", "Latam - Número Preimpreso"]},
    },
    "Nota de credito Cliente": {
        "requeridos": ["EXTERNAL ID", "Cliente", "Cuenta", "Moneda", "Fecha", "Subsidiaria",
                       "Latam - Tipo de Documento Fiscal", "Latam - Serie CxC", "Latam - Número Preimpreso",
                       "Latam - Modification Reason", "Latam - Tipo Documento Ref", "Latam - Documento Serie Ref",
                       "Latam - Documento Numero Ref", "Latam - Fecha del documento Ref",
                       "Artículos: Artículo", "Artículos: Cantidad", "Artículos: Tarifa", "Artículos: Tax Code"],
        "fechas": ["Fecha", "Latam - Fecha del documento Ref"],
        "numericos": ["Artículos: Cantidad", "Artículos: Tarifa", "Artículos: Importe", "Monto aplicado"],
        "lineas_desde": "Artículos: Artículo",
        "autocalcular": [("Artículos: Importe", "Artículos: Cantidad", "Artículos: Tarifa")],
        "cuentas": ["Cuenta"],
        "id_auto": {"prefijo": "NotaCredito", "campos": ["Cliente", "Latam - Serie CxC", "Latam - Número Preimpreso"]},
    },
    "Pago de factura de cliente": {
        "requeridos": ["ExternalID", "A/P ACCOUNT", "Beneficiario", "ACCOUNT", "CURRENCY", "EXCHANGE RATE",
                       "DATE", "POSTING PERIOD", "SUBSIDIARY", "Aplica", "Pago"],
        "fechas": ["DATE"],
        "numericos": ["EXCHANGE RATE", "Pago"],
        "periodo": "POSTING PERIOD",
        "lineas_desde": "Aplica",
        "cuentas": ["A/P ACCOUNT", "ACCOUNT"],
        "id_auto": {"prefijo": "Pago_FCliente", "campos": ["Beneficiario", "Aplica", "DATE"]},
    },
}


def norm(txt) -> str:
    return re.sub(r"\s+", " ", str(txt or "")).strip().lower()


@dataclass
class Lista:
    """Lista desplegable: valores exactos o texto crudo (cuando hay comas dentro de un valor)."""
    valores: set[str] = field(default_factory=set)
    crudo: str = ""
    opciones: list[str] = field(default_factory=list)  # textos originales para ofrecer en pantalla

    def acepta(self, valor: str) -> bool:
        v = norm(valor)
        if v in self.valores:
            return True
        return bool(self.crudo) and f",{v}," in f",{norm(self.crudo)},"


@dataclass
class Hoja:
    nombre: str
    encabezados: list[str]
    listas: dict[int, Lista]  # índice de columna (0-based) -> lista
    reglas: dict

    def idx(self, nombre: str) -> int:
        n = norm(nombre)
        for i, h in enumerate(self.encabezados):
            if norm(h) == n:
                return i
        raise KeyError(nombre)


@dataclass
class Plantilla:
    hojas: dict[str, Hoja]  # clave: norm(nombre)
    maestra: dict[str, list[str]]  # catálogos de DATA MAESTRA por encabezado "Nombre" (informativo)


_NS_X14 = "{http://schemas.microsoft.com/office/spreadsheetml/2009/9/main}"
_NS_XM = "{http://schemas.microsoft.com/office/excel/2006/main}"
_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def _rutas_hojas(z: zipfile.ZipFile) -> dict[str, str]:
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    destino = {r.get("Id"): r.get("Target") for r in rels}
    rid = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    salida = {}
    for s in wb.iter(f"{_NS}sheet"):
        t = destino[s.get(rid)]
        salida[s.get("name")] = "xl/" + t.lstrip("/").removeprefix("xl/")
    return salida


def _columnas(sqref: str) -> list[int]:
    cols = []
    for tok in sqref.split():
        letras = re.match(r"[A-Z]+", tok.split(":")[0])
        if letras:
            cols.append(column_index_from_string(letras.group()) - 1)
    return cols


def _resolver_formula(f: str, wb, hoja_actual: str) -> Lista:
    f = f.replace('"&"', "").strip()
    if f.startswith('"'):
        crudo = f.strip('"')
        partes = crudo.split(",")
        # si algún valor lleva comas internas, partirlo por comas sería engañoso
        ambigua = any(x != x.strip() for x in partes)
        return Lista(valores={norm(x) for x in partes}, crudo=crudo, opciones=[] if ambigua else partes)
    m = re.match(r"(?:'?([^'!]+)'?!)?\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)", f)
    if not m:
        return Lista()
    ws = wb[m.group(1) or hoja_actual]
    vals, ops = set(), []
    for fila in ws[f"{m.group(2)}{m.group(3)}:{m.group(4)}{m.group(5)}"]:
        for c in fila:
            if c.value is not None:
                vals.add(norm(c.value))
                ops.append(str(c.value).strip())
    return Lista(valores=vals, opciones=ops)


def _listas_de_hoja(xml: bytes, wb, nombre: str) -> dict[int, Lista]:
    raiz = ET.fromstring(xml)
    listas: dict[int, Lista] = {}
    for dv in raiz.iter():
        if not dv.tag.endswith("dataValidation") or dv.get("type") != "list":
            continue
        sqref = dv.get("sqref") or (dv.findtext(f"{_NS_XM}sqref") or "")
        f = dv.findtext(f"{_NS}formula1") or dv.findtext(f"{_NS_X14}formula1/{_NS_XM}f") or ""
        if not f or not sqref:
            continue
        lista = _resolver_formula(f, wb, nombre)
        for col in _columnas(sqref):
            listas[col] = lista
    return listas


def cargar_plantilla(ruta: Path = RUTA_PLANTILLA) -> Plantilla:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb = load_workbook(ruta, data_only=True)
    hojas: dict[str, Hoja] = {}
    with zipfile.ZipFile(ruta) as z:
        rutas = _rutas_hojas(z)
        for nombre, ws in ((n, wb[n]) for n in wb.sheetnames):
            if norm(nombre) == norm(HOJA_MAESTRA):
                continue
            enc = [str(c.value) if c.value is not None else "" for c in ws[1]]
            while enc and not enc[-1].strip():
                enc.pop()
            clave = next((k for k in REGLAS if norm(k) == norm(nombre)), None)
            hojas[norm(nombre)] = Hoja(
                nombre=nombre, encabezados=enc,
                listas=_listas_de_hoja(z.read(rutas[nombre]), wb, nombre),
                reglas=REGLAS.get(clave, {}),
            )
    maestra = {}
    if HOJA_MAESTRA in wb.sheetnames:
        ws = wb[HOJA_MAESTRA]
        for col in range(1, ws.max_column + 1):
            vals = [ws.cell(r, col).value for r in range(2, ws.max_row + 1)]
            vals = [str(v) for v in vals if v is not None]
            if vals:
                maestra[f"{col}:{ws.cell(1, col).value}"] = vals
    return Plantilla(hojas=hojas, maestra=maestra)
