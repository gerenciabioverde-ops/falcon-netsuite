"""Carga de un Excel/CSV a las grillas: busca la hoja, detecta la fila de encabezados y mapea por nombre.

Sirve tanto para el Excel de plantillas (una hoja por pestaña, encabezados en la fila 1) como para el libro de
proveedores (hoja «Vendor», encabezados en la fila 3 y filas de ayuda/ejemplo antes de los datos).
"""
from __future__ import annotations

import csv
import io
import warnings
from datetime import date, datetime
from decimal import Decimal, InvalidOperation

from openpyxl import load_workbook

from .esquema import Hoja, Plantilla, norm

MAX_FILAS = 2000
ALIAS_HOJAS = {"vendor": "proveedores", "vendors": "proveedores", "proveedor": "proveedores"}
# Etiquetas de la primera columna de las filas de ayuda del libro de carga (no son datos)
FILAS_DE_AYUDA = {"campos", "oracle netsuite standard", "tipo de campo", "mandatory", "mandatoriedad", "description",
                  "descripcion", "descripcción", "ejemplo", "comentarios", "fields"}


def _texto(v) -> str:
    if v is None:
        return ""
    if isinstance(v, (datetime, date)):
        return v.strftime("%d/%m/%Y")
    if isinstance(v, bool):
        return "T" if v else "F"
    if isinstance(v, float) or isinstance(v, int):
        try:
            return format(Decimal(str(v)).normalize(), "f")
        except InvalidOperation:
            return str(v)
    return str(v).strip()


def _encabezados(filas: list[list], hoja: Hoja) -> tuple[int, dict[int, int]] | None:
    """(índice de la fila de encabezados, {columna de la grilla: columna del archivo}) o None."""
    objetivos = [norm(e) for e in hoja.encabezados]
    mejor: tuple[int, dict[int, int]] | None = None
    for i, fila in enumerate(filas[:25]):
        presentes = [norm(c) for c in fila]
        usados: set[int] = set()
        mapa: dict[int, int] = {}
        for g, nombre in enumerate(objetivos):  # duplicados (p. ej. «Memo» dos veces): se asignan en orden
            for j, cabecera in enumerate(presentes):
                if cabecera == nombre and j not in usados:
                    mapa[g] = j
                    usados.add(j)
                    break
        if len(mapa) >= 3 and (mejor is None or len(mapa) > len(mejor[1])):
            mejor = (i, mapa)
    return mejor


def _filas_de_datos(filas: list[list], hoja: Hoja, avisos: list[str]) -> list[list[str]] | None:
    enc = _encabezados(filas, hoja)
    if enc is None:
        return None
    inicio, mapa = enc
    salida: list[list[str]] = []
    for fila in filas[inicio + 1:]:
        if fila and norm(fila[0]) in FILAS_DE_AYUDA:
            continue
        celdas = [""] * len(hoja.encabezados)
        for g, j in mapa.items():
            if j < len(fila):
                celdas[g] = _texto(fila[j])
        if any(celdas):
            salida.append(celdas)
    if hoja.nombre == "Proveedores":  # «No encontrado» (resultado del XLOOKUP del libro) no es un código postal
        i = hoja.idx("Código ZIP")
        for c in salida:
            if c[i] and not c[i].isdigit():
                c[i] = ""
    if len(salida) > MAX_FILAS:
        avisos.append(f"«{hoja.nombre.strip()}» tiene {len(salida)} filas; se cargaron las primeras {MAX_FILAS}.")
        salida = salida[:MAX_FILAS]
    return salida


def leer(nombre_archivo: str, contenido: bytes, plantilla: Plantilla, clave_csv: str = "") -> tuple[dict[str, list[list[str]]], list[str]]:
    """Devuelve ({clave de pestaña: filas}, avisos). Lanza ValueError si el archivo no sirve."""
    avisos: list[str] = []
    nombre = nombre_archivo.lower()
    resultado: dict[str, list[list[str]]] = {}

    if nombre.endswith(".csv"):
        if clave_csv not in plantilla.hojas:
            raise ValueError("Para un CSV cárguelo desde la pestaña a la que corresponde.")
        texto = contenido.decode("utf-8-sig", errors="replace")
        primera = texto.split("\n", 1)[0]
        delim = ";" if primera.count(";") > primera.count(",") else ","
        filas = [list(f) for f in csv.reader(io.StringIO(texto), delimiter=delim)]
        datos = _filas_de_datos(filas, plantilla.hojas[clave_csv], avisos)
        if datos is None:
            raise ValueError("No reconocí los encabezados del CSV para esta pestaña.")
        return {clave_csv: datos}, avisos

    if not nombre.endswith((".xlsx", ".xlsm")):
        raise ValueError("Suba un archivo .xlsx o .csv")
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb = load_workbook(io.BytesIO(contenido), read_only=True, data_only=True)
    for hoja_xl in wb.sheetnames:
        clave = ALIAS_HOJAS.get(norm(hoja_xl), norm(hoja_xl))
        if clave not in plantilla.hojas or clave in resultado:
            continue
        filas = [list(f) for f in wb[hoja_xl].iter_rows(values_only=True)]
        datos = _filas_de_datos(filas, plantilla.hojas[clave], avisos)
        if datos is None:
            avisos.append(f"La hoja «{hoja_xl.strip()}» no tiene los encabezados esperados; se omitió.")
        elif datos:
            resultado[clave] = datos
    if not resultado:
        raise ValueError("No encontré datos en ninguna hoja conocida (por ejemplo «Vendor» o las hojas de la plantilla).")
    return resultado, avisos
