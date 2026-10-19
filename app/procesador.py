"""Lectura del archivo subido, validación contra la plantilla y exportación a CSV NetSuite."""
from __future__ import annotations

import csv
import io
import re
import warnings
import zipfile
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal, InvalidOperation

from openpyxl import load_workbook

from .esquema import Hoja, Plantilla, norm

FORMATOS_FECHA = {
    "DD/MM/YYYY": "%d/%m/%Y",
    "MM/DD/YYYY": "%m/%d/%Y",
    "YYYY-MM-DD": "%Y-%m-%d",
}
_PERIODO = re.compile(r"^[A-Za-z]{3,9} \d{4}$")
_SEGURO = re.compile(r"^[\d.\s+\-*/()]+$")


@dataclass
class Problema:
    nivel: str  # "error" | "aviso"
    fila: int | None
    columna: str
    mensaje: str


@dataclass
class ResultadoHoja:
    nombre: str
    encabezados: list[str]
    filas: list[list[str]] = field(default_factory=list)  # ya formateadas para CSV
    numeros_fila: list[int] = field(default_factory=list)
    problemas: list[Problema] = field(default_factory=list)

    @property
    def errores(self) -> int:
        return sum(p.nivel == "error" for p in self.problemas)

    @property
    def avisos(self) -> int:
        return sum(p.nivel == "aviso" for p in self.problemas)

    @property
    def exportable(self) -> bool:
        return bool(self.filas) and self.errores == 0

    def csv_bytes(self) -> bytes:
        buf = io.StringIO(newline="")
        w = csv.writer(buf, lineterminator="\r\n")
        w.writerow(self.encabezados)
        w.writerows(self.filas)
        return buf.getvalue().encode("utf-8-sig")

    @property
    def archivo(self) -> str:
        return re.sub(r"\W+", "_", self.nombre.strip()).strip("_") + ".csv"


# ---------- lectura ----------

def _evaluar(formula: str, ws_f, ws_v, prof: int = 0):
    expr = formula[1:]
    if prof > 5:
        raise ValueError("fórmula demasiado anidada")

    def ref(m):
        c = ws_f[f"{m.group(1)}{m.group(2)}"]
        v = _valor_celda(c, ws_f, ws_v, prof + 1)
        if v in (None, ""):
            return "0"
        try:
            return str(Decimal(str(v)))
        except InvalidOperation:
            raise ValueError(f"la fórmula usa «{v}», que no es un número")

    sust = re.sub(r"\$?([A-Z]{1,3})\$?(\d+)", ref, expr)
    if not _SEGURO.match(sust):
        raise ValueError(f"fórmula no soportada: {formula}")
    sust = re.sub(r"\d+(?:\.\d+)?", lambda m: f"D('{m.group()}')", sust)
    return eval(sust, {"__builtins__": {}, "D": Decimal})  # noqa: S307 (expresión ya validada)


def _valor_celda(c, ws_f, ws_v, prof=0):
    v = c.value
    if isinstance(v, str) and v.startswith("="):
        cache = ws_v[c.coordinate].value if ws_v is not None else None
        if cache is not None:
            return cache
        return _evaluar(v, ws_f, ws_v, prof)
    return v


def leer_xlsx(contenido: bytes, plantilla: Plantilla) -> tuple[dict[str, list[tuple[int, list]]], list[str]]:
    """Devuelve {clave_hoja: [(nro_fila_excel, [valores])]} y avisos generales."""
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wf = load_workbook(io.BytesIO(contenido))
        wv = load_workbook(io.BytesIO(contenido), data_only=True)
    datos, avisos = {}, []
    for nombre in wf.sheetnames:
        clave = norm(nombre)
        if clave not in plantilla.hojas:
            if "maestra" not in clave:
                avisos.append(f"Hoja «{nombre.strip()}» no corresponde a ninguna plantilla y se ignoró.")
            continue
        ws_f, ws_v = wf[nombre], wv[nombre]
        n = len(plantilla.hojas[clave].encabezados)
        filas = [(1, [c.value for c in ws_f[1][:n]])]
        for r in range(2, ws_f.max_row + 1):
            fila = []
            for c in ws_f[r][:n]:
                try:
                    fila.append(_valor_celda(c, ws_f, ws_v))
                except Exception as e:  # fórmula no evaluable
                    fila.append(f"#ERROR {e}")
            fila += [None] * (n - len(fila))
            if any(v not in (None, "") for v in fila):
                filas.append((r, fila))
        datos[clave] = filas
    return datos, avisos


def leer_csv(contenido: bytes, clave_hoja: str, plantilla: Plantilla):
    texto = contenido.decode("utf-8-sig", errors="replace")
    delim = ";" if texto.split("\n", 1)[0].count(";") > texto.split("\n", 1)[0].count(",") else ","
    filas = list(csv.reader(io.StringIO(texto), delimiter=delim))
    n = len(plantilla.hojas[clave_hoja].encabezados)
    salida = []
    for i, f in enumerate(filas, start=1):
        f = (f + [None] * n)[:n]
        f = [None if (isinstance(v, str) and v == "") else v for v in f]
        if i == 1 or any(v is not None for v in f):
            salida.append((i, f))
    return {clave_hoja: salida}, []


# ---------- validación ----------

def _num(v) -> Decimal | None:
    if v in (None, ""):
        return None
    if isinstance(v, str):
        v = v.strip().replace(",", ".") if v.count(",") == 1 and "." not in v else v.strip()
    try:
        return Decimal(str(v))
    except InvalidOperation:
        raise ValueError(f"«{v}» no es un número")


def _fecha(v, fmt_entrada=("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y")) -> date:
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    for f in fmt_entrada:
        try:
            return datetime.strptime(str(v).strip(), f).date()
        except ValueError:
            pass
    raise ValueError(f"«{v}» no es una fecha válida (use DD/MM/AAAA)")


def _fmt_num(d: Decimal) -> str:
    s = format(d.normalize(), "f")
    return s


def validar_hoja(hoja: Hoja, filas: list[tuple[int, list]], formato_fecha: str) -> ResultadoHoja:
    res = ResultadoHoja(nombre=hoja.nombre, encabezados=hoja.encabezados)
    if not filas:
        return res
    r = hoja.reglas
    enc_recibido = filas[0][1]
    for i, esperado in enumerate(hoja.encabezados):
        if norm(enc_recibido[i]) != norm(esperado):
            res.problemas.append(Problema("error", 1, esperado,
                                          f"Encabezado esperado «{esperado.strip()}», recibido «{enc_recibido[i] or ''}»"))
    if res.errores:
        return res

    idx = hoja.idx
    fmt_out = FORMATOS_FECHA[formato_fecha]
    requeridos = [idx(c) for c in r.get("requeridos", [])]
    fechas = {idx(c) for c in r.get("fechas", [])}
    numericos = {idx(c) for c in r.get("numericos", [])}
    col_periodo = idx(r["periodo"]) if r.get("periodo") else None
    desde = idx(r["lineas_desde"]) if r.get("lineas_desde") else None
    col_id = 0
    cabeceras: dict[str, list[str]] = {}
    saldos: dict[str, Decimal] = {}

    for nro, fila in filas[1:]:
        salida: list[str] = []
        for i, v in enumerate(fila):
            nom = hoja.encabezados[i].strip()
            vacio = v is None or (isinstance(v, str) and not v.strip())
            if vacio:
                if i in requeridos:
                    res.problemas.append(Problema("error", nro, nom, "Campo obligatorio vacío"))
                salida.append("")
                continue
            try:
                if isinstance(v, str) and v.startswith("#ERROR"):
                    raise ValueError(v[1:])
                if i in fechas:
                    texto = _fecha(v).strftime(fmt_out)
                elif i in numericos:
                    texto = _fmt_num(_num(v))
                elif isinstance(v, float) and v.is_integer():
                    texto = str(int(v))
                elif isinstance(v, (datetime, date)):
                    texto = _fecha(v).strftime(fmt_out)
                else:
                    texto = str(v).strip()
                if i == col_periodo and not _PERIODO.match(texto):
                    raise ValueError(f"«{texto}» no tiene el formato de periodo «Oct 2026»")
                lista = hoja.listas.get(i)
                if lista and (lista.valores or lista.crudo) and not lista.acepta(texto):
                    res.problemas.append(Problema("error", nro, nom, f"«{texto[:60]}» no está en la lista permitida"))
                salida.append(texto)
            except ValueError as e:
                res.problemas.append(Problema("error", nro, nom, str(e)))
                salida.append(str(v))

        for a, b in r.get("alguno_de", []):
            if not salida[idx(a)] and not salida[idx(b)]:
                res.problemas.append(Problema("error", nro, f"{a} / {b}", "Debe informar al menos una de las dos líneas"))
        for si, entonces in r.get("dependientes", []):
            if salida[idx(si)]:
                for c in entonces:
                    if not salida[idx(c)]:
                        res.problemas.append(Problema("error", nro, c, f"Obligatorio cuando se informa «{si}»"))

        if "cuadre" in r:
            d, c = (salida[idx(x)] for x in r["cuadre"])
            if bool(d) == bool(c):
                res.problemas.append(Problema("error", nro, "Debit / Credit", "Cada línea debe tener Débito o Crédito (solo uno)"))
            else:
                try:
                    saldos[salida[col_id]] = saldos.get(salida[col_id], Decimal(0)) + Decimal(d or 0) - Decimal(c or 0)
                except InvalidOperation:
                    pass

        if desde is not None and salida[col_id]:
            cab = salida[:desde]
            previa = cabeceras.setdefault(salida[col_id], cab)
            if previa is not cab and previa != cab:
                dif = [hoja.encabezados[i].strip() for i in range(desde) if previa[i] != cab[i]]
                res.problemas.append(Problema("aviso", nro, ", ".join(dif[:3]),
                                              f"La cabecera difiere de la primera línea del ID «{salida[col_id]}»"))
        res.filas.append(salida)
        res.numeros_fila.append(nro)

    for ext, saldo in saldos.items():
        if saldo != 0:
            res.problemas.append(Problema("error", None, ext, f"Asiento descuadrado: Débito − Crédito = {_fmt_num(saldo)}"))
    return res


def procesar(datos: dict[str, list], plantilla: Plantilla, formato_fecha: str) -> list[ResultadoHoja]:
    return [validar_hoja(plantilla.hojas[k], filas, formato_fecha) for k, filas in datos.items()]


def empaquetar_zip(resultados: list[ResultadoHoja]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for r in resultados:
            if r.exportable:
                z.writestr(r.archivo, r.csv_bytes())
    return buf.getvalue()
