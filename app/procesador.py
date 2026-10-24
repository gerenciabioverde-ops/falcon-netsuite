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

from . import esquema, gastos_cuentas, homologacion, ids, periodos, proveedores
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
    encabezados_origen: list[str] | None = None  # columnas de la pestaña, cuando el CSV sale con otras (Proveedores)
    cuentas: list[str] = field(default_factory=list)  # columnas con cuentas: el CSV lleva el ID interno, el Excel el número del COA

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

    @property
    def archivo_xlsx(self) -> str:
        return self.archivo[:-4] + ".xlsx"

    def xlsx_bytes(self) -> bytes:
        """Las mismas filas del CSV en un libro de Excel (todo como texto, para no perder ceros ni formatos)."""
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = re.sub(r"[\\/*?:\[\]]", " ", self.nombre.strip())[:31] or "Hoja"
        ws.append(list(self.encabezados))
        # en el reporte las cuentas van como en el sistema (número del COA de NetSuite: 1203.00, 42100); el CSV lleva el ID interno
        id_a_coa = {i: n for n, i in homologacion.coa_a_id().items()}
        cols_cuenta = {self.encabezados.index(c) for c in self.cuentas if c in self.encabezados}
        for fila in self.filas:
            ws.append([("" if c is None else str(id_a_coa.get(str(c).strip(), c)) if k in cols_cuenta else ("" if c is None else str(c)))
                       for k, c in enumerate(fila)])
        for celda in ws[1]:
            celda.font = celda.font.copy(bold=True)
        for fila in ws.iter_rows(min_row=2):
            for celda in fila:
                celda.number_format = "@"
        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()


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
        v = v.strip()
        if re.fullmatch(r"-?\d{1,3}(,\d{3})+(\.\d+)?", v):          # 1,308.00 -> 1308.00
            v = v.replace(",", "")
        elif v.count(",") == 1 and "." not in v:                       # 12,5 -> 12.5
            v = v.replace(",", ".")
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


def _clave(v, es_fecha: bool = False) -> str:
    """Texto apto para formar un EXTERNAL ID (solo letras, números y guiones)."""
    if v is None or (isinstance(v, str) and not v.strip()):
        return ""
    if es_fecha:
        try:
            return _fecha(v).strftime("%Y%m%d")
        except ValueError:
            return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return re.sub(r"[^A-Za-z0-9]+", "-", str(v).strip()).strip("-")


def generar_ids(hoja: Hoja, filas: list[tuple[int, list]]) -> list[tuple[int, list]]:
    """Completa EXTERNAL ID vacíos de forma determinista: el mismo comprobante da siempre el mismo ID.

    - Facturas/notas/pagos: prefijo + campos clave (p. ej. VENDB_296720_F001_00000012).
    - Asientos: líneas consecutivas sin ID hasta que Débito = Crédito forman un asiento
      (FCPAAAAMM_0001: prefijo + año/mes de la fecha + correlativo del mes, que continúa tras los IDs ya escritos).
    """
    spec = hoja.reglas.get("id_auto")
    if not spec:
        return filas
    fechas = {hoja.idx(c) for c in hoja.reglas.get("fechas", [])}
    salida = [(nro, list(fila)) for nro, fila in filas]
    vacio = lambda v: v is None or (isinstance(v, str) and not v.strip())  # noqa: E731

    if spec.get("modo") == "asiento":
        i_deb, i_cre, i_fec = hoja.idx("Debit"), hoja.idx("Credit"), hoja.idx("Date")
        grupo: list[int] = []
        saldo = Decimal(0)
        correlativo: dict[str, int] = {}  # año/mes -> último número usado (parte de los IDs que el usuario ya escribió)
        patron = re.compile(rf"^{re.escape(spec['prefijo'])}(\d{{6}})_(\d+)$")
        for _, fila in salida:
            m = patron.match(str(fila[0] or "").strip())
            if m:
                correlativo[m.group(1)] = max(correlativo.get(m.group(1), 0), int(m.group(2)))
        for k, (nro, fila) in enumerate(salida):
            if not vacio(fila[0]):
                grupo, saldo = [], Decimal(0)
                continue
            grupo.append(k)
            try:
                saldo += (_num(fila[i_deb]) or 0) - (_num(fila[i_cre]) or 0)
            except ValueError:
                grupo, saldo = [], Decimal(0)
                continue
            if saldo == 0:
                mes = _clave(salida[grupo[0]][1][i_fec], True)[:6]
                if mes:
                    correlativo[mes] = correlativo.get(mes, 0) + 1
                    ident = f"{spec['prefijo']}{mes}_{correlativo[mes]:04d}"
                    for g in grupo:
                        salida[g][1][0] = ident
                grupo = []
        return salida

    cols = [(hoja.idx(c), hoja.idx(c) in fechas) for c in spec["campos"]]
    for nro, fila in salida:
        if vacio(fila[0]):
            if spec.get("sin_separador"):       # VENDB_ + RUC + Serie + Número, todo junto: VENDB_20513497769FA0100018047
                crudo = [("" if vacio(fila[i]) else str(fila[i]).strip()) for i, _ in cols]
                quitar = spec.get("quitar_primero", "")
                if quitar and crudo[0].startswith(quitar):
                    crudo[0] = crudo[0][len(quitar):]
                fila[0] = ids.id_comprobante(spec["prefijo"], *crudo) or fila[0]
                continue
            partes = [_clave(fila[i], f) for i, f in cols]
            if all(partes):
                fila[0] = "_".join([spec["prefijo"], *partes])
    return salida


def _ruc_valido(ruc: str) -> bool:
    """Dígito verificador del RUC (módulo 11, pesos de SUNAT)."""
    pesos = (5, 4, 3, 2, 7, 6, 5, 4, 3, 2)
    resto = 11 - sum(int(d) * p for d, p in zip(ruc[:10], pesos)) % 11
    return int(ruc[10]) == {10: 0, 11: 1}.get(resto, resto)


def validar_hoja(hoja: Hoja, filas: list[tuple[int, list]], formato_fecha: str, periodo: str | None = None) -> ResultadoHoja:
    res = ResultadoHoja(nombre=hoja.nombre, encabezados=hoja.encabezados,
                        cuentas=list(hoja.reglas.get("cuentas", [])))
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
    avisos_completar: dict[int, list[str]] = {}
    if r.get("completar") == "proveedores":
        completadas = []
        for nro, fila in filas[1:]:
            fila = list(fila)
            avisos_completar[nro] = proveedores.completar(fila, idx)
            completadas.append((nro, fila))
        filas = [filas[0], *completadas]
    etiq_periodo = periodos.etiqueta(periodo) if periodo and periodos.es_valido(periodo) else ""
    if etiq_periodo and r.get("periodo"):  # POSTING PERIOD vacío = período seleccionado
        i_per = idx(r["periodo"])
        filas = [filas[0], *[(n, [etiq_periodo if k == i_per and f[k] in (None, "") else f[k] for k in range(len(f))]) for n, f in filas[1:]]]
    for col, valor in r.get("por_defecto", {}).items():  # valor fijo cuando la celda se deja vacía (subsidiaria, ubicación)
        i_def = idx(col)
        valor, i_si = valor if isinstance(valor, tuple) else (valor, None)  # (valor, columna que debe tener dato)
        i_si = idx(i_si) if i_si else None
        filas = [filas[0], *[(n, [valor if k == i_def and f[k] in (None, "") and (i_si is None or f[i_si] not in (None, "")) else f[k]
                                  for k in range(len(f))]) for n, f in filas[1:]]]
    filas = [filas[0], *generar_ids(hoja, filas[1:])]
    fmt_out = FORMATOS_FECHA[formato_fecha]
    requeridos = [idx(c) for c in r.get("requeridos", [])]
    fechas = {idx(c) for c in r.get("fechas", [])}
    numericos = {idx(c) for c in r.get("numericos", [])}
    col_periodo = idx(r["periodo"]) if r.get("periodo") else None
    desde = idx(r["lineas_desde"]) if r.get("lineas_desde") else None
    col_id = 0
    cabeceras: dict[str, list[str]] = {}
    saldos: dict[str, Decimal] = {}

    autocalc = [(idx(i), idx(c), idx(t)) for i, c, t in r.get("autocalcular", [])]
    cuentas = {idx(c) for c in r.get("cuentas", [])}
    mapa_h = homologacion.mapa() if cuentas else {}
    coa_h = homologacion.coa_a_id() if cuentas else {}
    traducidas: dict[tuple, list[int]] = {}
    sin_id: dict[tuple, list[int]] = {}
    ids_cuenta = set(coa_h.values()) | set(mapa_h.values())
    ids_vistos: dict[str, int] = {}
    claves_vistas: dict[tuple, int] = {}

    for nro, fila in filas[1:]:
        fila = list(fila)
        for i_imp, i_cant, i_tar in autocalc:  # Importe = Cantidad × Tarifa si se deja vacío
            if fila[i_imp] in (None, "") and fila[i_cant] not in (None, "") and fila[i_tar] not in (None, ""):
                try:
                    fila[i_imp] = _num(fila[i_cant]) * _num(fila[i_tar])
                except (ValueError, TypeError, InvalidOperation):
                    pass
        salida: list[str] = []
        for i, v in enumerate(fila):
            nom = hoja.encabezados[i].strip()
            vacio = v is None or (isinstance(v, str) and not v.strip())
            if vacio:
                if i == 0 and r.get("id_auto"):
                    falta = ("las líneas consecutivas sin ID no cuadran (Débito ≠ Crédito)" if r["id_auto"].get("modo") == "asiento"
                             else "complete " + ", ".join(r["id_auto"]["campos"]))
                    res.problemas.append(Problema("error", nro, nom, f"No se pudo generar el EXTERNAL ID: {falta}"))
                elif i in requeridos:
                    opcional = r.get("opcional_si", {}).get(nom)       # p. ej. EXCHANGE RATE en USD: NetSuite usa su tipo de cambio de la fecha
                    if opcional and str(fila[idx(opcional[0])] or "").strip().upper() == opcional[1]:
                        pass
                    # una persona (Individual) lleva Nombres y Apellidos, no «Nombre de la empresa»
                    elif r.get("completar") == "proveedores" and nom == proveedores.C_NOMBRE and (
                            str(fila[idx(proveedores.C_TIPO)] or "").strip() == "Individual"
                            or any(str(fila[idx(c)] or "").strip() for c in (proveedores.C_NOMBRES, proveedores.C_APELLIDOS))):
                        pass
                    else:
                        res.problemas.append(Problema("error", nro, nom, "Campo obligatorio vacío"))
                salida.append("")
                continue
            try:
                if isinstance(v, str) and v.startswith("#ERROR"):
                    raise ValueError(v[1:])
                if i in cuentas and mapa_h:  # cuenta de Falcon (PCGE) o número del COA de NetSuite -> ID interno de NetSuite
                    nuevo = mapa_h.get(homologacion.normalizar(v)) or coa_h.get(homologacion.normalizar(v))
                    if nuevo is not None and nuevo != str(v).strip():
                        traducidas.setdefault((nom, str(v).strip(), nuevo), []).append(nro)
                        v = nuevo
                    elif nuevo is None and str(v).strip() not in ids_cuenta:
                        sin_id.setdefault((nom, str(v).strip()), []).append(nro)
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
                if r.get("bill_interno") and nom == r["bill_interno"] and texto.startswith("VENDB_"):
                    res.problemas.append(Problema("error", nro, nom, "BILL todavía es el External ID de la factura: cargue el export de facturas de NetSuite "
                                                  "(🔗 Cargar ID de facturas de NetSuite) para poner su Internal ID"))
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
            elif salida[col_id]:
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
        if r.get("ruc_o_dni") and salida[idx(r["ruc_o_dni"])]:
            c, v = r["ruc_o_dni"], salida[idx(r["ruc_o_dni"])]
            if v.isdigit() and len(v) == 11:
                if not _ruc_valido(v):
                    res.problemas.append(Problema("aviso", nro, c, f"El RUC {v} no pasa la verificación del dígito final; revíselo"))
            elif not (v.isdigit() and len(v) == 8):
                res.problemas.append(Problema("error", nro, c, "Debe ser un RUC de 11 dígitos o un DNI de 8"))
        if r.get("cuadre_total"):
            try:
                a, b, t = (_num(salida[idx(n)]) or Decimal(0) for n in r["cuadre_total"])
                if abs(a + b - t) > Decimal("0.02"):
                    res.problemas.append(Problema("error", nro, r["cuadre_total"][2],
                                                  f"{r['cuadre_total'][0]} + {r['cuadre_total'][1]} ({_fmt_num(a + b)}) no es igual a {r['cuadre_total'][2]} ({_fmt_num(t)})"))
            except (ValueError, InvalidOperation):
                pass
        if r.get("pagado"):  # Monto Pagado = Total − Monto Detr. (Sub-Total + IGV = Total, menos lo depositado en detracciones)
            try:
                pag, tot, det = r["pagado"]
                total_p = _num(salida[idx(tot)])
                salida[idx(pag)] = _fmt_num(total_p - (_num(salida[idx(det)]) or Decimal(0))) if total_p is not None else ""
            except (ValueError, InvalidOperation):
                pass
        if r.get("detraccion"):  # detracción SUNAT: porcentaje según el concepto y Monto Detr. = Total × porcentaje
            dd = r["detraccion"]
            sujeto = norm(salida[idx(dd["sujeto"])]) in ("sí", "si", "s", "1", "true", "t")
            concepto = salida[idx(dd["concepto"])]
            pct = esquema.DETRACCIONES.get(concepto)
            salida[idx(dd["porcentaje"])] = f"{pct * 100:g}%" if pct else ""
            if sujeto:
                try:
                    monto, total, igv = (_num(salida[idx(dd[k])]) or Decimal(0) for k in ("monto", "total", "igv"))
                except (ValueError, InvalidOperation):
                    monto = total = igv = Decimal(0)
                if not concepto or not pct:
                    res.problemas.append(Problema("error", nro, dd["concepto"], "Sujeto a detracción: indique el concepto de detracción"))
                else:
                    esperado = (total * Decimal(str(pct))).quantize(Decimal("0.01"))
                    if monto == 0:
                        salida[idx(dd["monto"])] = _fmt_num(esperado)
                    elif abs(monto - esperado) > Decimal("0.01"):
                        res.problemas.append(Problema("aviso", nro, dd["monto"], f"Monto Detr. ({_fmt_num(monto)}) difiere del cálculo {pct * 100:g}% × {_fmt_num(total)} = {_fmt_num(esperado)}"))
                codigo_det = concepto.split(" ", 1)[0]
                if not salida[idx(dd["tipo"])] and codigo_det in esquema.CODIGOS_SERVICIO | esquema.CODIGOS_EXONERADO:   # tipo de operación de la tabla de SUNAT
                    inicio = "Venta de bienes exonerada" if codigo_det in esquema.CODIGOS_EXONERADO else "Venta de bienes muebles"
                    salida[idx(dd["tipo"])] = next((o for o in esquema.LISTAS_EXTRA["Gastos"].get("Tipo Operación SUNAT", []) if o.startswith(inicio)), "")
                if codigo_det in esquema.CODIGOS_EXONERADO and esquema.norm(salida[idx("Tipo Doc")]).startswith("liquidaci"):
                    res.problemas.append(Problema("aviso", nro, dd["sujeto"], "Una liquidación de compra de café no lleva la detracción del 1.5%: revise si debe estar sujeta"))
                if igv == 0 and codigo_det not in esquema.CODIGOS_EXONERADO:
                    res.problemas.append(Problema("aviso", nro, dd["sujeto"], "La detracción aplica a operaciones afectas al IGV, y este comprobante tiene IGV 0"))
            else:
                try:
                    if (_num(salida[idx(dd["monto"])]) or 0) > 0:
                        res.problemas.append(Problema("aviso", nro, dd["sujeto"], "Tiene Monto Detr. pero no está marcado como sujeto a detracción"))
                except (ValueError, InvalidOperation):
                    pass
        if r.get("detraccion"):  # importe neto a pagar al proveedor = Total − detracción (si el comprobante está sujeto)
            dd = r["detraccion"]
            try:
                total_n = _num(salida[idx(dd["total"])])
                monto_n = (_num(salida[idx(dd["monto"])]) or Decimal(0)) if norm(salida[idx(dd["sujeto"])]) in ("sí", "si", "s", "1", "true", "t") else Decimal(0)
                salida[idx(dd["neto"])] = _fmt_num(total_n - monto_n) if total_n is not None else ""
            except (ValueError, InvalidOperation):
                salida[idx(dd["neto"])] = ""
        if r.get("metodo_pago"):  # Met. Pago vacío = según la Caja (cuenta bancaria -> transferencia; Caja Chica -> efectivo)
            c_caja, c_met = r["metodo_pago"]
            if not salida[idx(c_met)]:
                salida[idx(c_met)] = gastos_cuentas.metodo_pago(salida[idx(c_caja)])
                if not salida[idx(c_met)] and salida[idx(c_caja)]:
                    res.problemas.append(Problema("aviso", nro, c_met, f"No se pudo deducir el método de pago de la caja «{salida[idx(c_caja)][:40]}»"))
        if r.get("cuenta_pagar"):  # cuenta por pagar: PCGE según el tipo de documento (si está vacía) y su número en el COA de NetSuite
            cp = r["cuenta_pagar"]
            if not salida[idx(cp["pcge"])]:
                salida[idx(cp["pcge"])] = esquema.cuenta_pagar_por_tipo(salida[idx(cp["tipo"])])
            pcge_p = homologacion.normalizar(salida[idx(cp["pcge"])])
            if pcge_p and homologacion.mapa().get(pcge_p) is None:
                res.problemas.append(Problema("aviso", nro, cp["pcge"], f"La cuenta por pagar {pcge_p} no está en «Cuentas homologadas»; agréguela con su cuenta de NetSuite"))
                salida[idx(cp["ns"])] = ""
            else:
                salida[idx(cp["ns"])] = (homologacion.coa().get(pcge_p) or str(homologacion.mapa().get(pcge_p, ""))) if pcge_p else ""
        if r.get("cuenta_gasto"):  # Cuenta NetSuite = ID interno de la cuenta PCGE elegida (según «Cuentas homologadas»)
            c_pcge, c_ns = r["cuenta_gasto"]
            pcge = homologacion.normalizar(salida[idx(c_pcge)])
            if not pcge:
                res.problemas.append(Problema("aviso", nro, c_pcge, "Sin cuenta: use «Proponer cuentas» o elíjala"))
                salida[idx(c_ns)] = ""
            elif homologacion.mapa().get(pcge) is None:
                res.problemas.append(Problema("aviso", nro, c_pcge, f"La cuenta {pcge} no está en «Cuentas homologadas»"))
                salida[idx(c_ns)] = ""
            else:   # Cuenta NetSuite = número de la cuenta en el COA (p. ej. 9011.06); el ID interno se usa al pasar a Factura de Proveedor
                salida[idx(c_ns)] = homologacion.coa().get(pcge) or str(homologacion.mapa()[pcge])
        for c in r.get("avisar_vacios", []):
            if not salida[idx(c)]:
                res.problemas.append(Problema("aviso", nro, c, "Sin dato"))
        if r.get("unico_compuesto"):
            clave = tuple(salida[idx(c)].upper() for c in r["unico_compuesto"])
            if all(clave):
                if clave in claves_vistas:
                    res.problemas.append(Problema("error", nro, " + ".join(r["unico_compuesto"]),
                                                  f"Comprobante repetido: ya está en la fila {claves_vistas[clave]}"))
                else:
                    claves_vistas[clave] = nro
        if etiq_periodo:
            if r.get("periodo") and salida[idx(r["periodo"])] and norm(salida[idx(r["periodo"])]) != norm(etiq_periodo):
                res.problemas.append(Problema("aviso", nro, r["periodo"], f"El período contable ({salida[idx(r['periodo'])]}) no coincide con el período seleccionado ({etiq_periodo})"))
            if r.get("fecha_periodo") and fila[idx(r["fecha_periodo"])] not in (None, ""):
                try:
                    f_doc = _fecha(fila[idx(r["fecha_periodo"])])
                    if f"{f_doc.year:04d}-{f_doc.month:02d}" != periodo:
                        res.problemas.append(Problema("aviso", nro, r["fecha_periodo"], f"La fecha {f_doc.strftime('%d/%m/%Y')} no pertenece al período {etiq_periodo}"))
                except ValueError:
                    pass  # ya se informa como error de fecha
        doc = r.get("documento")
        if doc and salida[idx(doc["numero"])]:
            tipo, num = salida[idx(doc["tipo"])], salida[idx(doc["numero"])]
            if proveedores.es_ruc(tipo):
                if not (num.isdigit() and len(num) == 11):
                    res.problemas.append(Problema("error", nro, doc["numero"], "El RUC debe tener 11 dígitos"))
                elif not _ruc_valido(num):
                    res.problemas.append(Problema("aviso", nro, doc["numero"], f"El RUC {num} no pasa la verificación del dígito final; revíselo"))
            elif proveedores.es_dni(tipo) and not (num.isdigit() and len(num) == 8):
                res.problemas.append(Problema("error", nro, doc["numero"], "El DNI debe tener 8 dígitos"))
        for texto_aviso in avisos_completar.get(nro, []):
            col_aviso, texto_aviso = texto_aviso if isinstance(texto_aviso, tuple) else ("Código ZIP", texto_aviso)
            res.problemas.append(Problema("aviso", nro, col_aviso, texto_aviso))
        for c in r.get("cinco_digitos", []):
            v = salida[idx(c)]
            if v and not (v.isdigit() and len(v) == 5):
                res.problemas.append(Problema("aviso", nro, c, "El código postal SERPOST tiene 5 dígitos"))
        if r.get("unico_id") and salida[col_id]:
            if salida[col_id] in ids_vistos:
                res.problemas.append(Problema("error", nro, hoja.encabezados[col_id].strip(),
                                              f"Repetido: {salida[col_id]} ya está en la fila {ids_vistos[salida[col_id]]}"))
            else:
                ids_vistos[salida[col_id]] = nro
        res.filas.append(salida)
        res.numeros_fila.append(nro)

    for (col, origen), nros in sin_id.items():
        res.problemas.append(Problema("aviso", nros[0], col, f"La cuenta {origen} no está en «Cuentas homologadas» (no tiene ID interno de NetSuite): se envía tal cual y NetSuite puede rechazarla ({len(nros)} fila{'s' if len(nros) > 1 else ''})"))
    for (col, origen, destino), nros in traducidas.items():
        res.problemas.append(Problema("aviso", nros[0], col,
                                      f"Cuenta {origen} → {destino} (homologada, {len(nros)} fila{'s' if len(nros) > 1 else ''})"))
    for ext, saldo in saldos.items():
        if saldo != 0:
            res.problemas.append(Problema("error", None, ext, f"Asiento descuadrado: Débito − Crédito = {_fmt_num(saldo)}"))
    return res


def procesar(datos: dict[str, list], plantilla: Plantilla, formato_fecha: str, periodo: str | None = None) -> list[ResultadoHoja]:
    resultados = []
    for k, filas in datos.items():
        res = validar_hoja(plantilla.hojas[k], filas, formato_fecha, periodo)
        if plantilla.hojas[k].reglas.get("completar") == "proveedores" and res.filas:
            # el CSV de proveedores sale con las columnas del formulario de NetSuite (las validaciones siguen sobre las de la pestaña)
            res.filas, cuantas = proveedores.a_salida(res.encabezados, res.filas)
            if len(cuantas) == len(res.numeros_fila):       # una fila por moneda: el número de fila de la pestaña se repite
                res.numeros_fila = [n for n, c in zip(res.numeros_fila, cuantas) for _ in range(c)]
            res.encabezados_origen, res.encabezados = res.encabezados, list(proveedores.SALIDA)
        resultados.append(res)
    return resultados


def reporte_errores_xlsx(resultados: list[ResultadoHoja], periodo_etiqueta: str = "") -> bytes:
    """Excel con todos los errores y avisos de la validación: una fila por problema, con el número de fila de la pestaña
    (el «#» de la grilla), el identificador (primera columna), la columna, el valor que tiene y el mensaje."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill

    wb = Workbook()
    ws_r = wb.active
    ws_r.title = "Resumen"
    ws_r.append(["Plantilla", "Filas", "Errores", "Avisos", "¿Se generó el CSV?"])
    ws_d = wb.create_sheet("Errores y avisos")
    ws_d.append(["Plantilla", "Fila (#)", "Identificador", "Columna", "Valor actual", "Nivel", "Mensaje"])
    for r in resultados:
        ws_r.append([r.nombre.strip(), len(r.filas), r.errores, r.avisos, "Sí" if r.exportable else ("No: tiene errores" if r.errores else "No")])
        por_fila = dict(zip(r.numeros_fila, r.filas))
        for p in sorted(r.problemas, key=lambda p: (p.nivel != "error", p.fila or 0, p.columna)):
            fila = por_fila.get(p.fila or -1)
            col = r.encabezados.index(p.columna) if p.columna in r.encabezados else -1
            ws_d.append([r.nombre.strip(), p.fila, (fila[0] if fila else "") or "", p.columna,
                         (fila[col] if fila and col >= 0 else "") or "", "Error" if p.nivel == "error" else "Aviso", p.mensaje])
    if periodo_etiqueta:
        ws_r.append([]); ws_r.append([f"Período {periodo_etiqueta}"])
    for ws, anchos in ((ws_r, (28, 10, 10, 10, 24)), (ws_d, (24, 10, 40, 34, 30, 9, 90))):
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="1E40AF"); c.alignment = Alignment(vertical="center")
        for i, w in enumerate(anchos):
            ws.column_dimensions[chr(65 + i)].width = w
        ws.freeze_panes = "A2"
    ws_d.auto_filter.ref = ws_d.dimensions
    for fila in ws_d.iter_rows(min_row=2):
        if fila[5].value == "Error":
            fila[5].font = Font(bold=True, color="B91C1C")
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def empaquetar_zip(resultados: list[ResultadoHoja]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for r in resultados:
            if r.exportable:
                z.writestr(r.archivo, r.csv_bytes())
    return buf.getvalue()
