"""Padrón reducido de RUC de SUNAT: trae nombre, estado, condición y dirección de un RUC para completar proveedores.

SUNAT publica un ZIP (padron_reducido_ruc.zip, ~13 millones de filas, TXT separado por «|» en ISO-8859-1). No se guarda
completo: solo los RUC que este sistema usa (los de Gastos, Proveedores y Bancos de todos los períodos, más los que se
consultaron y no estaban) en datos_usuario/padron_sunat.db (SQLite). El TXT se lee al vuelo, sin descomprimir a disco.

Columnas del TXT: RUC | NOMBRE O RAZÓN SOCIAL | ESTADO DEL CONTRIBUYENTE | CONDICIÓN DE DOMICILIO | UBIGEO | TIPO DE VÍA |
NOMBRE DE VÍA | CÓDIGO DE ZONA | TIPO DE ZONA | NÚMERO | INTERIOR | LOTE | DEPARTAMENTO | MANZANA | KILÓMETRO.
El UBIGEO es el código de 6 dígitos; data/ubigeo_sunat.json lo traduce a «DEPARTAMENTO-PROVINCIA-DISTRITO» (el formato de
la lista de la plantilla; fuente de los códigos: paquete ubigeos-peru, licencia MIT).
"""
from __future__ import annotations

import json
import os
import re
import sqlite3
import tempfile
import threading
import time
import unicodedata
import urllib.request
import zipfile
from functools import lru_cache
from pathlib import Path
from typing import Iterator

URL_PADRON = "http://www2.sunat.gob.pe/padron_reducido_ruc.zip"
RAIZ = Path(__file__).resolve().parent.parent
RUTA_DB = RAIZ / "datos_usuario" / "padron_sunat.db"
RUTA_UBIGEOS = RAIZ / "data" / "ubigeo_sunat.json"
DIR_PERIODOS = RAIZ / "datos_usuario" / "periodos"

_RUC = re.compile(r"^(10|15|17|20)\d{9}$")
_CAB = {"razon": "nombre o razon social", "estado": "estado del contribuyente", "condicion": "condicion de domicilio",
        "ubigeo": "ubigeo", "tipo_via": "tipo de via", "nombre_via": "nombre de via", "cod_zona": "codigo de zona",
        "tipo_zona": "tipo de zona", "numero": "numero", "interior": "interior", "lote": "lote", "departamento": "departamento", "manzana": "manzana",
        "km": "kilometro"}
_POS_DEFECTO = {"ruc": 0, "razon": 1, "estado": 2, "condicion": 3, "ubigeo": 4, "tipo_via": 5, "nombre_via": 6, "cod_zona": 7,
                "tipo_zona": 8, "numero": 9, "interior": 10, "lote": 11, "departamento": 12, "manzana": 13, "km": 14}

_estado: dict = {"fase": "", "activo": False, "error": "", "leidas": 0, "encontrados": 0, "bytes": 0}
_candado = threading.Lock()


class PadronError(Exception):
    """No se pudo descargar o leer el padrón."""


def _sin_tildes(s: str) -> str:
    return " ".join("".join(c for c in unicodedata.normalize("NFKD", s or "") if not unicodedata.combining(c)).lower().split())


def es_ruc(valor) -> bool:
    return bool(_RUC.match(str(valor or "").strip()))


@lru_cache(maxsize=1)
def _ubigeos() -> dict[str, str]:
    try:
        return json.loads(RUTA_UBIGEOS.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def _conexion() -> sqlite3.Connection:
    RUTA_DB.parent.mkdir(parents=True, exist_ok=True)
    c = sqlite3.connect(RUTA_DB, timeout=30)
    c.execute("CREATE TABLE IF NOT EXISTS padron (ruc TEXT PRIMARY KEY, razon TEXT, estado TEXT, condicion TEXT, ubigeo TEXT, direccion TEXT)")
    c.execute("CREATE TABLE IF NOT EXISTS pedidos (ruc TEXT PRIMARY KEY)")
    c.execute("CREATE TABLE IF NOT EXISTS meta (clave TEXT PRIMARY KEY, valor TEXT)")
    return c


# ---------- consulta ----------

def formatear_razon(ruc: str, razon: str) -> str:
    """SUNAT escribe todo en mayúsculas. Una persona (RUC 10/15/17) se pone con la inicial en mayúscula, como los demás
    proveedores personas del sistema (Torres Nuñez Sadith); las empresas se dejan como las trae SUNAT."""
    if str(ruc)[:2] in ("10", "15", "17") and razon.isupper():
        return re.sub(r"[A-ZÁÉÍÓÚÑÜ]+", lambda m: m.group(0).capitalize(), razon)
    return razon


_ETIQUETAS_DIR = {"Nro": "Nro.", "Dpto": "Dpto.", "Int": "Int.", "Lote": "Lt.", "Mza": "Mz.", "Km": "Km."}


def formatear_direccion(texto: str) -> str:
    """Dirección como las que ya tiene el sistema (Cal. Porta Nro. 170 Dpto. 610): inicial mayúscula y Nro./Mz./Lt./Km. con punto."""
    t = re.sub(r"[A-ZÁÉÍÓÚÑÜ]+", lambda m: m.group(0).capitalize(), texto or "")
    return " ".join(_ETIQUETAS_DIR.get(w, w) for w in t.split())


def buscar(ruc: str) -> dict | None:
    """Datos del RUC en el padrón cargado, o None. Si no está, queda «pedido» para la próxima actualización."""
    ruc = str(ruc or "").strip()
    if not es_ruc(ruc):
        return None
    with _conexion() as c:
        f = c.execute("SELECT ruc, razon, estado, condicion, ubigeo, direccion FROM padron WHERE ruc = ?", (ruc,)).fetchone()
        if f is None:
            c.execute("INSERT OR IGNORE INTO pedidos (ruc) VALUES (?)", (ruc,))
            return None
    return {"ruc": f[0], "razon": formatear_razon(f[0], f[1] or ""), "estado": f[2] or "", "condicion": f[3] or "", "ubigeo_codigo": f[4] or "",
            "ubigeo": _ubigeos().get(f[4] or "", ""), "direccion": formatear_direccion(f[5] or "")}


def estado() -> dict:
    with _conexion() as c:
        n = c.execute("SELECT COUNT(*) FROM padron").fetchone()[0]
        pend = c.execute("SELECT COUNT(*) FROM pedidos WHERE ruc NOT IN (SELECT ruc FROM padron)").fetchone()[0]
        act = c.execute("SELECT valor FROM meta WHERE clave = 'actualizado'").fetchone()
    return {"cargados": n, "pendientes": pend, "actualizado": act[0] if act else "", "proceso": dict(_estado)}


# ---------- RUC de interés ----------

def _rucs_de_los_periodos() -> set[str]:
    from . import esquema, periodos, proveedores
    iG, iB, iP = (esquema.GASTOS_ENCABEZADOS.index("RUC"), esquema.BANCOS_MN_ENCABEZADOS.index("Número de RUC"),
                  proveedores.ENCABEZADOS.index("Latam - Tax Reg Number"))
    salida: set[str] = set()
    for ruta in DIR_PERIODOS.glob("*.json"):
        hojas = periodos.leer(ruta.stem).get("hojas") or {}
        for clave, i in (("gastos", iG), ("bancos mn", iB), ("proveedores", iP)):
            for f in (hojas.get(clave) or {}).get("filas") or []:
                if len(f) > i and es_ruc(f[i]):
                    salida.add(str(f[i]).strip())
    return salida


def rucs_de_interes() -> set[str]:
    with _conexion() as c:
        pedidos = {r[0] for r in c.execute("SELECT ruc FROM pedidos")}
        cargados = {r[0] for r in c.execute("SELECT ruc FROM padron")}
    return _rucs_de_los_periodos() | pedidos | cargados


# ---------- lectura del archivo ----------

def _posiciones(linea: str):
    cols = [_sin_tildes(c) for c in linea.split("|")]
    if not cols or cols[0] != "ruc":
        return None
    pos = {"ruc": 0}
    for k, nombre in _CAB.items():
        if nombre in cols:
            pos[k] = cols.index(nombre)
    return pos if "razon" in pos else None


def _lineas(ruta: str) -> Iterator[str]:
    import io
    if zipfile.is_zipfile(ruta):
        with zipfile.ZipFile(ruta) as z:
            nombres = [n for n in z.namelist() if n.lower().endswith(".txt")]
            if not nombres:
                raise PadronError("El ZIP no contiene ningún archivo .txt")
            with z.open(nombres[0]) as f:
                for crudo in io.TextIOWrapper(f, encoding="latin-1", newline=""):
                    yield crudo.rstrip("\r\n")
    else:
        with open(ruta, "r", encoding="latin-1", newline="") as f:
            for crudo in f:
                yield crudo.rstrip("\r\n")


def _v(campos: list[str], pos: dict, clave: str) -> str:
    i = pos.get(clave)
    if i is None or i >= len(campos):
        return ""
    v = campos[i].strip()
    return "" if v in ("-", "--") else v


def _direccion(campos: list[str], pos: dict) -> str:
    partes = [_v(campos, pos, "tipo_via"), _v(campos, pos, "nombre_via")]
    for etiqueta, clave in (("NRO", "numero"), ("DPTO", "departamento"), ("INT", "interior"), ("LOTE", "lote"), ("MZA", "manzana"), ("KM", "km")):
        valor = _v(campos, pos, clave)
        if valor:
            partes.append(f"{etiqueta} {valor}")
    zona = " ".join(p for p in (_v(campos, pos, "tipo_zona"), _v(campos, pos, "cod_zona")) if p)
    if zona:
        partes.append(zona)
    return " ".join(p for p in partes if p)


def cargar_archivo(ruta: str, rucs: set[str]) -> int:
    """Lee el padrón (ZIP o TXT) y guarda solo los RUC pedidos. Devuelve cuántos encontró."""
    pos = dict(_POS_DEFECTO)
    lote: list[tuple] = []
    encontrados = 0
    with _conexion() as c:
        def volcar():
            if lote:
                c.executemany("INSERT OR REPLACE INTO padron (ruc, razon, estado, condicion, ubigeo, direccion) VALUES (?,?,?,?,?,?)", lote)
                lote.clear()
        for n, linea in enumerate(_lineas(ruta), start=1):
            if n == 1:
                p = _posiciones(linea)
                if p:
                    pos = p
                    continue
            _estado["leidas"] = n
            corte = linea.find("|")
            if corte < 0 or linea[:corte] not in rucs:
                continue
            campos = linea.split("|")
            lote.append((campos[0], _v(campos, pos, "razon"), _v(campos, pos, "estado"), _v(campos, pos, "condicion"),
                         _v(campos, pos, "ubigeo"), _direccion(campos, pos)))
            encontrados += 1
            _estado["encontrados"] = encontrados
            if len(lote) >= 500:
                volcar()
        volcar()
        c.execute("DELETE FROM pedidos WHERE ruc IN (SELECT ruc FROM padron)")
        c.execute("INSERT OR REPLACE INTO meta (clave, valor) VALUES ('actualizado', ?)", (time.strftime("%d/%m/%Y %H:%M"),))
    return encontrados


def _descargar(destino: str) -> None:
    try:
        req = urllib.request.Request(URL_PADRON, headers={"User-Agent": "Mozilla/5.0 FalconNetsuite"})
        with urllib.request.urlopen(req, timeout=60) as r, open(destino, "wb") as f:
            while True:
                trozo = r.read(1024 * 1024)
                if not trozo:
                    break
                f.write(trozo)
                _estado["bytes"] += len(trozo)
    except OSError as e:
        raise PadronError(f"No se pudo descargar el padrón de SUNAT ({e}). Descárguelo desde {URL_PADRON} y súbalo con «Subir ZIP».") from e
    if not zipfile.is_zipfile(destino):
        raise PadronError("Lo descargado no es un ZIP válido (SUNAT pudo bloquear la descarga). Descárguelo a mano y súbalo con «Subir ZIP».")


def _trabajo(ruta_subida: str | None) -> None:
    ruta, descargado = ruta_subida, False
    try:
        rucs = rucs_de_interes()
        if not rucs:
            raise PadronError("No hay RUC en Gastos, Proveedores ni Bancos para buscar en el padrón.")
        if not ruta:
            _estado["fase"] = "Descargando el padrón de SUNAT…"
            fd, ruta = tempfile.mkstemp(suffix=".zip", prefix="padron_")
            os.close(fd)
            descargado = True
            _descargar(ruta)
        _estado["fase"] = f"Buscando {len(rucs)} RUC en el padrón…"
        cargar_archivo(ruta, rucs)
        _estado["fase"] = "Listo"
    except (PadronError, OSError, zipfile.BadZipFile, sqlite3.Error) as e:
        _estado["error"] = str(e)
        _estado["fase"] = "Error"
    finally:
        _estado["activo"] = False
        if ruta and (descargado or ruta_subida):       # el ZIP descargado o subido es temporal
            try:
                os.unlink(ruta)
            except OSError:
                pass


def iniciar_actualizacion(ruta_subida: str | None = None) -> bool:
    """Arranca la descarga/lectura en segundo plano; False si ya hay una en curso."""
    with _candado:
        if _estado["activo"]:
            return False
        _estado.update({"fase": "Iniciando…", "activo": True, "error": "", "leidas": 0, "encontrados": 0, "bytes": 0})
    threading.Thread(target=_trabajo, args=(ruta_subida,), daemon=True).start()
    return True
