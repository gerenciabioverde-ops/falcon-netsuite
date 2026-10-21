"""External ID de los comprobantes de compra: prefijo + RUC + Serie + Número, todo junto y de largo fijo."""
from __future__ import annotations

import re

LARGO_DOC, LARGO_SERIE, LARGO_NUMERO = 11, 4, 8       # VENDB_ + RUC (11) + Serie (4) + Número (8) = 29 caracteres


def _limpio(t) -> str:
    return re.sub(r"[^A-Za-z0-9]+", "", str(t or ""))


def id_comprobante(prefijo: str, doc, serie, numero=None) -> str:
    """VENDB_20513497769FA0100018047: el documento (un DNI de 8 dígitos también) se completa con ceros a la izquierda hasta 11, la serie hasta 4 y el número hasta 8.
    Se puede pasar «Serie-Número» en un solo texto (serie) y sin número; sin serie (planilla de movilidad) la serie es 0000.
    Devuelve '' si falta el documento o el número."""
    if numero is None:
        sn = str(serie or "").strip()
        serie, numero = sn.split("-", 1) if "-" in sn else ("", sn)
    doc, serie, numero = _limpio(doc), _limpio(serie), _limpio(numero)
    if not doc or not numero:
        return ""
    return f"{prefijo}_{doc.rjust(LARGO_DOC, '0')}{serie.rjust(LARGO_SERIE, '0')}{numero.rjust(LARGO_NUMERO, '0')}"
