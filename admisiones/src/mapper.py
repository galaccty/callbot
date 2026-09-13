"""Mapper: normaliza el DataFrame fuente al formato interno de trabajo.

Convierte encabezados, limpia espacios, maneja valores nulos y estandariza
campos textuales (Usuario, Estado_Entrega) para que el resto del pipeline
trabaje siempre con nombres y valores conocidos.
"""

from __future__ import annotations

import hashlib
from typing import List, Tuple

import pandas as pd


class MappingError(Exception):
    """Levantada cuando los encabezados del archivo no coinciden con lo esperado."""


def _norm(value: object) -> str:
    """Estandariza un valor a texto limpio (strip, NaN -> '')."""
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    val = str(value).strip()
    if val.lower() in {"nan", "none", "nat"}:
        return ""
    return val


def validate_headers(
    data: pd.DataFrame, expected: List[str]
) -> Tuple[List[str], List[str]]:
    """Compara los encabezados presentes vs. los esperados.

    Devuelve (encabezados_faltantes, encabezados_desconocidos). No lanza error:
    es el llamador quien decide cómo tratar las discrepancias.
    """
    actual = [c.strip() for c in map(str, data.columns) if str(c).strip()]
    expected_set = {str(e).strip() for e in expected}
    actual_set = set(actual)

    missing = [e for e in expected if e not in actual_set]
    unknown = [a for a in actual if a not in expected_set]
    return missing, unknown


def build_id(row: pd.Series) -> str:
    """Construye un identificador estable anti-duplicados a partir de campos clave.

    Sirve para detectar filas duplicadas aunque el 'Admision' venga vacío,
    sin inventar datos: se hace un hash de los campos que identifican la fila.
    """
    base = "|".join(
        _norm(row.get(c))
        for c in ["Agencia", "Admision", "Contrato", "Nombre", "Fecha"]
    )
    if not base.strip("|"):
        return ""
    return hashlib.md5(base.encode("utf-8")).hexdigest()


def apply_mapping(data: pd.DataFrame, expected: List[str]) -> pd.DataFrame:
    """Renombra y normaliza el DataFrame según el mapeo esperado.

    Pasos:
      1. Limpia nombres de columnas (strip).
      2. Reordena al orden esperado, ignorando columnas extra desconocidas.
      3. Convierte el encabezado numerico '#' en columna '#' (indice de origen).
      4. Aplica normalizacion de texto a las columnas clave.
      5. Clasifica 'Usuario vacio' como 'Sin usuario identificado'.
      6. Asigna el id anti-duplicados.
    """
    frame = data.copy()

    # Normalizar nombres de columnas
    frame.columns = [str(c).strip() for c in frame.columns]

    # Conservar solo / reordenar según order esperado (más las presentes).
    cols_esperadas = [str(e).strip() for e in expected]
    presentes = [c for c in cols_esperadas if c in frame.columns]
    extra = [c for c in frame.columns if c not in presentes]
    frame = frame[presentes + extra]

    # Normalización de texto en columnas clave
    for col in ["Admision", "Afiliado", "Nombre", "Contrato",
                "Eps", "Estado", "Usuario"]:
        if col in frame.columns:
            frame[col] = frame[col].map(_norm)

    # Agencia: normalizar y dejar con 3 dígitos (p.ej. 17 -> '017') para
    # coincidir con la condición "Agencia: 017" del informe.
    if "Agencia" in frame.columns:
        ag = frame["Agencia"].map(_norm)
        # Quitar el '.0' que deja un float entero (17.0 -> 17).
        ag = ag.map(lambda v: v[:-2] if v.endswith(".0") and v.count(".") == 1 else v)
        frame["Agencia"] = ag.map(
            lambda v: v.zfill(3) if v.isdigit() else v
        )

    # Estado_Entrega: generado por el loader desde el color; se estandariza
    # a capitalización (Entregada / Faltante).
    if "Estado_Entrega" in frame.columns:
        frame["Estado_Entrega"] = frame["Estado_Entrega"].map(
            lambda v: v.lower().capitalize() if v else ""
        )

    # Clasificar usuarios vacíos
    if "Usuario" in frame.columns:
        frame.loc[frame["Usuario"] == "", "Usuario"] = "Sin usuario identificado"

    # Id anti-duplicados
    frame["_id"] = frame.apply(build_id, axis=1)

    return frame
