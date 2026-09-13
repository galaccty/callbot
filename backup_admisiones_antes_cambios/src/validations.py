"""Validations: detección y reporte de errores en los datos cargados.

Objetivo: NUNCA borrar datos silenciosamente. Cada anomalía se reporta en
pantalla con localización (fila) y detalle, y el usuario decide cómo seguir.
"""

from __future__ import annotations

from typing import Dict, List

import pandas as pd


def _empty(value: object) -> bool:
    """¿El valor se considera vacío?"""
    if value is None:
        return True
    if isinstance(value, float) and pd.isna(value):
        return True
    s = str(value).strip()
    return s == "" or s.lower() in {"nan", "none", "nat"}


def validate(
    data: pd.DataFrame,
    expected_cols: List[str],
) -> Dict[str, List[str]]:
    """Recorre el DataFrame y acumula errores legibles por categoría.

    Devuelve un dict con las claves:
      - duplicados_admision: valores de 'Admision' repetidos (con filas).
      - duplicados_fila:     filas idénticas según el id anti-duplicados.
      - admision_vacia:      filas sin 'Admision'.
      - usuario_vacio:       filas sin 'Usuario' (no es error duro, se reporta).
      - filas_incompletas:   filas con campos críticos vacíos.
      - tipos_invalidos:     valores que no coinciden con el tipo esperado.

    Las filas se reportan con su índice (0 = primera fila de datos tras el
    encabezado) para que el usuario las ubique en el Excel original.
    """
    errors: Dict[str, List[str]] = {
        "duplicados_admision": [],
        "duplicados_fila": [],
        "admision_vacia": [],
        "usuario_vacio": [],
        "filas_incompletas": [],
        "tipos_invalidos": [],
    }

    if data.empty:
        return errors

    # 1) Admisiones duplicadas (reportar todas las filas implicadas)
    if "Admision" in data.columns:
        adm = data["Admision"].fillna("")
        dup_mask = adm.duplicated(keep=False)
        seen: set = set()
        for i, (val, is_dup) in enumerate(zip(adm, dup_mask)):
            if not is_dup or not str(val).strip():
                continue
            if val not in seen:
                seen.add(val)
                filas = [j for j, v in enumerate(adm) if v == val]
                errors["duplicados_admision"].append(
                    f"Admision '{val}' aparece {len(filas)} veces (filas dato {filas})"
                )

    # 2) Filas duplicadas completas (según id anti-duplicados)
    if "_id" in data.columns:
        ids = data["_id"].fillna("")
        seen_ids: set = set()
        for i, (val, is_dup) in enumerate(zip(ids, ids.duplicated(keep=False))):
            if not is_dup or not str(val).strip() or val in seen_ids:
                continue
            seen_ids.add(val)
            filas = [j for j, v in enumerate(ids) if v == val]
            errors["duplicados_fila"].append(
                f"Fila duplicada {len(filas)} veces (filas dato {filas})"
            )

    # 3) Admision vacío
    if "Admision" in data.columns:
        for i in data.index:
            if _empty(data.at[i, "Admision"]):
                errors["admision_vacia"].append(f"Fila dato {i}: 'Admision' vacío")

    # 4) Usuario vacío (informativo, no bloquea)
    if "Usuario" in data.columns:
        for i in data.index:
            if _empty(data.at[i, "Usuario"]):
                errors["usuario_vacio"].append(
                    f"Fila dato {i}: 'Usuario' vacío -> se clasifica como "
                    "'Sin usuario identificado'"
                )

    # 5) Filas incompletas (críticos: Nombre, Contrato, Agencia)
    criticos = [c for c in ["Agencia", "Nombre", "Contrato"] if c in data.columns]
    if criticos:
        for i in data.index:
            faltantes = [c for c in criticos if _empty(data.at[i, c])]
            if faltantes:
                errors["filas_incompletas"].append(
                    f"Fila dato {i}: campos vacíos {faltantes}"
                )

    # 6) Tipos inválidos: 'Saldo' debe ser numérico; 'Fecha' no vacía y parseable
    if "Saldo" in data.columns:
        for i in data.index:
            v = data.at[i, "Saldo"]
            if _empty(v):
                continue
            try:
                float(str(v).replace(",", "."))
            except (ValueError, TypeError):
                errors["tipos_invalidos"].append(
                    f"Fila dato {i}: 'Saldo' no numérico ({v!r})"
                )
    if "Fecha" in data.columns:
        for i in data.index:
            f = data.at[i, "Fecha"]
            if _empty(f):
                errors["tipos_invalidos"].append(f"Fila dato {i}: 'Fecha' vacía")

    return errors


def summarize(errors: Dict[str, List[str]]) -> str:
    """Convierte el dict de errores en un resumen legible para pantalla."""
    total = sum(len(v) for v in errors.values())
    partes = [f"Se encontraron {total} registros con errores."]
    etiquetas = {
        "duplicados_admision": "Admisiones duplicadas",
        "duplicados_fila": "Filas duplicadas",
        "admision_vacia": "Admisiones vacías",
        "usuario_vacio": "Usuarios vacíos (informativo)",
        "filas_incompletas": "Filas incompletas",
        "tipos_invalidos": "Tipos de dato inesperados",
    }
    for clave, lista in errors.items():
        if lista:
            partes.append(f"\n• {etiquetas[clave]} ({len(lista)}):")
            for detalle in lista[:15]:
                partes.append(f"    - {detalle}")
            if len(lista) > 15:
                partes.append(f"    - ... y {len(lista) - 15} más")
    return "\n".join(partes)
