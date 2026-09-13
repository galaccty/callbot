"""Loader: lectura del Excel fuente y ensamblado del pipeline de carga.

Coordina: lectura con openpyxl (para detectar encabezado y colores) -> mapeo
-> validación. El resultado es un LoadResult con el DataFrame limpio y la
lista de errores detectados (sin descartar ninguna fila).

Particularidad del informe real:
  - El encabezado NO está en la fila 1: hay 2-4 filas de títulos/condiciones.
    Se detecta automáticamente buscando la fila que contiene la MAYOR cantidad
    de columnas esperadas.
  - No existe la columna 'Estado_Entrega' en el origen: se DEDUCE del color
    de fondo de la celda 'Admision' (amarillo = Faltante, sin relleno =
    Entregada). Por eso leemos con openpyxl (a diferencia de pandas, que no
    conserva los estilos) y anexamos ese estado al DataFrame.
  - Hay pie de página ("SIIPS 5.0.0") y filas vacías que deben ignorarse.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import datetime
from typing import List, Optional, Tuple

import pandas as pd
import openpyxl

from .mapper import apply_mapping
from .validations import summarize, validate


@dataclass
class LoadResult:
    """Resultado de cargar y validar un archivo fuente."""
    data: pd.DataFrame
    filename: str
    headers_missing: List[str] = field(default_factory=list)
    headers_unknown: List[str] = field(default_factory=list)
    validation_errors: dict = field(default_factory=dict)
    blocked: bool = False        # True si hay errores que impiden proseguir
    blocking_reasons: List[str] = field(default_factory=list)
    color_warnings: List[str] = field(default_factory=list)

    @property
    def error_summary(self) -> str:
        return summarize(self.validation_errors)


# ── Lectura con openpyxl ─────────────────────────────────────────────
def _norm_header(value: object) -> str:
    """Normaliza el texto de un encabezado (strip, str)."""
    if value is None:
        return ""
    s = str(value).strip()
    if s.lower() in {"nan", "none", "nat"}:
        return ""
    return s


def _find_header_row(ws, expected_cols: List[str], max_scan: int = 15) -> Optional[int]:
    """Localiza la fila del encabezado dentro de las primeras filas.

    Recorre filas 1..max_scan buscando aquella con la MAYOR cantidad de
    columnas esperadas. Devuelve el índice (1-based) o None si no encuentra.
    """
    expected_set = {str(e).strip() for e in expected_cols if str(e).strip()}
    best_row, best_count = None, 0
    for r in range(1, max_scan + 1):
        found = set()
        for c in range(1, ws.max_column + 1):
            h = _norm_header(ws.cell(row=r, column=c).value)
            if h:
                found.add(h)
        matching = len(expected_set & found)
        if matching > best_count:
            best_count = matching
            best_row = r
    return best_row if best_count > 0 else None


def _is_number(value: object) -> bool:
    """¿El valor es numérico (propio del índice '#' del registro)?"""
    if value is None:
        return False
    try:
        float(str(value).strip().replace(",", "."))
        return True
    except (ValueError, TypeError):
        return False


def _cell_fill_rgb(cell) -> Optional[str]:
    """Devuelve el RGB (ARGB 8 hex) del relleno sólido de una celda, o None."""
    fill = cell.fill
    if fill is None:
        return None
    try:
        if fill.patternType != "solid":
            return None
        rgb = fill.fgColor.rgb
        if rgb is None:
            return None
        return str(rgb)
    except Exception:  # noqa: BLE001
        return None


def _parse_sheet(ws, expected_cols):
    """Lee la hoja: localiza encabezado y arma filas + colores de Admision.

    Devuelve (header_row, headers, rows, admin_col, colores).
      headers: dict {titulo_normalizado: indice_columna}.
      rows:    lista de dicts {titulo: valor} por fila de datos.
      colores: lista de RGB del relleno de la celda Admision (por fila).
    """
    header_row = _find_header_row(ws, expected_cols)
    if header_row is None:
        return None, {}, [], None, []

    headers = {}
    admin_col = None
    for c in range(1, ws.max_column + 1):
        h = _norm_header(ws.cell(row=header_row, column=c).value)
        if h:
            headers[h] = c
            if h == "Admision":
                admin_col = c

    rows = []
    colores = []
    for r in range(header_row + 1, ws.max_row + 1):
        first_val = ws.cell(row=r, column=1).value
        admin_val = ws.cell(row=r, column=admin_col).value if admin_col else None
        # Solo filas de datos: el índice '#' (columna 1) debe ser numérico.
        # Esto descarta el pie de página ("SIIPS 5.0.0") y filas vacías.
        if not _is_number(first_val):
            continue
        row = {}
        for name, col in headers.items():
            row[name] = ws.cell(row=r, column=col).value
        rows.append(row)
        colores.append(_cell_fill_rgb(ws.cell(row=r, column=admin_col))
                       if admin_col else None)

    return header_row, headers, rows, admin_col, colores


def _state_from_color(rgb: Optional[str], entregada_hex: str, faltante_hex: str) -> str:
    """Convierte el color de la celda 'Admision' en Estado_Entrega.

    Regla (del informe real):
      - celda amarilla  -> 'Entregada' (las que llegaron)
      - celda blanca/sin relleno -> 'Faltante' (las que faltan)
    """
    entregada_hex = entregada_hex.upper()
    faltante_hex = faltante_hex.upper()
    if rgb is None:
        return "Faltante"
    rgb = rgb.upper()
    if rgb in (entregada_hex, "FFFF00", "FFFFFF00"):
        return "Entregada"
    if rgb in ("FFFFFFFF", "00000000"):
        return "Faltante"
    # color no reconocido -> tratarlo conservadoramente como Faltante
    # (el llamador lo reporta en color_warnings)
    return "Faltante"


def load(
    input_path: str,
    expected_cols: List[str],
    state_colors: dict = None,
) -> LoadResult:
    """Carga y valida un Excel fuente.

    Parámetros:
      input_path:    ruta al archivo .xlsx entrante.
      expected_cols: nombres de columna esperados (del config).
      state_colors:  dict con claves 'entregada' y 'faltante' (hex de la
                     celda Admision). Regla: amarillo=Entregada, blanco=Faltante.

    Devuelve un LoadResult.
    """
    state_colors = state_colors or {}
    entregada_hex = str(state_colors.get("entregada", "FFFFFF00")).upper()
    faltante_hex = str(state_colors.get("faltante", "FFFFFFFF")).upper()

    result = LoadResult(data=pd.DataFrame(), filename=input_path)

    try:
        wb = openpyxl.load_workbook(input_path, data_only=False, read_only=True)
        ws = wb.worksheets[0]
        header_row, headers, rows, admin_col, colores = _parse_sheet(ws, expected_cols)
        wb.close()
    except Exception as exc:  # noqa: BLE001
        result.blocked = True
        result.blocking_reasons.append(
            f"No se pudo leer el archivo {input_path}: {exc}"
        )
        return result

    if header_row is None:
        result.blocked = True
        result.blocking_reasons.append(
            "No se encontró la fila de encabezados (buscando: "
            + ", ".join(expected_cols) + ")"
        )
        return result

    # Encabezados faltantes / desconocidos (Estado_Entrega es derivado)
    expected_set = {str(e).strip() for e in expected_cols}
    actual_set = set(headers.keys())
    missing = [e for e in expected_cols if e not in actual_set and e != "Estado_Entrega"]
    unknown = [h for h in headers if h not in expected_set and h != "Estado_Entrega"]
    result.headers_missing = missing
    result.headers_unknown = unknown

    if missing:
        result.blocked = True
        result.blocking_reasons.append(
            "Faltan columnas esperadas en el archivo: " + ", ".join(missing)
        )

    frame = pd.DataFrame(rows)
    if frame.empty:
        result.blocked = True
        result.blocking_reasons.append("El archivo no contiene filas de datos.")
        return result

    # Estado_Entrega: deducir del color de la celda Admision
    if admin_col and len(colores) == len(frame):
        frame["Estado_Entrega"] = [
            _state_from_color(rgb, entregada_hex, faltante_hex) for rgb in colores
        ]
        known = {entregada_hex, faltante_hex}
        for rgb, fila in zip(colores, frame.index):
            if rgb is not None and rgb.upper() not in known and \
               rgb.upper() not in ("FFFFFFFF", "00000000"):
                result.color_warnings.append(
                    f"Fila {fila}: color de Admision no reconocido ({rgb}) -> "
                    "se trató como Faltante."
                )

    # Mapear (normaliza, ordena, agrega _id, clasifica usuario vacío)
    frame = apply_mapping(frame, expected_cols)

    result.validation_errors = validate(frame, expected_cols)
    result.data = frame
    return result


def format_filename_date(filename: str) -> str:
    """Extrae una fecha legible del nombre del archivo, si es posible."""
    candidates = re.findall(r"(\d{4})[-_/](\d{1,2})[-_/](\d{1,2})", filename)
    if candidates:
        y, m, d = candidates[0]
        try:
            return datetime(int(y), int(m), int(d)).strftime("%d/%m/%Y")
        except ValueError:
            pass

    candidates = re.findall(r"(\d{1,2})[-_/](\d{1,2})[-_/](\d{4})", filename)
    if candidates:
        d, m, y = candidates[0]
        try:
            return datetime(int(y), int(m), int(d)).strftime("%d/%m/%Y")
        except ValueError:
            pass

    return ""
