"""Writer: construcción del Excel principal en UNA SOLA PASADA.

═══════════════════════════════════════════════════════════════════════
REGLAS ANTI-CORRUPCIÓN DE GRÁFICAS  (la causa del bug de la v1)
═══════════════════════════════════════════════════════════════════════

1) IDs de eje (axId) ÚNICOS en todo el libro.
   Excel exige que cada eje (x_axis / y_axis) de CADA gráfica tenga un
   axId distinto en TODO el libro. Si dos gráficas comparten el mismo
   axId, la parte XML de dibujo queda inválida y Excel muestra el diálogo
   de "reparación", descartando drawing*.xml. Por eso este módulo mantiene
   un contador GLOBAL `_AXID_COUNTER` que reserva valores por pares: cada
   gráfica toma (x, y) = (siguiente, siguiente+1) y NUNCA los reutiliza.
   Dejamos a un lado la asignación automática de openpyxl.

2) UNA sola pasada de guardado.
   Todo (datos, formato, tablas, gráficas) se arma sobre el MISMO objeto
   Workbook y se guarda una única vez al final con workbook.save(). Está
   PROHIBIDO abrir-guardar-reabrir el archivo a mitad de camino: cada
   save/load regenera IDs de dibujo y rompe las referencias de las
   gráficas ya insertadas, corrompiendo el archivo.

3) Anclas (celdas de inserción) únicas.
   Ninguna gráfica comparte la celda-ancla de otra. Cada gráfica recibe su
   propia posición explícita, evitando que dos objetos de dibujo apunten al
   mismo anchor.

4) Tipo de gráfica: barras 2D verticales (col), sin 3D. Tamaño consistente.
   Colores fijos por significado (Entregadas=verde, Faltantes=rojo,
   %Cumplimiento según rango configurable).

5) Referencias de datos EXACTAS. Las gráficas referencian bloques de celdas
   concretos (solo datos, sin incluir celdas de anclas ni filas de título de
   una altura distinta).
═══════════════════════════════════════════════════════════════════════
"""

from __future__ import annotations

import io
from datetime import datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule

# ── Paleta ───────────────────────────────────────────────────────────
AZUL = "1F4E78"
VERDE = "22B14C"
ROJO = "DC3545"
AMARILLO = "FFC107"
GRIS_FONDO = "F2F2F2"
GRIS_BORDE = "D9D9D9"

# ── Contador global de axId (único y nunca reutilizado) ──────────────
_AXID_COUNTER = 0


def _next_axids() -> tuple:
    """Reserva el siguiente par de axId únicos (x, y).

    El contador NUNCA retrocede, garantizando unicidad global. Arranca en
    100 para dejar margen frente a IDs internos bajos que openpyxl pudiera
    usar internamente.
    """
    global _AXID_COUNTER
    _AXID_COUNTER += 2
    return (100 + _AXID_COUNTER - 1, 100 + _AXID_COUNTER)


# ── Utilidades de formato ────────────────────────────────────────────
THIN = Side(style="thin", color=GRIS_BORDE)
BOX_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center")
LEFT = Alignment(horizontal="left", vertical="center")


def _setup_sheet(ws) -> None:
    ws.sheet_view.showGridLines = False


def _title_row(ws, row: int, text: str, span: int, fill: str = AZUL,
               font_color: str = "FFFFFF") -> None:
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = Font(bold=True, size=14, color=font_color)
    cell.fill = PatternFill("solid", fgColor=fill)
    cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    ws.row_dimensions[row].height = 28


def _kpi_card(ws, row: int, col: int, label: str, value: object,
              fill: str) -> None:
    """Dibuja una tarjeta KPI: la etiqueta encima de un valor grande."""
    ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + 1)
    ws.merge_cells(start_row=row + 1, start_column=col, end_row=row + 1,
                   end_column=col + 1)
    lb = ws.cell(row=row, column=col, value=label)
    lb.font = Font(bold=True, size=10, color="FFFFFF")
    lb.fill = PatternFill("solid", fgColor=fill)
    lb.alignment = CENTER
    val = ws.cell(row=row + 1, column=col, value=value)
    val.font = Font(bold=True, size=20, color="FFFFFF")
    val.fill = PatternFill("solid", fgColor=fill)
    val.alignment = CENTER
    for r in (row, row + 1):
        for c in range(col, col + 2):
            ws.cell(row=r, column=c).border = BOX_BORDER
    ws.row_dimensions[row].height = 22
    ws.row_dimensions[row + 1].height = 32


def _write_table(ws, start_row: int, start_col: int, df: pd.DataFrame,
                 col_widths: dict = None) -> int:
    """Escribe un DataFrame como tabla simple (encabezado sombreado AZUL).

    Devuelve la primera fila de datos (start_row + 1), útil para que las
    gráficas referencien exactamente el bloque de datos.
    """
    col_widths = col_widths or {}
    headers = [str(c) for c in df.columns]
    for j, h in enumerate(headers):
        cell = ws.cell(row=start_row, column=start_col + j, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=AZUL)
        cell.alignment = CENTER
        cell.border = BOX_BORDER
        w = col_widths.get(h, 16)
        ws.column_dimensions[get_column_letter(start_col + j)].width = w

    first_data = start_row + 1
    for i, (_, row_vals) in enumerate(df.iterrows()):
        r = first_data + i
        for j, h in enumerate(headers):
            v = row_vals[h]
            if isinstance(v, float) and v.is_integer():
                v = int(v)
            cell = ws.cell(row=r, column=start_col + j, value=v)
            cell.border = BOX_BORDER
            cell.alignment = CENTER if j != 0 else LEFT
    return first_data


def _make_chart(ws, series_first_col: int, series_last_col: int,
                title_col: int, title_row: int, cat_col: int,
                cat_row_min: int, cat_row_max: int,
                series_row_min: int, series_row_max: int,
                axid_x, axid_y, colors: list, chart_title: str,
                ancho_cm: float = 12.0, alto_cm: float = 8.0) -> BarChart:
    """Crea una BarChart 2D vertical con axId explícitos y únicos.

    Enfoque robusto (anti-corrupción):
      - Agregamos TODAS las series de una vez con `chart.add_data()` usando
        un Reference bidimensional (series_first_col..series_last_col) y
        tomando los títulos desde la fila de encabezado (titles_from_data).
        Es el método nativo de openpyxl y produce XML de serie siempre válido.
      - Luego pintamos cada serie con su color de significado.
      - Los axId (axid_x/axid_y) son SEMILLA únicos ya reservados por
        _next_axids(); garantiza unicidad global.

    Parámetros:
      series_first_col / series_last_col: rango de columnas de las series.
      title_col / title_row: celda de la que salen los nombres de serie
                            (columna de la primera serie, fila encabezado).
      cat_col:          columna de categorías.
      cat_row_min/max:  rango de filas de las categorías (datos reales).
      series_row_*:     rango de filas de los valores.
      colors:           lista de hex, uno por serie, en orden de columnas.
      chart_title:      texto del título de la gráfica.
    """
    chart = BarChart()
    chart.type = "col"
    chart.style = 10

    # Ancho/alto consistentes. openpyxl interpreta chart.width/height en
    # CENTÍMETROS (los convierte internamente a EMU = cm * 360000). NO hay que
    # multiplicar por 360000 aquí: si lo hiciéramos, el tamaño crecería ~360000x
    # y el gráfico quedaría fuera de rango (gigante y sin renderizar en Excel).
    chart.width = float(ancho_cm)
    chart.height = float(alto_cm)

    chart.title = chart_title

    # Referencia de datos: los títulos de serie se leen desde la fila de
    # encabezado (series_row_min - 1). `add_data` con titles_from_data=True
    # crea una serie por columna del rango, con su nombre correcto y XML
    # válido.
    header_row = series_row_min - 1
    if header_row >= 1:
        data = Reference(ws, min_col=series_first_col, max_col=series_last_col,
                         min_row=header_row, max_row=series_row_max)
    else:
        data = Reference(ws, min_col=series_first_col, max_col=series_last_col,
                         min_row=series_row_min, max_row=series_row_max)
    chart.add_data(data, titles_from_data=True)

    cats = Reference(ws, min_col=cat_col, min_row=cat_row_min,
                     max_row=cat_row_max)
    chart.set_categories(cats)

    # Colores de significado por serie
    for idx, color_hex in enumerate(colors[: len(chart.series)]):
        s = chart.series[idx]
        s.graphicalProperties.solidFill = color_hex

    # Etiquetas de datos: valor sobre cada barra
    labels = DataLabelList()
    labels.showVal = True
    chart.dataLabels = labels

    # ── axIds únicos + referencias cruzadas (regla anti-corrupción) ──
    # Cada gráfica del libro debe tener axIds DISTINTOS a los de las demás.
    # openpyxl usa por defecto los mismos (10 y 100) en TODAS sus gráficas:
    # eso provoca colisión de ejes y el diálogo de reparación. Aquí forzamos
    # valores únicos (reservados por _next_axids) y, CRUCIAL, actualizamos
    # también `crossAx` de cada eje para que apunte a su pareja. Si solo
    # cambiáramos axId y dejáramos el crossAx por defecto, el eje quedaría
    # referenciando un id inexistente y el archivo igualmente se consideraría
    # dañado.
    chart.x_axis.axId = axid_x
    chart.x_axis.crossAx = axid_y
    chart.y_axis.axId = axid_y
    chart.y_axis.crossAx = axid_x

    return chart


# ── Construcción principal ───────────────────────────────────────────
def build_report(metrics, data, config, output_stream,
                 faltantes_anterior: int = 0,
                 fecha_reporte: str = "") -> None:
    """Construye el reporte completo y lo escribe en output_stream (BytesIO).

    Regla #2: todo sobre un único Workbook, guardado una sola vez al final.
    """
    cfg_rangos = config.get("rangos_color", {})
    cfg_colores = config.get("colores", {})
    cfg_graf = config.get("graficas", {})
    ancho_cm = float(cfg_graf.get("ancho_cm", 12))
    alto_cm = float(cfg_graf.get("alto_cm", 8))
    reporte_nombre = config.get("reporte", {}).get("nombre", "Reporte")
    turno = config.get("reporte", {}).get("turno", "")

    def col(key: str, default: str) -> str:
        """Resuelve un color desde config (lista RGB 0-255) a hex."""
        rgb = cfg_colores.get(key)
        if rgb and isinstance(rgb, (list, tuple)) and len(rgb) == 3:
            return "".join(f"{int(v) & 0xFF:02X}" for v in rgb)
        return default

    c_entregada = col("entregadas", VERDE)
    c_faltante = col("faltantes", ROJO)
    c_alto = col("cumplimiento_alto", VERDE)
    c_medio = col("cumplimiento_medio", AMARILLO)
    c_rojo = col("cumplimiento_bajo", ROJO)

    if not fecha_reporte:
        fecha_reporte = datetime.now().strftime("%d/%m/%Y")

    wb = Workbook()

    # ── Portada ─────────────────────────────────────────────────────
    portada = wb.active
    portada.title = "Portada"
    _setup_sheet(portada)
    portada.column_dimensions["A"].width = 4
    portada.column_dimensions["B"].width = 30
    portada.column_dimensions["C"].width = 30
    portada.column_dimensions["D"].width = 30
    _title_row(portada, 3, reporte_nombre, 3)
    portada.cell(row=6, column=2,
                 value=f"Fecha del reporte: {fecha_reporte}").font = Font(bold=True, size=12)
    portada.cell(row=8, column=2, value="Turno:").font = Font(bold=True)
    portada.cell(row=8, column=3, value=turno)

    # ── Resumen ─────────────────────────────────────────────────────
    resumen = wb.create_sheet("Resumen")
    _setup_sheet(resumen)
    for c in range(1, 9):
        resumen.column_dimensions[get_column_letter(c)].width = 16

    _title_row(resumen, 1, "RESUMEN DE ADMISIONES", 8)

    kpi = metrics.kpis()
    cards = [
        ("TOTAL", kpi["total"], AZUL),
        ("ENTREGADAS", kpi["entregadas"], VERDE),
        ("FALTANTES", kpi["faltantes"], ROJO),
        ("% CUMPLIMIENTO", f"{kpi['cumplimiento']}%", AMARILLO),
    ]
    for idx, (label, value, color) in enumerate(cards):
        _kpi_card(resumen, 3, 1 + idx * 2, label, value, color)

    # Campo manual: Faltantes del informe anterior (NUNCA se sobrescribe)
    fia = resumen.cell(row=6, column=1,
                       value="Faltantes del informe anterior (manual)")
    fia.font = Font(bold=True)
    fia.alignment = LEFT
    resumen.merge_cells("A6:B6")
    cel = resumen.cell(row=6, column=3, value=faltantes_anterior)
    cel.border = BOX_BORDER
    cel.alignment = CENTER
    cel.fill = PatternFill("solid", fgColor="FFF3CD")
    cel.number_format = "0"

    # Tabla Estado / Cantidad (filas 8..)
    estado_df = metrics.estado_table()
    _title_row(resumen, 8, "Tabla Estado / Cantidad", 3)
    est_data_row = _write_table(resumen, 9, 1, estado_df, {"Estado": 25})

    # Tabla Rendimiento por Usuario (resumen, sin TOTAL aquí para no duplicar
    # la fila TOTAL GENERAL que va en Análisis; aquí mostramos tabla compacta)
    usu_row = est_data_row + len(estado_df) + 3  # dejar espacio
    _title_row(resumen, usu_row, "Rendimiento por Usuario (resumen)", 6)
    usr_df = metrics.usuario_table()  # sin TOTAL GENERAL
    if not usr_df.empty:
        _write_table(resumen, usu_row + 1, 1, usr_df,
                     {"Usuario": 28, "%Cumplimiento": 16})

    # ── Gráfica 1 del Resumen: Estado/Cantidad ─────────────────────
    # axId únicos (reservados aquí, nunca reutilizados)
    ax1 = _next_axids()
    chart1 = _make_chart(
        resumen,
        series_first_col=2, series_last_col=2,
        title_col=2, title_row=est_data_row - 1,
        cat_col=1, cat_row_min=est_data_row,
        cat_row_max=est_data_row + len(estado_df) - 1,
        series_row_min=est_data_row,
        series_row_max=est_data_row + len(estado_df) - 1,
        axid_x=ax1[0], axid_y=ax1[1], colors=[c_entregada],
        chart_title="Admisiones por Estado",
        ancho_cm=ancho_cm, alto_cm=alto_cm,
    )
    # Anchor propio (nunca compartido): celda E17 de la hoja Resumen
    resumen.add_chart(chart1, "E17")

    # ── Detalle_Admisiones ──────────────────────────────────────────
    detalle = wb.create_sheet("Detalle_Admisiones")
    _setup_sheet(detalle)
    # Columnas a EXCLUIR del detalle (técnicas o prescindibles): el id
    # anti-duplicados interno y los metadatos de Correo/Teléfono.
    detail_exclude = {"_id", "Correo", "Telefóno", "Teléfono", "Teléfono 2",
                      "Telefono", "Teléfono2"}
    detail_cols = [c for c in data.columns if c not in detail_exclude]
    if not data.empty and detail_cols:
        cols = detail_cols
        _title_row(detalle, 1, "DETALLE DE ADMISIONES", len(cols))
        for j, c in enumerate(cols):
            cell = detalle.cell(row=2, column=j + 1, value=c)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor=AZUL)
            cell.alignment = CENTER
            cell.border = BOX_BORDER
        for i, (_, row_vals) in enumerate(data.iterrows()):
            r = 3 + i
            for j, c in enumerate(cols):
                v = row_vals[c]
                if isinstance(v, float) and v.is_integer():
                    v = int(v)
                cell = detalle.cell(row=r, column=j + 1, value=v)
                cell.border = BOX_BORDER
        nrows = len(data)
        ncols = len(cols)
        ref = f"A2:{get_column_letter(ncols)}{2 + nrows}"
        tbl = Table(displayName="TblDetalle", ref=ref)
        tbl.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium9", showFirstColumn=False,
            showLastColumn=False, showRowStripes=True, showColumnStripes=False)
        detalle.add_table(tbl)
        # Formato condicional por Estado_Entrega
        try:
            est_col = cols.index("Estado_Entrega") + 1
        except ValueError:
            est_col = None
        if est_col:
            letter = get_column_letter(est_col)
            rng = f"{letter}3:{letter}{2 + nrows}"
            detalle.conditional_formatting.add(
                rng,
                CellIsRule(operator="equal", formula=['"Entregada"'],
                           fill=PatternFill("solid", fgColor="C6EFCE"),
                           font=Font(color="008000", bold=True)))
            detalle.conditional_formatting.add(
                rng,
                CellIsRule(operator="equal", formula=['"Faltante"'],
                           fill=PatternFill("solid", fgColor="FFC7CE"),
                           font=Font(color="FF0000", bold=True)))
    else:
        detalle.cell(row=2, column=1, value="Sin datos").font = Font(italic=True)

    # ── Analisis_Usuario ────────────────────────────────────────────
    analisis = wb.create_sheet("Analisis_Usuario")
    _setup_sheet(analisis)
    for c in range(1, 7):
        analisis.column_dimensions[get_column_letter(c)].width = 18
    _title_row(analisis, 1, "ANÁLISIS POR USUARIO", 6)

    usr_full = metrics.usuario_table_with_total()
    last_a = 3
    if not usr_full.empty:
        usr_data_row = _write_table(analisis, 3, 1, usr_full,
                                    {"Usuario": 28, "%Cumplimiento": 16})
        last_a = usr_data_row + len(usr_full) - 1
        # Resaltar fila TOTAL GENERAL
        for c in range(1, len(usr_full.columns) + 1):
            cell = analisis.cell(row=last_a, column=c)
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="DDEBF7")
        # Formato condicional por rango de cumplimiento
        cump_col = usr_full.columns.get_loc("%Cumplimiento") + 1
        letter = get_column_letter(cump_col)
        rng = f"{letter}{usr_data_row}:{letter}{last_a - 1}"
        vd = float(cfg_rangos.get("verde_desde", 70))
        vh = float(cfg_rangos.get("verde_hasta", 90))
        analisis.conditional_formatting.add(
            rng, CellIsRule(operator="lessThan", formula=[f"{vd}"],
                            fill=PatternFill("solid", fgColor="FFC7CE")))
        analisis.conditional_formatting.add(
            rng, CellIsRule(operator="lessThan", formula=[f"{vh}"],
                            fill=PatternFill("solid", fgColor="FFEB9C")))
        analisis.conditional_formatting.add(
            rng, CellIsRule(operator="greaterThanOrEqual", formula=[f"{vh}"],
                            fill=PatternFill("solid", fgColor="C6EFCE")))

        # ── Gráficas comparativas del Análisis ──────────────────────
        # Gráfica 2 (axIds únicos): Entregadas y Faltantes por usuario.
        # Datos: la tabla empieza en la fila 3 (encabezado) y los datos van
        # de usr_data_row a last_a - 1 (excluye la fila TOTAL GENERAL).
        ax2 = _next_axids()
        chart2 = _make_chart(
            analisis,
            series_first_col=2, series_last_col=3,
            title_col=2, title_row=usr_data_row - 1,
            cat_col=1, cat_row_min=usr_data_row,
            cat_row_max=last_a - 1,
            series_row_min=usr_data_row,
            series_row_max=last_a - 1,
            axid_x=ax2[0], axid_y=ax2[1],
            colors=[c_entregada, c_faltante],
            chart_title="Entregadas vs Faltantes por Usuario",
            ancho_cm=ancho_cm, alto_cm=alto_cm,
        )
        analisis.add_chart(chart2, f"A{last_a + 3}")

        # Gráfica 3 (axIds únicos): %Cumplimiento por usuario. Cada barra se
        # pinta con el color de SU rango (rojo/amarillo/verde) según el valor
        # de %Cumplimiento, cumpliendo el requisito de "color según rango, no
        # aleatorio". Se aplica como color por punto de dato (DataPoint).
        ax3 = _next_axids()
        chart3 = _make_chart(
            analisis,
            series_first_col=5, series_last_col=5,
            title_col=5, title_row=usr_data_row - 1,
            cat_col=1, cat_row_min=usr_data_row,
            cat_row_max=last_a - 1,
            series_row_min=usr_data_row,
            series_row_max=last_a - 1,
            axid_x=ax3[0], axid_y=ax3[1], colors=[c_alto],
            chart_title="% Cumplimiento por Usuario",
            ancho_cm=ancho_cm, alto_cm=alto_cm,
        )
        # Colorear cada barra por su rango de cumplimiento
        vd = float(cfg_rangos.get("verde_desde", 70))
        vh = float(cfg_rangos.get("verde_hasta", 90))
        serie3 = chart3.series[0]
        serie3.data_points = []
        for k, (_, fila) in enumerate(usr_full.iloc[:-1].iterrows()):
            pct = float(fila["%Cumplimiento"])
            if pct < vd:
                hexc = c_rojo
            elif pct < vh:
                hexc = c_medio
            else:
                hexc = c_alto
            from openpyxl.chart.marker import DataPoint
            dp = DataPoint(idx=k)
            dp.graphicalProperties.solidFill = hexc
            serie3.data_points.append(dp)
        analisis.add_chart(chart3, f"G{last_a + 3}")
    else:
        analisis.cell(row=3, column=1, value="Sin datos").font = Font(italic=True)

    # ── Evidencias ──────────────────────────────────────────────────
    ev = wb.create_sheet("Evidencias")
    _setup_sheet(ev)
    _title_row(ev, 1, "EVIDENCIAS (pegue 4 imágenes manualmente)", 6)
    for idx in range(4):
        r = 3 + idx * 2
        ev.cell(row=r, column=1,
                value=f"Evidencia {idx + 1}").font = Font(bold=True, size=12)
        ev.merge_cells(start_row=r + 1, start_column=1, end_row=r + 1, end_column=4)
        box = ev.cell(row=r + 1, column=1,
                      value="Haga clic y pegue aquí su imagen (Insertar > Imagen)")
        box.border = BOX_BORDER
        box.alignment = CENTER
        box.fill = PatternFill("solid", fgColor=GRIS_FONDO)
        ev.row_dimensions[r + 1].height = 40

    # ── Guardar UNA sola vez ────────────────────────────────────────
    wb.save(output_stream)
