"""Writer: construcción del Excel principal en UNA SOLA PASADA.

═══════════════════════════════════════════════════════════════════════
REGLAS ANTI-CORRUPCIÓN (lección del bug de la v1)
═══════════════════════════════════════════════════════════════════════

1) UNA sola pasada de guardado.
   Todo (datos, formato, tablas, gráficas) se arma sobre el MISMO objeto
   Workbook y se guarda una única vez al final con workbook.save(). Está
   PROHIBIDO abrir-guardar-reabrir el archivo a mitad de camino: cada
   save/load regenera IDs de dibujo y rompe las referencias de las
   gráficas ya insertadas, corrompiendo el archivo.

2) Gráficas CIRCULARES (pie 2D), sin leyenda.
   Las gráficas de barras se reemplazaron por pizzas (pie charts). Al no
   tener ejes, desaparece el antiguo riesgo de colisión de axId: este
   módulo ya no reserva IDs de eje. Se conservan las demás reglas duras.

3) Anclas (celdas de inserción) únicas.
   Ninguna gráfica comparte la celda-ancla de otra. Cada gráfica recibe su
   propia posición explícita, evitando que dos objetos de dibujo apunten al
   mismo anchor.

4) Etiquetas CON LÍNEA GUÍA (leader line).
   Cada etiqueta (nombre de categoría + valor) sale pegada a su porción con
   una rayita desde la zona señalada del círculo; no hay leyenda aparte.

5) Referencias de datos EXACTAS y VIVAS. Las gráficas referencian los
   bloques de celdas de las tablas del propio libro (Reference), nunca
   valores fijos: al cambiar los números en esas celdas la gráfica se
   actualiza, y con cada corrida se regenera desde los datos nuevos.

6) Portada automática. La hoja Portada reproduce la portada que se hace a
   mano (logo + textos corporativos + fecha del reporte en formato largo).
═══════════════════════════════════════════════════════════════════════
"""

from __future__ import annotations

import io
import os
from datetime import datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import DataPoint
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.styles.colors import Color
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule
from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor, AnchorMarker

# ── Paleta ───────────────────────────────────────────────────────────
AZUL = "1F4E78"
VERDE = "22B14C"
ROJO = "DC3545"
AMARILLO = "FFC107"
GRIS_FONDO = "F2F2F2"
GRIS_BORDE = "D9D9D9"
AZUL_PORTADA = "FF1B3A5C"    # azul corporativo (títulos de la portada)
DORADO_PORTADA = "FFC9971E"  # dorado corporativo (subtítulo)
GRIS_PORTADA = "FF555555"    # gris (fecha de la portada)

# Ruta del logo corporativo (assets/logo.png junto a la app)
_PROJECT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LOGO_PATH = os.path.join(_PROJECT_DIR, "assets", "logo.png")

# Meses en español para la fecha larga de la portada ("7 de septiembre de 2026")
_MESES = ["", "enero", "febrero", "marzo", "abril", "mayo", "junio",
          "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]


def _fecha_larga(fecha: str) -> str:
    """Convierte '07/09/2026' -> '7 de septiembre de 2026'."""
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%d/%m/%Y %H:%M"):
        try:
            d = datetime.strptime(fecha.strip(), fmt)
            return f"{d.day} de {_MESES[d.month]} de {d.year}"
        except ValueError:
            continue
    return fecha


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


def _make_pie_chart(ws, cat_col: int, data_col: int, header_row: int,
                    data_row_min: int, data_row_max: int,
                    point_colors: list, chart_title: str,
                    ancho_cm: float = 7.0, alto_cm: float = 4.5) -> PieChart:
    """Crea un PieChart 2D (circular) con etiqueta por porcion y linea guia.

    - Sin leyenda: cada categoria se conecta a su porcion con una rayita
      (leader line) desde la zona señalada del circulo hasta el texto.
    - Colores BASICOS por porcion (DataPoint), nunca aleatorios.
    - Referencia las CELDAS VIVAS de la tabla origen (no valores fijos).

    Parámetros:
      ws:              hoja donde vive la tabla origen.
      cat_col:         columna de categorias (nombres de las porciones).
      data_col:        columna con el valor numerico de cada porcion.
      header_row:      fila del encabezado (titulo de la serie).
      data_row_min/max: rango de filas de DATOS (excluye filas de titulo).
      point_colors:    lista de hex, UNO por fila de datos (porcion).
      chart_title:     titulo de la grafica.
    """
    chart = PieChart()
    chart.type = "pie"
    chart.style = 10

    # Tamaño (cm); openpyxl interpreta width/height en centímetros.
    chart.width = float(ancho_cm)
    chart.height = float(alto_cm)

    chart.title = chart_title
    chart.varyColors = False  # colores controlados por nosotros, no aleatorios

    # Serie: datos desde la celda de encabezado hasta la ultima fila de datos.
    data = Reference(ws, min_col=data_col, max_col=data_col,
                     min_row=header_row, max_row=data_row_max)
    chart.add_data(data, titles_from_data=True)

    cats = Reference(ws, min_col=cat_col, min_row=data_row_min,
                     max_row=data_row_max)
    chart.set_categories(cats)

    # Color por porcion (rojo/verde segun significado, uno por fila)
    n_puntos = data_row_max - data_row_min + 1
    if chart.series:
        serie = chart.series[0]
        serie.data_points = []
        for k in range(n_puntos):
            if k < len(point_colors):
                dp = DataPoint(idx=k)
                dp.graphicalProperties.solidFill = point_colors[k]
                serie.data_points.append(dp)

    # Etiquetas: categoria + valor, pegadas a su porcion con linea guia.
    labels = DataLabelList()
    labels.showCatName = True
    labels.showVal = True
    labels.showPercent = False
    labels.showLeaderLines = True
    labels.separator = "\n"
    labels.dLblPos = "outEnd"
    chart.dataLabels = labels
    chart.legend = None

    return chart


def _build_portada(ws, fecha_larga: str) -> None:
    """Replica la portada manual de referencia (logo + textos corporativos).

    El diseño copia exactamente hoja `Portada` del informe terminado:
    logo, textos, fusiones, altos/anchos y la fecha del reporte en formato
    largo ("Reporte SIIPS : 7 de septiembre de 2026").
    """
    _setup_sheet(ws)
    ws.column_dimensions["A"].width = 4
    for col in ("B", "C", "D"):
        ws.column_dimensions[col].width = 30

    altos = {3: 28.05, 27: 21, 33: 36.6, 35: 23.4, 39: 15.6, 40: 15.6}
    for row, h in altos.items():
        ws.row_dimensions[row].height = h

    for rng in ("C27:H27", "C29:H29", "C33:H33",
                "C35:H35", "C39:H39", "C40:H40"):
        ws.merge_cells(rng)

    # Textos corporativos (fijos)
    def _texto(coord, value, size, bold, color):
        cell = ws[coord]
        cell.value = value
        cell.font = Font(name="Calibri", size=size, bold=bold, color=color)
        cell.alignment = Alignment(horizontal="center")

    _texto("C27", "IPS INTEGRAL SOMOS SALUD S.A.S.", 16, True, AZUL_PORTADA)
    _texto("C33", "INFORME CALL CENTER", 28, True, AZUL_PORTADA)
    _texto("C35", "GESTIÓN Y CONTROL DE ADMISIONES", 18, True, DORADO_PORTADA)
    _texto("C40", f"Reporte SIIPS : {fecha_larga}", 12, False, GRIS_PORTADA)
    # Filas de separacion vacias conservan su estilo (banda / texto gris).
    ws["C29"].fill = PatternFill("solid", fgColor=Color(theme=4, tint=-0.25))
    ws["C39"].font = Font(name="Calibri", size=12, color=GRIS_PORTADA)
    ws["C39"].alignment = Alignment(horizontal="center")

    # Logo: misma posicion y tamano que en el informe terminado.
    try:
        img = Image(LOGO_PATH)
        img.anchor = TwoCellAnchor(
            editAs="oneCell",
            _from=AnchorMarker(col=2, colOff=1690204, row=14, rowOff=21757),
            to=AnchorMarker(col=5, colOff=330338, row=25, rowOff=122721),
        )
        ws.add_image(img)
    except Exception:  # noqa: BLE001
        # Sin logo no se aborta el reporte.
        pass


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
    fecha_larga = _fecha_larga(fecha_reporte)

    wb = Workbook()

    # ── Portada (automatica, replica la portada manual) ──────────────
    portada = wb.active
    portada.title = "Portada"
    _build_portada(portada, fecha_larga)

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

    # ── Gráfica 1 del Resumen: Admisiones por Estado (circular) ─────
    # Pizzas 2D sin ejes: ancla propia (nunca compartida) en E17.
    # Colores por categoría: Entregada=verde, Faltante=rojo.
    colores_est = [
        c_entregada if str(e) == "Entregada" else c_faltante
        for e in estado_df["Estado"]
    ]
    chart1 = _make_pie_chart(
        resumen,
        cat_col=1, data_col=2,
        header_row=est_data_row - 1,
        data_row_min=est_data_row,
        data_row_max=est_data_row + len(estado_df) - 1,
        point_colors=colores_est,
        chart_title="Admisiones por Estado",
        ancho_cm=ancho_cm, alto_cm=alto_cm,
    )
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

        # ── Gráficas circulares del Análisis ─────────────────────────
        # Datos: la tabla empieza en la fila 3 (encabezado) y los datos van
        # de usr_data_row a last_a - 1 (excluye la fila TOTAL GENERAL).
        n_usr = last_a - usr_data_row
        vd = float(cfg_rangos.get("verde_desde", 70))

        # Gráfica 2: "Entregadas por Usuario" (porciones todas verde).
        chart2 = _make_pie_chart(
            analisis,
            cat_col=1, data_col=2,
            header_row=usr_data_row - 1,
            data_row_min=usr_data_row, data_row_max=last_a - 1,
            point_colors=[c_entregada] * n_usr,
            chart_title="Entregadas por Usuario",
            ancho_cm=ancho_cm, alto_cm=alto_cm,
        )
        analisis.add_chart(chart2, f"A{last_a + 3}")

        # Gráfica 3: "Faltantes por Usuario" (porciones todas rojo).
        chart3 = _make_pie_chart(
            analisis,
            cat_col=1, data_col=3,
            header_row=usr_data_row - 1,
            data_row_min=usr_data_row, data_row_max=last_a - 1,
            point_colors=[c_faltante] * n_usr,
            chart_title="Faltantes por Usuario",
            ancho_cm=ancho_cm, alto_cm=alto_cm,
        )
        analisis.add_chart(chart3, f"G{last_a + 3}")

        # Gráfica 4: "% Cumplimiento por Usuario" (verde si cumple el umbral
        # 'verde_desde', rojo si no; colores basicos, no aleatorios).
        cols_cumple = [
            c_entregada if float(fila["%Cumplimiento"]) >= vd else c_rojo
            for (_, fila) in usr_full.iloc[:-1].iterrows()
        ]
        chart4 = _make_pie_chart(
            analisis,
            cat_col=1, data_col=5,
            header_row=usr_data_row - 1,
            data_row_min=usr_data_row, data_row_max=last_a - 1,
            point_colors=cols_cumple,
            chart_title="% Cumplimiento por Usuario",
            ancho_cm=ancho_cm, alto_cm=alto_cm,
        )
        analisis.add_chart(chart4, f"M{last_a + 3}")
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
