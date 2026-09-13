"""Prueba completa del pipeline en UN comando.

Uso:
    venv/Scripts/python.exe prueba.py [archivo_fuente.xlsx]

Sin argumento usa _temp_upload.xlsx. Genera prueba_reporte.xlsx y lo
valida (verify_ooxml) + apertura por Windows COM si hay Excel.
"""

import os
import sys
import win32com.client

import yaml

from src import loader, metrics, writer

APP_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG = os.path.join(APP_DIR, "config.yaml")
FUENTE = sys.argv[1] if len(sys.argv) > 1 else os.path.join(APP_DIR, "_temp_upload.xlsx")
SALIDA = os.path.join(APP_DIR, "prueba_reporte.xlsx")


def main() -> None:
    print("1) Cargando config...")
    config = yaml.safe_load(open(CONFIG, encoding="utf-8"))
    expected = [c.strip() for c in config.get("columnas_esperadas", [])]

    print(f"2) Leyendo fuente: {os.path.basename(FUENTE)}")
    res = loader.load(FUENTE, expected, state_colors=config.get("colores_estado", {}))
    if res.blocked:
        print("   BLOQUEADO:", res.blocking_reasons)
        sys.exit(1)
    print(f"   filas: {len(res.data)} | warnings color: {len(res.color_warnings)}")
    err = {k: len(v) for k, v in res.validation_errors.items() if v}
    print("   errores detectados:", err or "ninguno")

    met = metrics.Metrics(res.data, total_esperado=int(config.get("total_esperado", 0)))
    print("   KPIs:", met.kpis())

    print("3) Generando reporte...")
    import io
    out = io.BytesIO()
    writer.build_report(met, res.data, config, out,
                        faltantes_anterior=0, fecha_reporte="07/09/2026")
    with open(SALIDA, "wb") as fh:
        fh.write(out.getbuffer())
    print(f"   guardado: {os.path.basename(SALIDA)} ({len(out.getbuffer())} bytes)")

    print("4) Validando contra corrupcion (verify_ooxml)...")
    import verify_ooxml
    probs = verify_ooxml.verify_workbook(SALIDA)
    if probs:
        print("   PROBLEMAS:", *probs, sep="\n    - ")
        sys.exit(1)
    print("   OK: sin problemas de corrupcion.")

    print("5) Abriendo en Excel (COM)...")
    xl = win32com.client.Dispatch("Excel.Application")
    xl.DisplayAlerts = False
    xl.Visible = False
    try:
        wb = xl.Workbooks.Open(SALIDA, ReadOnly=True, CorruptLoad=0)
        for i in range(1, wb.Worksheets.Count + 1):
            ws = wb.Worksheets(i)
            print(f"   - {ws.Name}: {ws.ChartObjects().Count} graficas")
        port = wb.Sheets("Portada")
        print("   Portada C40:", port.Range("C40").Value)
        print("   Portada C33:", port.Range("C33").Value)
        wb.Close(SaveChanges=False)
    finally:
        xl.Quit()
    print("TODO OK. Abre 'prueba_reporte.xlsx' para revisar la portada y las graficas.")


if __name__ == "__main__":
    main()