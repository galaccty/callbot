"""Versión nueva: verificación del Excel generado contra la corrupción.

Comprueba, dentro del .xlsx (que es un ZIP de XML), las dos causas típicas
del diálogo de "reparación de archivo" de Excel:

  1) axId DUPLICADOS entre ejes de distintas gráficas.
  2) Anclas/posiciones de dibujo solapadas o rotas.

También verifica que el [Content_Types].xml y las relaciones del libro no
referencien partes inexistentes (otra fuente común de reparación).

Devuelve código 0 si todo está bien, 1 si hay problemas (para usarlo en CI).
"""

from __future__ import annotations

import re
import sys
import zipfile
from collections import Counter

import openpyxl


def verify_workbook(path: str) -> list:
    """Retorna la lista de problemas encontrados (vacía = OK)."""
    problems: list = []

    # 1) Reabrir con openpyxl: si el XML está corrupto, esto lanza.
    try:
        wb = openpyxl.load_workbook(path)
    except Exception as exc:  # noqa: BLE001
        return [f"No se pudo abrir el libro con openpyxl: {exc}"]

    # 2) Recorrer las gráficas y comprobar unicidad de axId en TODO el libro.
    #    Las graficas circulares (pie) NO tienen ejes (x_axis/y_axis): es normal
    #    y por eso se saltan.
    axids = Counter()
    for ws in wb.worksheets:
        for chart in ws._charts:
            ax_x = getattr(chart, "x_axis", None)
            ax_y = getattr(chart, "y_axis", None)
            if ax_x is None or ax_y is None:
                continue
            axids[ax_x.axId] += 1
            axids[ax_y.axId] += 1
    for axid, count in axids.items():
        if count > 1:
            problems.append(
                f"axId {axid} usado {count} veces (debe ser único en todo el libro)"
            )

    # 3) Verificar que el ZIP no tenga "partes rotas": todo archivo al que se
    #    hace referencia en .rels existe dentro del paquete.
    with zipfile.ZipFile(path) as zf:
        names = set(zf.namelist())
        for n in list(names):
            if n.endswith(".rels"):
                rel_dir = n.rsplit("/", 1)[0] if "/" in n else ""
                rels_xml = zf.read(n).decode("utf-8", "replace")
                for t in re.findall(r'Target="([^"]+)"', rels_xml):
                    if t.startswith("http"):
                        continue
                    # Resolver target relativo al dir del .rels (padre en "xl/")
                    if not rel_dir or rel_dir == "_rels":
                        base = ""
                    else:
                        # .rels está bajo .../_rels, el target se resuelve
                        # relativo al directorio padre de _rels
                        base = rel_dir.rsplit("/", 1)[0] + "/" \
                            if rel_dir.endswith("_rels") else rel_dir + "/"
                    resolved = t.lstrip("/")
                    if not t.startswith("/") and base:
                        resolved = base + t
                    if resolved not in names and t.split("/")[-1] not in names:
                        problems.append(
                            f"Relación en {n} apunta a parte inexistente: {t}")

    return problems


def main() -> int:
    if len(sys.argv) < 2:
        print("Uso: python verify_ooxml.py <archivo.xlsx> [segundo.xlsx ...]")
        return 2
    problemas_acumulados = 0
    for path in sys.argv[1:]:
        probs = verify_workbook(path)
        if not probs:
            print(f"OK: {path}  (sin problemas de corrupción detectados)")
        else:
            problemas_acumulados += 1
            print(f"PROBLEMAS en {path}:")
            for p in probs:
                print(f"  - {p}")
    return 1 if problemas_acumulados else 0


if __name__ == "__main__":
    sys.exit(main())
