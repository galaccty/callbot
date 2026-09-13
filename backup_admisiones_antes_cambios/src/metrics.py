"""Metrics: cálculo de KPIs y tablas derivadas a partir del DataFrame limpio.

Todas las funciones son puras (no escriben Excel); devuelven DataFrames y
valores que el writer usará. Esto separa la lógica de negocio de la capa de
presentación, facilitando pruebas.
"""

from __future__ import annotations

from typing import Dict

import pandas as pd


class Metrics:
    """Calcula todas las métricas de una corrida."""

    def __init__(self, data: pd.DataFrame, total_esperado: int = 0):
        self.data = data
        self.total_esperado = int(total_esperado) if total_esperado else 0
        self._total = len(data)
        self._entregadas = int(
            (data["Estado_Entrega"].str.lower() == "entregada").sum()
            if "Estado_Entrega" in data.columns else 0
        )
        self._faltantes = int(
            (data["Estado_Entrega"].str.lower() == "faltante").sum()
            if "Estado_Entrega" in data.columns else 0
        )

    @property
    def total(self) -> int:
        """Total de registros cargados (sin los manuales)."""
        return self._total

    @property
    def entregadas(self) -> int:
        return self._entregadas

    @property
    def faltantes(self) -> int:
        return self._faltantes

    @property
    def cumplimiento(self) -> float:
        """%Cumplimiento = entregadas / total (0 si no hay datos)."""
        if self._total == 0:
            return 0.0
        return round(self._entregadas / self._total * 100, 1)

    def estado_table(self) -> pd.DataFrame:
        """Tabla Estado / Cantidad."""
        if self.data.empty:
            return pd.DataFrame({"Estado": ["Entregada", "Faltante"],
                                 "Cantidad": [0, 0]})
        s = self.data["Estado_Entrega"].replace("", "Sin estado")
        counts = s.value_counts().reset_index()
        counts.columns = ["Estado", "Cantidad"]
        return counts

    def usuario_table(self) -> pd.DataFrame:
        """Rendimiento por Usuario: Entregadas/Faltantes/Total/%Cumplimiento."""
        if self.data.empty:
            return pd.DataFrame(
                columns=["Usuario", "Entregadas", "Faltantes", "Total", "%Cumplimiento"]
            )

        grp = self.data.groupby("Usuario")
        entregadas = grp["Estado_Entrega"].apply(
            lambda s: (s.str.lower() == "entregada").sum()
        )
        faltantes = grp["Estado_Entrega"].apply(
            lambda s: (s.str.lower() == "faltante").sum()
        )
        total = grp.size()
        cumple = (
            (entregadas / total.where(total > 0) * 100)
            .fillna(0)
            .round(1)
        )

        tabla = pd.DataFrame({
            "Usuario": entregadas.index,
            "Entregadas": entregadas.values,
            "Faltantes": faltantes.values,
            "Total": total.values,
            "%Cumplimiento": cumple.values,
        })
        tabla = tabla.sort_values("Total", ascending=False).reset_index(drop=True)
        return tabla

    def usuario_table_with_total(self) -> pd.DataFrame:
        """Tabla de usuario + fila TOTAL GENERAL."""
        tabla = self.usuario_table()
        if tabla.empty:
            return tabla
        total_row = pd.DataFrame([{
            "Usuario": "TOTAL GENERAL",
            "Entregadas": self.entregadas,
            "Faltantes": self.faltantes,
            "Total": self.total,
            "%Cumplimiento": self.cumplimiento,
        }])
        return pd.concat([tabla, total_row], ignore_index=True)

    def kpis(self) -> Dict[str, object]:
        """Conjunto de KPIs para las tarjetas del Resumen."""
        return {
            "total": self.total,
            "entregadas": self.entregadas,
            "faltantes": self.faltantes,
            "cumplimiento": self.cumplimiento,
            "total_esperado": self.total_esperado,
        }
