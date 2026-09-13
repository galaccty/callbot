"""Interfaz Streamlit 100% local para automatizar el reporte de admisiones.

Flujo:
  1. Sube el Excel fuente con un botón.
  2. Pulsa "Procesar y actualizar".
  3. El sistema valida los datos (reporta errores en pantalla) y genera el
     Excel principal con dashboards y gráficas.
  4. "Descargar reporte" entrega el .xlsx ya listo.

Ejecución:  streamlit run app.py
"""

from __future__ import annotations

import io
from datetime import datetime

import pandas as pd
import streamlit as st
import yaml

from src import loader, metrics, writer

APP_DIR = __file__.rsplit("\\", 1)[0] if "\\" in __file__ else __file__.rsplit("/", 1)[0]


def load_config() -> dict:
    path = f"{APP_DIR}/config.yaml"
    with open(path, "r", encoding="utf-8") as fh:
        return yaml.safe_load(fh)


def main() -> None:
    st.set_page_config(page_title="Automatización de Admisiones",
                       layout="wide")
    st.title("Automatización de Reporte de Admisiones")
    st.caption("Suba el Excel fuente, procese y descargue el reporte final.")

    config = load_config()
    expected = [c.strip() for c in config.get("columnas_esperadas", [])]
    cfg_graf = config.get("graficas", {})

    # ── Panel de configuración rápida (sin tocar código) ───────────
    with st.expander("Configuración de la corrida"):
        total_esperado = st.number_input(
            "Total esperado de admisiones", min_value=0,
            value=int(config.get("total_esperado", 0)),
            help="Se usa en los KPIs y en el % de cumplimiento.")
        faltantes_anterior = st.number_input(
            "Faltantes del informe anterior (manual)", min_value=0, value=0,
            help="El usuario lo llena a mano; el sistema nunca lo sobrescribe.")
        turno = st.text_input("Turno", value=config.get("reporte", {}).get("turno", ""))
        responsable = st.text_input(
            "Responsable", value=config.get("reporte", {}).get("responsable", ""))
        ancho_cm = st.number_input(
            "Ancho de gráficas (cm)", min_value=3.0,
            value=float(cfg_graf.get("ancho_cm", 7.0)), step=0.5)
        alto_cm = st.number_input(
            "Alto de gráficas (cm)", min_value=2.0,
            value=float(cfg_graf.get("alto_cm", 4.5)), step=0.5)

    uploaded = st.file_uploader(
        "1) Suba el Excel de admisiones (.xlsx)", type=["xlsx"])

    if uploaded is not None:
        st.success(f"Archivo listo: {uploaded.name}")

    if st.button("2) Procesar y actualizar", type="primary",
                 disabled=uploaded is None):
        if uploaded is None:
            st.error("Primero suba un archivo.")
            st.stop()

        # Guardar el archivo temporalmente para leerlo con pandas
        temp_path = f"{APP_DIR}/_temp_upload.xlsx"
        with open(temp_path, "wb") as fh:
            fh.write(uploaded.getbuffer())

        state_colors = config.get("colores_estado", {})
        with st.spinner("Procesando..."):
            res = loader.load(temp_path, expected, state_colors=state_colors)

        if res.color_warnings:
            for w in res.color_warnings[:20]:
                st.warning(w)

        # ── Mostrar validaciones siempre (sin borrar datos) ────────
        if res.blocked:
            st.error("NO se pudo procesar el archivo:")
            for razon in res.blocking_reasons:
                st.error(razon)
            if res.data.empty:
                st.stop()
            st.write("Primeras filas del archivo leído:")
            st.dataframe(res.data.head(20))
            st.stop()

        if res.validation_errors:
            total_err = sum(len(v) for v in res.validation_errors.values())
            if total_err:
                st.warning(res.error_summary)
            else:
                st.success("Validación sin errores.")

        # ── Calcular métricas ──────────────────────────────────────
        met = metrics.Metrics(res.data, total_esperado=total_esperado)

        # Fecha del reporte (nombre del archivo o fecha del sistema)
        fecha = loader.format_filename_date(uploaded.name) or datetime.now().strftime("%d/%m/%Y")

        # ── Mostrar KPIs en pantalla ───────────────────────────────
        kpi = met.kpis()
        cols = st.columns(5)
        cols[0].metric("Total", kpi["total"])
        cols[1].metric("Entregadas", kpi["entregadas"])
        cols[2].metric("Faltantes", kpi["faltantes"])
        cols[3].metric("% Cumplimiento", f"{kpi['cumplimiento']}%")
        cols[4].metric("Total esperado", kpi["total_esperado"])

        # Vista previa del detalle
        with st.expander("Vista previa del detalle (primeras 20 filas)"):
            st.dataframe(res.data.drop(columns=["_id"], errors="ignore").head(20))

        # ── Generar el Excel en una sola pasada ────────────────────
        output = io.BytesIO()
        writer.build_report(
            met, res.data, config, output,
            faltantes_anterior=int(faltantes_anterior),
            fecha_reporte=fecha,
        )
        output.seek(0)

        base_name = uploaded.name.replace(".xlsx", "").replace(".xls", "")
        out_name = f"Reporte_Admisiones_{base_name}.xlsx"

        st.download_button(
            "3) Descargar reporte (.xlsx)",
            data=output,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
        st.caption(
            "Recuerde: luego de descargar, pegue manualmente las 4 imágenes "
            "de evidencia en la hoja 'Evidencias'."
        )


if __name__ == "__main__":
    main()
