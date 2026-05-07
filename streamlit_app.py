from __future__ import annotations

import re
from pathlib import Path
from tempfile import TemporaryDirectory

import streamlit as st

from formatear_banners_hogar_final import format_workbook


st.set_page_config(page_title="Formato de Tablas", page_icon="📊", layout="centered")

st.title("Formato de Tablas")
st.caption("Sube un archivo .xlsx, aplica el formato automáticamente y descarga el resultado.")

uploaded_file = st.file_uploader("Archivo Excel", type=["xlsx"])

if uploaded_file is not None:
    input_name = Path(uploaded_file.name)
    default_output_name = f"{input_name.stem} FORMATEADO.xlsx"

    output_name = st.text_input("Nombre del archivo de salida", value=default_output_name)

    if st.button("Formatear archivo", type="primary"):
        if not output_name.strip().lower().endswith(".xlsx"):
            st.error("El archivo de salida debe terminar en .xlsx")
        else:
            try:
                progress_bar = st.progress(0, text="0% - Iniciando proceso")
                status_placeholder = st.empty()
                sheet_placeholder = st.empty()

                def on_progress(progress: float, message: str) -> None:
                    pct = min(100, max(0, int(round(progress * 100))))
                    progress_bar.progress(pct, text=f"{pct}% - {message}")
                    status_placeholder.info(message)
                    match = re.search(r"\(([^)]+)\):", message)
                    if match:
                        sheet_placeholder.info(f"Hoja actual: {match.group(1)}")
                    elif "Hojas objetivo" in message:
                        sheet_placeholder.info("Hoja actual: preparando procesamiento")

                with TemporaryDirectory() as temp_dir:
                    temp_dir_path = Path(temp_dir)
                    temp_input = temp_dir_path / input_name.name
                    temp_output = temp_dir_path / output_name.strip()

                    on_progress(0.02, "Preparando archivo de entrada")
                    temp_input.write_bytes(uploaded_file.getbuffer())
                    format_workbook(temp_input, temp_output, progress_callback=on_progress)
                    output_bytes = temp_output.read_bytes()
                    on_progress(1.0, "Archivo listo para descargar")

                st.success("Archivo formateado correctamente.")
                st.download_button(
                    label="Descargar archivo formateado",
                    data=output_bytes,
                    file_name=output_name.strip(),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as exc:
                st.error(f"Ocurrió un error al formatear: {exc}")

st.markdown("---")
st.markdown("Esta app usa la lógica de `formatear_banners_hogar_final.py`.")
