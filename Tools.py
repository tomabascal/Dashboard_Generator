import streamlit as st
import pandas as pd
import pptx
from pptx.dml.color import RGBColor
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN
import os
import time
import zipfile
import io
import shutil
from datetime import datetime
import re
import subprocess

def convert_pptx_to_pdf(pptx_path, pdf_path):
    """Convierte un archivo PPTX a PDF en Linux usando LibreOffice (funciona en Streamlit Cloud)."""
    try:
        subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf",
                       pptx_path, "--outdir", os.path.dirname(pdf_path)], check=True)
    except Exception as e:
        print(f"Error converting {pptx_path} to PDF: {e}")

def create_zip_of_presentations(folder_path):
    """Crea un archivo ZIP con todos los PPTX generados en la carpeta."""
    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        for file in os.listdir(folder_path):
            file_path = os.path.join(folder_path, file)
            if file.endswith(".pptx"):  # Evitamos incluir plantilla y Excel
                zipf.write(file_path, arcname=file)

    zip_buffer.seek(0)
    return zip_buffer

def get_filename_from_selection(row, selected_columns):
    """Genera el nombre del archivo según las columnas seleccionadas."""
    file_name_parts = [str(row[col]) for col in selected_columns if col in row]
    return "_".join(file_name_parts)

def update_text_of_textbox(presentation, column_letter, new_text):
    """Busca y reemplaza texto dentro de las cajas de texto que tengan el formato {A}, {B}, etc."""
    pattern = rf"\{{{column_letter}\}}"

    for slide in presentation.slides:
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text:
                if re.search(pattern, shape.text):
                    text_frame = shape.text_frame
                    for paragraph in text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.text = re.sub(pattern, str(new_text), run.text)

def process_files(ppt_file, excel_file, search_option, start_row, end_row, store_ids, selected_columns, output_format):
    """Genera reportes en formato PPTX o PDF en Streamlit Cloud con aviso de tiempos estimados."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    folder_name = f"Presentations_{timestamp}"
    os.makedirs(folder_name, exist_ok=True)

    temp_folder = "temp_files"
    os.makedirs(temp_folder, exist_ok=True)

    ppt_template_path = os.path.join(temp_folder, ppt_file.name)
    excel_file_path = os.path.join(temp_folder, excel_file.name)

    with open(ppt_template_path, "wb") as f:
        f.write(ppt_file.getbuffer())
    with open(excel_file_path, "wb") as f:
        f.write(excel_file.getbuffer())

    try:
        with pd.ExcelFile(excel_file_path) as xls:
            df1 = pd.read_excel(xls, sheet_name=0)
    except PermissionError as e:
        st.error(f"Error reading Excel file: {e}")
        return

    if search_option == 'rows':
        df_selected = df1.iloc[start_row:end_row + 1]
    elif search_option == 'store_id':
        store_id_list = [store_id.strip() for store_id in store_ids.split(',')]
        df_selected = df1[df1.iloc[:, 0].astype(str).isin(store_id_list)]
    else:
        df_selected = pd.DataFrame()

    total_files = len(df_selected)
    if total_files == 0:
        st.error("⚠️ No hay archivos para generar. Verifica los filtros.")
        return

    estimated_time = total_files * (5 if output_format == "PDF" else 1)
    st.info(f"⏳ Estimated time: ~{estimated_time} seconds")

    progress_bar = st.progress(0)
    progress_text = st.empty()

    current_file = 0
    start_time = time.time()

    for index, row in df_selected.iterrows():
        process_row(ppt_template_path, row, df1, index,
                    selected_columns, folder_name, output_format)
        current_file += 1
        progress = current_file / total_files
        progress_bar.progress(progress)
        elapsed_time = time.time() - start_time
        progress_text.write(f"📄 Generating {current_file}/{total_files} ({output_format}) - Elapsed time: {int(elapsed_time)}s")

    zip_path = f"{folder_name}.zip"
    shutil.make_archive(zip_path.replace(".zip", ""), 'zip', folder_name)

    with open(zip_path, "rb") as zip_file:
        st.download_button(
            label=f"📥 Download {total_files} reports ({output_format})",
            data=zip_file,
            file_name=f"{folder_name}.zip",
            mime="application/zip"
        )

    progress_text.write(f"✅ All reports have been generated in {output_format} format! Total time: {int(time.time() - start_time)}s")

def process_row(presentation_path, row, df1, index, selected_columns, output_folder, output_format):
    """Procesa una fila y genera un archivo PPTX o PDF en Streamlit Cloud."""
    presentation = pptx.Presentation(presentation_path)

    for col_idx, col_name in enumerate(row.index):
        column_letter = chr(65 + col_idx)
        update_text_of_textbox(presentation, column_letter, row[col_name])

    file_name = get_filename_from_selection(row, selected_columns)
    pptx_path = os.path.join(output_folder, f"{file_name}.pptx")

    presentation.save(pptx_path)

    if output_format == "PDF":
        pdf_path = os.path.join(output_folder, f"{file_name}.pdf")
        convert_pptx_to_pdf(pptx_path, pdf_path)
        os.remove(pptx_path)
