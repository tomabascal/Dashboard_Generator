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
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows

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
        wb = openpyxl.load_workbook(excel_file_path, data_only=True)
        ws = wb.active

        # Lee los valores formateados de las celdas
        data = []
        for row in ws.iter_rows(min_row=start_row, max_row=end_row):
            formatted_row = []
            for cell in row:
                if cell.number_format == 'General':
                    formatted_row.append(cell.value)
                else:
                    formatted_row.append(cell.value if cell.value is None else cell.number_format.format(cell.value))
            data.append(formatted_row)

        columns = [cell.value for cell in ws[1]]
        df1 = pd.DataFrame(data, columns=columns)
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


import re

def update_text_of_textbox(presentation, column_letter, text):
    pattern = r"\b{}\b".format(re.escape(column_letter))  # Asegúrate de que el patrón esté correctamente escapado
    for slide in presentation.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                if re.search(pattern, shape.text):
                    shape.text = text
                    return


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


# ========= 💡 Estilos para mejorar el diseño =========
st.markdown("""
    <style>
    div.stButton > button {
        width: 100%;
        height: 50px;
        font-size: 16px;
        border-radius: 10px;
        font-weight: bold;
    }
    </style>
""", unsafe_allow_html=True)

# ========= Título =========
st.title("Shopfully Dashboard Generator")

# Opción para elegir el formato de salida
st.markdown("### **Select Output Format**")
output_format = st.radio("Choose the file format:", ["PPTX", "PDF"])

# Mensaje de advertencia si el usuario elige PDF
if output_format == "PDF":
    st.warning(
        "⚠️ Converting to PDF may take extra time. Large batches of presentations might take several minutes.")

# ========= 📂 Upload de archivos con formato mejorado =========
st.markdown(
    "**Upload PPTX Template**  \n*(Text Box format that will be edited -> {Column Letter} For Example: `{A}`)*", unsafe_allow_html=True)
ppt_template = st.file_uploader("", type=["pptx"])

st.write("")  # Espaciado

st.markdown(
    "**Upload Excel File**  \n*(Column A must be `Store ID`)*", unsafe_allow_html=True)
data_file = st.file_uploader("", type=["xlsx"])

# ========= 🔍 Botones mejorados para "Search by" =========
st.markdown("### **Search by:**")  # Título en negrita y más grande
col1, col2 = st.columns(2)  # Dos columnas para alinear botones en mosaico

# Inicializar la variable de estado para la selección del filtro
if "search_option" not in st.session_state:
    st.session_state.search_option = "rows"  # Valor por defecto

# Botón 1 - Search by Rows
with col1:
    if st.button("🔢 Rows", use_container_width=True):
        st.session_state.search_option = "rows"

# Botón 2 - Search by Store ID
with col2:
    if st.button("🔍 Store ID", use_container_width=True):
        st.session_state.search_option = "store_id"

# Mostrar la opción seleccionada
st.markdown(f"**Selected: `{st.session_state.search_option}`**")

# ========= 🔢 Inputs para definir el rango de búsqueda =========
start_row, end_row, store_ids = None, None, None

if st.session_state.search_option == "rows":
    start_row = st.number_input("Start Row", min_value=0, step=1)
    end_row = st.number_input("End Row", min_value=0, step=1)

elif st.session_state.search_option == "store_id":
    store_ids = st.text_input("Enter Store IDs (comma-separated)")

# ========= 📝 Selección de columnas para el nombre del archivo =========
if data_file is not None:
    # Leer la primera hoja del Excel
    df = pd.read_excel(data_file, sheet_name=0)
    column_names = df.columns.tolist()

    selected_columns = st.multiselect(
        "📂 Select and order the columns for the file name:",
        column_names,
        default=column_names[:3]
    )

    def get_filename_from_selection(row, selected_columns):
        """Genera el nombre del archivo según las columnas seleccionadas."""
        file_name_parts = [str(row[col])
                           for col in selected_columns if col in row]
        return "_".join(file_name_parts)

    st.write("🔹 Example file name:", get_filename_from_selection(
        df.iloc[0], selected_columns))

# ========= 🚀 Botón de procesamiento =========
if st.button("Process"):
    if ppt_template and data_file:
        process_files(ppt_template, data_file, st.session_state.search_option,
                      start_row, end_row, store_ids, selected_columns, output_format)
    else:
        st.error("Please upload both files before processing.")
