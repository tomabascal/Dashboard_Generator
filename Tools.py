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
    """Busca y reemplaza texto dentro de las cajas de texto que tengan el formato {A}, {B}, etc., manteniendo el formato del PPTX."""
    pattern = rf"\{{{column_letter}\}}"

    for slide in presentation.slides:
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text:
                if re.search(pattern, shape.text):
                    text_frame = shape.text_frame
                    for paragraph in text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.text = re.sub(pattern, str(new_text), run.text)


def process_files(ppt_template_path, data_file, search_option, selected_columns, output_folder, output_format):
    """Procesa los archivos y genera las presentaciones."""
    df1 = pd.read_excel(data_file, engine='openpyxl')
    wb = openpyxl.load_workbook(data_file)
    sheet_name = wb.sheetnames[0]  # Ajusta según el nombre de tu hoja

    # Ajusta según tu lógica de selección
    df_selected = df1[df1['search_column'] == search_option]
    total_files = len(df_selected)

    estimated_time = total_files * (5 if output_format == "PDF" else 1)
    st.info(f"⏳ Estimated time: ~{estimated_time} seconds")

    progress_bar = st.progress(0)
    progress_text = st.empty()

    current_file = 0
    start_time = time.time()

    for index, row in df_selected.iterrows():
        process_row(ppt_template_path, row, df1, index, selected_columns,
                    output_folder, output_format, wb, sheet_name)
        current_file += 1
        progress = current_file / total_files
        progress_bar.progress(progress)
        elapsed_time = time.time() - start_time
        progress_text.write(f"📄 Generating {
                            current_file}/{total_files} ({output_format}) - Elapsed time: {int(elapsed_time)}s")

    zip_path = f"{output_folder}.zip"
    shutil.make_archive(zip_path.replace(".zip", ""), 'zip', output_folder)
    return zip_path


def process_row(presentation_path, row, df1, index, selected_columns, output_folder, output_format, wb, sheet_name):
    """Procesa una fila y genera un archivo PPTX o PDF en Streamlit Cloud."""
    presentation = pptx.Presentation(presentation_path)
    sheet = wb[sheet_name]

    for col_idx, col_name in enumerate(row.index):
        column_letter = chr(65 + col_idx)
        # Ajusta según el índice de fila y columna
        cell = sheet.cell(row=index + 2, column=col_idx + 1)
        formatted_text = format_cell_value(cell, df1.iloc[index, col_idx])
        update_text_of_textbox(presentation, column_letter, formatted_text)

    file_name = get_filename_from_selection(row, selected_columns)
    pptx_path = os.path.join(output_folder, f"{file_name}.pptx")

    presentation.save(pptx_path)

    if output_format == "PDF":
        pdf_path = os.path.join(output_folder, f"{file_name}.pdf")
        convert_pptx_to_pdf(pptx_path, pdf_path)
        os.remove(pptx_path)


def format_cell_value(cell, cell_value):
    """Formatea el valor de la celda según su tipo y formato."""
    if pd.isna(cell_value):
        return ""
    if isinstance(cell_value, (int, float)):
        if cell.number_format in ['0%', '0.00%']:
            return f"{cell_value * 100:.1f}%"
        if cell.number_format in ['Currency', 'Accounting']:
            return f"{cell_value:,.2f}€"
        return f"{cell_value:,.1f}"
    if isinstance(cell_value, pd.Timestamp):
        return cell_value.strftime("%d-%m-%Y")
    return str(cell_value)


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
        default=column_names[:1]
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
