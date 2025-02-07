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
from openpyxl import load_workbook



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


def process_files(ppt_file, excel_file, search_option, start_row, end_row, store_ids, selected_columns, output_format):
    """Genera reportes en formato PPTX o PDF en Streamlit Cloud respetando formatos del Excel."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    folder_name = f"Presentations_{timestamp}"
    os.makedirs(folder_name, exist_ok=True)

    temp_folder = "temp_files"
    os.makedirs(temp_folder, exist_ok=True)

    # Guardar archivos temporales
    ppt_template_path = os.path.join(temp_folder, ppt_file.name)
    excel_file_path = os.path.join(temp_folder, excel_file.name)

    with open(ppt_template_path, "wb") as f:
        f.write(ppt_file.getbuffer())
    with open(excel_file_path, "wb") as f:
        f.write(excel_file.getbuffer())

    # Leer el archivo Excel con pandas para filtrar datos
    try:
        with pd.ExcelFile(excel_file_path) as xls:
            df1 = pd.read_excel(xls, sheet_name=0)
    except PermissionError as e:
        st.error(f"Error reading Excel file: {e}")
        return

    # Ajustar los índices de las filas seleccionadas
    start_row_index = start_row - 1
    end_row_index = end_row - 1

    # Aplicar filtros según la opción seleccionada
    if search_option == 'rows':
        df_selected = df1.iloc[start_row_index:end_row_index + 1]
    elif search_option == 'store_id':
        store_id_list = [store_id.strip() for store_id in store_ids.split(',')]
        df_selected = df1[df1.iloc[:, 0].astype(str).isin(store_id_list)]
    else:
        df_selected = pd.DataFrame()

    total_files = len(df_selected)
    if total_files == 0:
        st.error("⚠️ No data found. Verify filters.")
        return

    st.info(f"⏳ Estimated time: ~{total_files} seconds")

    progress_bar = st.progress(0)
    progress_text = st.empty()

    # Verificación para evitar división por cero
    if total_files > 0:
        # Procesar las filas seleccionadas
        for index, row in df_selected.iterrows():
            # Actualizar la barra de progreso
            progress_value = (index + 1) / total_files
            progress_bar.progress(progress_value)
            progress_text.text(f"Processing file {index + 1} of {total_files}")

        st.success("✅ Processing complete!")
    else:
        st.error("⚠️ No data found for processing.")



def process_row(presentation_path, row, excel_file_path, index, selected_columns, output_folder, output_format):
    """Procesa una fila y genera un archivo PPTX o PDF respetando el formato original de Excel."""
    
    # Cargar la presentación de PowerPoint
    presentation = pptx.Presentation(presentation_path)

    # Cargar el archivo Excel con openpyxl para leer los formatos
    wb = load_workbook(excel_file_path, data_only=True)
    ws = wb.active  # Tomar la primera hoja del Excel

    for col_idx, col_name in enumerate(row.index):
        column_letter = chr(65 + col_idx)  # Convertir índice numérico en letra (A, B, C...)
        excel_cell = ws[f"{column_letter}{index + 2}"]  # Ajuste para coincidir con filas de Excel (base 1)

        # Obtener el valor formateado
        formatted_text = format_cell_value(excel_cell, wb, ws.title)
        
        # Reemplazar el texto en la presentación
        update_text_of_textbox(presentation, column_letter, formatted_text)

    # Generar el nombre del archivo basado en las columnas seleccionadas
    file_name = get_filename_from_selection(row, selected_columns)
    pptx_path = os.path.join(output_folder, f"{file_name}.pptx")

    # Guardar el archivo PPTX
    presentation.save(pptx_path)

    # Convertir a PDF si es necesario
    if output_format == "PDF":
        pdf_path = os.path.join(output_folder, f"{file_name}.pdf")
        convert_pptx_to_pdf(pptx_path, pdf_path)
        os.remove(pptx_path)  # Eliminar el PPTX después de la conversión


def format_cell_value(cell, wb, sheet_name):
    """Formatea y redondea el valor de la celda según su tipo y formato en Excel."""
    if cell is None or cell.value is None:
        return ""
    
    value = cell.value
    if isinstance(value, (int, float)):
        ws = wb[sheet_name]
        cell_format = ws[cell.coordinate].number_format

        # Limpiar caracteres extraños del formato (ej. \#,##0\ "€")
        cleaned_format = re.sub(r'[^\d.,%€$£]', '', cell_format)  

        # Identificar el símbolo de moneda si existe
        currency_symbol = next((symbol for symbol in ["€", "$", "£"] if symbol in cleaned_format), "")

        if currency_symbol:
            # Redondear a 1 decimal y eliminar el .0 si es entero
            rounded_value = round(value, 1)
            return f"{rounded_value:,.1f}".rstrip('0').rstrip('.') + f" {currency_symbol}"
        elif "%" in cleaned_format:
            # Redondear porcentaje a 1 decimal, pero nunca mostrar .0
            percentage = round(value * 100, 1)
            if percentage.is_integer():  # Si el porcentaje es un número entero
                return f"{int(percentage)}%"  # No mostrar decimales
            else:
                return f"{percentage:.1f}%"  # Mostrar un decimal
        else:
            # Redondear número normal a 1 decimal y eliminar el .0 si es entero
            rounded_value = round(value, 1)
            return f"{rounded_value:,.1f}".rstrip('0').rstrip('.')  # Redondeo de 1 decimal

    elif isinstance(value, datetime):
        return value.strftime("%d-%m-%Y")  # Formato de fecha

    return str(value)

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
