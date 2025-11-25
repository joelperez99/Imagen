import streamlit as st
import requests
from io import BytesIO
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
import re

st.set_page_config(page_title="Insertar imágenes en Excel", layout="centered")

st.title("📸 Insertar imágenes en Excel desde cualquier celda")
st.write(
    "Sube un archivo Excel (.xlsx). El sistema buscará **en todas las celdas** "
    "texto que contenga URLs (http/https) de imágenes. "
    "Por cada URL encontrada descargará la imagen y la insertará en la hoja.\n\n"
    "Si una celda tiene varias URLs, las imágenes se colocan en la misma fila, "
    "empezando en la columna de esa celda y moviéndose a la derecha (ej. F, G, H…)."
)

uploaded_file = st.file_uploader("Subir archivo Excel", type=["xlsx"])

if uploaded_file is not None:
    if st.button("🔄 Procesar archivo e insertar imágenes"):
        with st.spinner("Procesando archivo, descargando imágenes..."):
            try:
                # Cargar workbook desde el archivo subido
                file_bytes = BytesIO(uploaded_file.read())
                wb = openpyxl.load_workbook(file_bytes)
                ws = wb.active  # Cambia a wb["NombreHoja"] si quieres una hoja específica

                max_row = ws.max_row
                max_col = ws.max_column

                # Para la barra de progreso usamos número de filas de datos (sin encabezado)
                total_rows = max_row - 1 if max_row > 1 else 1
                progress = st.progress(0)
                processed_rows = 0

                # Recorremos desde la fila 2 (asumiendo fila 1 como encabezados;
                # si no tienes encabezados, cambia a range(1, max_row+1))
                for row in range(2, max_row + 1):
                    row_touched = False  # para saber si ajustamos altura

                    for col in range(1, max_col + 1):
                        cell = ws.cell(row=row, column=col)
                        text = str(cell.value).strip() if cell.value is not None else ""

                        if not text:
                            continue

                        # Extraer TODAS las posibles URLs en el texto de la celda
                        urls = re.findall(r"https?://\S+", text)

                        if not urls:
                            continue

                        # Para cada URL de esa celda, insertamos una imagen
                        for idx, img_url in enumerate(urls):
                            try:
                                resp = requests.get(img_url, timeout=10)
                                if resp.status_code == 200:
                                    img_bytes = BytesIO(resp.content)
                                    img = XLImage(img_bytes)

                                    # Tamaño de miniatura
                                    img.width = 100
                                    img.height = 100

                                    # Columna destino: la de la celda + índice de la URL
                                    col_idx = col + idx
                                    col_letter = get_column_letter(col_idx)
                                    anchor_cell = f"{col_letter}{row}"

                                    img.anchor = anchor_cell
                                    ws.add_image(img)

                                    # Ajustar ancho de la columna (solo si está muy chica)
                                    current_width = ws.column_dimensions[col_letter].width
                                    if current_width is None or current_width < 18:
                                        ws.column_dimensions[col_letter].width = 18

                                    row_touched = True
                                else:
                                    # Si falla la descarga, saltamos esa URL
                                    pass
                            except Exception:
                                # Si algo truena con esa URL, seguimos con las demás
                                pass

                    # Si en esa fila agregamos al menos una imagen, ajustamos la altura
                    if row_touched:
                        if ws.row_dimensions[row].height is None or ws.row_dimensions[row].height < 80:
                            ws.row_dimensions[row].height = 80

                    processed_rows += 1
                    progress.progress(processed_rows / total_rows)

                # Guardar el workbook en memoria
                output = BytesIO()
                wb.save(output)
                output.seek(0)

                st.success("¡Proceso terminado! Ya puedes descargar tu archivo con las imágenes insertadas.")
                st.download_button(
                    label="⬇️ Descargar Excel con imágenes",
                    data=output,
                    file_name="excel_con_imagenes.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

            except Exception as e:
                st.error(f"Ocurrió un error procesando el archivo: {e}")
else:
    st.info("Sube un archivo Excel para comenzar.")
