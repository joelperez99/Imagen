import streamlit as st
import requests
from io import BytesIO
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
import re

st.set_page_config(page_title="Insertar imágenes en Excel", layout="centered")

st.title("📸 Insertar imágenes en Excel desde URLs")
st.write(
    "Sube un archivo Excel que tenga una columna llamada "
    "**'Imágenes del anaquel AL LLEGAR'** con links de imágenes."
)
st.write("El sistema descargará las imágenes y las incrustará en sus celdas.")

uploaded_file = st.file_uploader("Subir archivo Excel", type=["xlsx"])

if uploaded_file is not None:
    # Botón para procesar
    if st.button("🔄 Procesar archivo e insertar imágenes"):
        with st.spinner("Procesando archivo, descargando imágenes..."):
            try:
                # Cargar workbook desde el archivo subido
                file_bytes = BytesIO(uploaded_file.read())
                wb = openpyxl.load_workbook(file_bytes)
                ws = wb.active  # puedes cambiar a wb["NombreHoja"] si quieres algo específico

                # Encontrar la columna que tenga el encabezado deseado
                header_name = "Imágenes del anaquel AL LLEGAR"
                image_col_idx = None

                for cell in ws[1]:  # primera fila = encabezados
                    if str(cell.value).strip() == header_name:
                        image_col_idx = cell.column
                        break

                if image_col_idx is None:
                    st.error(
                        f"No se encontró la columna con el encabezado: '{header_name}'.\n"
                        "Verifica que el nombre coincida exactamente."
                    )
                else:
                    max_row = ws.max_row
                    progress = st.progress(0)
                    processed = 0
                    total = max_row - 1  # sin el encabezado

                    # Recorremos filas de datos
                    for row in range(2, max_row + 1):
                        cell = ws.cell(row=row, column=image_col_idx)
                        url_text = str(cell.value).strip() if cell.value is not None else ""

                        if url_text:
                            # Buscar la primera URL en el texto (por si hay ' --- ')
                            urls = re.findall(r"https?://\S+", url_text)
                            if urls:
                                img_url = urls[0]

                                try:
                                    resp = requests.get(img_url, timeout=10)
                                    if resp.status_code == 200:
                                        img_bytes = BytesIO(resp.content)
                                        img = XLImage(img_bytes)

                                        # Ajustar tamaño de la miniatura (opcional)
                                        img.width = 100
                                        img.height = 100

                                        # Anclar imagen a la celda correspondiente
                                        col_letter = get_column_letter(image_col_idx)
                                        anchor_cell = f"{col_letter}{row}"
                                        img.anchor = anchor_cell
                                        ws.add_image(img)

                                        # Ajustar altura de la fila (opcional)
                                        ws.row_dimensions[row].height = 80
                                    else:
                                        # Si falla la descarga, simplemente se salta
                                        pass
                                except Exception:
                                    # Si hay cualquier error con esa URL, se ignora
                                    pass

                        processed += 1
                        progress.progress(processed / total)

                    # Guardar workbook en memoria
                    output = BytesIO()
                    wb.save(output)
                    output.seek(0)

                    st.success("¡Proceso terminado! Ya puedes descargar tu archivo con imágenes.")

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
