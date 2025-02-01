# -*- coding: utf-8 -*-
"""
Created on Fri Jan 31 23:02:57 2025

@author: betoh
"""

import streamlit as st
import pandas as pd
import pdfplumber
import re
from io import BytesIO

# --- Parámetros, patrones y constantes ---
PATRON_FECHA = re.compile(r'^\d{1,2}/[A-Z]{3}')
PATRON_MONTO = re.compile(r'(\d{1,3}(?:,\d{3})*\.\d{2})')

# Límites para clasificar transacciones (ajusta estos valores según tus necesidades)
LOWER_LIMIT = 417.3665  # Si el promedio de coordenadas es menor o igual a este valor => CARGO
UPPER_LIMIT = 424.75    # Si el promedio es mayor o igual a este valor => ABONO

# Mapa para convertir "02" a "FEB", "03" a "MAR", etc. (para nombrar la hoja)
MONTH_MAP = {
    "01": "ENE", "02": "FEB", "03": "MAR", "04": "ABR",
    "05": "MAY", "06": "JUN", "07": "JUL", "08": "AGO",
    "09": "SEP", "10": "OCT", "11": "NOV", "12": "DIC"
}

# --- Función para procesar un PDF y extraer datos en un DataFrame ---
def parsear_pdf_a_df(pdf_file, password=None):
    """
    Recibe un archivo PDF (tipo BytesIO o similar) y opcionalmente una contraseña.
    Devuelve:
      - Un DataFrame con columnas: Fecha, Tipo de movimiento, Monto, Descripcion y Categoría.
      - Un nombre sugerido para la hoja (por ejemplo, 'FEB_2024'), basado en la fecha de la 2ª página.
    """
    data = []
    sheet_name = None
    found_year = None

    # 1) Extraer la fecha de la 2ª página para el nombre de la hoja
    with pdfplumber.open(pdf_file, password=password) as pdf:
        if len(pdf.pages) >= 2:
            page2_text = pdf.pages[1].extract_text() or ""
            # Ejemplo: "DEL 13/02/2024 AL 12/03/2024"
            match_periodo = re.search(r'DEL\s+(\d{1,2}/\d{2}/\d{4})\s+AL', page2_text)
            if match_periodo:
                fecha_inicial = match_periodo.group(1)  # Ej: "13/02/2024"
                dd, mm, yyyy = fecha_inicial.split("/")
                found_year = yyyy
                mes_txt = MONTH_MAP.get(mm, mm)
                sheet_name = f"{mes_txt}_{yyyy}"
        if not sheet_name:
            sheet_name = "SIN_NOMBRE"
        if not found_year:
            found_year = "0000"

    # 2) Parsear las transacciones usando pdfplumber.extract_words()
    fecha_operacion = None
    tipo_movimiento = []
    monto = None
    x_monto = None
    descripcion_acumulada = []

    def guardar_registro():
        """Guarda el registro actual si está completo."""
        if fecha_operacion and tipo_movimiento and monto:
            desc = " ".join(descripcion_acumulada).strip()
            data.append({
                "Fecha": fecha_operacion,     # p.e. "13/FEB"
                "Tipo de movimiento": " ".join(tipo_movimiento).strip(),
                "Monto": monto,
                "x_monto": x_monto,          # Se usa para la clasificación
                "Descripcion": desc
            })

    with pdfplumber.open(pdf_file, password=password) as pdf:
        for page_index, page in enumerate(pdf.pages):
            words = page.extract_words()
            i = 0
            while i < len(words):
                w = words[i]
                txt = w["text"].strip()

                # 2.1) Si el token coincide con el formato de fecha ("13/FEB") => inicio de un nuevo movimiento
                if PATRON_FECHA.match(txt):
                    guardar_registro()
                    fecha_operacion = txt
                    tipo_movimiento = []
                    monto = None
                    x_monto = None
                    descripcion_acumulada = []

                    # Omitir todas las siguientes palabras que sean fechas (por ejemplo, fecha de liquidación)
                    j = i + 1
                    while j < len(words) and PATRON_FECHA.match(words[j]["text"].strip()):
                        j += 1

                    # Buscar el primer token que contenga un monto
                    temp_tipo = []
                    found_m = False
                    while j < len(words):
                        w2 = words[j]
                        txt2 = w2["text"].strip()
                        match_m = re.search(PATRON_MONTO, txt2)
                        if match_m:
                            monto = match_m.group(1)
                            # Calcula el promedio de x0 y x1 para una mejor determinación de posición
                            x0 = w2.get("x0", 0)
                            x1 = w2.get("x1", 0)
                            x_monto = (x0 + x1) / 2
                            found_m = True
                            tipo_movimiento = temp_tipo
                            i = j
                            break
                        else:
                            temp_tipo.append(txt2)
                            j += 1

                    if not found_m:
                        tipo_movimiento = temp_tipo
                        i = j
                else:
                    # 2.2) Para tokens que no inician un nuevo movimiento:
                    # Omitir tokens que empiecen con "Referencia" o que sean montos exactos
                    if re.match(r'(?i)^referencia', txt):
                        i += 1
                        continue
                    if re.fullmatch(PATRON_MONTO, txt):
                        i += 1
                        continue
                    # Agregar el token a la descripción
                    descripcion_acumulada.append(txt)
                    i += 1

    # Guardar el último registro
    guardar_registro()

    # 3) Crear el DataFrame y agregar el año a la fecha (por ejemplo, "13/FEB/2024")
    df = pd.DataFrame(data)
    if not df.empty:
        df["Fecha"] = df["Fecha"] + "/" + found_year
    else:
        df = pd.DataFrame(columns=["Fecha", "Tipo de movimiento", "Monto", "Descripcion", "x_monto"])

    # 4) Clasificar cada transacción según el promedio de coordenadas
    def clasificar_por_coordenadas(row):
        x_val = row["x_monto"] if pd.notnull(row["x_monto"]) else 0
        if x_val >= UPPER_LIMIT:
            return "ABONO"
        elif x_val <= LOWER_LIMIT:
            return "CARGO"
        else:
            # Si el promedio está en el rango intermedio, forzamos a "CARGO" (puedes ajustar esta lógica)
            return "CARGO"

    if not df.empty:
        df["Categoría"] = df.apply(clasificar_por_coordenadas, axis=1)
        df.drop(columns=["x_monto"], inplace=True, errors="ignore")

    return df, sheet_name

# --- Función principal de la aplicación Streamlit ---
def main():
    st.title("Lector de Estados de Cuenta")
    st.write("Sube uno o varios PDFs (o una carpeta en navegadores compatibles) y genera un Excel con una hoja por cada PDF.")

    # Permite seleccionar múltiples archivos PDF
    uploaded_files = st.file_uploader("Selecciona archivos PDF", type=["pdf"], accept_multiple_files=True)
    pdf_password = st.text_input("Contraseña para PDFs (opcional)", type="password")

    if st.button("Procesar y Descargar Excel"):
        if not uploaded_files:
            st.error("No se han subido archivos.")
        else:
            num_files = len(uploaded_files)
            progress_bar = st.progress(0)
            status_text = st.empty()

            output = BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                # Procesamos cada archivo y actualizamos la barra de progreso
                for i, pdf_file in enumerate(uploaded_files, start=1):
                    status_text.text(f"Procesando {pdf_file.name} ({i}/{num_files})")
                    try:
                        df, sheet_name = parsear_pdf_a_df(pdf_file, password=pdf_password or None)
                    except Exception as e:
                        st.error(f"Error al procesar {pdf_file.name}: {e}")
                        continue

                    original_sheet_name = sheet_name
                    counter = 1
                    # Evitar nombres de hoja duplicados
                    while sheet_name in writer.sheets:
                        sheet_name = f"{original_sheet_name}_{counter}"
                        counter += 1

                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    progress_bar.progress(i / num_files)
            output.seek(0)
            st.download_button(
                "Descargar Excel",
                data=output,
                file_name="estados_de_cuenta.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            status_text.text("¡Procesamiento completado!")

if __name__ == "__main__":
    main()
