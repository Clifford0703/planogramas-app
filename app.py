import base64
import io
import json
import os
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pypdfium2 as pdfium
import streamlit as st
from groq import Groq

# --- CONFIGURACIÓN DE PÁGINA STREAMLIT ---
st.set_page_config(
    page_title="Convertidor de Planogramas a Excel",
    page_icon="⚡",
    layout="centered",
)

st.title("⚡ Convertidor de Planogramas con Visión AI (Groq)")
st.write(
    "Utiliza **Groq Cloud (Llama 3.2 Vision)** para procesar visualmente "
    "cada página del PDF de forma ultra rápida y sin omitir ninguna columna."
)

st.sidebar.title("📌 Información")
st.sidebar.write("**Autor:** Alfredo HM")
st.sidebar.write("**Motor:** Groq Llama 3.2 Vision")

# --- CONFIGURACIÓN DE GROQ API ---
groq_api_key = st.secrets.get(
    "GROQ_API_KEY", os.environ.get("GROQ_API_KEY", "")
)

if not groq_api_key:
    groq_api_key = st.text_input("Ingresa tu GROQ_API_KEY:", type="password")

client = None
if groq_api_key:
    client = Groq(api_key=groq_api_key)


# --- FUNCIÓN DE EXTRACCIÓN CON GROQ VISION ---
def extraer_tablas_con_groq(pdf_bytes):
    pdf = pdfium.PdfDocument(pdf_bytes)
    todas_las_filas = []

    prompt = """
    Analiza esta página de planograma/reporte de implementación de retail.
    Extrae ÚNICAMENTE la tabla de productos detallada.
    Ignora encabezados generales, títulos de módulos, gráficos de pie/barras e imágenes del mueble.
    
    Devuelve un JSON estrictamente con la siguiente estructura (una lista de listas de cadenas de texto):
    {
        "filas": [
            ["Bandeja", "N°", "EAN", "Nombre", "Marca", "Desc_A", "Fabricante", "Caras", "Altura", "Profundidad", "Total Unid en Bandeja", "Total_Unidades"],
            ...
        ]
    }
    Instrucciones estrictas:
    1. Asegúrate de capturar TODAS las líneas que contengan un código EAN (numérico de 10-14 dígitos).
    2. Mantén exactamente las 12 columnas.
    3. Si una celda tiene un asterisco '*', consérvalo.
    """

    total_paginas = len(pdf)
    barra_progreso = st.progress(0)

    for i, page in enumerate(pdf):
        st.write(f"⚡ Escaneando página {i+1} de {total_paginas} con Groq...")

        # Renderizar página a imagen PNG en memoria y convertir a base64
        image = page.render(scale=2).to_pil()
        img_byte_arr = io.BytesIO()
        image.save(img_byte_arr, format="PNG")
        base64_image = base64.b64encode(img_byte_arr.getvalue()).decode("utf-8")

        try:
            # Llamada síncrona a Groq Vision (Ultra Rápida)
            response = client.chat.completions.create(
                model="llama-3.2-11b-vision-preview",
                response_format={"type": "json_object"},
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": prompt},
                            {
                                "type": "image_url",
                                "image_url": {
                                    "url": f"data:image/png;base64,{base64_image}"
                                },
                            },
                        ],
                    }
                ],
            )

            res_json = json.loads(response.choices[0].message.content)
            filas_pag = res_json.get("filas", [])

            for f in filas_pag:
                if f and f[0] not in ["Bandeja", "Reporte de Implementación"]:
                    todas_las_filas.append(f)

        except Exception as e:
            st.error(f"Error en página {i+1}: {e}")

        barra_progreso.progress((i + 1) / total_paginas)

    return todas_las_filas


# --- FUNCIÓN DE GENERACIÓN DE EXCEL ---
def generar_excel_en_memoria(datos_filas, titulo_categoria):
    wb = openpyxl.Workbook()

    ws_summary = wb.active
    ws_summary.title = "Resumen y KPIs"
    ws_data = wb.create_sheet(title="Reporte_Implementacion")

    ws_summary.views.sheetView[0].showGridLines = True
    ws_data.views.sheetView[0].showGridLines = True

    font_title = Font(name="Calibri", size=16, bold=True, color="1F497D")
    font_subtitle = Font(name="Calibri", size=11, italic=True, color="595959")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_body = Font(name="Calibri", size=10)
    font_bold = Font(name="Calibri", size=10, bold=True)

    fill_header = PatternFill(
        start_color="1F497D", end_color="1F497D", fill_type="solid"
    )
    fill_zebra = PatternFill(
        start_color="F2F5F9", end_color="F2F5F9", fill_type="solid"
    )
    fill_kpi = PatternFill(
        start_color="DCE6F1", end_color="DCE6F1", fill_type="solid"
    )

    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )
    header_border = Border(
        left=Side(style="thin", color="1F497D"),
        right=Side(style="thin", color="1F497D"),
        top=Side(style="thin", color="1F497D"),
        bottom=Side(style="medium", color="1F497D"),
    )

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    headers = [
        "Bandeja",
        "N°",
        "EAN",
        "Nombre",
        "Marca",
        "Desc_A",
        "Fabricante",
        "Caras",
        "Altura",
        "Profundidad",
        "Total Unid en Bandeja",
        "Total_Unidades",
    ]

    ws_data.cell(row=1, column=1, value="METRO HIPER MEJORADO").font = font_title
    ws_data.cell(
        row=2,
        column=1,
        value=f"Reporte de Implementación - Categoría: {titulo_categoria}",
    ).font = font_subtitle

    start_row = 4
    for col_idx, header in enumerate(headers, 1):
        cell = ws_data.cell(row=start_row, column=col_idx, value=header)
        (
            cell.font,
            cell.fill,
            cell.alignment,
            cell.border,
        ) = (font_header, fill_header, align_center, header_border)

    for r_idx, row_data in enumerate(datos_filas, start_row + 1):
        row_fill = fill_zebra if (r_idx % 2 == 0) else None

        while len(row_data) < 12:
            row_data.append("")

        for c_idx, val in enumerate(row_data[:12], 1):
            cell = ws_data.cell(row=r_idx, column=c_idx)
            cell.font, cell.border = font_body, thin_border
            if row_fill:
                cell.fill = row_fill

            if c_idx in [1, 2]:
                cell.value = str(val) if val != "" else ""
                cell.alignment = align_center
            elif c_idx == 3:
                cell.value = str(val)
                cell.alignment, cell.number_format = align_center, "@"
            elif c_idx in [4, 5, 6, 7]:
                cell.value = str(val)
                cell.alignment = align_left
            elif c_idx in [8, 9, 10, 11]:
                cell.value = int(val) if str(val).isdigit() else 0
                cell.alignment, cell.number_format = align_right, "#,##0"
            elif c_idx == 12:
                if str(val) == "*":
                    cell.value, cell.alignment = "*", align_center
                else:
                    cell.value = int(val) if str(val).isdigit() else val
                    cell.alignment, cell.number_format = align_right, "#,##0"

    tot_row = len(datos_filas) + start_row + 1
    ws_data.cell(row=tot_row, column=4, value="TOTAL GENERAL").font = font_bold
    ws_data.cell(row=tot_row, column=8, value=f"=SUM(H5:H{tot_row-1})").font = (
        font_bold
    )
    ws_data.cell(row=tot_row, column=11, value=f"=SUM(K5:K{tot_row-1})").font = (
        font_bold
    )

    for col_i in range(1, len(headers) + 1):
        c = ws_data.cell(row=tot_row, column=col_i)
        c.border = Border(
            top=Side(style="thin", color="1F497D"),
            bottom=Side(style="double", color="1F497D"),
        )
        if col_i in [8, 11]:
            c.number_format, c.alignment = "#,##0", align_right

    for col in ws_data.columns:
        col_letter = get_column_letter(col[0].column)
        ws_data.column_dimensions[col_letter].width = 16

    ws_data.column_dimensions["D"].width = 45
    ws_data.column_dimensions["G"].width = 35
    ws_data.freeze_panes = "A5"

    # PESTAÑA RESUMEN
    ws_summary.cell(
        row=1,
        column=1,
        value=f"DASHBOARD CATEGORÍA {titulo_categoria.upper()}",
    ).font = font_title
    ws_summary.cell(
        row=2, column=1, value="Resumen Métrico y Participación"
    ).font = font_subtitle

    kpis = [
        (
            "Total SKUs Registrados",
            f"=COUNTA(Reporte_Implementacion!C5:C{tot_row-1})",
            "B4",
            "C4",
            "B5",
            "C5",
        ),
        (
            "Total Caras Exhibidas",
            f"=SUM(Reporte_Implementacion!H5:H{tot_row-1})",
            "E4",
            "F4",
            "E5",
            "F5",
        ),
        (
            "Total Unidades en Bandeja",
            f"=SUM(Reporte_Implementacion!K5:K{tot_row-1})",
            "H4",
            "I4",
            "H5",
            "I5",
        ),
    ]

    for label, formula, top_l, top_r, bot_l, bot_r in kpis:
        ws_summary.merge_cells(f"{top_l}:{top_r}")
        ws_summary.merge_cells(f"{bot_l}:{bot_r}")

        c_lbl = ws_summary[top_l]
        (
            c_lbl.value,
            c_lbl.font,
            c_lbl.alignment,
            c_lbl.fill,
        ) = (
            label,
            Font(name="Calibri", size=10, bold=True, color="595959"),
            align_center,
            fill_kpi,
        )

        c_val = ws_summary[bot_l]
        (
            c_val.value,
            c_val.font,
            c_val.alignment,
            c_val.fill,
            c_val.number_format,
        ) = (
            formula,
            Font(name="Calibri", size=16, bold=True, color="1F497D"),
            align_center,
            fill_kpi,
            "#,##0",
        )

    # Resumen por Marca
    ws_summary.cell(row=8, column=2, value="Resumen por Marca").font = Font(
        name="Calibri", size=12, bold=True, color="1F497D"
    )
    s2_headers = ["Marca", "Caras Total", "Unidades en Bandeja", "% Caras"]
    for col_idx, h in enumerate(s2_headers, 2):
        c = ws_summary.cell(row=9, column=col_idx, value=h)
        c.font, c.fill, c.alignment = font_header, fill_header, align_center

    marcas_set = list(
        dict.fromkeys(
            [
                r[4]
                for r in datos_filas
                if len(r) > 4 and str(r[4]).strip() != ""
            ]
        )
    )

    for idx, m_name in enumerate(marcas_set, 10):
        ws_summary.cell(row=idx, column=2, value=m_name).font = font_body
        ws_summary.cell(
            row=idx,
            column=3,
            value=f"=SUMIF(Reporte_Implementacion!E$5:E${tot_row-1}, B{idx}, Reporte_Implementacion!H$5:H${tot_row-1})",
        ).font = font_body
        ws_summary.cell(
            row=idx,
            column=4,
            value=f"=SUMIF(Reporte_Implementacion!E$5:E${tot_row-1}, B{idx}, Reporte_Implementacion!K$5:K${tot_row-1})",
        ).font = font_body
        ws_summary.cell(
            row=idx, column=5, value=f"=C{idx}/C${10+len(marcas_set)}"
        ).font = font_body

        ws_summary.cell(row=idx, column=2).alignment = align_left
        ws_summary.cell(row=idx, column=3).alignment = align_right
        ws_summary.cell(row=idx, column=4).alignment = align_right
        ws_summary.cell(row=idx, column=5).alignment = align_right

        ws_summary.cell(row=idx, column=3).number_format = "#,##0"
        ws_summary.cell(row=idx, column=4).number_format = "#,##0"
        ws_summary.cell(row=idx, column=5).number_format = "0.00%"

    tot_m_row = 10 + len(marcas_set)
    ws_summary.cell(row=tot_m_row, column=2, value="Total").font = font_bold
    ws_summary.cell(
        row=tot_m_row, column=3, value=f"=SUM(C10:C{tot_m_row-1})"
    ).font = font_bold
    ws_summary.cell(
        row=tot_m_row, column=4, value=f"=SUM(D10:D{tot_m_row-1})"
    ).font = font_bold
    ws_summary.cell(
        row=tot_m_row, column=5, value=f"=SUM(E10:E{tot_m_row-1})"
    ).font = font_bold

    for c_i in range(2, 6):
        cell = ws_summary.cell(row=tot_m_row, column=c_i)
        cell.border = Border(
            top=Side(style="thin", color="1F497D"),
            bottom=Side(style="double", color="1F497D"),
        )
        if c_i in [3, 4]:
            cell.number_format, cell.alignment = "#,##0", align_right
        elif c_i == 5:
            cell.number_format, cell.alignment = "0.00%", align_right

    ws_summary.column_dimensions["B"].width = 28
    ws_summary.column_dimensions["C"].width = 16
    ws_summary.column_dimensions["D"].width = 22
    ws_summary.column_dimensions["E"].width = 16

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# --- INTERFAZ USUARIO ---
uploaded_file = st.file_uploader("Arrastra tu PDF aquí", type=["pdf"])

if uploaded_file is not None:
    categoria = st.text_input("Nombre de la Categoría (opcional)", "General")

    if st.button("Procesar y Convertir a Excel"):
        if not client:
            st.error(
                "Configura tu GROQ_API_KEY en los Secrets de Streamlit para continuar."
            )
        else:
            with st.spinner("Escaneando imágenes con la IA de Groq..."):
                pdf_bytes = uploaded_file.read()
                datos = extraer_tablas_con_groq(pdf_bytes)

                if datos:
                    st.success(
                        f"¡Éxito! Se procesaron {len(datos)} registros de forma impecable."
                    )
                    excel_bytes = generar_excel_en_memoria(datos, categoria)

                    st.download_button(
                        label="Descargar Excel",
                        data=excel_bytes,
                        file_name=f"Reporte_{categoria}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                else:
                    st.error("No se pudieron extraer los datos.")

st.markdown("---")
st.write("Desarrollado por **Alfredo HM**")
