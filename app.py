import io
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pdfplumber
import streamlit as st

# --- CONFIGURACIÓN DE PÁGINA STREAMLIT ---
st.set_page_config(
    page_title="Convertidor de Planogramas a Excel",
    page_icon="📊",
    layout="centered",
)

# --- CABECERA Y TÍTULO ---
st.title("📊 Convertidor de Planogramas a Excel")
st.write(
    "Sube tu archivo PDF de implementación para generar automáticamente "
    "el reporte en Excel con todas las columnas organizadas y estructuradas."
)

# --- PANEL LATERAL ---
st.sidebar.title("📌 Información")
st.sidebar.write("**Autor:** Alfredo HM")
st.sidebar.write("**Estado:** Listo para procesar")


# --- FUNCIÓN DE EXTRACCIÓN DE DATOS ---
def extraer_tabla_exacta_stream(pdf_file):
    datos_procesados = []
    patron_ean = re.compile(r"\b\d{10,14}\b")

    with pdfplumber.open(pdf_file) as pdf:
        for pagina in pdf.pages:
            words = pagina.extract_words(
                x_tolerance=3, y_tolerance=3, keep_blank_chars=False
            )
            if not words:
                continue

            # Agrupar por líneas (coordenada Y)
            lineas_dict = {}
            for w in words:
                y_pos = round(w["top"], 1)
                linea_clave = None
                for y_existente in lineas_dict.keys():
                    if abs(y_existente - y_pos) <= 3.0:
                        linea_clave = y_existente
                        break

                if linea_clave is None:
                    linea_clave = y_pos
                    lineas_dict[linea_clave] = []

                lineas_dict[linea_clave].append(w)

            ancho_pag = pagina.width

            # Asignar palabras a columnas según su posición horizontal (X)
            for y_pos in sorted(lineas_dict.keys()):
                words_linea = sorted(
                    lineas_dict[y_pos], key=lambda item: item["x0"]
                )
                texto_linea = " ".join([w["text"] for w in words_linea])

                match_ean = patron_ean.search(texto_linea)
                if not match_ean:
                    continue

                col = [""] * 12
                txt_nombre, txt_marca, txt_desc, txt_fabricante = [], [], [], []

                for w in words_linea:
                    x0 = w["x0"]
                    text = w["text"]
                    x_rel = x0 / ancho_pag

                    if x_rel < 0.08:
                        if not col[0]:
                            col[0] = text
                    elif x_rel < 0.12:
                        if text.isdigit() and not col[1]:
                            col[1] = text
                        elif not col[0]:
                            col[0] = text
                    elif x_rel < 0.22:
                        if patron_ean.match(text):
                            col[2] = text
                        elif not col[2] and text.isdigit():
                            col[2] = text
                    elif x_rel < 0.42:
                        txt_nombre.append(text)
                    elif x_rel < 0.52:
                        txt_marca.append(text)
                    elif x_rel < 0.62:
                        txt_desc.append(text)
                    elif x_rel < 0.78:
                        txt_fabricante.append(text)
                    elif x_rel < 0.82:
                        if not col[7]:
                            col[7] = text
                    elif x_rel < 0.86:
                        if not col[8]:
                            col[8] = text
                    elif x_rel < 0.90:
                        if not col[9]:
                            col[9] = text
                    elif x_rel < 0.95:
                        if not col[10]:
                            col[10] = text
                    else:
                        if not col[11]:
                            col[11] = text

                col[3] = " ".join(txt_nombre).strip()
                col[4] = " ".join(txt_marca).strip()
                col[5] = " ".join(txt_desc).strip()
                col[6] = " ".join(txt_fabricante).strip()

                if not col[2]:
                    col[2] = match_ean.group(0)

                datos_procesados.append(col)

    return datos_procesados


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
        with st.spinner("Procesando documento..."):
            datos = extraer_tabla_exacta_stream(uploaded_file)

            if datos:
                st.success(
                    f"¡Listo! Se extrajeron {len(datos)} filas correctamente."
                )
                excel_bytes = generar_excel_en_memoria(datos, categoria)

                st.download_button(
                    label="Descargar Excel",
                    data=excel_bytes,
                    file_name=f"Reporte_{categoria}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.error("No se encontraron registros válidos en el PDF.")

# --- PIE DE PÁGINA SIMPLE ---
st.markdown("---")
st.write("Desarrollado por **Alfredo HM**")
