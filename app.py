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
    layout="wide",
)

st.title("📊 Convertidor de Planogramas a Excel")
st.write(
    "Sube tu archivo PDF de implementación para extraer las tablas de productos "
    "de forma clara, completa y perfectamente estructurada en Excel."
)

st.sidebar.title("📌 Información")
st.sidebar.write("**Autor:** Alfredo HM")
st.sidebar.write("**Estado:** Listo para procesar")


# --- ALGORITMO DE EXTRACCIÓN POR BLOQUES DE ITEM Y COORDENADAS X ---
def extraer_tabla_perfecta(pdf_file):
    datos_procesados = []
    patron_ean = re.compile(r"\b\d{10,14}\b")

    with pdfplumber.open(pdf_file) as pdf:
        for pagina in pdf.pages:
            words = pagina.extract_words(
                x_tolerance=2, y_tolerance=2, keep_blank_chars=False
            )
            if not words:
                continue

            # 1. Agrupar palabras por altura vertical (Y)
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

            # 2. Encontrar cada ítem (fila principal donde inicia un EAN)
            bloques_items = []
            y_keys_ordenadas = sorted(lineas_dict.keys())

            for idx, y_pos in enumerate(y_keys_ordenadas):
                words_linea = sorted(
                    lineas_dict[y_pos], key=lambda item: item["x0"]
                )
                texto_linea = " ".join([w["text"] for w in words_linea])

                match_ean = patron_ean.search(texto_linea)
                if match_ean:
                    bloques_items.append((y_pos, match_ean.group(0)))

            if not bloques_items:
                continue

            ancho_pag = pagina.width

            # 3. Procesar cada ítem agrupando todas sus sublíneas verticales
            for i, (y_inicio, ean_val) in enumerate(bloques_items):
                y_fin = (
                    bloques_items[i + 1][0]
                    if i + 1 < len(bloques_items)
                    else y_inicio + 50.0
                )

                # Extraer todas las palabras dentro del rango vertical de este producto
                words_item = [
                    w
                    for y_k in y_keys_ordenadas
                    if y_inicio <= y_k < y_fin
                    for w in lineas_dict[y_k]
                ]

                # Clasificar palabras en las 12 columnas según su coordenada horizontal X
                # (Límites calibrados según la maquetación estándar del reporte)
                col_bandeja = []
                col_num = []
                col_ean = []
                col_nombre = []
                col_marca = []
                col_desc = []
                col_fab = []
                col_caras = []
                col_altura = []
                col_prof = []
                col_unid_band = []
                col_total_unid = []

                # Ordenar palabras de arriba a abajo y de izquierda a derecha
                words_item = sorted(
                    words_item, key=lambda w: (round(w["top"], 1), w["x0"])
                )

                for w in words_item:
                    x_rel = w["x0"] / ancho_pag
                    text = w["text"]

                    if x_rel < 0.07:
                        col_bandeja.append(text)
                    elif x_rel < 0.11:
                        if text.isdigit():
                            col_num.append(text)
                        else:
                            col_bandeja.append(text)
                    elif x_rel < 0.22:
                        if patron_ean.match(text) or text.isdigit():
                            col_ean.append(text)
                    elif x_rel < 0.35:
                        col_nombre.append(text)
                    elif x_rel < 0.45:
                        col_marca.append(text)
                    elif x_rel < 0.56:
                        col_desc.append(text)
                    elif x_rel < 0.70:
                        col_fab.append(text)
                    elif x_rel < 0.74:
                        col_caras.append(text)
                    elif x_rel < 0.79:
                        col_altura.append(text)
                    elif x_rel < 0.85:
                        col_prof.append(text)
                    elif x_rel < 0.92:
                        col_unid_band.append(text)
                    else:
                        col_total_unid.append(text)

                # Armar la fila final con los textos consolidados con espacios limpios
                row = [""] * 12
                row[0] = col_bandeja[0] if col_bandeja else ""
                row[1] = col_num[0] if col_num else ""
                row[2] = col_ean[0] if col_ean else ean_val
                row[3] = " ".join(col_nombre).strip()
                row[4] = " ".join(col_marca).strip()
                row[5] = " ".join(col_desc).strip()
                row[6] = " ".join(col_fab).strip()

                # Limpieza de valores numéricos finales
                row[7] = col_caras[0] if col_caras else "0"
                row[8] = col_altura[0] if col_altura else "0"
                row[9] = col_prof[0] if col_prof else "0"
                row[10] = col_unid_band[0] if col_unid_band else "0"
                row[11] = (
                    col_total_unid[0]
                    if col_total_unid
                    else (col_unid_band[0] if col_unid_band else "0")
                )

                # Ajustes de separación si alguna palabra rozó el límite entre Nombre y Marca
                # (Ejemplo: arreglar si 'MOUNTAIN' o 'ALTOMAYO' se coló al final del Nombre)
                if row[4] and row[3].endswith(row[4]):
                    row[3] = row[3][: -len(row[4])].strip()

                datos_procesados.append(row)

    return datos_procesados


# --- FUNCIÓN PARA GENERAR EL EXCEL CON ESTILOS IDÉNTICOS AL MODELO ---
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
            elif c_idx == 3:  # EAN formato texto estricto
                cell.value = str(val)
                cell.alignment, cell.number_format = align_center, "@"
            elif c_idx in [4, 5, 6, 7]:  # Nombre, Marca, Desc_A, Fabricante
                cell.value = str(val)
                cell.alignment = align_left
            elif c_idx in [8, 9, 10, 11]:  # Valores Numéricos
                cell.value = int(val) if str(val).isdigit() else 0
                cell.alignment, cell.number_format = align_right, "#,##0"
            elif c_idx == 12:  # Total_Unidades
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

    # --- PESTAÑA RESUMEN Y KPIS ---
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


# --- INTERFAZ STREAMLIT ---
uploaded_file = st.file_uploader("Arrastra tu PDF aquí", type=["pdf"])

if uploaded_file is not None:
    categoria = st.text_input("Nombre de la Categoría (opcional)", "General")

    if st.button("Procesar y Convertir a Excel"):
        with st.spinner("Procesando documento..."):
            datos = extraer_tabla_perfecta(uploaded_file)

            if datos:
                st.success(
                    f"¡Listo! Se extrajeron {len(datos)} filas con todas sus columnas estructuradas."
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

st.markdown("---")
st.write("Desarrollado por **Alfredo HM**")
