import re
import io
from datetime import datetime
import pandas as pd
import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference

# ==============================================================================
# CONFIGURACIÓN DE PÁGINA (STREAMLIT - CON / C5 PANAMÁ)
# ==============================================================================
st.set_page_config(
    page_title="Procesador Operación Vértice | CON-C5",
    page_icon="🛡️",
    layout="wide"
)

st.title("🛡️ Consola de Consolidación Operativa — Operación Vértice (CON-C5)")
st.markdown(
    "Pega los reportes operativos crudos (WhatsApp / Telegram / PDF) para clasificar automáticamente por "
    "**Bloque Horario**, **Institución (PN / SENAFRONT)**, **Punto de Control (BO)**, depurar reenvíos duplicados "
    "y exportar el libro de **Excel ejecutivo de 5 pestañas**."
)

# ==============================================================================
# REGLAS OPERATIVAS DE CLASIFICACIÓN (PANAMÁ — MINSEG / CON)
# ==============================================================================
SENAFRONT_BOS = {
    "BO-4", "BO-5", "BO-6", "BO-11", "BO-15", "BO-53", "BO-54", "BO-59", "BO-60", "BO-61"
}

def detectar_institucion(bo_code: str, texto: str) -> str:
    t_up = texto.upper()
    if "SENAFRONT" in t_up or "CENTINELAS DE LA PATRIA" in t_up or "BRIGADA" in t_up or "BRIOCC" in t_up:
        return "SENAFRONT"
    if "POLICIA NACIONAL" in t_up or "POLICÍA NACIONAL" in t_up or "DIOS Y PATRIA" in t_up or "ZONA DE POLIC" in t_up:
        return "Policía Nacional"
    if bo_code in SENAFRONT_BOS:
        return "SENAFRONT"
    return "Policía Nacional"

def detectar_bloque_horario(texto: str) -> str:
    # Busca formato militar DDHHMMNOV26 (ej. 181236NOV26 o 190150NOV26)
    m_mil = re.search(r"\b(18|19)(\d{2})(\d{2})NOV26\b", texto.upper())
    if m_mil:
        dia, hh, mm = m_mil.group(1), int(m_mil.group(2)), int(m_mil.group(3))
        mins = hh * 60 + mm
        if dia == "19" and mins <= 360:
            return "Bloque 4 (00:00-06:00)"
        if mins < 720:
            return "Bloque 1 (06:00-11:59)"
        elif mins <= 1045:  # hasta 17:25
            return "Bloque 2 (12:00-17:25)"
        else:
            return "Bloque 3 (17:26-23:59)"

    # Busca formato HH:MM
    m_hm = re.search(r"(?:HORA[S]?:?|A LAS)\s*(\d{1,2}):(\d{2})", texto.upper())
    if m_hm:
        hh, mm = int(m_hm.group(1)), int(m_hm.group(2))
        if "P.M." in texto.upper() and hh < 12:
            hh += 12
        mins = hh * 60 + mm
        if mins <= 360:
            return "Bloque 4 (00:00-06:00)"
        elif mins < 720:
            return "Bloque 1 (06:00-11:59)"
        elif mins <= 1045:
            return "Bloque 2 (12:00-17:25)"
        else:
            return "Bloque 3 (17:26-23:59)"

    if "BUENOS DIAS" in texto.upper() or "BUENOS DÍAS" in texto.upper():
        return "Bloque 4 (00:00-06:00)"
    if "BUENAS NOCHES" in texto.upper():
        return "Bloque 3 (17:26-23:59)"
    return "Bloque 2 (12:00-17:25)"

def parsear_reportes(texto_crudo: str):
    # Divide bloques que empiezan con BO-XX-
    bloques = re.split(r"(?=\bBO-\d+-)", texto_crudo.strip())
    registros = []
    vistos = set()

    for b in bloques:
        b = b.strip()
        if not b or not b.startswith("BO-"):
            continue

        # Huella de deduplicación (normaliza espacios y símbolos)
        huella = re.sub(r"\s+", " ", b.lower()).strip()
        if huella in vistos:
            continue
        vistos.add(huella)

        lineas = [l.strip() for l in b.splitlines() if l.strip()]
        encabezado = lineas[0]
        m_bo = re.match(r"^(BO-\d+)\s*-\s*([^\n]+)", encabezado)
        bo_code = m_bo.group(1) if m_bo else "BO-S/N"
        punto_control = encabezado

        cuerpo = "\n".join(lineas[1:])
        institucion = detectar_institucion(bo_code, b)
        bloque = detectar_bloque_horario(b)
        b_up = b.upper()

        # Clasificación de evento
        if "CONDUCCIÓN JUDICIAL" in b_up or "CONDUCCION JUDICIAL" in b_up or "CAPTURA" in b_up or "OFICIO" in b_up or "CINTILLO ROJO" in b_up:
            categoria = "Solicitado / Oficio"
        elif "INDOCUMENTADO" in b_up or "ESTADÍA VENCIDA" in b_up or "ESTADIA VENCIDA" in b_up or "ACTA DE INSPECCIÓN SANITARIA" in b_up or "CITACIÓN" in b_up:
            categoria = "Falta Administrativa / Identidad"
        elif "INFRACCIÓN" in b_up or "INFRACCION" in b_up or "BOLETA" in b_up or "GRÚA" in b_up or "GRUA" in b_up:
            categoria = "Infracción de Tránsito"
        else:
            categoria = "Verificación / Control"

        # Extracción numérica de personas y vehículos
        personas = 0
        vehiculos = 0
        boletas = 0

        m_pers = re.findall(r"(?:PERSONAS(?:\s+VERIFICADAS)?|CIUDADANOS(?:\s+NACIONALES|\s+EXTRANJEROS)?)\s*[:=.-]*\s*(\d+)", b_up)
        if m_pers:
            personas = sum(int(x) for x in m_pers)
        elif "PERSONA VERIFICADA" in b_up or "CIUDADANO VERIFICADO" in b_up or categoria == "Solicitado / Oficio":
            personas = 1

        m_veh = re.findall(r"(?:VEH[ÍI]CULOS(?:\s+PARTICULARES|\s+NACIONALES|\s+EXTRANJEROS)?|TAXIS?|BUSES?|MOTOS?|CAMI[ÓO]N(?:ES)?|VOLQUETES?|ARTICULADOS?|MULAS?)\s*[:=.-]*\s*(\d+)", b_up)
        if m_veh:
            vehiculos = sum(int(x) for x in m_veh)
        elif "PLACA" in b_up or "MATRÍCULA" in b_up or "MATRICULA" in b_up:
            vehiculos = 1

        m_bol = re.search(r"(\d+)\s*(?:BOLETAS|INFRACCIONES)", b_up)
        if m_bol:
            boletas = int(m_bol.group(1))
        elif categoria == "Infracción de Tránsito":
            boletas = 1

        # Extracción de cédula/pasaporte/placa
        m_doc = re.search(r"(?:C\.I\.P\.?|C[ÉE]DULA|PASAPORTE|RUEX)\s*[:=.-]*\s*([A-Z0-9-]+)", b_up)
        doc_id = m_doc.group(1) if m_doc else "N/A"

        m_placa = re.search(r"(?:PLACA|MATR[ÍI]CULA|MATRICULA)\s*[:=.-]*\s*([A-Z0-9-]+)", b_up)
        placa = m_placa.group(1) if m_placa else "N/A"

        registros.append({
            "Bloque Horario": bloque,
            "Institución": institucion,
            "Código BO": bo_code,
            "Punto de Control": punto_control,
            "Categoría": categoria,
            "Personas": personas,
            "Vehículos": vehiculos,
            "Boletas": boletas,
            "Documento / C.I.P.": doc_id,
            "Placa": placa,
            "Detalle Operativo": cuerpo[:350]
        })

    return pd.DataFrame(registros)

def generar_excel_operacion_vertice(df: pd.DataFrame) -> bytes:
    wb = openpyxl.Workbook()

    NAVY_HEADER = "1B365D"
    SLATE_SUBHEADER = "2C5282"
    ZEBRA_FILL = "F4F7FA"
    TOTAL_FILL = "E2E8F0"
    BORDER_COLOR = "CBD5E0"

    font_title = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    font_subtitle = Font(name="Calibri", size=10, italic=True, color="E2E8F0")
    font_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_regular = Font(name="Calibri", size=10, color="1A202C")
    font_bold = Font(name="Calibri", size=10, bold=True, color="1A202C")
    font_total = Font(name="Calibri", size=11, bold=True, color="1B365D")

    fill_title = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    fill_header = PatternFill(start_color=SLATE_SUBHEADER, end_color=SLATE_SUBHEADER, fill_type="solid")
    fill_zebra = PatternFill(start_color=ZEBRA_FILL, end_color=ZEBRA_FILL, fill_type="solid")
    fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    fill_total = PatternFill(start_color=TOTAL_FILL, end_color=TOTAL_FILL, fill_type="solid")

    thin_side = Side(style="thin", color=BORDER_COLOR)
    border_data = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    border_total = Border(left=thin_side, right=thin_side, top=Side(style="thin", color=NAVY_HEADER), bottom=Side(style="double", color=NAVY_HEADER))

    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def banner(ws, cols, title, sub):
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=cols)
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=cols)
        ws.cell(row=1, column=1, value=title).font = font_title
        ws.cell(row=2, column=1, value=sub).font = font_subtitle
        for r in (1, 2):
            for c in range(1, cols + 1):
                ws.cell(row=r, column=c).fill = fill_title
                ws.cell(row=r, column=c).alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[1].height = 26
        ws.row_dimensions[2].height = 18

    # Pestaña 1: Resumen Ejecutivo 4 Bloques
    ws_res = wb.active
    ws_res.title = "Resumen Ejecutivo 4 Bloques"
    banner(ws_res, 6, "OPERACIÓN VÉRTICE — CONSOLIDADO 4 BLOQUES OPERATIVOS", f"Generado automáticamente | Registros únicos procesados: {len(df)}")

    bloques_orden = [
        "Bloque 1 (06:00-11:59)",
        "Bloque 2 (12:00-17:25)",
        "Bloque 3 (17:26-23:59)",
        "Bloque 4 (00:00-06:00)"
    ]
    headers = ["Categoría Operativa"] + bloques_orden + ["Total Acumulado"]
    ws_res.row_dimensions[4].height = 26
    for idx, h in enumerate(headers, 1):
        c = ws_res.cell(row=4, column=idx, value=h)
        c.font = font_header
        c.fill = fill_header
        c.alignment = align_center
        c.border = border_data

    categorias = [
        ("Personas Solicitadas / Capturas (Oficios)", "Solicitado / Oficio", None),
        ("Infracciones de Tránsito Registradas", None, "Boletas"),
        ("Faltas Administrativas / Indocumentados", "Falta Administrativa / Identidad", None),
        ("Personas Verificadas (Conteo Depurado)", None, "Personas"),
        ("Vehículos Verificados (Conteo Depurado)", None, "Vehículos"),
    ]

    for r_i, (label, cat_filter, sum_col) in enumerate(categorias, start=5):
        ws_res.row_dimensions[r_i].height = 24
        fill = fill_zebra if r_i % 2 == 1 else fill_white
        c0 = ws_res.cell(row=r_i, column=1, value=label)
        c0.font = font_bold
        c0.fill = fill
        c0.border = border_data
        c0.alignment = align_left

        for b_i, b_name in enumerate(bloques_orden, start=2):
            sub_df = df[df["Bloque Horario"] == b_name]
            if cat_filter:
                val = int((sub_df["Categoría"] == cat_filter).sum())
            else:
                val = int(sub_df[sum_col].sum())
            c = ws_res.cell(row=r_i, column=b_i, value=val)
            c.font = font_regular
            c.fill = fill
            c.border = border_data
            c.alignment = align_center
            c.number_format = "#,##0"

        ctot = ws_res.cell(row=r_i, column=6, value=f"=SUM(B{r_i}:E{r_i})")
        ctot.font = font_bold
        ctot.fill = fill
        ctot.border = border_data
        ctot.alignment = align_center
        ctot.number_format = "#,##0"

    chart = BarChart()
    chart.type = "col"
    chart.title = "Comparativo Operativo por Bloque Horario"
    data_ref = Reference(ws_res, min_col=2, min_row=4, max_col=5, max_row=6)
    cats_ref = Reference(ws_res, min_col=1, min_row=5, max_row=6)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    chart.width = 16
    chart.height = 8
    ws_res.add_chart(chart, "A12")

    for col_l, w in zip(["A", "B", "C", "D", "E", "F"], [38, 20, 20, 20, 20, 18]):
        ws_res.column_dimensions[col_l].width = w

    # Pestañas de detalle
    pestanas = [
        ("Solicitados y Faltas", df[df["Categoría"].isin(["Solicitado / Oficio", "Falta Administrativa / Identidad"])]),
        ("Infracciones Tránsito", df[df["Boletas"] > 0]),
        ("Verificaciones por Punto", df),
    ]

    for sheet_name, sub in pestanas:
        ws = wb.create_sheet(sheet_name)
        banner(ws, len(df.columns), f"OPERACIÓN VÉRTICE — {sheet_name.upper()}", "Registros depurados y clasificados")
        ws.row_dimensions[4].height = 24
        for c_i, col_name in enumerate(df.columns, 1):
            c = ws.cell(row=4, column=c_i, value=col_name)
            c.font = font_header
            c.fill = fill_header
            c.alignment = align_center
            c.border = border_data

        for r_i, (_, row) in enumerate(sub.iterrows(), start=5):
            ws.row_dimensions[r_i].height = 28
            fill = fill_zebra if r_i % 2 == 1 else fill_white
            for c_i, val in enumerate(row.values, 1):
                c = ws.cell(row=r_i, column=c_i, value=val)
                c.font = font_regular
                c.fill = fill
                c.border = border_data
                c.alignment = align_center if c_i in (1, 2, 3, 6, 7, 8, 9, 10) else align_left

        ws.freeze_panes = "A5"

    for s in wb.worksheets:
        s.page_setup.orientation = s.ORIENTATION_LANDSCAPE
        s.page_setup.paperSize = s.PAPERSIZE_LETTER
        s.sheet_properties.pageSetUpPr.fitToPage = True
        s.page_setup.fitToWidth = 1
        s.page_setup.fitToHeight = 0

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

# ==============================================================================
# INTERFAZ DE USUARIO STREAMLIT
# ==============================================================================
texto_input = st.text_area(
    "📋 Pega aquí el bloque completo de reportes de WhatsApp / Telegram:",
    height=280,
    placeholder="Ejemplo:\n- BO-53-Chepo (Entrada principal)\n*3RA BRIGADA PANAMÁ ESTE/ BAT-CHEPO*..."
)

if st.button("⚙️ Procesar y Generar Cuadro Consolidado", type="primary"):
    if not texto_input.strip():
        st.warning("Por favor, pega los textos operativos antes de procesar.")
    else:
        df_resultado = parsear_reportes(texto_input)
        st.success(f"✅ Procesamiento completado: {len(df_resultado)} registros únicos depurados.")

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Solicitados / Oficios", int((df_resultado["Categoría"] == "Solicitado / Oficio").sum()))
        col2.metric("Infracciones Tránsito", int(df_resultado["Boletas"].sum()))
        col3.metric("Personas Verificadas", f"{int(df_resultado['Personas'].sum()):,}")
        col4.metric("Vehículos Verificados", f"{int(df_resultado['Vehículos'].sum()):,}")

        st.dataframe(df_resultado, use_container_width=True)

        excel_bytes = generar_excel_operacion_vertice(df_resultado)
        st.download_button(
            label="📥 Descargar Excel Operación Vértice (4 Bloques)",
            data=excel_bytes,
            file_name=f"Operacion_Vertice_Consolidado_{datetime.now().strftime('%d%b%y_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
