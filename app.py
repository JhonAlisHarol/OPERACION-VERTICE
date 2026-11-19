import re
import io
from datetime import datetime
import pandas as pd
import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference

# ==============================================================================
# CONFIGURACIÓN DE PÁGINA — CON-C5 PANAMÁ (MINSEG)
# ==============================================================================
st.set_page_config(
    page_title="Operación Vértice | CON-C5 Panamá",
    page_icon="🛡️",
    layout="wide"
)

st.title("🛡️ Consola Operativa — Operación Vértice (CON-C5)")
st.caption("Ministerio de Seguridad Pública (MINSEG) • Centro de Operación Nacional")

# ==============================================================================
# REGLAS OPERATIVAS DE CLASIFICACIÓN Y DEPURACIÓN
# ==============================================================================
SENAFRONT_BOS = {
    "BO-4", "BO-5", "BO-6", "BO-11", "BO-15", "BO-53", "BO-54", "BO-59", "BO-60", "BO-61"
}

BLOQUES_ORDEN = [
    "Bloque 1 (06:00-11:59)",
    "Bloque 2 (12:00-17:25)",
    "Bloque 3 (17:26-23:59)",
    "Bloque 4 (00:00-06:00)"
]

def detectar_institucion(bo_code: str, texto: str) -> str:
    t_up = texto.upper()
    if any(k in t_up for k in ["SENAFRONT", "CENTINELAS DE LA PATRIA", "BRIGADA", "BRIOCC", "B.G.M.Q.V", "BACTA"]):
        return "SENAFRONT"
    if any(k in t_up for k in ["POLICIA NACIONAL", "POLICÍA NACIONAL", "DIOS Y PATRIA", "ZONA DE POLIC", "ZONA POLICIAL"]):
        return "Policía Nacional"
    if bo_code in SENAFRONT_BOS:
        return "SENAFRONT"
    return "Policía Nacional"

def detectar_bloque_horario(texto: str) -> str:
    t_up = texto.upper()
    # Formato militar DDHHMMNOV26 (ej. 181236NOV26 o 190150NOV26)
    m_mil = re.search(r"\b(\d{2})(\d{2})(\d{2})NOV\d{2}\b", t_up)
    if m_mil:
        dia, hh, mm = int(m_mil.group(1)), int(m_mil.group(2)), int(m_mil.group(3))
        mins = hh * 60 + mm
        if dia >= 19 and mins <= 360:
            return "Bloque 4 (00:00-06:00)"
        if mins < 720:
            return "Bloque 1 (06:00-11:59)"
        elif mins <= 1045:  # hasta 17:25
            return "Bloque 2 (12:00-17:25)"
        else:
            return "Bloque 3 (17:26-23:59)"

    # Formato HH:MM
    m_hm = re.search(r"(?:HORA[S]?:?|A LAS)\s*(\d{1,2}):(\d{2})", t_up)
    if m_hm:
        hh, mm = int(m_hm.group(1)), int(m_hm.group(2))
        if ("P.M." in t_up or "PM" in t_up) and hh < 12:
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

    if "BUENOS DIAS" in t_up or "BUENOS DÍAS" in t_up:
        return "Bloque 4 (00:00-06:00)"
    if "BUENAS NOCHES" in t_up:
        return "Bloque 3 (17:26-23:59)"
    return "Bloque 2 (12:00-17:25)"

def parsear_reportes(texto_crudo: str) -> pd.DataFrame:
    bloques = re.split(r"(?=\bBO-\d+-)", texto_crudo.strip())
    registros = []
    vistos = set()

    for b in bloques:
        b = b.strip()
        if not b or not b.startswith("BO-"):
            continue

        # Huella normalizada para depurar reenvíos duplicados
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

        # Clasificación operativa
        if any(k in b_up for k in ["CONDUCCIÓN JUDICIAL", "CONDUCCION JUDICIAL", "CAPTURA", "OFICIO DE", "CINTILLO ROJO"]):
            categoria = "Solicitado / Oficio"
        elif any(k in b_up for k in ["INDOCUMENTADO", "ESTADÍA VENCIDA", "ESTADIA VENCIDA", "ACTA DE INSPECCIÓN SANITARIA", "BOLETA DE CITACIÓN", "ETRM"]):
            categoria = "Falta Administrativa / Identidad"
        elif any(k in b_up for k in ["INFRACCIÓN", "INFRACCION", "BOLETA #", "BOLETA N", "ARTÍCULO 240", "ARTICULO 240", "ARTÍCULO 132", "ARTICULO 132", "GRÚA", "GRUA"]):
            categoria = "Infracción de Tránsito"
        else:
            categoria = "Verificación / Control"

        # Extracción numérica
        personas = 0
        vehiculos = 0
        boletas = 0

        m_pers = re.findall(r"(?:PERSONAS(?:\s+VERIFICADAS)?|CIUDADANOS(?:\s+NACIONALES|\s+EXTRANJEROS)?)\s*[:=.-]*\s*(\d+)", b_up)
        if m_pers:
            personas = sum(int(x) for x in m_pers)
        elif any(k in b_up for k in ["PERSONA VERIFICADA", "CIUDADANO VERIFICADO"]) or categoria == "Solicitado / Oficio":
            personas = 1

        m_veh = re.findall(r"(?:VEH[ÍI]CULOS(?:\s+PARTICULARES|\s+NACIONALES|\s+EXTRANJEROS)?|TAXIS?|BUSES?|MOTOS?|CAMI[ÓO]N(?:ES)?|VOLQUETES?|ARTICULADOS?|MULAS?)\s*[:=.-]*\s*(\d+)", b_up)
        if m_veh:
            vehiculos = sum(int(x) for x in m_veh)
        elif any(k in b_up for k in ["PLACA", "MATRÍCULA", "MATRICULA"]):
            vehiculos = 1

        m_bol = re.search(r"(\d+)\s*(?:BOLETAS|INFRACCIONES)", b_up)
        if m_bol:
            boletas = int(m_bol.group(1))
        elif categoria == "Infracción de Tránsito":
            boletas = 1

        # Apoyo interinstitucional detectado
        apoyos = []
        for ent in ["SNM", "MIGRACIÓN", "MIGRACION", "MINSA", "ADUANAS", "MIDA", "ATTT", "DAN", "NARCÓTICOS", "DIJ", "DNIP"]:
            if ent in b_up:
                apoyos.append(ent)
        apoyo_str = ", ".join(dict.fromkeys(apoyos)) if apoyos else "N/A"

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
            "Apoyo Interinstitucional": apoyo_str,
            "Documento / C.I.P.": doc_id,
            "Placa": placa,
            "Detalle Operativo": cuerpo[:400]
        })

    return pd.DataFrame(registros)

def generar_excel_5_pestanas(df: pd.DataFrame) -> bytes:
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

    # 1. Resumen Ejecutivo 4 Bloques
    ws_res = wb.active
    ws_res.title = "Resumen Ejecutivo 4 Bloques"
    banner(ws_res, 6, "OPERACIÓN VÉRTICE — CONSOLIDADO GENERAL 4 BLOQUES (24 HORAS)", f"CON-C5 Panamá | Registros únicos depurados: {len(df)}")

    headers = ["Categoría Operativa"] + BLOQUES_ORDEN + ["Total Acumulado"]
    ws_res.row_dimensions[4].height = 28
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

        for b_i, b_name in enumerate(BLOQUES_ORDEN, start=2):
            sub_df = df[df["Bloque Horario"] == b_name]
            val = int((sub_df["Categoría"] == cat_filter).sum()) if cat_filter else int(sub_df[sum_col].sum())
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
    chart.title = "Infracciones de Tránsito por Bloque Horario"
    data_ref = Reference(ws_res, min_col=2, min_row=4, max_col=5, max_row=6)
    cats_ref = Reference(ws_res, min_col=1, min_row=5, max_row=6)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    chart.width = 16
    chart.height = 8
    ws_res.add_chart(chart, "A12")

    for col_l, w in zip(["A", "B", "C", "D", "E", "F"], [38, 20, 20, 20, 20, 18]):
        ws_res.column_dimensions[col_l].width = w

    # 4 pestañas de detalle operativo
    pestanas = [
        ("Solicitados y Faltas", df[df["Categoría"].isin(["Solicitado / Oficio", "Falta Administrativa / Identidad"])]),
        ("Infracciones Tránsito", df[df["Boletas"] > 0]),
        ("Verificaciones por Punto", df),
        ("Apoyo Interinstitucional", df[df["Apoyo Interinstitucional"] != "N/A"]),
    ]

    for sheet_name, sub in pestanas:
        ws = wb.create_sheet(sheet_name)
        banner(ws, len(df.columns), f"OPERACIÓN VÉRTICE — {sheet_name.upper()}", "Registros depurados y clasificados por Bloque e Institución")
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
                c.alignment = align_center if c_i in (1, 2, 3, 6, 7, 8, 10, 11) else align_left

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
# INTERFAZ PRINCIPAL (PESTAÑAS: PROCESADOR Y TABLERO EJECUTIVO)
# ==============================================================================
tab_proc, tab_dash = st.tabs(["📥 Procesador y Exportación Excel", "📊 Presentación Ejecutiva en Vivo"])

with tab_proc:
    texto_input = st.text_area(
        "Pega aquí los reportes operativos crudos (WhatsApp / Telegram / PDF):",
        height=260,
        placeholder="- BO-53-Chepo (Entrada principal)\n*3RA BRIGADA PANAMÁ ESTE/ BAT-CHEPO*..."
    )

    if st.button("⚙️ Procesar Reportes y Generar Excel", type="primary"):
        if not texto_input.strip():
            st.warning("Pega los textos operativos antes de procesar.")
        else:
            df_res = parsear_reportes(texto_input)
            st.session_state["df_vertice"] = df_res
            st.success(f"✅ {len(df_res)} registros únicos procesados y depurados.")

    if "df_vertice" in st.session_state and not st.session_state["df_vertice"].empty:
        df_res = st.session_state["df_vertice"]
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Solicitados / Oficios", int((df_res["Categoría"] == "Solicitado / Oficio").sum()))
        c2.metric("Infracciones Tránsito", int(df_res["Boletas"].sum()))
        c3.metric("Personas Verificadas", f"{int(df_res['Personas'].sum()):,}")
        c4.metric("Vehículos Verificados", f"{int(df_res['Vehículos'].sum()):,}")

        excel_bytes = generar_excel_5_pestanas(df_res)
        st.download_button(
            label="📥 Descargar Cuadro Consolidado Excel (5 Pestañas)",
            data=excel_bytes,
            file_name=f"Operacion_Vertice_4_Bloques_{datetime.now().strftime('%d%b%y_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.dataframe(df_res, use_container_width=True)

with tab_dash:
    if "df_vertice" not in st.session_state or st.session_state["df_vertice"].empty:
        st.info("Procesa primero un bloque de reportes en la pestaña anterior para visualizar el tablero ejecutivo.")
    else:
        df_res = st.session_state["df_vertice"]
        st.subheader("Resumen Ejecutivo por Bloque Horario e Institución")
        resumen_b = df_res.groupby(["Bloque Horario", "Institución"], as_index=False)[["Personas", "Vehículos", "Boletas"]].sum()
        st.dataframe(resumen_b, use_container_width=True)
        st.bar_chart(df_res.groupby("Bloque Horario")[["Personas", "Vehículos", "Boletas"]].sum())
