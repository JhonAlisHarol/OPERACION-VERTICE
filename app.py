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
    page_title="CONSOLIDADO 24 HORAS — Operación Vértice | CON-C5",
    page_icon="🛡️",
    layout="wide"
)

# ==============================================================================
# VALORES AUDITADOS DE REFERENCIA (EXCEL v5 - 24 HORAS)
# ==============================================================================
BASE_VALORES = {
    "Consolidado": {
        "oficios": 14,
        "infracciones": 136,
        "gruas": 3,
        "incidentes": 7,
        "personas": 4009,
        "vehiculos": 2738,
    },
    "Corte 1 (06:00-11:59)": {
        "oficios": 3,
        "infracciones": 14,
        "gruas": 0,
        "incidentes": 0,
        "personas": 0,
        "vehiculos": 0,
    },
    "Corte 2 (12:00-17:25)": {
        "oficios": 4,
        "infracciones": 62,
        "gruas": 2,
        "incidentes": 5,
        "personas": 2268,
        "vehiculos": 1741,
    },
    "Corte 3 (17:26-23:59)": {
        "oficios": 6,
        "infracciones": 54,
        "gruas": 1,
        "incidentes": 2,
        "personas": 1528,
        "vehiculos": 852,
    },
    "Corte 4 (00:00-06:00)": {
        "oficios": 1,
        "infracciones": 6,
        "gruas": 0,
        "incidentes": 0,
        "personas": 213,
        "vehiculos": 145,
    },
}

SENAFRONT_BOS = {
    "BO-4", "BO-5", "BO-6", "BO-11", "BO-15", "BO-53", "BO-54", "BO-59", "BO-60", "BO-61"
}

BLOQUES_ORDEN = [
    "Corte 1 (06:00-11:59)",
    "Corte 2 (12:00-17:25)",
    "Corte 3 (17:26-23:59)",
    "Corte 4 (00:00-06:00)"
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

def detectar_corte(texto: str) -> str:
    t_up = texto.upper()
    m_mil = re.search(r"\b(\d{2})(\d{2})(\d{2})NOV\d{2}\b", t_up)
    if m_mil:
        dia, hh, mm = int(m_mil.group(1)), int(m_mil.group(2)), int(m_mil.group(3))
        mins = hh * 60 + mm
        if dia >= 19 and mins <= 360:
            return "Corte 4 (00:00-06:00)"
        if mins < 720:
            return "Corte 1 (06:00-11:59)"
        elif mins <= 1045:
            return "Corte 2 (12:00-17:25)"
        else:
            return "Corte 3 (17:26-23:59)"

    m_hm = re.search(r"(?:HORA[S]?:?|A LAS)\s*(\d{1,2}):(\d{2})", t_up)
    if m_hm:
        hh, mm = int(m_hm.group(1)), int(m_hm.group(2))
        if ("P.M." in t_up or "PM" in t_up) and hh < 12:
            hh += 12
        mins = hh * 60 + mm
        if mins <= 360:
            return "Corte 4 (00:00-06:00)"
        elif mins < 720:
            return "Corte 1 (06:00-11:59)"
        elif mins <= 1045:
            return "Corte 2 (12:00-17:25)"
        else:
            return "Corte 3 (17:26-23:59)"

    if "BUENOS DIAS" in t_up or "BUENOS DÍAS" in t_up:
        return "Corte 4 (00:00-06:00)"
    if "BUENAS NOCHES" in t_up:
        return "Corte 3 (17:26-23:59)"
    return "Corte 2 (12:00-17:25)"

def parsear_texto_a_df(texto_crudo: str) -> pd.DataFrame:
    bloques = re.split(r"(?=\bBO-\d+-)", texto_crudo.strip())
    filas = []
    vistos = set()

    for b in bloques:
        b = b.strip()
        if not b or not b.startswith("BO-"):
            continue
        huella = re.sub(r"\s+", " ", b.lower()).strip()
        if huella in vistos:
            continue
        vistos.add(huella)

        lineas = [l.strip() for l in b.splitlines() if l.strip()]
        encabezado = lineas[0]
        m_bo = re.match(r"^(BO-\d+)\s*-\s*([^\n]+)", encabezado)
        bo_code = m_bo.group(1) if m_bo else "BO-S/N"
        corte = detectar_corte(b)
        inst = detectar_institucion(bo_code, b)
        b_up = b.upper()

        oficios = 1 if any(k in b_up for k in ["CONDUCCIÓN JUDICIAL", "CONDUCCION JUDICIAL", "CAPTURA", "OFICIO DE", "OFICIO N", "CINTILLO ROJO"]) else 0
        incidentes = 1 if any(k in b_up for k in ["INDOCUMENTADO", "ESTADÍA VENCIDA", "ESTADIA VENCIDA", "ACTA DE INSPECCIÓN SANITARIA", "BOLETA DE CITACIÓN", "ETRM", "C.V.I"]) else 0

        m_grua = re.search(r"REMOCI[ÓO]N\s+EN\s+GR[ÚU]A\s*[:=.-]*\s*(\d+)", b_up)
        gruas = int(m_grua.group(1)) if m_grua else (1 if any(k in b_up for k in ["GRÚA", "GRUA", "REMOVIDO EN GRÚA", "REMOVIDO EN GRUA", "VEHÍCULO EN GRÚA", "VEHICULO EN GRUA"]) else 0)

        boletas = 0
        if "INFRACCIÓN DE TRÁNSITO (00)" not in b_up and "INFRACCION DE TRANSITO (00)" not in b_up:
            m_bol_rev = re.search(r"(?:BOLETAS|INFRACCIONES\s+DE\s+TR[ÁA]NSITO)\s*[:=.-]*\s*(\d+)", b_up)
            m_bol = re.search(r"(\d+)\s*(?:BOLETAS|INFRACCIONES)", b_up)
            if m_bol_rev:
                boletas = int(m_bol_rev.group(1))
            elif m_bol:
                boletas = int(m_bol.group(1))
            elif any(k in b_up for k in ["INFRACCIÓN", "INFRACCION", "BOLETA #", "BOLETA N", "ARTÍCULO 240", "ARTICULO 240", "ARTÍCULO 132", "ARTICULO 132", "SANCIONADO", "SANCIONÓ"]):
                boletas = 1

        personas = 0
        vehiculos = 0
        m_pers = re.findall(r"(?:PERSONAS(?:\s+VERIFICADAS)?|CIUDADANOS(?:\s+NACIONALES|\s+EXTRANJEROS)?)\s*[:=.-]*\s*(\d+)", b_up)
        if m_pers:
            personas = sum(int(x) for x in m_pers)
        elif "PERSONA VERIFICADA" in b_up or "CIUDADANO VERIFICADO" in b_up or oficios or incidentes:
            personas = 1

        m_veh = re.findall(r"(?:VEH[ÍI]CULOS(?:\s+PARTICULARES|\s+NACIONALES|\s+EXTRANJEROS)?|TAXIS?|BUSES?|MOTOS?|CAMI[ÓO]N(?:ES)?|VOLQUETES?|ARTICULADOS?|MULAS?)\s*[:=.-]*\s*(\d+)", b_up)
        if m_veh:
            vehiculos = sum(int(x) for x in m_veh)
        elif "PLACA" in b_up or "MATRÍCULA" in b_up or "MATRICULA" in b_up or gruas > 0:
            vehiculos = 1

        filas.append({
            "Corte": corte,
            "Institución": inst,
            "Sector (BO)": encabezado,
            "Oficios": oficios,
            "Infracciones": boletas,
            "Grúas": gruas,
            "Incidentes": incidentes,
            "Personas": personas,
            "Vehículos": vehiculos,
            "Detalle": " ".join(lineas[1:])[:300]
        })

    return pd.DataFrame(filas)

def generar_excel_v5(df_extra: pd.DataFrame | None = None) -> bytes:
    wb = openpyxl.Workbook()

    HEADER_DARK = "23395D"
    SUBHEADER_STEEL = "3A608F"
    SECTION_SAGE = "2E5A53"
    ZEBRA_LIGHT = "F7F9FC"
    SUBTOTAL_TINT = "EEF3F8"
    TOTAL_GRAY = "E4E9F0"
    BORDER_GRAY = "CBD5E1"

    font_title = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    font_sub = Font(name="Calibri", size=10, italic=True, color="E2E8F0")
    font_hdr = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_sec = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_reg = Font(name="Calibri", size=10, bold=False, color="1E293B")
    font_bld = Font(name="Calibri", size=10, bold=True, color="1E293B")
    font_tot = Font(name="Calibri", size=11, bold=True, color="23395D")

    fill_title = PatternFill(start_color=HEADER_DARK, end_color=HEADER_DARK, fill_type="solid")
    fill_hdr = PatternFill(start_color=SUBHEADER_STEEL, end_color=SUBHEADER_STEEL, fill_type="solid")
    fill_sec = PatternFill(start_color=SECTION_SAGE, end_color=SECTION_SAGE, fill_type="solid")
    fill_zebra = PatternFill(start_color=ZEBRA_LIGHT, end_color=ZEBRA_LIGHT, fill_type="solid")
    fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    fill_sub = PatternFill(start_color=SUBTOTAL_TINT, end_color=SUBTOTAL_TINT, fill_type="solid")
    fill_tot = PatternFill(start_color=TOTAL_GRAY, end_color=TOTAL_GRAY, fill_type="solid")

    thin = Side(style="thin", color=BORDER_GRAY)
    med_navy = Side(style="medium", color=HEADER_DARK)
    double_navy = Side(style="double", color=HEADER_DARK)
    border_cell = Border(left=thin, right=thin, top=thin, bottom=thin)
    border_sub = Border(left=thin, right=thin, top=med_navy, bottom=med_navy)
    border_tot = Border(left=thin, right=thin, top=med_navy, bottom=double_navy)

    al_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    al_center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws_res = wb.active
    ws_res.title = "Resumen Ejecutivo 4 Cortes"
    ws_res.views.sheetView[0].showGridLines = True

    ws_res.merge_cells("A1:H1")
    ws_res.merge_cells("A2:H2")
    ws_res["A1"] = "OPERACIÓN VÉRTICE — CONSOLIDADO FINAL 4 CORTES (24 HORAS)"
    ws_res["A2"] = "18 al 19 de noviembre de 2026 | Corte 1 (06:00-11:59) | Corte 2 (12:00-17:25) | Corte 3 (17:26-23:59) | Corte 4 (00:00-06:00)"
    ws_res["A1"].font = font_title
    ws_res["A2"].font = font_sub
    for r in (1, 2):
        for c in range(1, 9):
            ws_res.cell(row=r, column=c).fill = fill_title
            ws_res.cell(row=r, column=c).alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws_res.row_dimensions[1].height = 26
    ws_res.row_dimensions[2].height = 18

    hdrs = ["Categoría / Tipo de Acción", "Corte 1
(06:00 - 11:59)", "Corte 2
(12:00 - 17:25)", "Corte 3
(17:26 - 23:59)", "Corte 4
(00:00 - 06:00)", "Total 24 Horas", "Distribución Institucional", "Notas"]
    ws_res.row_dimensions[4].height = 28
    for idx, h in enumerate(hdrs, 1):
        c = ws_res.cell(row=4, column=idx, value=h)
        c.font = font_hdr
        c.fill = fill_hdr
        c.alignment = al_center
        c.border = border_cell

    # Extra sums if user pasted new text
    extra_c3_inf = int(df_extra[df_extra["Corte"] == "Corte 3 (17:26-23:59)"]["Infracciones"].sum()) if df_extra is not None and not df_extra.empty else 0
    extra_c4_inf = int(df_extra[df_extra["Corte"] == "Corte 4 (00:00-06:00)"]["Infracciones"].sum()) if df_extra is not None and not df_extra.empty else 0

    filas_res = [
        ("Personas Solicitadas / Capturas (Oficios)", 3, 4, 6, 1, "=SUM(B5:E5)", "8 PN / 6 SNF", "Total 24h: 14 capturas por oficio"),
        ("Infracciones de Tránsito Registradas", 14, 62, 54 + extra_c3_inf, 6 + extra_c4_inf, "=SUM(B6:E6)", "74 PN / 62 SNF", "Total 24h: 136 boletas/sanciones"),
        ("Vehículos Removidos en Grúa", 0, 2, 1, 0, "=SUM(B7:E7)", "3 PN / 0 SNF", "Llano Bonito, Almirante y David BV0753"),
        ("Faltas Administrativas / Migratorias / Indocumentados", 0, 5, 2, 0, "=SUM(B8:E8)", "SNM / MINSA / PN / SNF", "Total 24h: 7 faltas admin / indocumentados"),
        ("Personas Verificadas (Conteo Depurado)", 0, 2268, 1528, 213, "=SUM(B9:E9)", "PN + SENAFRONT", "Total 24h: 4,009 personas verificadas"),
        ("Vehículos Verificados (Conteo Depurado)", 0, 1741, 852, 145, "=SUM(B10:E10)", "PN + SENAFRONT", "Total 24h: 2,738 vehículos verificados"),
    ]

    for r_i, vals in enumerate(filas_res, start=5):
        ws_res.row_dimensions[r_i].height = 25
        fl = fill_zebra if r_i % 2 == 1 else fill_white
        for c_i, v in enumerate(vals, 1):
            c = ws_res.cell(row=r_i, column=c_i, value=v)
            c.fill = fl
            c.border = border_cell
            if c_i == 1:
                c.font = font_bld
                c.alignment = al_left
            elif c_i in (2, 3, 4, 5, 6):
                c.font = font_tot if c_i == 6 else font_reg
                c.alignment = al_center
                c.number_format = "#,##0"
            else:
                c.font = font_reg
                c.alignment = al_left

    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = "Infracciones, Grúas y Capturas por Corte Horario"
    data_ref = Reference(ws_res, min_col=2, min_row=4, max_col=5, max_row=7)
    cats_ref = Reference(ws_res, min_col=1, min_row=5, max_row=7)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    chart.width = 18
    chart.height = 8.5
    ws_res.add_chart(chart, "A12")

    for col_l, w in zip(["A", "B", "C", "D", "E", "F", "G", "H"], [40, 16, 16, 16, 16, 16, 26, 36]):
        ws_res.column_dimensions[col_l].width = w

    for sheet in wb.worksheets:
        sheet.page_setup.orientation = sheet.ORIENTATION_LANDSCAPE
        sheet.page_setup.paperSize = sheet.PAPERSIZE_LETTER
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
        sheet.page_setup.fitToWidth = 1
        sheet.page_setup.fitToHeight = 0

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ==============================================================================
# ESTILO VISUAL TIPO TABLERO CON-C5 (MINSEG — CON PASO FIRME)
# ==============================================================================
st.markdown("""
<style>
  .top-banner {
    background: linear-gradient(90deg, #071936 0%, #0E2F63 55%, #071936 100%);
    border: 1px solid #1E498A;
    border-radius: 6px;
    padding: 14px 20px;
    margin-bottom: 12px;
    display: flex;
    justify-content: space-between;
    align-items: center;
  }
  .top-banner-title {
    color: #FFFFFF;
    font-size: 20px;
    font-weight: 800;
    letter-spacing: 0.5px;
    margin: 0;
  }
  .top-banner-sub {
    color: #93C5FD;
    font-size: 12px;
    margin: 3px 0 0 0;
  }
  .hero-header {
    background: #051329;
    text-align: center;
    padding: 10px 16px;
    border-bottom: 2px solid #1E498A;
    margin-bottom: 14px;
  }
  .hero-title {
    color: #FFFFFF;
    font-size: 26px;
    font-weight: 900;
    letter-spacing: 1px;
    margin: 0;
  }
  .hero-sub {
    color: #E2E8F0;
    font-size: 13px;
    margin: 4px 0 0 0;
  }
  .panel-box {
    background: #071936;
    border: 2px solid #00A8E8;
    border-radius: 4px;
    padding: 0 0 10px 0;
    margin-bottom: 14px;
  }
  .panel-box.green { border-color: #22C55E; }
  .panel-box.amber { border-color: #F59E0B; }
  .panel-box.orange { border-color: #F97316; }
  .panel-hdr {
    background: #0D2B5C;
    color: #FFFFFF;
    font-weight: 800;
    font-size: 14px;
    text-align: center;
    padding: 8px;
    border-bottom: 1px solid #1E498A;
    text-transform: uppercase;
  }
  .kpi-big {
    font-size: 42px;
    font-weight: 900;
    text-align: center;
    line-height: 1.1;
    margin-top: 10px;
  }
  .kpi-blue { color: #38BDF8; }
  .kpi-green { color: #4ADE80; }
  .kpi-amber { color: #FBBF24; }
  .kpi-orange { color: #FB923C; }
  .kpi-lbl {
    color: #FFFFFF;
    font-size: 12px;
    font-weight: 700;
    text-align: center;
    margin-top: 4px;
  }
  .kpi-foot {
    background: #041024;
    color: #CBD5E1;
    font-size: 11px;
    text-align: center;
    padding: 6px;
    margin-top: 8px;
    border-top: 1px solid #163A70;
  }
</style>
""", unsafe_allow_html=True)

# Encabezado superior y selector de Grupo de Horas (estilo imagen enviada)
col_b_izq, col_b_der = st.columns([3.4, 1.1])
with col_b_izq:
    st.markdown("""
    <div class="top-banner">
      <div>
        <div class="top-banner-title">GOBIERNO NACIONAL • CON PASO FIRME — MINISTERIO DE SEGURIDAD PÚBLICA</div>
        <div class="top-banner-sub">CENTRO DE OPERACIÓN NACIONAL (C.O.N. / C5) • OPERACIÓN VÉRTICE</div>
      </div>
    </div>
    """, unsafe_allow_html=True)
with col_b_der:
    grupo_sel = st.selectbox(
        "GRUPO DE HORAS ▼",
        options=[
            "Consolidado",
            "Corte 1 (06:00-11:59)",
            "Corte 2 (12:00-17:25)",
            "Corte 3 (17:26-23:59)",
            "Corte 4 (00:00-06:00)"
        ],
        index=0
    )

st.markdown(f"""
<div class="hero-header">
  <div class="hero-title">CONSOLIDADO 24 HORAS</div>
  <div class="hero-sub">Operación Vértice • Puntos de control • Estamento: Todos • Vista activa: <strong>{grupo_sel}</strong></div>
</div>
""", unsafe_allow_html=True)

vals_act = BASE_VALORES[grupo_sel].copy()

# Si el usuario procesó texto nuevo en la pestaña de carga, sumarlo dinámicamente
if "df_vertice_extra" in st.session_state and not st.session_state["df_vertice_extra"].empty:
    df_ex = st.session_state["df_vertice_extra"]
    if grupo_sel == "Consolidado":
        vals_act["oficios"] += int(df_ex["Oficios"].sum())
        vals_act["infracciones"] += int(df_ex["Infracciones"].sum())
        vals_act["gruas"] += int(df_ex["Grúas"].sum())
        vals_act["incidentes"] += int(df_ex["Incidentes"].sum())
        vals_act["personas"] += int(df_ex["Personas"].sum())
        vals_act["vehiculos"] += int(df_ex["Vehículos"].sum())
    else:
        sub_ex = df_ex[df_ex["Corte"] == grupo_sel]
        vals_act["oficios"] += int(sub_ex["Oficios"].sum())
        vals_act["infracciones"] += int(sub_ex["Infracciones"].sum())
        vals_act["gruas"] += int(sub_ex["Grúas"].sum())
        vals_act["incidentes"] += int(sub_ex["Incidentes"].sum())
        vals_act["personas"] += int(sub_ex["Personas"].sum())
        vals_act["vehiculos"] += int(sub_ex["Vehículos"].sum())

tab_pres, tab_proc = st.tabs(["📊 Presentación Ejecutiva (CON-C5)", "📥 Cargar Textos y Exportar Excel"])

with tab_pres:
    c1, c2, c3, c4 = st.columns([1.0, 1.1, 1.8, 1.3])

    with c1:
        st.markdown("""
        <div class="panel-box">
          <div class="panel-hdr">👥 PERSONAL</div>
          <div class="kpi-big kpi-blue">587</div>
          <div class="kpi-lbl">PERSONAL EMPEÑADO</div>
          <div class="kpi-foot">Turno 454 • Migración 33 • Tránsito 36</div>
        </div>
        """, unsafe_allow_html=True)

    with c2:
        st.markdown("""
        <div class="panel-box green">
          <div class="panel-hdr">📍 PUNTOS DE CONTROL</div>
          <div class="kpi-big kpi-green">63</div>
          <div class="kpi-lbl">INSTALADOS DE 63 ASIGNADOS</div>
          <div class="kpi-foot">Con reporte 24 • Sin reporte 39</div>
        </div>
        """, unsafe_allow_html=True)

    with c3:
        st.markdown(f"""
        <div class="panel-box amber">
          <div class="panel-hdr">🛡️ RESULTADOS OPERATIVOS</div>
          <div style="display:flex; justify-content:space-around; align-items:center; padding-top:6px;">
            <div>
              <div class="kpi-big kpi-amber" style="font-size:34px;">{vals_act['oficios']}</div>
              <div class="kpi-lbl" style="font-size:10px;">ACCIONES POLICIALES</div>
              <div style="color:#94A3B8; font-size:10px; text-align:center;">{grupo_sel}</div>
            </div>
            <div>
              <div class="kpi-big kpi-amber" style="font-size:34px;">{vals_act['personas']:,}</div>
              <div class="kpi-lbl" style="font-size:10px;">PERS. VERIFICADAS</div>
              <div style="color:#94A3B8; font-size:10px; text-align:center;">Vehículos: {vals_act['vehiculos']:,}</div>
            </div>
            <div>
              <div class="kpi-big kpi-amber" style="font-size:34px;">{vals_act['incidentes']}</div>
              <div class="kpi-lbl" style="font-size:10px;">INCIDENTES</div>
              <div style="color:#94A3B8; font-size:10px; text-align:center;">{grupo_sel}</div>
            </div>
          </div>
        </div>
        """, unsafe_allow_html=True)

    with c4:
        st.markdown(f"""
        <div class="panel-box orange">
          <div class="panel-hdr">🚘 TRÁNSITO</div>
          <div style="display:flex; justify-content:space-around; align-items:center; padding-top:6px;">
            <div>
              <div class="kpi-big kpi-orange" style="font-size:34px;">{vals_act['infracciones']}</div>
              <div class="kpi-lbl" style="font-size:10px;">INFRACCIONES</div>
              <div style="color:#94A3B8; font-size:10px; text-align:center;">{grupo_sel}</div>
            </div>
            <div>
              <div class="kpi-big kpi-orange" style="font-size:34px;">{vals_act['gruas']}</div>
              <div class="kpi-lbl" style="font-size:10px;">REMOLCADOS (GRÚA)</div>
              <div style="color:#94A3B8; font-size:10px; text-align:center;">{grupo_sel}</div>
            </div>
          </div>
        </div>
        """, unsafe_allow_html=True)

    # Fila inferior: Detalle de Acciones Policiales, Detalle de Incidentes y Novedades Relevantes
    b1, b2, b3 = st.columns([1.4, 1.4, 1.2])

    with b1:
        st.markdown('<div class="panel-box amber"><div class="panel-hdr" style="text-align:left;padding-left:12px;">DETALLE DE ACCIONES POLICIALES</div></div>', unsafe_allow_html=True)
        df_acc_det = pd.DataFrame([
            {"Tipo": "Persona solicitada (Oficios)", "Total": vals_act["oficios"]},
            {"Tipo": "Personas verificadas (PDA/SVC)", "Total": vals_act["personas"]},
            {"Tipo": "Vehículos verificados", "Total": vals_act["vehiculos"]},
        ])
        st.dataframe(df_acc_det, use_container_width=True, hide_index=True)

    with b2:
        st.markdown('<div class="panel-box amber"><div class="panel-hdr" style="text-align:left;padding-left:12px;">DETALLE DE INCIDENTES</div></div>', unsafe_allow_html=True)
        df_inc_det = pd.DataFrame([
            {"Tipo": "Falta administrativa / Migratoria / C.V.I.", "Total": vals_act["incidentes"]},
            {"Tipo": "Vehículos removidos en grúa", "Total": vals_act["gruas"]},
            {"Tipo": "Infracciones de tránsito", "Total": vals_act["infracciones"]},
        ])
        st.dataframe(df_inc_det, use_container_width=True, hide_index=True)

    with b3:
        st.markdown("""
        <div class="panel-box" style="padding:10px 14px; min-height:155px;">
          <div style="color:#38BDF8; font-weight:800; font-size:15px; margin-bottom:8px;">NOVEDADES RELEVANTES</div>
          <div style="color:#E2E8F0; font-size:12.5px; line-height:1.45;">
            • Grúas (3): Moto Nakeo Sport sin placa/licencia (AR8293, Todos Llanta), 1 en Almirante y Hyundai Accent BV0753 en David.<br>
            • Migración/Salud: Traslado ETRM San Vicente, Citación Simplus #2736, Acta MINSA #4006 y estadía vencida en Agua Fría.<br>
            • Armamento: Permiso vigente pistola Tanfoglio .40 (Z15339) en Capira.
          </div>
        </div>
        """, unsafe_allow_html=True)

with tab_proc:
    texto_pegar = st.text_area(
        "Pega aquí los textos de reportes operativos para extraer y sumar valores:",
        height=240,
        placeholder="- BO-53-Chepo (Entrada principal)..."
    )
    if st.button("⚙️ Extraer Valores y Actualizar Presentación", type="primary"):
        if texto_pegar.strip():
            df_ex = parsear_texto_a_df(texto_pegar)
            st.session_state["df_vertice_extra"] = df_ex
            st.success(f"✅ {len(df_ex)} registros extraídos y sumados a la presentación.")
            st.dataframe(df_ex, use_container_width=True, hide_index=True)

    excel_data = generar_excel_v5(st.session_state.get("df_vertice_extra"))
    st.download_button(
        label="📥 Descargar Excel Consolidado (4 Cortes)",
        data=excel_data,
        file_name=f"Operacion_Vertice_4_Cortes_{datetime.now().strftime('%d%b%y_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
