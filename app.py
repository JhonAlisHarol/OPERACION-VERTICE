import re
import io
from datetime import datetime
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
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

st.title("🛡️ Consola Operativa y Presentación Ejecutiva — Operación Vértice")
st.caption("Ministerio de Seguridad Pública (MINSEG) • Centro de Operación Nacional (CON-C5)")

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
    m_mil = re.search(r"\b(\d{2})(\d{2})(\d{2})NOV\d{2}\b", t_up)
    if m_mil:
        dia, hh, mm = int(m_mil.group(1)), int(m_mil.group(2)), int(m_mil.group(3))
        mins = hh * 60 + mm
        if dia >= 19 and mins <= 360:
            return "Bloque 4 (00:00-06:00)"
        if mins < 720:
            return "Bloque 1 (06:00-11:59)"
        elif mins <= 1045:
            return "Bloque 2 (12:00-17:25)"
        else:
            return "Bloque 3 (17:26-23:59)"

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

def clasificar_tipo_falta_transito(texto_up: str) -> str:
    if any(k in texto_up for k in ["ALCOHOLEMIA", "EMBRIAGUEZ", "CÓDIGO 46", "CODIGO 46"]):
        return "Embriaguez / Alcoholemia"
    if any(k in texto_up for k in ["DESACATO", "240"]):
        return "Desacato (Art. 240)"
    if any(k in texto_up for k in ["LICENCIA VENCIDA", "132", "NO PORTAR LICENCIA", "NO MANTENIA LICENCIA", "NO MANTENÍA LICENCIA"]):
        return "Sin Licencia / Licencia Vencida"
    if any(k in texto_up for k in ["LUCES", "#14", "53A", "53D", "53 INC", "REFLECTIVO", "REFLEXIVAS"]):
        return "Luces / Reflectivos"
    if any(k in texto_up for k in ["DESATENDER SEÑALES", "DESATENDER SENALES", "#59"]):
        return "Desatender Señales"
    if any(k in texto_up for k in ["GIRO INDEBIDO", "#58"]):
        return "Giro Indebido"
    if any(k in texto_up for k in ["CINTURÓN", "CINTURON", "737", "#37"]):
        return "Sin Cinturón de Seguridad"
    if any(k in texto_up for k in ["PAPEL AHUMADO", "VIDRIOS AHUMADO", "#21"]):
        return "Papel Ahumado"
    if any(k in texto_up for k in ["PÓLIZA", "POLIZA", "SEGURO"]):
        return "Sin Póliza / Seguro Vencido"
    if any(k in texto_up for k in ["SIN PLACA", "SIN METÁLICA", "SIN METALICA", "PLACA VISIBLE", "#10"]):
        return "Sin Placa / Placa No Visible"
    if any(k in texto_up for k in ["LLANTA LISA"]):
        return "Llanta Lisa"
    return "Otras Infracciones Reglamento"

def parsear_reportes(texto_crudo: str) -> pd.DataFrame:
    bloques = re.split(r"(?=\bBO-\d+-)", texto_crudo.strip())
    registros = []
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
        punto_control = encabezado

        cuerpo = "\n".join(lineas[1:])
        institucion = detectar_institucion(bo_code, b)
        bloque = detectar_bloque_horario(b)
        b_up = b.upper()

        # Clasificación principal de la acción
        if any(k in b_up for k in ["CONDUCCIÓN JUDICIAL", "CONDUCCION JUDICIAL", "CAPTURA", "OFICIO DE", "OFICIO N", "CINTILLO ROJO"]):
            categoria = "Solicitado / Oficio"
            subtipo = "Captura / Conducción Judicial"
        elif any(k in b_up for k in ["INDOCUMENTADO", "ESTADÍA VENCIDA", "ESTADIA VENCIDA", "ACTA DE INSPECCIÓN SANITARIA", "BOLETA DE CITACIÓN", "ETRM", "C.V.I"]):
            categoria = "Falta Administrativa / Identidad"
            if "SANITARIA" in b_up or "MINSA" in b_up:
                subtipo = "Decomiso Sanitario (MINSA)"
            elif "INDOCUMENTADO" in b_up or "C.V.I" in b_up:
                subtipo = "Conducción por Identidad (C.V.I.)"
            else:
                subtipo = "Falta Migratoria (SNM)"
        elif any(k in b_up for k in ["INFRACCIÓN", "INFRACCION", "BOLETA #", "BOLETA N", "ARTÍCULO 240", "ARTICULO 240", "ARTÍCULO 132", "ARTICULO 132", "GRÚA", "GRUA"]):
            categoria = "Infracción de Tránsito"
            subtipo = clasificar_tipo_falta_transito(b_up)
        else:
            categoria = "Verificación / Control"
            if "PROFILAXIS" in b_up:
                subtipo = "Profilaxis Social"
            elif "325" in b_up:
                subtipo = "Registro Art. 325 CPP"
            elif "ARMA" in b_up or "TONFLOGIO" in b_up or "TANFOGLIO" in b_up:
                subtipo = "Verificación de Armamento"
            else:
                subtipo = "Control Preventivo / PDA / SVC"

        # Extracción numérica de indicadores
        personas = 0
        vehiculos = 0
        boletas = 0
        gruas = 0
        oficios = 0
        faltas_admin = 0
        armas_verif = 0

        if categoria == "Solicitado / Oficio":
            oficios = 1
        if categoria == "Falta Administrativa / Identidad":
            faltas_admin = 1

        # Remoción en grúa
        m_grua = re.search(r"REMOCI[ÓO]N\s+EN\s+GR[ÚU]A\s*[:=.-]*\s*(\d+)", b_up)
        if m_grua:
            gruas = int(m_grua.group(1))
        elif any(k in b_up for k in ["GRÚA", "GRUA", "REMOVIDO EN GRÚA", "REMOVIDO EN GRUA", "VEHÍCULO EN GRÚA", "VEHICULO EN GRUA"]):
            gruas = 1

        # Armas verificadas con permiso
        m_arma = re.search(r"ARMA\s+CON\s+PERMISO\s*[:=.-]*\s*(\d+)", b_up)
        if m_arma:
            armas_verif = int(m_arma.group(1))
        elif "ARMA DE FUEGO" in b_up or "PERMISO DE PORTE" in b_up:
            armas_verif = 1

        # Personas verificadas
        m_pers = re.findall(r"(?:PERSONAS(?:\s+VERIFICADAS)?|CIUDADANOS(?:\s+NACIONALES|\s+EXTRANJEROS)?)\s*[:=.-]*\s*(\d+)", b_up)
        if m_pers:
            personas = sum(int(x) for x in m_pers)
        elif any(k in b_up for k in ["PERSONA VERIFICADA", "CIUDADANO VERIFICADO"]) or categoria in ("Solicitado / Oficio", "Falta Administrativa / Identidad"):
            personas = 1

        # Vehículos verificados
        m_veh = re.findall(r"(?:VEH[ÍI]CULOS(?:\s+PARTICULARES|\s+NACIONALES|\s+EXTRANJEROS)?|TAXIS?|BUSES?|MOTOS?|CAMI[ÓO]N(?:ES)?|VOLQUETES?|ARTICULADOS?|MULAS?)\s*[:=.-]*\s*(\d+)", b_up)
        if m_veh:
            vehiculos = sum(int(x) for x in m_veh)
        elif any(k in b_up for k in ["PLACA", "MATRÍCULA", "MATRICULA"]) or gruas > 0:
            vehiculos = 1

        # Boletas / Infracciones de tránsito
        m_bol = re.search(r"(\d+)\s*(?:BOLETAS|INFRACCIONES)", b_up)
        m_bol_rev = re.search(r"(?:BOLETAS|INFRACCIONES\s+DE\s+TR[ÁA]NSITO)\s*[:=.-]*\s*(\d+)", b_up)
        if m_bol_rev:
            boletas = int(m_bol_rev.group(1))
        elif m_bol:
            boletas = int(m_bol.group(1))
        elif categoria == "Infracción de Tránsito" or gruas > 0:
            boletas = 1

        # Apoyo interinstitucional
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
            "Tipo / Detalle de Falta": subtipo,
            "Oficios / Capturas": oficios,
            "Infracciones Tránsito": boletas,
            "Faltas Admin / Indoc": faltas_admin,
            "Vehículos en Grúa": gruas,
            "Armas Verificadas": armas_verif,
            "Personas Verificadas": personas,
            "Vehículos Verificados": vehiculos,
            "Apoyo Interinstitucional": apoyo_str,
            "Documento / C.I.P.": doc_id,
            "Placa": placa,
            "Narrativa Operativa": cuerpo[:420]
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

    fill_title = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    fill_header = PatternFill(start_color=SLATE_SUBHEADER, end_color=SLATE_SUBHEADER, fill_type="solid")
    fill_zebra = PatternFill(start_color=ZEBRA_FILL, end_color=ZEBRA_FILL, fill_type="solid")
    fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    thin_side = Side(style="thin", color=BORDER_COLOR)
    border_data = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

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

    ws_res = wb.active
    ws_res.title = "Resumen Ejecutivo 4 Bloques"
    banner(ws_res, 6, "OPERACIÓN VÉRTICE — CONSOLIDADO GENERAL 4 BLOQUES (24 HORAS)", f"CON-C5 Panamá | Registros únicos depurados: {len(df)}")

    headers = ["Indicador Operativo"] + BLOQUES_ORDEN + ["Total Acumulado"]
    ws_res.row_dimensions[4].height = 28
    for idx, h in enumerate(headers, 1):
        c = ws_res.cell(row=4, column=idx, value=h)
        c.font = font_header
        c.fill = fill_header
        c.alignment = align_center
        c.border = border_data

    metricas_resumen = [
        ("Personas Solicitadas / Capturas (Oficios)", "Oficios / Capturas"),
        ("Infracciones de Tránsito Registradas", "Infracciones Tránsito"),
        ("Vehículos Removidos en Grúa", "Vehículos en Grúa"),
        ("Faltas Administrativas / Migratorias / C.V.I.", "Faltas Admin / Indoc"),
        ("Armas de Fuego Verificadas (Con Permiso)", "Armas Verificadas"),
        ("Personas Verificadas (Conteo Depurado)", "Personas Verificadas"),
        ("Vehículos Verificados (Conteo Depurado)", "Vehículos Verificados"),
    ]

    for r_i, (label, col_name) in enumerate(metricas_resumen, start=5):
        ws_res.row_dimensions[r_i].height = 24
        fill = fill_zebra if r_i % 2 == 1 else fill_white
        c0 = ws_res.cell(row=r_i, column=1, value=label)
        c0.font = font_bold
        c0.fill = fill
        c0.border = border_data
        c0.alignment = align_left

        for b_i, b_name in enumerate(BLOQUES_ORDEN, start=2):
            val = int(df[df["Bloque Horario"] == b_name][col_name].sum())
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
    chart.title = "Infracciones, Capturas y Grúas por Bloque Horario"
    data_ref = Reference(ws_res, min_col=2, min_row=4, max_col=5, max_row=7)
    cats_ref = Reference(ws_res, min_col=1, min_row=5, max_row=7)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    chart.width = 16
    chart.height = 8
    ws_res.add_chart(chart, "A14")

    for col_l, w in zip(["A", "B", "C", "D", "E", "F"], [42, 20, 20, 20, 20, 18]):
        ws_res.column_dimensions[col_l].width = w

    pestanas = [
        ("Solicitados y Faltas", df[(df["Oficios / Capturas"] > 0) | (df["Faltas Admin / Indoc"] > 0)]),
        ("Infracciones y Grúas", df[(df["Infracciones Tránsito"] > 0) | (df["Vehículos en Grúa"] > 0)]),
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
                c.alignment = align_center if c_i in (1, 2, 3, 7, 8, 9, 10, 11, 12, 13, 15, 16) else align_left

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

def generar_html_presentacion(df: pd.DataFrame) -> str:
    tot_inf = int(df["Infracciones Tránsito"].sum())
    tot_ofi = int(df["Oficios / Capturas"].sum())
    tot_fal = int(df["Faltas Admin / Indoc"].sum())
    tot_gru = int(df["Vehículos en Grúa"].sum())
    tot_per = int(df["Personas Verificadas"].sum())
    tot_veh = int(df["Vehículos Verificados"].sum())
    tot_arm = int(df["Armas Verificadas"].sum())

    pn_inf = int(df[df["Institución"] == "Policía Nacional"]["Infracciones Tránsito"].sum())
    snf_inf = int(df[df["Institución"] == "SENAFRONT"]["Infracciones Tránsito"].sum())
    pn_ofi = int(df[df["Institución"] == "Policía Nacional"]["Oficios / Capturas"].sum())
    snf_ofi = int(df[df["Institución"] == "SENAFRONT"]["Oficios / Capturas"].sum())
    pn_per = int(df[df["Institución"] == "Policía Nacional"]["Personas Verificadas"].sum())
    snf_per = int(df[df["Institución"] == "SENAFRONT"]["Personas Verificadas"].sum())
    pn_veh = int(df[df["Institución"] == "Policía Nacional"]["Vehículos Verificados"].sum())
    snf_veh = int(df[df["Institución"] == "SENAFRONT"]["Vehículos Verificados"].sum())

    # Desglose de tipos de faltas de tránsito
    faltas_trans = (
        df[df["Infracciones Tránsito"] > 0]
        .groupby("Tipo / Detalle de Falta")["Infracciones Tránsito"]
        .sum()
        .sort_values(ascending=False)
    )
    filas_faltas_html = "".join(
        f"<tr><td>{k}</td><td style='text-align:center;font-weight:bold;'>{int(v)}</td></tr>"
        for k, v in faltas_trans.items()
    )

    # Desglose de otras acciones / faltas admin / grúas
    otras_acciones = df[(df["Oficios / Capturas"] > 0) | (df["Faltas Admin / Indoc"] > 0) | (df["Vehículos en Grúa"] > 0) | (df["Armas Verificadas"] > 0)]
    filas_relevantes_html = "".join(
        f"<tr><td>{r['Bloque Horario']}</td><td>{r['Institución']}</td><td>{r['Punto de Control']}</td><td><strong>{r['Tipo / Detalle de Falta']}</strong></td><td>{r['Narrativa Operativa'][:160]}...</td></tr>"
        for _, r in otras_acciones.head(15).iterrows()
    )

    return f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>Presentación Ejecutiva — Operación Vértice (CON-C5)</title>
<style>
  body {{ font-family: 'Segoe UI', Arial, sans-serif; background: #F0F4F8; color: #1A202C; margin: 0; padding: 24px; }}
  .header {{ background: #1B365D; color: #FFF; padding: 24px 32px; border-radius: 10px; border-bottom: 5px solid #319795; margin-bottom: 24px; }}
  .header h1 {{ margin: 0; font-size: 26px; letter-spacing: 0.5px; }}
  .header p {{ margin: 6px 0 0; color: #CBD5E0; font-size: 13px; }}
  .kpi-grid {{ display: grid; grid-template-columns: repeat(6, 1fr); gap: 14px; margin-bottom: 24px; }}
  .kpi {{ background: #FFF; padding: 16px; border-radius: 8px; border-left: 5px solid #1B365D; box-shadow: 0 2px 6px rgba(0,0,0,0.06); }}
  .kpi.red {{ border-left-color: #9B2C2C; }}
  .kpi.teal {{ border-left-color: #319795; }}
  .kpi.amber {{ border-left-color: #B7791F; }}
  .kpi .num {{ font-size: 28px; font-weight: 800; color: #1B365D; }}
  .kpi.red .num {{ color: #9B2C2C; }}
  .kpi.teal .num {{ color: #285E61; }}
  .kpi.amber .num {{ color: #975A16; }}
  .kpi .lbl {{ font-size: 12px; font-weight: 700; margin-top: 4px; text-transform: uppercase; }}
  .kpi .sub {{ font-size: 11px; color: #4A5568; margin-top: 4px; }}
  .section {{ background: #FFF; padding: 20px 24px; border-radius: 8px; box-shadow: 0 2px 6px rgba(0,0,0,0.06); margin-bottom: 20px; }}
  .section h2 {{ margin: 0 0 14px; font-size: 17px; color: #1B365D; border-bottom: 2px solid #E2E8F0; padding-bottom: 6px; }}
  .two-col {{ display: grid; grid-template-columns: 1fr 1.6fr; gap: 20px; }}
  table {{ width: 100%; border-collapse: collapse; font-size: 12px; }}
  th {{ background: #2C5282; color: #FFF; padding: 8px 10px; text-align: left; }}
  td {{ padding: 7px 10px; border-bottom: 1px solid #E2E8F0; }}
  tr:nth-child(even) {{ background: #F7FAFC; }}
  @media print {{ body {{ background: #FFF; padding: 0; }} .section {{ box-shadow: none; border: 1px solid #CBD5E0; }} }}
</style>
</head>
<body>
  <div class="header">
    <h1>MINSEG | CENTRO DE OPERACIÓN NACIONAL (CON-C5) — OPERACIÓN VÉRTICE</h1>
    <p>Informe Ejecutivo de Resultados Operativos • Generado el {datetime.now().strftime('%d/%m/%Y %H:%M hrs')}</p>
  </div>

  <div class="kpi-grid">
    <div class="kpi red"><div class="num">{tot_ofi:,}</div><div class="lbl">Solicitados / Oficios</div><div class="sub">PN: {pn_ofi} | SNF: {snf_ofi}</div></div>
    <div class="kpi"><div class="num">{tot_inf:,}</div><div class="lbl">Infracciones Tránsito</div><div class="sub">PN: {pn_inf} | SNF: {snf_inf}</div></div>
    <div class="kpi amber"><div class="num">{tot_gru:,}</div><div class="lbl">Vehículos en Grúa</div><div class="sub">Remociones por faltas graves</div></div>
    <div class="kpi amber"><div class="num">{tot_fal:,}</div><div class="lbl">Faltas Admin / Indoc</div><div class="sub">SNM, MINSA, C.V.I. y Armas ({tot_arm})</div></div>
    <div class="kpi teal"><div class="num">{tot_per:,}</div><div class="lbl">Personas Verificadas</div><div class="sub">PN: {pn_per:,} | SNF: {snf_per:,}</div></div>
    <div class="kpi teal"><div class="num">{tot_veh:,}</div><div class="lbl">Vehículos Verificados</div><div class="sub">PN: {pn_veh:,} | SNF: {snf_veh:,}</div></div>
  </div>

  <div class="two-col">
    <div class="section">
      <h2>Desglose por Tipo de Infracción de Tránsito</h2>
      <table>
        <thead><tr><th>Tipo de Falta / Boleta</th><th style="text-align:center;">Cantidad</th></tr></thead>
        <tbody>{filas_faltas_html}</tbody>
      </table>
    </div>
    <div class="section">
      <h2>Acciones Relevantes: Oficios, Faltas Administrativas y Grúas</h2>
      <table>
        <thead><tr><th>Bloque</th><th>Institución</th><th>Punto (BO)</th><th>Acción / Falta</th><th>Detalle</th></tr></thead>
        <tbody>{filas_relevantes_html}</tbody>
      </table>
    </div>
  </div>
</body>
</html>"""

# ==============================================================================
# INTERFAZ PRINCIPAL STREAMLIT
# ==============================================================================
tab_proc, tab_dash = st.tabs(["📥 1. Procesador de Reportes y Excel", "📊 2. Presentación Ejecutiva al Mando"])

with tab_proc:
    texto_input = st.text_area(
        "Pega aquí los reportes operativos crudos (WhatsApp / Telegram / PDF):",
        height=250,
        placeholder="- BO-53-Chepo (Entrada principal)\n*3RA BRIGADA PANAMÁ ESTE/ BAT-CHEPO*..."
    )

    if st.button("⚙️ Procesar Reportes y Actualizar Presentación", type="primary"):
        if not texto_input.strip():
            st.warning("Pega los textos operativos antes de procesar.")
        else:
            df_res = parsear_reportes(texto_input)
            st.session_state["df_vertice"] = df_res
            st.success(f"✅ {len(df_res)} registros únicos procesados y depurados exitosamente.")

    if "df_vertice" in st.session_state and not st.session_state["df_vertice"].empty:
        df_res = st.session_state["df_vertice"]

        # Fila 1 de KPIs (6 indicadores clave)
        k1, k2, k3, k4, k5, k6 = st.columns(6)
        k1.metric("📄 Solicitados / Oficios", int(df_res["Oficios / Capturas"].sum()))
        k2.metric("🚦 Infracciones Tránsito", int(df_res["Infracciones Tránsito"].sum()))
        k3.metric("🚛 Vehículos en Grúa", int(df_res["Vehículos en Grúa"].sum()))
        k4.metric("⚖️ Faltas Admin / Indoc", int(df_res["Faltas Admin / Indoc"].sum()))
        k5.metric("👥 Personas Verificadas", f"{int(df_res['Personas Verificadas'].sum()):,}")
        k6.metric("🚘 Vehículos Verificados", f"{int(df_res['Vehículos Verificados'].sum()):,}")

        excel_bytes = generar_excel_5_pestanas(df_res)
        st.download_button(
            label="📥 Descargar Cuadro Consolidado Excel (4 Bloques / 5 Pestañas)",
            data=excel_bytes,
            file_name=f"Operacion_Vertice_Consolidado_{datetime.now().strftime('%d%b%y_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.dataframe(df_res, use_container_width=True)

with tab_dash:
    if "df_vertice" not in st.session_state or st.session_state["df_vertice"].empty:
        st.info("Procesa primero un bloque de reportes en la pestaña '1. Procesador de Reportes y Excel' para desplegar la presentación.")
    else:
        df_res = st.session_state["df_vertice"]

        st.subheader("🎯 Tablero Ejecutivo — Resultados Relevantes Operación Vértice")
        m1, m2, m3, m4, m5, m6, m7 = st.columns(7)
        m1.metric("Oficios / Capturas", int(df_res["Oficios / Capturas"].sum()))
        m2.metric("Infracciones Tránsito", int(df_res["Infracciones Tránsito"].sum()))
        m3.metric("Vehículos en Grúa", int(df_res["Vehículos en Grúa"].sum()))
        m4.metric("Faltas Admin / Indoc", int(df_res["Faltas Admin / Indoc"].sum()))
        m5.metric("Armas Verificadas", int(df_res["Armas Verificadas"].sum()))
        m6.metric("Personas Verificadas", f"{int(df_res['Personas Verificadas'].sum()):,}")
        m7.metric("Vehículos Verificados", f"{int(df_res['Vehículos Verificados'].sum()):,}")

        html_pres = generar_html_presentacion(df_res)
        st.download_button(
            label="🖥️ Descargar Presentación Ejecutiva (HTML / Imprimir a PDF)",
            data=html_pres.encode("utf-8"),
            file_name=f"Presentacion_Operacion_Vertice_{datetime.now().strftime('%d%b%y_%H%M')}.html",
            mime="text/html"
        )

        st.divider()
        col_izq, col_der = st.columns(2)

        with col_izq:
            st.markdown("#### 🚦 Desglose por Tipo de Infracción de Tránsito")
            df_faltas = (
                df_res[df_res["Infracciones Tránsito"] > 0]
                .groupby(["Tipo / Detalle de Falta", "Institución"], as_index=False)["Infracciones Tránsito"]
                .sum()
                .sort_values(by="Infracciones Tránsito", ascending=False)
            )
            st.dataframe(df_faltas, use_container_width=True, hide_index=True)

        with col_der:
            st.markdown("#### 📊 Consolidado por Bloque Horario e Institución")
            df_bloques = df_res.groupby(["Bloque Horario", "Institución"], as_index=False)[
                ["Oficios / Capturas", "Infracciones Tránsito", "Vehículos en Grúa", "Faltas Admin / Indoc", "Personas Verificadas", "Vehículos Verificados"]
            ].sum()
            st.dataframe(df_bloques, use_container_width=True, hide_index=True)

        st.markdown("#### 🚨 Detalle de Acciones Relevantes (Oficios, Faltas Administrativas, Grúas y Armas)")
        df_relevante = df_res[
            (df_res["Oficios / Capturas"] > 0) |
            (df_res["Faltas Admin / Indoc"] > 0) |
            (df_res["Vehículos en Grúa"] > 0) |
            (df_res["Armas Verificadas"] > 0)
        ][["Bloque Horario", "Institución", "Punto de Control", "Categoría", "Tipo / Detalle de Falta", "Documento / C.I.P.", "Placa", "Narrativa Operativa"]]
        st.dataframe(df_relevante, use_container_width=True, hide_index=True)
