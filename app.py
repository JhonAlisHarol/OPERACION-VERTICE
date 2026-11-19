import re
import io
import hashlib
from datetime import datetime
import pandas as pd
import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule

# ==============================================================================
# CONFIGURACIÓN DE PÁGINA — CON-C5 PANAMÁ (MINSEG)
# ==============================================================================
st.set_page_config(
    page_title="CONSOLIDADO 24 HORAS — Operación Vértice | CON-C5",
    page_icon="🛡️",
    layout="wide"
)

# ==============================================================================
# VALORES BASE AUDITADOS (SIN DUPLICADOS 108 / 12 — 24 HORAS)
# ==============================================================================
BASE_VALORES = {
    "Consolidado": {
        "oficios": 14,
        "infracciones": 136,
        "gruas": 3,
        "faltas_admin": 7,
        "personas": 4009,
        "vehiculos": 2738,
        "otras": 0,
    },
    "Corte 1 (06:00-11:59)": {
        "oficios": 3,
        "infracciones": 14,
        "gruas": 0,
        "faltas_admin": 0,
        "personas": 0,
        "vehiculos": 0,
        "otras": 0,
    },
    "Corte 2 (12:00-17:25)": {
        "oficios": 4,
        "infracciones": 62,
        "gruas": 2,
        "faltas_admin": 5,
        "personas": 2268,
        "vehiculos": 1741,
        "otras": 0,
    },
    "Corte 3 (17:26-23:59)": {
        "oficios": 6,
        "infracciones": 54,
        "gruas": 1,
        "faltas_admin": 2,
        "personas": 1528,
        "vehiculos": 852,
        "otras": 0,
    },
    "Corte 4 (00:00-06:00)": {
        "oficios": 1,
        "infracciones": 6,
        "gruas": 0,
        "faltas_admin": 0,
        "personas": 213,
        "vehiculos": 145,
        "otras": 0,
    },
}

CORTES_ORDEN = [
    "Corte 1 (06:00-11:59)",
    "Corte 2 (12:00-17:25)",
    "Corte 3 (17:26-23:59)",
    "Corte 4 (00:00-06:00)"
]

SENAFRONT_BOS = {
    "BO-4", "BO-5", "BO-6", "BO-11", "BO-15", "BO-53", "BO-54", "BO-59", "BO-60", "BO-61"
}

def extraer_texto_archivos(archivos_subidos) -> str:
    """Extrae texto de archivos PDF, Word (.docx), Excel (.xlsx/.xls), CSV o TXT."""
    if not archivos_subidos:
        return ""
    partes = []
    for f in archivos_subidos:
        nombre = (f.name or "").lower()
        raw = f.getvalue()
        if nombre.endswith(".pdf"):
            try:
                import pypdf
                reader = pypdf.PdfReader(io.BytesIO(raw))
                for pg in reader.pages:
                    t = pg.extract_text() or ""
                    if t.strip():
                        partes.append(t)
            except Exception as e:
                partes.append(f"[Error leyendo PDF {f.name}: {e}]")
        elif nombre.endswith(".docx"):
            try:
                import docx
                doc = docx.Document(io.BytesIO(raw))
                for p in doc.paragraphs:
                    if p.text.strip():
                        partes.append(p.text)
                for tbl in doc.tables:
                    for row in tbl.rows:
                        partes.append(" | ".join(cell.text.strip() for cell in row.cells if cell.text.strip()))
            except Exception as e:
                partes.append(f"[Error leyendo DOCX {f.name}: {e}]")
        elif nombre.endswith(".xlsx") or nombre.endswith(".xls"):
            try:
                xls = pd.ExcelFile(io.BytesIO(raw))
                for sh in xls.sheet_names:
                    df_s = pd.read_excel(xls, sheet_name=sh)
                    partes.append(f"--- HOJA: {sh} ---")
                    partes.append(df_s.fillna("").to_string(index=False))
            except Exception as e:
                partes.append(f"[Error leyendo Excel {f.name}: {e}]")
        elif nombre.endswith(".csv"):
            try:
                df_c = pd.read_csv(io.BytesIO(raw))
                partes.append(df_c.fillna("").to_string(index=False))
            except Exception:
                partes.append(raw.decode("utf-8", errors="ignore"))
        else:
            partes.append(raw.decode("utf-8", errors="ignore"))
    return "\n".join(partes)

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
    m_mil = re.search(r"\b(\d{2})(\d{2})(\d{2})NOV\d{2,4}\b", t_up)
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

def enmascarar_identificadores(texto_up: str) -> str:
    s = re.sub(r"\b\d{2}\d{2}\d{2}NOV\d{2,4}\b", " [TS] ", texto_up)
    s = re.sub(r"\b\d{1,2}/\d{1,2}/\d{2,4}\b", " [FECHA] ", s)
    s = re.sub(r"\b\d{1,2}:\d{2}(?:\s*(?:A\.?M\.?|P\.?M\.?|HRS?))?\b", " [HORA] ", s)
    s = re.sub(r"B/\.?\s*\d+(?:\.\d{1,2})?", " [MONTO] ", s)
    s = re.sub(r"\b(?:[A-Z]-)?\d{1,2}-\d{3,5}-\d{3,5}\b", " [ID] ", s)
    s = re.sub(
        r"(?:BOLETA|INFRACCI[ÓO]N|FALTA|ART[ÍI]CULO|ARTICULO|C[ÓO]DIGO|OFICIO|SUBROGADO)\s*(?:N[°º#.\s-]*|INCISO\s*[A-Z]?\s*)?[A-Z0-9-]*",
        " [COD] ",
        s
    )
    s = re.sub(
        r"(?:PLACA|MATR[ÍI]CULA|MATRICULA|MET[ÁA]LICA|METALICA)\s*(?:N[°º#.\s-]*)?[A-Z]{1,3}\d{1,6}|\b[A-Z]{2}\d{3,6}\b|\b\d{6}\b",
        " [PLACA] ",
        s
    )
    return s

def calcular_huella_exacta(texto: str) -> str:
    limpio = re.sub(r"[^a-z0-9áéíóúñü]+", " ", texto.lower()).strip()
    limpio = re.sub(r"\s+", " ", limpio)
    return hashlib.md5(limpio.encode("utf-8")).hexdigest()

def extraer_claves_entidad(corte: str, bo_code: str, b_up: str) -> list[str]:
    claves = []
    for ced in re.findall(r"\b(?:[A-Z]-)?\d{1,2}-\d{3,5}-\d{3,5}\b", b_up):
        claves.append(f"{corte}|{bo_code}|ID:{ced}")
    for pas in re.findall(r"\b[A-Z]{2}\d{5,8}\b", b_up):
        claves.append(f"{corte}|{bo_code}|PAS:{pas}")
    for pl in re.findall(r"(?:PLACA|MATR[ÍI]CULA|MATRICULA|MET[ÁA]LICA|METALICA)\s*(?:N[°º#.\s-]*)?([A-Z]{1,3}\d{1,6}|\d{4,6})\b", b_up):
        claves.append(f"{corte}|{bo_code}|PLACA:{pl}")
    for bol in re.findall(r"(?:BOLETA|OFICIO|SUBROGADO)\s*(?:N[°º#.\s-]*)?([A-Z0-9-]{2,20})\b", b_up):
        if bol not in {"DE", "DEL", "POR", "SIN", "EN", "LA", "EL", "NO", "SI"}:
            claves.append(f"{corte}|{bo_code}|DOC:{bol}")
    return claves

def analizar_reportes_texto(texto_crudo: str) -> tuple[pd.DataFrame, dict]:
    bloques = re.split(r"(?=\bBO-\d+-)", texto_crudo.strip())
    filas = []

    huellas_vistas = set()
    entidades_vistas = set()
    semantica_sin_id = set()

    stats = {
        "total_bloques": 0,
        "duplicados_exactos": 0,
        "duplicados_entidad": 0,
        "duplicados_semanticos": 0,
        "rutina_o_vacio": 0,
        "unicos_validos": 0,
    }

    for b in bloques:
        b = b.strip()
        if not b or not b.startswith("BO-"):
            continue
        stats["total_bloques"] += 1

        huella = calcular_huella_exacta(b)
        if huella in huellas_vistas:
            stats["duplicados_exactos"] += 1
            continue
        huellas_vistas.add(huella)

        lineas = [l.strip() for l in b.splitlines() if l.strip()]
        encabezado = lineas[0]
        m_bo = re.match(r"^(BO-\d+)\s*-\s*([^\n]+)", encabezado)
        bo_code = m_bo.group(1) if m_bo else "BO-S/N"
        corte = detectar_corte(b)
        inst = detectar_institucion(bo_code, b)

        b_up = b.upper()
        b_mask = enmascarar_identificadores(b_up)

        es_rutina = bool(re.search(r"NOS MANTENEMOS\s+(?:EN\s+LA\s+)?REALIZANDO\s+(?:UNA\s+)?BOLETA", b_up)) and not any(
            tok in b_mask for tok in ["[PLACA]", "[ID]", "[COD]"]
        )
        if es_rutina or any(k in b_up for k in ["SE PUEDEN REGISTRAR INFRACCIONES", "NO SE MANTIENE BASTÓN", "NO SE MANTIENE BASTON"]):
            stats["rutina_o_vacio"] += 1
            continue

        claves_ent = extraer_claves_entidad(corte, bo_code, b_up)
        if claves_ent and all(k in entidades_vistas for k in claves_ent):
            stats["duplicados_entidad"] += 1
            continue
        for k in claves_ent:
            entidades_vistas.add(k)

        oficios = 1 if any(k in b_up for k in [
            "OFICIO DE CONDUCCI", "OFICIO DE CAPTURA", "OFICIO JUDICIAL", "OFICIO DE SUBROGADO",
            "CONDUCCI[ÓO]N JUDICIAL", "APREHENSI[ÓO]N POR OFICIO", "CINTILLO ROJO", "OFICIO N"
        ]) else 0

        faltas_admin = 1 if any(k in b_up for k in [
            "INDOCUMENTADO", "SIN DOCUMENTO", "ESTAD[ÍI]A VENCIDA", "MIGRACI[ÓO]N", "S.N.M", "SNM",
            "ACTA DE INSPECCI[ÓO]N SANITARIA", "MINSA", "BOLETA DE CITACI[ÓO]N", "C.V.I", "ETRM"
        ]) else 0

        m_grua_cnt = re.search(r"(?:REMOCI[ÓO]N\s+EN\s+GR[ÚU]A|VEH[ÍI]CULOS?\s+EN\s+GR[ÚU]A|REMOVIDO\s+EN\s+GR[ÚU]A)\s*[:=\s-]*(\d{1,2})\b", b_mask)
        if m_grua_cnt:
            gruas = int(m_grua_cnt.group(1))
        elif any(k in b_up for k in ["REMOCI[ÓO]N EN GR[ÚU]A", "REMOVIDO EN GR[ÚU]A", "VEH[ÍI]CULO EN GR[ÚU]A", "EN LA GR[ÚU]A"]):
            gruas = 1
        else:
            gruas = 0

        infracciones = 0
        if not re.search(r"(?:INFRACCI[ÓO]N\s+DE\s+TR[ÁA]NSITO|BOLETAS\s+PUESTAS)\s*[:=\s(-]*0+\b", b_mask):
            m_cnt_antes = re.findall(r"\b(\d{1,3})\s*(?:BOLETAS?|INFRACCIONES?|SANCIONES?)\b", b_mask)
            m_cnt_despues = re.findall(r"\b(?:BOLETAS?|INFRACCIONES?\s+DE\s+TR[ÁA]NSITO|SANCIONES?)\s*[:=\s-]+(\d{1,3})\b", b_mask)
            if m_cnt_despues:
                infracciones = sum(int(x) for x in m_cnt_despues)
            elif m_cnt_antes:
                infracciones = sum(int(x) for x in m_cnt_antes)
            elif any(k in b_up for k in [
                "SE COLOCA INFRACCI", "SE REALIZA INFRACCI", "SE REALIZA BOLETA",
                "SE SANCION", "ES SANCION", "BOLETA DE INFRACCI", "INFRACCI[ÓO]N DE TR[ÁA]NSITO",
                "SE LE RETIRA LICENCIA", "SE LE RETIENE LICENCIA", "EMBRIAGUEZ COMPROBADA"
            ]):
                infracciones = 1

        if infracciones == 1 and not any(tok in b_mask for tok in ["[PLACA]", "[ID]"]):
            firma_sem = f"{corte}|{bo_code}|{re.sub(r'[^A-Z0-9]', '', b_mask[:80])}"
            if firma_sem in semantica_sin_id:
                stats["duplicados_semanticos"] += 1
                infracciones = 0
            else:
                semantica_sin_id.add(firma_sem)

        personas = 0
        for pat in [
            r"\bPERSONAS(?:\s+VERIFICADAS)?\s*[:=\s.-]*(\d{1,4})\b",
            r"\bCIUDADANOS(?:\s+NACIONALES|\s+EXTRANJEROS)?\s*[:=\s.-]*(\d{1,4})\b",
            r"\b(\d{1,4})\s+CIUDADANOS?\s+(?:NACIONALES?|EXTRANJEROS?|VERIFICADOS?)\b",
        ]:
            found_p = re.findall(pat, b_mask)
            if found_p:
                personas = max(personas, sum(int(x) for x in found_p))

        vehiculos = 0
        for pat in [
            r"\bVEH[ÍI]CULOS(?:\s+VERIFICADOS|\s+PARTICULARES|\s+NACIONALES|\s+EXTRANJEROS)?\s*[:=\s.-]*(\d{1,4})\b",
            r"\bMOTOS?(?:\s+VERIFICADAS?)?\s*[:=\s.-]*(\d{1,4})\b",
            r"\bTAXIS?\s*[:=\s.-]*(\d{1,4})\b",
            r"\bBUSES?\s*[:=\s.-]*(\d{1,4})\b",
            r"\bCAMIONES?\s*[:=\s.-]*(\d{1,4})\b",
        ]:
            found_v = re.findall(pat, b_mask)
            if found_v:
                vehiculos += sum(int(x) for x in found_v)

        otras = 1 if (oficios == 0 and faltas_admin == 0 and gruas == 0 and infracciones == 0 and personas == 0 and vehiculos == 0 and any(
            k in b_up for k in ["ARMA", "PERMISO", "INSPECCI", "PROFILAXIS", "325", "NARC[ÓO]TICOS", "ADUANAS", "MIDA"]
        )) else 0

        stats["unicos_validos"] += 1
        filas.append({
            "Corte": corte,
            "Institución": inst,
            "Sector (BO)": encabezado,
            "Oficios / Capturas": oficios,
            "Infracciones": infracciones,
            "Grúas": gruas,
            "Faltas Admin / Indoc": faltas_admin,
            "Personas Verificadas": personas,
            "Vehículos Verificados": vehiculos,
            "Otras Acciones": otras,
            "Detalle": " ".join(lineas[1:])[:320]
        })

    return pd.DataFrame(filas), stats

def generar_excel_consolidado(df_extra: pd.DataFrame | None = None, modo_carga: str = "Sumar como reporte nuevo") -> bytes:
    wb = openpyxl.Workbook()

    HEADER_DARK = "23395D"
    SUBHEADER_STEEL = "3A608F"
    SECTION_SAGE = "2E5A53"
    ZEBRA_LIGHT = "F7F9FC"
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
    fill_tot = PatternFill(start_color=TOTAL_GRAY, end_color=TOTAL_GRAY, fill_type="solid")

    thin = Side(style="thin", color=BORDER_GRAY)
    med_navy = Side(style="medium", color=HEADER_DARK)
    double_navy = Side(style="double", color=HEADER_DARK)
    border_cell = Border(left=thin, right=thin, top=thin, bottom=thin)
    border_tot = Border(left=thin, right=thin, top=med_navy, bottom=double_navy)

    al_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    al_center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws_res = wb.active
    ws_res.title = "Resumen Ejecutivo 4 Cortes"
    ws_res.views.sheetView[0].showGridLines = True

    ws_res.merge_cells("A1:H1")
    ws_res.merge_cells("A2:H2")
    ws_res["A1"] = "OPERACIÓN VÉRTICE — CONSOLIDADO EXACTO 4 CORTES (SIN DUPLICADOS 108 / 12)"
    ws_res["A2"] = "18 al 19 de noviembre de 2026 | Corte 1 (06:00-11:59) | Corte 2 (12:00-17:25) | Corte 3 (17:26-23:59) | Corte 4 (00:00-06:00)"
    ws_res["A1"].font = font_title
    ws_res["A2"].font = font_sub
    for r in (1, 2):
        for c in range(1, 9):
            ws_res.cell(row=r, column=c).fill = fill_title
            ws_res.cell(row=r, column=c).alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws_res.row_dimensions[1].height = 26
    ws_res.row_dimensions[2].height = 18

    hdrs = [
        "Categoría / Tipo de Acción", "Corte 1\n(06:00 - 11:59)", "Corte 2\n(12:00 - 17:25)",
        "Corte 3\n(17:26 - 23:59)", "Corte 4\n(00:00 - 06:00)", "Total Exacto\n(24 Horas)",
        "Distribución Institucional", "Notas de Control"
    ]
    ws_res.row_dimensions[4].height = 28
    for idx, h in enumerate(hdrs, 1):
        c = ws_res.cell(row=4, column=idx, value=h)
        c.font = font_hdr
        c.fill = fill_hdr
        c.alignment = al_center
        c.border = border_cell

    reemplazar = (modo_carga == "Reemplazar valores base") and (df_extra is not None)

    def val_celda(base_v, col_k, corte_lbl):
        ex_v = int(df_extra[df_extra["Corte"] == corte_lbl][col_k].sum()) if (df_extra is not None and not df_extra.empty) else 0
        return ex_v if reemplazar else (base_v + ex_v)

    filas_res = [
        ("Personas Solicitadas / Capturas (Oficios)",
         val_celda(3, "Oficios / Capturas", "Corte 1 (06:00-11:59)"),
         val_celda(4, "Oficios / Capturas", "Corte 2 (12:00-17:25)"),
         val_celda(6, "Oficios / Capturas", "Corte 3 (17:26-23:59)"),
         val_celda(1, "Oficios / Capturas", "Corte 4 (00:00-06:00)"),
         "=SUM(B5:E5)", "8 PN / 6 SNF", "3 + 4 + 6 + 1 = 14 capturas por oficio"),
        ("Infracciones de Tránsito Registradas",
         val_celda(14, "Infracciones", "Corte 1 (06:00-11:59)"),
         val_celda(62, "Infracciones", "Corte 2 (12:00-17:25)"),
         val_celda(54, "Infracciones", "Corte 3 (17:26-23:59)"),
         val_celda(6, "Infracciones", "Corte 4 (00:00-06:00)"),
         "=SUM(B6:E6)", "74 PN / 62 SNF", "14 + 62 + 54 + 6 = 136 (sin 108/12)"),
        ("Vehículos Removidos en Grúa",
         val_celda(0, "Grúas", "Corte 1 (06:00-11:59)"),
         val_celda(2, "Grúas", "Corte 2 (12:00-17:25)"),
         val_celda(1, "Grúas", "Corte 3 (17:26-23:59)"),
         val_celda(0, "Grúas", "Corte 4 (00:00-06:00)"),
         "=SUM(B7:E7)", "3 PN / 0 SNF", "Llano Bonito, Almirante y David BV0753"),
        ("Faltas Administrativas / Migratorias / Indocumentados",
         val_celda(0, "Faltas Admin / Indoc", "Corte 1 (06:00-11:59)"),
         val_celda(5, "Faltas Admin / Indoc", "Corte 2 (12:00-17:25)"),
         val_celda(2, "Faltas Admin / Indoc", "Corte 3 (17:26-23:59)"),
         val_celda(0, "Faltas Admin / Indoc", "Corte 4 (00:00-06:00)"),
         "=SUM(B8:E8)", "SNM / MINSA / PN / SNF", "0 + 5 + 2 + 0 = 7 faltas admin / indocumentados"),
        ("Personas Verificadas (Conteo Depurado)",
         val_celda(0, "Personas Verificadas", "Corte 1 (06:00-11:59)"),
         val_celda(2268, "Personas Verificadas", "Corte 2 (12:00-17:25)"),
         val_celda(1528, "Personas Verificadas", "Corte 3 (17:26-23:59)"),
         val_celda(213, "Personas Verificadas", "Corte 4 (00:00-06:00)"),
         "=SUM(B9:E9)", "PN + SENAFRONT", "2,268 + 1,528 + 213 = 4,009 personas"),
        ("Vehículos Verificados (Conteo Depurado)",
         val_celda(0, "Vehículos Verificados", "Corte 1 (06:00-11:59)"),
         val_celda(1741, "Vehículos Verificados", "Corte 2 (12:00-17:25)"),
         val_celda(852, "Vehículos Verificados", "Corte 3 (17:26-23:59)"),
         val_celda(145, "Vehículos Verificados", "Corte 4 (00:00-06:00)"),
         "=SUM(B10:E10)", "PN + SENAFRONT", "1,741 + 852 + 145 = 2,738 vehículos"),
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

    ws_res.merge_cells("A12:H12")
    ws_res["A12"] = "DESGLOSE DE INFRACCIONES POR SECTOR (BO) — CORTE 3 (54) Y CORTE 4 (6)"
    ws_res["A12"].font = font_sec
    ws_res["A12"].fill = fill_sec
    ws_res["A12"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws_res.row_dimensions[12].height = 24

    hdr_sec = ["Sector / Punto de Control (BO)", "Institución", "Corte 3", "Corte 4", "Total C3+C4", "Detalle Operativo"]
    ws_res.merge_cells("F13:H13")
    ws_res.row_dimensions[13].height = 24
    for idx, h in enumerate(hdr_sec, 1):
        c = ws_res.cell(row=13, column=idx, value=h)
        c.font = font_hdr
        c.alignment = al_center
    for col_c in range(1, 9):
        ws_res.cell(row=13, column=col_c).fill = fill_hdr
        ws_res.cell(row=13, column=col_c).border = border_cell

    sectores = [
        ("BO-54 - Tanara (Panamá Este)", "SENAFRONT", 8, 2, "C3: 8 boletas (Art. 53A/53D/132C/240, #64, #03) | C4: 2 boletas (#14 y Art. 132C alcoholemia)"),
        ("BO-53 - Chepo (Entrada principal)", "SENAFRONT", 6, 0, "C3: 5 nocturnas (#14 x4, Art. 132C alcoholemia) + 1 reubicada de C2 (Art. 132C)"),
        ("BO-61 - Emberá Purú - Arimae (Darién)", "SENAFRONT", 6, 0, "C3: 5 nocturnas (Cód. 46 ebriedad, #16, señales, Art. 53D, #03) + 1 reubicada de C2 (luces reflexivas)"),
        ("BO-3 - Almirante (Entrada al puerto)", "Policía Nacional", 6, 0, "C3: 4 nocturnas SVC (#14 x3, #01 x1) + 2 reubicadas de C2 (Boleta #70 y 1 grúa)"),
        ("BO-24 - Divisa (Entrada)", "Policía Nacional", 6, 0, "C3: 6 boletas por diferentes infracciones al reglamento de tránsito (reubicadas de C2)"),
        ("BO-59 - Agua Fría (Darién)", "SENAFRONT", 5, 0, "C3: Consolidado 5 infracciones verificadas con Tránsito PN (Faltas #32, #70 y #14)"),
        ("BO-52 - Pacora (Las Garzas)", "Policía Nacional", 4, 0, "C3: Boletas #5034933, #5034934 (sin metálica), #5034935 y #5034936 (Inf. #10 y #40)"),
        ("BO-6 - Paso Canoas (Frontera / San Isidro)", "SENAFRONT", 2, 3, "C3: 2 boletas #14 (camión 957064 y Nissan pick-up) | C4: 3 boletas cierre turno 1800-0600 hrs"),
        ("BO-23 - Santiago (Frente a M/S Marticito)", "Policía Nacional", 3, 0, "C3: 3 boletas por luces (1 defectuosa + 2 inadecuadas)"),
        ("BO-60 - Metetí (Darién)", "SENAFRONT", 3, 0, "C3: 3 boletas nocturnas (Art. 53A luces inadecuadas + Art. 53D luces defectuosas x2)"),
        ("BO-56 - Corredor Sur (Entrada Llano Bonito)", "Policía Nacional", 2, 0, "C3: Alberto Blanco (CX4552 sin seguro) + 1 reubicada de C2 (Jomil Antala #59 giro indebido)"),
        ("BO-15 - Frontera Paso Canoas - Sector Bugaba (Portón)", "SENAFRONT", 0, 1, "C4: Infracción de tránsito #32 por exceso de pasajeros (04:25 hrs del 19NOV26)"),
        ("BO-19 - Atalaya (Ponuga, Veraguas)", "Policía Nacional", 1, 0, "C3: Boleta #3500092597418112026 — Código 735 (Orlando Enrique Silva, moto NC6634)"),
        ("BO-39 - Capira (Pueblo)", "Policía Nacional", 1, 0, "C3: Infracción #21 vidrios ahumados (Raúl Álvarez, Toyota Coaster 837182)"),
        ("BO-13 - Guadalupe (David)", "Policía Nacional", 1, 0, "C3: 01 vehículo Hyundai Accent blanco (BV0753) removido en grúa por desacato y licencia vencida"),
    ]

    r_idx = 14
    for i_s, (s_nom, s_inst, c3_v, c4_v, s_det) in enumerate(sectores):
        ws_res.row_dimensions[r_idx].height = 24
        ws_res.merge_cells(start_row=r_idx, start_column=6, end_row=r_idx, end_column=8)
        fl = fill_zebra if i_s % 2 == 1 else fill_white

        ws_res.cell(row=r_idx, column=1, value=s_nom).font = font_bld
        ws_res.cell(row=r_idx, column=2, value=s_inst).font = font_reg
        ws_res.cell(row=r_idx, column=3, value=c3_v).font = font_bld
        ws_res.cell(row=r_idx, column=4, value=c4_v).font = font_bld
        ws_res.cell(row=r_idx, column=5, value=f"=C{r_idx}+D{r_idx}").font = font_bld
        ws_res.cell(row=r_idx, column=6, value=s_det).font = font_reg

        for col_c in range(1, 9):
            cell = ws_res.cell(row=r_idx, column=col_c)
            cell.fill = fl
            cell.border = border_cell
            if col_c in (2, 3, 4, 5):
                cell.alignment = al_center
                if col_c in (3, 4, 5):
                    cell.number_format = "#,##0"
            else:
                cell.alignment = al_left
        r_idx += 1

    ws_res.row_dimensions[r_idx].height = 26
    ws_res.merge_cells(start_row=r_idx, start_column=6, end_row=r_idx, end_column=8)
    ws_res.cell(row=r_idx, column=1, value="TOTAL POR SECTOR (CORTES 3 Y 4)").font = font_tot
    ws_res.cell(row=r_idx, column=2, value="PN + SNF").font = font_tot
    ws_res.cell(row=r_idx, column=3, value=f"=SUM(C14:C{r_idx-1})").font = font_tot
    ws_res.cell(row=r_idx, column=4, value=f"=SUM(D14:D{r_idx-1})").font = font_tot
    ws_res.cell(row=r_idx, column=5, value=f"=SUM(E14:E{r_idx-1})").font = font_tot
    ws_res.cell(row=r_idx, column=6, value="Corte 3: 54 | Corte 4: 6 | Total C3+C4: 60 (Total 24h con C1: 14 y C2: 62 = 136)").font = font_tot
    for col_c in range(1, 9):
        cell = ws_res.cell(row=r_idx, column=col_c)
        cell.fill = fill_tot
        cell.border = border_tot
        if col_c in (2, 3, 4, 5):
            cell.alignment = al_center
            if col_c in (3, 4, 5):
                cell.number_format = "#,##0"
        else:
            cell.alignment = al_left

    rule_pos = CellIsRule(operator="greaterThan", formula=["0"], stopIfTrue=False,
                          fill=PatternFill(start_color="E6F4EA", end_color="E6F4EA", fill_type="solid"))
    ws_res.conditional_formatting.add("B5:E10", rule_pos)

    for col_l, w in zip(["A", "B", "C", "D", "E", "F", "G", "H"], [42, 18, 16, 16, 16, 22, 22, 24]):
        ws_res.column_dimensions[col_l].width = w

    ws_res.page_setup.orientation = ws_res.ORIENTATION_LANDSCAPE
    ws_res.page_setup.paperSize = ws_res.PAPERSIZE_LETTER
    ws_res.sheet_properties.pageSetUpPr.fitToPage = True
    ws_res.page_setup.fitToWidth = 1
    ws_res.page_setup.fitToHeight = 0

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
  }
  .top-banner-title { color: #FFFFFF; font-size: 20px; font-weight: 800; letter-spacing: 0.5px; margin: 0; }
  .top-banner-sub { color: #93C5FD; font-size: 12px; margin: 3px 0 0 0; }
  .hero-header {
    background: #051329; text-align: center; padding: 10px 16px;
    border-bottom: 2px solid #1E498A; margin-bottom: 14px;
  }
  .hero-title { color: #FFFFFF; font-size: 26px; font-weight: 900; letter-spacing: 1px; margin: 0; }
  .hero-sub { color: #E2E8F0; font-size: 13px; margin: 4px 0 0 0; }
  .panel-box {
    background: #071936; border: 2px solid #00A8E8;
    border-radius: 4px; padding: 0 0 10px 0; margin-bottom: 14px;
  }
  .panel-box.green { border-color: #22C55E; }
  .panel-box.amber { border-color: #F59E0B; }
  .panel-box.orange { border-color: #F97316; }
  .panel-hdr {
    background: #0D2B5C; color: #FFFFFF; font-weight: 800; font-size: 14px;
    text-align: center; padding: 8px; border-bottom: 1px solid #1E498A; text-transform: uppercase;
  }
  .kpi-big { font-size: 40px; font-weight: 900; text-align: center; line-height: 1.1; margin-top: 10px; }
  .kpi-blue { color: #38BDF8; }
  .kpi-green { color: #4ADE80; }
  .kpi-amber { color: #FBBF24; }
  .kpi-orange { color: #FB923C; }
  .kpi-lbl { color: #FFFFFF; font-size: 12px; font-weight: 700; text-align: center; margin-top: 4px; }
  .kpi-foot {
    background: #041024; color: #CBD5E1; font-size: 11px;
    text-align: center; padding: 6px; margin-top: 8px; border-top: 1px solid #163A70;
  }
</style>
""", unsafe_allow_html=True)

if "modo_carga_vertice" not in st.session_state:
    st.session_state["modo_carga_vertice"] = "Sumar como reporte nuevo"

col_b_izq, col_b_der = st.columns([3.4, 1.1])
with col_b_izq:
    st.markdown("""
    <div class="top-banner">
      <div class="top-banner-title">GOBIERNO NACIONAL • CON PASO FIRME — MINISTERIO DE SEGURIDAD PÚBLICA</div>
      <div class="top-banner-sub">CENTRO DE OPERACIÓN NACIONAL (C.O.N. / C5) • OPERACIÓN VÉRTICE</div>
    </div>
    """, unsafe_allow_html=True)
with col_b_der:
    grupo_sel = st.selectbox(
        "GRUPO DE HORAS ▼",
        options=["Consolidado"] + CORTES_ORDEN,
        index=0
    )

st.markdown(f"""
<div class="hero-header">
  <div class="hero-title">CONSOLIDADO 24 HORAS</div>
  <div class="hero-sub">Operación Vértice • Puntos de control • Estamento: Todos • Vista activa: <strong>{grupo_sel}</strong></div>
</div>
""", unsafe_allow_html=True)

tab_pres, tab_proc = st.tabs(["📊 Presentación Ejecutiva (CON-C5)", "📥 Cargar Textos, PDF / Word / Excel"])

with tab_proc:
    modo_carga = st.radio(
        "Modo de actualización al cargar o pegar datos:",
        options=["Sumar como reporte nuevo", "Reemplazar valores base"],
        index=0 if st.session_state["modo_carga_vertice"] == "Sumar como reporte nuevo" else 1,
        horizontal=True
    )
    st.session_state["modo_carga_vertice"] = modo_carga

    archivos_subidos = st.file_uploader(
        "Cargar archivos PDF, Word (.docx), Excel (.xlsx/.xls), CSV o TXT (actualización automática):",
        type=["pdf", "docx", "xlsx", "xls", "csv", "txt"],
        accept_multiple_files=True,
        key="uploader_vertice_auto"
    )
    texto_pegar = st.text_area(
        "O pega aquí directamente los textos de reportes operativos (actualización automática):",
        height=210,
        placeholder="- BO-53-Chepo (Entrada principal)...",
        key="texto_vertice_auto"
    )

    texto_archivos = extraer_texto_archivos(archivos_subidos)
    texto_combinado = "\n".join(x for x in [texto_archivos, texto_pegar] if x and x.strip())

    if texto_combinado.strip():
        df_ex, stats_dup = analizar_reportes_texto(texto_combinado)
        st.session_state["df_vertice_extra"] = df_ex
        st.session_state["stats_dup"] = stats_dup
        st.info(f"⚡ Actualización automática activa ({modo_carga}): {stats_dup['unicos_validos']} registros únicos de {stats_dup['total_bloques']} bloques detectados.")
        d1, d2, d3, d4 = st.columns(4)
        d1.metric("Duplicados Exactos", stats_dup["duplicados_exactos"])
        d2.metric("Duplicados por Placa/ID/Boleta", stats_dup["duplicados_entidad"])
        d3.metric("Duplicados Semánticos", stats_dup["duplicados_semanticos"])
        d4.metric("Rutina / Sin Boleta", stats_dup["rutina_o_vacio"])
        st.dataframe(df_ex, use_container_width=True, hide_index=True)
    else:
        st.session_state["df_vertice_extra"] = None
        st.session_state["stats_dup"] = None

    excel_data = generar_excel_consolidado(st.session_state.get("df_vertice_extra"), st.session_state["modo_carga_vertice"])
    st.download_button(
        label="📥 Descargar Excel Consolidado (4 Cortes — Sin Duplicados)",
        data=excel_data,
        file_name=f"Operacion_Vertice_4_Cortes_{datetime.now().strftime('%d%b%y_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

vals_act = BASE_VALORES[grupo_sel].copy()
df_ex_act = st.session_state.get("df_vertice_extra")
modo_act = st.session_state.get("modo_carga_vertice", "Sumar como reporte nuevo")

if df_ex_act is not None:
    sub_ex = df_ex_act if grupo_sel == "Consolidado" else df_ex_act[df_ex_act["Corte"] == grupo_sel]
    ex_oficios = int(sub_ex["Oficios / Capturas"].sum()) if not sub_ex.empty else 0
    ex_inf = int(sub_ex["Infracciones"].sum()) if not sub_ex.empty else 0
    ex_gruas = int(sub_ex["Grúas"].sum()) if not sub_ex.empty else 0
    ex_faltas = int(sub_ex["Faltas Admin / Indoc"].sum()) if not sub_ex.empty else 0
    ex_pers = int(sub_ex["Personas Verificadas"].sum()) if not sub_ex.empty else 0
    ex_veh = int(sub_ex["Vehículos Verificados"].sum()) if not sub_ex.empty else 0
    ex_otras = int(sub_ex["Otras Acciones"].sum()) if not sub_ex.empty else 0

    if modo_act == "Reemplazar valores base":
        vals_act["oficios"] = ex_oficios
        vals_act["infracciones"] = ex_inf
        vals_act["gruas"] = ex_gruas
        vals_act["faltas_admin"] = ex_faltas
        vals_act["personas"] = ex_pers
        vals_act["vehiculos"] = ex_veh
        vals_act["otras"] = ex_otras
    else:
        vals_act["oficios"] += ex_oficios
        vals_act["infracciones"] += ex_inf
        vals_act["gruas"] += ex_gruas
        vals_act["faltas_admin"] += ex_faltas
        vals_act["personas"] += ex_pers
        vals_act["vehiculos"] += ex_veh
        vals_act["otras"] += ex_otras

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
              <div class="kpi-big kpi-amber" style="font-size:32px;">{vals_act['oficios']}</div>
              <div class="kpi-lbl" style="font-size:10px;">OFICIOS / CAPTURAS</div>
              <div style="color:#94A3B8; font-size:10px; text-align:center;">{grupo_sel}</div>
            </div>
            <div>
              <div class="kpi-big kpi-amber" style="font-size:32px;">{vals_act['personas']:,}</div>
              <div class="kpi-lbl" style="font-size:10px;">PERS. VERIFICADAS</div>
              <div style="color:#94A3B8; font-size:10px; text-align:center;">Vehículos: {vals_act['vehiculos']:,}</div>
            </div>
            <div>
              <div class="kpi-big kpi-amber" style="font-size:32px;">{vals_act['faltas_admin']}</div>
              <div class="kpi-lbl" style="font-size:10px;">FALTAS ADMIN / INDOC</div>
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
              <div class="kpi-big kpi-orange" style="font-size:32px;">{vals_act['infracciones']}</div>
              <div class="kpi-lbl" style="font-size:10px;">INFRACCIONES</div>
              <div style="color:#94A3B8; font-size:10px; text-align:center;">{grupo_sel}</div>
            </div>
            <div>
              <div class="kpi-big kpi-orange" style="font-size:32px;">{vals_act['gruas']}</div>
              <div class="kpi-lbl" style="font-size:10px;">REMOLCADOS (GRÚA)</div>
              <div style="color:#94A3B8; font-size:10px; text-align:center;">{grupo_sel}</div>
            </div>
          </div>
        </div>
        """, unsafe_allow_html=True)

    b1, b2, b3 = st.columns([1.4, 1.4, 1.2])

    with b1:
        st.markdown('<div class="panel-box amber"><div class="panel-hdr" style="text-align:left;padding-left:12px;">DETALLE DE CONDUCCIÓN, CAPTURA Y VERIFICACIÓN</div></div>', unsafe_allow_html=True)
        df_acc_det = pd.DataFrame([
            {"Tipo": "Conducción / Captura / Cintillo / Oficio", "Total": vals_act["oficios"]},
            {"Tipo": "Personas verificadas (PDA/SVC)", "Total": vals_act["personas"]},
            {"Tipo": "Vehículos verificados", "Total": vals_act["vehiculos"]},
            {"Tipo": "Otras acciones operativas", "Total": vals_act["otras"]},
        ])
        st.dataframe(df_acc_det, use_container_width=True, hide_index=True)

    with b2:
        st.markdown('<div class="panel-box amber"><div class="panel-hdr" style="text-align:left;padding-left:12px;">DETALLE DE FALTAS, GRÚAS E INFRACCIONES</div></div>', unsafe_allow_html=True)
        df_inc_det = pd.DataFrame([
            {"Tipo": "Faltas administrativas / Indocumentados / Migración", "Total": vals_act["faltas_admin"]},
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
