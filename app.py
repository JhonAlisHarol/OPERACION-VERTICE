import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, date

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(
    page_title="CON-C5 | Sistema de Recursos Humanos",
    page_icon="🛡️",
    layout="wide"
)

# --- ESTILO INSTITUCIONAL CON-C5 ---
st.markdown("""
    <style>
    .stApp { background-color: #061325; color: #E2E8F0; }
    h1, h2, h3 { color: #38BDF8 !important; }
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] {
        background-color: #0F2342;
        border-radius: 6px 6px 0 0;
        color: #CBD5E1;
        padding: 10px 20px;
    }
    .stTabs [aria-selected="true"] {
        background-color: #0284C7 !important;
        color: #FFFFFF !important;
        font-weight: bold;
    }
    </style>
""", unsafe_allow_html=True)

# --- COLUMNAS OFICIALES (60 CAMPOS DE Plan_de_Recojo) ---
COLUMNAS_OFICIALES = [
    "TIMESTAMP", "USUARIO_REGISTRO", "RANGO", "POSIC", "PLACA_METALICA",
    "NOMBRE", "APELLIDO", "CEDULA", "FECHA_NACIMIENTO", "EDAD", "GENERO",
    "TELEFONO_PERSONAL", "ESTAMENTO", "NUM_PROMOCION", "DESCRIPCION_PROMOCION",
    "FECHA_INGRESO", "ANOS_SERVICIOS", "RES_OGD_ASIG", "ASIG_TRASLADO",
    "DIREC_DEPEN", "ASIGNADO_CENTRO", "U_CENTRO_ASIGNADO", "DEPARTAMENTO",
    "TURNO", "GRUPO", "FUNCION", "ESTADO_STATUS", "OPERAT_NO_OPERAT",
    "PERM_TEMP", "CONDICION", "APTO_NO_APTO", "NIVEL_SALUD", "OBSERVACION",
    "CERT_BUENA_SALUD", "FECHA_DIAGNOSTICO", "F_PROX_EVALUACION",
    "SALUD_OCUPACIONAL", "PENSIONADO", "PROVINCIA", "DISTRITO",
    "CORREGIMIENTO", "LUGAR", "TIPO_SANGRE", "NOMBRE_EMERGENCIA",
    "PARENTESCO", "TELEFONO_EMERGENCIA", "ESTADO_CIVIL", "DONANTE_SANGRE",
    "SECUNDARIA", "UNIVERSIDAD", "HIJOS_M", "HIJAS_F", "TOTAL_HIJOS",
    "TALLA_CAMISA", "TALLA_PANTALON", "TALLA_SUETER", "TALLA_BOTAS",
    "TALLA_ZAPATOS", "TALLA_QUEPI", "TALLA_CORREA"
]

# --- CONEXIÓN A GOOGLE SHEETS ---
@st.cache_resource
def conectar_google_sheets():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    client = gspread.authorize(creds)
    sheet_name = st.secrets.get("SHEET_NAME_RRHH", "Recursos_Humanos")
    ws_name = st.secrets.get("WORKSHEET_NAME_RRHH", "Plan_de_Recojo")
    sh = client.open(sheet_name)
    return sh.worksheet(ws_name)

def calcular_anios(fecha_ref):
    if not fecha_ref:
        return 0
    hoy = date.today()
    return hoy.year - fecha_ref.year - ((hoy.month, hoy.day) < (fecha_ref.month, fecha_ref.day))

def cargar_datos(ws):
    registros = ws.get_all_records()
    if not registros:
        return pd.DataFrame(columns=COLUMNAS_OFICIALES)
    return pd.DataFrame(registros)

# --- ENCABEZADO ---
st.title("🛡️ CENTRO DE OPERACIÓN NACIONAL (CON - C5)")
st.caption("Sistema de Gestión y Registro de Talento Humano · Conectado a Google Sheets (`Recursos_Humanos`)")

try:
    ws = conectar_google_sheets()
except Exception as e:
    st.error(f"Error de conexión con Google Sheets: {e}")
    st.stop()

# --- PESTAÑAS PRINCIPALES ---
tab_registro, tab_masivo, tab_base = st.tabs([
    "📝 Registro Individual (Plan de Recojo)",
    "📥 Carga Masiva (Excel / CSV)",
    "📊 Base Consolidada"
])

# ==============================================================================
# 1. REGISTRO INDIVIDUAL
# ==============================================================================
with tab_registro:
    with st.form("form_rrhh", clear_on_submit=True):
        st.subheader("1. Identificación y Datos Institucionales")
        c1, c2, c3, c4 = st.columns(4)
        usuario_reg = c1.text_input("Usuario que Registra *", placeholder="Ej: Sgto. 1ro. J. Diaz")
        estamento = c2.selectbox("Estamento *", ["Policía Nacional", "SENAN", "SENAFRONT", "Migración", "MINSEG", "MNJ", "Enlace Externo"])
        rango = c3.selectbox("Rango *", [
            "Comisionado", "Subcomisionado", "Mayor", "Capitán", "Teniente",
            "Subteniente", "Sargento 1ro.", "Sargento 2do.", "Cabo 1ro.",
            "Cabo 2do.", "Agente", "MNJ", "Licenciado(a)", "Administrativo"
        ])
        posic = c4.text_input("Posición *", placeholder="Ej: 24561")

        c5, c6, c7, c8 = st.columns(4)
        placa = c5.text_input("Placa Metálica")
        nombre = c6.text_input("Nombre *")
        apellido = c7.text_input("Apellido *")
        cedula = c8.text_input("Cédula *", placeholder="Ej: 8-000-0000")

        c9, c10, c11, c12 = st.columns(4)
        fecha_nac = c9.date_input("Fecha de Nacimiento", value=date(1990, 1, 1), min_value=date(1950, 1, 1), max_value=date.today())
        genero = c10.selectbox("Género", ["Masculino", "Femenino"])
        tel_personal = c11.text_input("Teléfono Personal", placeholder="6000-0000")
        estado_civil = c12.selectbox("Estado Civil", ["Soltero(a)", "Casado(a)", "Unido(a)", "Divorciado(a)", "Viudo(a)"])

        st.subheader("2. Trayectoria y Asignación Operativa CON-C5")
        c13, c14, c15, c16 = st.columns(4)
        num_promo = c13.text_input("N.º Promoción")
        desc_promo = c14.text_input("Descripción Promoción")
        fecha_ing = c15.date_input("Fecha de Ingreso", value=date(2015, 1, 1), min_value=date(1980, 1, 1), max_value=date.today())
        res_ogd = c16.text_input("Resolución / OGD Asignación")

        c17, c18, c19, c20 = st.columns(4)
        asig_traslado = c17.selectbox("Asignado / Traslado", ["Asignado", "Traslado", "Comisión", "Prestado"])
        direc_depen = c18.text_input("Dirección / Dependencia", value="MINSEG - CON")
        centro = c19.selectbox("Centro Asignado *", ["CON", "CORCOL", "COMCH", "COMDA", "COMAR", "COMSAM", "COMCHEP", "CEVIBO"])
        u_centro = c20.text_input("Último Centro Asignado")

        c21, c22, c23, c24 = st.columns(4)
        departamento = c21.selectbox("Departamento *", [
            "Call Center 104", "Despacho", "Videovigilancia", "Control de Calidad",
            "Extracción de Video", "Planeamiento", "Recursos Humanos",
            "Análisis / Estadística", "Seguridad de Instalaciones", "Soporte Técnico", "Enlace"
        ])
        turno = c22.selectbox("Turno", ["4x4 (05:00-17:00)", "4x4 (17:00-05:00)", "7x7", "Diurno Administrativo", "20x20"])
        grupo = c23.selectbox("Grupo", ["Grupo A", "Grupo B", "Grupo C", "Grupo D", "A-1", "A-2", "B-1", "B-2", "Administrativo", "N/A"])
        funcion = c24.text_input("Función Específica", placeholder="Ej: Operador de Cámaras / Despachador")

        c25, c26, c27, c28 = st.columns(4)
        estado_status = c25.selectbox("Estado / Status *", ["Activo", "Vacaciones", "Vacaciones por Jubilación", "Incapacidad", "Comisión", "Licencia", "Permiso"])
        operatividad = c26.selectbox("Operatividad *", ["Operativo", "No Operativo", "Operativo Administrativo"])
        perm_temp = c27.selectbox("Permanente / Temporal", ["Permanente", "Temporal"])
        condicion = c28.text_input("Condición", value="Normal")

        st.subheader("3. Salud, Residencia y Contacto de Emergencia")
        c29, c30, c31, c32 = st.columns(4)
        apto = c29.selectbox("Apto / No Apto", ["Apto", "No Apto", "Apto con Restricción"])
        nivel_salud = c30.text_input("Nivel de Salud")
        cert_salud = c31.selectbox("Certificado Buena Salud", ["Sí", "No", "En Trámite"])
        salud_ocup = c32.text_input("Salud Ocupacional")

        c33, c34, c35, c36 = st.columns(4)
        fecha_diag = c33.text_input("Fecha Diagnóstico (si aplica)", placeholder="DD/MM/AAAA")
        f_prox_eval = c34.text_input("Fecha Próx. Evaluación", placeholder="DD/MM/AAAA")
        pensionado = c35.selectbox("Pensionado", ["No", "Sí"])
        tipo_sangre = c36.selectbox("Tipo de Sangre", ["O+", "O-", "A+", "A-", "B+", "B-", "AB+", "AB-", "No Sabe"])

        c37, c38, c39, c40 = st.columns(4)
        provincia = c37.selectbox("Provincia", [
            "Panamá", "Panamá Oeste", "Colón", "Chiriquí", "Coclé",
            "Veraguas", "Herrera", "Los Santos", "Bocas del Toro",
            "Darién", "Comarca Ngäbe-Buglé", "Comarca Guna Yala", "Comarca Emberá-Wounaan"
        ])
        distrito = c38.text_input("Distrito")
        corregimiento = c39.text_input("Corregimiento")
        lugar = c40.text_input("Barriada / Calle / Casa")

        c41, c42, c43, c44 = st.columns(4)
        nombre_emerg = c41.text_input("Contacto de Emergencia")
        parentesco = c42.text_input("Parentesco")
        tel_emerg = c43.text_input("Teléfono Emergencia")
        donante = c44.selectbox("Donante de Sangre", ["No", "Sí"])

        st.subheader("4. Formación, Carga Familiar y Tallas de Uniforme")
        c45, c46, c47, c48 = st.columns(4)
        secundaria = c45.text_input("Nivel Secundario / Bachiller")
        universidad = c46.text_input("Estudios Universitarios / Postgrado")
        hijos_m = c47.number_input("Hijos Varones (M)", min_value=0, max_value=20, step=1)
        hijas_f = c48.number_input("Hijas Mujeres (F)", min_value=0, max_value=20, step=1)

        t1, t2, t3, t4, t5, t6, t7 = st.columns(7)
        t_camisa = t1.text_input("Camisa", placeholder="M / L")
        t_pantalon = t2.text_input("Pantalón", placeholder="32 / 34")
        t_sueter = t3.text_input("Suéter", placeholder="M / L")
        t_botas = t4.text_input("Botas", placeholder="9.5")
        t_zapatos = t5.text_input("Zapatos", placeholder="9.5")
        t_quepi = t6.text_input("Quepi", placeholder="7 1/4")
        t_correa = t7.text_input("Correa", placeholder="34")

        observacion = st.text_area("Observaciones Generales")
        enviado = st.form_submit_button("💾 Guardar Registro en Recursos Humanos", use_container_width=True)

        if enviado:
            if not nombre.strip() or not apellido.strip() or not cedula.strip():
                st.warning("⚠️ Complete los campos obligatorios: Nombre, Apellido y Cédula.")
            else:
                edad_calc = calcular_anios(fecha_nac)
                anios_serv_calc = calcular_anios(fecha_ing)
                total_hijos = int(hijos_m + hijas_f)
                timestamp_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                fila = [
                    timestamp_actual, usuario_reg.strip(), rango, posic.strip(), placa.strip(),
                    nombre.strip().upper(), apellido.strip().upper(), cedula.strip(),
                    fecha_nac.strftime("%Y-%m-%d"), edad_calc, genero, tel_personal.strip(),
                    estamento, num_promo.strip(), desc_promo.strip(), fecha_ing.strftime("%Y-%m-%d"),
                    anios_serv_calc, res_ogd.strip(), asig_traslado, direc_depen.strip(),
                    centro, u_centro.strip(), departamento, turno, grupo, funcion.strip(),
                    estado_status, operatividad, perm_temp, condicion.strip(), apto,
                    nivel_salud.strip(), observacion.strip(), cert_salud, fecha_diag.strip(),
                    f_prox_eval.strip(), salud_ocup.strip(), pensionado, provincia,
                    distrito.strip(), corregimiento.strip(), lugar.strip(), tipo_sangre,
                    nombre_emerg.strip(), parentesco.strip(), tel_emerg.strip(), estado_civil,
                    donante, secundaria.strip(), universidad.strip(), int(hijos_m), int(hijas_f),
                    total_hijos, t_camisa.strip(), t_pantalon.strip(), t_sueter.strip(),
                    t_botas.strip(), t_zapatos.strip(), t_quepi.strip(), t_correa.strip()
                ]
                ws.append_row(fila, value_input_option="USER_ENTERED")
                st.success(f"✅ Unidad **{rango} {nombre.upper()} {apellido.upper()}** registrada correctamente.")

# ==============================================================================
# 2. CARGA MASIVA DESDE EXCEL / CSV
# ==============================================================================
with tab_masivo:
    st.subheader("Importación Masiva desde Excel (.xlsx) o CSV")
    st.info("El archivo debe contener los encabezados exactos de `Plan_de_Recojo` (las columnas faltantes se rellenarán en blanco automáticamente).")

    plantilla_df = pd.DataFrame(columns=COLUMNAS_OFICIALES)
    st.download_button(
        label="📄 Descargar Plantilla CSV Oficial (60 Columnas)",
        data=plantilla_df.to_csv(index=False).encode("utf-8-sig"),
        file_name="plantilla_rrhh_con.csv",
        mime="text/csv"
    )

    archivo = st.file_uploader("Subir archivo de personal", type=["xlsx", "csv"])
    if archivo:
        df_subido = pd.read_csv(archivo) if archivo.name.endswith(".csv") else pd.read_excel(archivo)
        df_subido.columns = [c.strip().upper() for c in df_subido.columns]

        # Alinear estrictamente a las 60 columnas oficiales
        df_alineado = pd.DataFrame(columns=COLUMNAS_OFICIALES)
        for col in COLUMNAS_OFICIALES:
            if col in df_subido.columns:
                df_alineado[col] = df_subido[col].fillna("").astype(str)
            else:
                df_alineado[col] = ""

        df_alineado["TIMESTAMP"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        st.dataframe(df_alineado.head(15), use_container_width=True)

        if st.button(f"🚀 Subir {len(df_alineado)} registros a Google Sheets"):
            ws.append_rows(df_alineado.values.tolist(), value_input_option="USER_ENTERED")
            st.success(f"✅ Se importaron {len(df_alineado)} registros exitosamente.")

# ==============================================================================
# 3. BASE CONSOLIDADA
# ==============================================================================
with tab_base:
    st.subheader("Consulta General del Efectivo Registrado")
    df_actual = cargar_datos(ws)
    if df_actual.empty:
        st.info("Aún no hay registros cargados en `Plan_de_Recojo`.")
    else:
        col_f1, col_f2, col_f3 = st.columns(3)
        filtro_centro = col_f1.multiselect("Filtrar por Centro", sorted(df_actual["ASIGNADO_CENTRO"].astype(str).unique()))
        filtro_estamento = col_f2.multiselect("Filtrar por Estamento", sorted(df_actual["ESTAMENTO"].astype(str).unique()))
        busqueda = col_f3.text_input("Buscar por Cédula, Posición o Apellido")

        df_view = df_actual.copy()
        if filtro_centro:
            df_view = df_view[df_view["ASIGNADO_CENTRO"].isin(filtro_centro)]
        if filtro_estamento:
            df_view = df_view[df_view["ESTAMENTO"].isin(filtro_estamento)]
        if busqueda:
            mask = (
                df_view["CEDULA"].astype(str).str.contains(busqueda, case=False, na=False) |
                df_view["POSIC"].astype(str).str.contains(busqueda, case=False, na=False) |
                df_view["APELLIDO"].astype(str).str.contains(busqueda, case=False, na=False)
            )
            df_view = df_view[mask]

        m1, m2, m3 = st.columns(3)
        m1.metric("Total Efectivo Filtrado", len(df_view))
        m2.metric("Operativos", int((df_view["OPERAT_NO_OPERAT"] == "Operativo").sum()))
        m3.metric("No Operativos / Admin.", int((df_view["OPERAT_NO_OPERAT"] != "Operativo").sum()))

        st.dataframe(df_view, use_container_width=True)
