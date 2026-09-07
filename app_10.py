import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import date, datetime, timezone
from zoneinfo import ZoneInfo
import time
import re
import uuid

from streamlit_autorefresh import st_autorefresh

from supabase import create_client, Client

# =========================================================
# CONFIGURACION
# =========================================================
st.set_page_config(
    page_title="JC Control de Solicitudes — Despacho",
    page_icon="📝",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# =========================================================
# ESTILOS
# =========================================================
st.markdown(
    """
<style>
* { box-sizing:border-box; }
.stApp { background:#020914; color:#f3f4f6; font-family:Arial,Helvetica,sans-serif; }
.block-container { width:94%; max-width:1450px; padding-top:22px!important; padding-bottom:24px!important; margin:auto; }
header[data-testid="stHeader"] { background:transparent; }
div[data-testid="stToolbar"] { display:none; }
footer { display:none; }
h1,h2,h3 { color:#f3f4f6!important; font-family:Arial,Helvetica,sans-serif!important; }
.titulo { text-align:center; margin:4px 0 14px; font-size:30px; font-weight:900; line-height:1.2; width:100%; white-space:nowrap; }
.stButton>button,.stDownloadButton>button { border:1px solid #273246!important; border-radius:5px!important; min-height:30px!important; height:30px!important; padding:4px 8px!important; font-weight:800!important; color:white!important; background:#17263d!important; font-size:10px!important; }
.stButton>button:hover,.stDownloadButton>button:hover { filter:brightness(1.18); border-color:#315b86!important; }

/* BOTONES CANCELAR — rojo y respuesta visual al presionarlos */
.stFormSubmitButton > button[kind="primary"] {
    background:#7f1d1d!important;
    border-color:#ef4444!important;
    color:#ffffff!important;
}
.stFormSubmitButton > button[kind="primary"]:hover {
    background:#b91c1c!important;
    border-color:#f87171!important;
    filter:none!important;
}
.stFormSubmitButton > button[kind="primary"]:active {
    background:#ef4444!important;
    border-color:#fecaca!important;
    transform:scale(.98)!important;
}
div[data-testid="stForm"] { background:rgba(7,15,29,.55); border:1px solid #273246; border-radius:6px; padding:10px!important; }
div[data-testid="stForm"] label { color:#cbd5e1!important; font-size:10px!important; font-weight:700!important; }
div[data-baseweb="input"],div[data-baseweb="select"]>div,div[data-testid="stDateInput"]>div { background:#242630!important; color:#f3f4f6!important; border-radius:5px!important; border-color:transparent!important; min-height:33px!important; }
div[data-baseweb="input"] input,div[data-testid="stDateInput"] input { color:#f3f4f6!important; font-size:11px!important; }
div[data-baseweb="select"] span { color:#f3f4f6!important; font-size:11px!important; }
div[data-baseweb="select"] svg { fill:#f3f4f6!important; }
div[data-testid="stTextInput"] input { width:100%; height:33px; border:1px solid transparent; border-radius:5px; background:#242630!important; color:#f3f4f6!important; padding:0 9px; font-size:11px; }
/* =========================================================
   TABLA JC — ESTILO PROFESIONAL
   ========================================================= */
div[data-testid="stDataFrame"],
div[data-testid="stDataEditor"] {
    width:100%!important;
    border:1px solid #30445f!important;
    border-radius:12px!important;
    overflow:hidden!important;
    background:linear-gradient(180deg,#0b1526 0%,#07101d 100%)!important;
    box-shadow:0 8px 28px rgba(0,0,0,.28), inset 0 1px 0 rgba(255,255,255,.025)!important;
}

/* Barra superior del editor */
div[data-testid="stDataEditor"] > div {
    border-radius:12px!important;
}

/* Encabezados */
div[data-testid="stDataEditor"] [role="columnheader"] {
    background:#172a43!important;
    color:#f8fafc!important;
    font-weight:800!important;
    border-bottom:1px solid #3a5270!important;
}

/* Indicadores de color dentro de ESTADO y PRIORIDAD */
div[data-testid="stDataEditor"] [role="gridcell"] {
    font-size: 12px !important;
}

/* Encabezados de acciones */
div[data-testid="stDataEditor"] [role="columnheader"] {
    letter-spacing: .2px !important;
}

/* Botones de edición de la tabla */
div[data-testid="stDataEditor"] button {
    border-radius:6px!important;
}

/* Checkboxes EDITAR / ELIMINAR */
div[data-testid="stDataEditor"] input[type="checkbox"] {
    accent-color:#38bdf8!important;
}

/* Contenedor de la tabla */
div[data-testid="stDataEditor"] canvas {
    border-radius:10px!important;
}

/* Texto que acompaña el contador */
div[data-testid="stCaptionContainer"] {
    color:#94a3b8!important;
    font-size:11px!important;
    font-weight:700!important;
    padding:5px 2px 8px!important;
}

/* Separador */
hr { border-color:#273246!important; }
.footer { margin-top:30px; padding:15px 10px; border-top:1px solid #273246; text-align:center; color:#8f9bad; font-size:11px; }
.footer strong { color:#dbe2ea; }
.login-wrapper { max-width:380px; margin:8px auto 14px; text-align:center; }
.login-icon { font-size:30px; line-height:1; margin-bottom:5px; }
.login-title { font-size:27px; font-weight:800; line-height:1.15; }
.login-subtitle { font-size:12px; opacity:.62; margin-top:5px; }
</style>
""",
    unsafe_allow_html=True,
)

# =========================================================
# CONSTANTES
# =========================================================
COLS = [
    "CLIENTE", "NRO SOLICITUD - WO", "TIPO DE SOLICITUD", "CENTRO DE COSTO",
    "PRIORIDAD", "CANT - ITEMS", "ESTADO DE SOLICITUD", "DIRECCIÓN", "FECHA DE INGRESO"
]
RUTA_SHEET = "PROGRAMACION_RUTAS"
LOCK_TIMEOUT_SECONDS = 15 * 60
INACTIVITY_TIMEOUT_SECONDS = 8 * 60
AUTO_REFRESH_INTERVAL_MS = 60 * 1000
ZONA_HORARIA_APP = ZoneInfo("America/La_Paz")

CLIENTES = [
    "Seleccione una opcion", "BANCO SOL", "BANCO BNB", "BANCO FIE", "BANCO FORTALEZA",
    "PRENDAMAS", "MOLINO ANDINO", "PREVICOR CORREDORES", "WCS-BOLIVIA",
    "INDUSTRIA Y COMERCIO ALICONSUMO",
]
TIPOS = [
    "Seleccione una opcion", "EXTERNO", "INTERNO - INV.", "INTERNO - BPO.",
    "INDEXACION - BPO.", "REVISION INTERNA", "SERVICIOS", "ENVIO DE MATERIALES",
]
PRIORIDADES = ["Seleccione una opcion", "RUSH", "TURNO SIGUIENTE", "NORMAL"]

CENTROS_COSTO = [
    "Seleccione una opcion",
    "100 - REGIONAL CHUQUISACA",	"138 - BANCO NACIONAL DE BOLIVIA - PANDO",	"220 - TEMBLADERANI",	"271 - INGAVI",	"316 - AYACUCHO",	"410 - TAGARETE",	"608 - YACUIBA",	"912 - LABORATORIO LKM BOLIVIA S.A.",	"REGIONAL EL ALTO - 12 DE OCTUBRE",	"REGIONAL EL ALTO - VILLA ADELA",	"REGIONAL LA PAZ - TEMBLADERANI",	"REGIONAL SANTA CRUZ - GUARAYOS",	"REGIONAL SUCRE - CHARCAS",
    "101 - MERCADO CAMPESINO SUCRE",	"140 - INDEXACION - BNB LA PAZ",	"221 - LA PORTADA",	"272 - SAN ROQUE",	"317 - SOL AMIGO COCHABAMBA",	"411 - TACNA",	"609 - PALMARCITO",	"913 - PHARMATECH BOLIVIANA S.A.",	"REGIONAL EL ALTO - 16 DE JULIO",	"REGIONAL EL ALTO - VILLA DOLORES",	"REGIONAL LA PAZ - VILLA COPACABANA",	"REGIONAL SANTA CRUZ - KILOMETRO 6",	"REGIONAL SUCRE - MESA VERDE",
    "102 - NORMALIZACIÓN SUCRE",	"141 - INDEXACION - BNB COCHABAMBA",	"222 - PERIFERICA",	"273 - EL KENKO",	"318 - PANAMERICANA",	"500 - REGIONAL POTOSI",	"610 - MERCADO CAMPESINO YACUIBA",	"C008 - POLYSISTEMAS",	"REGIONAL EL ALTO - 21 DE OCTUBRE",	"REGIONAL EL ALTO - VILLA ESPERANZA",	"REGIONAL LA PAZ - VILLA FATIMA",	"REGIONAL SANTA CRUZ - LA RAMADA",	"REGIONAL SUCRE - MONTEAGUDO",
    "1020153022 - BNB VALORES S.A.",	"142 - INDEXACION - BNB ORURO",	"223 - PAMPAHASI",	"274 - QUISWARAS",	"319 - AGENCIA MOVÍL",	"501 - MERCADO UYUNI",	"612 - INDUSTRIA Y COMERCIO ALICONSUMO S.A.",	"REGIONAL BENI - RIBERALTA",	"REGIONAL EL ALTO - ACHACACHI",	"REGIONAL EL ALTO - VILLA YUNGUYO",	"REGIONAL LA PAZ - ZONA SUR",	"REGIONAL SANTA CRUZ - LOS LOTES",	"REGIONAL SUCRE - PADILLA",
    "103 - ESPAÑA",	"143 - INDEXACION - BNB TARIJA",	"224 - CHUQUIAGUILLO",	"275 - FRANZ TAMAYO",	"320 - CLIZA",	"502 - SOL AMIGO POTOSI",	"613 - INDUSTRIA Y COMERCIO ALICONSUMO - SANTA CRUZ",	"REGIONAL COCHABAMBA - ALALAY",	"REGIONAL EL ALTO - ACHOCALLA",	"REGIONAL LA PAZ - BUENOS AIRES",	"REGIONAL ORURO - BOLIVAR",	"REGIONAL SANTA CRUZ - LOS POCITOS",	"REGIONAL TARIJA - ANDALUCIA",
    "104 - SOL AMIGO SUCRE",	"146 - INDEXACIÓN – BNB BENI",	"225 - ACHUMANI",	"276 - VILLA INGENIO",	"321 - VINTO",	"503 - BOULEVARD",	"618 - ALPASUR S.A.",	"REGIONAL COCHABAMBA - CALA CALA",	"REGIONAL EL ALTO - BALLIVIAN",	"REGIONAL LA PAZ - CAMACHO",	"REGIONAL ORURO - ESPAÑA",	"REGIONAL SANTA CRUZ - LOS TUSEQUIS",	"REGIONAL TARIJA - ARANJUEZ",
    "105 - GERMAN MENDOZA",	"200 - REGIONAL LA PAZ",	"250 - REGIONAL EL ALTO",	"277 - FERROPETROL",	"322 - REPÚBLICA",	"505 - LAS BANDERAS",	"620 - MOLINO ANDINO S.A.",	"REGIONAL COCHABAMBA - CHIMORE",	"REGIONAL EL ALTO - COPACABANA",	"REGIONAL LA PAZ - CARANAVI",	"REGIONAL ORURO - LLALLAGUA",	"REGIONAL SANTA CRUZ - MAIRANA",	"REGIONAL TARIJA - BERMEJO",
    "106 - CHARCAS",	"201 - SAN PEDRO",	"251 - 16 DE JULIO",	"279 - CHACALTAYA",	"323 - TIQUIPAYA",	"506 - MURILLO",	"700 - REGIONAL SANTA CRUZ",	"REGIONAL COCHABAMBA - COLCAPIRHUA",	"REGIONAL EL ALTO - DESAGUADERO",	"REGIONAL LA PAZ - CHASQUIPAMPA",	"REGIONAL ORURO - PAGADOR",	"REGIONAL SANTA CRUZ - MERCADO ABASTO",	"REGIONAL TARIJA - CAMARGO",
    "107 - ZUDAÑEZ",	"202 - GARITA",	"252 - LA CEJA",	"280 - LAGUNAS EX PARADA 8",	"324 - QUINTANILLA",	"555 - PREVICOR CORREDORES Y ASES. DE SEG.",	"701 - CASCO VIEJO",	"REGIONAL COCHABAMBA - CRUCE TAQUIÑA",	"REGIONAL EL ALTO - LA CEJA",	"REGIONAL LA PAZ - CHULUMANI",	"REGIONAL ORURO - SUCURSAL ORURO",	"REGIONAL SANTA CRUZ - MERCADO FERRETERO",	"REGIONAL TARIJA - CULPINA",
    "108 - LAS AMERICAS",	"203 - VILLA FATIMA",	"253 - RIO SECO",	"295 - OFICINA NACIONAL",	"325 - JORDAN",	"557 - PRENDAMAS S.R.L.",	"702 - EL PARI",	"REGIONAL COCHABAMBA - ENTRE RIOS",	"REGIONAL EL ALTO - LIBERTAD",	"REGIONAL LA PAZ - COBIJA",	"REGIONAL POTOSI - 10 DE NOVIEMBRE",	"REGIONAL SANTA CRUZ - MONTERO",	"REGIONAL TARIJA - GUADALQUIVIR",
    "109 - LAJASTAMBO",	"204 - MIRAFLORES",	"254 - VILLA ADELA",	"300 - REGIONAL COCHABAMBA",	"326 - PLAZA BOLIVAR",	"559 - BANCO FORTALEZA S.A.",	"703 - MUTUALISTA",	"REGIONAL COCHABAMBA - HEROINAS",	"REGIONAL EL ALTO - NUEVO AMANECER",	"REGIONAL LA PAZ - CORIPATA",	"REGIONAL POTOSI - BETANZOS",	"REGIONAL SANTA CRUZ - NORTE",	"REGIONAL TARIJA - GUADALUPE",
    "1093 - BNB VALORES S.A.",	"205 - EL TEJAR",	"256 - VIACHA",	"301 - ESTEBAN ARCE",	"327 - PETROLERA",	"561 - INTERQUIMICA INDUSTRIAL S.A.",	"704 - 1RO. DE MAYO",	"REGIONAL COCHABAMBA - IVIRGARZAMA",	"REGIONAL EL ALTO - NUEVOS HORIZONTES",	"REGIONAL LA PAZ - COROICO",	"REGIONAL POTOSI - CERRO DE PLATA",	"REGIONAL SANTA CRUZ - PAMPA DE LA ISLA",	"REGIONAL TARIJA - LA TABLADA",
    "124-1 - BFIE LA PAZ",	"206 - ALONSO DE MENDOZA",	"258 - NORMALIZACIÓN EL ALTO",	"302 - SAN MARTIN",	"328 - LA CHIMBA",	"563 - DHL BOLIVIA S.R.L.",	"705 - MONTERO",	"REGIONAL COCHABAMBA - JORDAN",	"REGIONAL EL ALTO - OFICINA CENTRAL",	"REGIONAL LA PAZ - EL TEJAR",	"REGIONAL POTOSI - COTAGAITA",	"REGIONAL SANTA CRUZ - PLAN 3000",	"REGIONAL TARIJA - LUIS DE FUENTE",
    "124-2 - BFIE SANTA CRUZ",	"208 - SAN MIGUEL",	"259 - 12 DE OCTUBRE",	"303 - HUAYRA KHASA",	"329 - AMERICA",	"565 - BANCO FORTALEZA S.A. - TARIJA",	"706 - EL TORNO",	"REGIONAL COCHABAMBA - KANATA",	"REGIONAL EL ALTO - PACAJES",	"REGIONAL LA PAZ - GRAN PODER",	"REGIONAL POTOSI - JUNIN",	"REGIONAL SANTA CRUZ - SAN IGNACIO",	"REGIONAL TARIJA - VALLE DE CONCEPCION",
    "124-3 - BFIE COCHABAMBA",	"209 - BALLIVIAN",	"260 - SENKATA",	"305 - CRUCE TAQUIÑA",	"331 - EL AVION",	"565 - WCS-BOLIVIA",	"709 - PIRAI",	"REGIONAL COCHABAMBA - LA CANCHA",	"REGIONAL EL ALTO - PANAMERICANA",	"REGIONAL LA PAZ - LA PORTADA",	"REGIONAL POTOSI - NUEVA TERMINAL",	"REGIONAL SANTA CRUZ - SAN JULIAN",	"REGIONAL TARIJA - VILLAMONTES",
    "126 - ORURO - ESPAÑA",	"210 - CAMACHO",	"261 - BOLIVIA",	"306 - QUILLACOLLO",	"332 - VILLA PAGADOR",	"599 - ROCHE BOLIVIA S.R.L.",	"711 - PLAN 3000",	"REGIONAL COCHABAMBA - NATANIEL AGUIRRE",	"REGIONAL EL ALTO - PATACAMAYA",	"REGIONAL LA PAZ - MIRAFLORES",	"REGIONAL POTOSI - SAN ROQUE",	"REGIONAL SANTA CRUZ - SANTOS DUMONT",	"REGIONAL TARIJA - YACUIBA",
    "130 - BANCO NACIONAL DE BOLIVIA - LA PAZ",	"211 - CRUCE VILLA COPACABANA",	"262 - SATELITE",	"307 - COLCA PIRHUA",	"333 - PACATA",	"600 - REGIONAL TARIJA",	"713 - NORMALIZACIÓN SANTA CRUZ",	"REGIONAL COCHABAMBA - PACATA",	"REGIONAL EL ALTO - RIO SECO",	"REGIONAL LA PAZ - PALOS BLANCOS",	"REGIONAL POTOSI - TUPIZA",	"REGIONAL SANTA CRUZ - SATELITE NORTE",	
    "131 - BANCO NACIONAL DE BOLIVIA - SANTA CRUZ",	"212 - COTA COTA",	"263 - VILLA DOLORES",	"309 - MUYURINA",	"334 - COLOMI",	"601 - MERCADO CAMPESINO TARIJA",	"715 - LA GUARDIA",	"REGIONAL COCHABAMBA - PETROLERA",	"REGIONAL EL ALTO - ROMERO PAMPA",	"REGIONAL LA PAZ - PAMPAHASI",	"REGIONAL POTOSI - UYUNI",	"REGIONAL SANTA CRUZ - TRES CRUCES",	
    "132 - BANCO NACIONAL DE BOLIVIA - COCHABAMBA",	"213 - NORMALIZADORA LA PAZ",	"264 - SOL AMIGO EL ALTO",	"310 - NORMALIZACIÓN COCHABAMBA",	"367 - JTI BOLIVIA",	"602 - CENTRO TARIJA",	"716 - ALTO SAN PEDRO",	"REGIONAL COCHABAMBA - PUNATA",	"REGIONAL EL ALTO - SANTIAGO",	"REGIONAL LA PAZ - PERIFERICA",	"REGIONAL POTOSI - VILLAZON",	"REGIONAL SANTA CRUZ - VILLA 1RO DE MAYO",	
    "133 - BANCO NACIONAL DE BOLIVIA - SUCRE",	"214 - SOL AMIGO LA PAZ",	"265 - AGENCIA MOVÍL",	"311 - RECAUDADORA JORDAN",	"400 - REGIONAL ORURO",	"603 - SUR",	"718 - NORTE",	"REGIONAL COCHABAMBA - QUILLACOLLO",	"REGIONAL EL ALTO - SATELITE",	"REGIONAL LA PAZ - PLAZA EGUINO",	"REGIONAL SANTA CRUZ - ALTO SAN PEDRO",	"REGIONAL SANTA CRUZ - VILLA PRIMERO DE MAYO",	
    "134 - BANCO NACIONAL DE BOLIVIA - ORURO",	"216 - GRAN PODER",	"266 - MERCADO EL CARMEN RÍO SECO",	"312 - SACABA",	"401 - CENTRAL",	"604 - SOL AMIGO TARIJA",	"719 - SOL AMIGO SANTA CRUZ",	"REGIONAL COCHABAMBA - SACABA",	"REGIONAL EL ALTO - SENKATA",	"REGIONAL LA PAZ - RURRENABAQUE",	"REGIONAL SANTA CRUZ - BELEN",	"REGIONAL SANTA CRUZ - VIRGEN DE LUJAN",	
    "135 - BANCO NACIONAL DE BOLIVIA - POTOSI",	"217 - VINO TINTO",	"267 - SANTIAGO II",	"313 - VILLA GALINDO",	"407 - NORMALIZACIÓN ORURO",	"605 - 15 DE ABRIL",	"721 - ARROYO CONCEPCIÓN",	"REGIONAL COCHABAMBA - SAN MARTIN",	"REGIONAL EL ALTO - TERMINAL",	"REGIONAL LA PAZ - SAN MIGUEL",	"REGIONAL SANTA CRUZ - CASCO VIEJO",	"REGIONAL SANTA CRUZ - WARNES",	
    "136 - BANCO NACIONAL DE BOLIVIA - TARIJA",	"218 - VILLA ARMONIA",	"269 - VENTILLA",	"314 - PUNATA",	"408 - PUNTO AMIGO ORURO",	"606 - NORMALIZACIÓN TARIJA",	"722 - AGENCIA MÓVIL STC",	"REGIONAL COCHABAMBA - TAMBORADA",	"REGIONAL EL ALTO - VENTILLA CALAMARCA",	"REGIONAL LA PAZ - SAN PEDRO",	"REGIONAL SANTA CRUZ - EL CARMEN",	"REGIONAL SANTA CRUZ - YAPACANI",	
    "137 - BANCO NACIONAL DE BOLIVIA - BENI",	"219 - OBRAJES",	"270 - 12 DE OCTUBRE",	"315 - REMESADORA",	"409 - VIRGEN DEL SOCAVÓN",	"607 - TABLADITA",	"723 - PAMPA DE LA ISLA",	"REGIONAL COCHABAMBA - VINTO",	"REGIONAL EL ALTO - VIACHA",	"REGIONAL LA PAZ - SOPOCACHI",	"REGIONAL SANTA CRUZ - EQUIPETROL",	"REGIONAL SUCRE - 25 DE MAYO",	
    
]
ESTADOS = [
    "Seleccione una opcion", "POR EXTRAER", "POR REASIGNAR", "POR ENVIAR",
    "ENTREGADO", "ENVIADO", "ANULADO", "POR ETIQUETAR",
]
DIRECCIONES = [
    "Seleccione una opcion", "POLYSISTEMAS", "EVARISTO VALLE", "LA PAZ - CAMACHO",
    "SAN PEDRO OF NAL", "SAN MIGUEL", "ZONA SUR", "12 DE OCTUBLE", "PALENQUE",
    "SATELITE", "COCHABAMBA", "SANTA CRUZ", "SUCRE", "ORURO", "POTOSI",
    "Banco Sol", "Banco BNB", "Banco Fie", "Molino Andino", "Prendamas",
    "Banco Fortaleza", "Previcor Corredores", "Wcs - Bolivia",
    "Industria y comercio Aliconsumo", "Banco Nacional de Bolivia - La Paz",
    "Banco Nacional de Bolivia Sucre", "Banco Nacional de Bolivia Potosi", "Banco Nacional de Bolivia Tarija",
    "Banco Nacional de Bolivia Oruro",
]
REGIONALES = ["Seleccione una opcion", "LA PAZ", "COCHABAMBA", "SANTA CRUZ", "SUCRE", "ORURO", "POTOSI", "OTRA"]
ZONAS = ["Seleccione una opcion", "ZONA SUR", "ZONA ESTE", "ZONA CENTRO", "ZONA NORTE", "EL ALTO", "PROVINCIA", "OTRA"]
RUTA_COLS = ["REGIONAL", "CENTRO DE ACOPIO", "AGENCIAS", "FECHA LIMITE DE INGRESO (SE1)", "FECHA LIMITE DE INGRESO (SR1)", "FECHA DE RECOJO"]
CAJAS_COLS = ["SOLICITANTE", "CLIENTE", "AGENCIA", "FECHA DE SALIDA", "CANTIDAD DE CAJAS SOLICITADAS", "NRO. WORKORDER", "CAJAS REUTILIZADAS", "CANTIDAD DE CINTILLOS", "OBSERVACIONES"]

# =========================================================
# SESSION STATE
# =========================================================
def ss(name, default):
    if name not in st.session_state:
        st.session_state[name] = default

ss("rows", None)
ss("rutas", None)
ss("cajas", None)
ss("movimientos_stock", None)
ss("page", "solicitudes")
ss("caja_form_version", 0)
ss("stock_form_version", 0)
ss("mostrar_editor_historial_salidas", False)
ss("form_version", 0)
ss("tabla_version", 0)
ss("ruta_form_version", 0)
ss("editing", None)
ss("editing_key", "")
ss("session_id", uuid.uuid4().hex)
ss("registro_activo", False)
ss("ruta_activo", False)
ss("usuario_nombre", "")
ss("autenticado", False)

# =========================================================
# SUPABASE
# =========================================================
@st.cache_resource
def obtener_supabase() -> Client:
    try:
        url = st.secrets["SUPABASE_URL"]
        key = (
            st.secrets.get("SUPABASE_SECRET_KEY")
            or st.secrets.get("SUPABASE_SERVICE_ROLE_KEY")
            or st.secrets.get("SUPABASE_KEY")
        )
    except Exception as e:
        raise RuntimeError("Configura SUPABASE_URL y SUPABASE_KEY en Streamlit Secrets.") from e
    if not url or not key:
        raise RuntimeError("Faltan SUPABASE_URL y/o SUPABASE_KEY en Streamlit Secrets.")
    return create_client(url, key)


def db():
    return obtener_supabase()


def fecha_local_hoy():
    return datetime.now(ZONA_HORARIA_APP).date()


def convertir_fecha(valor):
    """Convierte ISO o DD/MM/YYYY sin invertir mes y día."""
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    texto = str(valor or "").strip()
    if not texto:
        return None
    for formato in ("%Y-%m-%d", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M:%S.%f", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(texto[:26], formato).date()
        except ValueError:
            pass
    convertido = pd.to_datetime(texto, errors="coerce", dayfirst=False)
    return None if pd.isna(convertido) else convertido.date()


def normalizar(v):
    return str(v if pd.notna(v) else "").strip().upper()

# =========================================================
# SOLICITUDES
# =========================================================
def cargar_solicitudes_supabase():
    response = db().table("solicitudes").select("*").order("id", desc=False).execute()
    data = response.data or []
    rows = []
    for r in data:
        rows.append({
            "_ID_": r.get("id"),
            "CLIENTE": r.get("cliente", ""),
            "NRO SOLICITUD - WO": r.get("nro_solicitud_wo", ""),
            "TIPO DE SOLICITUD": r.get("tipo_solicitud", ""),
            "CENTRO DE COSTO": r.get("centro_costo", ""),
            "PRIORIDAD": r.get("prioridad", ""),
            "CANT - ITEMS": r.get("cantidad_items", 0),
            "ESTADO DE SOLICITUD": r.get("estado_solicitud", ""),
            "DIRECCIÓN": r.get("direccion", ""),
            "FECHA DE INGRESO": r.get("fecha_ingreso", ""),
        })
    return pd.DataFrame(rows, columns=["_ID_"] + COLS).fillna("")


def guardar_solicitud_supabase(nuevo, registro_id=None):
    payload = {
        "cliente": nuevo["CLIENTE"],
        "nro_solicitud_wo": nuevo["NRO SOLICITUD - WO"],
        "tipo_solicitud": nuevo["TIPO DE SOLICITUD"],
        "centro_costo": nuevo.get("CENTRO DE COSTO", ""),
        "prioridad": nuevo["PRIORIDAD"],
        "cantidad_items": int(nuevo["CANT - ITEMS"]),
        "estado_solicitud": nuevo["ESTADO DE SOLICITUD"],
        "direccion": nuevo["DIRECCIÓN"],
        "fecha_ingreso": nuevo["FECHA DE INGRESO"],
    }
    q = db().table("solicitudes")
    return q.update(payload).eq("id", int(registro_id)).execute() if registro_id else q.insert(payload).execute()


def eliminar_solicitudes_supabase(ids):
    for registro_id in ids:
        db().table("solicitudes").delete().eq("id", int(registro_id)).execute()

def cargar_solicitudes_mes_supabase(fecha_inicio, fecha_fin_exclusiva):
    response = (
        db()
        .table("solicitudes")
        .select("*")
        .gte("fecha_ingreso", fecha_inicio.isoformat())
        .lt("fecha_ingreso", fecha_fin_exclusiva.isoformat())
        .order("fecha_ingreso", desc=False)
        .execute()
    )
    data = response.data or []
    rows = []
    for r in data:
        rows.append({
            "_ID_": r.get("id"),
            "CLIENTE": r.get("cliente", ""),
            "NRO SOLICITUD - WO": r.get("nro_solicitud_wo", ""),
            "TIPO DE SOLICITUD": r.get("tipo_solicitud", ""),
            "CENTRO DE COSTO": r.get("centro_costo", ""),
            "PRIORIDAD": r.get("prioridad", ""),
            "CANT - ITEMS": r.get("cantidad_items", 0),
            "ESTADO DE SOLICITUD": r.get("estado_solicitud", ""),
            "DIRECCIÓN": r.get("direccion", ""),
            "FECHA DE INGRESO": r.get("fecha_ingreso", ""),
        })
    return pd.DataFrame(rows, columns=["_ID_"] + COLS).fillna("")

# =========================================================
# PROGRAMACION DE RUTAS
# La tabla programacion_rutas usa columnas directas:
# id, regional, centro_acopio, agencias,
# fecha_limite_ingreso_se1, fecha_limite_ingreso_sr1, fecha_recojo.
# =========================================================
def cargar_rutas_supabase():
    response = (
        db()
        .table("programacion_rutas")
        .select("id, regional, centro_acopio, agencias, fecha_limite_ingreso_se1, fecha_limite_ingreso_sr1, fecha_recojo")
        .order("id", desc=False)
        .execute()
    )
    data = response.data or []
    rows = []
    for r in data:
        rows.append({
            "_ID_RUTA_": r.get("id"),
            "REGIONAL": r.get("regional", ""),
            "CENTRO DE ACOPIO": r.get("centro_acopio", ""),
            "AGENCIAS": r.get("agencias", ""),
            "FECHA LIMITE DE INGRESO (SE1)": r.get("fecha_limite_ingreso_se1", ""),
            "FECHA LIMITE DE INGRESO (SR1)": r.get("fecha_limite_ingreso_sr1", ""),
            "FECHA DE RECOJO": r.get("fecha_recojo", ""),
        })
    return pd.DataFrame(rows, columns=["_ID_RUTA_"] + RUTA_COLS).fillna("")


def guardar_ruta_supabase(datos, ruta_id=None):
    payload = {
        "regional": datos["REGIONAL"],
        "centro_acopio": datos["CENTRO DE ACOPIO"],
        "agencias": datos["AGENCIAS"],
        "fecha_limite_ingreso_se1": datos["FECHA LIMITE DE INGRESO (SE1)"],
        "fecha_limite_ingreso_sr1": datos["FECHA LIMITE DE INGRESO (SR1)"],
        "fecha_recojo": datos["FECHA DE RECOJO"],
    }
    q = db().table("programacion_rutas")
    if ruta_id:
        return q.update(payload).eq("id", int(ruta_id)).execute()
    return q.insert(payload).execute()


def eliminar_ruta_supabase(ruta_id):
    return db().table("programacion_rutas").delete().eq("id", int(ruta_id)).execute()


# =========================================================
# CONTROL DE CAJAS Y CINTILLOS
# =========================================================
def cargar_cajas_supabase():
    response = db().table("control_cajas").select("*").order("id", desc=False).execute()
    rows = []
    for r in response.data or []:
        rows.append({
            "_ID_CAJA_": r.get("id"),
            "SOLICITANTE": r.get("solicitante", ""),
            "CLIENTE": r.get("cliente", ""),
            "AGENCIA": r.get("agencia", ""),
            "FECHA DE SALIDA": r.get("fecha_salida", ""),
            "CANTIDAD DE CAJAS SOLICITADAS": r.get("cantidad_cajas_solicitadas", 0),
            "NRO. WORKORDER": r.get("nro_workorder", ""),
            "CAJAS REUTILIZADAS": r.get("cajas_reutilizadas", 0),
            "CANTIDAD DE CINTILLOS": r.get("cantidad_cintillos", 0),
            "OBSERVACIONES": r.get("observaciones", ""),
        })
    return pd.DataFrame(rows, columns=["_ID_CAJA_"] + CAJAS_COLS).fillna("")


def cargar_movimientos_stock_supabase():
    response = db().table("movimientos_stock").select("*").order("fecha", desc=True).order("id", desc=True).execute()
    return pd.DataFrame(response.data or [])


def guardar_movimiento_stock(fecha, tipo, cajas_nuevas=0, cajas_reutilizadas=0, cintillos=0, proveedor="", observaciones=""):
    payload = {
        "fecha": fecha.isoformat() if isinstance(fecha, date) else str(fecha),
        "tipo_movimiento": tipo,
        "cajas_nuevas": int(cajas_nuevas or 0),
        "cajas_reutilizadas": int(cajas_reutilizadas or 0),
        "cintillos": int(cintillos or 0),
        "proveedor": (proveedor or "").strip(),
        "observaciones": (observaciones or "").strip(),
        "usuario": st.session_state.get("usuario_nombre", ""),
    }
    return db().table("movimientos_stock").insert(payload).execute()


def obtener_stock_actual():
    mov = cargar_movimientos_stock_supabase()
    if mov.empty:
        return {"CAJAS NUEVAS": 0, "CAJAS REUTILIZADAS": 0, "CINTILLOS": 0}
    return {
        "CAJAS NUEVAS": int(pd.to_numeric(mov.get("cajas_nuevas", 0), errors="coerce").fillna(0).sum()),
        "CAJAS REUTILIZADAS": int(pd.to_numeric(mov.get("cajas_reutilizadas", 0), errors="coerce").fillna(0).sum()),
        "CINTILLOS": int(pd.to_numeric(mov.get("cintillos", 0), errors="coerce").fillna(0).sum()),
    }


def guardar_salida_cajas_supabase(datos):
    payload = {
        "solicitante": datos["SOLICITANTE"],
        "cliente": datos["CLIENTE"],
        "agencia": datos["AGENCIA"],
        "fecha_salida": datos["FECHA DE SALIDA"],
        "cantidad_cajas_solicitadas": int(datos["CANTIDAD DE CAJAS SOLICITADAS"]),
        "nro_workorder": datos["NRO. WORKORDER"],
        "cajas_reutilizadas": int(datos["CAJAS REUTILIZADAS"]),
        "cantidad_cintillos": int(datos["CANTIDAD DE CINTILLOS"]),
        "observaciones": datos["OBSERVACIONES"],
        "usuario": st.session_state.get("usuario_nombre", ""),
    }
    return db().table("control_cajas").insert(payload).execute()


def entero_seguro(valor, default=0):
    """Convierte valores de tabla a entero sin fallar con vacíos o NaN."""
    try:
        numero = pd.to_numeric(valor, errors="coerce")
        if pd.isna(numero):
            return int(default)
        return int(numero)
    except Exception:
        return int(default)


def actualizar_salida_cajas_supabase(datos, salida_id):
    """Actualiza un registro existente del historial de salidas."""
    payload = {
        "solicitante": datos["SOLICITANTE"],
        "cliente": datos["CLIENTE"],
        "agencia": datos["AGENCIA"],
        "fecha_salida": datos["FECHA DE SALIDA"],
        "cantidad_cajas_solicitadas": int(datos["CANTIDAD DE CAJAS SOLICITADAS"]),
        "nro_workorder": datos["NRO. WORKORDER"],
        "cajas_reutilizadas": int(datos["CAJAS REUTILIZADAS"]),
        "cantidad_cintillos": int(datos["CANTIDAD DE CINTILLOS"]),
        "observaciones": datos["OBSERVACIONES"],
        "usuario": st.session_state.get("usuario_nombre", ""),
    }
    return db().table("control_cajas").update(payload).eq("id", int(salida_id)).execute()

# =========================================================
# BLOQUEO
# =========================================================
def _ahora_utc_iso():
    """Devuelve una fecha/hora UTC compatible con PostgreSQL timestamptz."""
    return datetime.now(timezone.utc).isoformat()


def _bloqueo_expirado(valor):
    """Comprueba si last_activity (timestamptz o valor antiguo Unix) expiró."""
    if valor in (None, "", 0, "0"):
        return True

    try:
        # Compatibilidad con registros antiguos que pudieron guardar time.time().
        if isinstance(valor, (int, float)):
            ultima = datetime.fromtimestamp(float(valor), tz=timezone.utc)
        else:
            texto = str(valor).strip()
            try:
                numero = float(texto)
                ultima = datetime.fromtimestamp(numero, tz=timezone.utc)
            except ValueError:
                ultima = pd.to_datetime(texto, utc=True, errors="coerce")
                if pd.isna(ultima):
                    return True
                ultima = ultima.to_pydatetime()

        if ultima.tzinfo is None:
            ultima = ultima.replace(tzinfo=timezone.utc)

        return (datetime.now(timezone.utc) - ultima).total_seconds() > LOCK_TIMEOUT_SECONDS
    except Exception:
        return True


def leer_bloqueo_edicion():
    try:
        response = db().table("app_locks").select("*").eq("id", 1).maybe_single().execute()
        row = response.data
        if not row or not row.get("owner_id"):
            return None

        if _bloqueo_expirado(row.get("last_activity")):
            liberar_bloqueo_edicion(force=True)
            return None

        return row
    except Exception:
        return None


def adquirir_bloqueo_edicion(motivo="registro"):
    usuario = (st.session_state.get("usuario_nombre") or "Usuario").strip()
    session_id = st.session_state.session_id
    actual = leer_bloqueo_edicion()

    if actual and actual.get("owner_id") != session_id:
        return False, actual

    # IMPORTANTE:
    # app_locks.last_activity es timestamp with time zone (timestamptz).
    # No debemos enviar time.time(), porque devuelve un Unix timestamp
    # como 1788320350.168386 y PostgreSQL lo rechaza como fecha.
    payload = {
        "id": 1,
        # app_locks.recurso es NOT NULL en Supabase.
        # Usamos el motivo de la operación como recurso para que
        # nunca se envíe NULL a esa columna.
        "recurso": str(motivo or "registro").strip() or "registro",
        "owner_id": session_id,
        "usuario": usuario,
        "last_activity": _ahora_utc_iso(),
        "motivo": str(motivo or "registro").strip() or "registro",
    }

    db().table("app_locks").upsert(payload, on_conflict="id").execute()
    st.session_state.registro_activo = True
    return True, payload


def renovar_bloqueo_edicion():
    if not st.session_state.get("registro_activo") and not st.session_state.get("ruta_activo"):
        return

    try:
        db().table("app_locks").update(
            {"last_activity": _ahora_utc_iso()}
        ).eq("id", 1).eq("owner_id", st.session_state.session_id).execute()
    except Exception:
        pass


def liberar_bloqueo_edicion(force=False):
    try:
        q = db().table("app_locks").delete().eq("id", 1)
        if not force:
            q = q.eq("owner_id", st.session_state.session_id)
        q.execute()
    except Exception:
        pass
    st.session_state.registro_activo = False
    st.session_state.ruta_activo = False

# =========================================================
# EXPORTAR EXCEL
# =========================================================

def resumen_tipo_solicitud_wo(df):
    """Cuenta SE1, SE2, SE3, SR1 y SR2 en la columna NRO SOLICITUD - WO."""
    tipos = ["SE1", "SE2", "SE3", "SR1", "SR2"]
    conteo = {tipo: 0 for tipo in tipos}

    if df is None or df.empty or "NRO SOLICITUD - WO" not in df.columns:
        return conteo

    serie = df["NRO SOLICITUD - WO"].fillna("").astype(str).str.upper()
    for tipo in tipos:
        conteo[tipo] = int(
            serie.apply(
                lambda valor: len(re.findall(rf"(?<![A-Z0-9]){tipo}(?![A-Z0-9])", valor))
            ).sum()
        )
    return conteo


def excel_bytes_mensual(df):
    """Genera el respaldo mensual con detalle completo y hojas de resumen."""
    out = BytesIO()
    datos = df.drop(columns=["_ID_"], errors="ignore").copy()

    # Asegurar que NRO SOLICITUD - WO esté incluido y visible.
    columnas_preferidas = [
        "NRO SOLICITUD - WO", "CLIENTE", "TIPO DE SOLICITUD",
        "CENTRO DE COSTO", "PRIORIDAD", "CANT - ITEMS",
        "ESTADO DE SOLICITUD", "DIRECCIÓN", "FECHA DE INGRESO"
    ]
    columnas = [c for c in columnas_preferidas if c in datos.columns]
    resto = [c for c in datos.columns if c not in columnas]
    datos = datos[columnas + resto]

    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        datos.to_excel(writer, sheet_name="SOLICITUDES", index=False)

        # Resumen por estado
        if "ESTADO DE SOLICITUD" in datos.columns:
            resumen_estado = (
                datos["ESTADO DE SOLICITUD"]
                .fillna("SIN ESTADO")
                .astype(str)
                .value_counts()
                .rename_axis("ESTADO DE SOLICITUD")
                .reset_index(name="CANTIDAD")
            )
        else:
            resumen_estado = pd.DataFrame(columns=["ESTADO DE SOLICITUD", "CANTIDAD"])
        resumen_estado.to_excel(writer, sheet_name="RESUMEN_ESTADO", index=False)

        # Resumen por prioridad
        if "PRIORIDAD" in datos.columns:
            resumen_prioridad = (
                datos["PRIORIDAD"]
                .fillna("SIN PRIORIDAD")
                .astype(str)
                .value_counts()
                .rename_axis("PRIORIDAD")
                .reset_index(name="CANTIDAD")
            )
        else:
            resumen_prioridad = pd.DataFrame(columns=["PRIORIDAD", "CANTIDAD"])
        resumen_prioridad.to_excel(writer, sheet_name="RESUMEN_PRIORIDAD", index=False)

        # Resumen de tipos SE/SR desde NRO SOLICITUD - WO.
        # Se construye directamente para evitar referencias a variables locales
        # que puedan quedar sin inicializar al cambiar de mes.
        conteo_wo = resumen_tipo_solicitud_wo(datos)
        pd.DataFrame(
            {
                "TIPO WO": list(conteo_wo.keys()),
                "CANTIDAD": list(conteo_wo.values()),
            }
        ).to_excel(writer, sheet_name="RESUMEN_SE_SR", index=False)

        # Resumen general, incluyendo cantidad de solicitudes/WO.
        resumen_general = pd.DataFrame({
            "INDICADOR": [
                "TOTAL DE REGISTROS",
                "TOTAL DE SOLICITUDES / WO",
                "TOTAL DE ITEMS",
            ],
            "VALOR": [
                len(datos),
                datos["NRO SOLICITUD - WO"].nunique(dropna=True) if "NRO SOLICITUD - WO" in datos.columns else 0,
                pd.to_numeric(datos["CANT - ITEMS"], errors="coerce").fillna(0).sum()
                if "CANT - ITEMS" in datos.columns else 0,
            ]
        })
        resumen_general.to_excel(writer, sheet_name="RESUMEN_GENERAL", index=False)

        # Ajustar ancho de columnas para facilitar la lectura.
        for ws in writer.book.worksheets:
            for col_cells in ws.columns:
                max_len = 0
                for cell in col_cells:
                    value = "" if cell.value is None else str(cell.value)
                    max_len = max(max_len, len(value))
                ws.column_dimensions[col_cells[0].column_letter].width = min(max(max_len + 2, 12), 45)

    return out.getvalue()

def excel_bytes(df, rutas=None):
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.drop(columns=["_ID_"], errors="ignore").to_excel(writer, sheet_name="SOLICITUDES", index=False)
        if rutas is not None and not rutas.empty:
            rutas.drop(columns=["_ID_RUTA_"], errors="ignore").to_excel(writer, sheet_name=RUTA_SHEET, index=False)
    return out.getvalue()


def excel_bytes_historial_salidas(df):
    """Genera un archivo Excel del historial de salidas de cajas y cintillos."""
    out = BytesIO()
    datos = df.drop(columns=["_ID_CAJA_"], errors="ignore").copy()

    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        datos.to_excel(writer, sheet_name="HISTORIAL SALIDAS", index=False)
        ws = writer.sheets["HISTORIAL SALIDAS"]

        # Ajustar el ancho de las columnas para facilitar la lectura.
        for col_cells in ws.columns:
            max_len = 0
            for cell in col_cells:
                value = "" if cell.value is None else str(cell.value)
                max_len = max(max_len, len(value))
            ws.column_dimensions[col_cells[0].column_letter].width = min(
                max(max_len + 2, 12), 45
            )

    return out.getvalue()

# =========================================================
# LOGIN
# =========================================================
def obtener_usuarios_login():
    try:
        bloque = st.secrets.get("usuarios", {})
        if hasattr(bloque, "items"):
            return {str(k).strip().lower(): str(v) for k, v in bloque.items()}
    except Exception:
        pass
    return {}


def nombres_usuarios():
    return {
        "admin": "admin",
        "alfredo": "Alfredo",
        "jonathan": "Jonathan",
        "luis": "Luis",
    }


def obtener_rol():
    usuario = str(st.session_state.get("usuario_login", "") or "").strip().lower()
    return {
        "admin": "admin",
        "alfredo": "usuario",
        "jonathan": "usuario",
        "luis": "usuario",
    }.get(usuario, "usuario")


def es_admin():
    return obtener_rol() == "admin"


def puede_modificar():
    return obtener_rol() in ("admin", "usuario")


def mostrar_login():
    st.markdown(
        '<div class="login-wrapper"><div class="login-icon">🔐</div><div class="login-title">Iniciar sesión</div><div class="login-subtitle">Acceso al sistema de despacho</div></div>',
        unsafe_allow_html=True,
    )
    usuarios = obtener_usuarios_login()
    if not usuarios:
        st.error("❌ No hay usuarios configurados en Streamlit Secrets.")
        return
    _, centro, _ = st.columns([1, 1.25, 1])
    with centro:
        with st.form("login_jc_control"):
            usuario = st.text_input("👤 Usuario", placeholder="Ingrese su usuario")
            password = st.text_input("🔑 Contraseña", type="password", placeholder="Ingrese su contraseña")
            ingresar = st.form_submit_button("🔐 INGRESAR", use_container_width=True)
    if ingresar:
        key = usuario.strip().lower()
        if key in usuarios and password == usuarios[key]:
            st.session_state.autenticado = True
            st.session_state.usuario_login = key
            st.session_state.usuario_nombre = nombres_usuarios().get(key, usuario.strip())
            st.rerun()
        else:
            st.error("❌ Usuario o contraseña incorrectos.")

# =========================================================
# FORMULARIO PROGRAMACION DE RUTAS
# =========================================================
def formulario_nueva_ruta():
    st.markdown("### ➕ Registrar programación de ruta")
    with st.form(f"ruta_form_{st.session_state.ruta_form_version}", clear_on_submit=False):
        a, b = st.columns(2)
        with a:
            regional = st.selectbox("🌎 Regional", REGIONALES)
            centro_acopio = st.text_input("🏢 Centro de acopio", placeholder="Ej.: POLYSISTEMAS")
            agencias = st.text_input("🏪 Agencias", placeholder="Ej.: San Miguel, Zona Sur")
        with b:
            fecha_se1 = st.date_input("📅 Fecha límite de ingreso (SE1)", value=fecha_local_hoy(), format="DD/MM/YYYY")
            fecha_sr1 = st.date_input("📅 Fecha límite de ingreso (SR1)", value=fecha_local_hoy(), format="DD/MM/YYYY")
            fecha_recojo = st.date_input("🚚 Fecha de recojo", value=fecha_local_hoy(), format="DD/MM/YYYY")

        x, y = st.columns([4, 1])
        with x:
            guardar = st.form_submit_button("💾 GUARDAR PROGRAMACIÓN", use_container_width=True)
        with y:
            limpiar = st.form_submit_button("✕ CANCELAR", type="primary", use_container_width=True)

    if limpiar:
        liberar_bloqueo_edicion()
        st.session_state.ruta_form_version += 1
        st.rerun()

    if guardar:
        if regional == REGIONALES[0]:
            st.error("Selecciona una Regional antes de guardar.")
            return
        if not centro_acopio.strip():
            st.error("Ingresa el Centro de Acopio antes de guardar.")
            return
        if not agencias.strip():
            st.error("Ingresa las Agencias antes de guardar.")
            return

        if fecha_se1 > fecha_recojo or fecha_sr1 > fecha_recojo:
            st.error("La fecha de recojo no puede ser anterior a las fechas límite de ingreso.")
            return

        try:
            datos = {
                "REGIONAL": regional,
                "CENTRO DE ACOPIO": centro_acopio.strip(),
                "AGENCIAS": agencias.strip(),
                "FECHA LIMITE DE INGRESO (SE1)": fecha_se1.isoformat(),
                "FECHA LIMITE DE INGRESO (SR1)": fecha_sr1.isoformat(),
                "FECHA DE RECOJO": fecha_recojo.isoformat(),
            }
            guardar_ruta_supabase(datos)
            st.session_state.rutas = cargar_rutas_supabase()
            liberar_bloqueo_edicion()
            st.session_state.ruta_form_version += 1
            st.success("✅ Programación de ruta guardada correctamente en Supabase.")
            st.rerun()
        except Exception as e:
            liberar_bloqueo_edicion()
            st.error(f"❌ No se pudo guardar la programación: {e}")


# =========================================================
# CABECERA Y LOGIN
# =========================================================
st.markdown('<div class="titulo">📝 JC CONTROL DE SOLICITUDES — DESPACHO</div>', unsafe_allow_html=True)

if not st.session_state.autenticado:
    mostrar_login()
    st.stop()

# =========================================================
# CIERRE AUTOMÁTICO POR INACTIVIDAD
# 8 minutos sin interacción del usuario.
# =========================================================
_refresh_count = st_autorefresh(interval=AUTO_REFRESH_INTERVAL_MS, key="control_inactividad")
_ultimo_refresh = st.session_state.get("_ultimo_refresh_inactividad")
_ahora = time.time()
if "_ultima_actividad" not in st.session_state:
    st.session_state._ultima_actividad = _ahora
elif _ultimo_refresh is not None and _refresh_count == _ultimo_refresh:
    st.session_state._ultima_actividad = _ahora
st.session_state._ultimo_refresh_inactividad = _refresh_count

_inactividad = _ahora - st.session_state._ultima_actividad
if _inactividad >= INACTIVITY_TIMEOUT_SECONDS:
    liberar_bloqueo_edicion()
    for k in ["rows", "rutas", "cajas", "movimientos_stock"]:
        st.session_state[k] = None
    st.session_state.editing = None
    st.session_state.editing_key = ""
    st.session_state.usuario_nombre = ""
    st.session_state.autenticado = False
    st.session_state.page = "solicitudes"
    st.session_state.pop("_ultima_actividad", None)
    st.session_state.pop("_ultimo_refresh_inactividad", None)
    st.rerun()

if _inactividad >= 5 * 60:
    minutos_restantes = max(1, int((INACTIVITY_TIMEOUT_SECONDS - _inactividad + 59) // 60))
    st.warning(f"⏳ Tu sesión se cerrará por inactividad en aproximadamente {minutos_restantes} minuto(s).")

u1, u2, u3 = st.columns([1, 3, 1])
with u1:
    st.markdown(f"**👤 Usuario:** `{st.session_state.usuario_nombre}`")
    if es_admin():
        st.caption("🔴 ADMIN — ACCESO TOTAL")
    else:
        st.caption("🟢 USUARIO — REGISTRAR Y EDITAR")
with u2:
    st.markdown("**🗄️ Base de datos:** `PostgreSQL`")
with u3:
    if st.button("🚪 SALIR", use_container_width=True):
        liberar_bloqueo_edicion()
        for k in ["rows", "rutas", "cajas", "movimientos_stock"]:
            st.session_state[k] = None
        st.session_state.editing = None
        st.session_state.editing_key = ""
        st.session_state.usuario_nombre = ""
        st.session_state.autenticado = False
        st.session_state.page = "solicitudes"
        st.rerun()

# =========================================================
# CARGA INICIAL
# =========================================================
if st.session_state.rows is None or st.session_state.rutas is None or st.session_state.cajas is None or st.session_state.movimientos_stock is None:
    try:
        st.session_state.rows = cargar_solicitudes_supabase()
        st.session_state.rutas = cargar_rutas_supabase()
        st.session_state.cajas = cargar_cajas_supabase()
        st.session_state.movimientos_stock = cargar_movimientos_stock_supabase()
    except Exception as e:
        st.error(f"❌ No se pudieron cargar los datos desde Supabase: {e}")
        st.stop()

# =========================================================
# NAVEGACION
# =========================================================
c1, c2, c3, c4 = st.columns([1, 1.35, 1.45, 2.2])
with c1:
    if st.button("🔄 ACTUALIZAR", use_container_width=True):
        try:
            st.session_state.rows = cargar_solicitudes_supabase()
            st.session_state.rutas = cargar_rutas_supabase()
            st.session_state.editing = None
            st.session_state.editing_key = ""
            st.toast("🔄 Datos actualizados desde Supabase.", icon="🔄")
            st.rerun()
        except Exception as e:
            st.error(f"❌ No se pudieron actualizar los datos: {e}")
with c2:
    if st.button("🗓️ PROGRAMACIÓN DE RUTAS", use_container_width=True):
        liberar_bloqueo_edicion()
        st.session_state.editing = None
        st.session_state.editing_key = ""
        st.session_state.page = "rutas"
        st.rerun()
with c3:
    if st.button("📦 CONTROL DE CAJAS Y CINTILLOS", use_container_width=True):
        liberar_bloqueo_edicion()
        st.session_state.editing = None
        st.session_state.editing_key = ""
        st.session_state.page = "cajas"
        st.rerun()

# =========================================================
# PAGINA CONTROL DE CAJAS Y CINTILLOS
# =========================================================
if st.session_state.page == "cajas":
    st.subheader("📦 CONTROL DE CAJAS Y CINTILLOS")
    st.caption("Control de ingresos, stock actual y salidas de cajas y cintillos.")
    if st.button("⬅ VOLVER A SOLICITUDES", key="volver_cajas"):
        st.session_state.page = "solicitudes"
        st.rerun()

    try:
        stock = obtener_stock_actual()
        m1, m2, m3 = st.columns(3)
        m1.metric("📦 CAJAS NUEVAS", stock["CAJAS NUEVAS"])
        m2.metric("♻️ CAJAS REUTILIZADAS", stock["CAJAS REUTILIZADAS"])
        m3.metric("🏷️ CINTILLOS", stock["CINTILLOS"])
    except Exception as e:
        st.error(f"❌ No se pudo calcular el stock: {e}")
        st.stop()

    if es_admin():
        # Estado para limpiar y replegar automáticamente el formulario de stock.
        st.session_state.setdefault("mostrar_ingreso_stock", False)
        st.session_state.setdefault("stock_expander_version", 0)

        titulo_stock = (
            "➕ INGRESAR LLEGADA DE STOCK"
            + ("\u200b" * int(st.session_state.stock_expander_version))
        )

        with st.expander(
            titulo_stock,
            expanded=st.session_state.mostrar_ingreso_stock
        ):
            st.caption("🔴 Solo ADMIN puede aumentar o ajustar el stock.")
            with st.form(f"stock_form_{st.session_state.stock_form_version}"):
                a,b = st.columns(2)
                with a:
                    fecha_ingreso_stock = st.date_input(
                        "📅 Fecha de llegada",
                        value=fecha_local_hoy(),
                        format="DD/MM/YYYY"
                    )
                    cajas_nuevas_in = st.number_input(
                        "📦 Cajas nuevas recibidas",
                        min_value=0,
                        step=1,
                        value=0
                    )
                    cajas_reutilizadas_in = st.number_input(
                        "♻️ Cajas reutilizadas ingresadas",
                        min_value=0,
                        step=1,
                        value=0
                    )
                with b:
                    cintillos_in = st.number_input(
                        "🏷️ Cintillos recibidos",
                        min_value=0,
                        step=1,
                        value=0
                    )
                    proveedor = st.text_input("🚚 Proveedor / procedencia")
                    obs_ingreso = st.text_area("📝 Observaciones")

                x_stock, y_stock = st.columns([4, 1])
                with x_stock:
                    guardar_stock = st.form_submit_button(
                        "💾 GUARDAR INGRESO DE STOCK",
                        use_container_width=True
                    )
                with y_stock:
                    cancelar_stock = st.form_submit_button(
                        "✕ CANCELAR",
                        type="primary",
                        use_container_width=True
                    )

            # CANCELAR: no guarda nada, limpia el formulario y lo repliega.
            if cancelar_stock:
                st.session_state.stock_form_version += 1
                st.session_state.mostrar_ingreso_stock = False
                st.session_state.stock_expander_version += 1
                st.rerun()

            if guardar_stock:
                if int(cajas_nuevas_in) + int(cajas_reutilizadas_in) + int(cintillos_in) <= 0:
                    st.error("Ingresa al menos una cantidad mayor a cero.")
                else:
                    guardar_movimiento_stock(
                        fecha_ingreso_stock,
                        "INGRESO",
                        cajas_nuevas_in,
                        cajas_reutilizadas_in,
                        cintillos_in,
                        proveedor,
                        obs_ingreso
                    )
                    st.session_state.cajas = cargar_cajas_supabase()
                    st.session_state.movimientos_stock = cargar_movimientos_stock_supabase()
                    st.session_state.stock_form_version += 1
                    st.session_state.mostrar_ingreso_stock = False
                    st.session_state.stock_expander_version += 1
                    st.success("✅ Ingreso de stock registrado correctamente.")
                    st.rerun()
    else:
        st.info("🔒 Solo el usuario ADMIN puede ingresar o modificar el stock.")

    # Registro de salida como sección desplegable, igual que "Ingresar llegada de stock".
    # Se usa una versión invisible del título para forzar que Streamlit reconstruya
    # el expander cuando se presiona CANCELAR o se guarda el registro.
    st.session_state.setdefault("mostrar_registro_salida", False)
    st.session_state.setdefault("salida_expander_version", 0)

    titulo_salida = (
        "📝 REGISTRAR SALIDA DE CAJAS Y CINTILLOS"
        + ("\u200b" * int(st.session_state.salida_expander_version))
    )

    with st.expander(
        titulo_salida,
        expanded=st.session_state.mostrar_registro_salida,
    ):
        with st.form(f"salida_cajas_form_{st.session_state.caja_form_version}"):
            a,b = st.columns(2)
            with a:
                solicitante = st.text_input("SOLICITANTE")
                cliente_caja = st.selectbox("CLIENTE", CLIENTES[1:])
                agencia_caja = st.selectbox(
                    "AGENCIA",
                    CENTROS_COSTO[1:],
                    index=None,
                    placeholder="Seleccione una agencia"
                )
                fecha_salida = st.date_input("FECHA DE SALIDA", value=fecha_local_hoy(), format="DD/MM/YYYY")
                cajas_solicitadas = st.number_input("CANTIDAD DE CAJAS SOLICITADAS", min_value=0, step=1, value=0)
            with b:
                nro_wo_caja = st.text_input("NRO. WORKORDER")
                cajas_reutilizadas = st.number_input("CAJAS REUTILIZADAS", min_value=0, step=1, value=0)
                cintillos_salida = st.number_input("CANTIDAD DE CINTILLOS", min_value=0, step=1, value=0)
                observaciones_caja = st.text_area("OBSERVACIONES")

            x, y = st.columns([4, 1])
            with x:
                guardar_salida = st.form_submit_button("📤 REGISTRAR SALIDA", use_container_width=True)
            with y:
                cancelar_salida = st.form_submit_button("✕ CANCELAR", type="primary", use_container_width=True)

        if cancelar_salida:
            # Limpiar el formulario y cerrar visualmente el expander.
            st.session_state.caja_form_version += 1
            st.session_state.mostrar_registro_salida = False
            st.session_state.salida_expander_version += 1
            st.rerun()

        if guardar_salida:
            cajas_nuevas_salida = int(cajas_solicitadas) - int(cajas_reutilizadas)
            if not solicitante.strip() or not agencia_caja or (int(cajas_solicitadas) <= 0 and int(cintillos_salida) <= 0):
                st.error("Completa SOLICITANTE y AGENCIA, e ingresa al menos una cantidad mayor a cero de CAJAS o CINTILLOS.")
            elif cajas_nuevas_salida < 0:
                st.error("Las cajas reutilizadas no pueden ser mayores que las cajas solicitadas.")
            elif cajas_nuevas_salida > stock["CAJAS NUEVAS"] or int(cajas_reutilizadas) > stock["CAJAS REUTILIZADAS"] or int(cintillos_salida) > stock["CINTILLOS"]:
                st.error("❌ Stock insuficiente para registrar esta salida.")
            else:
                datos_salida = {"SOLICITANTE": solicitante.strip(), "CLIENTE": cliente_caja, "AGENCIA": agencia_caja.strip(), "FECHA DE SALIDA": fecha_salida.isoformat(), "CANTIDAD DE CAJAS SOLICITADAS": int(cajas_solicitadas), "NRO. WORKORDER": nro_wo_caja.strip(), "CAJAS REUTILIZADAS": int(cajas_reutilizadas), "CANTIDAD DE CINTILLOS": int(cintillos_salida), "OBSERVACIONES": observaciones_caja.strip()}
                guardar_salida_cajas_supabase(datos_salida)
                guardar_movimiento_stock(fecha_salida, "SALIDA", -cajas_nuevas_salida, -int(cajas_reutilizadas), -int(cintillos_salida), "", f"Salida WO: {nro_wo_caja.strip()}")
                st.session_state.cajas = cargar_cajas_supabase()
                st.session_state.movimientos_stock = cargar_movimientos_stock_supabase()
                st.session_state.caja_form_version += 1
                st.session_state.mostrar_registro_salida = False
                st.session_state.salida_expander_version += 1
                st.success("✅ Salida registrada y stock actualizado automáticamente.")
                st.rerun()

    st.markdown("### 📋 Historial de salidas")
    cajas_df = st.session_state.cajas if st.session_state.cajas is not None else cargar_cajas_supabase()

    # Solo el usuario ADMIN puede filtrar y descargar el historial completo.
    if es_admin():
        st.session_state.setdefault("mostrar_filtro_historial_salidas", False)

        b1, b2 = st.columns([1.2, 1.2])
        with b1:
            if st.button(
                "🔎 FILTRAR HISTORIAL",
                use_container_width=True,
                key="btn_filtro_historial_salidas",
            ):
                # No forzamos un segundo rerun: así no se vuelve a abrir
                # accidentalmente el formulario de REGISTRAR SALIDA.
                st.session_state.mostrar_filtro_historial_salidas = (
                    not st.session_state.mostrar_filtro_historial_salidas
                )

        # El historial descargado será exactamente el que se muestra después
        # de aplicar los filtros seleccionados.
        tabla_filtrada = cajas_df.copy()

        if st.session_state.mostrar_filtro_historial_salidas:
            st.markdown("#### 🔎 Filtros del historial de salidas")

            # Opciones del filtro por MES en español, por ejemplo:
            # SEPTIEMBRE 2026.
            meses_es = {
                1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
                5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
                9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE",
            }

            if cajas_df.empty:
                st.info("No hay registros de salidas para filtrar.")
            else:
                fechas_hist = pd.to_datetime(
                    cajas_df["FECHA DE SALIDA"], errors="coerce"
                ).dropna()

                if fechas_hist.empty:
                    opciones_mes = [("TODOS", None)]
                else:
                    periodos = sorted(
                        fechas_hist.dt.to_period("M").unique().tolist(),
                        reverse=True,
                    )
                    opciones_mes = [("TODOS", None)] + [
                        (f"{meses_es[int(periodo.month)]} {int(periodo.year)}", periodo)
                        for periodo in periodos
                    ]

                # ======================================================
                # FILTROS INDEPENDIENTES Y COMBINADOS
                # Cada selector se alimenta del historial COMPLETO.
                # Así:
                #   - MES funciona por separado.
                #   - CLIENTE funciona por separado.
                #   - Si se seleccionan ambos, se cumplen AMBAS condiciones.
                # ======================================================
                f1, f2 = st.columns(2)

                with f1:
                    etiquetas_mes = [etiqueta for etiqueta, _ in opciones_mes]
                    mes_historial = st.selectbox(
                        "📅 FILTRAR POR MES",
                        etiquetas_mes,
                        key="filtro_historial_mes",
                    )

                with f2:
                    clientes_historial = sorted(
                        [
                            str(cliente).strip()
                            for cliente in cajas_df["CLIENTE"].dropna().unique().tolist()
                            if str(cliente).strip()
                        ]
                    )
                    opciones_cliente = ["TODOS"] + clientes_historial
                    cliente_historial = st.selectbox(
                        "👤 FILTRAR POR CLIENTE",
                        opciones_cliente,
                        key="filtro_historial_cliente",
                    )

                # Empezamos siempre desde el historial completo.
                # De esta forma los filtros son independientes.
                tabla_filtrada = cajas_df.copy()

                # 1) FILTRO POR MES (si se seleccionó un mes).
                periodo_seleccionado = dict(opciones_mes).get(mes_historial)
                if periodo_seleccionado is not None:
                    periodos_tabla = pd.to_datetime(
                        tabla_filtrada["FECHA DE SALIDA"], errors="coerce"
                    ).dt.to_period("M")
                    tabla_filtrada = tabla_filtrada[
                        periodos_tabla == periodo_seleccionado
                    ]

                # 2) FILTRO POR CLIENTE (si se seleccionó un cliente).
                # Si también hay un mes seleccionado, este filtro se aplica
                # sobre el resultado anterior, cumpliendo AMBAS condiciones.
                if cliente_historial != "TODOS":
                    tabla_filtrada = tabla_filtrada[
                        tabla_filtrada["CLIENTE"].astype(str).str.strip() == cliente_historial
                    ]

                # Indicamos claramente qué filtro está activo.
                filtros_activos = []
                if mes_historial != "TODOS":
                    filtros_activos.append(f"Mes: {mes_historial}")
                if cliente_historial != "TODOS":
                    filtros_activos.append(f"Cliente: {cliente_historial}")

                if filtros_activos:
                    if len(filtros_activos) == 2:
                        st.caption(
                            "Mostrando registros que cumplen ambos filtros: "
                            + "  •  ".join(filtros_activos)
                        )
                    else:
                        st.caption(
                            "Mostrando registros filtrados por: "
                            + "  •  ".join(filtros_activos)
                        )
                else:
                    st.caption("Mostrando todo el historial de salidas.")

                # Resumen de cantidades correspondientes exactamente al filtro aplicado.
                total_cajas_filtradas = int(
                    pd.to_numeric(
                        tabla_filtrada["CANTIDAD DE CAJAS SOLICITADAS"],
                        errors="coerce",
                    ).fillna(0).sum()
                ) if "CANTIDAD DE CAJAS SOLICITADAS" in tabla_filtrada.columns else 0

                total_cintillos_filtrados = int(
                    pd.to_numeric(
                        tabla_filtrada["CANTIDAD DE CINTILLOS"],
                        errors="coerce",
                    ).fillna(0).sum()
                ) if "CANTIDAD DE CINTILLOS" in tabla_filtrada.columns else 0

                st.markdown("#### 📊 Totales del filtro")
                t1, t2 = st.columns(2)
                with t1:
                    st.metric("📦 TOTAL DE CAJAS SOLICITADAS", f"{total_cajas_filtradas:,}")
                with t2:
                    st.metric("🏷️ TOTAL DE CINTILLOS SOLICITADOS", f"{total_cintillos_filtrados:,}")

        with b2:
            nombre_archivo = (
                "historial_salidas_"
                + datetime.now(ZONA_HORARIA_APP).strftime("%Y%m%d_%H%M%S")
                + ".xlsx"
            )
            st.download_button(
                "⬇️ DESCARGAR HISTORIAL",
                data=excel_bytes_historial_salidas(tabla_filtrada),
                file_name=nombre_archivo,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="descargar_historial_salidas",
                disabled=tabla_filtrada.empty,
            )

        # La edición se selecciona directamente desde la columna ✏️ EDITAR de la tabla.
        st.session_state.setdefault("salida_id_editar_inline", None)
        salida_id_editar = st.session_state.get("salida_id_editar_inline")

        if salida_id_editar is not None:
            try:
                salida_id_editar = int(salida_id_editar)
                if salida_id_editar not in tabla_filtrada["_ID_CAJA_"].astype(int).tolist():
                    salida_id_editar = None
                    st.session_state.salida_id_editar_inline = None
            except Exception:
                salida_id_editar = None
                st.session_state.salida_id_editar_inline = None

        if salida_id_editar is not None:
            st.markdown("#### ✏️ Editar registro de salida")
            st.caption("🔴 Solo ADMIN puede editar. El stock se ajustará automáticamente según los cambios realizados.")

            registro_editar = tabla_filtrada[
                tabla_filtrada["_ID_CAJA_"].astype(int) == int(salida_id_editar)
            ].iloc[0].to_dict()

            cliente_actual = str(registro_editar.get("CLIENTE", "")).strip()
            clientes_edicion = list(CLIENTES[1:])
            if cliente_actual and cliente_actual not in clientes_edicion:
                clientes_edicion.append(cliente_actual)

            agencia_actual = str(registro_editar.get("AGENCIA", "")).strip()
            agencias_edicion = list(CENTROS_COSTO[1:])
            if agencia_actual and agencia_actual not in agencias_edicion:
                agencias_edicion.append(agencia_actual)

            fecha_edicion = convertir_fecha(registro_editar.get("FECHA DE SALIDA")) or fecha_local_hoy()

            with st.form(f"editar_salida_form_{int(salida_id_editar)}"):
                e1, e2 = st.columns(2)
                with e1:
                    edit_solicitante = st.text_input(
                        "SOLICITANTE", value=str(registro_editar.get("SOLICITANTE", ""))
                    )
                    edit_cliente = st.selectbox(
                        "CLIENTE",
                        clientes_edicion,
                        index=clientes_edicion.index(cliente_actual) if cliente_actual in clientes_edicion else 0,
                    )
                    edit_agencia = st.selectbox(
                        "AGENCIA",
                        agencias_edicion,
                        index=agencias_edicion.index(agencia_actual) if agencia_actual in agencias_edicion else 0,
                    )
                    edit_fecha = st.date_input(
                        "FECHA DE SALIDA", value=fecha_edicion, format="DD/MM/YYYY"
                    )
                    edit_cajas = st.number_input(
                        "CANTIDAD DE CAJAS SOLICITADAS",
                        min_value=0,
                        step=1,
                        value=entero_seguro(registro_editar.get("CANTIDAD DE CAJAS SOLICITADAS", 0)),
                    )
                with e2:
                    edit_wo = st.text_input(
                        "NRO. WORKORDER", value=str(registro_editar.get("NRO. WORKORDER", ""))
                    )
                    edit_reutilizadas = st.number_input(
                        "CAJAS REUTILIZADAS",
                        min_value=0,
                        step=1,
                        value=entero_seguro(registro_editar.get("CAJAS REUTILIZADAS", 0)),
                    )
                    edit_cintillos = st.number_input(
                        "CANTIDAD DE CINTILLOS",
                        min_value=0,
                        step=1,
                        value=entero_seguro(registro_editar.get("CANTIDAD DE CINTILLOS", 0)),
                    )
                    edit_observaciones = st.text_area(
                        "OBSERVACIONES", value=str(registro_editar.get("OBSERVACIONES", ""))
                    )

                bx, by = st.columns([4, 1])
                with bx:
                    guardar_edicion_salida = st.form_submit_button(
                        "💾 ACTUALIZAR REGISTRO", use_container_width=True
                    )
                with by:
                    cancelar_edicion_salida = st.form_submit_button(
                        "✕ CANCELAR", type="primary", use_container_width=True
                    )

            if cancelar_edicion_salida:
                # Cierra el formulario de edición y deselecciona el registro en la tabla.
                st.session_state.salida_id_editar_inline = None
                st.session_state["caja_tabla_version"] = st.session_state.get("caja_tabla_version", 0) + 1
                st.rerun()

            if guardar_edicion_salida:
                nuevo_consumo_cajas_nuevas = int(edit_cajas) - int(edit_reutilizadas)
                if not edit_solicitante.strip() or not edit_agencia:
                    st.error("Completa SOLICITANTE y AGENCIA antes de actualizar.")
                elif int(edit_cajas) <= 0 and int(edit_cintillos) <= 0:
                    st.error("Debes registrar al menos una cantidad mayor a cero de CAJAS o CINTILLOS.")
                elif nuevo_consumo_cajas_nuevas < 0:
                    st.error("Las cajas reutilizadas no pueden ser mayores que las cajas solicitadas.")
                else:
                    viejo_cajas = entero_seguro(registro_editar.get("CANTIDAD DE CAJAS SOLICITADAS", 0))
                    viejo_reutilizadas = entero_seguro(registro_editar.get("CAJAS REUTILIZADAS", 0))
                    viejo_cintillos = entero_seguro(registro_editar.get("CANTIDAD DE CINTILLOS", 0))
                    viejo_consumo_nuevas = viejo_cajas - viejo_reutilizadas

                    # Ajuste necesario para que el stock refleje exactamente el registro editado.
                    ajuste_nuevas = viejo_consumo_nuevas - nuevo_consumo_cajas_nuevas
                    ajuste_reutilizadas = viejo_reutilizadas - int(edit_reutilizadas)
                    ajuste_cintillos = viejo_cintillos - int(edit_cintillos)

                    stock_actual_edicion = obtener_stock_actual()
                    falta_nuevas = max(0, -ajuste_nuevas - stock_actual_edicion["CAJAS NUEVAS"])
                    falta_reutilizadas = max(0, -ajuste_reutilizadas - stock_actual_edicion["CAJAS REUTILIZADAS"])
                    falta_cintillos = max(0, -ajuste_cintillos - stock_actual_edicion["CINTILLOS"])

                    if falta_nuevas > 0 or falta_reutilizadas > 0 or falta_cintillos > 0:
                        st.error("❌ No se puede actualizar porque el cambio dejaría el stock en negativo.")
                    else:
                        datos_editados = {
                            "SOLICITANTE": edit_solicitante.strip(),
                            "CLIENTE": edit_cliente,
                            "AGENCIA": edit_agencia.strip(),
                            "FECHA DE SALIDA": edit_fecha.isoformat(),
                            "CANTIDAD DE CAJAS SOLICITADAS": int(edit_cajas),
                            "NRO. WORKORDER": edit_wo.strip(),
                            "CAJAS REUTILIZADAS": int(edit_reutilizadas),
                            "CANTIDAD DE CINTILLOS": int(edit_cintillos),
                            "OBSERVACIONES": edit_observaciones.strip(),
                        }
                        try:
                            actualizar_salida_cajas_supabase(datos_editados, int(salida_id_editar))
                            if ajuste_nuevas != 0 or ajuste_reutilizadas != 0 or ajuste_cintillos != 0:
                                guardar_movimiento_stock(
                                    edit_fecha,
                                    "AJUSTE",
                                    ajuste_nuevas,
                                    ajuste_reutilizadas,
                                    ajuste_cintillos,
                                    "",
                                    f"Ajuste por edición de salida ID {int(salida_id_editar)} · WO: {edit_wo.strip()}",
                                )
                            st.session_state.cajas = cargar_cajas_supabase()
                            st.session_state.movimientos_stock = cargar_movimientos_stock_supabase()
                            st.session_state.salida_id_editar_inline = None
                            st.session_state["caja_tabla_version"] = st.session_state.get("caja_tabla_version", 0) + 1
                            st.success("✅ Registro actualizado correctamente y stock ajustado.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"❌ No se pudo actualizar el registro: {e}")
    else:
        tabla_filtrada = cajas_df.copy()

    if tabla_filtrada.empty:
        if cajas_df.empty:
            st.info("No hay salidas registradas todavía.")
        else:
            st.info("No hay salidas que coincidan con los filtros seleccionados.")
    else:
        # ADMIN edita seleccionando la casilla ✏️ directamente al lado de cada registro,
        # igual que en la tabla principal de Solicitudes.
        if es_admin():
            tabla_cajas = tabla_filtrada.copy()
            tabla_cajas["✏️ EDITAR"] = False
            resultado_cajas = st.data_editor(
                tabla_cajas,
                use_container_width=True,
                hide_index=True,
                height=420,
                key=f"tabla_historial_salidas_{st.session_state.get('caja_tabla_version', 0)}",
                disabled=[c for c in tabla_cajas.columns if c not in ["✏️ EDITAR"]],
                column_config={
                    "_ID_CAJA_": None,
                    "FECHA DE SALIDA": st.column_config.DateColumn(
                        "FECHA DE SALIDA", format="DD/MM/YYYY", width=130
                    ),
                    "✏️ EDITAR": st.column_config.CheckboxColumn(
                        "EDIT.", default=False, width=60
                    ),
                },
            )

            editar_caja = resultado_cajas[resultado_cajas["✏️ EDITAR"] == True]
            if not editar_caja.empty:
                salida_seleccionada = int(editar_caja.iloc[0]["_ID_CAJA_"])
                if st.session_state.get("salida_id_editar_inline") != salida_seleccionada:
                    st.session_state.salida_id_editar_inline = salida_seleccionada
                    st.session_state["caja_tabla_version"] = st.session_state.get("caja_tabla_version", 0) + 1
                    st.rerun()
        else:
            tabla_cajas = tabla_filtrada.drop(columns=["_ID_CAJA_"], errors="ignore").copy()
            st.dataframe(
                tabla_cajas,
                use_container_width=True,
                hide_index=True,
                height=420,
                column_config={
                    "FECHA DE SALIDA": st.column_config.DateColumn(
                        "FECHA DE SALIDA", format="DD/MM/YYYY", width=130
                    )
                },
            )

    with st.expander("📊 VER HISTORIAL DE MOVIMIENTOS DE STOCK", expanded=False):
        mov = cargar_movimientos_stock_supabase()
        if mov.empty:
            st.info("No hay movimientos registrados todavía.")
        else:
            columnas_mov = [c for c in ["fecha", "tipo_movimiento", "cajas_nuevas", "cajas_reutilizadas", "cintillos", "proveedor", "observaciones", "usuario"] if c in mov.columns]
            st.dataframe(mov[columnas_mov], use_container_width=True, hide_index=True, height=350)

    st.markdown('<div class="footer"><strong>JC Control de Solicitudes — Almacén</strong><br>©JuanCarlosRamos - 2026 — Todos los derechos reservados</div>', unsafe_allow_html=True)
    st.stop()

# =========================================================
# PAGINA PROGRAMACION DE RUTAS
# =========================================================
if st.session_state.page == "rutas":
    st.subheader("🗓️ PROGRAMACIÓN DE RUTAS")
    st.caption("Registro y consulta de programación directamente desde Supabase PostgreSQL.")

    if st.button("⬅ VOLVER A SOLICITUDES"):
        liberar_bloqueo_edicion()
        st.session_state.page = "solicitudes"
        st.rerun()

    bloqueo = leer_bloqueo_edicion()
    if bloqueo and bloqueo.get("owner_id") != st.session_state.session_id:
        st.warning(f"🔒 {bloqueo.get('usuario', 'Otro usuario')} está realizando una operación de edición. Espera para registrar una ruta.")
    else:
        if not st.session_state.ruta_activo:
            if st.button("🔐 INICIAR REGISTRO DE RUTA"):
                ok, info = adquirir_bloqueo_edicion("nueva programación de ruta")
                if ok:
                    st.session_state.ruta_activo = True
                    st.session_state.ruta_form_version += 1
                    st.rerun()
                else:
                    st.warning(f"🔒 {info.get('usuario', 'Otro usuario')} está utilizando el sistema.")

    if st.session_state.ruta_activo:
        renovar_bloqueo_edicion()
        st.success("🔐 Registro de ruta activo. La edición está reservada para este usuario.")
        formulario_nueva_ruta()

    st.markdown("### 📋 Programaciones registradas")
    rutas = st.session_state.rutas
    if rutas is None or rutas.empty:
        st.info("No hay programaciones registradas todavía.")
    else:
        f1, f2 = st.columns([1, 2])
        with f1:
            regional_filtro = st.selectbox("Regional", ["TODOS"] + REGIONALES[1:], key="ruta_regional")
        with f2:
            buscar_ruta = st.text_input("Buscar", placeholder="Centro de acopio, agencia o regional...", key="buscar_ruta")

        datos = rutas.copy()
        if regional_filtro != "TODOS":
            datos = datos[datos["REGIONAL"].map(normalizar).eq(normalizar(regional_filtro))]
        if buscar_ruta:
            texto = normalizar(buscar_ruta)
            mask = datos.astype(str).apply(lambda col: col.map(lambda x: texto in normalizar(x))).any(axis=1)
            datos = datos[mask]

        st.caption(f"📋 {len(datos)} de {len(rutas)} programaciones encontradas")
        tabla_rutas = datos.drop(columns=["_ID_RUTA_"], errors="ignore").copy()

        # Formato de fechas para que la tabla sea más compacta y uniforme.
        for col in [
            "FECHA LIMITE DE INGRESO (SE1)",
            "FECHA LIMITE DE INGRESO (SR1)",
            "FECHA DE RECOJO",
        ]:
            if col in tabla_rutas.columns:
                tabla_rutas[col] = pd.to_datetime(
                    tabla_rutas[col], errors="coerce"
                )

        st.dataframe(
            tabla_rutas,
            use_container_width=True,
            hide_index=True,
            height=500,
            column_config={
                "REGIONAL": st.column_config.TextColumn(
                    "REGIONAL", width=120
                ),
                "CENTRO DE ACOPIO": st.column_config.TextColumn(
                    "CENTRO DE ACOPIO", width=190
                ),
                "AGENCIAS": st.column_config.TextColumn(
                    "AGENCIAS", width=250
                ),
                "FECHA LIMITE DE INGRESO (SE1)": st.column_config.DateColumn(
                    "FECHA LÍMITE SE1", format="DD/MM/YYYY", width=150
                ),
                "FECHA LIMITE DE INGRESO (SR1)": st.column_config.DateColumn(
                    "FECHA LÍMITE SR1", format="DD/MM/YYYY", width=150
                ),
                "FECHA DE RECOJO": st.column_config.DateColumn(
                    "FECHA DE RECOJO", format="DD/MM/YYYY", width=145
                ),
            },
        )
        _, col_descarga_ruta, _ = st.columns([1, 2, 1])
        with col_descarga_ruta:
            st.download_button(
                "📥 DESCARGAR PROGRAMACIÓN",
                data=excel_bytes(st.session_state.rows, datos),
                file_name=f"PROGRAMACION_RUTAS_{date.today().isoformat()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    st.markdown('<div class="footer"><strong>JC Control de Solicitudes — Almacén</strong><br>©JuanCarlosRamos - 2026 — Todos los derechos reservados</div>', unsafe_allow_html=True)
    st.stop()


# =========================================================
# PAGINA SOLICITUDES
# =========================================================
df = st.session_state.rows
st.markdown("### 📁 Base de datos compartida")
st.caption("🗄️ Supabase PostgreSQL · Datos compartidos entre todos los usuarios.")

bloqueo_actual = leer_bloqueo_edicion()
if bloqueo_actual and bloqueo_actual.get("owner_id") != st.session_state.session_id:
    st.warning(f"🔒 BASE DE DATOS EN USO — {bloqueo_actual.get('usuario', 'otro usuario')} está registrando o editando. Puedes consultar, pero espera para modificar.")

# =========================================================
# NUEVA SOLICITUD
# =========================================================
if st.session_state.editing is None and not st.session_state.registro_activo:
    if st.button("🔐 INICIAR NUEVA SOLICITUD"):
        ok, info = adquirir_bloqueo_edicion("nueva solicitud")
        if ok:
            st.session_state.registro_activo = True
            st.session_state.form_version += 1
            st.rerun()
        else:
            st.warning(f"🔒 {info.get('usuario', 'Otro usuario')} está utilizando el sistema.")

# =========================================================
# FORMULARIO SOLICITUD
# =========================================================
if st.session_state.registro_activo or st.session_state.editing is not None:
    renovar_bloqueo_edicion()
    editing = st.session_state.editing
    if editing is not None and editing in df.index:
        r = df.loc[editing]
        valores = {"cliente":r["CLIENTE"],"numero":r["NRO SOLICITUD - WO"],"tipo":r["TIPO DE SOLICITUD"],"centro_costo":r.get("CENTRO DE COSTO",""),"prioridad":r["PRIORIDAD"],"cantidad_items":r["CANT - ITEMS"],"estado":r["ESTADO DE SOLICITUD"],"direccion":r["DIRECCIÓN"],"fecha":r["FECHA DE INGRESO"]}
        titulo_form = "✏️ Editar solicitud"
    else:
        valores = {"cliente":CLIENTES[0],"numero":"","tipo":TIPOS[0],"centro_costo":"","prioridad":PRIORIDADES[0],"cantidad_items":0,"estado":ESTADOS[0],"direccion":DIRECCIONES[0],"fecha":str(date.today())}
        titulo_form = "📥 Registrar nueva solicitud"

    with st.form(f"solicitud_form_{st.session_state.form_version}"):
        st.markdown(f"**{titulo_form}**")
        a,b = st.columns(2)
        with a:
            cliente = st.selectbox("Cliente", CLIENTES, index=CLIENTES.index(valores["cliente"]) if valores["cliente"] in CLIENTES else 0)
            numero = st.text_input("Nro. Solicitud / WO", value=str(valores["numero"] or ""), placeholder="Ejemplo: SE2-26-17658")
            tipo = st.selectbox("Tipo de Solicitud", TIPOS, index=TIPOS.index(valores["tipo"]) if valores["tipo"] in TIPOS else 0)
            centro_actual = str(valores.get("centro_costo", "") or "").strip()
            opciones_centro = CENTROS_COSTO.copy()
            if centro_actual and centro_actual not in opciones_centro:
                opciones_centro.append(centro_actual)
            centro_costo = st.selectbox(
                "Centro de Costo",
                opciones_centro,
                index=opciones_centro.index(centro_actual) if centro_actual in opciones_centro else 0,
            )
            prioridad = st.selectbox("Prioridad", PRIORIDADES, index=PRIORIDADES.index(valores["prioridad"]) if valores["prioridad"] in PRIORIDADES else 0)
        with b:
            try: cantidad_default = max(0, int(float(str(valores.get("cantidad_items",0)).replace(",","."))))
            except Exception: cantidad_default = 0
            cantidad_items = st.number_input("Cantidad de ítems", min_value=0, step=1, value=cantidad_default, format="%d")
            estado = st.selectbox("Estado de Solicitud", ESTADOS, index=ESTADOS.index(valores["estado"]) if valores["estado"] in ESTADOS else 0)
            direccion = st.selectbox("Dirección", DIRECCIONES, index=DIRECCIONES.index(valores["direccion"]) if valores["direccion"] in DIRECCIONES else 0)
            fecha_default = convertir_fecha(valores.get("fecha")) or fecha_local_hoy()
            fecha = st.date_input("📅 Fecha de ingreso", value=fecha_default, format="DD/MM/YYYY")
        x,y = st.columns([4,1])
        with x: guardar = st.form_submit_button("💾 ACTUALIZAR REGISTRO" if editing is not None else "📥 AGREGAR REGISTRO", use_container_width=True)
        with y: cancelar = st.form_submit_button("✕ CANCELAR", type="primary", use_container_width=True)

    if cancelar:
        liberar_bloqueo_edicion()
        st.session_state.editing = None
        st.session_state.editing_key = ""
        st.session_state.form_version += 1
        st.rerun()

    if guardar:
        if editing is not None and not puede_modificar():
            st.error("⛔ Tu usuario no tiene permiso para editar solicitudes.")
            st.stop()
        if editing is None and not puede_modificar():
            st.error("⛔ Tu usuario no tiene permiso para registrar solicitudes.")
            st.stop()
        numero_norm = normalizar(numero)
        if not numero_norm or cliente == CLIENTES[0] or tipo == TIPOS[0] or prioridad == PRIORIDADES[0] or estado == ESTADOS[0] or direccion == DIRECCIONES[0]:
            st.error("Completa todos los campos obligatorios antes de guardar.")
        else:
            try:
                latest = cargar_solicitudes_supabase()
                nuevo = {"CLIENTE":cliente,"NRO SOLICITUD - WO":numero.strip(),"TIPO DE SOLICITUD":tipo,"CENTRO DE COSTO":centro_costo.strip(),"PRIORIDAD":prioridad,"CANT - ITEMS":int(cantidad_items),"ESTADO DE SOLICITUD":estado,"DIRECCIÓN":direccion,"FECHA DE INGRESO":fecha.isoformat()}
                repetidos = latest[latest["NRO SOLICITUD - WO"].map(normalizar).eq(numero_norm)]
                if editing is not None:
                    registro_id = int(df.loc[editing,"_ID_"])
                    if not repetidos.empty and int(repetidos.iloc[0]["_ID_"]) != registro_id:
                        st.error("Ya existe otra solicitud con ese Nro. Solicitud / WO.")
                    else:
                        guardar_solicitud_supabase(nuevo, registro_id)
                        st.session_state.rows = cargar_solicitudes_supabase()
                        liberar_bloqueo_edicion()
                        st.session_state.editing = None
                        st.session_state.editing_key = ""
                        st.success("✅ Registro actualizado correctamente en Supabase.")
                        st.rerun()
                else:
                    if not repetidos.empty:
                        st.error("Ya existe una solicitud con ese Nro. Solicitud / WO.")
                    else:
                        guardar_solicitud_supabase(nuevo)
                        st.session_state.rows = cargar_solicitudes_supabase()
                        liberar_bloqueo_edicion()
                        st.success("✅ Nueva solicitud agregada correctamente en Supabase.")
                        st.rerun()
            except Exception as e:
                liberar_bloqueo_edicion()
                st.error(f"❌ No se pudo guardar en Supabase: {e}")

# =========================================================
# FILTROS
# =========================================================
st.markdown("### Resultados")
buscar_col, cliente_col, estado_col, fecha_col = st.columns([3,1.4,1.4,1.2])
with buscar_col:
    buscar = st.text_input("Buscar solicitud...", placeholder="Cliente, WO, dirección, tipo...", key="buscar_solicitud_live")
with cliente_col:
    filtro_cliente = st.selectbox("Cliente", ["TODOS"] + CLIENTES[1:], key="filtro_cliente")
with estado_col:
    filtro_estado = st.selectbox("Estado", ["TODOS"] + ESTADOS[1:], key="filtro_estado_solicitud")
with fecha_col:
    filtro_fecha = st.date_input("📅 Fecha", value=None, key="filtro_fecha_solicitud")

vista = df.copy()
if buscar:
    texto = normalizar(buscar)
    vista = vista[vista.astype(str).apply(lambda col: col.map(lambda x: texto in normalizar(x))).any(axis=1)]
if filtro_cliente != "TODOS":
    vista = vista[vista["CLIENTE"].map(normalizar).eq(normalizar(filtro_cliente))]
if filtro_estado != "TODOS":
    vista = vista[vista["ESTADO DE SOLICITUD"].map(normalizar).eq(normalizar(filtro_estado))]
if filtro_fecha:
    fechas = vista["FECHA DE INGRESO"].apply(convertir_fecha)
    vista = vista[fechas.apply(lambda f: f == filtro_fecha)]

# =========================================================
# ALERTAS DE ATENCIÓN POR ANTIGÜEDAD
# =========================================================
# Las solicitudes ENVIADAS, ENTREGADAS o ANULADAS no generan alertas.
# 1 día de antigüedad  -> advertencia
# 2 o más días          -> atención urgente
estados_sin_alerta = {"ENVIADO", "ENTREGADO", "ANULADO"}

alertas = df.copy()

# Convertir cada fecha con la misma lógica usada por el sistema:
# - ISO de Supabase: YYYY-MM-DD
# - Fecha mostrada/guardada como DD/MM/YYYY
# Esto evita que las solicitudes antiguas se interpreten con mes y día invertidos.
alertas["_FECHA_ALERTA"] = alertas["FECHA DE INGRESO"].apply(convertir_fecha)

hoy_app = fecha_local_hoy()
alertas["_DIAS_ALERTA"] = alertas["_FECHA_ALERTA"].apply(
    lambda f: (hoy_app - f).days if f is not None else None
)

alertas = alertas[
    (~alertas["ESTADO DE SOLICITUD"].astype(str).str.strip().str.upper().isin(estados_sin_alerta))
    & alertas["_FECHA_ALERTA"].notna()
    & (alertas["_DIAS_ALERTA"] >= 1)
]

if not alertas.empty:
    urgentes = alertas[alertas["_DIAS_ALERTA"] >= 2]
    advertencias = alertas[alertas["_DIAS_ALERTA"] == 1]

    st.markdown("### 🔔 Alertas de atención")
    met1, met2, met3 = st.columns(3)
    with met1:
        st.metric("🚨 Atención urgente (2+ días)", len(urgentes))
    with met2:
        st.metric("⚠️ Atención (1 día)", len(advertencias))
    with met3:
        st.metric("📌 Total pendientes", len(alertas))

    # Botón desplegable para no ocupar espacio con las tablas de alertas.
    with st.expander("🔽 Ver tabla de alertas de atención", expanded=False):
        if not urgentes.empty:
            st.error(
                f"🚨 Hay **{len(urgentes)} solicitud(es)** con 2 o más días de antigüedad que requieren atención urgente."
            )
            detalle_urg = urgentes[["NRO SOLICITUD - WO", "CLIENTE", "FECHA DE INGRESO", "ESTADO DE SOLICITUD", "_DIAS_ALERTA"]].copy()
            detalle_urg["FECHA DE INGRESO"] = detalle_urg["FECHA DE INGRESO"].apply(
                lambda f: convertir_fecha(f).strftime("%d/%m/%Y") if convertir_fecha(f) else ""
            )
            detalle_urg["ATENCIÓN"] = detalle_urg["_DIAS_ALERTA"].map(lambda x: f"🚨 {int(x)} días")
            detalle_urg = detalle_urg.drop(columns=["_DIAS_ALERTA"])
            st.dataframe(detalle_urg, use_container_width=True, hide_index=True)

        if not advertencias.empty:
            st.warning(
                f"⚠️ Hay **{len(advertencias)} solicitud(es)** que llevan 1 día desde su registro."
            )
            detalle_adv = advertencias[["NRO SOLICITUD - WO", "CLIENTE", "FECHA DE INGRESO", "ESTADO DE SOLICITUD"]].copy()
            detalle_adv["FECHA DE INGRESO"] = detalle_adv["FECHA DE INGRESO"].apply(
                lambda f: convertir_fecha(f).strftime("%d/%m/%Y") if convertir_fecha(f) else ""
            )
            detalle_adv["ATENCIÓN"] = "⚠️ 1 día"
            st.dataframe(detalle_adv, use_container_width=True, hide_index=True)
else:
    st.success("✅ No hay solicitudes pendientes con más de 1 día de antigüedad que requieran atención.")

st.caption(f"{len(vista)} de {len(df)} registros")

if len(vista):
    tabla = vista[COLS].copy()
    tabla["_INDICE_REAL_"] = vista.index
    tabla["✏️ EDITAR"] = False
    tabla["🗑️ ELIMINAR"] = False

    # =====================================================
    # COLORES VISUALES PARA ESTADO Y PRIORIDAD
    # =====================================================
    # Se muestran como indicadores de color dentro de la
    # tabla sin modificar los valores originales de Supabase.
    estado_colores = {
        "POR EXTRAER": "🟡 POR EXTRAER",
        "POR REASIGNAR": "🟣 POR REASIGNAR",
        "POR ENVIAR": "🟠 POR ENVIAR",
        "ENTREGADO": "🟢 ENTREGADO",
        "ENVIADO": "🔵 ENVIADO",
        "ANULADO": "🔴 ANULADO",
        "POR ETIQUETAR": "🟡 POR ETIQUETAR",
    }

    prioridad_colores = {
        "RUSH": "🔴 RUSH",
        "TURNO SIGUIENTE": "🟠 TURNO SIGUIENTE",
        "NORMAL": "🟢 NORMAL",
    }

    tabla["ESTADO DE SOLICITUD"] = (
        tabla["ESTADO DE SOLICITUD"]
        .map(lambda x: estado_colores.get(str(x).strip(), str(x)))
    )

    tabla["PRIORIDAD"] = (
        tabla["PRIORIDAD"]
        .map(lambda x: prioridad_colores.get(str(x).strip(), str(x)))
    )
    resultado = st.data_editor(
        tabla,
        use_container_width=True,
        hide_index=True,
        height=470,
        key=f"tabla_solicitudes_{st.session_state.tabla_version}",
        disabled=COLS + ["_INDICE_REAL_"],
        column_config={
            "_INDICE_REAL_": None,

            # Anchos fijos para que todas las columnas entren de forma
            # equilibrada en la pantalla.
            "CLIENTE": st.column_config.TextColumn("CLIENTE", width=130),
            "NRO SOLICITUD - WO": st.column_config.TextColumn("N° SOLICITUD / WO", width=160),
            "TIPO DE SOLICITUD": st.column_config.TextColumn("TIPO", width=125),
            "CENTRO DE COSTO": st.column_config.TextColumn("CENTRO DE COSTO", width=125),
            "PRIORIDAD": st.column_config.TextColumn("PRIORIDAD", width=160),
            "CANT - ITEMS": st.column_config.NumberColumn(
                "ÍTEMS", min_value=0, step=1, format="%d", width=70
            ),
            "ESTADO DE SOLICITUD": st.column_config.TextColumn("ESTADO", width=130),
            "DIRECCIÓN": st.column_config.TextColumn("DIRECCIÓN", width=150),
            "FECHA DE INGRESO": st.column_config.DateColumn(
                "FECHA", format="DD/MM/YYYY", width=105
            ),
            "✏️ EDITAR": st.column_config.CheckboxColumn("EDIT.", default=False, width=60),
            "🗑️ ELIMINAR": st.column_config.CheckboxColumn(
                "ELIM.", default=False, width=60, disabled=not es_admin()
            ),
        },
    )

    editar = resultado[resultado["✏️ EDITAR"] == True]
    if not editar.empty:
        indice_real = editar.iloc[0]["_INDICE_REAL_"]
        ok, info = adquirir_bloqueo_edicion("edición de solicitud")
        if ok:
            st.session_state.editing = int(indice_real)
            st.session_state.editing_key = normalizar(df.loc[int(indice_real), "NRO SOLICITUD - WO"])
            st.session_state.form_version += 1
            st.session_state.tabla_version += 1
            st.rerun()
        else:
            st.warning(f"🔒 {info.get('usuario', 'Otro usuario')} está editando.")

    eliminar = resultado[resultado["🗑️ ELIMINAR"] == True]
    if not es_admin():
        eliminar = resultado.iloc[0:0]
    if not eliminar.empty:
        ok, info = adquirir_bloqueo_edicion("eliminación de solicitud")
        if not ok:
            st.warning(f"🔒 {info.get('usuario', 'Otro usuario')} está utilizando el sistema.")
        else:
            ids = [int(df.loc[int(i), "_ID_"]) for i in eliminar["_INDICE_REAL_"].tolist()]
            try:
                eliminar_solicitudes_supabase(ids)
                st.session_state.rows = cargar_solicitudes_supabase()
                liberar_bloqueo_edicion()
                st.success("✅ Registro(s) eliminado(s) correctamente de Supabase.")
                st.session_state.tabla_version += 1
                st.rerun()
            except Exception as e:
                liberar_bloqueo_edicion()
                st.error(f"❌ No se pudo eliminar: {e}")
else:
    st.info("No hay registros que coincidan con los filtros.")

# =========================================================
# ARCHIVO MENSUAL Y LIMPIEZA DE BASE DE DATOS
# Solo ADMIN puede eliminar registros históricos.
# Primero se descarga el mes y luego se solicita confirmación
# explícita antes de borrar esos registros de Supabase.
# =========================================================
if es_admin():
    with st.expander("📦 ARCHIVAR Y LIMPIAR SOLICITUDES POR MES", expanded=False):
        st.warning(
            "⚠️ Esta opción descarga los registros de un mes y permite eliminarlos de Supabase. "
            "Se recomienda conservar el archivo Excel como respaldo antes de eliminar."
        )

        hoy = fecha_local_hoy()
        anos_disponibles = sorted(
            {d.year for d in pd.to_datetime(df["FECHA DE INGRESO"], errors="coerce").dropna()},
            reverse=True,
        )
        if hoy.year not in anos_disponibles:
            anos_disponibles.insert(0, hoy.year)
        meses = {
            1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
            5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
            9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"
        }

        ac1, ac2 = st.columns(2)
        with ac1:
            anio_archivo = st.selectbox(
                "Año", anos_disponibles, key="anio_archivo_mensual"
            )
        with ac2:
            mes_archivo = st.selectbox(
                "Mes", list(meses.keys()),
                format_func=lambda m: meses[m],
                key="mes_archivo_mensual"
            )

        import calendar
        ultimo_dia = calendar.monthrange(anio_archivo, mes_archivo)[1]
        fecha_inicio_archivo = date(anio_archivo, mes_archivo, 1)
        fecha_fin_archivo = date(anio_archivo, mes_archivo, ultimo_dia)
        fecha_fin_exclusiva = date(anio_archivo + (mes_archivo == 12), 1 if mes_archivo == 12 else mes_archivo + 1, 1)

        # La confirmación de borrado pertenece exclusivamente al mes seleccionado.
        # Si el usuario cambia de año/mes, se limpia para evitar borrar otro mes por error.
        _mes_seleccionado = (anio_archivo, mes_archivo)
        _mes_previo = st.session_state.get("_mes_confirmacion_borrado")
        if _mes_previo != _mes_seleccionado:
            st.session_state.pop("confirmar_borrado_mensual", None)
            st.session_state["_mes_confirmacion_borrado"] = _mes_seleccionado

        # Después de un borrado exitoso se solicita nuevamente la confirmación.
        if st.session_state.pop("_reset_confirmar_borrado_mensual", False):
            st.session_state.pop("confirmar_borrado_mensual", None)

        try:
            registros_mes = cargar_solicitudes_mes_supabase(
                fecha_inicio_archivo, fecha_fin_exclusiva
            )
            cantidad_mes = len(registros_mes)

            st.info(
                f"📅 {meses[mes_archivo]} {anio_archivo}: **{cantidad_mes} registro(s)** encontrados "
                f"({fecha_inicio_archivo.strftime('%d/%m/%Y')} al {fecha_fin_archivo.strftime('%d/%m/%Y')})."
            )

            if cantidad_mes > 0:
                archivo_mes = excel_bytes_mensual(registros_mes[COLS])
                st.info(
                    "📊 El Excel incluirá el detalle completo con **NRO SOLICITUD - WO**, "
                    "además de resúmenes por estado, prioridad, SE1/SE2/SE3/SR1/SR2 y un resumen general."
                )

                _, col_descarga_mes, _ = st.columns([1, 2, 1])
                with col_descarga_mes:
                    st.download_button(
                        "📥 DESCARGAR ESTE MES EN EXCEL",
                        data=archivo_mes,
                        file_name=f"SOLICITUDES_{anio_archivo}_{mes_archivo:02d}_{meses[mes_archivo]}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="descargar_archivo_mensual",
                    )

                confirmar_borrado = st.checkbox(
                    "Confirmo que ya descargué y guardé el archivo de este mes y deseo eliminar estos registros de Supabase.",
                    key="confirmar_borrado_mensual",
                )

                if st.button(
                    f"🗑️ ELIMINAR {cantidad_mes} REGISTRO(S) DE {meses[mes_archivo]} {anio_archivo}",
                    type="secondary",
                    use_container_width=True,
                    disabled=not confirmar_borrado,
                    key="eliminar_mes_supabase",
                ):
                    ids_mes = [int(x) for x in registros_mes["_ID_"].tolist() if pd.notna(x)]
                    try:
                        eliminar_solicitudes_supabase(ids_mes)
                        st.session_state.rows = cargar_solicitudes_supabase()
                        st.session_state["_reset_confirmar_borrado_mensual"] = True
                        st.success(
                            f"✅ Se eliminaron {len(ids_mes)} registro(s) de {meses[mes_archivo]} {anio_archivo} de Supabase."
                        )
                        time.sleep(1)
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ No se pudo completar la limpieza mensual: {e}")
            else:
                st.success("✅ No existen registros para el mes seleccionado.")
        except Exception as e:
            st.error(f"❌ No se pudo consultar el mes seleccionado: {e}")

# =========================================================
# DESCARGA
# =========================================================
_, col_descarga_filtrados, _ = st.columns([1, 2, 1])
with col_descarga_filtrados:
    st.download_button(
        "📥 DESCARGAR FILTRADOS",
        data=excel_bytes(vista[COLS], None),
        file_name=f"SOLICITUDES_FILTRADAS_{date.today().isoformat()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

st.markdown('<div class="footer"><strong>JC Control de Solicitudes — Almacén</strong><br>©JuanCarlosRamos - 2026 — Todos los derechos reservados</div>', unsafe_allow_html=True)
