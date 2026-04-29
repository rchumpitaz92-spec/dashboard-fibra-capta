"""
Dashboard Revisiones Fibra Capta - Streamlit App
=================================================
Solo el admin puede subir los archivos Excel.
Todos los demás solo ven el dashboard.

ARCHIVOS QUE MANEJA:
1. REVISIONES_FIBRA_CAPTA_*.xlsx → pestañas Front, BO, Aprobado vs KO
2. CORTE_CAPTA_*.xlsx (hoja "Volcado Bo") → pestañas Corte x Hora, Efectividad
"""

import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import json
import re
from pathlib import Path
from datetime import datetime

# ============== CONFIGURACIÓN ==============
st.set_page_config(
    page_title="Dashboard Revisiones Fibra Capta",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

DATA_FILE = Path("data_actual.json")
META_FILE = Path("data_meta.json")
CORTE_DATA_FILE = Path("corte_data.json")
CORTE_META_FILE = Path("corte_meta.json")
DASHBOARD_TEMPLATE = Path("dashboard_template.html")


# ============== REVISIONES ==============
REQUIRED_COLUMNS = [
    'COORDINADOR', 'AGENTE FRONT', 'AGENTE BO', 'ESTADO', 'SEMANA', 'DIA',
    'MOTIVO_CANCELACION', 'MOTIVO_CANCELACION3', 'Motivo Bo', 'Sub Motivo Bo',
    'SLA_AGENDA', 'Solicitud', 'FAMILIA', 'SLA_INSTALACION'
]


def validate_excel(uploaded_file) -> dict:
    result = {'ok': True, 'errors': [], 'warnings': [], 'info': {}}
    try:
        df = pd.read_excel(uploaded_file, sheet_name=0, header=0)
    except Exception as e:
        result['ok'] = False
        result['errors'].append(f"No se pudo leer el archivo: {e}")
        return result

    missing_cols = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing_cols:
        result['ok'] = False
        result['errors'].append(
            f"Faltan columnas obligatorias: **{', '.join(missing_cols)}**. "
            "Verifica que estés subiendo el archivo correcto (REVISIONES, no CORTE)."
        )
        return result

    if len(df) == 0:
        result['ok'] = False
        result['errors'].append("El archivo está vacío.")
        return result

    valid = df[df['ESTADO'].isin(['TERMINADO','CANCELADO','CAIDA 1RA LINEA','PENDIENTE'])]
    valid = valid[valid['COORDINADOR'].notna()]
    new_count = len(valid)

    if new_count == 0:
        result['ok'] = False
        result['errors'].append("El archivo no tiene casos válidos.")
        return result

    existing_meta = load_meta()
    if existing_meta:
        old_count = existing_meta.get('total_records', 0)
        diff = new_count - old_count
        result['info']['old_count'] = old_count
        result['info']['diff'] = diff
        if diff < 0:
            result['warnings'].append(
                f"⚠️ El archivo tiene **{abs(diff)} casos MENOS** que el actual ({new_count} vs {old_count})."
            )
        elif diff == 0:
            result['warnings'].append(f"ℹ️ Mismo número de casos ({new_count}).")

    result['info']['casos_validos'] = new_count
    result['info']['coordinadores'] = sorted(valid['COORDINADOR'].dropna().unique().tolist())
    result['info']['estados'] = valid['ESTADO'].value_counts().to_dict()
    return result


def process_excel(uploaded_file) -> tuple:
    df = pd.read_excel(uploaded_file, sheet_name=0, header=0)
    valid = df[df['ESTADO'].isin(['TERMINADO','CANCELADO','CAIDA 1RA LINEA','PENDIENTE'])].copy()
    valid = valid[valid['COORDINADOR'].notna()]

    out = pd.DataFrame({
        'coord': valid['COORDINADOR'].fillna(''),
        'agente': valid['AGENTE FRONT'].fillna(''),
        'agente_bo': valid['AGENTE BO'].fillna(''),
        'estado': valid['ESTADO'].fillna(''),
        'semana': valid['SEMANA'].fillna(''),
        'dia': valid['DIA'].fillna('').astype(str),
        'motivo': valid['MOTIVO_CANCELACION'].fillna(''),
        'submotivo': valid['MOTIVO_CANCELACION3'].fillna(''),
        'caida_motivo': valid.apply(lambda r: r['Motivo Bo'] if r['ESTADO']=='CAIDA 1RA LINEA' else '', axis=1).fillna(''),
        'caida_submotivo': valid.apply(lambda r: r['Sub Motivo Bo'] if r['ESTADO']=='CAIDA 1RA LINEA' else '', axis=1).fillna(''),
        'modalidad': valid['SLA_AGENDA'].fillna(''),
        'solicitud': valid['Solicitud'].fillna(''),
        'familia': valid['FAMILIA'].fillna(''),
        'sla_inst': valid['SLA_INSTALACION'].fillna(''),
    })
    for c in out.columns:
        out[c] = out[c].astype(str).str.strip().replace({'nan':'','None':'','NaT':'','0':''})
    return out.to_dict('records'), len(out)


def load_data():
    if DATA_FILE.exists():
        with open(DATA_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return None


def load_meta():
    if META_FILE.exists():
        with open(META_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return None


def save_data(records, filename):
    with open(DATA_FILE, 'w', encoding='utf-8') as f:
        json.dump(records, f, ensure_ascii=False, separators=(',',':'))
    meta = {
        'last_update': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'total_records': len(records),
        'filename': filename
    }
    with open(META_FILE, 'w', encoding='utf-8') as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)


# ============== CORTE ==============
CORTE_REQUIRED = ['Resultado Bo', 'HORA BO', 'AGENTE BO', 'COORDINADOR',
                  'Solicitud', 'Motivo Bo', 'Sub Motivo Bo', 'SLA_AGENDA',
                  'SEMANA', 'DIA FRONT']


def validate_corte(uploaded_file) -> dict:
    result = {'ok': True, 'errors': [], 'warnings': [], 'info': {}}
    try:
        xl = pd.ExcelFile(uploaded_file)
    except Exception as e:
        result['ok'] = False
        result['errors'].append(f"No se pudo abrir el archivo: {e}")
        return result

    sheet_match = next((s for s in xl.sheet_names
                        if s.lower().replace(' ', '') == 'volcadobo'), None)
    if not sheet_match:
        result['ok'] = False
        result['errors'].append(
            f"No se encontró la hoja **'Volcado Bo'**. "
            f"Hojas en el archivo: {', '.join(xl.sheet_names[:5])}..."
        )
        return result

    try:
        df = pd.read_excel(uploaded_file, sheet_name=sheet_match, header=0)
    except Exception as e:
        result['ok'] = False
        result['errors'].append(f"Error leyendo Volcado Bo: {e}")
        return result

    missing = [c for c in CORTE_REQUIRED if c not in df.columns]
    if missing:
        result['ok'] = False
        result['errors'].append(f"Faltan columnas en Volcado Bo: **{', '.join(missing)}**")
        return result

    valid = df[df['Resultado Bo'].isin(['APROBADO','CAIDA','OBSERVADO','PENDIENTE'])]
    new_count = len(valid)
    if new_count == 0:
        result['ok'] = False
        result['errors'].append("Ningún caso con Resultado Bo válido.")
        return result

    existing_meta = load_corte_meta()
    if existing_meta:
        old_count = existing_meta.get('total_records', 0)
        diff = new_count - old_count
        result['info']['old_count'] = old_count
        result['info']['diff'] = diff
        if diff < 0:
            result['warnings'].append(
                f"⚠️ Tiene **{abs(diff)} casos MENOS** que el actual ({new_count} vs {old_count})."
            )
        elif diff == 0:
            result['warnings'].append(f"ℹ️ Mismo número de casos ({new_count}).")

    result['info']['new_count'] = new_count
    result['info']['sheet'] = sheet_match
    result['info']['resultados'] = valid['Resultado Bo'].value_counts().to_dict()
    result['info']['agentes_bo'] = len(valid['AGENTE BO'].dropna().unique())
    return result


def process_corte(uploaded_file) -> tuple:
    xl = pd.ExcelFile(uploaded_file)
    sheet_match = next((s for s in xl.sheet_names
                        if s.lower().replace(' ', '') == 'volcadobo'), None)
    df = pd.read_excel(uploaded_file, sheet_name=sheet_match, header=0)
    valid = df[df['Resultado Bo'].isin(['APROBADO','CAIDA','OBSERVADO','PENDIENTE'])].copy()

    n = len(valid)

    # Construir fecha_front en formato YYYY-MM-DD
    if 'Fecha Front' in valid.columns:
        fecha_front_str = pd.to_datetime(valid['Fecha Front'], errors='coerce').dt.strftime('%Y-%m-%d').fillna('')
    else:
        fecha_front_str = pd.Series(['']*n, index=valid.index)

    out = pd.DataFrame({
        'agente_bo': valid['AGENTE BO'].fillna(''),
        'coord': valid['COORDINADOR'].fillna(''),
        'resultado': valid['Resultado Bo'].fillna(''),
        'hora_bo': valid['HORA BO'].fillna(0).astype(int),
        'hora_front': (valid['HORA FRONT'] if 'HORA FRONT' in valid.columns else pd.Series([0]*n, index=valid.index)).fillna(0).astype(int),
        'dia_front': valid['DIA FRONT'].fillna(0).astype(int),
        'fecha_front': fecha_front_str,
        'semana': valid['SEMANA'].fillna(''),
        'solicitud': valid['Solicitud'].fillna(''),
        'anti': (valid['ANTI'] if 'ANTI' in valid.columns else pd.Series(['']*n, index=valid.index)).fillna(''),
        'motivo_ko': valid['Motivo Bo'].fillna(''),
        'submotivo_ko': valid['Sub Motivo Bo'].fillna(''),
        'sla_agenda': valid['SLA_AGENDA'].fillna(''),
        'rango_delivery': (valid['RANGO_DELIVERY'] if 'RANGO_DELIVERY' in valid.columns else pd.Series(['']*n, index=valid.index)).fillna(''),
        'tipo_llamada': (valid['Tipo De Llamada'] if 'Tipo De Llamada' in valid.columns else pd.Series(['']*n, index=valid.index)).fillna(''),
    })
    for c in ['agente_bo','coord','resultado','semana','solicitud','anti','motivo_ko','submotivo_ko','sla_agenda','tipo_llamada','rango_delivery','fecha_front']:
        out[c] = out[c].astype(str).str.strip().replace({'nan':'','None':''})
    return out.to_dict('records'), len(out)


def load_corte_data():
    if CORTE_DATA_FILE.exists():
        with open(CORTE_DATA_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return None


def load_corte_meta():
    if CORTE_META_FILE.exists():
        with open(CORTE_META_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return None


def save_corte_data(records, filename):
    with open(CORTE_DATA_FILE, 'w', encoding='utf-8') as f:
        json.dump(records, f, ensure_ascii=False, separators=(',',':'))
    meta = {
        'last_update': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'total_records': len(records),
        'filename': filename
    }
    with open(CORTE_META_FILE, 'w', encoding='utf-8') as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)


# ============== DASHBOARD ==============
def render_dashboard(records, corte_records, is_admin=False):
    if not DASHBOARD_TEMPLATE.exists():
        st.error("No se encuentra dashboard_template.html en el repositorio.")
        return None

    with open(DASHBOARD_TEMPLATE, 'r', encoding='utf-8') as f:
        template = f.read()

    data_json = json.dumps(records, ensure_ascii=False, separators=(',',':'))
    html = re.sub(r'const RAW = \[.*?\];', f'const RAW = {data_json};', template, count=1, flags=re.DOTALL)

    total = len(records)
    html = re.sub(r'\d+ casos', f'{total} casos', html)

    # Inyectar flag de admin (true/false en JS)
    is_admin_js = 'true' if is_admin else 'false'
    admin_inject = f'<script>window.IS_ADMIN = {is_admin_js};</script>\n'

    # Inyectar data del corte si existe
    inject_block = admin_inject
    if corte_records:
        corte_json = json.dumps(corte_records, ensure_ascii=False, separators=(',',':'))
        inject_block += f'<script>window.PRELOADED_CORTE_DATA = {corte_json};</script>\n'

    html = html.replace(
        '<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js',
        inject_block + '<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js',
        1
    )

    return html


# ============== AUTH ==============
def check_admin():
    if 'is_admin' not in st.session_state:
        st.session_state.is_admin = False
    return st.session_state.is_admin


def admin_login_form():
    with st.sidebar:
        st.markdown("### 🔐 Modo administrador")
        if not check_admin():
            password_input = st.text_input("Clave de admin", type="password", key="pwd_input")
            if st.button("Iniciar sesión"):
                try:
                    correct_pwd = st.secrets["admin_password"]
                except (KeyError, FileNotFoundError):
                    correct_pwd = "cambiar_esta_clave"
                if password_input == correct_pwd:
                    st.session_state.is_admin = True
                    st.rerun()
                else:
                    st.error("Clave incorrecta")
        else:
            st.success("✓ Sesión de admin activa")
            if st.button("Cerrar sesión"):
                st.session_state.is_admin = False
                st.rerun()


# ============== UI ==============
def admin_upload_section(label, key_prefix, uploader_key, validate_fn, process_fn, save_fn):
    """Renderiza una zona de subida (admin) reutilizable para Revisiones y Corte."""
    uploaded = st.file_uploader(
        f"Arrastra el Excel de {label}",
        type=['xlsx', 'xls'],
        key=uploader_key
    )

    if uploaded is None:
        return

    st.markdown(f"📁 **{uploaded.name}** ({uploaded.size/1024:.1f} KB)")
    with st.spinner("Validando..."):
        val = validate_fn(uploaded)

    if not val['ok']:
        st.error("❌ No se puede subir")
        for err in val['errors']:
            st.markdown(f"- {err}")
        return

    info = val['info']
    new_count = info.get('casos_validos', info.get('new_count', 0))

    c1, c2, c3 = st.columns(3)
    with c1:
        st.metric("Casos válidos", new_count)
    with c2:
        if info.get('old_count') is not None:
            diff = info.get('diff', 0)
            label_val = f"+{diff}" if diff > 0 else (f"{diff}" if diff < 0 else "Sin cambios")
            st.metric("vs actual", label_val, delta=f"{diff}" if diff != 0 else None)
        else:
            st.metric("Estado", "Primera carga")
    with c3:
        if 'coordinadores' in info:
            st.metric("Coordinadores", len(info['coordinadores']))
        elif 'agentes_bo' in info:
            st.metric("Agentes BO", info['agentes_bo'])

    # Distribución (estados o resultados)
    dist = info.get('estados') or info.get('resultados')
    if dist:
        cols = st.columns(len(dist))
        for i, (k, v) in enumerate(dist.items()):
            with cols[i]:
                st.metric(k, v)

    if 'sheet' in info:
        st.info(f"📑 Hoja detectada: **{info['sheet']}**")

    btn_label = f"🚀 Actualizar {label}"
    if val['warnings']:
        for w in val['warnings']:
            st.warning(w)
        confirm = st.checkbox("✅ He revisado las advertencias", key=f"{key_prefix}_conf")
        if confirm and st.button(btn_label, type="primary", key=f"{key_prefix}_btn_w"):
            try:
                with st.spinner("Actualizando..."):
                    records, total = process_fn(uploaded)
                    save_fn(records, uploaded.name)
                st.success(f"✓ {label} actualizado con **{total} casos**")
                st.balloons()
                st.rerun()
            except Exception as e:
                st.error(f"Error: {e}")
                st.exception(e)
    else:
        st.success("✓ Validación OK")
        if st.button(btn_label, type="primary", key=f"{key_prefix}_btn"):
            try:
                with st.spinner("Actualizando..."):
                    records, total = process_fn(uploaded)
                    save_fn(records, uploaded.name)
                st.success(f"✓ {label} actualizado con **{total} casos**")
                st.balloons()
                st.rerun()
            except Exception as e:
                st.error(f"Error: {e}")
                st.exception(e)


def main():
    admin_login_form()

    meta = load_meta()
    corte_meta = load_corte_meta()

    # ===== Header =====
    col1, col2 = st.columns([3, 1])
    with col1:
        st.markdown("# 📊 Dashboard Revisiones Fibra Capta")
        captions = []
        if meta:
            captions.append(f"📋 **Revisiones:** {meta['last_update']} · {meta['total_records']} casos")
        else:
            captions.append("📋 **Revisiones:** sin data")
        if corte_meta:
            captions.append(f"⏱ **Corte:** {corte_meta['last_update']} · {corte_meta['total_records']} casos")
        else:
            captions.append("⏱ **Corte:** sin data")
        st.caption(" · ".join(captions))
    with col2:
        if check_admin():
            st.info("👤 Eres admin")

    st.markdown("---")

    # ===== Panel admin =====
    if check_admin():
        st.markdown("### 📤 Panel del administrador")
        admin_tab1, admin_tab2 = st.tabs(["📋 Subir REVISIONES", "⏱ Subir CORTE"])

        with admin_tab1:
            st.markdown(
                "**Archivo:** `REVISIONES_FIBRA_CAPTA_*.xlsx` — actualiza pestañas Front, BO, Aprobado vs KO."
            )
            admin_upload_section(
                "REVISIONES", "rev", "uploader_rev",
                validate_excel, process_excel, save_data
            )

        with admin_tab2:
            st.markdown(
                "**Archivo:** `CORTE_CAPTA_*.xlsx` (debe contener hoja **Volcado Bo**) — actualiza pestañas Corte x Hora y Efectividad."
            )
            admin_upload_section(
                "CORTE", "corte", "uploader_corte",
                validate_corte, process_corte, save_corte_data
            )

        st.markdown("---")

    # ===== Dashboard visible para todos =====
    records = load_data()
    corte_records = load_corte_data()

    if records is None:
        st.warning("⏳ El dashboard aún no tiene datos de Revisiones cargados.")
        st.markdown("**Si eres administrador:** ingresa con tu clave en el panel lateral (←) y sube el archivo de Revisiones.")
    else:
        html = render_dashboard(records, corte_records, is_admin=check_admin())
        if html:
            components.html(html, height=4500, scrolling=True)


if __name__ == "__main__":
    main()
