"""
Dashboard Revisiones Fibra Capta - Streamlit App
=================================================
Solo el admin puede subir el archivo Excel.
Todos los demás solo ven el dashboard (no pueden modificar nada).
"""

import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import json
import os
from pathlib import Path

# ============== CONFIGURACIÓN ==============
st.set_page_config(
    page_title="Dashboard Revisiones Fibra Capta",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

DATA_FILE = Path("data_actual.json")
META_FILE = Path("data_meta.json")
DASHBOARD_TEMPLATE = Path("dashboard_template.html")


# ============== FUNCIONES ==============
REQUIRED_COLUMNS = [
    'COORDINADOR', 'AGENTE FRONT', 'AGENTE BO', 'ESTADO', 'SEMANA', 'DIA',
    'MOTIVO_CANCELACION', 'MOTIVO_CANCELACION3', 'Motivo Bo', 'Sub Motivo Bo',
    'SLA_AGENDA', 'Solicitud', 'FAMILIA', 'SLA_INSTALACION'
]


def validate_excel(uploaded_file) -> dict:
    """
    Valida el archivo antes de procesarlo. Devuelve dict con:
      - 'ok': bool si todo está bien
      - 'errors': lista de problemas críticos (impiden subir)
      - 'warnings': lista de advertencias (se puede subir pero hay que revisar)
      - 'info': info útil para el resumen
    """
    result = {'ok': True, 'errors': [], 'warnings': [], 'info': {}}

    try:
        df = pd.read_excel(uploaded_file, sheet_name=0, header=0)
    except Exception as e:
        result['ok'] = False
        result['errors'].append(f"No se pudo leer el archivo: {e}")
        return result

    # ===== Validación 1: Columnas requeridas =====
    missing_cols = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing_cols:
        result['ok'] = False
        result['errors'].append(
            f"Faltan columnas obligatorias: **{', '.join(missing_cols)}**. "
            "Verifica que estés subiendo el archivo correcto."
        )
        return result

    # ===== Validación 2: Archivo no vacío =====
    if len(df) == 0:
        result['ok'] = False
        result['errors'].append("El archivo está vacío (no tiene filas).")
        return result

    # ===== Procesar para contar casos válidos =====
    valid = df[df['ESTADO'].isin(['TERMINADO','CANCELADO','CAIDA 1RA LINEA','PENDIENTE'])]
    valid = valid[valid['COORDINADOR'].notna()]
    new_count = len(valid)

    if new_count == 0:
        result['ok'] = False
        result['errors'].append(
            "El archivo no tiene ningún caso válido (con ESTADO en TERMINADO/CANCELADO/CAIDA 1RA LINEA/PENDIENTE)."
        )
        return result

    # ===== Validación 3: Comparación con dashboard actual =====
    existing_meta = load_meta()
    if existing_meta:
        old_count = existing_meta.get('total_records', 0)
        diff = new_count - old_count

        result['info']['old_count'] = old_count
        result['info']['new_count'] = new_count
        result['info']['diff'] = diff
        result['info']['old_filename'] = existing_meta.get('filename', '?')
        result['info']['old_date'] = existing_meta.get('last_update', '?')

        if diff < 0:
            # Tiene MENOS casos que el dashboard actual: alerta fuerte
            result['warnings'].append(
                f"⚠️ El archivo nuevo tiene **{abs(diff)} casos MENOS** que el dashboard actual "
                f"({new_count} vs {old_count}). ¿Estás seguro que es el archivo correcto?"
            )
        elif diff == 0:
            result['warnings'].append(
                f"ℹ️ El archivo tiene la misma cantidad de casos que el actual ({new_count}). "
                "Si es el mismo archivo, no hace falta volver a subirlo."
            )
        # Si diff > 0: todo bien, no warning
    else:
        result['info']['new_count'] = new_count
        result['info']['old_count'] = None
        result['info']['diff'] = None

    # ===== Validación 4: Estados sospechosos =====
    estados_invalidos = df[~df['ESTADO'].isin(['TERMINADO','CANCELADO','CAIDA 1RA LINEA','PENDIENTE'])]['ESTADO'].dropna().unique()
    estados_invalidos = [e for e in estados_invalidos if e and str(e).strip()]
    if len(estados_invalidos) > 0:
        result['warnings'].append(
            f"Hay {len(estados_invalidos)} estado(s) que no se reconocen y se ignorarán: "
            f"{', '.join(str(e) for e in estados_invalidos[:5])}"
        )

    # ===== Validación 5: Coordinador vacío =====
    sin_coord = df['COORDINADOR'].isna().sum()
    if sin_coord > 0:
        result['warnings'].append(
            f"Hay {sin_coord} fila(s) sin COORDINADOR — esas se ignorarán."
        )

    # ===== Info adicional para mostrar al admin =====
    result['info']['total_filas_archivo'] = len(df)
    result['info']['casos_validos'] = new_count
    result['info']['coordinadores'] = sorted(valid['COORDINADOR'].dropna().unique().tolist())
    result['info']['semanas'] = sorted(valid['SEMANA'].dropna().unique().tolist())
    result['info']['estados'] = valid['ESTADO'].value_counts().to_dict()

    return result


def process_excel(uploaded_file) -> tuple:
    """Procesa el archivo Excel y devuelve los registros limpios."""
    df = pd.read_excel(uploaded_file, sheet_name=0, header=0)

    # Filtrar filas válidas
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
        out[c] = out[c].astype(str).str.strip()
        out[c] = out[c].replace({'nan':'','None':'','NaT':'','0':''})

    return out.to_dict('records'), len(out)


def load_data():
    """Carga la data guardada del disco."""
    if DATA_FILE.exists():
        with open(DATA_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return None


def load_meta():
    """Carga metadata (fecha de actualización, total casos)."""
    if META_FILE.exists():
        with open(META_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return None


def save_data(records, filename):
    """Guarda la data y metadata."""
    from datetime import datetime
    with open(DATA_FILE, 'w', encoding='utf-8') as f:
        json.dump(records, f, ensure_ascii=False, separators=(',',':'))
    meta = {
        'last_update': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'total_records': len(records),
        'filename': filename
    }
    with open(META_FILE, 'w', encoding='utf-8') as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)


def render_dashboard(records):
    """Genera el HTML del dashboard con la data inyectada."""
    if not DASHBOARD_TEMPLATE.exists():
        st.error("No se encuentra el template del dashboard.")
        return

    with open(DASHBOARD_TEMPLATE, 'r', encoding='utf-8') as f:
        template = f.read()

    data_json = json.dumps(records, ensure_ascii=False, separators=(',',':'))

    # Reemplazar el RAW del template
    import re
    html = re.sub(r'const RAW = \[.*?\];', f'const RAW = {data_json};', template, count=1, flags=re.DOTALL)

    # También reemplazar el conteo de casos en el subtítulo y footer
    total = len(records)
    html = re.sub(r'\d+ casos', f'{total} casos', html)

    return html


# ============== AUTENTICACIÓN ==============
def check_admin():
    """Verifica si el usuario es admin (con clave guardada en secrets)."""
    if 'is_admin' not in st.session_state:
        st.session_state.is_admin = False
    return st.session_state.is_admin


def admin_login_form():
    """Formulario de login en sidebar."""
    with st.sidebar:
        st.markdown("### 🔐 Modo administrador")
        if not check_admin():
            password_input = st.text_input("Clave de admin", type="password", key="pwd_input")
            if st.button("Iniciar sesión"):
                # La clave real viene de st.secrets (configurada en Streamlit Cloud)
                try:
                    correct_pwd = st.secrets["admin_password"]
                except (KeyError, FileNotFoundError):
                    correct_pwd = "cambiar_esta_clave"  # fallback para desarrollo local
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


# ============== INTERFAZ ==============
def main():
    # Sidebar con login y panel de admin
    admin_login_form()

    # Cargar metadata
    meta = load_meta()

    # ===== Header =====
    col1, col2 = st.columns([3, 1])
    with col1:
        st.markdown("# 📊 Dashboard Revisiones Fibra Capta")
        if meta:
            st.caption(f"Última actualización: **{meta['last_update']}** · {meta['total_records']} casos · Archivo: `{meta['filename']}`")
        else:
            st.caption("Sin datos cargados todavía")
    with col2:
        if check_admin():
            st.info("👤 Eres admin")

    st.markdown("---")

    # ===== Panel de Admin (subir archivo) =====
    if check_admin():
        with st.expander("📤 Subir nuevo archivo Excel (solo admin)", expanded=(meta is None)):
            st.markdown("""
            **Instrucciones:**
            1. Selecciona el archivo `.xlsx` con la data del día
            2. Se validará automáticamente antes de actualizar
            3. Click en "Confirmar y actualizar" → todos verán la nueva versión al refrescar
            """)

            uploaded = st.file_uploader(
                "Arrastra el Excel aquí o haz click para seleccionar",
                type=['xlsx', 'xls'],
                help="El archivo se procesará y los gráficos se actualizarán automáticamente",
                key="excel_uploader"
            )

            if uploaded is not None:
                st.markdown(f"📁 Archivo cargado: **{uploaded.name}** ({uploaded.size/1024:.1f} KB)")

                # ===== EJECUTAR VALIDACIÓN =====
                with st.spinner("Validando archivo..."):
                    validation = validate_excel(uploaded)

                # ===== ERRORES CRÍTICOS =====
                if not validation['ok']:
                    st.error("❌ **No se puede subir este archivo**")
                    for err in validation['errors']:
                        st.markdown(f"- {err}")
                    st.markdown("**Corrige el archivo y vuelve a intentarlo.**")
                else:
                    # ===== RESUMEN DEL CONTENIDO =====
                    st.markdown("### 📋 Resumen del archivo")

                    info = validation['info']

                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Casos válidos", info['casos_validos'])
                    with col2:
                        if info.get('old_count') is not None:
                            diff = info['diff']
                            if diff > 0:
                                st.metric("vs Dashboard actual", f"+{diff} casos", delta=f"{diff}")
                            elif diff < 0:
                                st.metric("vs Dashboard actual", f"{diff} casos", delta=f"{diff}")
                            else:
                                st.metric("vs Dashboard actual", "Sin cambios", delta="0")
                        else:
                            st.metric("Dashboard actual", "Vacío")
                    with col3:
                        st.metric("Coordinadores", len(info.get('coordinadores', [])))

                    # Distribución de estados
                    if info.get('estados'):
                        st.markdown("**Distribución de estados:**")
                        cols = st.columns(len(info['estados']))
                        for i, (estado, cnt) in enumerate(info['estados'].items()):
                            with cols[i]:
                                st.metric(estado, cnt)

                    # ===== ADVERTENCIAS =====
                    if validation['warnings']:
                        st.markdown("### ⚠️ Advertencias")
                        for w in validation['warnings']:
                            st.warning(w)

                        # Si hay advertencias, requiere doble confirmación
                        st.markdown("---")
                        confirmar = st.checkbox(
                            "✅ He revisado las advertencias y quiero subir este archivo de todas formas",
                            key="confirm_warnings"
                        )

                        if confirmar:
                            if st.button("🚀 Confirmar y actualizar dashboard", type="primary"):
                                try:
                                    with st.spinner("Actualizando..."):
                                        records, total = process_excel(uploaded)
                                        save_data(records, uploaded.name)
                                    st.success(f"✓ Dashboard actualizado con **{total} casos**")
                                    st.balloons()
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"Error procesando el archivo: {e}")
                                    st.exception(e)
                    else:
                        # Sin advertencias: subida directa con un solo botón
                        st.success("✓ Validación OK · sin alertas")
                        if st.button("🚀 Actualizar dashboard", type="primary"):
                            try:
                                with st.spinner("Actualizando..."):
                                    records, total = process_excel(uploaded)
                                    save_data(records, uploaded.name)
                                st.success(f"✓ Dashboard actualizado con **{total} casos**")
                                st.balloons()
                                st.rerun()
                            except Exception as e:
                                st.error(f"Error procesando el archivo: {e}")
                                st.exception(e)

        st.markdown("---")

    # ===== Dashboard (visible para todos) =====
    records = load_data()

    if records is None:
        st.warning("⏳ El dashboard aún no tiene datos cargados. El administrador debe subir el archivo Excel.")
        st.markdown("""
        **Si eres administrador:**
        - Ingresa con tu clave en el panel lateral izquierdo (←)
        - Sube el archivo Excel del día
        """)
    else:
        # Renderizar el dashboard HTML
        html = render_dashboard(records)
        if html:
            components.html(html, height=4000, scrolling=True)


if __name__ == "__main__":
    main()
