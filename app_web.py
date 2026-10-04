from datetime import datetime
import io
import os
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Gestión de Locales y Mesas Electorales",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Estilos CSS optimizados para adaptabilidad en Celulares y PC
st.markdown(
    """
    <style>
    .main { background-color: #E5ECEA; }

    .header-container {
        display: flex;
        justify-content: space-between;
        align-items: center;
        background-color: #ffffff;
        padding: 8px 12px;
        border-radius: 6px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        margin-bottom: 6px;
        flex-wrap: wrap;
        gap: 8px;
    }
    .header-title {
        font-size: 16px;
        font-weight: 700;
        color: #31333F;
        margin: 0;
    }
    .header-badge-inst {
        background-color: #f1f8f5;
        border-left: 3px solid #28a745;
        padding: 6px 10px;
        border-radius: 4px;
        font-weight: 600;
        color: #31333F;
        font-size: 12px;
    }

    .metric-container-grid-3 {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(130px, 1fr));
        gap: 6px;
        margin-bottom: 8px;
    }

    .metric-container-grid-2 {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
        gap: 6px;
        margin-bottom: 8px;
    }

    .metric-container-grid-4 {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(110px, 1fr));
        gap: 6px;
        margin-bottom: 8px;
    }

    .metric-card {
        background-color: #f8f9fa;
        padding: 8px 10px;
        border-radius: 6px;
        font-weight: 600;
        color: #31333F;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        font-size: 12px;
    }

    .card-local-obs { border-left: 4px solid #ffc107; }
    .card-local-sini { border-left: 4px solid #fd7e14; }
    .card-local-ext { border-left: 4px solid #dc3545; }

    .metric-local-inst {
        background-color: #f1f8f5;
        padding: 6px 10px;
        border-radius: 6px;
        font-weight: 600;
        color: #31333F;
        font-size: 12px;
        border-left: 4px solid #28a745;
        margin-bottom: 6px;
        display: inline-block;
    }

    .metric-container-local-grid-2 {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(140px, 1fr));
        gap: 6px;
        margin-bottom: 6px;
    }

    .metric-container-local-grid-3 {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(120px, 1fr));
        gap: 6px;
        margin-bottom: 6px;
    }

    .metric-container-local-grid-4 {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(100px, 1fr));
        gap: 6px;
        margin-bottom: 6px;
    }

    .metric-local-card {
        padding: 6px 8px;
        border-radius: 4px;
        font-size: 11px;
        font-weight: 600;
        color: #31333F;
    }

    .local-norm { background-color: #f2f9fa; border-left: 3px solid #17a2b8; }
    .local-obs { background-color: #fefcf0; border-left: 3px solid #ffc107; }
    .local-sini { background-color: #fff8f3; border-left: 3px solid #fd7e14; }
    .local-ext { background-color: #fdf2f2; border-left: 3px solid #dc3545; }

    .local-header-box {
        display: flex;
        align-items: center;
        background-color: #ffffff;
        padding: 8px 12px;
        border-radius: 6px;
        border: 1px solid #e0e0e0;
        margin-bottom: 6px;
        flex-wrap: wrap;
        gap: 6px;
    }

    .badge-codigo {
        background-color: #e8f4fd;
        color: #0056b3;
        padding: 3px 8px;
        border-radius: 4px;
        font-size: 13px;
        font-weight: 700;
        margin-right: 4px;
    }
    .badge-mesas {
        background-color: #f1f3f5;
        color: #495057;
        padding: 3px 8px;
        border-radius: 4px;
        font-size: 12px;
        font-weight: 600;
        text-transform: uppercase;
        margin-left: 4px;
    }
    .badge-electores {
        background-color: #e2f0d9;
        color: #27692f;
        padding: 3px 8px;
        border-radius: 4px;
        font-size: 12px;
        font-weight: 700;
        margin-left: 4px;
        border-left: 3px solid #28a745;
    }

    .floating-window {
        background-color: #ffffff;
        border: 1px solid #dcdcdc;
        border-radius: 6px;
        padding: 10px 12px;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.05);
        margin-top: 6px;
        margin-bottom: 6px;
    }

    .stButton>button {
        min-height: 38px;
        font-size: 13px;
        font-weight: 600;
    }
    </style>
""",
    unsafe_allow_html=True,
)


@st.cache_data
def cargar_datos_excel():
    ruta_excel = os.path.join("Data_LV", "Data_LV.xlsx")
    if not os.path.exists(ruta_excel):
        return None
    try:
        return pd.read_excel(ruta_excel, sheet_name="DATA", header=None, dtype=str)
    except Exception:
        return None


df_raw = cargar_datos_excel()

if df_raw is None:
    st.error(
        "❌ No se pudo cargar la hoja 'DATA' del archivo 'Data_LV/Data_LV.xlsx'."
    )
    st.stop()


@st.cache_data
def procesar_estructura(df):
    locales_dict = {}
    cache_mesas = {}
    for i in range(len(df)):
        fila = df.iloc[i]
        if len(fila) > 5:
            distrito, colegio, codigo, mesa = (
                str(fila.iloc[1]).strip(),
                str(fila.iloc[2]).strip(),
                str(fila.iloc[3]).strip(),
                str(fila.iloc[5]).strip(),
            )
            if (
                not distrito
                or distrito.lower() == "nan"
                or distrito.upper() == "DISTRITO"
                or not colegio
                or colegio.lower() == "nan"
                or not mesa
                or mesa.lower() == "nan"
            ):
                continue
            if "." in codigo:
                codigo = codigo.split(".")[0]
            if "." in mesa:
                mesa = mesa.split(".")[0]
            if mesa.isdigit() and len(mesa) <= 6:
                mesa = mesa.zfill(6)

            val_col = lambda idx, default="": (
                str(fila.iloc[idx]).strip()
                if len(fila) > idx
                and pd.notna(fila.iloc[idx])
                and str(fila.iloc[idx]).lower() != "nan"
                else default
            )

            cache_mesas[mesa] = {
                "row_index": i,
                "total_electores": val_col(6, "0").split(".")[0],
                "estado_mesa": val_col(7, "INSTALADA"),
                "estado_acta": val_col(8, "NORMAL"),
                "obs_tipo_eleccion": val_col(9, ""),
                "tipo_observacion": val_col(10, "NINGUNA"),
                "estado_digitacion": val_col(11, ""),
                "resolucion": val_col(12, ""),
                "memorandum": val_col(13, ""),
                "fecha_resolucion": val_col(14, ""),
                "indicador_sino": val_col(15, "NO"),
                "observaciones": val_col(16, ""),
                "electores_votaron": val_col(17, "0").split(".")[0],
            }

            if codigo not in locales_dict:
                locales_dict[codigo] = {
                    "Distrito": distrito.upper(),
                    "Codigo": codigo,
                    "Colegio": colegio,
                    "Mesas": [],
                }
            if mesa not in locales_dict[codigo]["Mesas"]:
                locales_dict[codigo]["Mesas"].append(mesa)

    for codigo in locales_dict:
        locales_dict[codigo]["Mesas"].sort(
            key=lambda x: int(x) if x.isdigit() else x
        )
        locales_dict[codigo]["total_mesas_real"] = len(locales_dict[codigo]["Mesas"])
        locales_dict[codigo]["Mesas_Reales"] = list(locales_dict[codigo]["Mesas"])

    return list(locales_dict.values()), cache_mesas


datos_procesados, cache_mesas = procesar_estructura(df_raw)


def guardar_en_excel(numero_mesa, **kwargs):
    ruta_excel = os.path.join("Data_LV", "Data_LV.xlsx")
    df = cargar_datos_excel()
    if df is None:
        return False

    df_to_save = df.copy()
    while df_to_save.shape[1] < 18:
        df_to_save[df_to_save.shape[1]] = ""

    row_idx = cache_mesas.get(numero_mesa, {}).get("row_index")
    if row_idx is not None and row_idx < len(df_to_save):
        dig_actual_mesa = (
            cache_mesas.get(numero_mesa, {}).get("estado_digitacion", "").strip()
        )
        nuevo_dig_ingresado = kwargs.get("estado_digitacion", "")

        if (
            dig_actual_mesa.isdigit()
            and int(dig_actual_mesa) > 0
            and not nuevo_dig_ingresado
        ):
            orden_eliminado = int(dig_actual_mesa)

            for m_key, m_info in cache_mesas.items():
                val_d = m_info.get("estado_digitacion", "").strip()
                if val_d.isdigit():
                    num_d = int(val_d)
                    if num_d > orden_eliminado:
                        nuevo_num_reducido = f"{num_d - 1:03d}"
                        r_idx = m_info.get("row_index")
                        if r_idx is not None and r_idx < len(df_to_save):
                            df_to_save.iloc[r_idx, 11] = nuevo_num_reducido

        mapeo_cols = {
            7: "estado_mesa",
            8: "estado_acta",
            9: "obs_tipo_eleccion",
            10: "tipo_observacion",
            11: "estado_digitacion",
            12: "resolucion",
            13: "memorandum",
            14: "fecha_resolucion",
            15: "indicador_sino",
            16: "observaciones",
            17: "electores_votaron",
        }
        for col_idx, key in mapeo_cols.items():
            val = kwargs.get(key, "")
            df_to_save.iloc[row_idx, col_idx] = (
                str(val).strip().upper()
                if val
                else (
                    "NO"
                    if key == "indicador_sino"
                    else "0"
                    if key == "electores_votaron"
                    else ""
                )
            )

        try:
            with pd.ExcelWriter(ruta_excel, engine="openpyxl", mode="w") as writer:
                df_to_save.to_excel(writer, sheet_name="DATA", header=False, index=False)
            st.cache_data.clear()
            return True
        except Exception as e:
            st.error(f"Error al guardar en Excel: {e}")
            return False
    return False


def resetear_todo_en_excel():
    ruta_excel = os.path.join("Data_LV", "Data_LV.xlsx")
    df = cargar_datos_excel()
    if df is None:
        return False

    df_to_save = df.copy()
    while df_to_save.shape[1] < 18:
        df_to_save[df_to_save.shape[1]] = ""

    indices_a_resetear = [
        info.get("row_index")
        for info in cache_mesas.values()
        if info.get("row_index") is not None
    ]

    if indices_a_resetear:
        df_to_save.loc[indices_a_resetear, 7] = "INSTALADA"
        df_to_save.loc[indices_a_resetear, 8] = "NORMAL"
        df_to_save.loc[indices_a_resetear, 9] = ""
        df_to_save.loc[indices_a_resetear, 10] = "NINGUNA"
        df_to_save.loc[indices_a_resetear, 11] = ""
        df_to_save.loc[indices_a_resetear, 12] = ""
        df_to_save.loc[indices_a_resetear, 13] = ""
        df_to_save.loc[indices_a_resetear, 14] = ""
        df_to_save.loc[indices_a_resetear, 15] = "NO"
        df_to_save.loc[indices_a_resetear, 16] = ""
        df_to_save.loc[indices_a_resetear, 17] = "0"

    try:
        with pd.ExcelWriter(ruta_excel, engine="openpyxl", mode="w") as writer:
            df_to_save.to_excel(writer, sheet_name="DATA", header=False, index=False)
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"Error al resetear datos en Excel: {e}")
        return False


# --- SISTEMA DE AUTENTICACIÓN ---
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False

if not st.session_state["autenticado"]:
    st.title("🔐 Acceso al Sistema Electoral")
    st.info("Ingresa tus credenciales autorizadas.")

    with st.form("login_form"):
        usuario = st.text_input("Usuario")
        clave = st.text_input("Contraseña", type="password")
        submit_login = st.form_submit_button("Ingresar")

        if submit_login:
            if usuario == "admin" and clave == "electoral2026":
                st.session_state["autenticado"] = True
                st.success("¡Acceso concedido!")
                st.rerun()
            else:
                st.error("❌ Usuario o contraseña incorrectos.")
    st.stop()

# --- INTERFAZ PRINCIPAL ---

with st.sidebar:
    st.markdown("### 🏢 SoftCourse")
    st.markdown(
        "<p style='color: #555; font-size: 12px; margin-top:"
        " -10px;'>Software a su servicio...</p>",
        unsafe_allow_html=True,
    )
    st.markdown("---")

    st.header("🔍 Panel de Control")
    if st.button("🚪 Cerrar Sesión"):
        st.session_state["autenticado"] = False
        for k in [
            "local_codigo_activo",
            "mesa_busqueda_activa",
            "mesa_activa",
            "tipo_busq_anterior",
            "confirmar_reset",
        ]:
            if k in st.session_state:
                del st.session_state[k]
        st.rerun()

    if st.button("🔄 Resetear Todo", use_container_width=True):
        st.session_state["confirmar_reset"] = True

    if st.session_state.get("confirmar_reset", False):
        st.sidebar.markdown(
            """
            <div style="background-color: #fff3cd; border: 1px solid #ffeeba; padding: 10px; border-radius: 6px; margin-top: 6px; font-size: 12px; color: #856404;">
                <b>⚠ ¿Estás seguro de resetear todo?</b><br>
                Se restablecerán todas las mesas a INSTALADA, actas a NORMAL y votantes a 0.
            </div>
            """,
            unsafe_allow_html=True,
        )
        col_s1, col_s2 = st.sidebar.columns(2)
        with col_s1:
            if st.button("Sí, Resetear", use_container_width=True, key="btn_conf_si"):
                exito_reset = resetear_todo_en_excel()
                if exito_reset:
                    st.success("¡Sistema reseteado con éxito!")
                    for k in [
                        "local_codigo_activo",
                        "mesa_busqueda_activa",
                        "mesa_activa",
                        "confirmar_reset",
                    ]:
                        if k in st.session_state:
                            del st.session_state[k]
                    st.rerun()
        with col_s2:
            if st.button("Cancelar", use_container_width=True, key="btn_conf_no"):
                del st.session_state["confirmar_reset"]
                st.rerun()

    tipo_busq = st.sidebar.selectbox(
        "Buscar por:",
        [
            "Distrito",
            "Nombre de Colegio / Local",
            "Número de Mesa",
            "Código",
            "Estado de Acta",
            "Estado de Mesa",
            "📋 Actas Ingresadas (Orden)",
        ],
    )

if "tipo_busq_anterior" not in st.session_state:
    st.session_state["tipo_busq_anterior"] = tipo_busq

if st.session_state["tipo_busq_anterior"] != tipo_busq:
    st.session_state["tipo_busq_anterior"] = tipo_busq
    for k in ["mesa_activa", "local_codigo_activo"]:
        if k in st.session_state:
            del st.session_state[k]
    st.rerun()

# =========================================================================
# VISTA: LISTA DE ACTAS INGRESADAS (CON DETALLES COMPLETOS Y EXPORTACIÓN)
# =========================================================================
if tipo_busq == "📋 Actas Ingresadas (Orden)":
    st.title("📋 Listado de Actas Ingresadas (Orden de Ingreso)")
    st.markdown(
        "Visualización en secuencia del orden en que fueron ingresadas las"
        " actas (del 001 al 830)."
    )

    info_mesas_map = {}
    for local in datos_procesados:
        dist = local["Distrito"]
        col = local["Colegio"]
        for m in local["Mesas_Reales"]:
            info_mesas_map[m] = {"Colegio": col, "Distrito": dist}

    actas_ingresadas_lista = []
    for mesa_id, info in cache_mesas.items():
        dig_val = info.get("estado_digitacion", "").strip()
        if dig_val.isdigit() and int(dig_val) > 0:
            extras = info_mesas_map.get(
                mesa_id, {"Colegio": "Desconocido", "Distrito": "Desconocido"}
            )
            actas_ingresadas_lista.append((int(dig_val), mesa_id, info, extras))

    actas_ingresadas_lista.sort(key=lambda x: x[0])

    if not actas_ingresadas_lista:
        st.info("Aún no hay actas marcadas como 'Acta Ingresada' en el sistema.")
    else:
        col_inf_1, col_inf_2 = st.columns([3, 1])
        with col_inf_1:
            st.markdown(
                f"**Total de actas ingresadas hasta el momento:**"
                f" {len(actas_ingresadas_lista)} / 830"
            )

        with col_inf_2:
            data_export = []
            for orden_num, m_id, info_m, extras in actas_ingresadas_lista:
                data_export.append({
                    "Orden Ingreso": f"{orden_num:03d}",
                    "Número de Mesa": m_id,
                    "Distrito": extras["Distrito"],
                    "Local de Votación": extras["Colegio"],
                    "Estado de Mesa": info_m.get("estado_mesa", ""),
                    "Estado de Acta": info_m.get("estado_acta", ""),
                    "Votantes": info_m.get("electores_votaron", "0"),
                    "Total Electores": info_m.get("total_electores", "0"),
                    "Tipo Elección Observada": info_m.get("obs_tipo_eleccion", ""),
                    "Tipo de Observación": info_m.get("tipo_observacion", ""),
                    "Resolución": info_m.get("resolucion", ""),
                    "Memorándum": info_m.get("memorandum", ""),
                    "Fecha Resolución": info_m.get("fecha_resolucion", ""),
                    "Resuelta?": info_m.get("indicador_sino", "NO"),
                    "Otras Observaciones": info_m.get("observaciones", ""),
                })
            df_export = pd.DataFrame(data_export)

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df_export.to_excel(writer, index=False, sheet_name="Actas_Ingresadas")
            excel_data = output.getvalue()

            st.download_button(
                label="📥 Exportar Excel Completo",
                data=excel_data,
                file_name="Detalle_Actas_Ingresadas.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ),
                use_container_width=True,
            )

        st.markdown("---")
        for orden_num, m_id, info_m, extras in actas_ingresadas_lista:
            orden_str = f"{orden_num:03d}"
            st.markdown(
                f"**{orden_str} - Mesa {m_id}** "
                f"(Estado Acta: {info_m.get('estado_acta', 'NORMAL')}) | "
                f"🏫 **Local:** {extras['Colegio']} | "
                f"📍 **Distrito:** {extras['Distrito']}"
            )

# =========================================================================
# VISTAS DE GESTIÓN Y BÚSQUEDA TRADICIONAL
# =========================================================================
else:
    st.title("🗳️ Sistema de Gestión de Locales y Mesas Electorales")

    buscando_por_mesa = False
    mesa_busq_input = ""
    est_acta_sel = ""

    if tipo_busq == "Distrito":
        opciones = sorted(list(set(item["Distrito"] for item in datos_procesados)))
        val_sel = st.sidebar.selectbox("Selecciona Distrito", opciones)
        resultados = [
            item for item in datos_procesados if item["Distrito"] == val_sel
        ]
    elif tipo_busq == "Nombre de Colegio / Local":
        opciones = sorted(list(set(item["Colegio"] for item in datos_procesados)))
        val_sel = st.sidebar.selectbox("Selecciona Colegio", opciones)
        resultados = [
            item for item in datos_procesados if item["Colegio"].upper() == val_sel
        ]
    elif tipo_busq == "Código":
        opciones = sorted(list(set(item["Codigo"] for item in datos_procesados)))
        val_sel = st.sidebar.selectbox("Selecciona Código", opciones)
        resultados = [
            item for item in datos_procesados if item["Codigo"] == val_sel
        ]
    elif tipo_busq == "Número de Mesa":
        mesa_input = st.sidebar.text_input("Ingresa N° de Mesa").strip()
        resultados = []
        if mesa_input:
            mesa_busq = mesa_input.zfill(6)
            buscando_por_mesa = True
            mesa_busq_input = mesa_busq
            for item in datos_procesados:
                if mesa_busq in item["Mesas_Reales"]:
                    item_copia = item.copy()
                    item_copia["Mesas"] = [mesa_busq]
                    resultados.append(item_copia)
    elif tipo_busq == "Estado de Acta":
        est_acta_sel = st.sidebar.selectbox(
            "Selecciona Estado de Acta",
            ["NORMAL", "OBSERVADA", "SINIESTRADA", "EXTRAVIADA", "RESUELTA"],
        )
        resultados = []
        for item in datos_procesados:
            mesas_filtradas = []
            for m in item["Mesas_Reales"]:
                d_m = cache_mesas.get(m, {})
                e_mesa = d_m.get("estado_mesa", "").upper()
                e_acta = d_m.get("estado_acta", "").upper()
                ind_sino = d_m.get("indicador_sino", "").upper()

                if e_mesa in ["INSTALADA", "TARDÍA"]:
                    if est_acta_sel == "RESUELTA":
                        if ind_sino == "SI":
                            mesas_filtradas.append(m)
                    else:
                        if e_acta == est_acta_sel:
                            mesas_filtradas.append(m)

            if mesas_filtradas:
                item_copia = item.copy()
                item_copia["Mesas"] = sorted(
                    mesas_filtradas, key=lambda x: int(x) if x.isdigit() else x
                )
                resultados.append(item_copia)
    elif tipo_busq == "Estado de Mesa":
        est_mesa_sel = st.sidebar.selectbox(
            "Selecciona Estado de Mesa", ["INSTALADA", "NO INSTALADA"]
        )
        resultados = []
        for item in datos_procesados:
            mesas_filtradas = []
            for m in item["Mesas_Reales"]:
                e_mesa = cache_mesas.get(m, {}).get("estado_mesa", "").upper()
                if est_mesa_sel == "INSTALADA" and e_mesa in ["INSTALADA", "TARDÍA"]:
                    mesas_filtradas.append(m)
                elif est_mesa_sel == "NO INSTALADA" and e_mesa == "NO INSTALADA":
                    mesas_filtradas.append(m)
            if mesas_filtradas:
                item_copia = item.copy()
                item_copia["Mesas"] = sorted(
                    mesas_filtradas, key=lambda x: int(x) if x.isdigit() else x
                )
                resultados.append(item_copia)

    st.markdown("---")

    if not resultados:
        st.info(
            "💡 Selecciona un criterio o ingresa un número de mesa válido en la"
            " barra lateral izquierda."
        )
    else:
        codigo_activo = st.session_state.get("local_codigo_activo")

        if buscando_por_mesa and resultados:
            codigo_activo = resultados[0]["Codigo"]
            st.session_state["local_codigo_activo"] = codigo_activo

        if codigo_activo:
            resultados_filtrados = [
                item for item in resultados if item["Codigo"] == codigo_activo
            ]
            if not resultados_filtrados:
                if "local_codigo_activo" in st.session_state:
                    del st.session_state["local_codigo_activo"]
                resultados_filtrados = resultados
        else:
            resultados_filtrados = resultados

        if "mesa_activa" in st.session_state:
            m_activa = st.session_state["mesa_activa"]
            d_prev = cache_mesas.get(m_activa, {})

            filtro_actual_es_resuelta = (
                tipo_busq == "Estado de Acta" and est_acta_sel == "RESUELTA"
            )
            es_resuelta_real = d_prev.get("indicador_sino", "").upper() == "SI"

            modo_solo_lectura = filtro_actual_es_resuelta and es_resuelta_real
            tiene_historial_resolucion = (
                es_resuelta_real
                or bool(d_prev.get("resolucion", "").strip())
                or d_prev.get("tipo_observacion", "NINGUNA") != "NINGUNA"
            )

            if st.button("⬅️ Volver al listado"):
                del st.session_state["mesa_activa"]
                st.rerun()

            titulo_ventana = (
                f"👁 Visualización de Mesa N°: {m_activa} (RESUELTA)"
                if modo_solo_lectura
                else f"🗂 Gestión de Mesa N°: {m_activa}"
            )

            st.markdown(
                f"""
            <div class="floating-window">
                <div style="font-size: 15px; font-weight: 700; color: {'#27692f' if modo_solo_lectura else '#1f77b4'}; border-bottom: 1px solid #eaeaea; padding-bottom: 6px; margin-bottom: 8px;">
                    {titulo_ventana}
                </div>
            """,
                unsafe_allow_html=True,
            )

            col_izq, col_der = st.columns(2)

            with col_izq:
                opciones_est_mesa = ["INSTALADA", "NO INSTALADA", "TARDÍA"]
                idx_est_mesa = (
                    opciones_est_mesa.index(d_prev.get("estado_mesa", "INSTALADA"))
                    if d_prev.get("estado_mesa", "INSTALADA") in opciones_est_mesa
                    else 0
                )
                nuevo_est_mesa = st.selectbox(
                    "Estado de Mesa",
                    opciones_est_mesa,
                    index=idx_est_mesa,
                    key=f"est_mesa_edit_{m_activa}",
                    disabled=modo_solo_lectura,
                )

                if nuevo_est_mesa == "NO INSTALADA":
                    nuevo_est_acta = "NO APLICA"
                    electores_votaron = "0"
                    st.markdown(
                        "<p style='font-size:11px; color:#721C24; margin:0;'>ℹ No aplica acta</p>",
                        unsafe_allow_html=True,
                    )
                else:
                    opciones_est_acta = [
                        "NORMAL",
                        "OBSERVADA",
                        "SINIESTRADA",
                        "EXTRAVIADA",
                    ]
                    val_acta_prev = d_prev.get("estado_acta", "NORMAL")
                    if val_acta_prev not in opciones_est_acta:
                        val_acta_prev = "NORMAL"
                    idx_est_acta = opciones_est_acta.index(val_acta_prev)

                    nuevo_est_acta = st.selectbox(
                        "Estado de Acta",
                        opciones_est_acta,
                        index=idx_est_acta,
                        key=f"est_acta_edit_{m_activa}",
                        disabled=modo_solo_lectura,
                    )

                    electores_votaron = st.text_input(
                        "Votantes",
                        value=d_prev.get("electores_votaron", "0"),
                        key=f"electores_edit_{m_activa}",
                        disabled=modo_solo_lectura,
                    )

            with col_der:
                dig_actual = d_prev.get("estado_digitacion", "").strip()
                ya_ingresada_inicial = dig_actual.isdigit() and int(dig_actual) > 0

                acta_ingresada = st.checkbox(
                    "Acta Ingresada",
                    value=ya_ingresada_inicial,
                    key=f"chk_ingresada_{m_activa}",
                    disabled=modo_solo_lectura,
                    help=(
                        "Marca para asignar automáticamente orden correlativo del 001"
                        " al 830"
                    ),
                )

                if ya_ingresada_inicial:
                    st.markdown(
                        f"<p style='font-size:11px; color:#0056b3; margin:0;'>📌 Orden"
                        f" asignado actual: <b>{int(dig_actual):03d}</b></p>",
                        unsafe_allow_html=True,
                    )

                observaciones = st.text_input(
                    "Otras Observaciones",
                    value=d_prev.get("observaciones", ""),
                    key=f"obs_edit_{m_activa}",
                    disabled=modo_solo_lectura,
                )

            chk_prov = False
            chk_dist = False
            nuevo_tipo_obs = "NINGUNA"
            resolucion = ""
            memorandum = ""
            fecha_resolucion = ""
            indicador_sino = "NO"

            mostrar_seccion_obs = (nuevo_est_mesa != "NO INSTALADA") and (
                nuevo_est_acta in ["OBSERVADA", "SINIESTRADA", "EXTRAVIADA"]
                or tiene_historial_resolucion
            )

            if mostrar_seccion_obs:
                st.markdown("<hr style='margin: 6px 0;'>", unsafe_allow_html=True)
                st.markdown(
                    "<p style='font-size: 12px; font-weight: 600; margin: 0 0 4px"
                    " 0;'>📝 Detalles de Observación / Resolución</p>",
                    unsafe_allow_html=True,
                )

                val_actual_obs = d_prev.get("obs_tipo_eleccion", "").upper()
                c_chk1, c_chk2 = st.columns(2)

                if nuevo_est_acta in ["SINIESTRADA", "EXTRAVIADA"]:
                    st.session_state[f"chk_prov_edit_{m_activa}"] = True
                    st.session_state[f"chk_dist_edit_{m_activa}"] = True
                    with c_chk1:
                        chk_prov = st.checkbox(
                            "Provincial",
                            disabled=True,
                            key=f"chk_prov_edit_{m_activa}",
                        )
                    with c_chk2:
                        chk_dist = st.checkbox(
                            "Distrital", disabled=True, key=f"chk_dist_edit_{m_activa}"
                        )
                else:
                    with c_chk1:
                        chk_prov = st.checkbox(
                            "Provincial",
                            value=any(x in val_actual_obs for x in ["PROVINCIAL", "AMBAS"]),
                            key=f"chk_prov_edit_{m_activa}",
                            disabled=modo_solo_lectura,
                        )
                    with c_chk2:
                        chk_dist = st.checkbox(
                            "Distrital",
                            value=any(x in val_actual_obs for x in ["DISTRITAL", "AMBAS"]),
                            key=f"chk_dist_edit_{m_activa}",
                            disabled=modo_solo_lectura,
                        )

                tipos_obs_lista = [
                    "NINGUNA",
                    "ACTA CON ERROR ARITMETICO",
                    "ACTA CON VOTOS IMPUGNADOS",
                    "ACTA CON ILEGIBILIDAD",
                    "ACTA INCOMPLETA",
                    "ACTA CON SOLICITUD DE NULIDAD DE LA MESA",
                    "ACTA SIN DATOS",
                    "ACTA SIN FIRMA",
                    "ACTA CON MAS DE UNA OBSERVACION",
                    "OTRAS OBSERVACIONES",
                ]
                t_obs_actual = d_prev.get("tipo_observacion", "NINGUNA")
                idx_t_obs = (
                    tipos_obs_lista.index(t_obs_actual)
                    if t_obs_actual in tipos_obs_lista
                    else 0
                )
                nuevo_tipo_obs = st.selectbox(
                    "Tipo de Observación",
                    tipos_obs_lista,
                    index=idx_t_obs,
                    key=f"t_obs_edit_{m_activa}",
                    disabled=modo_solo_lectura,
                )

            rc1, rc2, rc3, rc4 = st.columns([1, 1, 1, 0.9])
            with rc1:
                resolucion = st.text_input(
                    "Resolución",
                    value=d_prev.get("resolucion", ""),
                    key=f"res_edit_{m_activa}",
                    disabled=modo_solo_lectura,
                )
            with rc2:
                memorandum = st.text_input(
                    "Memorándum",
                    value=d_prev.get("memorandum", ""),
                    key=f"mem_edit_{m_activa}",
                    disabled=modo_solo_lectura,
                )
            with rc3:
                fecha_resolucion = st.text_input(
                    "Fecha (DD/MM/AAAA)",
                    value=d_prev.get("fecha_resolucion", ""),
                    key=f"f_res_edit_{m_activa}",
                    disabled=modo_solo_lectura,
                )
            with rc4:
                opciones_sino = ["NO", "SI"]
                idx_sino = (
                    opciones_sino.index(d_prev.get("indicador_sino", "NO"))
                    if d_prev.get("indicador_sino", "NO") in opciones_sino
                    else 0
                )
                indicador_sino = st.selectbox(
                    "Resuelta?",
                    opciones_sino,
                    index=idx_sino,
                    key=f"ind_sino_edit_{m_activa}",
                    disabled=modo_solo_lectura,
                )

            if modo_solo_lectura:
                st.info(
                    "🔒 Esta acta ya se encuentra **RESUELTA** (Vista de Histórico)."
                    " Sus datos son de solo visualización."
                )
                if st.button("Cerrar Vista", key=f"btn_cerrar_res_{m_activa}"):
                    del st.session_state["mesa_activa"]
                    st.rerun()
            else:
                bloquear_guardado = False
                if nuevo_est_mesa != "NO INSTALADA" and nuevo_est_acta == "OBSERVADA":
                    eleccion_seleccionada = chk_prov or chk_dist
                    tipo_obs_seleccionado = nuevo_tipo_obs != "NINGUNA"

                    if not (eleccion_seleccionada and tipo_obs_seleccionado):
                        bloquear_guardado = True
                        st.warning(
                            "⚠️ **Faltan datos:** Marca elección y Tipo de Observación."
                        )

                if ya_ingresada_inicial and not acta_ingresada:
                    st.warning(
                        "⚠️ **Atención:** Has desmarcado 'Acta Ingresada'. Al guardar,"
                        " perderá su posición y las actas posteriores descenderán para"
                        " cerrar el espacio."
                    )

                if bloquear_guardado:
                    st.button(
                        "💾 Guardar Cambios (Incompleto)",
                        disabled=True,
                        key=f"btn_guardar_inc_{m_activa}",
                    )
                else:
                    if st.button(
                        "💾 Guardar Cambios en Excel", key=f"btn_guardar_ok_{m_activa}"
                    ):
                        if nuevo_est_mesa == "NO INSTALADA":
                            val_acta_guardar = ""
                            v_votos = 0
                        else:
                            val_acta_guardar = nuevo_est_acta
                            try:
                                v_votos = int(electores_votaron)
                                t_electores = int(d_prev.get("total_electores", 0) or 0)
                                if v_votos < 0 or v_votos > t_electores:
                                    st.error(f"⚠ Votos entre 0 y {t_electores}.")
                                    st.stop()
                            except ValueError:
                                st.error("⚠ Número de electores inválido.")
                                st.stop()

                        if nuevo_est_mesa != "NO INSTALADA" and fecha_resolucion.strip():
                            try:
                                datetime.strptime(fecha_resolucion.strip(), "%d/%m/%Y")
                            except ValueError:
                                st.error("⚠ Fecha formato DD/MM/AAAA.")
                                st.stop()

                        if acta_ingresada:
                            if ya_ingresada_inicial:
                                nuevo_dig_val = dig_actual
                            else:
                                max_n = 0
                                for _, inf_m in cache_mesas.items():
                                    val_d = inf_m.get("estado_digitacion", "").strip()
                                    if val_d.isdigit():
                                        num_val = int(val_d)
                                        if num_val > max_n:
                                            max_n = num_val
                                nuevo_n = max_n + 1 if max_n < 830 else 830
                                nuevo_dig_val = f"{nuevo_n:03d}"
                        else:
                            nuevo_dig_val = ""

                        if nuevo_est_mesa != "NO INSTALADA":
                            if nuevo_est_acta in ["SINIESTRADA", "EXTRAVIADA"]:
                                obs_tipo_res = "AMBAS"
                            else:
                                obs_tipo_res = (
                                    "AMBAS"
                                    if chk_prov and chk_dist
                                    else "PROVINCIAL"
                                    if chk_prov
                                    else "DISTRITAL"
                                )
                            t_obs_res = nuevo_tipo_obs
                        else:
                            obs_tipo_res = ""
                            t_obs_res = ""
                            resolucion = ""
                            memorandum = ""
                            fecha_resolucion = ""
                            indicador_sino = "NO"

                        exito = guardar_en_excel(
                            numero_mesa=m_activa,
                            estado_mesa=nuevo_est_mesa,
                            estado_acta=val_acta_guardar,
                            obs_tipo_eleccion=obs_tipo_res,
                            tipo_observacion=t_obs_res,
                            estado_digitacion=nuevo_dig_val,
                            resolucion=resolucion,
                            memorandum=memorandum,
                            fecha_resolucion=fecha_resolucion,
                            indicador_sino=indicador_sino,
                            observaciones=observaciones,
                            electores_votaron=str(v_votos),
                        )

                        if exito:
                            st.success(f"✅ ¡Mesa {m_activa} guardada con éxito!")
                            del st.session_state["mesa_activa"]
                            st.rerun()

            st.markdown("</div>", unsafe_allow_html=True)

        else:
            # --- CÁLCULO DE ELECTORES GLOBALES Y MESAS (TOTAL GENERAL ABSOLUTO) ---
            es_filtro_resuelta = (
                tipo_busq == "Estado de Acta" and est_acta_sel == "RESUELTA"
            )
            es_filtro_observada = (
                tipo_busq == "Estado de Acta" and est_acta_sel == "OBSERVADA"
            )
            es_filtro_siniestrada = (
                tipo_busq == "Estado de Acta" and est_acta_sel == "SINIESTRADA"
            )
            es_filtro_extraviada = (
                tipo_busq == "Estado de Acta" and est_acta_sel == "EXTRAVIADA"
            )

            total_locales_encontrados = len(resultados_filtrados)
            total_mesas_reales_global = sum(
                local["total_mesas_real"] for local in resultados_filtrados
            )

            meses_instaladas_global = 0
            normales_global = 0
            observadas_global = 0
            siniestradas_global = 0
            extraviadas_global = 0

            resueltas_obs_global = 0
            resueltas_sini_global = 0
            resueltas_ext_global = 0

            siniestradas_encontradas_global = 0
            siniestradas_no_encontradas_global = 0
            extraviadas_encontradas_global = 0
            extraviadas_no_encontradas_global = 0

            tipos_obs_posibles = [
                "ACTA CON ERROR ARITMETICO",
                "ACTA CON VOTOS IMPUGNADOS",
                "ACTA CON ILEGIBILIDAD",
                "ACTA INCOMPLETA",
                "ACTA CON SOLICITUD DE NULIDAD DE LA MESA",
                "ACTA SIN DATOS",
                "ACTA SIN FIRMA",
                "ACTA CON MAS DE UNA OBSERVACION",
                "OTRAS OBSERVACIONES",
            ]
            conteo_tipos_obs_global = {t: 0 for t in tipos_obs_posibles}

            n_electores_global = 0
            x_electores_global = 0

            for local_item in datos_procesados:
                for m in local_item["Mesas_Reales"]:
                    d_m = cache_mesas.get(m, {})
                    try:
                        x_electores_global += int(d_m.get("total_electores", 0) or 0)
                    except ValueError:
                        pass
                    try:
                        n_electores_global += int(d_m.get("electores_votaron", 0) or 0)
                    except ValueError:
                        pass

                    est_m = d_m.get("estado_mesa", "").upper()
                    est_a = d_m.get("estado_acta", "").upper()

                    if est_m in ["INSTALADA", "TARDÍA"]:
                        meses_instaladas_global += 1
                        if est_a == "NORMAL":
                            normales_global += 1
                        elif est_a == "OBSERVADA":
                            observadas_global += 1
                        elif est_a == "SINIESTRADA":
                            siniestradas_global += 1
                        elif est_a == "EXTRAVIADA":
                            extraviadas_global += 1

            for local in resultados_filtrados:
                for m in local["Mesas_Reales"]:
                    d_m = cache_mesas.get(m, {})
                    est_m = d_m.get("estado_mesa", "").upper()
                    est_a = d_m.get("estado_acta", "").upper()
                    ind_sino = d_m.get("indicador_sino", "").upper()
                    t_obs = d_m.get("tipo_observacion", "").upper()
                    res_val = d_m.get("resolucion", "").strip()

                    if es_filtro_resuelta:
                        if ind_sino == "SI":
                            if est_a == "OBSERVADA":
                                resueltas_obs_global += 1
                            elif est_a == "SINIESTRADA":
                                resueltas_sini_global += 1
                            elif est_a == "EXTRAVIADA":
                                resueltas_ext_global += 1
                    elif es_filtro_observada:
                        if est_m in ["INSTALADA", "TARDÍA"] and est_a == "OBSERVADA":
                            if t_obs in conteo_tipos_obs_global:
                                conteo_tipos_obs_global[t_obs] += 1
                    elif es_filtro_siniestrada:
                        if est_m in ["INSTALADA", "TARDÍA"] and est_a == "SINIESTRADA":
                            if bool(res_val):
                                siniestradas_encontradas_global += 1
                            else:
                                siniestradas_no_encontradas_global += 1
                    elif es_filtro_extraviada:
                        if est_m in ["INSTALADA", "TARDÍA"] and est_a == "EXTRAVIADA":
                            if bool(res_val):
                                extraviadas_encontradas_global += 1
                            else:
                                extraviadas_no_encontradas_global += 1

            if es_filtro_resuelta:
                total_mesas_reales_global = (
                    resueltas_obs_global + resueltas_sini_global + resueltas_ext_global
                )
            elif es_filtro_observada:
                total_mesas_reales_global = sum(conteo_tipos_obs_global.values())
            elif es_filtro_siniestrada:
                total_mesas_reales_global = (
                    siniestradas_encontradas_global
                    + siniestradas_no_encontradas_global
                )
            elif es_filtro_extraviada:
                total_mesas_reales_global = (
                    extraviadas_encontradas_global + extraviadas_no_encontradas_global
                )

            if codigo_activo and not buscando_por_mesa:
                if st.button("⬅ Volver a la lista de locales"):
                    del st.session_state["local_codigo_activo"]
                    st.rerun()

            if es_filtro_resuelta:
                titulo_principal = f"📌 Actas Resueltas ({total_mesas_reales_global})"
            elif es_filtro_observada:
                titulo_principal = f"📌 Actas Observadas ({total_mesas_reales_global})"
            elif es_filtro_siniestrada:
                titulo_principal = f"📌 Actas Siniestradas ({total_mesas_reales_global})"
            elif es_filtro_extraviada:
                titulo_principal = f"📌 Actas Extraviadas ({total_mesas_reales_global})"
            else:
                titulo_principal = f"📍 Locales Encontrados ({total_locales_encontrados})"

            st.markdown(
                f"""
    <div class="header-container">
        <h3 class="header-title">{titulo_principal}</h3>
        <div style="display: flex; gap: 8px; align-items: center; flex-wrap: wrap;">
            <div class="header-badge-inst">
                👥 Electores Globales: <b>{n_electores_global:,} / {x_electores_global:,}</b>
            </div>
            <div class="header-badge-inst">
                🟢 Mesas Instaladas Globales: <b>{meses_instaladas_global} / 830</b>
            </div>
        </div>
    </div>
""",
                unsafe_allow_html=True,
            )

            if es_filtro_resuelta:
                st.markdown(
                    f"""
        <div class="metric-container-grid-3">
            <div class="metric-card card-local-obs">📊 Observadas: <b>{resueltas_obs_global} / {total_mesas_reales_global}</b></div>
            <div class="metric-card card-local-sini">🦹‍♂ Siniestradas: <b>{resueltas_sini_global} / {total_mesas_reales_global}</b></div>
            <div class="metric-card card-local-ext">❌ Extraviadas: <b>{resueltas_ext_global} / {total_mesas_reales_global}</b></div>
        </div>
    """,
                    unsafe_allow_html=True,
                )
            elif es_filtro_observada:
                tipos_activos_global = [
                    (t, count)
                    for t, count in conteo_tipos_obs_global.items()
                    if count > 0
                ]
                if not tipos_activos_global:
                    st.info("No hay actas observadas registradas.")
                else:
                    cols_obs = st.columns(min(3, len(tipos_activos_global)))
                    for idx_t, (t_name, t_count) in enumerate(tipos_activos_global):
                        with cols_obs[idx_t % len(cols_obs)]:
                            st.markdown(
                                f"""
                <div class="metric-card card-local-obs" style="margin-bottom: 4px;">
                    ⚠️ {t_name.title()}: <b>{t_count} / {total_mesas_reales_global}</b>
                </div>
            """,
                                unsafe_allow_html=True,
                            )
            elif es_filtro_siniestrada:
                st.markdown(
                    f"""
        <div class="metric-container-grid-2">
            <div class="metric-card card-local-sini">✅ Siniestradas (Con Resolución): <b>{siniestradas_encontradas_global} / {total_mesas_reales_global}</b></div>
            <div class="metric-card card-local-sini">❌ Siniestradas (Sin Resolución): <b>{siniestradas_no_encontradas_global} / {total_mesas_reales_global}</b></div>
        </div>
    """,
                    unsafe_allow_html=True,
                )
            elif es_filtro_extraviada:
                st.markdown(
                    f"""
        <div class="metric-container-grid-2">
            <div class="metric-card card-local-ext">✅ Extraviadas (Con Resolución): <b>{extraviadas_encontradas_global} / {total_mesas_reales_global}</b></div>
            <div class="metric-card card-local-ext">❌ Extraviadas (Sin Resolución): <b>{extraviadas_no_encontradas_global} / {total_mesas_reales_global}</b></div>
        </div>
    """,
                    unsafe_allow_html=True,
                )
            else:
                st.markdown(
                    f"""
        <div class="metric-container-grid-4">
            <div class="metric-card card-local-norm">📘 Normales: <b>{normales_global} / {meses_instaladas_global}</b></div>
            <div class="metric-card card-local-obs">📊 Observadas: <b>{observadas_global} / {meses_instaladas_global}</b></div>
            <div class="metric-card card-local-sini">🦹‍♂ Siniestradas: <b>{siniestradas_global} / {meses_instaladas_global}</b></div>
            <div class="metric-card card-local-ext">❌ Extraviadas: <b>{extraviadas_global} / {meses_instaladas_global}</b></div>
        </div>
    """,
                    unsafe_allow_html=True,
                )

            for local in resultados_filtrados:
                codigo_local = local["Codigo"]

                if es_filtro_resuelta:
                    obs_l = sini_l = ext_l = 0
                    for m in local["Mesas_Reales"]:
                        d_m = cache_mesas.get(m, {})
                        if d_m.get("indicador_sino", "").upper() == "SI":
                            ea = d_m.get("estado_acta", "").upper()
                            if ea == "OBSERVADA":
                                obs_l += 1
                            elif ea == "SINIESTRADA":
                                sini_l += 1
                            elif ea == "EXTRAVIADA":
                                ext_l += 1

                    total_mesas_local_count = obs_l + sini_l + ext_l
                    st.markdown(
                        f"""
          <div class="metric-container-local-grid-3">
              <div class="metric-local-card local-obs">📊 Observadas: <b>{obs_l} / {total_mesas_local_count}</b></div>
              <div class="metric-local-card local-sini">🦹‍♂ Siniestradas: <b>{sini_l} / {total_mesas_local_count}</b></div>
              <div class="metric-local-card local-ext">❌ Extraviadas: <b>{ext_l} / {total_mesas_local_count}</b></div>
          </div>
      """,
                        unsafe_allow_html=True,
                    )
                elif es_filtro_observada:
                    conteo_tipos_obs_local = {t: 0 for t in tipos_obs_posibles}
                    for m in local["Mesas_Reales"]:
                        d_m = cache_mesas.get(m, {})
                        if (
                            d_m.get("estado_acta", "").upper() == "OBSERVADA"
                            and d_m.get("estado_mesa", "").upper()
                            in ["INSTALADA", "TARDÍA"]
                        ):
                            t_obs_m = d_m.get("tipo_observacion", "").upper()
                            if t_obs_m in conteo_tipos_obs_local:
                                conteo_tipos_obs_local[t_obs_m] += 1

                    total_mesas_local_count = sum(conteo_tipos_obs_local.values())
                    tipos_activos_local = [
                        (t, count)
                        for t, count in conteo_tipos_obs_local.items()
                        if count > 0
                    ]
                    if tipos_activos_local:
                        cols_lobs = st.columns(min(3, len(tipos_activos_local)))
                        for idx_lt, (t_name, t_count) in enumerate(tipos_activos_local):
                            with cols_lobs[idx_lt % len(cols_lobs)]:
                                st.markdown(
                                    f"""
                  <div class="metric-local-card local-obs" style="margin-bottom: 3px;">
                      ⚠️ {t_name.title()}: <b>{t_count} / {total_mesas_local_count}</b>
                  </div>
              """,
                                    unsafe_allow_html=True,
                                )
                elif es_filtro_siniestrada:
                    sini_enc = 0
                    sini_no_enc = 0
                    for m in local["Mesas_Reales"]:
                        d_m = cache_mesas.get(m, {})
                        if (
                            d_m.get("estado_acta", "").upper() == "SINIESTRADA"
                            and d_m.get("estado_mesa", "").upper()
                            in ["INSTALADA", "TARDÍA"]
                        ):
                            if bool(d_m.get("resolucion", "").strip()):
                                sini_enc += 1
                            else:
                                sini_no_enc += 1
                    total_mesas_local_count = sini_enc + sini_no_enc
                    st.markdown(
                        f"""
          <div class="metric-container-local-grid-2">
              <div class="metric-local-card local-sini">✅ Siniestradas (Con Res.): <b>{sini_enc} / {total_mesas_local_count}</b></div>
              <div class="metric-local-card local-sini">❌ Siniestradas (Sin Res.): <b>{sini_no_enc} / {total_mesas_local_count}</b></div>
          </div>
      """,
                        unsafe_allow_html=True,
                    )
                elif es_filtro_extraviada:
                    ext_enc = 0
                    ext_no_enc = 0
                    for m in local["Mesas_Reales"]:
                        d_m = cache_mesas.get(m, {})
                        if (
                            d_m.get("estado_acta", "").upper() == "EXTRAVIADA"
                            and d_m.get("estado_mesa", "").upper()
                            in ["INSTALADA", "TARDÍA"]
                        ):
                            if bool(d_m.get("resolucion", "").strip()):
                                ext_enc += 1
                            else:
                                ext_no_enc += 1
                    total_mesas_local_count = ext_enc + ext_no_enc
                    st.markdown(
                        f"""
          <div class="metric-container-local-grid-2">
              <div class="metric-local-card local-ext">✅ Extraviadas (Con Res.): <b>{ext_enc} / {total_mesas_local_count}</b></div>
              <div class="metric-local-card local-ext">❌ Extraviadas (Sin Res.): <b>{ext_no_enc} / {total_mesas_local_count}</b></div>
          </div>
      """,
                        unsafe_allow_html=True,
                    )
                else:
                    total_mesas_local_count = local["total_mesas_real"]

                n_electores_local = 0
                x_electores_local = 0
                for m in local["Mesas_Reales"]:
                    d_m = cache_mesas.get(m, {})
                    try:
                        x_electores_local += int(d_m.get("total_electores", 0) or 0)
                    except ValueError:
                        pass
                    try:
                        n_electores_local += int(d_m.get("electores_votaron", 0) or 0)
                    except ValueError:
                        pass

                etiqueta_mesas_local = (
                    f"({total_mesas_local_count} MESA)"
                    if total_mesas_local_count == 1
                    else f"({total_mesas_local_count} MESAS)"
                )

                col_info, col_btn = st.columns([5, 1])
                with col_info:
                    st.markdown(
                        f"""
            <div class="local-header-box">
                <span style="font-size: 14px; margin-right: 4px;">🏫</span>
                <span class="badge-codigo">[{codigo_local}]</span>
                <span style="font-size: 13px; font-weight: 700; color: #2c3e50; flex: 1; min-width: 180px;">{local['Colegio']} — DISTRITO: {local['Distrito']}</span>
                <span class="badge-mesas">{etiqueta_mesas_local}</span>
                <span class="badge-electores">👥 Votantes: {n_electores_local:,} / {x_electores_local:,}</span>
            </div>
        """,
                        unsafe_allow_html=True,
                    )

                with col_btn:
                    if codigo_activo == codigo_local and not buscando_por_mesa:
                        if st.button(
                            "Cerrar", key=f"cerrar_{codigo_local}", use_container_width=True
                        ):
                            del st.session_state["local_codigo_activo"]
                            st.rerun()
                    elif not buscando_por_mesa:
                        if st.button(
                            "Ver Mesas",
                            key=f"abrir_{codigo_local}",
                            use_container_width=True,
                        ):
                            st.session_state["local_codigo_activo"] = codigo_local
                            st.rerun()

                if codigo_activo == codigo_local:
                    if not (
                        es_filtro_resuelta
                        or es_filtro_observada
                        or es_filtro_siniestrada
                        or es_filtro_extraviada
                    ):
                        if buscando_por_mesa and mesa_busq_input in local["Mesas_Reales"]:
                            m_obj = mesa_busq_input
                            d_m = cache_mesas.get(m_obj, {})
                            est_m = d_m.get("estado_mesa", "").upper()
                            est_a = d_m.get("estado_acta", "").upper()

                            inst_val_local = 1 if est_m in ["INSTALADA", "TARDÍA"] else 0
                            norm_m = (
                                1 if (inst_val_local == 1 and est_a == "NORMAL") else 0
                            )
                            obs_m = (
                                1 if (inst_val_local == 1 and est_a == "OBSERVADA") else 0
                            )
                            sini_m = (
                                1 if (inst_val_local == 1 and est_a == "SINIESTRADA") else 0
                            )
                            ext_m = (
                                1 if (inst_val_local == 1 and est_a == "EXTRAVIADA") else 0
                            )

                            st.markdown(
                                f"""
            <div class="metric-local-inst">🟢 Instaladas en Local: <b>{inst_val_local} / 1</b></div>
            <div class="metric-container-local-grid-4">
                <div class="metric-local-card local-norm">📘 Normales: <b>{norm_m} / 1</b></div>
                <div class="metric-local-card local-obs">📊 Observadas: <b>{obs_m} / 1</b></div>
                <div class="metric-local-card local-sini">🦹‍♂ Siniestradas: <b>{sini_m} / 1</b></div>
                <div class="metric-local-card local-ext">❌ Extraviadas: <b>{ext_m} / 1</b></div>
            </div>
        """,
                                unsafe_allow_html=True,
                            )
                        else:
                            normales_local = instaladas_local = observadas_local = (
                                siniestradas_local
                            ) = extraviadas_local = 0

                            for m in local["Mesas_Reales"]:
                                d_m = cache_mesas.get(m, {})
                                est_m = d_m.get("estado_mesa", "").upper()
                                est_a = d_m.get("estado_acta", "").upper()

                                if est_m in ["INSTALADA", "TARDÍA"]:
                                    instaladas_local += 1
                                    if est_a == "NORMAL":
                                        normales_local += 1
                                    elif est_a == "OBSERVADA":
                                        observadas_local += 1
                                    elif est_a == "SINIESTRADA":
                                        siniestradas_local += 1
                                    elif est_a == "EXTRAVIADA":
                                        extraviadas_local += 1

                            st.markdown(
                                f"""
            <div class="metric-local-inst">🟢 Instaladas en Local: <b>{instaladas_local} / {total_mesas_local_count}</b></div>
            <div class="metric-container-local-grid-4">
                <div class="metric-local-card local-norm">📘 Normales: <b>{normales_local} / {instaladas_local}</b></div>
                <div class="metric-local-card local-obs">📊 Observadas: <b>{observadas_local} / {instaladas_local}</b></div>
                <div class="metric-local-card local-sini">🦹‍♂ Siniestradas: <b>{siniestradas_local} / {instaladas_local}</b></div>
                <div class="metric-local-card local-ext">❌ Extraviadas: <b>{extraviadas_local} / {instaladas_local}</b></div>
            </div>
        """,
                                unsafe_allow_html=True,
                            )

                    st.markdown(
                        "<p style='font-size: 12px; font-weight: 600; margin: 6px 0 4px"
                        " 0;'>📋 Mesas del Local:</p>",
                        unsafe_allow_html=True,
                    )

                    mesas_a_mostrar = local["Mesas"]
                    cols_por_fila = 5
                    for idx_m in range(0, len(mesas_a_mostrar), cols_por_fila):
                        sub_mesas = mesas_a_mostrar[idx_m : idx_m + cols_por_fila]
                        cols_grid = st.columns(cols_por_fila)
                        for sub_i, m_num in enumerate(sub_mesas):
                            with cols_grid[sub_i]:
                                d_mesa = cache_mesas.get(m_num, {})
                                e_mesa = d_mesa.get("estado_mesa", "").upper()
                                e_acta = d_mesa.get("estado_acta", "").upper()
                                ind_sino = d_mesa.get("indicador_sino", "").upper()
                                dig_m = d_mesa.get("estado_digitacion", "").strip()

                                es_res_mesa = ind_sino == "SI"

                                if e_mesa == "NO INSTALADA":
                                    color_borde = "#721C24"
                                    bg_btn = "#f8d7da"
                                elif e_acta == "NORMAL":
                                    color_borde = (
                                        "#27692f" if es_res_mesa else "#17a2b8"
                                    )
                                    bg_btn = "#e8f8f5" if es_res_mesa else "#f2f9fa"
                                elif e_acta == "OBSERVADA":
                                    color_borde = "#ffc107"
                                    bg_btn = "#fefcf0"
                                elif e_acta == "SINIESTRADA":
                                    color_borde = "#fd7e14"
                                    bg_btn = "#fff8f3"
                                elif e_acta == "EXTRAVIADA":
                                    color_borde = "#dc3545"
                                    bg_btn = "#fdf2f2"
                                else:
                                    color_borde = "#6c757d"
                                    bg_btn = "#f8f9fa"

                                icono_estado = (
                                    "✖"
                                    if e_mesa == "NO INSTALADA"
                                    else "✔"
                                    if e_acta == "NORMAL"
                                    else "⚠"
                                    if e_acta == "OBSERVADA"
                                    else "⚡"
                                    if e_acta == "SINIESTRADA"
                                    else "❌"
                                )
                                if es_res_mesa:
                                    icono_estado = "🔒"

                                etiqueta_boton = f"{m_num} {icono_estado}"
                                if dig_m.isdigit() and int(dig_m) > 0:
                                    etiqueta_boton += f" [{int(dig_m):03d}]"

                                if st.button(
                                    etiqueta_boton,
                                    key=f"btn_mesa_{m_num}_{codigo_local}_{idx_m}_{sub_i}",
                                    use_container_width=True,
                                ):
                                    st.session_state["mesa_activa"] = m_num
                                    st.rerun()
