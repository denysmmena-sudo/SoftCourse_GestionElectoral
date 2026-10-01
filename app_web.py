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

# Estilos CSS optimizados
st.markdown(
    """
    <style>
    .main { background-color: #E5ECEA; }

    .header-container {
        display: flex;
        justify-content: space-between;
        align-items: center;
        background-color: #ffffff;
        padding: 6px 12px;
        border-radius: 6px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        margin-bottom: 4px;
    }
    .header-title {
        font-size: 15px;
        font-weight: 700;
        color: #31333F;
        margin: 0;
    }
    .header-badge-inst {
        background-color: #f1f8f5;
        border-left: 3px solid #28a745;
        padding: 4px 8px;
        border-radius: 4px;
        font-weight: 600;
        color: #31333F;
        font-size: 11px;
    }

    .metric-container-grid-4 {
        display: grid;
        grid-template-columns: repeat(4, 1fr);
        gap: 4px;
        margin-bottom: 6px;
    }

    .metric-card {
        background-color: #f8f9fa;
        padding: 4px 6px;
        border-radius: 4px;
        font-weight: 600;
        color: #31333F;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        font-size: 10px;
    }

    .card-local-norm { border-left: 3px solid #17a2b8; }
    .card-local-obs { border-left: 3px solid #ffc107; }
    .card-local-sini { border-left: 3px solid #fd7e14; }
    .card-local-ext { border-left: 3px solid #dc3545; }

    .metric-local-inst {
        background-color: #f1f8f5;
        padding: 4px 6px;
        border-radius: 4px;
        font-weight: 600;
        color: #31333F;
        font-size: 10px;
        border-left: 3px solid #28a745;
        margin-bottom: 3px;
        display: inline-block;
    }

    .metric-container-local-grid-4 {
        display: grid;
        grid-template-columns: repeat(4, 1fr);
        gap: 4px;
        margin-bottom: 4px;
    }

    .metric-local-card {
        padding: 3px 4px;
        border-radius: 4px;
        font-size: 9px;
        font-weight: 600;
        color: #31333F;
    }

    .local-norm { background-color: #f2f9fa; border-left: 2px solid #17a2b8; }
    .local-obs { background-color: #fefcf0; border-left: 2px solid #ffc107; }
    .local-sini { background-color: #fff8f3; border-left: 2px solid #fd7e14; }
    .local-ext { background-color: #fdf2f2; border-left: 2px solid #dc3545; }

    .badge-codigo {
        background-color: #e8f4fd;
        color: #0056b3;
        padding: 2px 6px;
        border-radius: 4px;
        font-size: 12px;
        font-weight: 700;
        margin-right: 6px;
    }
    .badge-mesas {
        background-color: #f1f3f5;
        color: #495057;
        padding: 2px 6px;
        border-radius: 4px;
        font-size: 11px;
        font-weight: 600;
        text-transform: uppercase;
        margin-left: 6px;
    }
    .badge-electores {
        background-color: #e2f0d9;
        color: #27692f;
        padding: 2px 6px;
        border-radius: 4px;
        font-size: 11px;
        font-weight: 700;
        margin-left: 6px;
        border-left: 3px solid #28a745;
    }

    .floating-window {
        background-color: #ffffff;
        border: 1px solid #dcdcdc;
        border-radius: 6px;
        padding: 6px 8px;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.05);
        margin-top: 4px;
        margin-bottom: 4px;
    }

    .stTextInput, .stSelectbox {
        margin-bottom: -12px !important;
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

# Sidebar con branding SoftCourse y opciones solicitadas
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
        <div style="background-color: #fff3cd; border: 1px solid #ffeeba; padding: 8px; border-radius: 6px; margin-top: 6px; font-size: 11px; color: #856404;">
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

# Detectar cambio de filtro para limpiar estados activos
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

  # Mapeo rápido de mesas con sus respectivos locales y distritos
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
      actas_ingresadas_lista.append(
          (int(dig_val), mesa_id, info, extras)
      )

  # Ordenar por el número de orden asignado
  actas_ingresadas_lista.sort(key=lambda x: x[0])

  if not actas_ingresadas_lista:
    st.info(
        "Aún no hay actas marcadas como 'Acta Ingresada' en el sistema."
    )
  else:
    col_inf_1, col_inf_2 = st.columns([3, 1])
    with col_inf_1:
      st.markdown(
          f"**Total de actas ingresadas hasta el momento:**"
          f" {len(actas_ingresadas_lista)} / 830"
      )

    with col_inf_2:
      # Preparar DataFrame completo con TODOS los detalles actualizados de cada mesa
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
        item_copia["Mesas"] = mesas_filtradas
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
        item_copia["Mesas"] = mesas_filtradas
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

    # Si hay una mesa activa para editar o visualizar
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

      if st.button("⬅️ Volver a las mesas"):
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
                <div style="font-size: 14px; font-weight: 700; color: {'#27692f' if modo_solo_lectura else '#1f77b4'}; border-bottom: 1px solid #eaeaea; padding-bottom: 4px; margin-bottom: 6px;">
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
              "<p style='font-size:10px; color:#721C24; margin:0;'>ℹ No aplica"
              " acta</p>",
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
        ya_ingresada_inicial = (
            dig_actual.isdigit() and int(dig_val := dig_actual) > 0
        )

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
              f"<p style='font-size:10px; color:#0056b3; margin:0;'>📌 Orden"
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
        st.markdown("<hr style='margin: 4px 0;'>", unsafe_allow_html=True)
        st.markdown(
            "<p style='font-size: 11px; font-weight: 600; margin: 0 0 2px"
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

        rc1, rc2, rc3, rc4 = st.columns([1, 1, 1, 0.8])
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

      st.markdown("<div style='margin-top: 4px;'></div>", unsafe_allow_html=True)

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
              "⚠️ **Atención:** Has desmarcado 'Acta Ingresada'. Si guardas,"
              " perderá su posición asignada."
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
                t_electores = int(
                    d_prev.get("total_electores", 0) or 0
                )
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
      # --- CÁLCULO DE ELECTORES GLOBALES (n / x) ---
      total_locales_encontrados = len(resultados_filtrados)
      total_mesas_reales_global = sum(
          local["total_mesas_real"] for local in resultados_filtrados
      )

      meses_instaladas_global = 0
      normales_global = 0
      observadas_global = 0
      siniestradas_global = 0
      extraviadas_global = 0

      n_electores_global = 0
      x_electores_global = 0

      for local in resultados_filtrados:
        for m in local["Mesas_Reales"]:
          d_m = cache_mesas.get(m, {})
          est_m = d_m.get("estado_mesa", "").upper()
          est_a = d_m.get("estado_acta", "").upper()

          try:
            x_electores_global += int(d_m.get("total_electores", 0) or 0)
          except ValueError:
            pass

          try:
            n_electores_global += int(d_m.get("electores_votaron", 0) or 0)
          except ValueError:
            pass

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

      if codigo_activo and not buscando_por_mesa:
        if st.button("⬅ Volver a la lista de locales"):
          del st.session_state["local_codigo_activo"]
          st.rerun()

      st.markdown(
          f"""
        <div class="header-container">
            <h3 class="header-title">📍 Locales Encontrados ({total_locales_encontrados})</h3>
            <div style="display: flex; gap: 8px; align-items: center;">
                <div class="header-badge-inst">
                    👥 Electores Globales: <b>{n_electores_global:,} / {x_electores_global:,}</b>
                </div>
                <div class="header-badge-inst">
                    🟢 Mesas Instaladas Globales: <b>{meses_instaladas_global} / {total_mesas_reales_global}</b>
                </div>
            </div>
        </div>
    """,
          unsafe_allow_html=True,
      )

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

        col_info, col_btn = st.columns([5, 1])
        with col_info:
          st.markdown(
              f"""
                <div style="display: flex; align-items: center; background-color: #ffffff; padding: 6px 10px; border-radius: 6px; border: 1px solid #e0e0e0; margin-bottom: 4px;">
                    <span style="font-size: 14px; margin-right: 6px;">🏫</span>
                    <span class="badge-codigo">[{codigo_local}]</span>
                    <span style="font-size: 13px; font-weight: 700; color: #2c3e50; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">{local['Colegio']} — DISTRITO: {local['Distrito']}</span>
                    <span class="badge-mesas">({total_mesas_local_count} MESAS)</span>
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
                "Ver Mesas", key=f"abrir_{codigo_local}", use_container_width=True
            ):
              st.session_state["local_codigo_activo"] = codigo_local
              st.rerun()

        if codigo_activo == codigo_local:
          if buscando_por_mesa and mesa_busq_input in local["Mesas_Reales"]:
            m_obj = mesa_busq_input
            d_m = cache_mesas.get(m_obj, {})
            est_m = d_m.get("estado_mesa", "").upper()
            est_a = d_m.get("estado_acta", "").upper()

            inst_val_local = 1 if est_m in ["INSTALADA", "TARDÍA"] else 0
            norm_m = 1 if (inst_val_local == 1 and est_a == "NORMAL") else 0
            obs_m = 1 if (inst_val_local == 1 and est_a == "OBSERVADA") else 0
            sini_m = (
                1 if (inst_val_local == 1 and est_a == "SINIESTRADA") else 0
            )
            ext_m = 1 if (inst_val_local == 1 and est_a == "EXTRAVIADA") else 0

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
            normales_local = (
                instaladas_local
            ) = (
                observadas_local
            ) = siniestradas_local = extraviadas_local = 0

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
                    <div class="metric-local-card local-norm">📘 Normales: <b>{normales_local} / {instaladas_local if instaladas_local > 0 else 1}</b></div>
                    <div class="metric-local-card local-obs">📊 Observadas: <b>{observadas_local} / {instaladas_local if instaladas_local > 0 else 1}</b></div>
                    <div class="metric-local-card local-sini">🦹‍♂ Siniestradas: <b>{siniestradas_local} / {instaladas_local if instaladas_local > 0 else 1}</b></div>
                    <div class="metric-local-card local-ext">❌ Extraviadas: <b>{extraviadas_local} / {instaladas_local if instaladas_local > 0 else 1}</b></div>
                </div>
            """,
                unsafe_allow_html=True,
            )

          st.markdown("#### 🗳 Mesas Electorales")

          cols = st.columns(6)
          for idx, mesa in enumerate(local["Mesas"]):
            datos_m = cache_mesas.get(mesa, {})
            est_acta = datos_m.get("estado_acta", "NORMAL").upper()
            est_mesa = datos_m.get("estado_mesa", "INSTALADA").upper()

            votos_mesa = datos_m.get("electores_votaron", "0")
            total_el_mesa = datos_m.get("total_electores", "0")

            prefix = "[T] " if est_mesa == "TARDÍA" else ""

            bg, txt, brd = "#FFFFFF", "#31333F", "#d6d6d6"
            if est_mesa == "NO INSTALADA":
              bg, txt, brd = "#F5C6CB", "#721C24", "#F5C6CB"
              label_estado = "NO INSTALADA"
            elif est_acta in ["EXTRAVIADA", "SINIESTRADA"]:
              bg, txt, brd = "#FFD8A8", "#D9480F", "#FFD8A8"
              label_estado = est_acta
            elif est_acta == "OBSERVADA":
              bg, txt, brd = "#FFEEBA", "#856404", "#FFEEBA"
              label_estado = est_acta
            elif est_mesa == "TARDÍA":
              bg, txt, brd = "#FFF3CD", "#856404", "#FFEEBA"
              label_estado = est_acta if est_acta else "NORMAL"
            else:
              label_estado = est_acta

            with cols[idx % 6]:
              st.markdown(
                  f"""
                        <div style="background-color: {bg}; color: {txt}; border: 1px solid {brd}; border-radius: 4px; padding: 4px 2px; text-align: center; font-weight: 600; font-size: 10px; margin-bottom: 2px;">
                            {prefix}Mesa {mesa}<br>
                            <span style="font-size: 7px; font-weight: normal;">{label_estado}</span><br>
                            <span style="font-size: 8px; font-weight: bold; color: #1e7e34;">👤 {votos_mesa}/{total_el_mesa}</span>
                        </div>
                    """,
                  unsafe_allow_html=True,
              )

              if st.button(
                  "Editar",
                  key=f"edit_m_{codigo_local}_{mesa}",
                  use_container_width=True,
              ):
                st.session_state["mesa_activa"] = mesa
                st.rerun()