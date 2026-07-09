
import pandas as pd
import streamlit as st

from core.engine import (
    DataControlHubSchemaError,
    process_data_control_hub,
    build_result_excel_from_result,
    send_excel_by_email,
)


# ============================================================
# CONFIGURACIÓN GENERAL
# ============================================================

st.set_page_config(
    page_title="Data Control Hub MVP",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Data Control Hub")
st.caption("MVP para detectar inconsistencias entre personas, activos y artículos, priorizar alertas y gestionar acciones pendientes.")


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:
    st.header("Configuración")

    umbral = st.slider(
        "Umbral costo reparación",
        min_value=0.10,
        max_value=1.00,
        value=0.70,
        step=0.05,
        help="La regla R008 se activa si costo_reparacion > costo_reposicion * umbral."
    )

    st.markdown("---")
    st.subheader("Archivos")

    source_file = st.file_uploader(
        "Subir Excel operativo",
        type=["xlsx"],
        help="Debe contener personas_rrhh, activos_dispositivos y articulos_reparacion_baja."
    )

    previous_file = st.file_uploader(
        "Subir resultado anterior opcional",
        type=["xlsx"],
        help="Sirve para conservar estado_gestion, comentarios y responsables anteriores."
    )

    process_button = st.button("Procesar archivo", type="primary", use_container_width=True)


# ============================================================
# PROCESAMIENTO
# ============================================================

if process_button:
    if source_file is None:
        st.error("Debes subir un Excel operativo antes de procesar.")
    else:
        with st.spinner("Procesando archivo..."):
            try:
                result = process_data_control_hub(
                    source_file=source_file,
                    previous_file=previous_file,
                    umbral_costo_reparacion=umbral,
                    source_filename=source_file.name
                )
                st.session_state["result"] = result
                st.session_state["acciones_editadas"] = result["acciones_pendientes"].copy()
                st.success("Archivo procesado correctamente.")
            except DataControlHubSchemaError as e:
                st.error(str(e))
                if e.schema_issues is not None and not e.schema_issues.empty:
                    st.dataframe(e.schema_issues, use_container_width=True)
            except Exception as e:
                st.exception(e)


# ============================================================
# CONTENIDO PRINCIPAL
# ============================================================

if "result" not in st.session_state:
    st.info("Sube un Excel operativo y presiona **Procesar archivo** para comenzar.")

    st.markdown("""
    ### Flujo del MVP

    1. Subes el Excel operativo mensual.
    2. Opcionalmente subes el resultado anterior para conservar gestión.
    3. El sistema valida datos, cruza fuentes y genera alertas.
    4. Revisas acciones pendientes.
    5. Editas estados, responsables y comentarios.
    6. Descargas o envías el resultado.
    """)

else:
    result = st.session_state["result"]

    tab_resumen, tab_validacion, tab_acciones, tab_historial, tab_reglas, tab_exportar = st.tabs([
        "Resumen ejecutivo",
        "Validación de datos",
        "Acciones pendientes",
        "Historial",
        "Reglas",
        "Exportar / Enviar"
    ])

    # --------------------------------------------------------
    # RESUMEN
    # --------------------------------------------------------
    with tab_resumen:
        st.subheader("Resumen ejecutivo")

        dashboard = result["dashboard_resumen"]

        def metric_value(name):
            row = dashboard[dashboard["indicador"] == name]
            if row.empty:
                return 0
            return row.iloc[0]["valor"]

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Alertas vigentes", metric_value("Alertas vigentes"))
        c2.metric("Críticas", metric_value("Alertas críticas"))
        c3.metric("Altas", metric_value("Alertas altas"))
        c4.metric("Vencidas", metric_value("Alertas vencidas"))
        c5.metric("Nuevas", metric_value("Nuevas detecciones"))

        st.markdown("#### Indicadores")
        st.dataframe(dashboard, use_container_width=True, hide_index=True)

        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown("#### Alertas por regla")
            st.dataframe(result["alertas_por_regla"], use_container_width=True, hide_index=True)
        with col_b:
            st.markdown("#### Alertas por estado de gestión")
            st.dataframe(result["alertas_por_estado_gestion"], use_container_width=True, hide_index=True)

    # --------------------------------------------------------
    # VALIDACIÓN
    # --------------------------------------------------------
    with tab_validacion:
        st.subheader("Validación de datos")
        st.write("Esta sección muestra problemas de calidad que pueden afectar la lectura del modelo.")

        validacion = result["validacion_datos"]

        if "severidad" in validacion.columns:
            severidades = ["Todas"] + sorted(validacion["severidad"].dropna().astype(str).unique().tolist())
            filtro_sev = st.selectbox("Filtrar por severidad", severidades)
            if filtro_sev != "Todas":
                validacion = validacion[validacion["severidad"].astype(str) == filtro_sev]

        st.dataframe(validacion, use_container_width=True, hide_index=True)

    # --------------------------------------------------------
    # ACCIONES
    # --------------------------------------------------------
    with tab_acciones:
        st.subheader("Acciones pendientes")
        st.write("Esta es la bandeja principal de gestión. Puedes editar estados, responsables, fechas y comentarios.")

        acciones = st.session_state["acciones_editadas"].copy()

        col1, col2, col3, col4 = st.columns(4)

        with col1:
            filtro_criticidad = st.selectbox(
                "Criticidad",
                ["Todas"] + sorted(acciones["criticidad"].dropna().astype(str).unique().tolist())
            )
        with col2:
            filtro_estado = st.selectbox(
                "Estado gestión",
                ["Todos"] + sorted(acciones["estado_gestion"].dropna().astype(str).unique().tolist())
            )
        with col3:
            filtro_regla = st.selectbox(
                "Regla",
                ["Todas"] + sorted(acciones["codigo_regla"].dropna().astype(str).unique().tolist())
            )
        with col4:
            filtro_vencida = st.selectbox(
                "Vencida",
                ["Todas", "Sí", "No"]
            )

        acciones_filtradas = acciones.copy()
        if filtro_criticidad != "Todas":
            acciones_filtradas = acciones_filtradas[acciones_filtradas["criticidad"].astype(str) == filtro_criticidad]
        if filtro_estado != "Todos":
            acciones_filtradas = acciones_filtradas[acciones_filtradas["estado_gestion"].astype(str) == filtro_estado]
        if filtro_regla != "Todas":
            acciones_filtradas = acciones_filtradas[acciones_filtradas["codigo_regla"].astype(str) == filtro_regla]
        if filtro_vencida != "Todas":
            acciones_filtradas = acciones_filtradas[acciones_filtradas["alerta_vencida"].astype(str) == filtro_vencida]

        st.caption(f"Mostrando {len(acciones_filtradas)} de {len(acciones)} acciones.")

        editable_cols = [
            "estado_gestion",
            "responsable_manual",
            "fecha_compromiso",
            "fecha_cierre",
            "decision_final",
            "comentario_manual",
            "evidencia",
            "usuario_gestion",
        ]

        edited = st.data_editor(
            acciones_filtradas,
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            column_config={
                "estado_gestion": st.column_config.SelectboxColumn(
                    "estado_gestion",
                    options=["Nueva", "En revisión", "En gestión", "Bloqueada", "Resuelta", "Descartada", "Reabierta"],
                    required=False
                ),
            },
            disabled=[c for c in acciones_filtradas.columns if c not in editable_cols],
            key="acciones_editor"
        )

        if st.button("Guardar cambios en sesión", use_container_width=True):
            # Actualiza solo las filas visibles/editadas, usando id_alerta como llave.
            acciones_base = st.session_state["acciones_editadas"].copy()
            edited_index = edited.set_index("id_alerta")
            acciones_base = acciones_base.set_index("id_alerta")

            for idx in edited_index.index:
                for col in editable_cols:
                    if col in acciones_base.columns and col in edited_index.columns:
                        acciones_base.loc[idx, col] = edited_index.loc[idx, col]

            st.session_state["acciones_editadas"] = acciones_base.reset_index()
            st.success("Cambios guardados en la sesión. Recuerda descargar el Excel actualizado.")

    # --------------------------------------------------------
    # HISTORIAL
    # --------------------------------------------------------
    with tab_historial:
        st.subheader("Historial de alertas")
        st.write("Permite revisar alertas vigentes, persistentes o no vigentes.")
        st.dataframe(result["historial_alertas"], use_container_width=True, hide_index=True)

    # --------------------------------------------------------
    # REGLAS
    # --------------------------------------------------------
    with tab_reglas:
        st.subheader("Reglas aplicadas")
        st.write("Matriz oficial de reglas del MVP.")
        st.dataframe(result["reglas_aplicadas"], use_container_width=True, hide_index=True)

    # --------------------------------------------------------
    # EXPORTAR / ENVIAR
    # --------------------------------------------------------
    with tab_exportar:
        st.subheader("Exportar o enviar resultado")

        acciones_finales = st.session_state.get("acciones_editadas", result["acciones_pendientes"])
        excel_bytes = build_result_excel_from_result(result, acciones_pendientes_editadas=acciones_finales)
        filename = "Data_Control_Hub_resultado_streamlit_v1.xlsx"

        st.download_button(
            label="Descargar Excel resultado",
            data=excel_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

        st.markdown("---")
        st.markdown("### Enviar por correo")

        with st.expander("Configurar envío por correo"):
            correo_remitente = st.text_input("Correo remitente")
            correo_destinatario = st.text_input("Correo destinatario")
            password_app = st.text_input("Contraseña de aplicación / SMTP", type="password")
            smtp_server = st.text_input("Servidor SMTP", value="smtp.gmail.com")
            smtp_port = st.number_input("Puerto SMTP", value=587, step=1)

            asunto = st.text_input("Asunto", value="Data Control Hub - Resultado de alertas")
            cuerpo = st.text_area(
                "Mensaje",
                value="Hola,\n\nSe adjunta el resultado generado por Data Control Hub.\n\nSaludos."
            )

            if st.button("Enviar Excel por correo", use_container_width=True):
                if not correo_remitente or not correo_destinatario or not password_app:
                    st.error("Debes completar remitente, destinatario y contraseña de aplicación.")
                else:
                    try:
                        send_excel_by_email(
                            excel_bytes=excel_bytes,
                            filename=filename,
                            correo_remitente=correo_remitente,
                            password_app=password_app,
                            correo_destinatario=correo_destinatario,
                            asunto=asunto,
                            cuerpo=cuerpo,
                            smtp_server=smtp_server,
                            smtp_port=int(smtp_port)
                        )
                        st.success(f"Correo enviado correctamente a {correo_destinatario}.")
                    except Exception as e:
                        st.error("No se pudo enviar el correo. El Excel fue generado correctamente y puedes descargarlo.")
                        st.exception(e)
