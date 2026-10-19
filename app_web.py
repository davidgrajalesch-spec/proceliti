import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta, time
import io

try:
    import msal
    import requests
except ImportError:
    os.system("pip install msal requests")
    import msal
    import requests

try:
    import docx
    from docx import Document
except ImportError:
    os.system("pip install python-docx")
    import docx
    from docx import Document

st.set_page_config(page_title="ProceLiti - Control de Comparendos D12", page_icon="⚖️", layout="wide")

ARCHIVO_EXCEL = 'TablaProcesos.xlsx'

# ==============================================================================
# FUNCIONES DE INTEGRACIÓN CON MICROSOFT GRAPH / ONEDRIVE
# ==============================================================================
def obtener_token():
    try:
        tenant_id = st.secrets["AZURE_TENANT_ID"]
        client_id = st.secrets["AZURE_CLIENT_ID"]
        client_secret = st.secrets["AZURE_CLIENT_SECRET"]
        
        authority = f"https://login.microsoftonline.com/{tenant_id}"
        app = msal.ConfidentialClientApplication(
            client_id,
            authority=authority,
            client_credential=client_secret,
        )
        result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        if "access_token" in result:
            return result["access_token"]
        return None
    except Exception:
        return None

def cargar_datos():
    columnas_requeridas = [
        'Numero_Comparendo', 'Cedula_Ciudadano', 'Nombre_Ciudadano', 'Placa_Vehiculo', 
        'Fecha_Infraccion', 'Nombre_Agente', 'Abogado_Impugnacion', 'Fecha_Version_Libre', 
        'Version_Conductor', 'Pruebas_Decretadas_E1', 'Abogado_Pruebas', 'Fecha_Pruebas', 
        'Errores_Diligenciamiento', 'Errores_Procedimiento', 'Hubo_Pago', 'Menciona_No_Pago',
        'Precedentes_Favorables', 'Nota_Adicional', 'Abogado_Fallo', 'Fecha_Fallo', 
        'Sentido_Fallo', 'Se_Presenta_Apelacion', '__PowerAppsId__'
    ]

    token = obtener_token()
    df = None

    if token:
        user_email = st.secrets["ONEDRIVE_USER_EMAIL"]
        # Busca en la raíz de OneDrive. Si lo ubicas dentro de la subcarpeta especificada,
        # la URL puede cambiarse a: .../root:/1.%20David%20Grajales/Prueba%20Poweapps/TablaProcesos.xlsx:/content
        url = f"https://graph.microsoft.com/v1.0/users/{user_email}/drive/root:/TablaProcesos.xlsx:/content"
        headers = {"Authorization": f"Bearer {token}"}
        
        try:
            response = requests.get(url, headers=headers)
            if response.status_code == 200:
                df = pd.read_excel(io.BytesIO(response.content), dtype=str)
        except Exception:
            df = None

    # Respaldo local si no conecta a la API o si falla el token
    if df is None:
        if not os.path.exists(ARCHIVO_EXCEL):
            df_vacio = pd.DataFrame(columns=columnas_requeridas)
            df_vacio.to_excel(ARCHIVO_EXCEL, index=False)
            return df_vacio
        df = pd.read_excel(ARCHIVO_EXCEL, dtype=str)

    for col in columnas_requeridas:
        if col not in df.columns:
            df[col] = ""
            
    return df.fillna("")

def guardar_datos(df):
    token = obtener_token()
    guardado_en_cloud = False

    if token:
        user_email = st.secrets["ONEDRIVE_USER_EMAIL"]
        url = f"https://graph.microsoft.com/v1.0/users/{user_email}/drive/root:/TablaProcesos.xlsx:/content"
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        }
        
        output = io.BytesIO()
        df.to_excel(output, index=False)
        output.seek(0)
        
        try:
            response = requests.put(url, headers=headers, data=output.getvalue())
            if response.status_code in [200, 201]:
                guardado_en_cloud = True
        except Exception:
            guardado_en_cloud = False

    # Guardado local de respaldo
    df.to_excel(ARCHIVO_EXCEL, index=False)


# ==============================================================================
# CARGA INICIAL Y ESTADOS DE SESIÓN
# ==============================================================================
df = cargar_datos()

if 'ficha_activa' not in st.session_state:
    st.session_state.ficha_activa = None
if 'editando_ficha' not in st.session_state:
    st.session_state.editando_ficha = False
if 'vista_actual' not in st.session_state:
    st.session_state.vista_actual = "Master Procesos"

st.markdown("<h2 style='text-align: center; color: #ff5722;'>⚖️ PROCELITI - GESTIÓN JURÍDICA DE COMPARENDOS D12</h2>", unsafe_allow_html=True)
st.markdown("---")

opciones_menu = ["Master Procesos", "Asignación de Audiencias", "Registrar Nuevo Proceso"]
idx_menu = 0
if st.session_state.vista_actual in opciones_menu:
    idx_menu = opciones_menu.index(st.session_state.vista_actual)

menu_seleccionado = st.sidebar.selectbox("Menú de Navegación", opciones_menu, index=idx_menu)
if menu_seleccionado != st.session_state.vista_actual and st.session_state.vista_actual != "Ficha Control de audiencia":
    st.session_state.vista_actual = menu_seleccionado

ojo_svg = """<svg xmlns="http://www.w3.org/2000/svg" width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#ff5722" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="vertical-align: middle;"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path><circle cx="12" cy="12" r="3"></circle></svg>"""

# ==============================================================================
# 1. MÓDULO: MASTER PROCESOS
# ==============================================================================
if st.session_state.vista_actual == "Master Procesos":
    st.subheader("📊 Master Procesos - Listado General")
    
    busqueda = st.text_input("🔍 Buscar por número de comparendo, placa o cédula:")
    
    if busqueda:
        resultado = df[
            df['Numero_Comparendo'].str.contains(busqueda, case=False, na=False) |
            df['Placa_Vehiculo'].str.contains(busqueda, case=False, na=False) |
            df['Cedula_Ciudadano'].str.contains(busqueda, case=False, na=False)
        ]
    else:
        resultado = df
        
    st.markdown(f"### Seleccione expediente haciendo clic en el ícono {ojo_svg}:", unsafe_allow_html=True)
    
    cols_tabla = [
        'Numero_Comparendo', 'Cedula_Ciudadano', 'Nombre_Ciudadano', 
        'Placa_Vehiculo', 'Fecha_Infraccion', 'Nombre_Agente', 
        'Abogado_Impugnacion', 'Fecha_Version_Libre', 'Fecha_Pruebas', 'Fecha_Fallo'
    ]
    
    for c in cols_tabla:
        if c not in resultado.columns:
            resultado[c] = ""

    header_cols = st.columns([1, 2, 2, 3, 2, 2, 2])
    header_cols[0].markdown("**Ver Ficha**")
    header_cols[1].markdown("**Comparendo**")
    header_cols[2].markdown("**Cédula**")
    header_cols[3].markdown("**Nombre**")
    header_cols[4].markdown("**Placa**")
    header_cols[5].markdown("**Fecha Infr.**")
    header_cols[6].markdown("**Abogado Imp.**")
    st.markdown("---")

    for idx, row in resultado.iterrows():
        c_cols = st.columns([1, 2, 2, 3, 2, 2, 2])
        with c_cols[0]:
            if st.button("👁️", key=f"ojo_{row['Numero_Comparendo']}_{idx}"):
                st.session_state.ficha_activa = row['Numero_Comparendo']
                st.session_state.editando_ficha = False
                st.session_state.vista_actual = "Ficha Control de audiencia"
                st.rerun()
        c_cols[1].text(row['Numero_Comparendo'])
        c_cols[2].text(row['Cedula_Ciudadano'])
        c_cols[3].text(row['Nombre_Ciudadano'])
        c_cols[4].text(row['Placa_Vehiculo'])
        c_cols[5].text(row['Fecha_Infraccion'])
        c_cols[6].text(row['Abogado_Impugnacion'])

    st.markdown("---")
    limite = st.selectbox("Ver resultados:", [10, 50, 100], index=0, key="limite_master")

# ==============================================================================
# 2. MÓDULO: ASIGNACIÓN DE AUDIENCIAS
# ==============================================================================
elif st.session_state.vista_actual == "Asignación de Audiencias":
    st.subheader("📅 Asignación de Audiencias y Cronograma Semanal")
    
    st.markdown("#### Tabla General Ordenada Cronológicamente por Fecha de Impugnación")
    cols_asig = ['Numero_Comparendo', 'Cedula_Ciudadano', 'Nombre_Ciudadano', 'Placa_Vehiculo', 'Fecha_Version_Libre', 'Fecha_Pruebas', 'Fecha_Fallo']
    
    df_ordenado = df.sort_values(by='Fecha_Version_Libre', ascending=True)
    limite_asig = st.selectbox("Ver resultados en tabla general:", [10, 50, 100], index=0, key="lim_asig")
    st.dataframe(df_ordenado[cols_asig].head(limite_asig), use_container_width=True)
    
    st.markdown("---")
    st.markdown("#### Selección de Semana a Asignar")
    
    hoy = datetime.now()
    inicio_lunes = hoy - timedelta(days=hoy.weekday())
    fin_viernes = inicio_lunes + timedelta(days=4)
    
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        f_inicio = st.date_input("Fecha Inicio Semana", value=inicio_lunes)
    with col_f2:
        f_fin = st.date_input("Fecha Fin Semana", value=fin_viernes)
        
    if st.button("📥 Filtrar y Asignar Semana"):
        st.session_state['f_inicio'] = f_inicio.strftime("%Y-%m-%d")
        st.session_state['f_fin'] = f_fin.strftime("%Y-%m-%d")
        
    if 'f_inicio' in st.session_state:
        fi = st.session_state['f_inicio']
        ff = st.session_state['f_fin']
        st.info(f"Mostrando audiencias programadas (Impugnación, Pruebas y Fallo) entre {fi} y {ff}")
        
        mask_imp = (df['Fecha_Version_Libre'] >= fi) & (df['Fecha_Version_Libre'] <= ff)
        mask_pru = (df['Fecha_Pruebas'] >= fi) & (df['Fecha_Pruebas'] <= ff)
        mask_fal = (df['Fecha_Fallo'] >= fi) & (df['Fecha_Fallo'] <= ff)
        
        df_semana = df[mask_imp | mask_pru | mask_fal]
        
        if len(df_semana) > 0:
            st.markdown("### Tabla de Asignación Semanal (Todas las Etapas)")
            abogados_lista = [
                "David Grajales", "Diego Pachon", "Jose Cortes", 
                "Vivian Oliveros", "Alejandra Robles", "Ingrid Cifuentes", 
                "Daniela Oliveros", "Otro"
            ]
            
            with st.form("form_asignacion_semanal"):
                actualizaciones = {}
                for idx, row in df_semana.iterrows():
                    tipo_aud = []
                    fecha_aud_val = ""
                    if fi <= str(row['Fecha_Version_Libre']) <= ff:
                        tipo_aud.append("Impugnación")
                        fecha_aud_val = row['Fecha_Version_Libre']
                    if fi <= str(row['Fecha_Pruebas']) <= ff:
                        tipo_aud.append("Pruebas")
                        fecha_aud_val = row['Fecha_Pruebas']
                    if fi <= str(row['Fecha_Fallo']) <= ff:
                        tipo_aud.append("Fallo")
                        fecha_aud_val = row['Fecha_Fallo']
                        
                    tipo_str = " / ".join(tipo_aud)
                    
                    st.markdown(f"**Comparendo:** {row['Numero_Comparendo']} | **Ciudadano:** {row['Nombre_Ciudadano']} | **Placa:** {row['Placa_Vehiculo']} | **Etapa(s):** {tipo_str}")
                    
                    col_s1, col_s2, col_s3 = st.columns(3)
                    with col_s1:
                        st.text(f"Audiencia: {tipo_str}")
                    with col_s2:
                        st.text(f"Fecha/Hora: {fecha_aud_val}")
                    with col_s3:
                        col_abogado_etapa = 'Abogado_Impugnacion'
                        if "Pruebas" in tipo_str and row['Abogado_Pruebas'] != "":
                            col_abogado_etapa = 'Abogado_Pruebas'
                        elif "Fallo" in tipo_str and row['Abogado_Fallo'] != "":
                            col_abogado_etapa = 'Abogado_Fallo'

                        val_actual = row[col_abogado_etapa]
                        if val_actual in abogados_lista:
                            idx_def = abogados_lista.index(val_actual)
                        elif val_actual != "":
                            idx_def = abogados_lista.index("Otro")
                        else:
                            idx_def = 0
                        
                        abv_sel = st.selectbox(f"Abogado Asignado ({row['Numero_Comparendo']})", abogados_lista, index=idx_def, key=f"abv_sel_{idx}")
                        
                        abv_otro = ""
                        if abv_sel == "Otro":
                            default_val_otro = val_actual if val_actual not in abogados_lista else ""
                            abv_otro = st.text_input(f"Escriba nombre del abogado (Obligatorio)", value=default_val_otro, key=f"otro_input_{idx}")
                            
                        actualizaciones[idx] = (col_abogado_etapa, abv_otro if abv_sel == "Otro" else abv_sel)
                    st.markdown("---")
                    
                col_btn1, col_btn2 = st.columns(2)
                with col_btn1:
                    btn_asignar = st.form_submit_button("Asignar")
                with col_btn2:
                    btn_modificar = st.form_submit_button("Modificar")
                    
                if btn_asignar or btn_modificar:
                    for idx, (col_etapa, nuevo_abv) in actualizaciones.items():
                        df.loc[idx, col_etapa] = nuevo_abv
                    guardar_datos(df)
                    st.success("✅ ¡Asignaciones actualizadas y guardadas en el Master exitosamente!")
        else:
            st.warning("No hay audiencias registradas (ni de impugnación, ni de pruebas, ni de fallo) en el rango de fechas seleccionado.")

# ==============================================================================
# 3. MÓDULO: REGISTRAR NUEVO PROCESO
# ==============================================================================
elif st.session_state.vista_actual == "Registrar Nuevo Proceso":
    st.subheader("➕ Registrar Nuevo Proceso")
    
    with st.form("form_nuevo_proceso"):
        c1, c2 = st.columns(2)
        with c1:
            num_comp = st.text_input("Número de Comparendo")
            cedula = st.text_input("Cédula del Infractor")
            nombre = st.text_input("Nombre Completo S.C.")
        with c2:
            placa = st.text_input("Placa del Vehículo")
            fecha_inf = st.date_input("Fecha de Presunta Infracción", value=datetime.now())
            agente = st.text_input("Nombre del Agente de Tránsito")
            
        st.markdown("**Fecha y Hora Audiencia de Impugnación (Versión Libre)**")
        col_fi1, col_fi2 = st.columns(2)
        with col_fi1:
            f_imp_date = st.date_input("Fecha Impugnación", value=datetime.now())
        with col_fi2:
            f_imp_time = st.time_input("Hora Impugnación", value=time(10, 0))
            
        submitted = st.form_submit_button("Guardar en el Master y Generar Ficha")
        if submitted and num_comp:
            fecha_hora_str = f"{f_imp_date.strftime('%Y-%m-%d')} {f_imp_time.strftime('%H:%M')}"
            nuevo = {
                'Numero_Comparendo': str(num_comp), 
                'Cedula_Ciudadano': str(cedula),
                'Nombre_Ciudadano': str(nombre).upper(), 
                'Placa_Vehiculo': str(placa.upper()),
                'Fecha_Infraccion': fecha_inf.strftime('%d/%m/%Y'), 
                'Nombre_Agente': str(agente).upper(),
                'Abogado_Impugnacion': "", 
                'Fecha_Version_Libre': str(fecha_hora_str),
                'Version_Conductor': "", 
                'Pruebas_Decretadas_E1': "", 
                'Abogado_Pruebas': "", 
                'Fecha_Pruebas': "",
                'Errores_Diligenciamiento': "", 
                'Errores_Procedimiento': "", 
                'Hubo_Pago': "", 
                'Menciona_No_Pago': "",
                'Precedentes_Favorables': "", 
                'Nota_Adicional': "", 
                'Abogado_Fallo': "", 
                'Fecha_Fallo': "",
                'Sentido_Fallo': "", 
                'Se_Presenta_Apelacion': "",
                '__PowerAppsId__': str(os.urandom(16).hex())
            }
            df = pd.concat([df, pd.DataFrame([nuevo])], ignore_index=True)
            guardar_datos(df)
            st.success(f"¡Proceso {num_comp} registrado en el Master exitosamente!")

# ==============================================================================
# 4. MÓDULO: FICHA CONTROL DE AUDIENCIA (Acceso exclusivo por ícono de Ojo)
# ==============================================================================
elif st.session_state.vista_actual == "Ficha Control de audiencia":
    st.subheader("📝 FICHA CONTROL AUDIENCIAS PROCESO CONTRAVENCIONAL")
    
    if st.button("⬅️ Volver a Master Procesos"):
        st.session_state.vista_actual = "Master Procesos"
        st.rerun()

    if st.session_state.ficha_activa is None:
        st.info("⚠️ Por favor, seleccione un caso haciendo clic en el ícono de Ojo (👁️) en el módulo **Master Procesos**.")
    else:
        comp_activo = st.session_state.ficha_activa
        idx_caso = df[df['Numero_Comparendo'] == comp_activo].index
        
        if len(idx_caso) == 0:
            st.error("No se encontró el registro activo.")
        else:
            i = idx_caso[0]
            r = df.loc[i]
            disabled_state = not st.session_state.editando_ficha
            
            st.markdown(f"### Expediente Comparendo: {r['Numero_Comparendo']}")
            
            with st.form("form_ficha_completa"):
                # (Primer Cuadro "Datos Básicos")
                st.markdown("---")
                st.markdown("#### (Primer Cuadro \"Datos Básicos\")")
                col_b1, col_b2 = st.columns(2)
                with col_b1:
                    nombre_sc = st.text_input("Nombre S.C:", value=r['Nombre_Ciudadano'], disabled=disabled_state)
                    cc_sc = st.text_input("CC.:", value=r['Cedula_Ciudadano'], disabled=disabled_state)
                    nro_exp = st.text_input("Nro. expediente:", value="", disabled=disabled_state)
                with col_b2:
                    nro_comp = st.text_input("Nro. Comparendo:", value=r['Numero_Comparendo'], disabled=disabled_state)
                    placa_v = st.text_input("Placa del vehículo:", value=r['Placa_Vehiculo'], disabled=disabled_state)
                    f_inf = st.text_input("Fecha de la presunta infracción:", value=r['Fecha_Infraccion'], disabled=disabled_state)
                
                n_agente = st.text_input("Nombre del agente:", value=r['Nombre_Agente'], disabled=disabled_state)
                
                # (Segundo Cuadro "Audiencia Versión Libre")
                st.markdown("---")
                st.markdown("#### (Segundo Cuadro \"Audiencia Versión Libre\" es la misma de Impugnación)")
                st.markdown("##### 1. Generales audiencia")
                abv_vl = st.text_input("Abogado audiencia Versión Libre:", value=r['Abogado_Impugnacion'], disabled=disabled_state)
                
                try:
                    dt_vl_default = datetime.strptime(r['Fecha_Version_Libre'], "%Y-%m-%d %H:%M")
                except:
                    dt_vl_default = datetime.now()
                
                col_fv1, col_fv2 = st.columns(2)
                with col_fv1:
                    f_vl_date = st.date_input("Fecha audiencia versión libre (Calendario)", value=dt_vl_default.date(), disabled=disabled_state)
                with col_fv2:
                    f_vl_time = st.time_input("Hora versión libre", value=dt_vl_default.time(), disabled=disabled_state)
                
                version_cond = st.text_area("Versión del SC:", value=r['Version_Conductor'], disabled=disabled_state, height=150)
                
                st.markdown("##### 2. Pruebas Decretadas:")
                opts_p = ["", "SI", "NO"]
                def get_idx_p(val_str, name):
                    if val_str == "": return 0
                    if f"{name}: SI" in val_str: return 1
                    if f"{name}: NO" in val_str: return 2
                    return 0

                p1 = st.selectbox("1. Declaración agente", opts_p, index=get_idx_p(str(r['Pruebas_Decretadas_E1']), "Declaración agente"), disabled=disabled_state)
                p2 = st.selectbox("2. Certificado seguridad vial", opts_p, index=get_idx_p(str(r['Pruebas_Decretadas_E1']), "Certificado seguridad vial"), disabled=disabled_state)
                p3 = st.selectbox("3. Certificado antecedentes procuraduría", opts_p, index=get_idx_p(str(r['Pruebas_Decretadas_E1']), "Certificado antecedentes procuraduría"), disabled=disabled_state)
                
                tiene_otra_p = "Otra:" in str(r['Pruebas_Decretadas_E1'])
                p_otro_sel = st.selectbox("4. Otra (Selector de Si y No)", ["NO", "SI"], index=1 if tiene_otra_p else 0, disabled=disabled_state)
                p_otro_txt = ""
                if p_otro_sel == "SI":
                    default_otra_txt = str(r['Pruebas_Decretadas_E1']).split("Otra:")[-1].strip() if tiene_otra_p else ""
                    p_otro_txt = st.text_input("Especifique otra prueba decretada (casilla obligatoria)", value=default_otra_txt, disabled=disabled_state)
                
                st.markdown("##### 3. Errores en diligenciamiento de Comparendo:")
                e1 = st.checkbox("Casilla 10, sin diligenciar fecha de expedición/vencimiento, dirección, teléfono, municipio, edad, correo electrónico", value="Casilla 10" in str(r['Errores_Diligenciamiento']), key="ck_err_dil_10", disabled=disabled_state)
                e2 = st.checkbox("Casilla 13, sin diligenciar datos del propietario", value="Casilla 13" in str(r['Errores_Diligenciamiento']), key="ck_err_dil_13", disabled=disabled_state)
                e3 = st.checkbox("Casilla 16, sin diligenciar dirección del patio, placa de la grúa, consecutivo", value="Casilla 16" in str(r['Errores_Diligenciamiento']), key="ck_err_dil_16", disabled=disabled_state)
                e4 = st.checkbox("Casilla 17, sin diligenciar", value="Casilla 17" in str(r['Errores_Diligenciamiento']), key="ck_err_dil_17", disabled=disabled_state)
                e5 = st.checkbox("Casilla 18, quien firma como testigo es un agente de tránsito", value="Casilla 18" in str(r['Errores_Diligenciamiento']), key="ck_err_dil_18", disabled=disabled_state)
                e6 = st.checkbox("En las casillas sin marcar no se consignó ningún indicativo que determine el no diligenciamiento de estas.", value="Sin marcar" in str(r['Errores_Diligenciamiento']), key="ck_err_dil_sin_marcar", disabled=disabled_state)
                
                default_ed_txt = ""
                if "Otro Errores Diligenciamiento:" in str(r['Errores_Diligenciamiento']):
                    default_ed_txt = str(r['Errores_Diligenciamiento']).split("Otro Errores Diligenciamiento:")[-1].strip()
                err_dil_otro_txt = st.text_input("Otro Errores Diligenciamiento (Cuadro de texto abierto)", value=default_ed_txt, disabled=disabled_state)

                # (Tercer Cuadro "Audiencia De Pruebas")
                st.markdown("---")
                st.markdown("#### (Tercer Cuadro \"Audiencia De Pruebas\")")
                st.markdown("##### 1. Generalidades Audiencia:")
                abv_p = st.text_input("Abogado segunda audiencia:", value=r['Abogado_Pruebas'], disabled=disabled_state)
                
                try:
                    dt_p_default = datetime.strptime(r['Fecha_Pruebas'], "%Y-%m-%d %H:%M")
                except:
                    dt_p_default = datetime.now()

                col_fp1, col_fp2 = st.columns(2)
                with col_fp1:
                    f_p_date = st.date_input("Fecha audiencia de pruebas (Calendario)", value=dt_p_default.date(), disabled=disabled_state)
                with col_fp2:
                    f_p_time = st.time_input("Hora audiencia de pruebas", value=dt_p_default.time(), disabled=disabled_state)
                
                st.markdown("##### 2. Identificación de Errores en procedimiento (en formato tabla):")
                ep1 = st.checkbox("El agente de tránsito interroga al conductor y a su acompañante.", value="El agente de tránsito interroga al conductor y a su acompañante." in str(r['Errores_Procedimiento']), key="ck_ep_1", disabled=disabled_state)
                ep2 = st.checkbox("No informan al presunto infractor a cerca de la razón del comparendo.", value="No informan al presunto infractor a cerca de la razón del comparendo." in str(r['Errores_Procedimiento']), key="ck_ep_2", disabled=disabled_state)
                ep3 = st.checkbox("El conductor recibe por parte del agente de tránsito un mal trato y amenazas.", value="El conductor recibe por parte del agente de tránsito un mal trato y amenazas." in str(r['Errores_Procedimiento']), key="ck_ep_3", disabled=disabled_state)
                ep4 = st.checkbox("Se remolco más de un vehículo a la vez en la misma grúa.", value="Se remolco más de un vehículo a la vez en la misma grúa." in str(r['Errores_Procedimiento']), key="ck_ep_4", disabled=disabled_state)
                ep5 = st.checkbox("Duración excesiva del tiempo de realización del procedimiento de imposición del comparendo.", value="Duración excesiva del tiempo de realización del procedimiento de imposición del comparendo." in str(r['Errores_Procedimiento']), key="ck_ep_5", disabled=disabled_state)
                ep6 = st.checkbox("El agente que solicita inicialmente los documentos al conductor es diferente al agente que se los entrega al terminar el procedimiento", value="El agente que solicita inicialmente los documentos al conductor es diferente al agente que se los entrega al terminar el procedimiento" in str(r['Errores_Procedimiento']), key="ck_ep_6", disabled=disabled_state)
                ep7 = st.checkbox("El comparendo lo impone un agente diferente al que detuvo la marcha del vehículo.", value="El comparendo lo impone un agente diferente al que detuvo la marcha del vehículo." in str(r['Errores_Procedimiento']), key="ck_ep_7", disabled=disabled_state)
                ep8 = st.checkbox("El retén no cumple con los requisitos de ley puesto que carecía de la señalización y número de agentes exigidos por la norma (en el puesto de control se encontraba un solo agente).", value="El retén no cumple con los requisitos de ley puesto que carecía de la señalización y número de agentes exigidos por la norma" in str(r['Errores_Procedimiento']), key="ck_ep_8", disabled=disabled_state)
                ep9 = st.checkbox("Procedimiento efectuado al conductor estando solo en el vehículo.", value="Procedimiento efectuado al conductor estando solo en el vehículo." in str(r['Errores_Procedimiento']), key="ck_ep_9", disabled=disabled_state)
                ep10 = st.checkbox("Procedimiento efectuado por varios agentes", value="Procedimiento efectuado por varios agentes" in str(r['Errores_Procedimiento']), key="ck_ep_10", disabled=disabled_state)
                ep11 = st.checkbox("Violación de derechos fundamentales", value="Violación de derechos fundamentales" in str(r['Errores_Procedimiento']), key="ck_ep_11", disabled=disabled_state)
                ep12 = st.checkbox("En desarrollo de la audiencia pública de impugnación, el operador jurídico interrogo al conductor durante la versión libre.", value="En desarrollo de la audiencia pública de impugnación, el operador jurídico interrogo al conductor durante la versión libre." in str(r['Errores_Procedimiento']), key="ck_ep_12", disabled=disabled_state)
                
                default_ep_txt = ""
                if "Otro Error Procedimiento:" in str(r['Errores_Procedimiento']):
                    default_ep_txt = str(r['Errores_Procedimiento']).split("Otro Error Procedimiento:")[-1].strip()
                ep_otro_txt = st.text_input("Otro (Identificación de Errores en procedimiento - Cuadro de texto abierto)", value=default_ep_txt, disabled=disabled_state)

                st.markdown("##### 3. Pago:")
                opts_pago = ["", "SI", "NO"]
                hubo_pago = st.selectbox("¿Hubo pago?", opts_pago, index=0 if r['Hubo_Pago']=="" else opts_pago.index(r['Hubo_Pago']), disabled=disabled_state)
                menc_pago = st.selectbox("Se menciona que no hubo pago en la versión libre", opts_pago, index=0 if r['Menciona_No_Pago']=="" else opts_pago.index(r['Menciona_No_Pago']), disabled=disabled_state)
                
                st.markdown("##### 4. Identificación de precedentes favorables (Formato Tabla)")
                prec1 = st.checkbox("Conductor manifiesta estar acompañado en el procedimiento", value="Conductor manifiesta estar acompañado en el procedimiento" in str(r['Precedentes_Favorables']), key="ck_prec_1", disabled=disabled_state)
                prec2 = st.checkbox("Conductor manifiesta estar solo durante el procedimiento", value="Conductor manifiesta estar solo durante el procedimiento" in str(r['Precedentes_Favorables']), key="ck_prec_2", disabled=disabled_state)
                prec3 = st.checkbox("Agente responde de forma incorrecta el interrogatorio", value="Agente responde de forma incorrecta el interrogatorio" in str(r['Precedentes_Favorables']), key="ck_prec_3", disabled=disabled_state)
                prec4 = st.checkbox("Agente manifiesta no recordar el procedimiento", value="Agente manifiesta no recordar el procedimiento" in str(r['Precedentes_Favorables']), key="ck_prec_4", disabled=disabled_state)
                prec5 = st.checkbox("Procedimiento efectuado por varios agentes", value="Procedimiento efectuado por varios agentes" in str(r['Precedentes_Favorables']), key="ck_prec_5", disabled=disabled_state)
                prec6 = st.checkbox("Agente acepta haber recolectado información durante imposición", value="Agente acepta haber recolectado información durante imposición" in str(r['Precedentes_Favorables']), key="ck_prec_6", disabled=disabled_state)
                prec7 = st.checkbox("Agente acepta NO observar pago", value="Agente acepta NO observar pago" in str(r['Precedentes_Favorables']), key="ck_prec_7", disabled=disabled_state)
                prec8 = st.checkbox("Agente rechaza contar con otro medio probatorio", value="Agente rechaza contar con otro medio probatorio" in str(r['Precedentes_Favorables']), key="ck_prec_8", disabled=disabled_state)
                prec9 = st.checkbox("Agente aporta pantallazo o fotografía como prueba", value="Agente aporta pantallazo o fotografía como prueba" in str(r['Precedentes_Favorables']), key="ck_prec_9", disabled=disabled_state)
                prec10 = st.checkbox("Agente aporta prueba videográfica", value="Agente aporta prueba videográfica" in str(r['Precedentes_Favorables']), key="ck_prec_10", disabled=disabled_state)
                prec11 = st.checkbox("Agente conversa con un tercero y después persigue al conductor.", value="Agente conversa con un tercero y después persigue al conductor." in str(r['Precedentes_Favorables']), key="ck_prec_11", disabled=disabled_state)
                prec12 = st.checkbox("Conductor no escucha manifestaciones del tercero", value="Conductor no escucha manifestaciones del tercero" in str(r['Precedentes_Favorables']), key="ck_prec_12", disabled=disabled_state)
                prec13 = st.checkbox("Comparendo sin firmas", value="Comparendo sin firmas" in str(r['Precedentes_Favorables']), key="ck_prec_13", disabled=disabled_state)
                prec14 = st.checkbox("Conductor abordado en sitios distintos", value="Conductor abordado en sitios distintos" in str(r['Precedentes_Favorables']), key="ck_prec_14", disabled=disabled_state)
                prec15 = st.checkbox("El conductor abandona el lugar de los hechos", value="El conductor abandona el lugar de los hechos" in str(r['Precedentes_Favorables']), key="ck_prec_15", disabled=disabled_state)
                prec16 = st.checkbox("Error en la identificación del acompañante del SC", value="Error en la identificación del acompañante del SC" in str(r['Precedentes_Favorables']), key="ck_prec_16", disabled=disabled_state)
                prec17 = st.checkbox("Falencias en el diligenciamiento del comparendo", value="Falencias en el diligenciamiento del comparendo" in str(r['Precedentes_Favorables']), key="ck_prec_17", disabled=disabled_state)
                prec18 = st.checkbox("Caducidad desde la fecha de los hechos al fallo de primera instancia", value="Caducidad desde la fecha de los hechos al fallo de primera instancia" in str(r['Precedentes_Favorables']), key="ck_prec_18", disabled=disabled_state)
                
                default_pr_txt = ""
                if "Otro Precedente:" in str(r['Precedentes_Favorables']):
                    default_pr_txt = str(r['Precedentes_Favorables']).split("Otro Precedente:")[-1].strip()
                prec_otro_txt = st.text_input("Otro Precedentes Favorables (Cuadro de texto abierto)", value=default_pr_txt, disabled=disabled_state)

                st.markdown("##### 5. Nota Adicional:")
                nota_adic = st.text_area("Entrada de texto para que el abogado diligencie comentario importante a tener en cuenta del caso", value=r['Nota_Adicional'], disabled=disabled_state)
                
                # (Cuarto Cuadro "Audiencia De Fallo")
                st.markdown("---")
                st.markdown("#### (Cuarto Cuadro \"Audiencia De Fallo\")")
                abv_f = st.text_input("Abogado audiencia fallo:", value=r['Abogado_Fallo'], disabled=disabled_state)
                
                try:
                    dt_f_default = datetime.strptime(r['Fecha_Fallo'], "%Y-%m-%d %H:%M")
                except:
                    dt_f_default = datetime.now()

                col_ff1, col_ff2 = st.columns(2)
                with col_ff1:
                    f_f_date = st.date_input("Fecha audiencia de fallo (Calendario)", value=dt_f_default.date(), disabled=disabled_state)
                with col_ff2:
                    f_f_time = st.time_input("Hora audiencia de fallo", value=dt_f_default.time(), disabled=disabled_state)

                sent_f = st.selectbox("Sentido del Fallo:", ["", "FAVORABLE", "DESFAVORABLE"], index=0 if r['Sentido_Fallo']=="" else (1 if r['Sentido_Fallo']=="FAVORABLE" else 2), disabled=disabled_state)
                apela = st.selectbox("Se Presenta Apelación:", ["", "SI", "NO"], index=0 if r['Se_Presenta_Apelacion']=="" else (1 if r['Se_Presenta_Apelacion']=="SI" else 2), disabled=disabled_state)
                
                st.markdown("---")
                col_btn_f1, col_btn_f2 = st.columns(2)
                with col_btn_f1:
                    if st.form_submit_button("✏️ Habilitar Edición"):
                        st.session_state.editando_ficha = True
                        st.rerun()
                with col_btn_f2:
                    if st.form_submit_button("💾 Guardar Cambios en Ficha"):
                        df.loc[i, 'Nombre_Ciudadano'] = nombre_sc
                        df.loc[i, 'Cedula_Ciudadano'] = cc_sc
                        df.loc[i, 'Numero_Comparendo'] = nro_comp
                        df.loc[i, 'Placa_Vehiculo'] = placa_v
                        df.loc[i, 'Fecha_Infraccion'] = f_inf
                        df.loc[i, 'Nombre_Agente'] = n_agente
                        
                        df.loc[i, 'Abogado_Impugnacion'] = abv_vl
                        df.loc[i, 'Fecha_Version_Libre'] = f"{f_vl_date.strftime('%Y-%m-%d')} {f_vl_time.strftime('%H:%M')}"
                        df.loc[i, 'Version_Conductor'] = version_cond
                        
                        pruebas_list = []
                        if p1 != "": pruebas_list.append(f"Declaración agente: {p1}")
                        if p2 != "": pruebas_list.append(f"Certificado seguridad vial: {p2}")
                        if p3 != "": pruebas_list.append(f"Certificado antecedentes procuraduría: {p3}")
                        if p_otro_sel == "SI" and p_otro_txt != "": pruebas_list.append(f"Otra: {p_otro_txt}")
                        df.loc[i, 'Pruebas_Decretadas_E1'] = " | ".join(pruebas_list)
                        
                        err_dil_list = []
                        if e1: err_dil_list.append("Casilla 10")
                        if e2: err_dil_list.append("Casilla 13")
                        if e3: err_dil_list.append("Casilla 16")
                        if e4: err_dil_list.append("Casilla 17")
                        if e5: err_dil_list.append("Casilla 18")
                        if e6: err_dil_list.append("Sin marcar")
                        if err_dil_otro_txt != "": err_dil_list.append(f"Otro Errores Diligenciamiento: {err_dil_otro_txt}")
                        df.loc[i, 'Errores_Diligenciamiento'] = " | ".join(err_dil_list)
                        
                        df.loc[i, 'Abogado_Pruebas'] = abv_p
                        df.loc[i, 'Fecha_Pruebas'] = f"{f_p_date.strftime('%Y-%m-%d')} {f_p_time.strftime('%H:%M')}"
                        
                        err_proc_list = []
                        if ep1: err_proc_list.append("El agente de tránsito interroga al conductor y a su acompañante.")
                        if ep2: err_proc_list.append("No informan al presunto infractor a cerca de la razón del comparendo.")
                        if ep3: err_proc_list.append("El conductor recibe por parte del agente de tránsito un mal trato y amenazas.")
                        if ep4: err_proc_list.append("Se remolco más de un vehículo a la vez en la misma grúa.")
                        if ep5: err_proc_list.append("Duración excesiva del tiempo de realización del procedimiento de imposición del comparendo.")
                        if ep6: err_proc_list.append("El agente que solicita inicialmente los documentos al conductor es diferente al agente que se los entrega al terminar el procedimiento")
                        if ep7: err_proc_list.append("El comparendo lo impone un agente diferente al que detuvo la marcha del vehículo.")
                        if ep8: err_proc_list.append("El retén no cumple con los requisitos de ley puesto que carecía de la señalización y número de agentes exigidos por la norma")
                        if ep9: err_proc_list.append("Procedimiento efectuado al conductor estando solo en el vehículo.")
                        if ep10: err_proc_list.append("Procedimiento efectuado por varios agentes")
                        if ep11: err_proc_list.append("Violación de derechos fundamentales")
                        if ep12: err_proc_list.append("En desarrollo de la audiencia pública de impugnación, el operador jurídico interrogo al conductor durante la versión libre.")
                        if ep_otro_txt != "": err_proc_list.append(f"Otro Error Procedimiento: {ep_otro_txt}")
                        df.loc[i, 'Errores_Procedimiento'] = " | ".join(err_proc_list)
                        
                        df.loc[i, 'Hubo_Pago'] = hubo_pago
                        df.loc[i, 'Menciona_No_Pago'] = menc_pago
                        
                        prec_list = []
                        if prec1: prec_list.append("Conductor manifiesta estar acompañado en el procedimiento")
                        if prec2: prec_list.append("Conductor manifiesta estar solo durante el procedimiento")
                        if prec3: prec_list.append("Agente responde de forma incorrecta el interrogatorio")
                        if prec4: prec_list.append("Agente manifiesta no recordar el procedimiento")
                        if prec5: prec_list.append("Procedimiento efectuado por varios agentes")
                        if prec6: prec_list.append("Agente acepta haber recolectado información durante imposición")
                        if prec7: prec_list.append("Agente acepta NO observar pago")
                        if prec8: prec_list.append("Agente rechaza contar con otro medio probatorio")
                        if prec9: prec_list.append("Agente aporta pantallazo o fotografía como prueba")
                        if prec10: prec_list.append("Agente aporta prueba videográfica")
                        if prec11: prec_list.append("Agente conversa con un tercero y después persigue al conductor.")
                        if prec12: prec_list.append("Conductor no escucha manifestaciones del tercero")
                        if prec13: prec_list.append("Comparendo sin firmas")
                        if prec14: prec_list.append("Conductor abordado en sitios distintos")
                        if prec15: prec_list.append("El conductor abandona el lugar de los hechos")
                        if prec16: prec_list.append("Error en la identificación del acompañante del SC")
                        if prec17: prec_list.append("Falencias en el diligenciamiento del comparendo")
                        if prec18: prec_list.append("Caducidad desde la fecha de los hechos al fallo de primera instancia")
                        if prec_otro_txt != "": prec_list.append(f"Otro Precedente: {prec_otro_txt}")
                        df.loc[i, 'Precedentes_Favorables'] = " | ".join(prec_list)
                        
                        df.loc[i, 'Nota_Adicional'] = nota_adic
                        df.loc[i, 'Abogado_Fallo'] = abv_f
                        df.loc[i, 'Fecha_Fallo'] = f"{f_f_date.strftime('%Y-%m-%d')} {f_f_time.strftime('%H:%M')}"
                        df.loc[i, 'Sentido_Fallo'] = sent_f
                        df.loc[i, 'Se_Presenta_Apelacion'] = apela
                        
                        guardar_datos(df)
                        st.session_state.editando_ficha = False
                        st.success("✅ ¡Ficha guardada exitosamente en OneDrive!")
                        st.rerun()
