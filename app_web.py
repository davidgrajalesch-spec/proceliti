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
    if not token:
        st.error("⚠️ Error MSAL: No se pudo obtener el Token de Azure (Revisa Client ID / Secret / Tenant ID).")
    else:
        user_email = st.secrets.get("ONEDRIVE_USER_EMAIL", "")
        # Consulta directamente en OneDrive
        url = f"https://graph.microsoft.com/v1.0/users/{user_email}/drive/root:/1.%20David%20Grajales/Prueba%20Poweapps/TablaProcesos.xlsx:/content"
        headers = {"Authorization": f"Bearer {token}"}
        
        try:
            response = requests.get(url, headers=headers)
            if response.status_code == 200:
                st.success("✅ Conexión exitosa a OneDrive.")
                df = pd.read_excel(io.BytesIO(response.content), dtype=str)
                for col in columnas_requeridas:
                    if col not in df.columns:
                        df[col] = ""
                return df.fillna("")
            else:
                st.warning(f"⚠️ Microsoft Graph devolvió Estado {response.status_code}: {response.text}")
        except Exception as e:
            st.error(f"⚠️ Error de conexión HTTP: {str(e)}")

    # Respaldo local
    if not os.path.exists(ARCHIVO_EXCEL):
        df_vacio = pd.DataFrame(columns=columnas_requeridas)
        df_vacio.to_excel(ARCHIVO_EXCEL, index=False)
        return df_vacio
    
    df = pd.read_excel(ARCHIVO_EXCEL, dtype=str)
    for col in columnas_requeridas:
        if col not in df.columns:
            df[col] = ""
    return df.fillna("")
