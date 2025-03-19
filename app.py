import streamlit as st
import pandas as pd
import numpy as np
import pyodbc
import re
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
import datetime

# Configuración SharePoint
sharepoint_url = "https://ofimundocl.sharepoint.com/sites/BusinessIntelligence"
sharepoint_relative_path = "/sites/BusinessIntelligence/Documentos%20compartidos/Paneles%20Logísticos/GC_Entregables_CA/Modelo%20Repuestos/ModeloAbastecimiento.xlsx"
usuario = "carancibia@ofimundo.cl"
password = "cb*84629"  # ⚠ Puedes proteger esto con st.secrets si prefieres

# Subir archivo a SharePoint
def subir_a_sharepoint(file_path):
    try:
        ctx_auth = AuthenticationContext(sharepoint_url)
        if ctx_auth.acquire_token_for_user(usuario, password):
            ctx = ClientContext(sharepoint_url, ctx_auth)
            with open(file_path, 'rb') as file:
                file_content = file.read()
            target_folder_url = sharepoint_relative_path.rsplit('/', 1)[0]
            target_file_name = sharepoint_relative_path.rsplit('/', 1)[-1]
            target_folder = ctx.web.get_folder_by_server_relative_url(target_folder_url)
            target_folder.upload_file(target_file_name, file_content).execute_query()
            st.success("Archivo subido exitosamente a SharePoint.")
        else:
            st.error("Error al autenticar en SharePoint.")
    except Exception as e:
        st.error(f"Error al subir a SharePoint: {e}")

# Conexión SQL
def conectar_sql():
    try:
        conn = pyodbc.connect('DSN=SQLKIA;UID=carancibia;PWD=cb*84629;')
        return conn
    except Exception as e:
        st.error(f"Error al conectar a SQL Server: {e}")
        return None

# Cargar datos
def cargar_datos(conn, query):
    try:
        return pd.read_sql(query, conn)
    except Exception as e:
        st.error(f"Error al ejecutar query: {e}")
        return None

# Normalizar códigos
def crear_columna_normalizada(df, columna_original, columna_nueva):
    sufijos = ["-REC", "-R", "REC", "-1", "-A"]
    patron = r"(" + "|".join(re.escape(s) for s in sufijos) + r")$"
    df[columna_nueva] = df[columna_original].str.replace(patron, "", regex=True).str.strip()
    return df

# Agrupar y pivotear
def agrupar_y_pivotear_repuestos(df, anio_inicio=2020):
    df["AÑO"] = pd.to_datetime(df["fecha_llamada"]).dt.year
    df = df[df["AÑO"] >= anio_inicio]
    df_pivot = (
        df.groupby(["CODIGO PRODUCTO NORM", "AÑO"], as_index=False)
        .agg({"cantidad": "sum"})
        .pivot(index="CODIGO PRODUCTO NORM", columns="AÑO", values="cantidad")
        .fillna(0)
        .reset_index()
    )
    df_pivot.columns = [str(col).strip() if isinstance(col, int) else col for col in df_pivot.columns]
    return df_pivot

# Modelo de abastecimiento
def crear_modelo(df_repuestos, df_inventario):
    df_repuestos = crear_columna_normalizada(df_repuestos, "CODIGO PRODUCTO", "CODIGO PRODUCTO NORM")
    df_inventario = crear_columna_normalizada(df_inventario, "CodProd", "CODIGO PRODUCTO NORM")
    df_repuestos_pivot = agrupar_y_pivotear_repuestos(df_repuestos)

    df_inventario = df_inventario[df_inventario["codigo_bodega"].astype(str) == "8"]
    df_inventario_reducido = df_inventario[["CODIGO PRODUCTO NORM", "DesProd", "disponible_en_bodega"]]

    df_modelo = pd.merge(df_repuestos_pivot, df_inventario_reducido, on="CODIGO PRODUCTO NORM", how="inner")
    df_modelo = df_modelo.drop_duplicates(subset=["CODIGO PRODUCTO NORM"])

    columnas_años = [col for col in df_modelo.columns if col.isdigit()]
    anio_actual = datetime.datetime.now().year
    anios_para_promedio = [col for col in columnas_años if 2020 <= int(col) <= anio_actual - 1]

    df_modelo['PREDICCION_ANUAL'] = df_modelo[anios_para_promedio].mean(axis=1)
    df_modelo['PREDICCION_MENSUAL'] = np.ceil(df_modelo['PREDICCION_ANUAL'] / 12)
    df_modelo['STOCK_DE_SEGURIDAD'] = np.ceil(df_modelo['PREDICCION_MENSUAL'] / 2).astype(int)
    df_modelo['STOCK_NECESARIO_TOTAL'] = (df_modelo['PREDICCION_MENSUAL'] + df_modelo['STOCK_DE_SEGURIDAD']).astype(int)

    def calcular_compra(row):
        return max(0, row['STOCK_NECESARIO_TOTAL'] - row['disponible_en_bodega'])

    df_modelo['COMPRA_RECOMENDADA'] = df_modelo.apply(calcular_compra, axis=1).astype(int)

    return df_modelo

# STREAMLIT APP
st.title("📦 Modelo de Abastecimiento de Repuestos")

if st.button("🔄 Actualizar modelo"):
    conn = conectar_sql()
    if conn:
        df_repuestos = cargar_datos(conn, "SELECT * FROM CONTROLGESTION.REPOSITORIO.VT_DATOS_REPUESTOS")
        df_inventario = cargar_datos(conn, "SELECT * FROM CONTROLGESTION.REPOSITORIO.VT_DATOS_INVENTARIO")
        if df_repuestos is not None and df_inventario is not None:
            df_modelo = crear_modelo(df_repuestos, df_inventario)

            # Guardar en Excel y subir
            try:
                excel_path = "ModeloAbastecimiento.xlsx"
                df_modelo.to_excel(excel_path, index=False)
                subir_a_sharepoint(excel_path)
                st.success("✅ Modelo generado y subido a SharePoint.")
                st.dataframe(df_modelo)
            except Exception as e:
                st.error(f"Error al guardar o subir el archivo: {e}")
        else:
            st.error("❌ No se pudieron cargar los datos desde SQL.")