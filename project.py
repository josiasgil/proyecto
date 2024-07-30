import streamlit as st
import pandas as pd
from sklearn.linear_model import LinearRegression
from sklearn.ensemble import RandomForestRegressor
import matplotlib.pyplot as plt
import seaborn as sns

# Función para cargar y analizar datos
def analizar_datos(df, columna_x, columna_y):
    # Escalar los datos para que estén entre 1 y 100
    df[columna_x] = (df[columna_x] - df[columna_x].min()) / (df[columna_x].max() - df[columna_x].min()) * 100
    df[columna_y] = (df[columna_y] - df[columna_y].min()) / (df[columna_y].max() - df[columna_y].min()) * 100

    # Seleccionar columnas
    X = df[[columna_x]].values
    y = df[columna_y].values

    # Regresión Lineal
    modelo_lineal = LinearRegression()
    modelo_lineal.fit(X, y)
    prediccion_lineal = modelo_lineal.predict(X)

    # Random Forest
    modelo_rf = RandomForestRegressor(n_estimators=100)
    modelo_rf.fit(X, y)
    prediccion_rf = modelo_rf.predict(X)

    # Guardar datos analizados en out.xlsx
    df_resultados = pd.DataFrame({
        columna_x: df[columna_x],
        columna_y: y,
        'Predicción Lineal': prediccion_lineal,
        'Predicción Random Forest': prediccion_rf
    })
    df_resultados.to_excel('out.xlsx', index=False)

    return df_resultados, prediccion_lineal, prediccion_rf

# Configuración de la aplicación
st.title('Proceso ETL y Análisis de Datos')
st.write("Selecciona un archivo Excel, columnas para el análisis y observa los resultados.")

# Cargar archivo
archivo = st.file_uploader("Sube un archivo Excel", type=["xlsx"])

if archivo:
    # Leer archivo
    df = pd.read_excel(archivo)
    st.write("Datos cargados:")
    st.write(df)

    # Selección de columna X y columna Y
    columnas = df.columns.tolist()
    columna_x = st.selectbox("Selecciona la columna X (independiente)", columnas)
    columna_y = st.selectbox("Selecciona la columna Y (dependiente)", columnas)

    # Selección de filtro por filas
    columna_filtro = st.selectbox("Selecciona una columna para filtrar filas", columnas)
    valores_unicos = df[columna_filtro].unique()
    valor_filtro = st.selectbox("Selecciona el valor de filtrado", valores_unicos)

    # Filtrar los datos
    df_filtrado = df[df[columna_filtro] == valor_filtro]

    if st.button("Analizar"):
        # Analizar datos y obtener resultados
        df_resultados, prediccion_lineal, prediccion_rf = analizar_datos(df_filtrado, columna_x, columna_y)
        st.write("Datos analizados:")
        st.write(df_resultados)

        # Gráficos de Barras y Tortas
        fig, ax = plt.subplots(1, 2, figsize=(15, 5))

        # Gráfico de Barras
        sns.barplot(x=df_filtrado[columna_y], y=df_filtrado.index, ax=ax[0], orient='h')
        ax[0].set_title('Distribución de la Variable Dependiente')
        ax[0].set_xlabel(columna_y)
        ax[0].set_ylabel('Conteo')

        # Gráfico de Tortas
        df_filtrado[columna_y].value_counts().plot.pie(autopct='%1.1f%%', ax=ax[1])
        ax[1].set_title('Proporción de Categorías')
        ax[1].set_ylabel('')

        st.pyplot(fig)

        # Análisis descriptivo
        st.write(f"**Análisis descriptivo**:")
        st.write(f"Se observa que la columna **{columna_y}** muestra una distribución en la que las categorías tienen las siguientes proporciones:")
        st.write(df_filtrado[columna_y].value_counts(normalize=True) * 100)

        # Gráficos de Regresión Lineal y Random Forest
        fig, ax = plt.subplots(2, 1, figsize=(10, 10))

        # Gráfico de Regresión Lineal
        sns.scatterplot(x=df_filtrado[columna_x], y=df_filtrado[columna_y], ax=ax[0])
        sns.lineplot(x=df_filtrado[columna_x], y=prediccion_lineal, color='red', ax=ax[0])
        ax[0].set_title('Regresión Lineal (Escala 1-100)')
        ax[0].set_xlabel(columna_x)
        ax[0].set_ylabel(columna_y)

        # Gráfico de Random Forest
        sns.scatterplot(x=df_filtrado[columna_x], y=df_filtrado[columna_y], ax=ax[1])
        sns.lineplot(x=df_filtrado[columna_x], y=prediccion_rf, color='green', ax=ax[1])
        ax[1].set_title('Random Forest (Escala 1-100)')
        ax[1].set_xlabel(columna_x)
        ax[1].set_ylabel(columna_y)

        st.pyplot(fig)

        st.success('El análisis ha sido completado y los resultados se han guardado en "out.xlsx".')
