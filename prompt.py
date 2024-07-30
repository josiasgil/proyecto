import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Progressbar
import os
import re
import matplotlib.pyplot as plt

def cargar_archivos_del_directorio(directorio):
    archivos = [f for f in os.listdir(directorio) if f.endswith('.xlsx')]
    return archivos

def extraer_fecha_del_nombre_archivo(nombre_archivo):
    coincidencia = re.search(r'\d{4}\.\d{2}\.\d{2}', nombre_archivo)
    if coincidencia:
        año, mes, día = coincidencia.group(0).split('.')
        return int(año), int(mes), int(día)
    return None, None, None

def procesar_archivos(directorio, rango_columnas, fila_inicio, limite_datos):
    todos_datos = []
    archivos = cargar_archivos_del_directorio(directorio)
    conteo_datos = 0
    
    for archivo in archivos:
        try:
            ruta = os.path.join(directorio, archivo)
            df = pd.read_excel(ruta, sheet_name="ITEM_O", skiprows=fila_inicio-1, usecols=rango_columnas)
            año, mes, día = extraer_fecha_del_nombre_archivo(archivo)
            if año is not None:
                df['AÑO'] = año
                df['MES'] = mes
                df['DÍA'] = día
                todos_datos.append(df)
                
                conteo_datos += len(df)
                if conteo_datos >= limite_datos:
                    break
        except Exception as e:
            print(f"Error al procesar el archivo {archivo}: {e}")

    df_combinado = pd.concat(todos_datos, ignore_index=True)
    return df_combinado

def guardar_en_excel(df, archivo_salida):
    try:
        df.to_excel(archivo_salida, index=False)
    except Exception as e:
        print(f"Error al guardar el archivo {archivo_salida}: {e}")

def generar_histograma(df, directorio_salida, limite_datos):
    try:
        plt.figure()
        plt.hist(df.select_dtypes(include=['number']).values.flatten(), bins=20)
        plt.title(f'Histograma de Datos (Límite: {limite_datos})')
        plt.xlabel('Valor')
        plt.ylabel('Frecuencia')
        ruta_grafico = os.path.join(directorio_salida, 'histograma_datos.png')
        plt.savefig(ruta_grafico)
        plt.close()
    except Exception as e:
        print(f"Error al generar el histograma: {e}")

def generar_grafico_pastel_con_reporte(estadisticas, titulo, directorio_salida):
    try:
        plt.figure(figsize=(10, 6))

        # Crear gráfico de pastel
        plt.subplot(2, 1, 1)
        plt.pie(estadisticas.values(), labels=estadisticas.keys(), autopct='%1.1f%%', startangle=140)
        plt.title(titulo)

        # Agregar reporte conceptual
        plt.subplot(2, 1, 2)
        reporte = "\n".join([f"{key}: {value:.2f}" for key, value in estadisticas.items()])
        plt.text(0.1, 0.9, reporte, fontsize=12, ha='left')
        plt.axis('off')

        ruta_grafico = os.path.join(directorio_salida, f'{titulo.replace(" ", "_")}.png')
        plt.tight_layout()
        plt.savefig(ruta_grafico)
        plt.close()
    except Exception as e:
        print(f"Error al generar el gráfico con reporte: {e}")

def calcular_y_graficar_estadisticas(df, directorio_salida):
    try:
        if df.empty:
            messagebox.showwarning("Sin Datos", "No hay datos para analizar. Por favor, ejecute el proceso ETL primero.")
            return

        estadisticas = {
            'Promedio': df.select_dtypes(include=['number']).mean().mean(),
            'Moda': df.select_dtypes(include=['number']).mode().iloc[0].mean(),
            'Mediana': df.select_dtypes(include=['number']).median().mean(),
            'Desviación Estándar': df.select_dtypes(include=['number']).std().mean()
        }

        generar_grafico_pastel_con_reporte(estadisticas, 'Estadísticas Pie Chart', directorio_salida)
        messagebox.showinfo("Éxito", f"Gráfico de estadísticas guardado exitosamente en {directorio_salida}.")
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error al generar los gráficos de estadísticas: {e}")

def ejecutar_etl():
    try:
        directorio = entrada_directorio.get()
        rango_columnas = entrada_columnas.get()
        fila_inicio = int(entrada_fila_inicio.get())
        limite_datos = int(entrada_limite_datos.get())
        
        global df_final
        df_final = procesar_archivos(directorio, rango_columnas, fila_inicio, limite_datos)
        
        # Guardar automáticamente el archivo con el nombre "Out.xlsx"
        archivo_salida = os.path.join(directorio, "Out.xlsx")
        guardar_en_excel(df_final, archivo_salida)
        
        messagebox.showinfo("Éxito", f"Datos procesados y guardados exitosamente como {archivo_salida}.")
        mostrar_dataframe(df_final)
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error: {e}")

def guardar_archivo():
    try:
        if df_final is not None:
            archivo_salida = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                         filetypes=[("Archivos Excel", "*.xlsx")],
                                                         title="Guardar Como")
            if archivo_salida:
                guardar_en_excel(df_final, archivo_salida)
                messagebox.showinfo("Guardado", f"Archivo guardado exitosamente como {archivo_salida}")
        else:
            messagebox.showwarning("Sin Datos", "No hay datos para guardar. Por favor, ejecute el proceso ETL primero.")
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error al guardar el archivo: {e}")

def descargar_graficos():
    try:
        if df_final is not None:
            directorio_salida = filedialog.askdirectory(title="Seleccionar Directorio para Gráficos")
            if directorio_salida:
                limite_datos = int(entrada_limite_datos.get())
                generar_histograma(df_final, directorio_salida, limite_datos)
                messagebox.showinfo("Éxito", f"Gráficos guardados exitosamente en {directorio_salida}.")
        else:
            messagebox.showwarning("Sin Datos", "No hay datos para graficar. Por favor, ejecute el proceso ETL primero.")
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error al generar los gráficos: {e}")

def descargar_graficos_estadisticas():
    try:
        if df_final is not None:
            directorio_salida = filedialog.askdirectory(title="Seleccionar Directorio para Gráficos de Estadísticas")
            if directorio_salida:
                calcular_y_graficar_estadisticas(df_final, directorio_salida)
        else:
            messagebox.showwarning("Sin Datos", "No hay datos para graficar. Por favor, ejecute el proceso ETL primero.")
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error al generar los gráficos de estadísticas: {e}")

def mostrar_dataframe(df):
    top = tk.Toplevel(root)
    top.title("Vista Previa de Datos")

    texto = tk.Text(top)
    texto.pack(expand=True, fill='both')
    texto.insert(tk.END, df.to_string())

def en_buscar():
    directorio = filedialog.askdirectory()
    entrada_directorio.delete(0, tk.END)
    entrada_directorio.insert(0, directorio)

def crear_interfaz():
    global entrada_directorio, entrada_columnas, entrada_fila_inicio, entrada_limite_datos, df_final, root

    root = tk.Tk()
    root.title("Proceso ETL")

    tk.Label(root, text="Seleccionar Directorio:").pack(padx=10, pady=5)
    entrada_directorio = tk.Entry(root, width=50)
    entrada_directorio.pack(padx=10, pady=5)
    tk.Button(root, text="Buscar", command=en_buscar).pack(padx=10, pady=5)

    tk.Label(root, text="Rango de Columnas (ej. A:I):").pack(padx=10, pady=5)
    entrada_columnas = tk.Entry(root, width=50)
    entrada_columnas.pack(padx=10, pady=5)

    tk.Label(root, text="Fila Inicial:").pack(padx=10, pady=5)
    entrada_fila_inicio = tk.Entry(root, width=50)
    entrada_fila_inicio.pack(padx=10, pady=5)

    tk.Label(root, text="Límite de Datos (Número de Filas):").pack(padx=10, pady=5)
    entrada_limite_datos = tk.Entry(root, width=50)
    entrada_limite_datos.pack(padx=10, pady=5)

    tk.Button(root, text="Ejecutar ETL", command=ejecutar_etl).pack(padx=10, pady=20)
    tk.Button(root, text="Guardar Archivo", command=guardar_archivo).pack(padx=10, pady=5)
    tk.Button(root, text="Descargar Gráficos", command=descargar_graficos).pack(padx=10, pady=5)
    tk.Button(root, text="Descargar Gráficos de Estadísticas", command=descargar_graficos_estadisticas).pack(padx=10, pady=5)

    progreso = Progressbar(root, orient='horizontal', length=300, mode='indeterminate')
    progreso.pack(padx=10, pady=5)
    progreso.start()

    root.mainloop()

if __name__ == "__main__":
    df_final = None
    crear_interfaz()
