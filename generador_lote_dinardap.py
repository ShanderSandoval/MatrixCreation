import os
import pandas as pd
import openpyxl

def extraer_lote_dinardap(ruta_matriz, ruta_salida, callback=None):
    def notificar(mensaje, porcentaje=None):
        print(mensaje)
        if callback:
            callback(mensaje, porcentaje)

    notificar("🚀 Iniciando extracción de Lote Dinardap...", 10)
    
    if not os.path.exists(ruta_matriz):
        raise FileNotFoundError(f"No se encontró la matriz en la ruta: {ruta_matriz}")
        
    try:
        # Leer el archivo de la matriz creada
        df = pd.read_excel(ruta_matriz, dtype=str)
        df.columns = df.columns.str.strip().str.upper()
        
        # Buscar las columnas de cédula y nombre
        col_cedula = None
        col_nombre = None
        
        posibles_cedulas = ['CEDULA_O_PASAPORTE', 'CEDULA', 'DOCUMENTO', 'CEDULA_STR']
        posibles_nombres = ['NOMBRE', 'NOMBRE OFICIAL', 'ESTUDIANTE']
        
        for c in posibles_cedulas:
            if c in df.columns:
                col_cedula = c
                break
                
        for n in posibles_nombres:
            if n in df.columns:
                col_nombre = n
                break
                
        if not col_cedula or not col_nombre:
            raise KeyError("No se encontraron las columnas de Cédula o Nombre en el archivo seleccionado.")
            
        # Filtrar registros válidos (omitir alertas o celdas vacías)
        df_filtrado = df[df[col_cedula].notna() & (df[col_cedula].str.strip() != "")].copy()
        # Omitir filas que sean de llenado manual o fallidas
        df_filtrado = df_filtrado[~df_filtrado[col_cedula].str.contains('ILEGIBLE|NO DETECTADA|FALLIDA', case=False, na=False)]
        
        # Construir el dataframe final del lote Dinardap
        df_lote = pd.DataFrame({
            'CEDULA': df_filtrado[col_cedula].str.strip(),
            'NOMBRE': df_filtrado[col_nombre].str.strip()
        })
        
        # Guardar archivo Excel limpio
        archivo_salida = os.path.join(ruta_salida, "Lote_Dinardap_Estudiantes.xlsx")
        df_lote.to_excel(archivo_salida, index=False)
        
        notificar(f"🎉 LOTE DINARDAP GENERADO CON ÉXITO.\n📄 Guardado en: {archivo_salida}", 100)
        
    except Exception as e:
        notificar(f"❌ Error al procesar el lote Dinardap: {e}", 100)
        raise e
