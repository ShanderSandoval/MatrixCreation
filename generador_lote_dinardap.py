import os
import pandas as pd

def extraer_lote_dinardap(ruta_matriz, ruta_salida, callback=None):
    def notificar(mensaje, porcentaje=None):
        print(mensaje)
        if callback:
            callback(mensaje, porcentaje)

    notificar("🚀 Iniciando extracción de Lote Dinardap (2 Columnas con Encabezado)...", 10)
    
    if not os.path.exists(ruta_matriz):
        raise FileNotFoundError(f"No se encontró la matriz en la ruta: {ruta_matriz}")
        
    try:
        # Búsqueda dinámica de la fila de encabezados reales
        df_temp = pd.read_excel(ruta_matriz, header=None, dtype=str)
        fila_encabezado = 0
        
        for idx, row in df_temp.head(10).iterrows():
            row_str = " ".join(row.fillna("").astype(str)).upper()
            if "CEDULA" in row_str or "PASAPORTE" in row_str:
                fila_encabezado = idx
                break
                
        # Lectura aplicando el encabezado correcto
        df = pd.read_excel(ruta_matriz, header=fila_encabezado, dtype=str)
        df.columns = df.columns.str.strip().str.upper()
        
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
            raise KeyError("No se encontraron las columnas de Cédula o Nombre en la matriz.")
            
        # Filtrado de celdas vacías o con errores de OCR
        df_filtrado = df[df[col_cedula].notna() & (df[col_cedula].str.strip() != "")].copy()
        df_filtrado = df_filtrado[~df_filtrado[col_cedula].str.contains('ILEGIBLE|NO DETECTADA|FALLIDA', case=False, na=False)]
        
        # Creación del DataFrame final con 2 columnas
        df_lote = pd.DataFrame({
            'CEDULA': df_filtrado[col_cedula].str.strip(),
            'NOMBRE': df_filtrado[col_nombre].str.strip()
        })
        
        archivo_salida = os.path.join(ruta_salida, "Lote_Dinardap_Estudiantes.xlsx")
        
        # Exportación CON encabezados (header=True)
        df_lote.to_excel(archivo_salida, index=False, header=True)
        
        notificar(f"🎉 LOTE DINARDAP GENERADO CON ÉXITO.\n📄 Guardado en: {archivo_salida}", 100)
        
    except Exception as e:
        notificar(f"❌ Error al procesar el lote Dinardap: {e}", 100)
        raise e