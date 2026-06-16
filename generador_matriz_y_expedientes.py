import os
import fitz  # PyMuPDF
import re
import pandas as pd
import io
from PIL import Image
import pytesseract
import openpyxl 

# ⚙️ CONFIGURACIÓN AUTOMÁTICA CORES-PLATAFORMA PARA TESSERACT
if os.name == 'nt':  # Si estás en Windows
    ruta_windows = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    if os.path.exists(ruta_windows):
        pytesseract.pytesseract.tesseract_cmd = ruta_windows
else:  # Si estás en macOS / Linux
    ruta_mac = '/opt/homebrew/bin/tesseract'
    if os.path.exists(ruta_mac):
        pytesseract.pytesseract.tesseract_cmd = ruta_mac

MAPEO_BANCOS = {
    'BANCO BOLIVARIANO': '1007',
    'BANCO DE GUAYAQUIL': '1006',
    'BANCO DE LOJA': '1025',
    'BANCO DEL AUSTRO': '1004',
    'BANCO DEL PACIFICO': '1028',
    'BANCO DEL PICHINCHA': '1029',
    'BANCO PRODUBANCO': '1033',
    'COOPERATIVA DE AHORRO Y CREDITO 29 DE OCTUBRE LTDA.': '1122',
    'COOPERATIVA DE AHORRO Y CREDITO VICENTINA MANUEL ESTEBAN GODOY ORTEGA LTDA.': '2129',
}

def limpiar_nan(valor):
    if pd.isna(valor) or str(valor).strip().lower() in ['nan', 'nat', 'none', 'null']:
        return ""
    return str(valor).strip()

def limpiar_numero_formato(valor):
    texto = limpiar_nan(valor)
    if texto.endswith('.0'):
        return texto[:-2]
    return texto

def generar_sistema_y_matriz(ruta_pdf_maestro, ruta_excel_general, ruta_destino_carpetas, archivo_matriz_salida, ruta_plantilla="PLANTILLA.xlsx"):
    print("🚀 [Programa Bienestar Estudiantil] Iniciando Consolidación con Nueva Plantilla (Valor: 400)...")
    
    try:
        hojas = pd.read_excel(ruta_excel_general, sheet_name=None)
        dfs_unificados = []
        for nombre_hoja, df in hojas.items():
            dfs_unificados.append(df)
        
        df_general = pd.concat(dfs_unificados, ignore_index=True)
        df_general['CEDULA_STR'] = df_general['CEDULA'].apply(lambda x: limpiar_numero_formato(x).zfill(10))
        print(f"📊 Base de datos cargada ({len(df_general)} registros).")
    except Exception as e:
        print(f"❌ Error crítico al leer el Excel General: {e}")
        return

    if not os.path.exists(ruta_destino_carpetas):
        os.makedirs(ruta_destino_carpetas)

    registros_matriz_final = []
    doc_maestro = fitz.open(ruta_pdf_maestro)
    contador_estudiante = 1
    
    print("🔎 Procesando PDF Maestro mediante Tesseract OCR...")
    
    for num_pagina in range(len(doc_maestro)):
        pagina = doc_maestro[num_pagina]
        pix = pagina.get_pixmap(matrix=fitz.Matrix(3, 3))  
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        texto_crudo = pytesseract.image_to_string(img, lang='spa')
        texto = texto_crudo.replace('\n', ' ')
        
        if "CONTRATO" not in texto.upper() and "DBU-" not in texto.upper():
            continue 

        match_contrato = re.search(r"(\d{4}-\d{4}-(?:VUL|BEA)-\d+)", texto, re.IGNORECASE)
        contrato_final = match_contrato.group(1).strip().upper() if match_contrato else "NO_DETECTADO"
        tipo_beca = "EXCELENCIA" if "BEA" in contrato_final else "VULNERABILIDAD" if "VUL" in contrato_final else "OTRA"

        cedula_extraida = None
        # 🚨 EL LIMPIADOR ANTI-GUIONES: Tolera espacios y guiones del OCR
        patron_principal = r"c[ée]dula.*?pasaporte\D{1,20}([A-Z0-9\s\-]{9,18})"
        match_1 = re.search(patron_principal, texto, re.IGNORECASE)
        if match_1:
            cedula_sucia = match_1.group(1)
            # Exprimimos la cédula: eliminamos cualquier espacio o guion
            cedula_extraida = re.sub(r'[\s\-]', '', cedula_sucia).strip()

        if not cedula_extraida:
            continue

        match_excel = df_general[df_general['CEDULA_STR'] == cedula_extraida]
        if not match_excel.empty:
            datos_estudiante = match_excel.iloc[0]
            nombre_oficial = str(datos_estudiante['NOMBRE']).strip()
            
            # --- CREACIÓN DE EXPEDIENTES ---
            nombre_carpeta = f"{contador_estudiante}_{cedula_extraida}_{nombre_oficial.replace(' ', '_')}"
            ruta_nueva_carpeta = os.path.join(ruta_destino_carpetas, nombre_carpeta)
            if not os.path.exists(ruta_nueva_carpeta):
                os.makedirs(ruta_nueva_carpeta)
            
            ruta_destino_pdf = os.path.join(ruta_nueva_carpeta, f"C_{cedula_extraida}.pdf")
            nuevo_pdf = fitz.open()
            nuevo_pdf.insert_pdf(doc_maestro, from_page=num_pagina, to_page=num_pagina)
            nuevo_pdf.save(ruta_destino_pdf)
            nuevo_pdf.close()
            
            banco_estudiante = limpiar_nan(datos_estudiante.get('ENTIDAD BANCARIA', '')).upper()
            codigo_banco = MAPEO_BANCOS.get(banco_estudiante, "")

            # 🚨 NUEVA ESTRUCTURA DE COLUMNAS (Sin Promedio, con VALOR Fijo en 400)
            registros_matriz_final.append({
                'No  BECARIO': contador_estudiante,
                'NUM_CONTRATO': contrato_final,
                'PERIODO': limpiar_nan(datos_estudiante.get('PERIODO', '')),
                'FACULTAD': limpiar_nan(datos_estudiante.get('FACULTAD', '')),
                'CARRERA': limpiar_nan(datos_estudiante.get('CARRERA', '')),
                'TIPO_DOCUMENTO': 'CEDULA' if cedula_extraida.isdigit() and len(cedula_extraida) <= 10 else 'PASAPORTE',
                'CEDULA_O_PASAPORTE': cedula_extraida,
                'NOMBRE': nombre_oficial,
                'CORREO ELECTRONICO': limpiar_nan(datos_estudiante.get('CORREO INSTITUCIONAL', '')),
                'TELEFONO ': limpiar_numero_formato(datos_estudiante.get('CELULAR', '')).zfill(10) if limpiar_numero_formato(datos_estudiante.get('CELULAR', '')) else "",
                'DIRECCION': limpiar_nan(datos_estudiante.get('DIRECCION', '')),
                'NUM_CUENTA': limpiar_numero_formato(datos_estudiante.get('No CUENTA', '')),
                'TIPO_CUENTA': limpiar_nan(datos_estudiante.get('TIPO CUENTA', '')) if pd.isna(datos_estudiante.get('TIPO CUENTA', '')) else str(datos_estudiante.get('TIPO CUENTA', '')).upper(),
                'ENTIDAD_BANCARIA': banco_estudiante,
                'CODIGO BANCO': codigo_banco,
                'VALOR': 400  # <--- VALOR FIJO INYECTADO AQUÍ
            })
            
            print(f"✅ Procesado #{contador_estudiante}: {nombre_oficial} [{tipo_beca}]")
            contador_estudiante += 1

    doc_maestro.close()

    # =======================================================
    # 4. EXPORTACIÓN: EL ACORDEÓN INVISIBLE A NUEVA PLANTILLA
    # =======================================================
    if registros_matriz_final:
        try:
            print(f"📝 Abriendo nueva plantilla original: {ruta_plantilla} ...")
            wb = openpyxl.load_workbook(ruta_plantilla)
            ws = wb.active
            
            # 📌 Inyectar los datos de los estudiantes
            fila_estudiante = 2
            for registro in registros_matriz_final:
                ws.cell(row=fila_estudiante, column=1, value=registro['No  BECARIO'])
                ws.cell(row=fila_estudiante, column=2, value=registro['NUM_CONTRATO'])
                ws.cell(row=fila_estudiante, column=3, value=registro['PERIODO'])
                ws.cell(row=fila_estudiante, column=4, value=registro['FACULTAD'])
                ws.cell(row=fila_estudiante, column=5, value=registro['CARRERA'])
                ws.cell(row=fila_estudiante, column=6, value=registro['TIPO_DOCUMENTO'])
                
                celda_cedula = ws.cell(row=fila_estudiante, column=7, value=registro['CEDULA_O_PASAPORTE'])
                celda_cedula.number_format = '@'
                
                ws.cell(row=fila_estudiante, column=8, value=registro['NOMBRE'])
                ws.cell(row=fila_estudiante, column=9, value=registro['CORREO ELECTRONICO'])
                
                celda_telefono = ws.cell(row=fila_estudiante, column=10, value=registro['TELEFONO '])
                celda_telefono.number_format = '@'
                
                ws.cell(row=fila_estudiante, column=11, value=registro['DIRECCION'])
                
                celda_cuenta = ws.cell(row=fila_estudiante, column=12, value=registro['NUM_CUENTA'])
                celda_cuenta.number_format = '@'
                
                ws.cell(row=fila_estudiante, column=13, value=registro['TIPO_CUENTA'])
                ws.cell(row=fila_estudiante, column=14, value=registro['ENTIDAD_BANCARIA'])
                ws.cell(row=fila_estudiante, column=15, value=registro['CODIGO BANCO'])
                
                # 📌 Columna 16: VALOR FIJO 400
                ws.cell(row=fila_estudiante, column=16, value=registro['VALOR'])
                
                fila_estudiante += 1

            # ====================================================================
            # 🧹 LA SOLUCIÓN DEFINITIVA: OCULTAR FILAS EN LUGAR DE BORRAR
            # ====================================================================
            total_estudiantes = len(registros_matriz_final)
            capacidad_plantilla = 100  # La plantilla tiene 100 filas para estudiantes
            
            if total_estudiantes < capacidad_plantilla:
                fila_inicio_vacia = total_estudiantes + 2
                
                # Bucle mágico: OCULTAMOS las filas sobrantes
                for r in range(fila_inicio_vacia, capacidad_plantilla + 2):
                    ws.row_dimensions[r].hidden = True
                
                print(f"🧹 Acordeón Invisible: Se ocultaron las filas vacías de la {fila_inicio_vacia} a la 101.")
                print(f"🎯 El bloque de firmas subió intacto con su formato original al 100%.")

            # Guardamos el archivo final
            wb.save(archivo_matriz_salida)
            print(f"\n🎉 ¡PROCESO COMPLETADO EXITOSAMENTE!")
            print(f"📊 Matriz perfecta guardada en: {archivo_matriz_salida}")
            
        except Exception as e:
            print(f"❌ Error al procesar la plantilla: {e}")
    else:
        print("⚠️ No se generaron registros válidos.")

if __name__ == "__main__":
    pass