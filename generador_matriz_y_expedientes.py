import os
import fitz  # PyMuPDF
import re
import pandas as pd
import io
from PIL import Image
import pytesseract
import openpyxl 

# 📍 Obtenemos la ruta exacta de la carpeta donde está tu programa
RUTA_PROYECTO = os.path.dirname(os.path.abspath(__file__))

# ⚙️ CONFIGURACIÓN AUTOMÁTICA CORES-PLATAFORMA PARA TESSERACT
if os.name == 'nt':
    ruta_windows = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    if os.path.exists(ruta_windows):
        pytesseract.pytesseract.tesseract_cmd = ruta_windows
    
    # 🔥 EL TRUCO PORTABLE
    ruta_tessdata_local = os.path.join(RUTA_PROYECTO, 'tessdata')
    os.environ['TESSDATA_PREFIX'] = ruta_tessdata_local
else:
    ruta_mac = '/opt/homebrew/bin/tesseract'
    if os.path.exists(ruta_mac):
        pytesseract.pytesseract.tesseract_cmd = ruta_mac

# Mapeo de emergencia por si falla el archivo de bancos
MAPEO_BANCOS_BASICO = {
    'BANCO DEL PICHINCHA': '1029', 'BANCO DEL PACIFICO': '1028', 'BANCO DE GUAYAQUIL': '1006'
}

def limpiar_nan(valor):
    if pd.isna(valor) or str(valor).strip().lower() in ['nan', 'nat', 'none', 'null']: return ""
    return str(valor).strip()

def limpiar_numero_formato(valor):
    texto = limpiar_nan(valor)
    return texto[:-2] if texto.endswith('.0') else texto

def generar_sistema_y_matriz(ruta_pdf_maestro, ruta_excel_general, ruta_salida_grupo, callback=None):
    
    def notificar(mensaje=None, progreso=None, texto_progreso=None):
        if mensaje: print(mensaje)
        if callback: callback(mensaje, progreso, texto_progreso)

    notificar("🚀 Iniciando sistema de Consolidación Inteligente...", 2, "Preparando...")

    ruta_plantilla = os.path.join(RUTA_PROYECTO, "PLANTILLA.xlsx")
    archivo_matriz_salida = os.path.join(ruta_salida_grupo, "Matriz_Contratos_Registrados.xlsx")

    if not os.path.exists(ruta_plantilla):
        notificar("❌ ERROR: No se encontró 'PLANTILLA.xlsx' en la carpeta del programa.", 0, "Error")
        raise FileNotFoundError("Plantilla no encontrada.")

    # =================================================================
    # 🏦 CARGA DINÁMICA DEL DICCIONARIO DE BANCOS (bancos.xlsx o .csv)
    # =================================================================
    dict_bancos = {}
    ruta_bancos_xlsx = os.path.join(RUTA_PROYECTO, "bancos.xlsx")
    ruta_bancos_csv = os.path.join(RUTA_PROYECTO, "bancos.csv")
    
    try:
        if os.path.exists(ruta_bancos_xlsx):
            df_bancos = pd.read_excel(ruta_bancos_xlsx)
            notificar("🏦 Archivo 'bancos.xlsx' detectado y cargado.", 3, "Cargando Bancos...")
        elif os.path.exists(ruta_bancos_csv):
            df_bancos = pd.read_csv(ruta_bancos_csv)
            notificar("🏦 Archivo 'bancos.csv' detectado y cargado.", 3, "Cargando Bancos...")
        else:
            df_bancos = pd.DataFrame()
            notificar("⚠️ No hay 'bancos.xlsx' en la carpeta. Se usará diccionario básico.", 3, "Aviso")
            dict_bancos = MAPEO_BANCOS_BASICO

        if not df_bancos.empty:
            for _, row in df_bancos.iterrows():
                # Busca las columnas 'Nombre' y 'CODIGO' de tu archivo
                nombre_b = str(row.get('Nombre', '')).strip().upper()
                codigo_b = str(row.get('CODIGO', '')).strip()
                if nombre_b and nombre_b != 'NAN':
                    dict_bancos[nombre_b] = codigo_b
    except Exception as e:
        notificar(f"❌ Error al cargar lista de bancos: {e}", 3)

    # Carga de la Base General
    try:
        hojas = pd.read_excel(ruta_excel_general, sheet_name=None)
        df_general = pd.concat(hojas.values(), ignore_index=True)
        df_general['CEDULA_STR'] = df_general['CEDULA'].apply(lambda x: limpiar_numero_formato(x).zfill(10))
        notificar(f"📊 Base de datos Excel cargada ({len(df_general)} registros).", 5, "Cargando Excel...")
    except Exception as e:
        notificar(f"❌ Error crítico al leer el Excel General: {e}", 0)
        raise e

    if not os.path.exists(ruta_salida_grupo):
        os.makedirs(ruta_salida_grupo)

    registros_matriz_final = []
    doc_maestro = fitz.open(ruta_pdf_maestro)
    contador_estudiante = 1
    total_paginas = len(doc_maestro)
    
    notificar("🔎 Escaneando contratos y aplicando Salvavidas de Datos...", 10, f"0/{total_paginas}")
    
    for num_pagina in range(total_paginas):
        progreso_actual = 10 + int((num_pagina / total_paginas) * 75)
        texto_contador = f"{num_pagina + 1}/{total_paginas}"
        
        notificar(None, progreso_actual, texto_contador)
        
        pagina = doc_maestro[num_pagina]
        pix = pagina.get_pixmap(matrix=fitz.Matrix(3, 3))  
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        texto = pytesseract.image_to_string(img, lang='spa').replace('\n', ' ')
        
        if "CONTRATO" not in texto.upper() and "DBU-" not in texto.upper():
            continue 

        match_contrato = re.search(r"(\d{4}-\d{4}-(?:VUL|BEA)-\d+)", texto, re.IGNORECASE)
        contrato_final = match_contrato.group(1).strip().upper() if match_contrato else "NO_DETECTADO"
        tipo_beca = "EXCELENCIA" if "BEA" in contrato_final else "VULNERABILIDAD" if "VUL" in contrato_final else "OTRA"

        cedula_extraida = None
        patron_principal = r"c[ée]dula.*?pasaporte\D{1,20}([A-Z0-9\s\-]{9,18})"
        match_1 = re.search(patron_principal, texto, re.IGNORECASE)
        if match_1:
            cedula_sucia = match_1.group(1)
            cedula_extraida = re.sub(r'[\s\-]', '', cedula_sucia).strip()

        if not cedula_extraida:
            continue

        match_excel = df_general[df_general['CEDULA_STR'] == cedula_extraida]
        if not match_excel.empty:
            datos_estudiante = match_excel.iloc[0]
            nombre_oficial = str(datos_estudiante['NOMBRE']).strip()
            
            nombre_carpeta = f"{contador_estudiante}_{cedula_extraida}_{nombre_oficial.replace(' ', '_')}"
            ruta_nueva_carpeta = os.path.join(ruta_salida_grupo, nombre_carpeta)
            os.makedirs(ruta_nueva_carpeta, exist_ok=True)
            
            ruta_destino_pdf = os.path.join(ruta_nueva_carpeta, f"C_{cedula_extraida}.pdf")
            nuevo_pdf = fitz.open()
            nuevo_pdf.insert_pdf(doc_maestro, from_page=num_pagina, to_page=num_pagina)
            nuevo_pdf.save(ruta_destino_pdf)
            nuevo_pdf.close()
            
            # ====================================================================
            # 🛟 APLICACIÓN DE SALVAVIDAS (FALLBACKS) MEJORADOS
            # ====================================================================
            
            # 1. TELÉFONO
            tel_raw = limpiar_numero_formato(datos_estudiante.get('CELULAR', ''))
            telefono_final = tel_raw.zfill(10) if tel_raw else "022542160"

            # 2. DIRECCIÓN
            dir_raw = limpiar_nan(datos_estudiante.get('DIRECCION', ''))
            direccion_final = dir_raw if dir_raw else "Ciudadela Universitaria"

            # 3. TIPO Y NÚMERO DE CUENTA
            cuenta_final = limpiar_numero_formato(datos_estudiante.get('No CUENTA', ''))
            tipo_cuenta_final = limpiar_nan(datos_estudiante.get('TIPO CUENTA', ''))

            # Si nos falta la cuenta o el tipo, usamos la Regex Letal
            if not cuenta_final or not tipo_cuenta_final:
                # Busca la palabra cuenta, luego atrapa (ahorros/corriente) opcionalmente, salta basura, y atrapa el número
                match_cuenta = re.search(r"cuenta\s*(?:de\s*)?:?\s*(ahorros?|corriente)?.*?(\d{7,15})", texto, re.IGNORECASE)
                if match_cuenta:
                    if not tipo_cuenta_final and match_cuenta.group(1):
                        tipo = match_cuenta.group(1).upper()
                        tipo_cuenta_final = "AHORROS" if "AHORRO" in tipo else "CORRIENTE"
                    
                    if not cuenta_final and match_cuenta.group(2):
                        cuenta_final = match_cuenta.group(2).strip()

            tipo_cuenta_final = str(tipo_cuenta_final).upper() if tipo_cuenta_final else ""

            # 4. ENTIDAD BANCARIA Y CÓDIGO
            banco_raw = limpiar_nan(datos_estudiante.get('ENTIDAD BANCARIA', '')).upper()
            codigo_banco_final = ""
            
            if banco_raw in dict_bancos:
                codigo_banco_final = dict_bancos[banco_raw]
            else:
                banco_encontrado = ""
                # Ordenamos los bancos por longitud para evitar que "BANCO DEL AUSTRO" se corte solo en "BANCO"
                bancos_ordenados = sorted(dict_bancos.keys(), key=len, reverse=True)
                for nombre_b in bancos_ordenados:
                    if nombre_b in texto.upper() and len(nombre_b) > 4:
                        banco_encontrado = nombre_b
                        codigo_banco_final = dict_bancos[nombre_b]
                        break
                
                # Rescate Visual: Si Tesseract leyó el banco con algún error leve, atrapamos la palabra BANCO/COOP
                if not banco_encontrado:
                    match_banco_regex = re.search(r"(BANCO\s+DE\s+\w+|BANCO\s+DEL\s+\w+|BANCO\s+\w+|COOPERATIVA\s+[\w\s]+)", texto, re.IGNORECASE)
                    if match_banco_regex:
                        posible_banco = match_banco_regex.group(1).strip().upper()
                        # Cruzamos la captura con nuestro diccionario
                        for nombre_b, cod_b in dict_bancos.items():
                            if nombre_b in posible_banco or posible_banco in nombre_b:
                                banco_encontrado = nombre_b
                                codigo_banco_final = cod_b
                                break

                if banco_encontrado:
                    banco_raw = banco_encontrado

            # Registramos al estudiante
            registros_matriz_final.append({
                'No  BECARIO': contador_estudiante, 'NUM_CONTRATO': contrato_final,
                'PERIODO': limpiar_nan(datos_estudiante.get('PERIODO', '')),
                'FACULTAD': limpiar_nan(datos_estudiante.get('FACULTAD', '')),
                'CARRERA': limpiar_nan(datos_estudiante.get('CARRERA', '')),
                'TIPO_DOCUMENTO': 'CEDULA' if cedula_extraida.isdigit() and len(cedula_extraida) <= 10 else 'PASAPORTE',
                'CEDULA_O_PASAPORTE': cedula_extraida, 'NOMBRE': nombre_oficial,
                'CORREO ELECTRONICO': limpiar_nan(datos_estudiante.get('CORREO INSTITUCIONAL', '')),
                'TELEFONO ': telefono_final,
                'DIRECCION': direccion_final,
                'NUM_CUENTA': cuenta_final,
                'TIPO_CUENTA': tipo_cuenta_final,
                'ENTIDAD_BANCARIA': banco_raw, 
                'CODIGO BANCO': codigo_banco_final,
                'VALOR': 400
            })
            
            notificar(f"✅ Procesado #{contador_estudiante}: {nombre_oficial} [{tipo_beca}]", progreso_actual, texto_contador)
            contador_estudiante += 1

    doc_maestro.close()

    # =======================================================
    # 4. EXPORTACIÓN Y ACORDEÓN INVISIBLE
    # =======================================================
    texto_final = f"{total_paginas}/{total_paginas}"
    if registros_matriz_final:
        try:
            notificar(f"📝 Escribiendo datos y rellenando salvavidas en la plantilla...", 90, texto_final)
            wb = openpyxl.load_workbook(ruta_plantilla)
            ws = wb.active
            
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
                ws.cell(row=fila_estudiante, column=16, value=registro['VALOR']) 
                
                fila_estudiante += 1

            # ACORDEÓN INVISIBLE
            total_estudiantes = len(registros_matriz_final)
            capacidad_plantilla = 100  
            if total_estudiantes < capacidad_plantilla:
                fila_inicio_vacia = total_estudiantes + 2
                for r in range(fila_inicio_vacia, capacidad_plantilla + 2):
                    ws.row_dimensions[r].hidden = True
                notificar(f"🧹 Acordeón Invisible aplicado con éxito.", 95, texto_final)

            wb.save(archivo_matriz_salida)
            notificar(f"🎉 ¡PROCESO COMPLETADO EXITOSAMENTE!\n📁 Guardado en:\n{ruta_salida_grupo}", 100, texto_final)
            
        except Exception as e:
            notificar(f"❌ Error al guardar Excel: {e}", 100, "Error")
            raise e
    else:
        notificar("⚠️ Terminado: No se detectaron registros válidos en el PDF.", 100, texto_final)

if __name__ == "__main__":
    pass