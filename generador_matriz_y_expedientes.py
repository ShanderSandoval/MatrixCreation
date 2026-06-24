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
    ruta_windows = r'C:\Teseract\tesseract.exe'
    if os.path.exists(ruta_windows):
        pytesseract.pytesseract.tesseract_cmd = ruta_windows
    
    ruta_tessdata_local = os.path.join(RUTA_PROYECTO, 'tessdata')
    os.environ['TESSDATA_PREFIX'] = ruta_tessdata_local
else:
    ruta_mac = '/opt/homebrew/bin/tesseract'
    if os.path.exists(ruta_mac):
        pytesseract.pytesseract.tesseract_cmd = ruta_mac

def limpiar_nan(valor):
    if pd.isna(valor) or str(valor).strip().lower() in ['nan', 'nat', 'none', 'null']: return ""
    return str(valor).strip()

def limpiar_numero_formato(valor):
    texto = limpiar_nan(valor)
    return texto[:-2] if texto.endswith('.0') else texto

def obtener_valor_columna(df_row, posibles_nombres):
    for col in posibles_nombres:
        if col in df_row.index:
            return df_row[col]
    return ""

def formatear_documento_identidad(valor):
    texto = limpiar_numero_formato(valor)
    if texto.isdigit() and len(texto) < 10:
        return texto.zfill(10)
    return texto.upper()

def generar_sistema_y_matriz(ruta_pdf_maestro, ruta_excel_general, ruta_salida_grupo, crear_carpetas=True, exportar_lote_dinardap=False, callback=None):
    
    def notificar(mensaje=None, progreso=None, texto_progreso=None):
        if mensaje: print(mensaje)
        if callback: callback(mensaje, progreso, texto_progreso)

    notificar("🚀 Iniciando sistema de Consolidación Inteligente (V13 - Dinardap Integrado)...", 2, "Preparando...")

    ruta_plantilla = os.path.join(RUTA_PROYECTO, "PLANTILLA.xlsx")
    archivo_matriz_salida = os.path.join(ruta_salida_grupo, "Matriz_Contratos_Registrados.xlsx")

    if not os.path.exists(ruta_plantilla):
        notificar("❌ ERROR: No se encontró 'PLANTILLA.xlsx' en la carpeta.", 0, "Error")
        raise FileNotFoundError("Plantilla no encontrada.")

    # 🏦 Carga dinámica del diccionario de Bancos
    dict_bancos = {
        'BANCO BOLIVARIANO': '1007', 'BANCO DE GUAYAQUIL': '1006', 'BANCO DE LOJA': '1025',
        'BANCO DEL AUSTRO': '1004', 'BANCO DEL PACIFICO': '1028', 'BANCO DEL PICHINCHA': '1029',
        'BANCO PRODUBANCO': '1033', 'COOPERATIVA DE AHORRO Y CREDITO 29 DE OCTUBRE LTDA.': '1122',
        'COOPERATIVA DE AHORRO Y CREDITO VICENTINA MANUEL ESTEBAN GODOY ORTEGA LTDA.': '2129'
    }
    
    ruta_bancos_xlsx = os.path.join(RUTA_PROYECTO, "bancos.xlsx")
    try:
        if os.path.exists(ruta_bancos_xlsx):
            df_bancos = pd.read_excel(ruta_bancos_xlsx)
            for _, row in df_bancos.iterrows():
                nombre_b = str(row.get('Nombre', '')).strip().upper()
                codigo_b = str(row.get('CODIGO', '')).strip()
                if nombre_b and nombre_b != 'NAN': 
                    dict_bancos[nombre_b] = codigo_b
            notificar("🏦 Base de datos 'bancos.xlsx' cargada con éxito.", 3, "Bancos listos")
    except Exception as e:
        notificar(f"⚠️ Error al cargar bancos.xlsx, usando respaldo interno.", 3)

    # Carga del Excel General
    try:
        hojas = pd.read_excel(ruta_excel_general, sheet_name=None, dtype=str)
        df_general = pd.concat(hojas.values(), ignore_index=True)
        df_general.columns = df_general.columns.str.strip().str.upper()
        df_general['CEDULA_STR'] = df_general['CEDULA'].apply(formatear_documento_identidad)
        notificar(f"📊 Base de datos Excel cargada ({len(df_general)} registros).", 5, "Cargando Excel...")
    except Exception as e:
        notificar(f"❌ Error crítico al leer el Excel General: {e}", 0)
        raise e

    if not os.path.exists(ruta_salida_grupo):
        os.makedirs(ruta_salida_grupo)

    registros_matriz_final = []
    paginas_fallidas = [] 
    
    doc_maestro = fitz.open(ruta_pdf_maestro)
    total_paginas = len(doc_maestro)
    
    notificar("🔎 Escaneando contratos con Inteligencia Artificial...", 10, f"0/{total_paginas}")
    
    for num_pagina in range(total_paginas):
        numero_real_pagina = num_pagina + 1  
        
        progreso_actual = 10 + int((numero_real_pagina / total_paginas) * 75)
        texto_contador = f"{numero_real_pagina}/{total_paginas}"
        notificar(None, progreso_actual, texto_contador)
        
        pagina = doc_maestro[num_pagina]
        pix = pagina.get_pixmap(matrix=fitz.Matrix(3, 3))  
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        texto = pytesseract.image_to_string(img, lang='spa').replace('\n', ' ')
        
        texto_corregido = texto.replace("8EA", "BEA").replace("8UL", "VUL").upper()
        
        # ====================================================================
        # 🚨 ERROR 1: DOCUMENTO ILEGIBLE
        # ====================================================================
        if "CONTRATO" not in texto_corregido and "DBU-" not in texto_corregido:
            paginas_fallidas.append(numero_real_pagina)
            registros_matriz_final.append({
                'No  BECARIO': numero_real_pagina,
                'NUM_CONTRATO': f"⚠️ PÁG {numero_real_pagina} FALLIDA",
                'PERIODO': 'LLENAR MANUAL', 'FACULTAD': 'LLENAR MANUAL', 'CARRERA': 'LLENAR MANUAL',
                'TIPO_DOCUMENTO': '', 'CEDULA_O_PASAPORTE': 'ILEGIBLE',
                'NOMBRE': "❌ RAZÓN: El OCR no reconoció la palabra CONTRATO (Escaneo muy borroso)",
                'PROMEDIO': '', 'CORREO ELECTRONICO': '', 'TELEFONO ': '', 'DIRECCION': '',
                'NUM_CUENTA': '', 'TIPO_CUENTA': '', 'ENTIDAD_BANCARIA': '', 'CODIGO BANCO': '', 'VALOR': ''
            })
            continue 

        match_contrato = re.search(r"(\d{4}\s*-\s*\d{4}\s*-\s*(?:VUL|BEA)\s*-\s*\d+)", texto_corregido)
        contrato_final = match_contrato.group(1).replace(" ", "") if match_contrato else "NO_DETECTADO"
        tipo_beca = "EXCELENCIA" if "BEA" in contrato_final else "VULNERABILIDAD" if "VUL" in contrato_final else "OTRA"

        match_periodo = re.search(r"PERIODO ACAD[EÉ]MICO\s+(.*?)\s+CONTRATO", texto_corregido)
        periodo_pdf = match_periodo.group(1).strip() if match_periodo else ""

        # MOTOR CERO ERRORES: BÚSQUEDA INVERSA (CÉDULAS Y PASAPORTES)
        cedula_extraida = None
        texto_limpio_busqueda = re.sub(r'[\s\-]', '', texto_corregido) 
        
        patron_zona = r"(?:C[EÉ]DULA|PASAPORTE|C\.C\.)(?:NRO|NO)?([A-Z0-9]{7,20})"
        match_zona = re.search(patron_zona, texto_limpio_busqueda)
        
        if match_zona:
            zona_sucia = match_zona.group(1)
            for doc_excel in df_general['CEDULA_STR'].unique():
                if len(doc_excel) >= 7 and doc_excel in zona_sucia:
                    cedula_extraida = doc_excel
                    break
        
        if not cedula_extraida:
            for doc_excel in df_general['CEDULA_STR'].unique():
                if len(doc_excel) >= 7 and doc_excel in texto_limpio_busqueda:
                    cedula_extraida = doc_excel
                    break

        # ====================================================================
        # 🚨 ERROR 2: DOCUMENTO (CÉDULA) NO DETECTADO
        # ====================================================================
        if not cedula_extraida:
            paginas_fallidas.append(numero_real_pagina)
            registros_matriz_final.append({
                'No  BECARIO': numero_real_pagina,
                'NUM_CONTRATO': f"⚠️ PÁG {numero_real_pagina} FALLIDA",
                'PERIODO': 'LLENAR MANUAL', 'FACULTAD': 'LLENAR MANUAL', 'CARRERA': 'LLENAR MANUAL',
                'TIPO_DOCUMENTO': '', 'CEDULA_O_PASAPORTE': 'NO DETECTADA',
                'NOMBRE': "❌ RAZÓN: No se detectó ningún documento de identidad válido en esta página",
                'PROMEDIO': '', 'CORREO ELECTRONICO': '', 'TELEFONO ': '', 'DIRECCION': '',
                'NUM_CUENTA': '', 'TIPO_CUENTA': '', 'ENTIDAD_BANCARIA': '', 'CODIGO BANCO': '', 'VALOR': ''
            })
            continue

        match_excel = df_general[df_general['CEDULA_STR'] == cedula_extraida]
        
        # ====================================================================
        # 🚨 ERROR 3: ESTUDIANTE NO ESTÁ EN EL EXCEL
        # ====================================================================
        if match_excel.empty:
            paginas_fallidas.append(numero_real_pagina)
            registros_matriz_final.append({
                'No  BECARIO': numero_real_pagina,
                'NUM_CONTRATO': f"⚠️ PÁG {numero_real_pagina} FALLIDA",
                'PERIODO': 'LLENAR MANUAL', 'FACULTAD': 'LLENAR MANUAL', 'CARRERA': 'LLENAR MANUAL',
                'TIPO_DOCUMENTO': '', 'CEDULA_O_PASAPORTE': cedula_extraida,
                'NOMBRE': f"❌ RAZÓN: El documento {cedula_extraida} NO EXISTE en el Excel base",
                'PROMEDIO': '', 'CORREO ELECTRONICO': '', 'TELEFONO ': '', 'DIRECCION': '',
                'NUM_CUENTA': '', 'TIPO_CUENTA': '', 'ENTIDAD_BANCARIA': '', 'CODIGO BANCO': '', 'VALOR': ''
            })
            continue
            
        datos_estudiante = match_excel.iloc[0]
        nombre_oficial = str(datos_estudiante['NOMBRE']).strip()
        
        if crear_carpetas:
            nombre_carpeta = f"{numero_real_pagina}_{cedula_extraida}_{nombre_oficial.replace(' ', '_')}"
            ruta_nueva_carpeta = os.path.join(ruta_salida_grupo, nombre_carpeta)
            os.makedirs(ruta_nueva_carpeta, exist_ok=True)
            
            ruta_destino_pdf = os.path.join(ruta_nueva_carpeta, f"C_{cedula_extraida}.pdf")
            nuevo_pdf = fitz.open()
            nuevo_pdf.insert_pdf(doc_maestro, from_page=num_pagina, to_page=num_pagina)
            nuevo_pdf.save(ruta_destino_pdf)
            nuevo_pdf.close()
        
        # --- 🛠️ EXTRACCIÓN DINÁMICA ---
        periodo_excel = limpiar_nan(obtener_valor_columna(datos_estudiante, ['* PERIODO', 'PERIODO', 'PERÍODO']))
        periodo_val = periodo_pdf if periodo_pdf else periodo_excel
        
        facultad_val = limpiar_nan(obtener_valor_columna(datos_estudiante, ['FACULTAD']))
        carrera_val = limpiar_nan(obtener_valor_columna(datos_estudiante, ['CARRERA']))
        correo_val = limpiar_nan(obtener_valor_columna(datos_estudiante, ['CORREO', 'CORREO INSTITUCIONAL', 'EMAIL']))
        
        promedio_raw = limpiar_nan(obtener_valor_columna(datos_estudiante, ['GENERAL', 'PROMEDIO']))
        promedio_final = ""
        if tipo_beca == "EXCELENCIA" and promedio_raw:
            try:
                prom_str = str(promedio_raw).replace(",", ".")
                promedio_final = float(prom_str)
            except ValueError:
                promedio_final = promedio_raw

        tel_raw = limpiar_numero_formato(obtener_valor_columna(datos_estudiante, ['CELULAR', 'TELEFONO', 'TELÉFONO']))
        telefono_final = tel_raw.zfill(10) if tel_raw else "022542160"

        dir_raw = limpiar_nan(obtener_valor_columna(datos_estudiante, ['DIRECCION', 'DIRECCIÓN']))
        direccion_final = dir_raw if dir_raw else "Ciudadela Universitaria"

        cuenta_final = limpiar_numero_formato(obtener_valor_columna(datos_estudiante, ['NO. CUENTA', 'NO CUENTA', 'CUENTA', 'N° CUENTA']))
        tipo_cuenta_final = limpiar_nan(obtener_valor_columna(datos_estudiante, ['TIPO CUENTA', 'TIPO_CUENTA']))

        match_bloque = re.search(r"CUENTA\s*[:\-]?\s*(AHORROS?|CORRIENTE)?\D*?(\d{7,15})", texto_corregido)
        if match_bloque:
            if not tipo_cuenta_final and match_bloque.group(1):
                tipo_cuenta_final = "AHORROS" if "AHORRO" in match_bloque.group(1) else "CORRIENTE"
            if not cuenta_final and match_bloque.group(2):
                cuenta_final = match_bloque.group(2).strip()

        tipo_cuenta_final = str(tipo_cuenta_final).upper() if tipo_cuenta_final else ""

        banco_raw = limpiar_nan(obtener_valor_columna(datos_estudiante, ['ENTIDAD BANCARIA', 'BANCO'])).upper()
        codigo_banco_final = ""
        banco_encontrado = ""
        bancos_ordenados = sorted(dict_bancos.keys(), key=len, reverse=True)

        if banco_raw:
            for nombre_b in bancos_ordenados:
                if banco_raw in nombre_b or nombre_b in banco_raw:
                    banco_encontrado = nombre_b
                    codigo_banco_final = dict_bancos[nombre_b]
                    break

        if not banco_encontrado:
            for nombre_b in bancos_ordenados:
                if nombre_b in texto_corregido:
                    banco_encontrado = nombre_b
                    codigo_banco_final = dict_bancos[nombre_b]
                    break

        if not banco_encontrado:
            if "PICHINCHA" in texto_corregido:
                banco_encontrado, codigo_banco_final = "BANCO DEL PICHINCHA", "1029"
            elif "GUAYAQUIL" in texto_corregido:
                banco_encontrado, codigo_banco_final = "BANCO DE GUAYAQUIL", "1006"
            elif "PACIFICO" in texto_corregido or "PACÍFICO" in texto_corregido:
                banco_encontrado, codigo_banco_final = "BANCO DEL PACIFICO", "1028"
            elif "AUSTRO" in texto_corregido:
                banco_encontrado, codigo_banco_final = "BANCO DEL AUSTRO", "1004"
            elif "BOLIVARIANO" in texto_corregido:
                banco_encontrado, codigo_banco_final = "BANCO BOLIVARIANO", "1007"
            elif "LOJA" in texto_corregido:
                banco_encontrado, codigo_banco_final = "BANCO DE LOJA", "1025"
            elif "29 DE OCTUBRE" in texto_corregido:
                banco_encontrado, codigo_banco_final = "COOPERATIVA DE AHORRO Y CREDITO 29 DE OCTUBRE LTDA.", "1122"

        if banco_encontrado: 
            banco_raw = banco_encontrado

        registros_matriz_final.append({
            'No  BECARIO': numero_real_pagina, 
            'NUM_CONTRATO': contrato_final,
            'PERIODO': periodo_val,
            'FACULTAD': facultad_val,
            'CARRERA': carrera_val,
            'TIPO_DOCUMENTO': 'CEDULA' if cedula_extraida.isdigit() and len(cedula_extraida) <= 10 else 'PASAPORTE',
            'CEDULA_O_PASAPORTE': cedula_extraida, 
            'NOMBRE': nombre_oficial,
            'PROMEDIO': promedio_final,
            'CORREO ELECTRONICO': correo_val,
            'TELEFONO ': telefono_final, 
            'DIRECCION': direccion_final, 
            'NUM_CUENTA': cuenta_final,
            'TIPO_CUENTA': tipo_cuenta_final, 
            'ENTIDAD_BANCARIA': banco_raw, 
            'CODIGO BANCO': codigo_banco_final, 
            'VALOR': 400
        })
        
        notificar(f"✅ Procesado: {nombre_oficial} (Pág {numero_real_pagina})", progreso_actual, texto_contador)

    doc_maestro.close()

    # =======================================================
    # 4. EXPORTACIÓN Y ACORDEÓN INVISIBLE
    # =======================================================
    texto_final = f"{total_paginas}/{total_paginas}"
    if registros_matriz_final:
        try:
            notificar("📝 Escribiendo matriz y aplicando acordeón...", 95, texto_final)
            wb = openpyxl.load_workbook(ruta_plantilla)
            ws = wb.active
            
            fila_estudiante = 2
            lista_dinardap_rapida = [] # Memoria temporal por si piden lote directo
            
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
                ws.cell(row=fila_estudiante, column=9, value=registro['PROMEDIO']) 
                ws.cell(row=fila_estudiante, column=10, value=registro['CORREO ELECTRONICO'])
                
                celda_telefono = ws.cell(row=fila_estudiante, column=11, value=registro['TELEFONO '])
                celda_telefono.number_format = '@'
                
                ws.cell(row=fila_estudiante, column=12, value=registro['DIRECCION'])
                
                celda_cuenta = ws.cell(row=fila_estudiante, column=13, value=registro['NUM_CUENTA'])
                celda_cuenta.number_format = '@'
                
                ws.cell(row=fila_estudiante, column=14, value=registro['TIPO_CUENTA'])
                ws.cell(row=fila_estudiante, column=15, value=registro['ENTIDAD_BANCARIA'])
                ws.cell(row=fila_estudiante, column=16, value=registro['CODIGO BANCO'])
                ws.cell(row=fila_estudiante, column=17, value=registro['VALOR'])
                
                # Guardar en lote Dinardap solo los válidos
                if registro['CEDULA_O_PASAPORTE'] and registro['CEDULA_O_PASAPORTE'] not in ['ILEGIBLE', 'NO DETECTADA']:
                    if "FALLIDA" not in registro['NUM_CONTRATO']:
                        lista_dinardap_rapida.append({
                            'CEDULA': registro['CEDULA_O_PASAPORTE'],
                            'NOMBRE': registro['NOMBRE']
                        })
                
                fila_estudiante += 1

            total_estudiantes = len(registros_matriz_final)
            capacidad_plantilla = 100  
            if total_estudiantes < capacidad_plantilla:
                fila_inicio_vacia = total_estudiantes + 2
                for r in range(fila_inicio_vacia, capacidad_plantilla + 2):
                    ws.row_dimensions[r].hidden = True

            wb.save(archivo_matriz_salida)
            
            # 🔥 EXTRACCIÓN OPCIONAL EN FASE 1 DEL LOTE DINARDAP 🔥
            if exportar_lote_dinardap and lista_dinardap_rapida:
                df_din = pd.DataFrame(lista_dinardap_rapida)
                ruta_din_salida = os.path.join(ruta_salida_grupo, "Lote_Dinardap_Fase1.xlsx")
                df_din.to_excel(ruta_din_salida, index=False)
                notificar(f"📄 Archivo complementario masivo 'Lote_Dinardap_Fase1.xlsx' exportado.")

            if paginas_fallidas:
                notificar(f"⚠️ MATRIZ GENERADA CON {len(paginas_fallidas)} EXCEPCIONES.\nLas filas fallidas mantuvieron su orden.\n📁 Guardado en: {ruta_salida_grupo}", 100, texto_final)
            else:
                notificar(f"🎉 MATRIZ GENERADA CON ÉXITO Y SIN ERRORES.\n📁 Guardado en: {ruta_salida_grupo}", 100, texto_final)
        
        except Exception as e:
            notificar(f"❌ Error al guardar Excel: {e}", 100, "Error")
            raise e
    else:
        notificar("⚠️ No se detectaron registros válidos en el PDF.", 100, texto_final)
