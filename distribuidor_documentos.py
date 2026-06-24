import os
import shutil
import re
import fitz  # PyMuPDF
import io
from PIL import Image
import pytesseract
import shutil

# Esta línea verifica si tesseract está en el PATH del sistema
if shutil.which("tesseract"):
    # Si existe en el PATH, no necesita configuración manual
    pass
else:
    # Si no lo encuentra, solo entonces forzamos la ruta estándar por si acaso
    ruta_estandar = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    if os.path.exists(ruta_estandar):
        pytesseract.pytesseract.tesseract_cmd = ruta_estandar
    else:
        print("⚠️ Advertencia: Tesseract no encontrado en el PATH ni en la ruta estándar.")

def distribuir_documentos(ruta_grupo, origenes, ruta_pdf_contratos=None, callback=None):
    def notificar(mensaje, porcentaje=None):
        print(mensaje)
        if callback:
            callback(mensaje, porcentaje)

    notificar("🚀 Iniciando Alimentación de Carpetas...", 5)

    estudiantes = {}
    if not os.path.exists(ruta_grupo):
        raise FileNotFoundError("La carpeta del grupo seleccionada no existe.")

    # 1. Mapear las carpetas de los estudiantes
    carpetas_en_grupo = [d for d in os.listdir(ruta_grupo) if os.path.isdir(os.path.join(ruta_grupo, d))]
    
    for carpeta in carpetas_en_grupo:
        match = re.match(r"^\d+_([A-Za-z0-9]{7,15})_", carpeta)
        if match:
            cedula = match.group(1).upper()
            estudiantes[cedula] = os.path.join(ruta_grupo, carpeta)

    notificar(f"📁 Se detectaron {len(estudiantes)} carpetas de estudiantes válidas.", 15)

    total_tipos = len(origenes)
    
    # 🔥 CORRECCIÓN: Ahora no se detiene si le pasas el PDF maestro
    if total_tipos == 0 and not ruta_pdf_contratos:
        notificar("⚠️ No se seleccionó ningún tipo de documento extra para alimentar.", 100)
        return

    archivos_procesados = 0
    progreso_base = 20
    elementos_a_procesar = total_tipos + (1 if ruta_pdf_contratos else 0)
    salto_progreso = 80 / elementos_a_procesar
    paso_actual = 0

    # =================================================================
    # 1. DISTRIBUIR BANCARIOS, DINARDAP Y RECORDS (Desde Carpetas)
    # =================================================================
    for prefijo, ruta_origen in origenes.items():
        if not os.path.exists(ruta_origen):
            notificar(f"⚠️ La carpeta de origen para '{prefijo}_' no existe. Saltando...", progreso_base + (paso_actual * salto_progreso))
            paso_actual += 1
            continue

        archivos = [f for f in os.listdir(ruta_origen) if f.upper().startswith(prefijo) and f.lower().endswith('.pdf')]
        notificar(f"🔍 Escaneando {len(archivos)} archivos tipo '{prefijo}_'...", progreso_base + (paso_actual * salto_progreso))

        asignados_este_prefijo = set()

        for archivo in archivos:
            match_archivo = re.search(r"^[A-Z]_([A-Za-z0-9]+)", archivo, re.IGNORECASE)
            if match_archivo:
                cedula_archivo = match_archivo.group(1).upper()
                
                if cedula_archivo in estudiantes:
                    ruta_origen_archivo = os.path.join(ruta_origen, archivo)
                    ruta_destino_archivo = os.path.join(estudiantes[cedula_archivo], archivo)
                    
                    shutil.copy2(ruta_origen_archivo, ruta_destino_archivo)
                    archivos_procesados += 1
                    asignados_este_prefijo.add(cedula_archivo)
                    notificar(f"✅ [{prefijo}] Asignado a: {os.path.basename(estudiantes[cedula_archivo])}")

        # 🚨 REPORTE DIRECTO EN PANTALLA DE FALTANTES 🚨
        faltantes = set(estudiantes.keys()) - asignados_este_prefijo
        if faltantes:
            nombre_documento = "Certificados Bancarios" if prefijo == "B" else "Dinardap" if prefijo == "I" else "Récords" if prefijo == "R" else f"Documento {prefijo}"
            notificar(f"\n=======================================================")
            notificar(f"⚠️ ATENCIÓN: FALTARON {nombre_documento} PARA {len(faltantes)} ESTUDIANTES:")
            for cedula_faltante in sorted(faltantes):
                nombre_carpeta = os.path.basename(estudiantes[cedula_faltante])
                notificar(f"❌ {cedula_faltante}  -->  (Falta en {nombre_carpeta})")
            notificar(f"=======================================================\n")
        
        paso_actual += 1

    # =================================================================
    # 2. DISTRIBUIR CONTRATOS DESDE EL PDF MAESTRO (Cortar y Pegar)
    # =================================================================
    if ruta_pdf_contratos and os.path.exists(ruta_pdf_contratos):
        notificar(f"📄 Desfragmentando PDF Maestro de Contratos Unidos...", progreso_base + (paso_actual * salto_progreso))
        try:
            doc_maestro = fitz.open(ruta_pdf_contratos)
            total_paginas = len(doc_maestro)
            asignados_contratos = set()

            for num_pagina in range(total_paginas):
                pagina = doc_maestro[num_pagina]
                pix = pagina.get_pixmap(matrix=fitz.Matrix(2, 2))  
                img = Image.open(io.BytesIO(pix.tobytes("png")))
                texto = pytesseract.image_to_string(img, lang='spa').upper()
                
                texto_limpio = re.sub(r'[\s\-]', '', texto)
                cedula_encontrada = None
                
                # Búsqueda Inversa con cédulas de los estudiantes registrados
                for cedula_est in estudiantes.keys():
                    if len(cedula_est) >= 7 and cedula_est in texto_limpio:
                        cedula_encontrada = cedula_est
                        break
                        
                if cedula_encontrada:
                    ruta_destino = os.path.join(estudiantes[cedula_encontrada], f"C_{cedula_encontrada}.pdf")
                    nuevo_pdf = fitz.open()
                    nuevo_pdf.insert_pdf(doc_maestro, from_page=num_pagina, to_page=num_pagina)
                    nuevo_pdf.save(ruta_destino)
                    nuevo_pdf.close()
                    archivos_procesados += 1
                    asignados_contratos.add(cedula_encontrada)
                    notificar(f"✅ [C] Contrato extraído y asignado a: {os.path.basename(estudiantes[cedula_encontrada])}")

            doc_maestro.close()
            
            # 🚨 REPORTE DE CONTRATOS FALTANTES 🚨
            faltantes_c = set(estudiantes.keys()) - asignados_contratos
            if faltantes_c:
                notificar(f"\n=======================================================")
                notificar(f"⚠️ ATENCIÓN: FALTARON CONTRATOS (C_) PARA {len(faltantes_c)} ESTUDIANTES:")
                for cedula_faltante in sorted(faltantes_c):
                    nombre_carpeta = os.path.basename(estudiantes[cedula_faltante])
                    notificar(f"❌ {cedula_faltante}  -->  (Falta en {nombre_carpeta})")
                notificar(f"=======================================================\n")
                
        except Exception as e:
            notificar(f"⚠️ Error al procesar el PDF Maestro: {e}")

    notificar(f"🎉 Alimentación completada: {archivos_procesados} documentos asignados exitosamente.", 100)