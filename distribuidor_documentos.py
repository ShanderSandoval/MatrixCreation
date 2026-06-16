import os
import shutil
import re

def distribuir_documentos(ruta_grupo, origenes, callback=None):
    """
    ruta_grupo: Carpeta donde están los expedientes de los estudiantes (ej. Resultados_Expedientes).
    origenes: Diccionario con las rutas de las carpetas de origen si fueron seleccionadas.
              Ej: {'B': 'ruta/bancos', 'I': 'ruta/dinardap', 'R': 'ruta/records'}
    """
    def notificar(mensaje, porcentaje=None):
        print(mensaje)
        if callback:
            callback(mensaje, porcentaje)

    notificar("🚀 Iniciando Alimentación de Carpetas...", 5)

    # 1. Mapear las carpetas de los estudiantes por número de cédula
    estudiantes = {}
    if not os.path.exists(ruta_grupo):
        raise FileNotFoundError("La carpeta del grupo seleccionada no existe.")

    carpetas_en_grupo = [d for d in os.listdir(ruta_grupo) if os.path.isdir(os.path.join(ruta_grupo, d))]
    
    for carpeta in carpetas_en_grupo:
        # Extraemos la cédula del nombre de la carpeta (Ej: 1_1751264340_Mites_Yacelga)
        match = re.match(r"^\d+_(\d{9,15})_", carpeta)
        if match:
            cedula = match.group(1)
            estudiantes[cedula] = os.path.join(ruta_grupo, carpeta)

    notificar(f"📁 Se detectaron {len(estudiantes)} carpetas de estudiantes válidas.", 15)

    # 2. Procesar cada tipo de documento seleccionado
    total_tipos = len(origenes)
    if total_tipos == 0:
        notificar("⚠️ No se seleccionó ningún tipo de documento para alimentar.", 100)
        return

    archivos_procesados = 0
    progreso_base = 20
    salto_progreso = 80 / total_tipos

    for i, (prefijo, ruta_origen) in enumerate(origenes.items()):
        if not os.path.exists(ruta_origen):
            notificar(f"⚠️ La carpeta de origen para '{prefijo}_' no existe. Saltando...", progreso_base + (i * salto_progreso))
            continue

        archivos = [f for f in os.listdir(ruta_origen) if f.upper().startswith(prefijo) and f.lower().endswith('.pdf')]
        notificar(f"🔍 Escaneando {len(archivos)} archivos tipo '{prefijo}_'...", progreso_base + (i * salto_progreso))

        for archivo in archivos:
            # Extraer cédula del archivo (Ej: B_1751264340.pdf -> 1751264340)
            match_archivo = re.search(r"^[A-Z]_(\d+)", archivo, re.IGNORECASE)
            if match_archivo:
                cedula_archivo = match_archivo.group(1)
                
                # Si la cédula del archivo coincide con una carpeta de estudiante
                if cedula_archivo in estudiantes:
                    ruta_origen_archivo = os.path.join(ruta_origen, archivo)
                    ruta_destino_archivo = os.path.join(estudiantes[cedula_archivo], archivo)
                    
                    # Copiamos el archivo (usamos copy2 para conservar metadatos y no borrar el original)
                    shutil.copy2(ruta_origen_archivo, ruta_destino_archivo)
                    archivos_procesados += 1
                    notificar(f"✅ [{prefijo}] Asignado a: {os.path.basename(estudiantes[cedula_archivo])}")

    notificar(f"🎉 Alimentación completada: {archivos_procesados} documentos asignados exitosamente.", 100)