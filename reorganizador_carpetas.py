import os
import re

def reorganizar_carpetas(ruta_grupo, callback=None):
    def notificar(mensaje, porcentaje=None):
        print(mensaje)
        if callback:
            callback(mensaje, porcentaje)

    notificar("🚀 Iniciando Reorganización Secuencial...", 5)

    if not os.path.exists(ruta_grupo):
        raise FileNotFoundError("La carpeta del grupo seleccionada no existe.")

    # Listar solo las carpetas
    carpetas = [d for d in os.listdir(ruta_grupo) if os.path.isdir(os.path.join(ruta_grupo, d))]
    
    carpetas_validas = []
    
    # Filtrar carpetas que empiezan con un número seguido de un guion bajo
    for carpeta in carpetas:
        match = re.match(r"^(\d+)_(.+)$", carpeta)
        if match:
            numero_actual = int(match.group(1))
            resto_nombre = match.group(2)
            carpetas_validas.append((numero_actual, carpeta, resto_nombre))

    if not carpetas_validas:
        notificar("⚠️ No se encontraron carpetas con el formato 'Numero_Cedula...'.", 100)
        return

    # Ordenar las carpetas por su número actual para mantener el orden lógico
    carpetas_validas.sort(key=lambda x: x[0])

    total_carpetas = len(carpetas_validas)
    carpetas_renombradas = 0
    progreso_base = 20
    salto = 70 / total_carpetas if total_carpetas > 0 else 0

    # Iterar y reasignar números consecutivos (1, 2, 3...)
    for indice, (numero_actual, nombre_original, resto_nombre) in enumerate(carpetas_validas):
        numero_correcto = indice + 1
        
        # Solo renombramos si el número actual no coincide con el orden correcto
        if numero_actual != numero_correcto:
            nuevo_nombre = f"{numero_correcto}_{resto_nombre}"
            
            ruta_vieja = os.path.join(ruta_grupo, nombre_original)
            ruta_nueva = os.path.join(ruta_grupo, nuevo_nombre)
            
            os.rename(ruta_vieja, ruta_nueva)
            carpetas_renombradas += 1
            
            notificar(f"🔄 Ajustado: {nombre_original} ➔ {nuevo_nombre}", progreso_base + (indice * salto))

    notificar(f"🎉 Reorden completado. Se ajustaron {carpetas_renombradas} carpetas. Orden perfecto de 1 a {total_carpetas}.", 100)
