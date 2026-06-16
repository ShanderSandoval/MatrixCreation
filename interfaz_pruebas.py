import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk, scrolledtext
import threading
import os

# Importamos nuestra función maestra
from generador_matriz_y_expedientes import generar_sistema_y_matriz

class InterfazBienestar:
    def __init__(self, root):
        self.root = root
        self.root.title("Programa Bienestar Estudiantil - Consolidador Automático")
        self.root.geometry("750x630")
        self.root.configure(padx=20, pady=20)
        
        # Variables para almacenar las rutas
        self.ruta_pdf = tk.StringVar()
        self.ruta_excel = tk.StringVar()
        self.ruta_salida = tk.StringVar(value="Esperando selección de Excel...") # Estado inicial

        self.crear_widgets()

    def crear_widgets(self):
        # Título
        lbl_titulo = tk.Label(self.root, text="🎓 Automatización de Contratos y Expedientes", font=("Arial", 14, "bold"))
        lbl_titulo.pack(pady=(0, 15))

        # 1. PDF Maestro
        frame_pdf = tk.Frame(self.root)
        frame_pdf.pack(fill="x", pady=5)
        tk.Button(frame_pdf, text="1. Seleccionar PDF Contratos", width=30, command=self.seleccionar_pdf).pack(side="left")
        tk.Label(frame_pdf, textvariable=self.ruta_pdf, fg="blue", wraplength=450).pack(side="left", padx=10)

        # 2. Excel General (Aquí se autoconfigura la ruta de salida)
        frame_excel = tk.Frame(self.root)
        frame_excel.pack(fill="x", pady=5)
        tk.Button(frame_excel, text="2. Seleccionar Excel General", width=30, command=self.seleccionar_excel).pack(side="left")
        tk.Label(frame_excel, textvariable=self.ruta_excel, fg="blue", wraplength=450).pack(side="left", padx=10)

        # 3. Carpeta de Salida (Opcional cambiarla)
        frame_salida = tk.Frame(self.root)
        frame_salida.pack(fill="x", pady=5)
        tk.Button(frame_salida, text="3. Cambiar Salida (Opcional)", width=30, command=self.seleccionar_salida).pack(side="left")
        tk.Label(frame_salida, textvariable=self.ruta_salida, fg="gray", wraplength=450).pack(side="left", padx=10)

        # --- SECCIÓN DE PROGRESO ---
        self.lbl_indicador_contador = tk.Label(self.root, text="Procesando: 0/0 (0%)", font=("Arial", 11, "bold"), fg="#2196F3")
        self.lbl_indicador_contador.pack(pady=(15, 0))

        self.progreso_var = tk.DoubleVar()
        self.progreso = ttk.Progressbar(self.root, variable=self.progreso_var, maximum=100)
        self.progreso.pack(fill="x", pady=5)

        # Pantalla de registros (Consola en vivo)
        tk.Label(self.root, text="Registro de Procesamiento (En Vivo):", font=("Arial", 10, "bold")).pack(anchor="w")
        self.consola = scrolledtext.ScrolledText(self.root, height=13, state='disabled', bg="#1e1e1e", fg="#4CAF50", font=("Consolas", 10))
        self.consola.pack(fill="both", expand=True, pady=5)

        # Botón Procesar
        self.btn_procesar = tk.Button(self.root, text="🚀 PROCESAR GRUPO", font=("Arial", 12, "bold"), bg="#4CAF50", fg="white", command=self.iniciar_proceso)
        self.btn_procesar.pack(fill="x", pady=10)

    # --- Funciones de Selección ---
    def seleccionar_pdf(self):
        ruta = filedialog.askopenfilename(title="Seleccionar PDF Maestro", filetypes=[("Archivos PDF", "*.pdf")])
        if ruta: self.ruta_pdf.set(ruta)

    def seleccionar_excel(self):
        ruta = filedialog.askopenfilename(title="Seleccionar Excel Base", filetypes=[("Archivos Excel", "*.xlsx *.xls")])
        if ruta: 
            self.ruta_excel.set(ruta)
            
            # 🔥 LA MAGIA DE LA RUTA AUTOMÁTICA 🔥
            # Obtenemos la carpeta donde vive el Excel que seleccionó el usuario
            carpeta_del_excel = os.path.dirname(ruta)
            # Creamos automáticamente una carpeta "Resultados_Expedientes" ahí mismo
            ruta_automatica = os.path.join(carpeta_del_excel, "Resultados_Expedientes")
            self.ruta_salida.set(ruta_automatica)

    def seleccionar_salida(self):
        ruta = filedialog.askdirectory(title="Seleccionar Carpeta para Resultados")
        if ruta: self.ruta_salida.set(ruta) # El usuario sobrescribe la ruta automática

    # --- Sistema de Mensajes a la Pantalla ---
    def actualizar_pantalla(self, mensaje=None, porcentaje=None, texto_progreso=None):
        self.root.after(0, self._escribir_consola, mensaje, porcentaje, texto_progreso)

    def _escribir_consola(self, mensaje, porcentaje, texto_progreso):
        if mensaje:
            self.consola.config(state='normal')
            self.consola.insert(tk.END, mensaje + "\n")
            self.consola.see(tk.END) # Auto-scroll hacia abajo
            self.consola.config(state='disabled')
        if porcentaje is not None:
            self.progreso_var.set(porcentaje)
        if texto_progreso and porcentaje is not None:
            # Actualiza el texto flotante de arriba: "Procesando: 1/10 (15%)"
            self.lbl_indicador_contador.config(text=f"Procesando: {texto_progreso} ({int(porcentaje)}%)")

    # --- Lógica de Ejecución ---
    def iniciar_proceso(self):
        if not all([self.ruta_pdf.get(), self.ruta_excel.get()]):
            messagebox.showwarning("Faltan datos", "Por favor, selecciona el PDF y el Excel antes de continuar.")
            return

        if self.ruta_salida.get() == "Esperando selección de Excel...":
            messagebox.showwarning("Ruta de Salida", "Por favor, selecciona el archivo Excel primero.")
            return

        # Limpiamos la consola
        self.consola.config(state='normal')
        self.consola.delete('1.0', tk.END)
        self.consola.config(state='disabled')
        
        self.btn_procesar.config(state="disabled", text="⏳ Procesando... Por favor espera")
        self.progreso_var.set(0)
        self.lbl_indicador_contador.config(text="Preparando entorno...")

        # Lanzar el proceso en un hilo separado
        hilo = threading.Thread(target=self.ejecutar_robot)
        hilo.start()

    def ejecutar_robot(self):
        try:
            generar_sistema_y_matriz(
                ruta_pdf_maestro=self.ruta_pdf.get(),
                ruta_excel_general=self.ruta_excel.get(),
                ruta_salida_grupo=self.ruta_salida.get(), 
                callback=self.actualizar_pantalla
            )
            self.root.after(0, self.proceso_exitoso)
        except Exception as e:
            self.root.after(0, lambda: self.proceso_fallido(e))

    def proceso_exitoso(self):
        self.btn_procesar.config(state="normal", text="🚀 PROCESAR GRUPO")
        self.lbl_indicador_contador.config(text="¡Completado al 100%!", fg="#4CAF50")
        messagebox.showinfo("¡Éxito!", f"El proceso ha finalizado.\nRevisa la carpeta:\n{self.ruta_salida.get()}")

    def proceso_fallido(self, error):
        self.btn_procesar.config(state="normal", text="🚀 PROCESAR GRUPO")
        self.lbl_indicador_contador.config(text="Fallo en el proceso", fg="red")
        self.actualizar_pantalla(f"\n❌ ERROR CRÍTICO: {str(error)}")
        messagebox.showerror("Error de Procesamiento", f"Ocurrió un error inesperado:\n{str(error)}")

if __name__ == "__main__":
    ventana_principal = tk.Tk()
    app = InterfazBienestar(ventana_principal)
    ventana_principal.mainloop()