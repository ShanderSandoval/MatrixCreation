import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk, scrolledtext
import threading
import os

from generador_matriz_y_expedientes import generar_sistema_y_matriz
from distribuidor_documentos import distribuir_documentos
from reorganizador_carpetas import reorganizar_carpetas
from generador_lote_dinardap import extraer_lote_dinardap

class InterfazBienestar:
    def __init__(self, root):
        self.root = root
        self.root.title("Programa Bienestar Estudiantil - Automatización Integral")
        self.root.geometry("850x780")
        self.root.configure(padx=10, pady=10)

        # Variables Pestaña 1
        self.ruta_pdf = tk.StringVar()
        self.ruta_excel = tk.StringVar()
        self.ruta_salida_matriz = tk.StringVar(value="Esperando selección de Excel...")
        self.opcion_crear_carpetas = tk.BooleanVar(value=True)
        self.opcion_lote_dinardap = tk.BooleanVar(value=False)
        
        # Variables Pestaña 2
        self.ruta_grupo_alimento = tk.StringVar()
        self.check_c = tk.BooleanVar(value=False)
        self.ruta_c = tk.StringVar()
        self.check_b = tk.BooleanVar(value=True)
        self.ruta_b = tk.StringVar()
        self.check_i = tk.BooleanVar(value=False)
        self.ruta_i = tk.StringVar()
        self.check_r = tk.BooleanVar(value=False)
        self.ruta_r = tk.StringVar()

        # Variables Pestaña 3
        self.ruta_grupo_renombrar = tk.StringVar()
        
        # Variables Lote Dinardap desde Matriz Existente
        self.ruta_matriz_existente = tk.StringVar()
        self.ruta_salida_lote_existente = tk.StringVar()

        self.crear_widgets()

    def crear_widgets(self):
        lbl_titulo = tk.Label(self.root, text="🎓 Sistema de Becas - Bienestar Estudiantil", font=("Arial", 16, "bold"), fg="#2c3e50")
        lbl_titulo.pack(pady=(0, 10))

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True)

        # ==================== PESTAÑA 1 ====================
        self.tab1 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab1, text="  📑 Fase 1: Matriz y Contratos  ")
        
        frame_t1 = tk.Frame(self.tab1, padx=20, pady=15)
        frame_t1.pack(fill="both", expand=True)

        f_pdf = tk.Frame(frame_t1); f_pdf.pack(fill="x", pady=6)
        tk.Button(f_pdf, text="1. PDF Contratos (C_)", width=25, command=self.sel_pdf).pack(side="left")
        tk.Label(f_pdf, textvariable=self.ruta_pdf, fg="blue", wraplength=450).pack(side="left", padx=10)

        f_excel = tk.Frame(frame_t1); f_excel.pack(fill="x", pady=6)
        tk.Button(f_excel, text="2. Excel Estudiantes", width=25, command=self.sel_excel).pack(side="left")
        tk.Label(f_excel, textvariable=self.ruta_excel, fg="blue", wraplength=450).pack(side="left", padx=10)

        f_salida = tk.Frame(frame_t1); f_salida.pack(fill="x", pady=6)
        tk.Button(f_salida, text="3. Cambiar Salida (Opcional)", width=25, command=self.sel_salida_matriz).pack(side="left")
        tk.Label(f_salida, textvariable=self.ruta_salida_matriz, fg="gray", wraplength=450).pack(side="left", padx=10)

        f_opcion = tk.Frame(frame_t1); f_opcion.pack(fill="x", pady=5)
        tk.Checkbutton(f_opcion, text="Crear Carpetas de Estudiantes y extraer PDFs individuales", 
                       variable=self.opcion_crear_carpetas, font=("Arial", 10, "bold"), fg="#2196F3").pack(side="left")
                       
        f_opcion2 = tk.Frame(frame_t1); f_opcion2.pack(fill="x", pady=5)
        tk.Checkbutton(f_opcion2, text="Generar adicionalmente archivo de Lote Dinardap (.xlsx)", 
                       variable=self.opcion_lote_dinardap, font=("Arial", 10, "bold"), fg="#9C27B0").pack(side="left")

        self.btn_procesar_t1 = tk.Button(frame_t1, text="🚀 PROCESAR MATRIZ Y CONTRATOS", font=("Arial", 11, "bold"), bg="#4CAF50", fg="white", command=self.iniciar_proceso_matriz)
        self.btn_procesar_t1.pack(fill="x", pady=15)

        # ==================== PESTAÑA 2 ====================
        self.tab2 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab2, text="  📂 Fase 2: Alimentar Expedientes  ")

        frame_t2 = tk.Frame(self.tab2, padx=20, pady=10)
        frame_t2.pack(fill="both", expand=True)
        
        # --- SUB-ZONA: EXTRAER LOTE DINARDAP DESDE MATRIZ EXISTENTE ---
        lbl_sub = tk.Label(frame_t2, text="🛠️ Utilidad: Obtener Lote Dinardap desde una Matriz ya existente", font=("Arial", 10, "bold"), fg="#9C27B0")
        lbl_sub.pack(anchor="w", pady=(0, 2))
        
        f_existente_1 = tk.Frame(frame_t2); f_existente_1.pack(fill="x", pady=3)
        tk.Button(f_existente_1, text="Seleccionar Matriz Excel", width=25, command=self.sel_matriz_existente).pack(side="left")
        tk.Label(f_existente_1, textvariable=self.ruta_matriz_existente, fg="blue", wraplength=450).pack(side="left", padx=10)
        
        self.btn_lote_externo = tk.Button(frame_t2, text="📋 EXTRAER EXCEL FORMATO DINARDAP (CEDULA Y NOMBRE)", font=("Arial", 9, "bold"), bg="#E91E63", fg="white", command=self.iniciar_proceso_lote_existente)
        self.btn_lote_externo.pack(fill="x", pady=(3, 15))
        
        # --- ZONA ALIMENTACIÓN NORMAL ---
        tk.Label(frame_t2, text="---------------------------------------------------------------------------------------------------------", fg="gray").pack()
        tk.Label(frame_t2, text="Distribuidor de Documentos (Inyectar Dinardap, Bancos, Récords):", font=("Arial", 10, "bold")).pack(anchor="w", pady=(5,5))
        
        f_grupo = tk.Frame(frame_t2); f_grupo.pack(fill="x", pady=3)
        tk.Button(f_grupo, text="Seleccionar Carpeta Grupo", width=25, command=self.sel_grupo_alimento).pack(side="left")
        tk.Label(f_grupo, textvariable=self.ruta_grupo_alimento, fg="blue", wraplength=450).pack(side="left", padx=10)

        tk.Label(frame_t2, text="Archivos a inyectar (Orígenes):", font=("Arial", 10, "bold")).pack(anchor="w", pady=(5,0))

        f_c = tk.Frame(frame_t2); f_c.pack(fill="x", pady=3)
        tk.Checkbutton(f_c, text="Contratos (C_) PDF Maestro", variable=self.check_c, width=22, anchor="w").pack(side="left")
        tk.Button(f_c, text="Buscar Archivo .pdf", width=15, command=lambda: self.sel_origen_archivo(self.ruta_c)).pack(side="left")
        tk.Label(f_c, textvariable=self.ruta_c, fg="gray", wraplength=350).pack(side="left", padx=10)

        f_b = tk.Frame(frame_t2); f_b.pack(fill="x", pady=3)
        tk.Checkbutton(f_b, text="Bancarios (B_) Carpeta", variable=self.check_b, width=22, anchor="w").pack(side="left")
        tk.Button(f_b, text="Buscar Carpeta", width=15, command=lambda: self.sel_origen_carpeta(self.ruta_b)).pack(side="left")
        tk.Label(f_b, textvariable=self.ruta_b, fg="gray", wraplength=350).pack(side="left", padx=10)

        f_i = tk.Frame(frame_t2); f_i.pack(fill="x", pady=3)
        tk.Checkbutton(f_i, text="Dinardap (I_) Carpeta", variable=self.check_i, width=22, anchor="w").pack(side="left")
        tk.Button(f_i, text="Buscar Carpeta", width=15, command=lambda: self.sel_origen_carpeta(self.ruta_i)).pack(side="left")
        tk.Label(f_i, textvariable=self.ruta_i, fg="gray", wraplength=350).pack(side="left", padx=10)

        f_r = tk.Frame(frame_t2); f_r.pack(fill="x", pady=3)
        tk.Checkbutton(f_r, text="Récords (R_) Carpeta", variable=self.check_r, width=22, anchor="w").pack(side="left")
        tk.Button(f_r, text="Buscar Carpeta", width=15, command=lambda: self.sel_origen_carpeta(self.ruta_r)).pack(side="left")
        tk.Label(f_r, textvariable=self.ruta_r, fg="gray", wraplength=350).pack(side="left", padx=10)

        self.btn_procesar_t2 = tk.Button(frame_t2, text="⚡ ALIMENTAR CARPETAS EXPEDIENTES", font=("Arial", 11, "bold"), bg="#2196F3", fg="white", command=self.iniciar_proceso_alimento)
        self.btn_procesar_t2.pack(fill="x", pady=10)

        # ==================== PESTAÑA 3 ====================
        self.tab3 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab3, text="  🛠️ Fase 3: Arreglos y Reorden  ")

        frame_t3 = tk.Frame(self.tab3, padx=20, pady=20)
        frame_t3.pack(fill="both", expand=True)

        tk.Label(frame_t3, text="¿Eliminaste estudiantes de la matriz o carpetas?", font=("Arial", 12, "bold")).pack(anchor="w")
        tk.Label(frame_t3, text="Esta herramienta analiza la carpeta del grupo, identifica los números faltantes (ej: si eliminaste el 24, el 25 pasará a ser el 24) y reescribe los nombres para mantener un orden numérico perfecto sin afectar el interior.", font=("Arial", 10), justify="left", wraplength=700).pack(anchor="w", pady=(5, 20))

        tk.Label(frame_t3, text="Selecciona la Carpeta del Grupo:", font=("Arial", 10, "bold")).pack(anchor="w")
        f_grupo_r = tk.Frame(frame_t3); f_grupo_r.pack(fill="x", pady=5)
        tk.Button(f_grupo_r, text="Seleccionar Carpeta", width=25, command=self.sel_grupo_renombrar).pack(side="left")
        tk.Label(f_grupo_r, textvariable=self.ruta_grupo_renombrar, fg="blue", wraplength=450).pack(side="left", padx=10)

        self.btn_procesar_t3 = tk.Button(frame_t3, text="🛠️ REORDENAR CARPETAS SECUENCIALMENTE", font=("Arial", 11, "bold"), bg="#FF9800", fg="white", command=self.iniciar_proceso_renombrar)
        self.btn_procesar_t3.pack(fill="x", pady=25)

        # ==================== CONSOLA Y PROGRESO ====================
        f_consola = tk.Frame(self.root)
        f_consola.pack(fill="both", expand=True, pady=5)

        self.lbl_indicador = tk.Label(f_consola, text="Esperando instrucciones...", font=("Arial", 10, "bold"), fg="#2196F3")
        self.lbl_indicador.pack()

        self.progreso_var = tk.DoubleVar()
        self.progreso = ttk.Progressbar(f_consola, variable=self.progreso_var, maximum=100)
        self.progreso.pack(fill="x", pady=5)

        self.consola = scrolledtext.ScrolledText(f_consola, height=10, state='disabled', bg="#1e1e1e", fg="#4CAF50", font=("Consolas", 9))
        self.consola.pack(fill="both", expand=True)

    # --- SELECCIONES ---
    def sel_pdf(self):
        r = filedialog.askopenfilename(title="Seleccionar PDF", filetypes=[("PDF", "*.pdf")])
        if r: self.ruta_pdf.set(r)

    def sel_excel(self):
        # 🔥 Acepta ambos formatos
        r = filedialog.askopenfilename(title="Seleccionar Excel", filetypes=[("Excel", "*.xlsx *.xls")])
        if r:
            self.ruta_excel.set(r)
            self.ruta_salida_matriz.set(os.path.join(os.path.dirname(r), "Resultados_Expedientes"))

    def sel_salida_matriz(self):
        r = filedialog.askdirectory(title="Seleccionar Salida")
        if r: self.ruta_salida_matriz.set(r)

    def sel_grupo_alimento(self):
        r = filedialog.askdirectory(title="Seleccionar Carpeta del Grupo (Destino)")
        if r: self.ruta_grupo_alimento.set(r)

    def sel_origen_carpeta(self, var_ruta):
        r = filedialog.askdirectory(title="Seleccionar Carpeta de Origen")
        if r: var_ruta.set(r)

    def sel_origen_archivo(self, var_ruta):
        r = filedialog.askopenfilename(title="Seleccionar PDF", filetypes=[("Archivos PDF", "*.pdf")])
        if r: var_ruta.set(r)

    def sel_grupo_renombrar(self):
        r = filedialog.askdirectory(title="Seleccionar Carpeta del Grupo a Reordenar")
        if r: self.ruta_grupo_renombrar.set(r)
        
    def sel_matriz_existente(self):
        # 🔥 CORRECCIÓN AQUÍ: Ahora la Pestaña 2 acepta tanto .xlsx como .xls para extraer la Dinardap
        r = filedialog.askopenfilename(title="Seleccionar Matriz Creada", filetypes=[("Excel", "*.xlsx *.xls")])
        if r:
            self.ruta_matriz_existente.set(r)
            self.ruta_salida_lote_existente.set(os.path.dirname(r))

    # --- COMUNICACIÓN Y EJECUCIÓN ---
    def actualizar_pantalla(self, mensaje=None, porcentaje=None, texto_extra=None):
        self.root.after(0, self._escribir_consola, mensaje, porcentaje, texto_extra)

    def _escribir_consola(self, mensaje, porcentaje, texto_extra):
        if mensaje is not None:
            self.consola.config(state='normal')
            # Corrección de formato para evitar el SyntaxError
            self.consola.insert(tk.END, str(mensaje) + "\n")
            self.consola.see(tk.END)
            self.consola.config(state='disabled')
        if porcentaje is not None:
            self.progreso_var.set(porcentaje)
        if texto_extra is not None:
            self.lbl_indicador.config(text=f"{texto_extra} ({int(porcentaje or 0)}%)")

    def iniciar_proceso_matriz(self):
        if not all([self.ruta_pdf.get(), self.ruta_excel.get()]):
            messagebox.showwarning("Faltan datos", "Selecciona el PDF y el Excel.")
            return
        self._preparar_ui()
        threading.Thread(target=self._hilo_matriz).start()

    def iniciar_proceso_alimento(self):
        if not self.ruta_grupo_alimento.get():
            messagebox.showwarning("Faltan datos", "Selecciona la carpeta del grupo de destino.")
            return

        origenes = {}
        if self.check_b.get() and self.ruta_b.get(): origenes['B'] = self.ruta_b.get()
        if self.check_i.get() and self.ruta_i.get(): origenes['I'] = self.ruta_i.get()
        if self.check_r.get() and self.ruta_r.get(): origenes['R'] = self.ruta_r.get()

        ruta_contratos = self.ruta_c.get() if self.check_c.get() else None

        if not origenes and not ruta_contratos:
            messagebox.showwarning("Faltan orígenes", "Selecciona al menos una casilla.")
            return

        self._preparar_ui()
        threading.Thread(target=self._hilo_alimento, args=(origenes, ruta_contratos)).start()

    def iniciar_proceso_renombrar(self):
        if not self.ruta_grupo_renombrar.get():
            messagebox.showwarning("Faltan datos", "Selecciona la carpeta que deseas reordenar.")
            return
        self._preparar_ui()
        threading.Thread(target=self._hilo_renombrar).start()
        
    def iniciar_proceso_lote_existente(self):
        if not self.ruta_matriz_existente.get():
            messagebox.showwarning("Faltan datos", "Selecciona primero la matriz excel ya creada.")
            return
        self._preparar_ui()
        threading.Thread(target=self._hilo_lote_existente).start()

    def _preparar_ui(self):
        self.consola.config(state='normal'); self.consola.delete('1.0', tk.END); self.consola.config(state='disabled')
        self.btn_procesar_t1.config(state="disabled")
        self.btn_procesar_t2.config(state="disabled")
        self.btn_procesar_t3.config(state="disabled")
        self.btn_lote_externo.config(state="disabled")

    def _hilo_matriz(self):
        try:
            generar_sistema_y_matriz(
                self.ruta_pdf.get(), 
                self.ruta_excel.get(), 
                self.ruta_salida_matriz.get(), 
                crear_carpetas=self.opcion_crear_carpetas.get(),
                exportar_lote_dinardap=self.opcion_lote_dinardap.get(),
                callback=self.actualizar_pantalla
            )
            self.root.after(0, lambda: self._fin_proceso("✅ Fase 1 Completada."))
        except Exception as error_capturado:
            mensaje = str(error_capturado)
            self.root.after(0, lambda m=mensaje: self._error_proceso(m))
    
    def _hilo_alimento(self, origenes, ruta_contratos):
        try:
            distribuir_documentos(self.ruta_grupo_alimento.get(), origenes, ruta_pdf_contratos=ruta_contratos, callback=self.actualizar_pantalla)
            self.root.after(0, lambda: self._fin_proceso("✅ Alimentación Completada."))
        except Exception as error_capturado:
            mensaje = str(error_capturado)
            self.root.after(0, lambda m=mensaje: self._error_proceso(m))

    def _hilo_renombrar(self):
        try:
            reorganizar_carpetas(self.ruta_grupo_renombrar.get(), callback=self.actualizar_pantalla)
            self.root.after(0, lambda: self._fin_proceso("✅ Reordenamiento Completado."))
        except Exception as error_capturado:
            mensaje = str(error_capturado)
            self.root.after(0, lambda m=mensaje: self._error_proceso(m))
            
    def _hilo_lote_existente(self):
        try:
            extraer_lote_dinardap(self.ruta_matriz_existente.get(), self.ruta_salida_lote_existente.get(), callback=self.actualizar_pantalla)
            self.root.after(0, lambda: self._fin_proceso("✅ Lote Dinardap Extraído."))
        except Exception as error_capturado:
            mensaje = str(error_capturado)
            self.root.after(0, lambda m=mensaje: self._error_proceso(m))

    def _fin_proceso(self, mensaje):
        self.btn_procesar_t1.config(state="normal")
        self.btn_procesar_t2.config(state="normal")
        self.btn_procesar_t3.config(state="normal")
        self.btn_lote_externo.config(state="normal")
        self.lbl_indicador.config(text=mensaje, fg="#4CAF50")
        messagebox.showinfo("Proceso Terminado", mensaje)

    def _error_proceso(self, error):
        self.btn_procesar_t1.config(state="normal")
        self.btn_procesar_t2.config(state="normal")
        self.btn_procesar_t3.config(state="normal")
        self.btn_lote_externo.config(state="normal")
        self.lbl_indicador.config(text="❌ Error en el proceso", fg="red")
        
        # Corrección del mensaje de error para que no de SyntaxError
        mensaje_final = "\n❌ ERROR: " + str(error)
        self.actualizar_pantalla(mensaje_final)
        messagebox.showerror("Error", str(error))

if __name__ == "__main__":
    root = tk.Tk()
    app = InterfazBienestar(root)
    root.mainloop()