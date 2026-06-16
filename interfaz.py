import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk, scrolledtext
import threading
import os

# Importamos nuestros dos motores lógicos
from generador_matriz_y_expedientes import generar_sistema_y_matriz
from distribuidor_documentos import distribuir_documentos

class InterfazBienestar:
    def __init__(self, root):
        self.root = root
        self.root.title("Programa Bienestar Estudiantil - Automatización Integral")
        self.root.geometry("800x650")
        self.root.configure(padx=10, pady=10)

        # Variables Pestaña 1 (Matriz)
        self.ruta_pdf = tk.StringVar()
        self.ruta_excel = tk.StringVar()
        self.ruta_salida_matriz = tk.StringVar(value="Esperando selección de Excel...")
        
        # Variables Pestaña 2 (Alimentación)
        self.ruta_grupo_alimento = tk.StringVar()
        
        self.check_b = tk.BooleanVar(value=True)
        self.ruta_b = tk.StringVar()
        
        self.check_i = tk.BooleanVar(value=False)
        self.ruta_i = tk.StringVar()
        
        self.check_r = tk.BooleanVar(value=False)
        self.ruta_r = tk.StringVar()

        self.crear_widgets()

    def crear_widgets(self):
        lbl_titulo = tk.Label(self.root, text="🎓 Sistema de Becas - Bienestar Estudiantil", font=("Arial", 16, "bold"), fg="#2c3e50")
        lbl_titulo.pack(pady=(0, 10))

        # Crear el contenedor de Pestañas
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True)

        # =========================================================
        # PESTAÑA 1: CONSOLIDACIÓN DE MATRIZ Y CONTRATOS
        # =========================================================
        self.tab1 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab1, text="  📑 Fase 1: Matriz y Contratos  ")
        
        frame_t1 = tk.Frame(self.tab1, padx=20, pady=20)
        frame_t1.pack(fill="both", expand=True)

        # 1. PDF Maestro
        f_pdf = tk.Frame(frame_t1)
        f_pdf.pack(fill="x", pady=8)
        tk.Button(f_pdf, text="1. PDF Contratos (C_)", width=25, command=self.sel_pdf).pack(side="left")
        tk.Label(f_pdf, textvariable=self.ruta_pdf, fg="blue", wraplength=450).pack(side="left", padx=10)

        # 2. Excel General
        f_excel = tk.Frame(frame_t1)
        f_excel.pack(fill="x", pady=8)
        tk.Button(f_excel, text="2. Excel Estudiantes", width=25, command=self.sel_excel).pack(side="left")
        tk.Label(f_excel, textvariable=self.ruta_excel, fg="blue", wraplength=450).pack(side="left", padx=10)

        # 3. Salida
        f_salida = tk.Frame(frame_t1)
        f_salida.pack(fill="x", pady=8)
        tk.Button(f_salida, text="3. Cambiar Salida (Opcional)", width=25, command=self.sel_salida_matriz).pack(side="left")
        tk.Label(f_salida, textvariable=self.ruta_salida_matriz, fg="gray", wraplength=450).pack(side="left", padx=10)

        self.btn_procesar_t1 = tk.Button(frame_t1, text="🚀 PROCESAR MATRIZ Y CARPETAS", font=("Arial", 11, "bold"), bg="#4CAF50", fg="white", command=self.iniciar_proceso_matriz)
        self.btn_procesar_t1.pack(fill="x", pady=15)

        # =========================================================
        # PESTAÑA 2: ALIMENTACIÓN DE DOCUMENTOS
        # =========================================================
        self.tab2 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab2, text="  📂 Fase 2: Alimentar Expedientes  ")

        frame_t2 = tk.Frame(self.tab2, padx=20, pady=20)
        frame_t2.pack(fill="both", expand=True)
        
        tk.Label(frame_t2, text="Destino (Carpetas de Estudiantes):", font=("Arial", 10, "bold")).pack(anchor="w")
        f_grupo = tk.Frame(frame_t2)
        f_grupo.pack(fill="x", pady=5)
        tk.Button(f_grupo, text="Seleccionar Carpeta Grupo", width=25, command=self.sel_grupo_alimento).pack(side="left")
        tk.Label(f_grupo, textvariable=self.ruta_grupo_alimento, fg="blue", wraplength=450).pack(side="left", padx=10)

        tk.Label(frame_t2, text="Archivos a inyectar (Orígenes):", font=("Arial", 10, "bold")).pack(anchor="w", pady=(15,0))

        # Checkboxes y Botones
        f_b = tk.Frame(frame_t2)
        f_b.pack(fill="x", pady=5)
        tk.Checkbutton(f_b, text="Bancarios (B_)", variable=self.check_b, width=15, anchor="w").pack(side="left")
        tk.Button(f_b, text="Buscar Carpeta", width=15, command=lambda: self.sel_origen(self.ruta_b)).pack(side="left")
        tk.Label(f_b, textvariable=self.ruta_b, fg="gray", wraplength=400).pack(side="left", padx=10)

        f_i = tk.Frame(frame_t2)
        f_i.pack(fill="x", pady=5)
        tk.Checkbutton(f_i, text="Dinardap (I_)", variable=self.check_i, width=15, anchor="w").pack(side="left")
        tk.Button(f_i, text="Buscar Carpeta", width=15, command=lambda: self.sel_origen(self.ruta_i)).pack(side="left")
        tk.Label(f_i, textvariable=self.ruta_i, fg="gray", wraplength=400).pack(side="left", padx=10)

        f_r = tk.Frame(frame_t2)
        f_r.pack(fill="x", pady=5)
        tk.Checkbutton(f_r, text="Récords (R_)", variable=self.check_r, width=15, anchor="w").pack(side="left")
        tk.Button(f_r, text="Buscar Carpeta", width=15, command=lambda: self.sel_origen(self.ruta_r)).pack(side="left")
        tk.Label(f_r, textvariable=self.ruta_r, fg="gray", wraplength=400).pack(side="left", padx=10)

        self.btn_procesar_t2 = tk.Button(frame_t2, text="⚡ ALIMENTAR CARPETAS", font=("Arial", 11, "bold"), bg="#2196F3", fg="white", command=self.iniciar_proceso_alimento)
        self.btn_procesar_t2.pack(fill="x", pady=15)

        # =========================================================
        # CONSOLA Y PROGRESO (COMPARTIDO PARA AMBAS PESTAÑAS)
        # =========================================================
        f_consola = tk.Frame(self.root)
        f_consola.pack(fill="both", expand=True, pady=5)

        self.lbl_indicador = tk.Label(f_consola, text="Esperando instrucciones...", font=("Arial", 10, "bold"), fg="#2c3e50")
        self.lbl_indicador.pack()

        self.progreso_var = tk.DoubleVar()
        self.progreso = ttk.Progressbar(f_consola, variable=self.progreso_var, maximum=100)
        self.progreso.pack(fill="x", pady=5)

        self.consola = scrolledtext.ScrolledText(f_consola, height=10, state='disabled', bg="#1e1e1e", fg="#4CAF50", font=("Consolas", 9))
        self.consola.pack(fill="both", expand=True)

    # --- SELECCIÓN PESTAÑA 1 ---
    def sel_pdf(self):
        r = filedialog.askopenfilename(title="Seleccionar PDF", filetypes=[("PDF", "*.pdf")])
        if r: self.ruta_pdf.set(r)

    def sel_excel(self):
        r = filedialog.askopenfilename(title="Seleccionar Excel", filetypes=[("Excel", "*.xlsx *.xls")])
        if r:
            self.ruta_excel.set(r)
            ruta_auto = os.path.join(os.path.dirname(r), "Resultados_Expedientes")
            self.ruta_salida_matriz.set(ruta_auto)

    def sel_salida_matriz(self):
        r = filedialog.askdirectory(title="Seleccionar Salida")
        if r: self.ruta_salida_matriz.set(r)

    # --- SELECCIÓN PESTAÑA 2 ---
    def sel_grupo_alimento(self):
        r = filedialog.askdirectory(title="Seleccionar Carpeta del Grupo (Destino)")
        if r: self.ruta_grupo_alimento.set(r)

    def sel_origen(self, var_ruta):
        r = filedialog.askdirectory(title="Seleccionar Carpeta de Origen")
        if r: var_ruta.set(r)

    # --- COMUNICACIÓN CONSOLA ---
    def actualizar_pantalla(self, mensaje=None, porcentaje=None, texto_extra=None):
        self.root.after(0, self._escribir_consola, mensaje, porcentaje, texto_extra)

    def _escribir_consola(self, mensaje, porcentaje, texto_extra):
        if mensaje:
            self.consola.config(state='normal')
            self.consola.insert(tk.END, mensaje + "\n")
            self.consola.see(tk.END)
            self.consola.config(state='disabled')
        if porcentaje is not None:
            self.progreso_var.set(porcentaje)
        if texto_extra is not None:
            self.lbl_indicador.config(text=f"{texto_extra} ({int(porcentaje or 0)}%)")

    # --- EJECUCIÓN PESTAÑA 1 ---
    def iniciar_proceso_matriz(self):
        if not all([self.ruta_pdf.get(), self.ruta_excel.get()]):
            messagebox.showwarning("Faltan datos", "Selecciona el PDF y el Excel.")
            return
        
        self.consola.config(state='normal'); self.consola.delete('1.0', tk.END); self.consola.config(state='disabled')
        self.btn_procesar_t1.config(state="disabled")
        self.btn_procesar_t2.config(state="disabled")
        
        threading.Thread(target=self._hilo_matriz).start()

    def _hilo_matriz(self):
        try:
            generar_sistema_y_matriz(self.ruta_pdf.get(), self.ruta_excel.get(), self.ruta_salida_matriz.get(), self.actualizar_pantalla)
            self.root.after(0, lambda: self._fin_proceso("✅ Fase 1 Completada."))
        except Exception as e:
            self.root.after(0, lambda: self._error_proceso(e))

    # --- EJECUCIÓN PESTAÑA 2 ---
    def iniciar_proceso_alimento(self):
        if not self.ruta_grupo_alimento.get():
            messagebox.showwarning("Faltan datos", "Selecciona la carpeta del grupo donde están los expedientes.")
            return

        origenes = {}
        if self.check_b.get() and self.ruta_b.get(): origenes['B'] = self.ruta_b.get()
        if self.check_i.get() and self.ruta_i.get(): origenes['I'] = self.ruta_i.get()
        if self.check_r.get() and self.ruta_r.get(): origenes['R'] = self.ruta_r.get()

        if not origenes:
            messagebox.showwarning("Faltan orígenes", "Selecciona al menos una casilla y su carpeta correspondiente.")
            return

        self.consola.config(state='normal'); self.consola.delete('1.0', tk.END); self.consola.config(state='disabled')
        self.btn_procesar_t1.config(state="disabled")
        self.btn_procesar_t2.config(state="disabled")
        
        threading.Thread(target=self._hilo_alimento, args=(origenes,)).start()

    def _hilo_alimento(self, origenes):
        try:
            distribuir_documentos(self.ruta_grupo_alimento.get(), origenes, self.actualizar_pantalla)
            self.root.after(0, lambda: self._fin_proceso("✅ Fase 2 Completada."))
        except Exception as e:
            self.root.after(0, lambda: self._error_proceso(e))

    # --- MANEJO DE FIN DE PROCESOS ---
    def _fin_proceso(self, mensaje):
        self.btn_procesar_t1.config(state="normal")
        self.btn_procesar_t2.config(state="normal")
        self.lbl_indicador.config(text=mensaje, fg="#4CAF50")
        messagebox.showinfo("Proceso Terminado", mensaje)

    def _error_proceso(self, error):
        self.btn_procesar_t1.config(state="normal")
        self.btn_procesar_t2.config(state="normal")
        self.lbl_indicador.config(text="❌ Error en el proceso", fg="red")
        self.actualizar_pantalla(f"\n❌ ERROR: {str(error)}")
        messagebox.showerror("Error", str(error))

if __name__ == "__main__":
    root = tk.Tk()
    app = InterfazBienestar(root)
    root.mainloop()