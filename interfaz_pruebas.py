import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from generador_matriz_y_expedientes import generar_sistema_y_matriz

class TextRedirector:
    """Redirige sys.stdout a un widget de texto de Tkinter en tiempo real."""
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, string):
        self.text_widget.configure(state='normal')
        self.text_widget.insert('end', string)
        self.text_widget.see('end')
        self.text_widget.configure(state='disabled')

    def flush(self):
        pass

class GeneradorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Programa Bienestar Estudiantil - Consolidación")
        self.root.geometry("750x650")
        self.root.configure(bg="#1e1e2e")  # Fondo oscuro moderno estilo Catppuccin
        
        # Estilo general
        self.configurar_estilos()
        
        # Título Principal
        title_label = tk.Label(
            root, 
            text="Consolidador de Matriz y Expedientes", 
            bg="#1e1e2e", 
            fg="#89b4fa", 
            font=('Helvetica', 16, 'bold')
        )
        title_label.pack(pady=15)
        
        # Frame de Formulario
        entry_frame = tk.Frame(root, bg="#1e1e2e")
        entry_frame.pack(padx=25, pady=10, fill='x')
        
        # Variables de control
        self.ruta_pdf = tk.StringVar()
        self.ruta_excel = tk.StringVar()
        self.ruta_carpetas = tk.StringVar()
        self.ruta_matriz = tk.StringVar()
        
        # Configurar campos de entrada
        self.crear_campo_seleccion(entry_frame, "PDF Maestro de Contratos:", self.ruta_pdf, self.seleccionar_pdf, 0)
        self.crear_campo_seleccion(entry_frame, "Excel General de Datos:", self.ruta_excel, self.seleccionar_excel, 1)
        self.crear_campo_seleccion(entry_frame, "Carpeta de Expedientes:", self.ruta_carpetas, self.seleccionar_carpeta, 2)
        self.crear_campo_seleccion(entry_frame, "Archivo Matriz (Salida):", self.ruta_matriz, self.seleccionar_matriz, 3)
        
        # Botón Ejecutar
        self.btn_ejecutar = tk.Button(
            root, 
            text="Iniciar Consolidación 🚀", 
            command=self.ejecutar_proceso, 
            bg="#a6e3a1", 
            fg="#11111b", 
            font=('Helvetica', 12, 'bold'), 
            activebackground="#94e2d5",
            relief="flat",
            cursor="hand2",
            height=2
        )
        self.btn_ejecutar.pack(pady=20)
        
        # Sección de Consola en Vivo
        console_label = tk.Label(
            root, 
            text="Salida del proceso en tiempo real:", 
            bg="#1e1e2e", 
            fg="#a6adc8", 
            font=('Helvetica', 10, 'italic')
        )
        console_label.pack(anchor='w', padx=25)
        
        self.txt_console = tk.Text(
            root, 
            height=14, 
            bg="#11111b", 
            fg="#a6e3a1", 
            font=('Courier New', 10), 
            state='disabled',
            relief="sunken",
            bd=2
        )
        self.txt_console.pack(padx=25, pady=5, fill='both', expand=True)
        
        # Redirección de logs
        self.stdout_orig = sys.stdout
        sys.stdout = TextRedirector(self.txt_console)
        
    def configurar_estilos(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TLabel', background='#1e1e2e', foreground='#cdd6f4', font=('Helvetica', 10))
        
    def crear_campo_seleccion(self, master, label_text, var, command, row):
        # Etiqueta
        lbl = tk.Label(master, text=label_text, bg="#1e1e2e", fg="#cdd6f4", font=('Helvetica', 10, 'bold'), width=25, anchor='w')
        lbl.grid(row=row, column=0, sticky='w', pady=8)
        
        # Entrada de Texto
        ent = tk.Entry(master, textvariable=var, bg="#313244", fg="#cdd6f4", insertbackground="white", relief="flat", bd=5, width=45)
        ent.grid(row=row, column=1, padx=10, pady=8, sticky='ew')
        
        # Botón
        btn = tk.Button(master, text="Buscar", command=command, bg="#45475a", fg="white", activebackground="#585b70", relief="flat", cursor="hand2")
        btn.grid(row=row, column=2, padx=5, pady=8)
        
        # Hacer que la columna central se expanda
        master.columnconfigure(1, weight=1)
        
    def seleccionar_pdf(self):
        ruta = filedialog.askopenfilename(title="Seleccionar PDF Maestro de Contratos", filetypes=[("Archivos PDF", "*.pdf")])
        if ruta:
            self.ruta_pdf.set(ruta)
            
    def seleccionar_excel(self):
        ruta = filedialog.askopenfilename(title="Seleccionar Excel General", filetypes=[("Archivos de Excel", "*.xlsx *.xls")])
        if ruta:
            self.ruta_excel.set(ruta)
            
            # Autocompletar sugerencias inteligentes basadas en el directorio del Excel seleccionado
            directorio_padre = os.path.dirname(ruta)
            if not self.ruta_matriz.get():
                self.ruta_matriz.set(os.path.join(directorio_padre, "Matriz_Contratos_Registrados.xlsx"))
            if not self.ruta_carpetas.get():
                self.ruta_carpetas.set(os.path.join(directorio_padre, "Expedientes_Organizados"))
            
    def seleccionar_carpeta(self):
        ruta = filedialog.askdirectory(title="Seleccionar Carpeta de Destino para Expedientes")
        if ruta:
            self.ruta_carpetas.set(ruta)
            
    def seleccionar_matriz(self):
        ruta = filedialog.asksaveasfilename(title="Guardar Matriz de Salida", defaultextension=".xlsx", filetypes=[("Archivo Excel", "*.xlsx")])
        if ruta:
            self.ruta_matriz.set(ruta)
            
    def ejecutar_proceso(self):
        pdf = self.ruta_pdf.get().strip()
        excel = self.ruta_excel.get().strip()
        carpetas = self.ruta_carpetas.get().strip()
        matriz = self.ruta_matriz.get().strip()
        
        if not pdf or not excel or not carpetas or not matriz:
            messagebox.showwarning("Campos incompletos", "Por favor, selecciona todas las rutas necesarias antes de ejecutar el proceso.")
            return
            
        # Deshabilitar interfaz mientras se procesa
        self.btn_ejecutar.config(state='disabled', text="Procesando... ⏳", bg="#45475a")
        
        # Ejecutar en hilo secundario para evitar congelar la interfaz gráfica
        def run_thread():
            try:
                generar_sistema_y_matriz(pdf, excel, carpetas, matriz)
                messagebox.showinfo("Éxito", "¡El proceso de consolidación de matriz y expedientes finalizó correctamente!")
            except Exception as e:
                print(f"\n❌ Error durante el proceso: {e}")
                messagebox.showerror("Error", f"Ocurrió un error inesperado:\n{e}")
            finally:
                self.btn_ejecutar.config(state='normal', text="Iniciar Consolidación 🚀", bg="#a6e3a1")
                
        threading.Thread(target=run_thread, daemon=True).start()

    def __del__(self):
        # Restaurar sys.stdout al destruir la app
        sys.stdout = self.stdout_orig

if __name__ == "__main__":
    root = tk.Tk()
    app = GeneradorGUI(root)
    
    def on_closing():
        sys.stdout = app.stdout_orig
        root.destroy()
        
    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()
