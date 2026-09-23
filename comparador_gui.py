import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import threading

def normalizar_status(valor):
    """Convierte los valores de status a un formato estándar"""
    if pd.isna(valor):
        return "sin_dato"
    
    val = str(valor).strip().lower()
    
    if val in ["con", "activo", "active", "activado"]:
        return "activo"
    elif val in ["soft", "cortado", "suspendido", "corte", "inactivo"]:
        return "cortado"
    else:
        return val

class ComparadorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Comparador de Archivos Excel")
        self.root.geometry("800x600")
        
        # Variables para las rutas de archivos
        self.archivo1 = tk.StringVar()
        self.archivo2 = tk.StringVar()
        
        self.crear_interfaz()
    
    def crear_interfaz(self):
        # Título
        titulo = tk.Label(self.root, text="🔍 COMPARADOR DE ARCHIVOS EXCEL", 
                         font=("Arial", 16, "bold"))
        titulo.pack(pady=10)
        
        # Frame para Archivo 1
        frame1 = tk.Frame(self.root)
        frame1.pack(fill=tk.X, padx=20, pady=5)
        
        tk.Label(frame1, text="Archivo 1:", font=("Arial", 10, "bold")).pack(side=tk.LEFT)
        tk.Entry(frame1, textvariable=self.archivo1, width=50).pack(side=tk.LEFT, padx=5)
        tk.Button(frame1, text="Seleccionar", command=self.seleccionar_archivo1).pack(side=tk.LEFT)
        
        # Frame para Archivo 2
        frame2 = tk.Frame(self.root)
        frame2.pack(fill=tk.X, padx=20, pady=5)
        
        tk.Label(frame2, text="Archivo 2:", font=("Arial", 10, "bold")).pack(side=tk.LEFT)
        tk.Entry(frame2, textvariable=self.archivo2, width=50).pack(side=tk.LEFT, padx=5)
        tk.Button(frame2, text="Seleccionar", command=self.seleccionar_archivo2).pack(side=tk.LEFT)
        
        # Botón de ejecutar
        self.btn_ejecutar = tk.Button(self.root, text="🚀 EJECUTAR COMPARACIÓN", 
                                      command=self.ejecutar_comparacion,
                                      font=("Arial", 12, "bold"),
                                      bg="#4CAF50", fg="white",
                                      padx=20, pady=10)
        self.btn_ejecutar.pack(pady=20)
        
        # Área de texto para mostrar resultados
        self.texto_resultados = scrolledtext.ScrolledText(self.root, width=90, height=20, 
                                                          font=("Courier", 9))
        self.texto_resultados.pack(padx=20, pady=10, fill=tk.BOTH, expand=True)
        
        # Botón para limpiar
        tk.Button(self.root, text="Limpiar", command=self.limpiar_resultados).pack(pady=5)
        
        # Mensaje inicial
        self.mostrar_mensaje("Bienvenido al Comparador de Archivos Excel\n")
        self.mostrar_mensaje("Selecciona los dos archivos que deseas comparar.\n")
    
    def seleccionar_archivo1(self):
        archivo = filedialog.askopenfilename(
            title="Selecciona el PRIMER archivo Excel",
            filetypes=[("Archivos Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")]
        )
        if archivo:
            self.archivo1.set(archivo)
            self.mostrar_mensaje(f"\n✅ Archivo 1 seleccionado: {os.path.basename(archivo)}")
    
    def seleccionar_archivo2(self):
        archivo = filedialog.askopenfilename(
            title="Selecciona el SEGUNDO archivo Excel",
            filetypes=[("Archivos Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")]
        )
        if archivo:
            self.archivo2.set(archivo)
            self.mostrar_mensaje(f"\n✅ Archivo 2 seleccionado: {os.path.basename(archivo)}")
    
    def mostrar_mensaje(self, mensaje):
        self.texto_resultados.insert(tk.END, mensaje + "\n")
        self.texto_resultados.see(tk.END)
        self.root.update_idletasks()
    
    def limpiar_resultados(self):
        self.texto_resultados.delete(1.0, tk.END)
    
    def ejecutar_comparacion(self):
        archivo1 = self.archivo1.get()
        archivo2 = self.archivo2.get()
        
        if not archivo1 or not archivo2:
            messagebox.showerror("Error", "Debes seleccionar ambos archivos")
            return
        
        if archivo1 == archivo2:
            messagebox.showerror("Error", "No puedes seleccionar el mismo archivo dos veces")
            return
        
        # Deshabilitar botón mientras procesa
        self.btn_ejecutar.config(state=tk.DISABLED, text="Procesando...")
        self.mostrar_mensaje("\n" + "="*60)
        self.mostrar_mensaje("🔄 Iniciando comparación...")
        
        # Ejecutar en un hilo separado para no congelar la interfaz
        hilo = threading.Thread(target=self.comparar_archivos, args=(archivo1, archivo2))
        hilo.start()
    
    def comparar_archivos(self, archivo1, archivo2):
        try:
            self.mostrar_mensaje(f"\n📂 Cargando: {os.path.basename(archivo1)}")
            df1 = pd.read_excel(archivo1)
            
            self.mostrar_mensaje(f"📂 Cargando: {os.path.basename(archivo2)}")
            df2 = pd.read_excel(archivo2)
            
            # Normalizar nombres de columnas
            df1.columns = df1.columns.str.strip().str.lower()
            df2.columns = df2.columns.str.strip().str.lower()
            
            # Identificadores
            cols_id = ['abonado', 'serial']
            
            # Buscar columnas de status
            col_status_1 = None
            col_status_2 = None
            
            for col in df1.columns:
                if 'status' in col or 'operativo' in col:
                    col_status_1 = col
                    break
            
            for col in df2.columns:
                if 'sts' in col or 'ctto' in col or 'status' in col:
                    col_status_2 = col
                    break
            
            # Verificar columnas
            for col in cols_id:
                if col not in df1.columns or col not in df2.columns:
                    self.mostrar_mensaje(f"\n❌ Error: Columna '{col}' no encontrada")
                    self.mostrar_mensaje(f"Columnas Archivo 1: {list(df1.columns)}")
                    self.mostrar_mensaje(f"Columnas Archivo 2: {list(df2.columns)}")
                    self.habilitar_boton()
                    return
            
            if not col_status_1:
                self.mostrar_mensaje(f"\n❌ No se encontró columna de status en Archivo 1")
                self.mostrar_mensaje(f"Columnas disponibles: {list(df1.columns)}")
                self.habilitar_boton()
                return
            
            if not col_status_2:
                self.mostrar_mensaje(f"\n❌ No se encontró columna de status en Archivo 2")
                self.mostrar_mensaje(f"Columnas disponibles: {list(df2.columns)}")
                self.habilitar_boton()
                return
            
            self.mostrar_mensaje(f"\n✅ Columna status Archivo 1: '{col_status_1}'")
            self.mostrar_mensaje(f"✅ Columna status Archivo 2: '{col_status_2}'")
            
            # Limpiar identificadores
            for col in cols_id:
                df1[col] = df1[col].astype(str).str.strip()
                df2[col] = df2[col].astype(str).str.strip()
            
            self.mostrar_mensaje("\n🔄 Normalizando estados y cruzando información...")
            
            # Normalizar status
            df1['status_normalizado'] = df1[col_status_1].apply(normalizar_status)
            df2['status_normalizado'] = df2[col_status_2].apply(normalizar_status)
            
            # Unir archivos
            merged = pd.merge(df1, df2, on=cols_id, how='outer', 
                            suffixes=('_arch1', '_arch2'), indicator=True)
            
            # Evaluar diferencias
            def evaluar_diferencia(row):
                if row['_merge'] == 'left_only':
                    return '⚠️ Solo existe en Archivo 1'
                elif row['_merge'] == 'right_only':
                    return '⚠️ Solo existe en Archivo 2'
                else:
                    status1 = str(row['status_normalizado_arch1']).strip()
                    status2 = str(row['status_normalizado_arch2']).strip()
                    
                    if status1 != status2:
                        orig1 = row[f'{col_status_1}_arch1']
                        orig2 = row[f'{col_status_2}_arch2']
                        return f'❌ Status Diferente: "{orig1}" vs "{orig2}"'
                
                return None
            
            merged['Resultado_Comparacion'] = merged.apply(evaluar_diferencia, axis=1)
            
            # Filtrar diferencias
            diferencias = merged[merged['Resultado_Comparacion'].notna()].copy()
            
            # Guardar resultados
            if not diferencias.empty:
                diferencias = diferencias.drop(columns=['_merge', 'status_normalizado_arch1', 
                                                       'status_normalizado_arch2'])
                
                # Guardar en la misma carpeta del archivo 1
                carpeta = os.path.dirname(archivo1)
                archivo_salida = os.path.join(carpeta, "reporte_diferencias.xlsx")
                
                diferencias.to_excel(archivo_salida, index=False)
                
                self.mostrar_mensaje(f"\n{'='*60}")
                self.mostrar_mensaje(f"✅ ¡Proceso terminado con éxito!")
                self.mostrar_mensaje(f"🔍 Se encontraron {len(diferencias)} servicios con diferencias")
                self.mostrar_mensaje(f"💾 Reporte guardado en:\n   {archivo_salida}")
                
                messagebox.showinfo("Éxito", 
                                  f"Se encontraron {len(diferencias)} diferencias.\n\n"
                                  f"Reporte guardado en:\n{archivo_salida}")
            else:
                self.mostrar_mensaje(f"\n{'='*60}")
                self.mostrar_mensaje("✅ ¡Proceso terminado!")
                self.mostrar_mensaje("🎉 No se encontraron diferencias")
                
                messagebox.showinfo("Éxito", "No se encontraron diferencias.\nLos datos son idénticos.")
            
        except Exception as e:
            self.mostrar_mensaje(f"\n❌ Error inesperado: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error:\n{str(e)}")
        
        finally:
            self.habilitar_boton()
    
    def habilitar_boton(self):
        self.root.after(0, lambda: self.btn_ejecutar.config(state=tk.NORMAL, text="🚀 EJECUTAR COMPARACIÓN"))

# Ejecutar la aplicación
if __name__ == "__main__":
    root = tk.Tk()
    app = ComparadorGUI(root)
    root.mainloop()