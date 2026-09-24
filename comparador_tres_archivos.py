import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import threading

# ============================================================
# FUNCIÓN PARA CARGAR ARCHIVOS
# ============================================================
def cargar_archivo(ruta, sheet_name=None):
    """Carga archivo Excel o CSV"""
    extension = os.path.splitext(ruta)[1].lower()
    
    if extension in ['.xlsx', '.xls']:
        if sheet_name:
            return pd.read_excel(ruta, sheet_name=sheet_name)
        else:
            return pd.read_excel(ruta, sheet_name=0)
    elif extension == '.csv':
        try:
            return pd.read_csv(ruta, sep=None, engine='python', encoding='utf-8')
        except Exception:
            return pd.read_csv(ruta, sep=None, engine='python', encoding='latin-1')
    else:
        raise ValueError(f"Extensión no soportada: {extension}")

# ============================================================
# NORMALIZACIÓN DE STATUS
# ============================================================
def normalizar_status(valor):
    """Convierte CON/SOFT y variantes a estándar"""
    if pd.isna(valor) or str(valor).strip() == "":
        return "sin_dato"
    
    val = str(valor).strip().upper()
    
    if val in ["CON", "ACTIVO", "ACTIVE", "ACTIVADO", "ENABLED", "UP", "ONLINE"]:
        return "ACTIVO"
    elif val in ["SOFT", "CORTADO", "SUSPENDIDO", "CORTE", "INACTIVO", "DISABLED", "DOWN", "OFFLINE"]:
        return "CORTADO"
    else:
        return val

# ============================================================
# CLASE PRINCIPAL
# ============================================================
class ComparadorTresArchivosGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("🔍 Comparador BOSS - GTAI - ACS")
        self.root.geometry("950x750")
        
        self.archivo_boss = tk.StringVar()
        self.archivo_gtai = tk.StringVar()
        self.archivo_acs = tk.StringVar()
        
        self.crear_interfaz()
    
    def crear_interfaz(self):
        titulo = tk.Label(self.root, 
                         text="🔍 COMPARADOR DE SERVICIOS\nBOSS → GTAI → ACS", 
                         font=("Arial", 14, "bold"))
        titulo.pack(pady=10)
        
        desc = tk.Label(self.root, 
                       text="GTAI tiene 2 pestañas: 'Redes Propias' y 'Redes Compartidas'\n"
                            "Redes Propias: BOSS + GTAI + ACS\n"
                            "Redes Compartidas: Solo BOSS + GTAI\n"
                            "Formatos: .xlsx, .xls, .csv",
                       font=("Arial", 9), justify="center")
        desc.pack(pady=5)
        
        frame1 = tk.LabelFrame(self.root, text="1️⃣ Archivo BOSS", 
                               font=("Arial", 10, "bold"), padx=10, pady=5)
        frame1.pack(fill=tk.X, padx=20, pady=5)
        tk.Entry(frame1, textvariable=self.archivo_boss, width=70).pack(side=tk.LEFT, padx=5)
        tk.Button(frame1, text="📁 Seleccionar", command=self.seleccionar_boss).pack(side=tk.LEFT)
        
        frame2 = tk.LabelFrame(self.root, text="2️⃣ Archivo GTAI (con pestañas)", 
                               font=("Arial", 10, "bold"), padx=10, pady=5)
        frame2.pack(fill=tk.X, padx=20, pady=5)
        tk.Entry(frame2, textvariable=self.archivo_gtai, width=70).pack(side=tk.LEFT, padx=5)
        tk.Button(frame2, text="📁 Seleccionar", command=self.seleccionar_gtai).pack(side=tk.LEFT)
        
        frame3 = tk.LabelFrame(self.root, text="3️⃣ Archivo ACS", 
                               font=("Arial", 10, "bold"), padx=10, pady=5)
        frame3.pack(fill=tk.X, padx=20, pady=5)
        tk.Entry(frame3, textvariable=self.archivo_acs, width=70).pack(side=tk.LEFT, padx=5)
        tk.Button(frame3, text="📁 Seleccionar", command=self.seleccionar_acs).pack(side=tk.LEFT)
        
        self.btn_ejecutar = tk.Button(self.root, text="🚀 EJECUTAR COMPARACIÓN", 
                                      command=self.ejecutar_comparacion,
                                      font=("Arial", 12, "bold"),
                                      bg="#4CAF50", fg="white",
                                      padx=20, pady=10)
        self.btn_ejecutar.pack(pady=15)
        
        self.progress = ttk.Progressbar(self.root, mode='indeterminate', length=400)
        self.progress.pack(pady=5)
        
        self.texto_resultados = scrolledtext.ScrolledText(self.root, width=110, height=20, 
                                                          font=("Consolas", 9))
        self.texto_resultados.pack(padx=20, pady=10, fill=tk.BOTH, expand=True)
        
        tk.Button(self.root, text="🗑️ Limpiar Log", command=self.limpiar_resultados).pack(pady=5)
        
        self.mostrar_mensaje("👋 Bienvenido al Comparador BOSS - GTAI - ACS")
        self.mostrar_mensaje("📌 Selecciona los 3 archivos y presiona EJECUTAR\n")
    
    def seleccionar_boss(self):
        archivo = filedialog.askopenfilename(
            title="Selecciona el archivo BOSS",
            filetypes=[("Archivos Excel/CSV", "*.xlsx *.xls *.csv"), ("Todos", "*.*")]
        )
        if archivo:
            self.archivo_boss.set(archivo)
            self.mostrar_mensaje(f"✅ BOSS: {os.path.basename(archivo)}")
    
    def seleccionar_gtai(self):
        archivo = filedialog.askopenfilename(
            title="Selecciona el archivo GTAI",
            filetypes=[("Archivos Excel/CSV", "*.xlsx *.xls *.csv"), ("Todos", "*.*")]
        )
        if archivo:
            self.archivo_gtai.set(archivo)
            self.mostrar_mensaje(f"✅ GTAI: {os.path.basename(archivo)}")
    
    def seleccionar_acs(self):
        archivo = filedialog.askopenfilename(
            title="Selecciona el archivo ACS",
            filetypes=[("Archivos Excel/CSV", "*.xlsx *.xls *.csv"), ("Todos", "*.*")]
        )
        if archivo:
            self.archivo_acs.set(archivo)
            self.mostrar_mensaje(f"✅ ACS: {os.path.basename(archivo)}")
    
    def mostrar_mensaje(self, mensaje):
        self.texto_resultados.insert(tk.END, mensaje + "\n")
        self.texto_resultados.see(tk.END)
        self.root.update_idletasks()
    
    def limpiar_resultados(self):
        self.texto_resultados.delete(1.0, tk.END)
    
    def ejecutar_comparacion(self):
        b = self.archivo_boss.get()
        g = self.archivo_gtai.get()
        a = self.archivo_acs.get()
        
        if not b or not g or not a:
            messagebox.showerror("Error", "Debes seleccionar los 3 archivos")
            return
        
        self.btn_ejecutar.config(state=tk.DISABLED, text="Procesando...")
        self.progress.start(10)
        self.mostrar_mensaje("\n" + "="*70)
        self.mostrar_mensaje(" Iniciando comparación...")
        
        hilo = threading.Thread(target=self.comparar_tres_archivos, args=(b, g, a))
        hilo.start()
    
    def buscar_hoja(self, ruta_gtai, palabra_clave):
        """Busca una hoja que contenga la palabra clave (ignora mayúsculas/minúsculas)"""
        with pd.ExcelFile(ruta_gtai) as xls:
            for hoja in xls.sheet_names:
                if palabra_clave.lower() in hoja.lower():
                    return hoja  # Retorna el nombre ORIGINAL de la hoja
        return None
    
    def comparar_tres_archivos(self, ruta_boss, ruta_gtai, ruta_acs):
        try:
            # ===== PASO 1: CARGAR ARCHIVOS =====
            self.mostrar_mensaje("\n PASO 1: Cargando archivos...")
            
            df_boss = cargar_archivo(ruta_boss)
            df_boss.columns = df_boss.columns.str.strip().str.lower()
            self.mostrar_mensaje(f"   ✅ BOSS: {len(df_boss)} registros")
            
            ruta_gtai_path = ruta_gtai
            extension_gtai = os.path.splitext(ruta_gtai_path)[1].lower()
            
            if extension_gtai in ['.xlsx', '.xls']:
                self.mostrar_mensaje(f"    Buscando hojas en GTAI...")
                
                hoja_propias = self.buscar_hoja(ruta_gtai_path, 'propia')
                hoja_compartidas = self.buscar_hoja(ruta_gtai_path, 'compartida')
                
                if hoja_propias:
                    df_gtai_propias = cargar_archivo(ruta_gtai_path, sheet_name=hoja_propias)
                    df_gtai_propias.columns = df_gtai_propias.columns.str.strip().str.lower()
                    self.mostrar_mensaje(f"   ✅ Redes Propias ({hoja_propias}): {len(df_gtai_propias)} registros")
                else:
                    self.mostrar_mensaje("   ❌ No se encontró hoja de 'Redes Propias'")
                    df_gtai_propias = pd.DataFrame()
                
                if hoja_compartidas:
                    df_gtai_compartidas = cargar_archivo(ruta_gtai_path, sheet_name=hoja_compartidas)
                    df_gtai_compartidas.columns = df_gtai_compartidas.columns.str.strip().str.lower()
                    self.mostrar_mensaje(f"   ✅ Redes Compartidas ({hoja_compartidas}): {len(df_gtai_compartidas)} registros")
                else:
                    self.mostrar_mensaje("   ❌ No se encontró hoja de 'Redes Compartidas'")
                    df_gtai_compartidas = pd.DataFrame()
            else:
                self.mostrar_mensaje("   ⚠️ GTAI es CSV, cargando todo como Redes Propias")
                df_gtai_propias = cargar_archivo(ruta_gtai_path)
                df_gtai_propias.columns = df_gtai_propias.columns.str.strip().str.lower()
                df_gtai_compartidas = pd.DataFrame()
            
            df_acs = cargar_archivo(ruta_acs)
            df_acs.columns = df_acs.columns.str.strip().str.lower()
            self.mostrar_mensaje(f"   ✅ ACS: {len(df_acs)} registros")
            
            # ===== PASO 2: IDENTIFICAR COLUMNAS =====
            self.mostrar_mensaje("\n PASO 2: Identificando columnas...")
            
            col_boss_contrato = self.buscar_columna(df_boss, ['ctto ftth', 'ctto_ftth'])
            col_boss_serial = self.buscar_columna(df_boss, ['serial'])
            col_boss_status = self.buscar_columna(df_boss, ['sts ctto', 'sts_ctto'])
            col_boss_ciudad = self.buscar_columna(df_boss, ['nbr_ciudad', 'ciudad'])
            
            df_gtai_muestra = df_gtai_propias if not df_gtai_propias.empty else df_gtai_compartidas
            
            col_gtai_contrato = self.buscar_columna(df_gtai_muestra, ['docm_contrato', 'docm contrato'])
            col_gtai_serial = self.buscar_columna(df_gtai_muestra, ['cod pon', 'codpon', 'cod_pon'])
            col_gtai_status = self.buscar_columna(df_gtai_muestra, ['estatus', 'status'])
            
            col_acs_serial = self.buscar_columna(df_acs, ['serialnumber', 'serial number', 'serial'])
            col_acs_status = self.buscar_columna(df_acs, ['generalstatus', 'general status'])
            
            columnas_requeridas = {
                'BOSS contrato': col_boss_contrato, 'BOSS serial': col_boss_serial, 
                'BOSS status': col_boss_status, 'BOSS ciudad': col_boss_ciudad,
                'GTAI contrato': col_gtai_contrato, 'GTAI serial': col_gtai_serial, 
                'GTAI status': col_gtai_status,
                'ACS serial': col_acs_serial, 'ACS status': col_acs_status
            }
            
            faltantes = [k for k, v in columnas_requeridas.items() if v is None]
            if faltantes:
                self.mostrar_mensaje(f"\n❌ Columnas no encontradas: {faltantes}")
                self.finalizar()
                return
            
            self.mostrar_mensaje("   ✅ Todas las columnas identificadas")
            
            # ===== PASO 3: NORMALIZAR DATOS =====
            self.mostrar_mensaje("\n PASO 3: Normalizando datos...")
            
            for df, col in [(df_boss, col_boss_serial), (df_acs, col_acs_serial)]:
                df[col] = df[col].astype(str).str.strip().str.upper()
                df[col] = df[col].replace('NAN', '').replace('NONE', '')
            
            if not df_gtai_propias.empty:
                df_gtai_propias[col_gtai_serial] = df_gtai_propias[col_gtai_serial].astype(str).str.strip().str.upper()
                df_gtai_propias[col_gtai_serial] = df_gtai_propias[col_gtai_serial].replace('NAN', '').replace('NONE', '')
            
            if not df_gtai_compartidas.empty:
                df_gtai_compartidas[col_gtai_serial] = df_gtai_compartidas[col_gtai_serial].astype(str).str.strip().str.upper()
                df_gtai_compartidas[col_gtai_serial] = df_gtai_compartidas[col_gtai_serial].replace('NAN', '').replace('NONE', '')
            
            df_boss[col_boss_contrato] = df_boss[col_boss_contrato].astype(str).str.strip()
            if not df_gtai_propias.empty:
                df_gtai_propias[col_gtai_contrato] = df_gtai_propias[col_gtai_contrato].astype(str).str.strip()
            if not df_gtai_compartidas.empty:
                df_gtai_compartidas[col_gtai_contrato] = df_gtai_compartidas[col_gtai_contrato].astype(str).str.strip()
            
            df_boss[col_boss_ciudad] = df_boss[col_boss_ciudad].astype(str).str.strip().str.title()
            df_boss[col_boss_ciudad] = df_boss[col_boss_ciudad].replace('Nan', '-')
            
            df_boss['status_norm'] = df_boss[col_boss_status].apply(normalizar_status)
            if not df_gtai_propias.empty:
                df_gtai_propias['status_norm'] = df_gtai_propias[col_gtai_status].apply(normalizar_status)
            if not df_gtai_compartidas.empty:
                df_gtai_compartidas['status_norm'] = df_gtai_compartidas[col_gtai_status].apply(normalizar_status)
            df_acs['status_norm'] = df_acs[col_acs_status].apply(normalizar_status)
            
            # ===== PASO 4: VALIDAR REDES PROPIAS (BOSS + GTAI + ACS) =====
            self.mostrar_mensaje("\n" + "="*70)
            self.mostrar_mensaje(" PASO 4: Validando REDES PROPIAS (BOSS + GTAI + ACS)")
            self.mostrar_mensaje("="*70)
            
            status_definitivo_propias = []
            incongruencia_acs_propias = []
            propias_incongruencia = []
            propias_sin_match = []
            no_en_acs_propias = []
            
            if not df_gtai_propias.empty:
                boss_prep = df_boss[[col_boss_contrato, col_boss_serial, col_boss_status, col_boss_ciudad, 'status_norm']].copy()
                boss_prep.columns = ['contrato', 'serial', 'status_boss_orig', 'ciudad', 'status_boss']
                
                gtai_propias_prep = df_gtai_propias[[col_gtai_contrato, col_gtai_serial, col_gtai_status, 'status_norm']].copy()
                gtai_propias_prep.columns = ['contrato', 'serial', 'status_gtai_orig', 'status_gtai']
                
                merged_bg_propias = pd.merge(boss_prep, gtai_propias_prep, on=['contrato', 'serial'], how='outer', indicator=True)
                
                for _, row in merged_bg_propias.iterrows():
                    en_boss = row['_merge'] in ['both', 'left_only']
                    en_gtai = row['_merge'] in ['both', 'right_only']
                    
                    if row['_merge'] == 'both':
                        sb = str(row.get('status_boss', '')).strip()
                        sg = str(row.get('status_gtai', '')).strip()
                        
                        if sb == sg and sb not in ['nan', 'sin_dato', '']:
                            acs_prep = df_acs[[col_acs_serial, col_acs_status, 'status_norm']].copy()
                            acs_prep.columns = ['serial', 'status_acs_orig', 'status_acs']
                            
                            match_acs = acs_prep[acs_prep['serial'] == row['serial']]
                            
                            if not match_acs.empty:
                                sa = str(match_acs.iloc[0]['status_acs']).strip()
                                sa_orig = str(match_acs.iloc[0]['status_acs_orig']).strip()
                                
                                if sb == sa and sb not in ['nan', 'sin_dato', '']:
                                    status_definitivo_propias.append({
                                        'Serial': row['serial'],
                                        'Contrato': row.get('contrato', '-'),
                                        'Ciudad': row.get('ciudad', '-'),
                                        'Tipo Red': 'Propia',
                                        'Status BOSS': row.get('status_boss_orig', '-'),
                                        'Status GTAI': row.get('status_gtai_orig', '-'),
                                        'Status ACS': sa_orig,
                                        'Status Normalizado': sb,
                                        'Resultado': '✅ COINCIDEN LOS 3'
                                    })
                                else:
                                    incongruencia_acs_propias.append({
                                        'Serial': row['serial'],
                                        'Contrato': row.get('contrato', '-'),
                                        'Ciudad': row.get('ciudad', '-'),
                                        'Tipo Red': 'Propia',
                                        'Status BOSS': row.get('status_boss_orig', '-'),
                                        'Status GTAI': row.get('status_gtai_orig', '-'),
                                        'Status ACS': sa_orig,
                                        'Status Norm BOSS/GTAI': sb,
                                        'Status Norm ACS': sa,
                                        'Resultado': '⚠️ Status difiere en ACS'
                                    })
                            else:
                                no_en_acs_propias.append({
                                    'Serial': row['serial'],
                                    'Contrato': row.get('contrato', '-'),
                                    'Ciudad': row.get('ciudad', '-'),
                                    'Tipo Red': 'Propia',
                                    'Status BOSS': row.get('status_boss_orig', '-'),
                                    'Status GTAI': row.get('status_gtai_orig', '-'),
                                    'Status ACS': '-',
                                    'Resultado': '❌ Serial no encontrado en ACS'
                                })
                        else:
                            propias_incongruencia.append({
                                'Serial': row['serial'],
                                'Contrato': row.get('contrato', '-'),
                                'Ciudad': row.get('ciudad', '-'),
                                'Tipo Red': 'Propia',
                                'Status BOSS (original)': row.get('status_boss_orig', '-'),
                                'Status GTAI (original)': row.get('status_gtai_orig', '-'),
                                'Status BOSS (normalizado)': sb,
                                'Status GTAI (normalizado)': sg,
                                'Motivo': '⚠️ Status difiere entre BOSS y GTAI'
                            })
                    else:
                        propias_sin_match.append({
                            'Serial': row['serial'],
                            'Contrato': row.get('contrato', '-'),
                            'Ciudad': row.get('ciudad', '-') if en_boss else '-',
                            'Tipo Red': 'Propia',
                            'Existe en BOSS': '✅ Sí' if en_boss else '❌ No',
                            'Existe en GTAI': '✅ Sí' if en_gtai else '❌ No',
                            'Status BOSS': row.get('status_boss_orig', '-') if en_boss else '-',
                            'Status GTAI': row.get('status_gtai_orig', '-') if en_gtai else '-',
                            'Motivo': '❌ No existe en ambos archivos'
                        })
                
                self.mostrar_mensaje(f"   ✅ Status Definitivo: {len(status_definitivo_propias)}")
                self.mostrar_mensaje(f"   ⚠️ Incongruencia BOSS↔GTAI: {len(propias_incongruencia)}")
                self.mostrar_mensaje(f"   ⚠️ Incongruencia con ACS: {len(incongruencia_acs_propias)}")
                self.mostrar_mensaje(f"   ❌ Sin match BOSS↔GTAI: {len(propias_sin_match)}")
                self.mostrar_mensaje(f"   ❌ No existe en ACS: {len(no_en_acs_propias)}")
            else:
                self.mostrar_mensaje("   ️ No hay registros de redes propias")
            
            # ===== PASO 5: VALIDAR REDES COMPARTIDAS (Solo BOSS + GTAI) =====
            self.mostrar_mensaje("\n" + "="*70)
            self.mostrar_mensaje("🔗 PASO 5: Validando REDES COMPARTIDAS (Solo BOSS + GTAI)")
            self.mostrar_mensaje("="*70)
            
            compartidas_status_ok = []
            compartidas_incongruencia = []
            compartidas_sin_match = []
            
            if not df_gtai_compartidas.empty:
                boss_prep = df_boss[[col_boss_contrato, col_boss_serial, col_boss_status, col_boss_ciudad, 'status_norm']].copy()
                boss_prep.columns = ['contrato', 'serial', 'status_boss_orig', 'ciudad', 'status_boss']
                
                gtai_compartidas_prep = df_gtai_compartidas[[col_gtai_contrato, col_gtai_serial, col_gtai_status, 'status_norm']].copy()
                gtai_compartidas_prep.columns = ['contrato', 'serial', 'status_gtai_orig', 'status_gtai']
                
                merged_bg_compartidas = pd.merge(boss_prep, gtai_compartidas_prep, on=['contrato', 'serial'], how='outer', indicator=True)
                
                for _, row in merged_bg_compartidas.iterrows():
                    en_boss = row['_merge'] in ['both', 'left_only']
                    en_gtai = row['_merge'] in ['both', 'right_only']
                    
                    if row['_merge'] == 'both':
                        sb = str(row.get('status_boss', '')).strip()
                        sg = str(row.get('status_gtai', '')).strip()
                        
                        if sb == sg and sb not in ['nan', 'sin_dato', '']:
                            compartidas_status_ok.append({
                                'Serial': row['serial'],
                                'Contrato': row.get('contrato', '-'),
                                'Ciudad': row.get('ciudad', '-'),
                                'Tipo Red': 'Compartida',
                                'Status BOSS': row.get('status_boss_orig', '-'),
                                'Status GTAI': row.get('status_gtai_orig', '-'),
                                'Status ACS': '-',
                                'Status Normalizado': sb,
                                'Resultado': '✅ COINCIDEN BOSS y GTAI'
                            })
                        else:
                            compartidas_incongruencia.append({
                                'Serial': row['serial'],
                                'Contrato': row.get('contrato', '-'),
                                'Ciudad': row.get('ciudad', '-'),
                                'Tipo Red': 'Compartida',
                                'Status BOSS (original)': row.get('status_boss_orig', '-'),
                                'Status GTAI (original)': row.get('status_gtai_orig', '-'),
                                'Status BOSS (normalizado)': sb,
                                'Status GTAI (normalizado)': sg,
                                'Motivo': '⚠️ Status difiere entre BOSS y GTAI'
                            })
                    else:
                        compartidas_sin_match.append({
                            'Serial': row['serial'],
                            'Contrato': row.get('contrato', '-'),
                            'Ciudad': row.get('ciudad', '-') if en_boss else '-',
                            'Tipo Red': 'Compartida',
                            'Existe en BOSS': '✅ Sí' if en_boss else '❌ No',
                            'Existe en GTAI': '✅ Sí' if en_gtai else '❌ No',
                            'Status BOSS': row.get('status_boss_orig', '-') if en_boss else '-',
                            'Status GTAI': row.get('status_gtai_orig', '-') if en_gtai else '-',
                            'Motivo': '❌ No existe en ambos archivos'
                        })
                
                self.mostrar_mensaje(f"   ✅ Status Definitivo: {len(compartidas_status_ok)}")
                self.mostrar_mensaje(f"   ⚠️ Incongruencia: {len(compartidas_incongruencia)}")
                self.mostrar_mensaje(f"   ❌ Sin match: {len(compartidas_sin_match)}")
            else:
                self.mostrar_mensaje("   ⚠️ No hay registros de redes compartidas")
            
            # ===== PASO 6: GENERAR EXCEL =====
            self.mostrar_mensaje("\n PASO 6: Generando archivo Excel...")
            
            df_definitivo = pd.DataFrame(status_definitivo_propias + compartidas_status_ok)
            df_incongruencia = pd.DataFrame(incongruencia_acs_propias + propias_incongruencia + compartidas_incongruencia)
            df_sin_match = pd.DataFrame(propias_sin_match + no_en_acs_propias + compartidas_sin_match)
            
            carpeta = os.path.dirname(ruta_boss)
            archivo_salida = os.path.join(carpeta, "resultado_validacion_servicios.xlsx")
            
            with pd.ExcelWriter(archivo_salida, engine='openpyxl') as writer:
                if not df_definitivo.empty:
                    df_definitivo.to_excel(writer, sheet_name='✅ Status Definitivo', index=False)
                else:
                    pd.DataFrame({'Mensaje': ['No hay servicios con status definitivo']}).to_excel(writer, sheet_name='✅ Status Definitivo', index=False)
                
                if not df_incongruencia.empty:
                    df_incongruencia.to_excel(writer, sheet_name='⚠️ Incongruencias', index=False)
                else:
                    pd.DataFrame({'Mensaje': ['No hay incongruencias']}).to_excel(writer, sheet_name='⚠️ Incongruencias', index=False)
                
                if not df_sin_match.empty:
                    df_sin_match.to_excel(writer, sheet_name='❌ Sin Match', index=False)
                else:
                    pd.DataFrame({'Mensaje': ['Todos los seriales hicieron match']}).to_excel(writer, sheet_name='❌ Sin Match', index=False)
                
                resumen = pd.DataFrame({
                    'Categoría': [
                        '✅ Status Definitivo (Redes Propias)',
                        '✅ Status Definitivo (Redes Compartidas)',
                        '⚠️ Incongruencia BOSSGTAI (Propias)',
                        '️ Incongruencia con ACS (Propias)',
                        '⚠️ Incongruencia BOSS↔GTAI (Compartidas)',
                        '❌ Sin match BOSS↔GTAI (Propias)',
                        '❌ No existe en ACS (Propias)',
                        ' Sin match BOSS↔GTAI (Compartidas)',
                        'TOTAL PROCESADO'
                    ],
                    'Cantidad': [
                        len(status_definitivo_propias),
                        len(compartidas_status_ok),
                        len(propias_incongruencia),
                        len(incongruencia_acs_propias),
                        len(compartidas_incongruencia),
                        len(propias_sin_match),
                        len(no_en_acs_propias),
                        len(compartidas_sin_match),
                        len(status_definitivo_propias) + len(compartidas_status_ok) + len(propias_incongruencia) + 
                        len(incongruencia_acs_propias) + len(compartidas_incongruencia) + len(propias_sin_match) + 
                        len(no_en_acs_propias) + len(compartidas_sin_match)
                    ]
                })
                resumen.to_excel(writer, sheet_name='📊 Resumen', index=False)
            
            total = len(status_definitivo_propias) + len(compartidas_status_ok) + len(propias_incongruencia) + len(incongruencia_acs_propias) + len(compartidas_incongruencia) + len(propias_sin_match) + len(no_en_acs_propias) + len(compartidas_sin_match)
            
            self.mostrar_mensaje("\n" + "="*70)
            self.mostrar_mensaje("✅ ¡PROCESO COMPLETADO!")
            self.mostrar_mensaje("="*70)
            self.mostrar_mensaje(f"\n📊 REDES PROPIAS:")
            self.mostrar_mensaje(f"   ✅ Status Definitivo: {len(status_definitivo_propias)}")
            self.mostrar_mensaje(f"   ⚠️ Incongruencia BOSS↔GTAI: {len(propias_incongruencia)}")
            self.mostrar_mensaje(f"   ⚠️ Incongruencia con ACS: {len(incongruencia_acs_propias)}")
            self.mostrar_mensaje(f"   ❌ Sin match BOSS↔GTAI: {len(propias_sin_match)}")
            self.mostrar_mensaje(f"   ❌ No existe en ACS: {len(no_en_acs_propias)}")
            
            self.mostrar_mensaje(f"\n📊 REDES COMPARTIDAS:")
            self.mostrar_mensaje(f"   ✅ Status Definitivo: {len(compartidas_status_ok)}")
            self.mostrar_mensaje(f"   ⚠️ Incongruencia: {len(compartidas_incongruencia)}")
            self.mostrar_mensaje(f"   ❌ Sin match: {len(compartidas_sin_match)}")
            
            self.mostrar_mensaje(f"\n📊 TOTAL GENERAL: {total}")
            self.mostrar_mensaje(f"\n💾 Archivo guardado en:\n   {archivo_salida}")
            
            messagebox.showinfo("✅ Proceso Completado", 
                              f"✅ Definitivo: {len(status_definitivo_propias) + len(compartidas_status_ok)}\n"
                              f"⚠️ Incongruencias: {len(propias_incongruencia) + len(incongruencia_acs_propias) + len(compartidas_incongruencia)}\n"
                              f"❌ Sin Match: {len(propias_sin_match) + len(no_en_acs_propias) + len(compartidas_sin_match)}\n\n"
                              f"Archivo:\n{archivo_salida}")
            
        except Exception as e:
            self.mostrar_mensaje(f"\n❌ ERROR: {str(e)}")
            import traceback
            self.mostrar_mensaje(traceback.format_exc())
            messagebox.showerror("Error", f"Ocurrió un error:\n{str(e)}")
        finally:
            self.finalizar()
    
    def buscar_columna(self, df, nombres_posibles):
        """Busca una columna que coincida con alguno de los nombres"""
        for nombre in nombres_posibles:
            nombre_lower = nombre.lower()
            for col in df.columns:
                if col.lower() == nombre_lower:
                    return col
        for nombre in nombres_posibles:
            nombre_lower = nombre.lower()
            for col in df.columns:
                if nombre_lower in col.lower():
                    return col
        return None
    
    def finalizar(self):
        self.progress.stop()
        self.root.after(0, lambda: self.btn_ejecutar.config(
            state=tk.NORMAL, text="🚀 EJECUTAR COMPARACIÓN"))

if __name__ == "__main__":
    root = tk.Tk()
    app = ComparadorTresArchivosGUI(root)
    root.mainloop()