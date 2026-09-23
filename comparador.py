import pandas as pd
import os
import glob

def normalizar_status(valor):
    """
    Convierte los valores de status a un formato estándar para comparar.
    """
    if pd.isna(valor):
        return "sin_dato"
    
    val = str(valor).strip().lower()
    
    # Mapeo de equivalencias
    if val in ["con", "activo", "active", "activado"]:
        return "activo"
    elif val in ["soft", "cortado", "suspendido", "corte", "inactivo"]:
        return "cortado"
    else:
        return val

def listar_archivos_excel():
    """Lista todos los archivos Excel en el directorio actual"""
    archivos = glob.glob("*.xlsx") + glob.glob("*.xls")
    return sorted(archivos)

def mostrar_menu_archivos(archivos, titulo):
    """Muestra un menú numerado de archivos"""
    print(f"\n{titulo}")
    print("=" * 60)
    for i, archivo in enumerate(archivos, 1):
        print(f"{i}. {archivo}")
    print("=" * 60)

def seleccionar_archivo(archivos, mensaje):
    """Permite al usuario seleccionar un archivo del menú"""
    while True:
        try:
            seleccion = input(f"\n{mensaje} (número): ").strip()
            indice = int(seleccion) - 1
            
            if 0 <= indice < len(archivos):
                return archivos[indice]
            else:
                print("❌ Número inválido. Intenta de nuevo.")
        except ValueError:
            print("❌ Por favor ingresa un número válido.")

def comparar_servicios(archivo1, archivo2, archivo_salida):
    print(f"\n📂 Archivo 1: {archivo1}")
    print(f"📂 Archivo 2: {archivo2}")
    print("Cargando archivos...")
    
    # 1. Cargar los archivos Excel
    df1 = pd.read_excel(archivo1)
    df2 = pd.read_excel(archivo2)

    # 2. Normalizar nombres de columnas (minúsculas y sin espacios extra)
    df1.columns = df1.columns.str.strip().str.lower()
    df2.columns = df2.columns.str.strip().str.lower()

    # 3. CONFIGURACIÓN DE COLUMNAS
    cols_id = ['abonado', 'serial'] 
    
    # Buscar automáticamente la columna de status en cada archivo
    col_status_1 = None
    col_status_2 = None
    
    # Buscar en archivo 1
    for col in df1.columns:
        if 'status' in col or 'operativo' in col:
            col_status_1 = col
            break
    
    # Buscar en archivo 2
    for col in df2.columns:
        if 'sts' in col or 'ctto' in col or 'status' in col:
            col_status_2 = col
            break

    # Verificar que las columnas existan
    for col in cols_id:
        if col not in df1.columns or col not in df2.columns:
            print(f"\n❌ Error: La columna '{col}' no se encontró en ambos archivos.")
            print(f"Columnas en Archivo 1: {list(df1.columns)}")
            print(f"Columnas en Archivo 2: {list(df2.columns)}")
            return False
    
    if not col_status_1:
        print(f"\n❌ Error: No se encontró columna de status en Archivo 1.")
        print(f"Columnas disponibles: {list(df1.columns)}")
        return False
    
    if not col_status_2:
        print(f"\n❌ Error: No se encontró columna de status en Archivo 2.")
        print(f"Columnas disponibles: {list(df2.columns)}")
        return False

    print(f"✅ Columna status Archivo 1: '{col_status_1}'")
    print(f"✅ Columna status Archivo 2: '{col_status_2}'")

    # 4. Limpiar identificadores
    for col in cols_id:
        df1[col] = df1[col].astype(str).str.strip()
        df2[col] = df2[col].astype(str).str.strip()

    print("🔄 Normalizando estados y cruzando información...")

    # 5. Crear columnas normalizadas
    df1['status_normalizado'] = df1[col_status_1].apply(normalizar_status)
    df2['status_normalizado'] = df2[col_status_2].apply(normalizar_status)

    # 6. Unir ambos archivos
    merged = pd.merge(
        df1, df2, 
        on=cols_id, 
        how='outer', 
        suffixes=('_arch1', '_arch2'), 
        indicator=True
    )

    # 7. Función para evaluar diferencias
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

    # 8. Filtrar solo las diferencias
    diferencias = merged[merged['Resultado_Comparacion'].notna()].copy()

    # 9. Exportar resultados
    if not diferencias.empty:
        diferencias = diferencias.drop(columns=['_merge', 'status_normalizado_arch1', 'status_normalizado_arch2'])
        
        diferencias.to_excel(archivo_salida, index=False)
        print(f"\n✅ ¡Proceso terminado con éxito!")
        print(f"🔍 Se encontraron {len(diferencias)} servicios con diferencias.")
        print(f"💾 Reporte guardado en: {os.path.abspath(archivo_salida)}")
        return True
    else:
        print("\n✅ ¡Proceso terminado!")
        print("🎉 No se encontraron diferencias. Los datos de status son idénticos en ambos archivos.")
        return True

# ==========================================
# MENÚ INTERACTIVO PRINCIPAL
# ==========================================
if __name__ == "__main__":
    print("=" * 60)
    print("🔍 COMPARADOR DE ARCHIVOS EXCEL")
    print("=" * 60)
    
    while True:
        # Listar archivos Excel disponibles
        archivos_excel = listar_archivos_excel()
        
        if len(archivos_excel) < 2:
            print("\n❌ No hay suficientes archivos Excel en esta carpeta.")
            print("Por favor coloca al menos 2 archivos .xlsx en esta carpeta.")
            print(f"\nArchivos encontrados: {archivos_excel}")
            input("\nPresiona Enter para salir...")
            break
        
        # Mostrar menú
        mostrar_menu_archivos(archivos_excel, "Archivos Excel disponibles")
        
        # Seleccionar archivo 1
        archivo1 = seleccionar_archivo(archivos_excel, "Selecciona el PRIMER archivo")
        
        # Seleccionar archivo 2 (no puede ser el mismo)
        while True:
            archivo2 = seleccionar_archivo(archivos_excel, "Selecciona el SEGUNDO archivo")
            if archivo2 != archivo1:
                break
            print("❌ No puedes seleccionar el mismo archivo dos veces. Elige otro.")
        
        # Nombre del archivo de salida
        archivo_salida = "reporte_diferencias.xlsx"
        
        # Confirmar
        print("\n" + "=" * 60)
        print("📋 RESUMEN DE SELECCIÓN")
        print("=" * 60)
        print(f"Archivo 1: {archivo1}")
        print(f"Archivo 2: {archivo2}")
        print(f"Reporte:   {archivo_salida}")
        print("=" * 60)
        
        confirmar = input("\n¿Deseas continuar con la comparación? (s/n): ").strip().lower()
        
        if confirmar == 's':
            try:
                comparar_servicios(archivo1, archivo2, archivo_salida)
            except Exception as e:
                print(f"\n❌ Ocurrió un error inesperado: {e}")
        else:
            print("\n❌ Operación cancelada.")
        
        # Preguntar si quiere hacer otra comparación
        otra = input("\n¿Deseas hacer otra comparación? (s/n): ").strip().lower()
        if otra != 's':
            print("\n¡Hasta luego!")
            break
        
        print("\n" + "=" * 60)