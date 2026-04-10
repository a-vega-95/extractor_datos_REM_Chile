import openpyxl
import pandas as pd
import os
import sys

# 1. Configuración de Ingesta y Rutas
ANIO = "2025"
MESES = ["06", "12"]
BASE_DIR = r"C:\DESARROLLO\EXTRACCION_DATOS_REM\data_rem"
RUTA_DICCIONARIO = r"C:\DESARROLLO\EXTRACCION_DATOS_REM\diccionario_centros\COD_CENTROS_SALUD.CSV"
RUTA_SALIDA_CARPETA = r"C:\DESARROLLO\EXTRACCION_DATOS_REM\resumen\P"

# Configuración de extracción Serie P4 (Úlceras Activas Tratadas)
CONFIG_EXTRACCION = [
    {"fila": "65", "label": "Curación Convencional"},
    {"fila": "66", "label": "Curación Avanzada"},
    {"fila": "67", "label": "Con Ayuda Técnica de Descarga"}
]

GRUPOS_ETARIOS = [
    "15 a 19 años", "20 a 24 años", "25 a 29 años", "30 a 34 años", 
    "35 a 39 años", "40 a 44 años", "45 a 49 años", "50 a 54 años", 
    "55 a 59 años", "60 a 64 años", "65 a 69 años", "70 a 74 años", 
    "75 a 79 años", "80 y más años"
]

# 2. Carga de Diccionario de Centros
try:
    df_centros = pd.read_csv(RUTA_DICCIONARIO)
except Exception as e:
    print(f"Error al cargar el diccionario de centros: {e}")
    sys.exit(1)

datos_maestros = []

print(f"--- Iniciando Extracción Masiva Serie P ({ANIO}) ---")

# 3. Validación de Existencia y Procesamiento Persistente
for mes in MESES:
    print(f"\nVerificando Periodo: {mes}/{ANIO}...")
    
    for _, centro in df_centros.iterrows():
        codigo = str(centro['COD_CENTRO'])
        nombre = centro['NOMBRE']
        
        # Construir ruta esperada del archivo xlsm
        archivo_nombre = f"{codigo}P.xlsm"
        ruta_archivo = os.path.join(BASE_DIR, ANIO, "P", mes, archivo_nombre)
        
        # Validar si el archivo existe
        if not os.path.exists(ruta_archivo):
            # Reintento con extensión .xls
            ruta_archivo_alt = ruta_archivo.replace(".xlsm", ".xls")
            if os.path.exists(ruta_archivo_alt):
                ruta_archivo = ruta_archivo_alt
            else:
                print(f"\n[ERROR DE INGESTA] Proceso detenido.")
                print(f"Falta archivo: {archivo_nombre}")
                print(f"Centro: {codigo} - {nombre}")
                print(f"Periodo faltante: {mes}/{ANIO}")
                sys.exit(1)
        
        # Definir fecha de corte para el reporte
        fecha_corte = f"30/06/{ANIO}" if mes == "06" else f"31/12/{ANIO}"
        
        # Apertura de archivo y extracción
        print(f" > Extrayendo: {codigo} - {nombre}...", end="\r")
        try:
            wb = openpyxl.load_workbook(ruta_archivo, data_only=True)
            hoja = wb['P4']
            
            for config in CONFIG_EXTRACCION:
                num_fila = config["fila"]
                tratamiento_label = config["label"]
                
                # Obtener fila completa de datos (F a AG)
                celdas = hoja[f'F{num_fila}':f'AG{num_fila}'][0]

                for i, grupo in enumerate(GRUPOS_ETARIOS):
                    # Metadatos base (campos fijos)
                    meta_base = {
                        "Fecha Corte": fecha_corte,
                        "Codigo DEIS": codigo,
                        "Nombre Centro": nombre,
                        "Serie": "P",
                        "Hoja": "P4",
                        "Tratamiento": tratamiento_label,
                        "Grupo Etario": grupo
                    }
                    
                    # Hombre (índice par)
                    val_h = celdas[i*2].value
                    val_h = float(val_h) if isinstance(val_h, (int, float)) else 0
                    reg_h = meta_base.copy()
                    reg_h.update({
                        "Género": "Hombre", 
                        "Cantidad": val_h,
                        "Ruta Origen": ruta_archivo
                    })
                    datos_maestros.append(reg_h)
                    
                    # Mujer (índice impar)
                    val_m = celdas[i*2 + 1].value
                    val_m = float(val_m) if isinstance(val_m, (int, float)) else 0
                    reg_m = meta_base.copy()
                    reg_m.update({
                        "Género": "Mujer", 
                        "Cantidad": val_m,
                        "Ruta Origen": ruta_archivo
                    })
                    datos_maestros.append(reg_m)
        
        except Exception as e:
            print(f"\nError al procesar el archivo {archivo_nombre}: {e}")
            continue

print(f"\n\n--- Extracción finalizada con éxito ---")

# 4. Consolidación y Exportación
if datos_maestros:
    df_final = pd.DataFrame(datos_maestros)
    
    if not os.path.exists(RUTA_SALIDA_CARPETA):
        os.makedirs(RUTA_SALIDA_CARPETA)
    
    ruta_xlsx = os.path.join(RUTA_SALIDA_CARPETA, "p_ulceras_activas_pie_tratadas.xlsx")
    df_final.to_excel(ruta_xlsx, index=False, sheet_name="TRATAMIENTO")
    
    print(f"Producto consolidado generado en: {ruta_xlsx}")
    print(f"Total de registros procesados: {len(df_final)}")
else:
    print("No se encontraron datos para procesar.")