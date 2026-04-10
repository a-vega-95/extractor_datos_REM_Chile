# Extractor de Datos REM Chile

Este proyecto contiene una suite de herramientas en Python diseñadas para la extracción masiva y automatizada de datos desde formularios **REM (Resumen Estadístico Mensual)** del sistema de salud chileno.

## 🚀 Propósito
El objetivo principal es transformar datos complejos almacenados en formularios `.xlsm` y `.xls` en tablas estructuradas listas para ser consumidas por herramientas de análisis como **Power BI**. El sistema está optimizado para manejar grandes volúmenes de centros de salud y periodos históricos de forma consistente.

## 🛠️ Herramientas Incluidas (Scripts)

El proyecto se divide en módulos especializados por serie y temática:

1.  **`1_extractor_serie_p_riesgo_ulceracion.py`**: Extrae la evaluación vigente de riesgo de ulceración (Serie P, Hoja P4, Filas 61-64).
2.  **`2_extractor_serie_p_ulceras_activas_pie_tratadas.py`**: Captura datos sobre tratamientos de úlceras activas (Curación convencional, avanzada y ayudas técnicas).
3.  **`3_extractor_serie_p_amputacion_dm.py`**: Registro de amputaciones por pie diabético (Serie P, Hoja P4, Fila 68).
4.  **`4_extractor_serie_bm_curacion_avanzada.py`**: Motor de búsqueda inteligente para la Serie BM (Mensual) que identifica curaciones avanzadas desglosadas por tipo de establecimiento.

## ✨ Características Principales

*   **Búsqueda Inteligente**: Los scripts de la Serie BM no dependen de coordenadas fijas; buscan palabras clave (ej: *"Curación Avanzada Pie Diabético"*) para localizar los datos incluso si el formulario cambia de formato.
*   **Validación de Ingesta Persistente**: Si falta un archivo de algún centro o periodo según el diccionario maestro, el proceso se detiene y notifica el error exacto para garantizar la integridad de la base de datos.
*   **Normalización de Metadatos**: Todos los archivos generados incluyen automáticamente:
    *   Código DEIS y Nombre del Centro (vía diccionario).
    *   Fecha de Corte (calculada según periodo mensual o semestral).
    *   Serie y Hoja de origen.
    *   Ruta absoluta del archivo fuente para fines de auditoría.
*   **Unpivot de Datos**: Los datos se entregan en formato largo (long format), ideal para modelos de datos relacionales.

## 📂 Estructura del Proyecto

```text
├── src/                        # Scripts de extracción (.py)
├── diccionario_centros/        # COD_CENTROS_SALUD.CSV (Maestro de centros)
├── data_rem/                   # Directorio de entrada (Archivos XLSM originales)
│   ├── 2025/
│   │   ├── P/ (06, 12)
│   │   └── BM/ (01 a 12)
└── resumen/                    # Resultados procesados (.xlsx)
    ├── P/
    └── BM/
```

## 📋 Requisitos

*   **Python 3.10+**
*   Librerías necesarias:
    ```bash
    pip install pandas openpyxl
    ```

## ⚙️ Configuración

1.  Asegúrate de que tus archivos REM sigan la estructura de carpetas: `data_rem\AÑO\SERIE\MES\archivo.xlsm`.
2.  Mantén actualizado el archivo `diccionario_centros\COD_CENTROS_SALUD.CSV` con los códigos DEIS vigentes.

## 👤 Autor
*   **Anghello Vega** - [a-vega-95](https://github.com/a-vega-95)
