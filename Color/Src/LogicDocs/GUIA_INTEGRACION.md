# Manual de Arquitectura e Integración — Colorimetría
## Proyecto: Color (Colorimetria) · C# 7.3 · .NET Framework 4.8

---

## 1. Introducción

Este documento describe la arquitectura modular del sistema de colorimetría de Coats, diseñado para la extracción, validación y corrección de recetas textiles mediante visión artificial (OpenCV), OCR (Tesseract) y motores matemáticos con paridad total al estándar de Excel de Coats.

---

## 2. Arquitectura del Sistema

El proyecto sigue una estructura modular organizada en cuatro capas principales:

```
Color/
├── Src/
│   ├── extractors/         ← Visión artificial y OCR
│   └── Services/           ← Lógica de negocio, persistencia y utilidades
├── Presentation/
│   └── Forms/              ← Interfaz de usuario (WinForms)
└── Data/
    ├── Models/             ← Modelos de datos
    ├── Tolerancias/        ← Configuración de tolerancias
    ├── big_data/           ← Datos históricos grandes
    └── tesdata/            ← Modelos de idioma de Tesseract
```

---

### A. Extractores (`Color.Src.extractors`)

Responsables de la visión artificial y OCR sobre los reportes de imagen.

| Archivo | Responsabilidad |
|---|---|
| `ShadeReportExtractor.cs` | Formato "Shade History Report". Extrae ingredientes (Código, Nombre, %) y valores LAB del estándar via Regex + Tesseract. |
| `Dataextraxtor.cs` | Motor principal OCR. Extrae tablas de medición (L\*, a\*, b\*, Chroma, Hue) para iluminantes D65, TL84 y A. Incluye escalado adaptativo y binarización Otsu. |
| `DynamicSplitGridExtractor.cs` | Extrae datos de reportes "Bulk" con cuadrículas dinámicas multi-columna. |
| `Opencvtabledetector.cs` | Motor de visión artificial con OpenCV. Detecta y segmenta regiones de interés (tablas) en las imágenes del reporte. |
| `Formconfirmacionocr.cs` | *Nota: ubicado en `extractors/` por acoplamiento histórico.* Interfaz de validación humana de datos OCR. |

---

### B. Servicios (`Color.Src.Services`)

Contienen toda la lógica de negocio, persistencia y utilidades de soporte.

| Archivo | Responsabilidad |
|---|---|
| `Colorimetriccalculator.cs` | Calcula deltas (ΔL, Δa, Δb, ΔE) y diferencias de color entre el estándar y la muestra. |
| `RecipeCorrector.cs` | Motor lógico que aplica heurísticas para generar las recetas correctivas R1 (Luminosidad), R2 (Croma), R3 (Tono). |
| `ColorDecisionEngine.cs` | Motor de decisión de cuadrante. Implementa la lógica equivalente a las fórmulas `IF` y `XLOOKUP` del Excel de Coats para determinar el brillo (`Brighter/Duller`) y el color opuesto en el diagrama CIELAB. |
| `ReportFormatRouter.cs` | Clasificador de formato de imagen. Detecta automáticamente si el reporte es `LegacyCombinedFormat` o `DynamicSplitGridFormat` usando OCR rápido en el 15% superior de la imagen y heurística geométrica de Hough Lines como fallback. |
| `TextileMetadataExtractor.cs` | Extrae metadatos del encabezado del reporte (ShadeName, DyeingClass, Substrate, Count/Ply, FiberType) con un pipeline de 4 pases OCR progresivos y binarización Otsu adaptativa. |
| `Historialservices.cs` | Servicio de persistencia dual (ver Sección 4). |
| `ColorimetriaApiClient.cs` | Cliente REST opcional para un backend externo de corrección colorimétrica (`http://localhost:5000`). |

---

### C. Interfaz de Usuario (`Color.Presentation.Forms`)

| Archivo | Responsabilidad |
|---|---|
| `Form1.cs` | Ventana principal. Control de flujo, carga de imágenes, invocación del `ReportFormatRouter` y manejo de operaciones asíncronas (`async/await`). Todas las ventanas tienen `MinimizeBox = true`. |
| `Formresultados.cs` | Dashboard de análisis. Muestra Bloques 1 y 2 (datos CIELAB, deltas, diagnósticos) y la tabla de receta correctiva (R1, R2, R3) con colores diferenciados por colorante. |
| `FormHistorial.cs` | Visualizador del historial. Lee de SQL Server (V4) con fallback a CSV legacy. |
| `FormGraficoCielab.cs` | Formulario de gráfico CIELAB interactivo. |
| `CielabChartControl.cs` | Control custom que renderiza el diagrama CIELAB con precisión geométrica CIE. |
| `IluminantReportBlock.cs` | Bloque de UI reutilizable que representa los datos de un único iluminante (Chroma/Hue, Deltas, Desviaciones). |

---

### D. Modelos de Datos (`Color.Data`)

| Archivo | Responsabilidad |
|---|---|
| `Data/Models/TextileMetadata.cs` | DTO con los campos del encabezado del reporte: `ShadeName`, `DyeingClass`, `Substrate`, `CountPly`, `FiberType`. |
| `Data/Tolerancias/FormConfigTolerancias.cs` | Formulario de configuración de tolerancias. Gestiona 3 perfiles predefinidos (DE=0.60, 1.20, 1.80) y 1 perfil manual con cálculo automático de ejes. |

---

## 3. Flujo Crítico de Trabajo

El sistema opera en un ciclo de 5 pasos principales:

```
[Imágenes PNG] → [ReportFormatRouter] → [OCR + Extracción] → [FormConfirmacionOCR] → [Cálculo] → [Resultados + Guardado]
```

1. **Carga y Clasificación**:
   - El usuario carga dos imágenes PNG (Mediciones y Receta).
   - `Form1` invoca `ReportFormatRouter.DetermineFormat()` para determinar si el reporte es formato Legacy o DynamicSplitGrid.
   - `TextileMetadataExtractor` extrae automáticamente el encabezado (Shade Name, Lot No, etc.).

2. **Procesamiento OCR Asíncrono**:
   - La extracción de datos se ejecuta en un hilo de fondo (`Task.Run` + `async/await`) para no bloquear la UI.
   - Las imágenes se procesan con downscaling previo para reducir la carga de memoria.
   - Se extraen datos para los iluminantes D65, TL84 y A.

3. **Validación Humana**:
   - Se abre `FormConfirmacionOCR` con los datos extraídos.
   - El usuario verifica y corrige los valores numéricos OCR antes del cálculo.

4. **Análisis y Resultados**:
   - `ColorimetricCalculator` genera los deltas (ΔL, Δa, Δb, ΔChroma, ΔHue, ΔE).
   - `ColorDecisionEngine` determina el veredicto de cuadrante y brillo.
   - `RecipeCorrector` genera las 3 recetas correctivas (R1, R2, R3).
   - `FormResultados` muestra todo el reporte visual con paridad al Excel de Coats.

5. **Persistencia**:
   - El usuario puede guardar el análisis. `HistorialService.GuardarAnalisisCompleto()` persiste en SQL Server (V4).
   - `FormHistorial` permite consultar el historial histórico.

---

## 4. Gestión de Historial — Persistencia Dual (V4)

El sistema implementa una arquitectura de persistencia en dos capas:

### 4.1 Capa Principal: SQL Server (V4)

Conexión por defecto: `Server=(localdb)\MSSQLLocalDB;Database=ColorimetriaDB;Trusted_Connection=True;`

**Tabla `tbl_analisis_cabecera`**:
| Campo | Tipo | Descripción |
|---|---|---|
| `Id_Lote` | `INT IDENTITY` (PK) | Identificador único del análisis |
| `ShadeName` | `NVARCHAR` | Nombre del shade analizado |
| `LotNo` | `NVARCHAR` | Número de lote |
| `FechaRegistro` | `DATETIME` | Timestamp del análisis |
| `DeltaE_TL84` / `CMC_TL84` / `Status_TL84` | Numérico / `NVARCHAR` | Resultados iluminante TL84 |
| `DeltaE_A` / `CMC_A` / `Status_A` | Numérico / `NVARCHAR` | Resultados iluminante A |

**Tabla `tbl_analisis_detalle`** (FK → `Id_Lote`):
| Campo | Descripción |
|---|---|
| `DyeCode`, `DyeName` | Código y nombre del colorante |
| `Concentration_Original` | Concentración original (5 decimales, con `%`) |
| `Proportion_Original` | Proporción relativa (1 decimal, con `%`) |
| `R1_Con/Part/Ajuste_Percentage` | Receta R1 (Luminosidad) |
| `R2_Con/Part/Ajuste_Percentage` | Receta R2 (Croma) |
| `R3_Con/Part/Ajuste_Percentage` | Receta R3 (Tono) |

### 4.2 Capa Legacy: CSV (Fallback)

Archivo: `DB_Coats_Consolidado.csv` — ubicado en el directorio raíz del ejecutable.

- **Formato**: CSV delimitado por `;`, codificación UTF-8, modo Append-Only.
- **Compatibilidad**: El lector incluye lógica de migración para registros de 18, 20 y 21 columnas (evolución histórica del esquema).
- **Cuando se usa**: Como fallback cuando SQL Server no está disponible, o para exportación de datos planos.

---

## 5. Clasificación Automática de Formato de Reporte

`ReportFormatRouter` implementa un pipeline de detección en dos etapas:

1. **Etapa 1 — OCR rápido del 15% superior** (título del reporte):
   - Si contiene `BULK`, `CHEESES` o `COL GROUP` → `DynamicSplitGridFormat`
   - Si contiene `PASS / FAIL`, `SHADE HISTORY` o `EQUATION` → `LegacyCombinedFormat`

2. **Etapa 2 — Heurística geométrica** (si el OCR no es concluyente):
   - Aplica umbralización adaptativa y detección de líneas horizontales via `HoughLinesP` en el 50% inferior de la imagen.
   - Si hay ≥ 4 líneas horizontales estructurales → `LegacyCombinedFormat`.
   - En caso contrario → `DynamicSplitGridFormat`.

---

## 6. Motor de Decisión de Color (`ColorDecisionEngine`)

Implementa la paridad exacta con Excel de Coats:

- **Brillo** (`IF(H32>0, "Brighter", "Duller")`): Basado en el signo de ΔL.
- **Mapeo de cuadrante** (`XLOOKUP` en tabla de 8 cuadrantes): Diccionario de color opuesto para cada dirección en el espacio CIELAB (Yellower, Bluer, Greener, Redder con modificadores).

**Tabla de cuadrantes implementada**:
| Entrada | Color opuesto |
|---|---|
| Yellower (Greener) | Bluer (Redder) |
| Yellower (Redder) | Bluer (Greener) |
| Greener (Bluer) | Redder (Yellower) |
| Greener (Yellower) | Redder (Bluer) |
| Bluer (Redder) | Yellower (Greener) |
| Bluer (Greener) | Yellower (Redder) |
| Redder (Yellower) | Greener (Bluer) |
| Redder (Bluer) | Greener (Yellower) |

---

## 7. Motores Matemáticos y Paridad Excel

### A. Concentración y Proporción de Colorantes
- **Concentración**: Almacenada con 5 decimales de precisión (`F5`), visualizada con símbolo `%`.
- **Proporción**: `(Concentración_i / Σ Concentraciones) × 100`, presentada con 1 decimal (`F1%`).
- **Ajuste**: `|Receta_i / Concentración_original - 1| × 100` (variación relativa porcentual).

### B. Diferencias de Color (Deltas)
- `ΔL = Lote_L − Estándar_L` (Luminosidad)
- `Δa = Lote_a − Estándar_a` (Eje Rojo/Verde)
- `Δb = Lote_b − Estándar_b` (Eje Amarillo/Azul)
- `ΔE = √(ΔL² + Δa² + Δb²)` (Diferencia total)

### C. Parámetros CMC (2:1)
El sistema usa el estándar industrial **CMC(2:1)** para la evaluación de "PASS/FAIL":

- **S_l**: `0.511` (si L < 16) | `(0.040975 × L) / (1.0 + 0.01765 × L)`
- **S_c**: `(0.0638 × C) / (1.0 + 0.0131 × C) + 0.638`
- **S_h**: `S_c × (f × T + 1.0 − f)`
- Donde `f = √(C⁴ / (C⁴ + 1900))` y `T` depende del ángulo de matiz `h`.
- **CMC ΔE** = `√((ΔL/(l×S_l))² + (ΔC/S_c)² + (ΔH/S_h)²)`

### D. Fórmula de Tolerancias
Los ejes D_L, D_C y D_H se calculan a partir del ΔE global con la fórmula:
```
Eje = √(ΔE² / 3)
```
Esta formula equivale exactamente a la implementación del Excel de Coats.

---

## 8. Configuración de Tolerancias

`FormConfigTolerancias` ofrece 3 perfiles predefinidos y 1 perfil manual:

| Perfil | ΔE | ΔL | ΔC | ΔHue |
|---|---|---|---|---|
| Estricto | 0.60 | 0.346 | 0.346 | 0.346 |
| Estándar | 1.20 | 0.693 | 0.693 | 0.693 |
| Flexible | 1.80 | 1.039 | 1.039 | 1.039 |
| Manual | Usuario | Auto | Auto | Auto |

Los valores se persisten en `Properties.Settings.Default` y son leídos en tiempo real por `FormResultados`.

---

## 9. Dependencias Técnicas

| Dependencia | Uso | Notas de Despliegue |
|---|---|---|
| **Tesseract (`tessdata/`)** | OCR — modelos `eng.traineddata`, `spa.traineddata` | Carpeta `tessdata/` debe estar en el directorio del ejecutable. |
| **OpenCvSharp4** (DLLs nativas) | Visión artificial, detección de tabla, segmentación | Las DLLs de runtime Win x64 deben incluirse en el instalador. |
| **`System.Data.SqlClient`** | Persistencia SQL Server / LocalDB | Requiere LocalDB instalado en máquina destino. |
| **`Newtonsoft.Json`** | Manejo de configuraciones y datos estructurados | Incluido via NuGet. |
| **`ClosedXML` / `ExcelDataReader`** | Exportación y lectura de Excel | Incluido via NuGet. |

---

## 10. Notas de Mantenimiento

- **Regex de Receta**: Si el formato del reporte de Coats cambia, la expresión regular en `ShadeReportExtractor.cs` debe ser actualizada.
- **Nuevos Iluminantes**: Para agregar un iluminante nuevo, actualizar `Dataextraxtor.cs` y la función `UpdateData` en `IluminantReportBlock.cs`.
- **Clasificador de formato**: Si aparece un nuevo tipo de reporte, agregar la palabra clave en `ReportFormatRouter.DetermineFormat()`.
- **Resolución de imagen**: Se recomienda capturar imágenes a 300 DPI para maximizar la precisión del OCR. El sistema aplica escalado adaptativo (2x–4x) según la nitidez medida automáticamente.
- **API REST**: `ColorimetriaApiClient` es un módulo opcional preparado para integración futura con un backend de ML. Si el endpoint no está disponible, `IsApiAvailable()` retorna `false` y el sistema opera en modo standalone sin interrupciones.

---

## 11. Historial de Versiones

| Fecha | Versión | Cambios Principales |
|---|---|---|
| Mayo 2026 | V1–V3 | OCR básico, CSV legacy, Forms iniciales. |
| Junio 2026 | V4 | Migración a SQL Server (LocalDB), persistencia relacional, formulario historial, motor `ColorDecisionEngine`, `ReportFormatRouter`, `TextileMetadataExtractor`, exportación PDF, chart CIELAB, optimización asíncrona, deploy ClickOnce. |
| Julio 2026 | V4.1 | Habilitación de minimización en todos los formularios, corrección de DLLs en instalador ClickOnce, `FormConfigTolerancias` con perfiles de tolerancia y tarjeta manual, visualización `%` en columna Dye, separadores visuales en tabla de receta correctiva. |

---
*Última actualización: Julio 2026*
