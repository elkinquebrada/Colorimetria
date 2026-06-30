# Manual de Arquitectura e Integración — Colorimetría
## Proyecto: Color (Colorimetria) · C# 7.3

---

## 1. Introducción
Este documento describe la arquitectura modular del sistema de colorimetría, diseñado para la extracción, validación y corrección de recetas textiles mediante OCR y modelos matemáticos.

---

## 2. Arquitectura del Sistema

El proyecto sigue una estructura modular para facilitar el mantenimiento y la escalabilidad:

### A. Extractores (`Color.extractors`)
Encargados de la visión artificial y OCR.
- **`ShadeReportExtractor.cs`**: Especializado en el formato "Shade History Report". Extrae ingredientes de la receta (Código, Nombre, %) y valores LAB del estándar.
- **`ColorimetricDataExtractor.cs`**: Extrae tablas de mediciones (L*, a*, b*, Chroma, Hue) para múltiples iluminantes (D65, TL84, A).
- **`Opencvtabledetector.cs`**: Utiliza OpenCV para segmentar y detectar regiones de interés en las imágenes.

### B. Lógica y Cálculos (`Color`)
- **`ColorimetricCalculator.cs`**: Calcula deltas y diferencias de color entre el estándar y la muestra.
- **`RecipeCorrector.cs`**: Motor lógico que aplica heurísticas para corregir porcentajes de colorantes basándose en las desviaciones de LAB.

### C. Servicios (`Color.Services`)
- **`HistorialService.cs`**: Gestiona la persistencia de datos en el archivo `DB_Colorimetria.csv`. Permite guardar mediciones detalladas y exportar el historial completo.

### D. Interfaz de Usuario (`Color.Forms`)
- **`Form1.cs`**: Ventana principal para carga de imágenes y control de flujo.
- **`FormConfirmacionOCR.cs`**: Interfaz de validación humana. Permite corregir errores de lectura del OCR antes del cálculo.
- **`FormResultados.cs`**: Dashboard final con comparativas y recomendaciones de corrección.
- **`FormHistorial.cs`**: Visualizador de la base de datos histórica.

---

## 3. Flujo Crítico de Trabajo

El sistema opera en un ciclo de 4 pasos principales:

1.  **Captura y Validación**:
    - El usuario carga dos imágenes PNG (Mediciones y Receta).
    - `Form1` valida el formato de la receta en tiempo real usando `ShadeReportExtractor`.
2.  **Procesamiento OCR & Confirmación**:
    - Se extraen los datos de ambas imágenes.
    - Se abre `FormConfirmacionOCR` donde el usuario verifica que los números sean correctos (validación decimal y numérica integrada).
3.  **Análisis y Resultados**:
    - `ColorimetricCalculator` genera los deltas.
    - `RecipeCorrector` genera la nueva receta sugerida.
    - Se visualiza todo en `FormResultados`.
4.  **Persistencia**:
    - El usuario puede guardar el análisis en el historial (`HistorialService`), alimentando la base de datos para futuros análisis de tendencias.

---

## 4. Gestión de Historial (Base de Datos)

El historial se almacena en `DB_Colorimetria.csv` en la carpeta raíz del ejecutable.

- **Formato**: CSV delimitado por punto y coma (`;`) con codificación UTF-8.
- **Columnas**:
  - `FechaHora`, `ShadeName`, `Iluminante`, `Lightness`, `Chroma`, `Diagnostico_L`, `Correccion_L`, etc.
- **Exportación**: El sistema permite exportar a Excel multihioja (en desarrollo/integración avanzada) y CSV estructurado.

---

## 5. Dependencias Técnicas

Para que el sistema compile y funcione, se requieren los siguientes componentes en el directorio de salida:

- **`tessdata/`**: Carpeta con los modelos de lenguaje de Tesseract (ej. `spa.traineddata`, `eng.traineddata`).
- **DLLs de OpenCV**: Necesarias para el preprocesamiento de imágenes.
- **`Newtonsoft.Json`**: Para manejo de configuraciones y datos estructurados.

---

## 6. Notas de Mantenimiento

- **Regex de Receta**: Si el formato del reporte de Coats cambia, la expresión regular en `ShadeReportExtractor.cs` debe ser actualizada.
- **Tolerancias**: Las tolerancias de color se configuran a través del botón "Tolerancias" y afectan los diagnósticos de "Pasa/Falla".
- **Resolución**: Se recomienda capturar imágenes a 300 DPI para maximizar la precisión del OCR.

---

## 7. Motores Matemáticos y Paridad Excel

Para asegurar la paridad 1:1 con los estándares de Coats en Excel, el sistema implementa las siguientes lógicas:

### A. Variación Relativa (Motor de Decisión)
Utilizada para determinar el impacto porcentual en la receta:
- **Fórmula**: `Variación = (Estándar - Lote) / Estándar`
- **Precisión**: Se maneja con 8 decimales en el motor (`decimal`) y se redondea para visualización.

### B. Diferencias de Color (Deltas)
Utilizados para la representación gráfica y diagnósticos visuales:
- `ΔL = Lote_L - Estándar_L` (Luminosidad)
- `Δa = Lote_a - Estándar_a` (Eje Rojo/Verde)
- `Δb = Lote_b - Estándar_b` (Eje Amarillo/Azul)

### C. Parámetros CMC (l:c)
El sistema utiliza el estándar industrial CMC(2:1) para la evaluación de "Pasa/Falla":
- **Semi-ejes (S_l, S_c, S_h)**:
  - `S_l`: 0.511 (si L < 16) o `(0.040975 * L) / (1.0 + 0.01765 * L)`
  - `S_c`: `(0.0638 * C) / (1.0 + 0.0131 * C) + 0.638`
  - `S_h`: `S_c * (f * T + 1.0 - f)`
  - Donde `f = √(C^4 / (C^4 + 1900))` y `T` depende del ángulo de matiz `h`.

---
*Ultima actualización: Mayo 2026*
