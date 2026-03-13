# Guía de Integración — Extractor de Recetas
## Proyecto: Color (Colorimetria) · C# 7.3

---

## Archivos generados

| Archivo | Qué hace |
|---|---|
| `ShadeReportExtractor.cs` | Clase principal. OCR + extracción de receta e LAB. **Agregar al proyecto.** |
| `Form1_IntegracionReceta.cs` | Parche con los cambios a aplicar en `Form1.cs`. **No reemplazar, solo fusionar.** |

---

## Paso 1 — Agregar ShadeReportExtractor.cs

1. En Visual Studio, clic derecho sobre el proyecto **Color**
2. **Agregar → Elemento existente**
3. Seleccionar `ShadeReportExtractor.cs`

---

## Paso 2 — Fusionar cambios en Form1.cs

Abre `Form1_IntegracionReceta.cs` y aplica cada sección marcada con `► AGREGAR` o `► MODIFICAR`:

### 2a. Campos de instancia
Dentro de la clase `Form1`, agrega:
```csharp
private readonly ShadeReportExtractor _shadeExtractor =
    new ShadeReportExtractor(@".\tessdata");

private ShadeExtractionResult _lastShadeResult;
private Bitmap _bitmapMediciones;
private Bitmap _bitmapReceta;
```

### 2b. Reemplazar BtnIniciarEscaneo_Click
Sustituye el cuerpo de tu botón "Iniciar escaneo" con el método del parche.  
El nuevo flujo es:
1. Valida que ambas imágenes estén cargadas
2. OCR imagen derecha → extrae **receta** con `ShadeReportExtractor`
3. OCR imagen izquierda → extrae **mediciones** con `ColorimetricDataExtractor`
4. Abre `FormResultados` con el texto combinado

### 2c. Agregar métodos nuevos
Copia al final de la clase `Form1` (antes del `}` de cierre):
- `MostrarResultadosCombinados(...)`
- `BuildResumenReceta(...)`
- `CargarImagenMediciones(...)`
- `CargarImagenReceta(...)`

---

## Paso 3 — Conectar la carga de imágenes

Busca en tu `Form1.cs` dónde se asignan las imágenes a los PictureBox.  
En ese punto llama a los nuevos métodos:

```csharp
// Cuando el usuario carga la imagen de MEDICIONES (izquierda)
CargarImagenMediciones(rutaSeleccionada);
picMediciones.Image = _bitmapMediciones;

// Cuando el usuario carga la imagen de RECETA (derecha)
CargarImagenReceta(rutaSeleccionada);
picReceta.Image = _bitmapReceta;
```

Si ya guardas las rutas en variables, simplemente asigna el Bitmap:
```csharp
_bitmapMediciones = new Bitmap(_rutaMediciones);
_bitmapReceta     = new Bitmap(_rutaReceta);
```

---

## Paso 4 — Verificar namespaces

En `ShadeReportExtractor.cs` el namespace es `Color` (igual que el proyecto).  
No necesitas agregar `using` adicionales en `Form1.cs`.

---

## Flujo completo al presionar "Iniciar escaneo"

```
[Imagen Mediciones]          [Imagen Receta]
       │                            │
ColorimetricDataExtractor    ShadeReportExtractor
       │                            │
 List<ColorimetricRow>      ShadeExtractionResult
       │                       ├── List<RecipeItem>
       │                       └── LabValues
       └──────────┬─────────────────┘
                  │
        MostrarResultadosCombinados()
                  │
           FormResultados
         ┌────────┴────────┐
   Mediciones/OCR     CMC/Recomendación
```

---

## Notas

- La regex de receta detecta el patrón: `8 dígitos + nombre + número%`  
  Ejemplo: `12345678  DISPERSE BLUE 3R  4.50%`
- Si la imagen tiene baja resolución, Tesseract puede fallar. Usa imágenes escaneadas a 300 DPI mínimo.
- Los valores LAB se extraen buscando la cabecera `L A B dL da dB cde P/F` y tomando la primera fila de datos siguiente.
