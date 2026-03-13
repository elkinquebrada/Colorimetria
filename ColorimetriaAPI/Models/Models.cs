// Models.cs — Modelos de datos de la API Colorimetría
namespace ColorimetriaAPI.Models
{
    // ── Request ────────────────────────────────────────────────────────────────

    /// <summary>
    /// Fila colorimétrica CIELab enviada desde el cliente Windows Forms.
    /// Solo contiene números — sin imagen, nombre de lote ni receta.
    /// </summary>
    public class ColorimetricRow
    {
        /// <summary>Iluminante: "D65", "TL84", "A"</summary>
        public string Illuminant { get; set; }

        /// <summary>Tipo: "Std" o "Lot"</summary>
        public string Type { get; set; }

        public double L { get; set; }
        public double A { get; set; }
        public double B { get; set; }
        public double Chroma { get; set; }
        public double Hue { get; set; }

        /// <summary>Marcado por el OCR local como sospechoso</summary>
        public bool NeedsReview { get; set; }
    }

    /// <summary>
    /// Reporte OCR completo enviado al endpoint /correct
    /// </summary>
    public class OcrReportRequest
    {
        /// <summary>Umbral de error de coherencia para detectar tokens erróneos (default 0.35)</summary>
        public double ChromaThreshold { get; set; } = 0.35;

        public List<ColorimetricRow> Measures { get; set; } = new();
    }

    // ── Response ───────────────────────────────────────────────────────────────

    /// <summary>
    /// Resultado de corrección de un token individual
    /// </summary>
    public class CorrectionResult
    {
        public string Field { get; set; }
        public string Illuminant { get; set; }
        public string Type { get; set; }
        public double OriginalValue { get; set; }
        public double CorrectedValue { get; set; }
        public double OriginalCoherenceError { get; set; }
        public double NewCoherenceError { get; set; }
        public bool Accepted { get; set; }
        public string Reason { get; set; }
    }

    public class CorrectReportResponse
    {
        public bool Success { get; set; }
        public int TokensAnalyzed { get; set; }
        public int TokensCorrected { get; set; }
        public List<CorrectionResult> Corrections { get; set; } = new();
        public List<ColorimetricRow> CorrectedMeasures { get; set; } = new();
        public string Error { get; set; }
    }

    // ── Token interno (detección) ──────────────────────────────────────────────

    public class ErrorToken
    {
        public string Field { get; set; }
        public string Illuminant { get; set; }
        public string Type { get; set; }
        public double OcrValue { get; set; }
        public double A { get; set; }
        public double B { get; set; }
        public double Chroma { get; set; }
        public double CoherenceError { get; set; }
    }

    public class ClaudeRawCorrection
    {
        public int Idx { get; set; }
        public double Corrected { get; set; }
        public string Reason { get; set; }
    }
}
