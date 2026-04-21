using System;
using System.Collections.Generic;
using System.Drawing;

namespace ColorimetriaAPI.Models
{
    public class ColorimetricRow
    {
        public string Illuminant { get; set; }
        public string Type { get; set; }
        public double L { get; set; }
        public double A { get; set; }
        public double B { get; set; }
        public double Chroma { get; set; }
        public double Hue { get; set; }
        public bool NeedsReview { get; set; }
    }

    public class OcrReportRequest
    {
        public double ChromaThreshold { get; set; } = 0.35;
        // Inicialización compatible con C# 7.3
        public List<ColorimetricRow> Measures { get; set; } = new List<ColorimetricRow>();
    }

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
        public List<CorrectionResult> Corrections { get; set; } = new List<CorrectionResult>();
        public List<ColorimetricRow> CorrectedMeasures { get; set; } = new List<ColorimetricRow>();
        public string Error { get; set; }
    }
}