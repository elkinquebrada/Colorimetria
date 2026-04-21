using System.Collections.Generic;

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
    }

    public class OcrReportRequest
    {
        public double ChromaThreshold { get; set; }
        public List<ColorimetricRow> Measures { get; set; }
    }

    public class CorrectReportResponse
    {
        public bool Success { get; set; }
        public string Error { get; set; }
        public List<ColorimetricRow> CorrectedMeasures { get; set; }
    }
}