using System;
using System.Collections.Generic;
using System.Linq;

namespace Color
{
    public sealed class ColorCorrectionResult
    {
        public string Illuminant { get; set; } = "";

        // Diferencias (Lot - Std)
        public double DeltaL { get; set; }
        public double DeltaA { get; set; }
        public double DeltaB { get; set; }

        // Valores absolutos
        public double AbsDeltaL { get; set; }
        public double AbsDeltaA { get; set; }
        public double AbsDeltaB { get; set; }

        // ΔChroma: diferencia de croma (Chroma_Lot - Chroma_Std), con Chroma = √(a² + b²)
        public double DeltaChroma { get; set; }

        // Diferencia de tono angular (±180°) — calculada; no necesariamente mostrada en UI
        public double DeltaHue { get; set; }

        // ΔE*ab (CIE76)
        public double DeltaE { get; set; }

        // % a corregir por canal (L/A/B) = (|Δ| / |Std|) * 100 * sign(Std)
        public double PercentL { get; set; }
        public double PercentA { get; set; }
        public double PercentB { get; set; }

        // Indicadores de calidad opcionales
        public string LightnessFlag { get; set; } = "";
        public string ChromaFlag { get; set; } = "";
    }

    // ================================
    // CALCULADORA COLORIMÉTRICA
    // ================================
    public static class ColorimetricCalculator
    {

        public static List<ColorCorrectionResult> Calculate(List<ColorimetricRow> rows)
        {
            if (rows == null || rows.Count == 0)
                return new List<ColorCorrectionResult>();

            var results = new List<ColorCorrectionResult>();

            var illuminants = rows
                .Select(r => r.Illuminant)
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x, StringComparer.OrdinalIgnoreCase);

            foreach (var illuminant in illuminants)
            {
                var std = rows.FirstOrDefault(r =>
                    string.Equals(r.Illuminant, illuminant, StringComparison.OrdinalIgnoreCase) &&
                    r.Type.Equals("Std", StringComparison.OrdinalIgnoreCase));

                var lot = rows.FirstOrDefault(r =>
                    string.Equals(r.Illuminant, illuminant, StringComparison.OrdinalIgnoreCase) &&
                    r.Type.Equals("Lot", StringComparison.OrdinalIgnoreCase));

                if (std == null || lot == null)
                    continue;

                // Δ cartesiana
                double dL = lot.L - std.L;
                double dA = lot.A - std.A;
                double dB = lot.B - std.B;

                // ΔE (CIE76)
                double dE = Math.Sqrt(dL * dL + dA * dA + dB * dB);

                // ΔC (croma)
                double chromaStd = Math.Sqrt(std.A * std.A + std.B * std.B);
                double chromaLot = Math.Sqrt(lot.A * lot.A + lot.B * lot.B);
                double dChroma = chromaLot - chromaStd;

                // Δh angular ±180°
                double dHue = lot.Hue - std.Hue;
                if (Math.Abs(dHue) > 180.0)
                    dHue = dHue > 0 ? dHue - 360.0 : dHue + 360.0;

                // % a corregir por canal (con signo del estándar)
                double pctL = (std.L != 0) ? (Math.Abs(dL) / Math.Abs(std.L)) * 100.0 * Math.Sign(std.L) : double.NaN;
                double pctA = (std.A != 0) ? (Math.Abs(dA) / Math.Abs(std.A)) * 100.0 * Math.Sign(std.A) : double.NaN;
                double pctB = (std.B != 0) ? (Math.Abs(dB) / Math.Abs(std.B)) * 100.0 * Math.Sign(std.B) : double.NaN;

                results.Add(new ColorCorrectionResult
                {
                    Illuminant = illuminant,
                    DeltaL = Math.Round(dL, 2),
                    DeltaA = Math.Round(dA, 2),
                    DeltaB = Math.Round(dB, 2),
                    AbsDeltaL = Math.Round(Math.Abs(dL), 2),
                    AbsDeltaA = Math.Round(Math.Abs(dA), 2),
                    AbsDeltaB = Math.Round(Math.Abs(dB), 2),
                    DeltaChroma = Math.Round(dChroma, 2),
                    DeltaHue = Math.Round(dHue, 2),
                    DeltaE = Math.Round(dE, 2),
                    PercentL = Math.Round(pctL, 2),
                    PercentA = Math.Round(pctA, 2),
                    PercentB = Math.Round(pctB, 2),
                    LightnessFlag = "",
                    ChromaFlag = ""
                });
            }

            return results;
        }
    }
}