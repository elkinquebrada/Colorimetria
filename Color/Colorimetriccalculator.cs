using System;
using System.Collections.Generic;
using System.Linq;

namespace Color
{
    // ================================
    // RESULTADO DE CORRECCIÓN DE COLOR
    // ================================
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

        // Diferencia de tono angular (±180°)
        public double DeltaHue { get; set; }

        // ΔE*ab (CIE76)
        public double DeltaE { get; set; }

        // % a corregir por canal = Δ / Std  (con el signo propio de la diferencia)
        public double PercentL { get; set; }
        public double PercentA { get; set; }
        public double PercentB { get; set; }

        // Indicadores de calidad
        public string LightnessFlag { get; set; } = "";

        // dA > 0 → "Rojo (aumentar) o Verde (disminuir)";  dA < 0 → "Rojo (disminuir) o Verde (aumentar)"
        public string ChromaAFlag { get; set; } = "";

        // dB > 0 → "Amarillo (aumentar) o Azul (disminuir)"; dB < 0 → "Azul (disminuir) o Amarillo (aumentar)"
        public string ChromaBFlag { get; set; } = "";

        // Alias de compatibilidad (mantiene el campo original apuntando al flag de L)
        public string ChromaFlag
        {
            get => LightnessFlag;
            set => LightnessFlag = value;
        }
    }

    // ================================
    // RESULTADO CMC(2:1)
    // ================================
    public sealed class CmcResult
    {
        public string Illuminant { get; set; } = "";

        // Componentes CMC(2:1) calculadas desde ΔL, ΔChroma, ΔHue
        public double Lightness    { get; set; }   
        public double Chroma       { get; set; }   
        public double Hue          { get; set; }   
        public double CmcValue     { get; set; }   

        // Conversiones: valor absoluto × 10  (para uso en receta)
        public double ConversionLightness { get; set; }
        public double ConversionChroma    { get; set; }
    }

    // ================================
    // RESULTADO DE RECETA POR ILUMINANTE
    // ================================
    public sealed class RecipeDyeResult
    {
        public string DyeName         { get; set; } = "";
        public double OriginalAmount  { get; set; }   
        public double Calc1Normalized { get; set; }   
        public double Calc2Amount     { get; set; }   
        public double Calc2Normalized { get; set; }   
        public double Calc3Amount     { get; set; }   
        public double Calc3Normalized { get; set; }   
    }

    public sealed class RecipeResult
    {
        public string Illuminant { get; set; } = "";
        public List<RecipeDyeResult> Dyes { get; set; } = new List<RecipeDyeResult>();

        public double TotalOriginal    { get; set; }
        public double TotalCalc2       { get; set; }
        public double TotalCalc3       { get; set; }

        public double VariationLightness { get; set; }  
        public double VariationChroma    { get; set; }  
    }

    // ================================
    // RESULTADO DE TOLERANCIA (límites de una banda)
    // ================================
    public sealed class ToleranceResult
    {
        public double DE { get; set; }
        public double DL { get; set; }  
        public double DC { get; set; }   
        public double DH { get; set; }   
    }

    // ================================
    // EVALUACIÓN DE TOLERANCIA POR ILUMINANTE
    // ================================
    /// Resultado de evaluar un iluminante contra los límites de una banda de tolerancia.
    public sealed class IlluminantToleranceCheck
    {
        /// Nombre del iluminante (p. ej. "D65", "TL84", "CWF")
        public string Illuminant { get; set; } = "";

        ///true si TODOS los componentes están dentro del límite.
        public bool Passes { get; set; }

        // Valores medidos
        public double MeasuredDE { get; set; }
        public double MeasuredDL { get; set; }   
        public double MeasuredDC { get; set; }   
        public double MeasuredDH { get; set; }   

        // Límites aplicados
        public double LimitDE { get; set; }
        public double LimitDL { get; set; }
        public double LimitDC { get; set; }
        public double LimitDH { get; set; }

        // Componente(s) que superan el límite (vacío si cumple)
        public List<string> FailingComponents { get; set; } = new List<string>();

        /// Se muestran DE y el componente dominante (el de mayor valor absoluto entre DL/DC/DH).
        public string Summary
        {
            get
            {
                string status = Passes ? "CUMPLE" : "NO CUMPLE";
                string highlight = string.Equals(Illuminant, "D65", StringComparison.OrdinalIgnoreCase) ? " (PRINCIPAL)" : "";
                return $"{Illuminant,-6} -> {status,-10} (DE={MeasuredDE:F2} {DominantComponentLabel()}){highlight}";
            }
        }

        private string DominantComponentLabel()
        {
            double absL = Math.Abs(MeasuredDL);
            double absC = Math.Abs(MeasuredDC);
            double absH = Math.Abs(MeasuredDH);

            if (absL >= absC && absL >= absH)
                return $"DL={MeasuredDL:+0.00;-0.00;0.00}";
            if (absC >= absH)
                return $"DC={MeasuredDC:+0.00;-0.00;0.00}";
            return $"DH={MeasuredDH:+0.00;-0.00;0.00}";
        }
    }

    /// Resumen de evaluación de todos los iluminantes contra una banda de tolerancia.
    public sealed class ToleranceEvaluationResult
    {
        ///Banda de tolerancia usada (DE límite, p. ej. 1.20).
        public ToleranceResult Band { get; set; } = new ToleranceResult();

        ///Evaluación de cada iluminante.
        public List<IlluminantToleranceCheck> Checks { get; set; } = new List<IlluminantToleranceCheck>();

        ///true si todos los iluminantes cumplen.
        public bool AllPass => Checks.All(c => c.Passes);

        /// Bloque de texto listo para mostrar, 
        public string FormatReport()
        {
            var sb = new System.Text.StringBuilder();
            sb.AppendLine("ESTADO L/ΔE (tolerancias):");
            sb.AppendLine($" DL≤{Band.DL:F2} DC≤{Band.DC:F2} DH≤{Band.DH:F2} DE≤{Band.DE:F2}");
            foreach (var c in Checks)
                sb.AppendLine(c.Summary);
            return sb.ToString().TrimEnd();
        }
    }

    // ================================
    // CALCULADORA COLORIMÉTRICA
    // ================================
    public static class ColorimetricCalculator
    {
        // ------------------------------------------------------------------
        // 1. CORRECCIÓN DE COLOR  (ΔL, Δa, Δb, ΔE, ΔC, ΔH, %, flags)
        // ------------------------------------------------------------------
        public static List<ColorCorrectionResult> Calculate(List<ColorimetricRow> rows)
        {
            if (rows == null || rows.Count == 0)
                return new List<ColorCorrectionResult>();

            var results = new List<ColorCorrectionResult>();

            var standardOrder = new List<string> { "D65", "TL84", "A" };
            var illuminants = rows
                .Select(r => r.Illuminant)
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => 
                {
                    int index = standardOrder.FindIndex(s => string.Equals(s, x, StringComparison.OrdinalIgnoreCase));
                    return index == -1 ? int.MaxValue : index;
                })
                .ThenBy(x => x, StringComparer.OrdinalIgnoreCase);

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
                double dChroma   = chromaLot - chromaStd;

                // Δh angular ±180°
                double dHue = lot.Hue - std.Hue;
                if (Math.Abs(dHue) > 180.0)
                    dHue = dHue > 0 ? dHue - 360.0 : dHue + 360.0;

                // % a corregir = Δ / Std  (fórmula de la hoja: =+E3/B3, es decir |Δ|/Std)
                // El Excel usa =ABS(Δ)/Std, lo que con Std negativo da un % negativo.
                double pctL = (std.L != 0) ? (Math.Abs(dL) / std.L) * 100.0 : double.NaN;
                double pctA = (std.A != 0) ? (Math.Abs(dA) / std.A) * 100.0 : double.NaN;
                double pctB = (std.B != 0) ? (Math.Abs(dB) / std.B) * 100.0 : double.NaN;

                // Flags de calidad
                string lightnessFlag = dL < 0 ? "Oscuro" : dL > 0 ? "Claro" : "";
                string chromaAFlag   = dA > 0
                    ? "Rojo (aumentar) o Verde (disminuir)"
                    : dA < 0
                        ? "Rojo (disminuir) o Verde (aumentar)"
                        : "";
                string chromaBFlag   = dB < 0
                    ? "Azul (disminuir) o Amarillo (aumentar)"
                    : dB > 0
                        ? "Amarillo (aumentar) o Azul (disminuir)"
                        : "";

                results.Add(new ColorCorrectionResult
                {
                    Illuminant  = illuminant,
                    DeltaL      = Math.Round(dL,  4),
                    DeltaA      = Math.Round(dA,  4),
                    DeltaB      = Math.Round(dB,  4),
                    AbsDeltaL   = Math.Round(Math.Abs(dL), 4),
                    AbsDeltaA   = Math.Round(Math.Abs(dA), 4),
                    AbsDeltaB   = Math.Round(Math.Abs(dB), 4),
                    DeltaChroma = Math.Round(dChroma, 4),
                    DeltaHue    = Math.Round(dHue,    4),
                    DeltaE      = Math.Round(dE,      4),
                    PercentL    = double.IsNaN(pctL) ? double.NaN : Math.Round(pctL, 6),
                    PercentA    = double.IsNaN(pctA) ? double.NaN : Math.Round(pctA, 6),
                    PercentB    = double.IsNaN(pctB) ? double.NaN : Math.Round(pctB, 6),
                    LightnessFlag = lightnessFlag,
                    ChromaAFlag   = chromaAFlag,
                    ChromaBFlag   = chromaBFlag
                });
            }

            return results;
        }

        // ------------------------------------------------------------------
        // 2. CMC(2:1)
        //    Fórmulas del Excel (hoja CALCULO RECETA):
        // ------------------------------------------------------------------
        public static List<CmcResult> CalculateCmc(List<ColorCorrectionResult> corrections)
        {
            if (corrections == null || corrections.Count == 0)
                return new List<CmcResult>();

            var results = new List<CmcResult>();

            foreach (var c in corrections)
            {
                // CMC(2:1): √( (ΔL/lightnessTerm)² + (ΔC/chromaTerm)² + (ΔH/hueTerm)² )
                double convL = Math.Round(Math.Abs(c.DeltaL * 10.0), 4);
                double convC = Math.Round(Math.Abs(c.DeltaChroma * 10.0), 4);

                results.Add(new CmcResult
                {
                    Illuminant          = c.Illuminant,
                    Lightness           = c.DeltaL,
                    Chroma              = c.DeltaChroma,
                    Hue                 = c.DeltaHue,
                    CmcValue            = c.DeltaE,   
                    ConversionLightness = convL,
                    ConversionChroma    = convC
                });
            }

            return results;
        }

        // Sobrecarga: recibe valores CMC directamente del espectrofotómetro
        public static CmcResult BuildCmcResult(
            string illuminant,
            double cmcLightness,
            double cmcChroma,
            double cmcHue,
            double cmcValue)
        {
            return new CmcResult
            {
                Illuminant          = illuminant,
                Lightness           = cmcLightness,
                Chroma              = cmcChroma,
                Hue                 = cmcHue,
                CmcValue            = cmcValue,
                ConversionLightness = Math.Round(Math.Abs(cmcLightness * 10.0), 4),
                ConversionChroma    = Math.Round(Math.Abs(cmcChroma    * 10.0), 4)
            };
        }

        // ------------------------------------------------------------------
        // 3. CÁLCULO DE RECETA
        // ------------------------------------------------------------------
        public static RecipeResult CalculateRecipe(
            string illuminant,
            List<(string name, double amount)> dyes,
            CmcResult cmc)
        {
            if (dyes == null || dyes.Count == 0 || cmc == null)
                return new RecipeResult { Illuminant = illuminant };

            double convL = cmc.ConversionLightness;
            double convC = cmc.ConversionChroma;

            double totalOriginal = dyes.Sum(d => d.amount);

            // Cálculo 2
            var calc2 = dyes.Select(d => d.amount * (1.0 - convL / 100.0)).ToList();
            double totalCalc2 = calc2.Sum();

            // Cálculo 3
            var calc3 = calc2.Select(v => v * (1.0 - convC / 100.0)).ToList();
            double totalCalc3 = calc3.Sum();

            double totalCalc2Normalized = totalOriginal > 0 ? totalCalc2 / totalOriginal : 0.0;

            var dyeResults = new List<RecipeDyeResult>();
            for (int i = 0; i < dyes.Count; i++)
            {
                dyeResults.Add(new RecipeDyeResult
                {
                    DyeName         = dyes[i].name,
                    OriginalAmount  = dyes[i].amount,
                    Calc1Normalized = totalOriginal > 0
                        ? Math.Round(dyes[i].amount / totalOriginal, 6) : double.NaN,
                    Calc2Amount     = Math.Round(calc2[i], 6),
                    Calc2Normalized = totalCalc2 > 0
                        ? Math.Round(calc2[i] / totalCalc2, 6) : double.NaN,
                    Calc3Amount     = Math.Round(calc3[i], 6),
                    Calc3Normalized = totalCalc3 > 0
                        ? Math.Round(calc3[i] / totalCalc3, 6) : double.NaN
                });
            }

            double varLightness = totalOriginal > 0
                ? Math.Round(totalCalc2 / totalOriginal - 1.0, 6) : double.NaN;
            double varChroma = totalCalc2Normalized > 0
                ? Math.Round(totalCalc3 / totalCalc2Normalized - 1.0, 6) : double.NaN;

            return new RecipeResult
            {
                Illuminant        = illuminant,
                Dyes              = dyeResults,
                TotalOriginal     = Math.Round(totalOriginal, 6),
                TotalCalc2        = Math.Round(totalCalc2,    6),
                TotalCalc3        = Math.Round(totalCalc3,    6),
                VariationLightness = varLightness,
                VariationChroma    = varChroma
            };
        }

        // ------------------------------------------------------------------
        // 4. TOLERANCIA
        //    Fórmula del Excel (hoja TOLERANCIA):
        // ------------------------------------------------------------------
        public static ToleranceResult CalculateTolerance(double de)
        {
            double component = Math.Sqrt((de * de) / 3.0);
            return new ToleranceResult
            {
                DE = Math.Round(de,        5),
                DL = Math.Round(component, 5),
                DC = Math.Round(component, 5),
                DH = Math.Round(component, 5)
            };
        }

        /// Calcula tres bandas de tolerancia a partir de los tres valores DE habituales
        public static List<ToleranceResult> CalculateToleranceBands(IEnumerable<double> deValues)
        {
            return deValues?.Select(CalculateTolerance).ToList()
                   ?? new List<ToleranceResult>();
        }

        // ------------------------------------------------------------------
        //  EVALUACIÓN DE TOLERANCIA POR ILUMINANTE
        // ------------------------------------------------------------------
        public static ToleranceEvaluationResult EvaluateTolerance(
            List<ColorCorrectionResult> corrections,
            ToleranceResult band)
        {
            var evaluation = new ToleranceEvaluationResult { Band = band };

            if (corrections == null || corrections.Count == 0)
                return evaluation;

            foreach (var c in corrections)
            {
                double absDE = Math.Abs(c.DeltaE);
                double absDL = Math.Abs(c.DeltaL);
                double absDC = Math.Abs(c.DeltaChroma);
                double absDH = Math.Abs(c.DeltaHue);

                var failing = new List<string>();
                if (absDE > band.DE) failing.Add("DE");
                if (absDL > band.DL) failing.Add("DL");
                if (absDC > band.DC) failing.Add("DC");
                if (absDH > band.DH) failing.Add("DH");

                evaluation.Checks.Add(new IlluminantToleranceCheck
                {
                    Illuminant        = c.Illuminant,
                    Passes            = failing.Count == 0,
                    MeasuredDE        = Math.Round(absDE,         2),
                    MeasuredDL        = Math.Round(c.DeltaL,      2),
                    MeasuredDC        = Math.Round(c.DeltaChroma, 2),
                    MeasuredDH        = Math.Round(c.DeltaHue,    2),
                    LimitDE           = band.DE,
                    LimitDL           = band.DL,
                    LimitDC           = band.DC,
                    LimitDH           = band.DH,
                    FailingComponents = failing
                });
            }

            return evaluation;
        }

        public static ToleranceEvaluationResult EvaluateTolerance(
            List<ColorCorrectionResult> corrections,
            double deLimitBand)
        {
            var band = CalculateTolerance(deLimitBand);
            return EvaluateTolerance(corrections, band);
        }

        /// Evalúa contra múltiples bandas de tolerancia de una sola vez.
        /// Útil para mostrar el estado en las tres bandas (0.6 / 1.2 / 1.8) simultáneamente.
        public static List<ToleranceEvaluationResult> EvaluateToleranceBands(
            List<ColorCorrectionResult> corrections,
            IEnumerable<double> deLimitBands)
        {
            return deLimitBands?
                .Select(de => EvaluateTolerance(corrections, de))
                .ToList()
                ?? new List<ToleranceEvaluationResult>();
        }
    }
}