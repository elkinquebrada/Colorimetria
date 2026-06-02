using System;
using System.Collections.Generic;
using System.Linq;

namespace Color
{
    // ========================================================================
    // RESULTADO DE CORRECCION DE COLOR (Fase 2 - Motor Experto)
    // ========================================================================
    public sealed class ColorCorrectionResult
    {
        public string Illuminant { get; set; } = "";
        public string ShadeName { get; set; } = "";

        // Escenarios de Correccion (Fase 2 - Paridad Excel Coats)
        public decimal FactorL { get; set; }
        public decimal FactorA { get; set; }
        public decimal FactorB { get; set; }
        public decimal FactorC { get; set; }
        public decimal FactorH { get; set; }

        // Valores Base (Auditoria)
        public double StdL { get; set; }
        public double StdA { get; set; }
        public double StdB { get; set; }
        public double StdC { get; set; }
        public double StdH { get; set; }
        public double LotL { get; set; }
        public double LotA { get; set; }
        public double LotB { get; set; }
        public double LotC { get; set; }
        public double LotH { get; set; }

        // Diagnostico Experto
        public string GlobalStatus { get; set; }
        public bool FlagAlertMetamerism { get; set; }
        public string MetamerismAlert { get; set; } = "";
        public bool Success { get; set; }
        public string Message { get; set; }

        // Nombres de Colorantes (Inyectados desde RecipeCorrector)
        public string PrimaryDyeName { get; set; } = "Colorante Principal";
        public string SecondaryDyeName { get; set; } = "Colorante de Brillo";
        public string TonerDyeName { get; set; } = "Matizador";

        // Deltas (Double para UI y Graficos)
        public double DeltaL { get; set; }
        public double DeltaA { get; set; }
        public double DeltaB { get; set; }
        public double DeltaChroma { get; set; }
        public double DeltaHue { get; set; }
        public double DeltaE { get; set; }
        public double CmcValue { get; set; }

        // Porcentajes de Variacion (Standard Excel Coats: (Std-Lot)/Std * 100)
        public double PercentL => (double)(FactorL * 100);
        public double PercentA => (double)(FactorA * 100);
        public double PercentB => (double)(FactorB * 100);
        public double PercentChroma => (double)(FactorC * 100);
        public double PercentHue => Math.Abs(DeltaHue);
        public double PorcentajeRecetaL => PercentL; // Alias para compatibilidad UI

        // Valores de Ejes CMC Especi≠ficos
        public double CmcLightness { get; set; }
        public double CmcChroma { get; set; }
        public double CmcHue { get; set; }

        // Valores Absolutos para Graficos
        public double AbsDeltaL => Math.Abs(DeltaL);
        public double AbsDeltaA => Math.Abs(DeltaA);
        public double AbsDeltaB => Math.Abs(DeltaB);

        // Textos del OCR (Espejo)
        public string OcrValueL { get; set; }
        public string OcrValueC { get; set; }
        public string OcrValueH { get; set; }

        public string OcrImpactoL { get; set; }
        public string OcrImpactoC { get; set; }
        public string OcrImpactoH { get; set; }

        public int FactorIntL { get; set; }
        public int FactorIntC { get; set; }
        public int FactorIntH { get; set; }

        // --- Propiedades de Diagnostico Dinamico ---
        public string DiagnosticoL => !string.IsNullOrEmpty(OcrImpactoL) ? OcrImpactoL : GetInternalDiagnosis("L");
        public string DiagnosticoLoteL => ColorimetricCalculator.GetEngineeringDiagnosis("dl", DeltaL, ImpactoLoteL);
        public string ImpactoRecetaL => ColorimetricCalculator.GetImpactoLRecipe(DeltaL);
        public string ImpactoLoteL => ColorimetricCalculator.GetImpactoLLot(DeltaL);
        public string RecomendacionRecetaL => ColorimetricCalculator.GetInstLRecipe(DeltaL, Math.Abs(PercentL), PrimaryDyeName);
        public string RecomendacionLoteL => ColorimetricCalculator.GetInstLLot(DeltaL, Math.Abs(PercentL), PrimaryDyeName);

        public string DiagnosisC => !string.IsNullOrEmpty(OcrImpactoC) ? OcrImpactoC : GetInternalDiagnosis("C");

        // DeltaChroma = Std - Lot:
        public string DescripcionC => (DeltaChroma > 0 ? "Opaco" : "Brillante");
        public string RecommendationC => ColorimetricCalculator.GetRecommendationC_Expert(DeltaL, DeltaChroma, Math.Abs(PercentChroma), SecondaryDyeName, PrimaryDyeName);

        public string DiagnosisH => !string.IsNullOrEmpty(OcrImpactoH) ? OcrImpactoH : GetInternalDiagnosis("H");

        // Direccion del viraje: eje dominante (|da| vs |db|)
        public string ImpactoMatiz => ColorimetricCalculator.GetHueDirection(DeltaA, DeltaB);
        public string RecomendacionMatiz => $"{(DeltaHue > 0 ? "Aumentar" : "Disminuir")} {TonerDyeName} {Math.Abs(DeltaHue):F2}%";

        private string GetInternalDiagnosis(string eje)
        {
            if (eje == "L") return ColorimetricCalculator.GetEngineeringDiagnosis("dl", DeltaL, ImpactoRecetaL);
            if (eje == "C") return ColorimetricCalculator.GetEngineeringDiagnosis("da", DeltaChroma, DescripcionC);
            if (eje == "H") return ColorimetricCalculator.GetEngineeringDiagnosis("db", DeltaHue, ImpactoMatiz);
            return "";
        }

        public bool Pass { get; set; }
    }

    public sealed class CmcResult
    {
        public string Illuminant { get; set; } = "";
        public double Lightness { get; set; }
        public double Chroma { get; set; }
        public double Hue { get; set; }
        public double CmcValue { get; set; }
    }

    public sealed class RecipeResult
    {
        public string Illuminant { get; set; } = "";
        public List<RecipeDyeResult> Dyes { get; set; } = new List<RecipeDyeResult>();
        public double TotalOriginal { get; set; }
        public double TotalCalc2 { get; set; }
        public double TotalCalc3 { get; set; }
    }

    public sealed class RecipeDyeResult
    {
        public string DyeName { get; set; } = "";
        public double OriginalAmount { get; set; }
        public double Calc1Normalized { get; set; }
        public double Calc2Amount { get; set; }
        public double Calc3Amount { get; set; }
    }

    public sealed class ToleranceResult
    {
        public double DE { get; set; }
        public double DL { get; set; }
        public double DC { get; set; }
        public double DH { get; set; }
    }

    // ========================================================================
    // CALCULADORA COLORIMETRICA (INDUSTRIAL STANDARD ENGINE)
    // ========================================================================
    public static class ColorimetricCalculator
    {
        /// Motor de Decision.
        /// F√≥rmula: (Std - Lot) / Std
        public static ColorCorrectionResult CalculateIndustrialCorrection(OcrReport report)
        {
            var all = CalculateAllIlluminants(report);
            return all.FirstOrDefault(r => r.Illuminant == "D65") ?? all.FirstOrDefault() ?? new ColorCorrectionResult { Success = false };
        }

        public static List<ColorCorrectionResult> CalculateAllIlluminants(OcrReport report)
        {
            var results = new List<ColorCorrectionResult>();
            if (report == null || report.Measures.Count == 0) return results;

            // 1. D65 (Principal)
            var resD65 = CalculateForIlluminant(report, "D65");
            if (resD65 != null) results.Add(resD65);

            // 2. TL84 (Secundario)
            var resTL84 = CalculateForIlluminant(report, "TL84");
            if (resTL84 != null) results.Add(resTL84);

            // 3. A / CWF / Otros
            var resA = CalculateForIlluminant(report, "A") ?? CalculateForIlluminant(report, "CWF");
            if (resA != null) results.Add(resA);

            // Post-procesamiento: Metamerismo (D65 vs TL84)
            if (resD65 != null && resTL84 != null)
            {
                if (Math.Sign(resD65.DeltaL) != Math.Sign(resTL84.DeltaL) && Math.Abs(resD65.DeltaL) > 0.1)
                {
                    resD65.FlagAlertMetamerism = true;
                    resD65.MetamerismAlert = "Inconsistencia Metam√©rica (D65 vs TL84)";
                }
            }

            return results;
        }

        private static ColorCorrectionResult CalculateForIlluminant(OcrReport report, string illuminantName)
        {
            var std = report.Measures.FirstOrDefault(m => m.Illuminant.ToUpper().Contains(illuminantName.ToUpper()) && m.Type.ToUpper().Contains("STD"));
            var lot = report.Measures.FirstOrDefault(m => m.Illuminant.ToUpper().Contains(illuminantName.ToUpper()) && (m.Type.ToUpper().Contains("LOT") || m.Type.ToUpper().Contains("SPL")));

            if (std == null || lot == null) return null;

            var res = new ColorCorrectionResult { Success = true, Illuminant = illuminantName };

            decimal sL = (decimal)std.L;
            decimal sA = (decimal)std.A;
            decimal sB = (decimal)std.B;
            decimal sC = (decimal)std.Chroma;
            decimal sH = (decimal)std.Hue;

            decimal lL = (decimal)lot.L;
            decimal lA = (decimal)lot.A;
            decimal lB = (decimal)lot.B;
            decimal lC = (decimal)lot.Chroma;
            decimal lH = (decimal)lot.Hue;

            // MOTOR DE SENSIBILIDAD ADAPTATIVA 
            // Switch de Chroma: evalua para determinar el carril de ajuste.
            // CARRIL A (sC > 15): Colores Vivos  Proporcion Lineal  (Std-Lot)/Std
            // CARRIL B Colores Oscuros Factor Fijo √ó0.15 en ejes cromaticos
            bool esColorOscuro = (sC <= 15m);

            if (esColorOscuro)
            {
                // CARRIL B Matriz de ejes para colores oscuros/negros
                // dL: Proporcional    (sL - lL) / sL   [mantiene fuerza para negros]
                res.FactorL = sL != 0 ? Math.Round((sL - lL) / sL, 8) : 0;
                // da: Ponderado     (sA - lA) * 0.15  [estabiliza matiz]
                res.FactorA = Math.Round((sA - lA) * 0.15m, 8);
                // db: Ponderado       (sB - lB) * 0.15  [estabiliza matiz]
                res.FactorB = Math.Round((sB - lB) * 0.15m, 8);
                // dC: Ponderado       (sC - lC) * 0.15  [evita ajustes absurdos de saturacion]
                res.FactorC = Math.Round((sC - lC) * 0.15m, 8);
            }
            else
            {
                // CARRIL A  Proporcion Lineal estandar para todos los ejes
                res.FactorL = sL != 0 ? Math.Round((sL - lL) / sL, 8) : 0;
                res.FactorA = sA != 0 ? Math.Round((sA - lA) / sA, 8) : 0;
                res.FactorB = sB != 0 ? Math.Round((sB - lB) / sB, 8) : 0;
                res.FactorC = sC != 0 ? Math.Round((sC - lC) / sC, 8) : 0;
            }

            // dH: siempre ponderado ó0.15 en ambos carriles (evita virajes bruscos)
            decimal dH_Raw = lH - sH;
            if (dH_Raw > 180) dH_Raw -= 360;
            if (dH_Raw < -180) dH_Raw += 360;
            res.FactorH = Math.Round(dH_Raw * 0.15m, 8);

            // Deltas CIE (Lot - Std segun paridad con Excel)
            res.DeltaL = (double)(lL - sL);
            res.DeltaA = (double)(lA - sA);
            res.DeltaB = (double)(lB - sB);
            res.DeltaChroma = (double)(lC - sC);

            // Delta Hue (CIE) dH = sign(h_Lot - h_Std) * sqrt(da^2 + db^2 - dC^2)
            double hDiff = (double)(lH - sH);
            if (hDiff > 180) hDiff -= 360;
            if (hDiff < -180) hDiff += 360;
            double dH_sq = res.DeltaA * res.DeltaA + res.DeltaB * res.DeltaB - res.DeltaChroma * res.DeltaChroma;
            res.DeltaHue = dH_sq >= 0 ? Math.Sign(hDiff) * Math.Sqrt(dH_sq) : 0;

            // C√ÅLCULO DE MOTOR CMC (2:1) SEGUN HOJA EXCEL
            var (sl, sc, sh) = CalculateCmcSemiAxes((double)sL, (double)sC, (double)sH);
            res.CmcLightness = sl > 0 ? res.DeltaL / (2.0 * sl) : 0;
            res.CmcChroma = sc > 0 ? res.DeltaChroma / sc : 0;
            res.CmcHue = sh > 0 ? res.DeltaHue / sh : 0;
            res.DeltaE = Math.Sqrt(res.CmcLightness * res.CmcLightness + res.CmcChroma * res.CmcChroma + res.CmcHue * res.CmcHue);
            res.CmcValue = res.DeltaE;

            // Intentar recuperar los valores OCR si estaban presentes
            var cmc = report.CmcDifferences?.FirstOrDefault(c => c.Illuminant.ToUpper().Contains(illuminantName.ToUpper()));
            if (cmc != null)
            {
                // Solo si el OCR dictamina≥ un DE lo sobrescribimos si se considerara necesario, pero el Excel Engine manda.
                // res.DeltaE = cmc.DeltaCMC; 
                // res.CmcValue = cmc.DeltaCMC;
                res.OcrValueL = $"{cmc.DeltaLightness.ToString(System.Globalization.CultureInfo.InvariantCulture)} {cmc.LightnessFlagOcr ?? ""}".Trim();
                res.OcrValueC = $"{cmc.DeltaChroma.ToString(System.Globalization.CultureInfo.InvariantCulture)} {cmc.ChromaFlagOcr ?? ""}".Trim();
                res.OcrValueH = $"{cmc.DeltaHue.ToString(System.Globalization.CultureInfo.InvariantCulture)} {cmc.HueFlagOcr ?? ""}".Trim();

                res.OcrImpactoL = cmc.LightnessFlagOcr ?? "";
                res.OcrImpactoC = cmc.ChromaFlagOcr ?? "";
                res.OcrImpactoH = cmc.HueFlagOcr ?? "";

                // Calculamos temporalmente el factor entero para tenerlo estructurado
                res.FactorIntL = 0;
                res.FactorIntC = 0;
                res.FactorIntH = 0;

                if (double.TryParse(System.Text.RegularExpressions.Regex.Match(res.OcrValueL, @"[-+]?[0-9]*\.?[0-9]+").Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double parsedL))
                    res.FactorIntL = (int)Math.Round(parsedL * 100, MidpointRounding.AwayFromZero);
                if (double.TryParse(System.Text.RegularExpressions.Regex.Match(res.OcrValueC, @"[-+]?[0-9]*\.?[0-9]+").Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double parsedC))
                    res.FactorIntC = (int)Math.Round(parsedC * 100, MidpointRounding.AwayFromZero);
                if (double.TryParse(System.Text.RegularExpressions.Regex.Match(res.OcrValueH, @"[-+]?[0-9]*\.?[0-9]+").Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double parsedH))
                    res.FactorIntH = (int)Math.Round(parsedH * 100, MidpointRounding.AwayFromZero);

                // REGLA "CAPA ESPEJO" (Innegociable) 
                // Si la Variacion Visual que ve el cliente es 0% (por redondeo u OCR),
                // el Factor de ajuste de la receta para ese eje se fuerza a 0.
                // Esto garantiza coherencia total: 0% visual  0% de ajuste en receta.
                if (res.FactorIntL == 0) res.FactorL = 0m;
                if (res.FactorIntC == 0) res.FactorC = 0m;
                if (res.FactorIntH == 0) res.FactorH = 0m;
            }

            res.StdL = std.L; res.StdA = std.A; res.StdB = std.B; res.StdC = std.Chroma; res.StdH = std.Hue;
            res.LotL = lot.L; res.LotA = lot.A; res.LotB = lot.B; res.LotC = lot.Chroma; res.LotH = lot.Hue;

            return res;
        }

        // --- HELPERS ---

        public static (double sl, double sc, double sh) CalculateCmcSemiAxes(double L1, double C1, double h1)
        {
            double f = Math.Sqrt(Math.Pow(C1, 4) / (Math.Pow(C1, 4) + 1900.0));
            double T = (h1 >= 164.0 && h1 <= 345.0)
                ? 0.56 + Math.Abs(0.2 * Math.Cos((Math.PI / 180.0) * (h1 + 168.0)))
                : 0.36 + Math.Abs(0.4 * Math.Cos((Math.PI / 180.0) * (h1 + 35.0)));
            double sl = L1 < 16.0 ? 0.511 : (0.040975 * L1) / (1.0 + 0.01765 * L1);
            double sc = (0.0638 * C1) / (1.0 + 0.0131 * C1) + 0.638;
            double sh = sc * (f * T + 1.0 - f);
            return (sl, sc, sh);
        }

        public static string GetDiagL_Expert(double dL) => (Math.Abs(dL) > 0.5 ? "Desviacion Cri≠tica" : "Desviacion Moderada");
        public static string GetImpactoLRecipe(double dL) => (dL < 0 ? "Mas Claro" : "Mas Oscuro");
        public static string GetImpactoLLot(double dL) => (dL < 0 ? "Brillante" : "");
        public static string GetInstLRecipe(double dL, double varL, string name) => $"{(dL < 0 ? "INCREMENTAR" : "REDUCIR")} {Math.Abs(varL):F2}%";
        public static string GetInstLLot(double dL, double varL, string name) => $"{(dL < 0 ? "ADICIONAR" : "REDUCIR")} {Math.Abs(varL):F2}%";
        public static string GetDiagC_Expert(double dC) => (dC < 0 ? "Saturado" : "Opaco");
        public static string GetDiagH_Expert(double dH, double tol) => (dH < 0 ? "Viraje (+)" : "Viraje (-)");


        public static string FormatDelta(double value)
        {
            if (Math.Abs(value) < 0.05) return "0"; 
            return value.ToString("+0.00;-0.00;0", System.Globalization.CultureInfo.InvariantCulture);
        }

        public static string GetHueDirection(double dA, double dB)
        {
            if (Math.Abs(dA) >= Math.Abs(dB))
            {
                return dA > 0 ? "Redder" : "Greener";
            }
            else
            {
                return dB > 0 ? "Yellower" : "Bluer";
            }
        }

        // --- MODULO DE INGENIER√çA TEXTIL (DIAGNOSTICO FINAL) ---
        public static string GetEngineeringDiagnosis(string eje, double delta, string impacto)
        {
            switch (eje.ToUpper())
            {
                case "DL":
                case "L":
                    return delta < 0 ? "Oscuro (Deep)" : "Claro (Thin)";

                case "DC":
                case "C":
                    return delta > 0 ? "Brighter" : "Duller";

                case "DH":
                case "DB":
                case "H":
                    return GetHueDirection(0, delta); // Usamos db por defecto para H en este contexto
            }

            return "OK";
        }

        // Logica de Ciclo Industrial (Matriz Diagonal)
        public static string GetRecommendationC_Expert(double dL, double dC, double varC, string secName, string priName)
        {
            if (Math.Abs(dC) < 0.05) return $"Verificar {priName}";

            if (dL > 0) // Oscuro (Std > Lot)
            {
                return dC < 0
                    ? $"restar {secName} {varC:F2}%"
                    : $"restar {priName} {varC:F2}%";
            }
            else // Claro (Std < Lot)
            {
                return dC < 0
                    ? $"sumar {priName} (opaco) {varC:F2}%"
                    : $"sumar {secName} {varC:F2}%";
            }
        }
    }
}