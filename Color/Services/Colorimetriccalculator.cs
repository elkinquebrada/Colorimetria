using System;
using System.Collections.Generic;
using System.Linq;

namespace Color
{
    // ========================================================================
    // RESULTADO DE CORRECCIÓN DE COLOR (Fase 2 - Motor Experto)
    // ========================================================================
    public sealed class ColorCorrectionResult
    {
        public string Illuminant { get; set; } = "";
        public string ShadeName { get; set; } = "";

        // Escenarios de Corrección (Fase 2 - Paridad Excel Coats)
        public decimal FactorL { get; set; } 
        public decimal FactorA { get; set; } 
        public decimal FactorB { get; set; } 
        public decimal FactorC { get; set; } 
        public decimal FactorH { get; set; } 

        // Valores Base (Auditoría)
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

        // Diagnóstico Experto
        public string GlobalStatus { get; set; }
        public bool FlagAlertMetamerism { get; set; }
        public string MetamerismAlert { get; set; } = "";
        public bool Success { get; set; }
        public string Message { get; set; }

        // Nombres de Colorantes (Inyectados desde RecipeCorrector)
        public string PrimaryDyeName { get; set; } = "Colorante Principal";
        public string SecondaryDyeName { get; set; } = "Colorante de Brillo";
        public string TonerDyeName { get; set; } = "Matizador";

        // Deltas (Double para UI y Gráficos)
        public double DeltaL { get; set; }
        public double DeltaA { get; set; }
        public double DeltaB { get; set; }
        public double DeltaChroma { get; set; }
        public double DeltaHue { get; set; }
        public double DeltaE { get; set; }
        public double CmcValue { get; set; }

        // Porcentajes de Variación (Standard Excel Coats: (Std-Lot)/Std * 100)
        public double PercentL => (double)(FactorL * 100);
        public double PercentA => (double)(FactorA * 100);
        public double PercentB => (double)(FactorB * 100);
        public double PercentChroma => (double)(FactorC * 100);
        public double PercentHue => Math.Abs(DeltaHue);
        public double PorcentajeRecetaL => PercentL; // Alias para compatibilidad UI

        // Valores Absolutos para Gráficos
        public double AbsDeltaL => Math.Abs(DeltaL);
        public double AbsDeltaA => Math.Abs(DeltaA);
        public double AbsDeltaB => Math.Abs(DeltaB);

        // --- Propiedades de Diagnóstico Dinámico ---
        public string DiagnosticoL => Math.Abs(DeltaL) < 0.2 ? "✔" : ColorimetricCalculator.GetEngineeringDiagnosis("dl", DeltaL, ImpactoRecetaL);
        public string DiagnosticoLoteL => Math.Abs(DeltaL) < 0.2 ? "✔" : ColorimetricCalculator.GetEngineeringDiagnosis("dl", DeltaL, ImpactoLoteL);
        public string ImpactoRecetaL => ColorimetricCalculator.GetImpactoLRecipe(DeltaL);
        public string ImpactoLoteL => ColorimetricCalculator.GetImpactoLLot(DeltaL);
        public string RecomendacionRecetaL => ColorimetricCalculator.GetInstLRecipe(DeltaL, Math.Abs(PercentL), PrimaryDyeName);
        public string RecomendacionLoteL => ColorimetricCalculator.GetInstLLot(DeltaL, Math.Abs(PercentL), PrimaryDyeName);

        public string DiagnosisC => Math.Abs(DeltaChroma) < 0.15 ? "✔" : ColorimetricCalculator.GetEngineeringDiagnosis("da", DeltaChroma, DescripcionC);

        // DeltaChroma = Std - Lot:
        public string DescripcionC => (Math.Abs(DeltaChroma) < 0.15) ? "✔" : (DeltaChroma > 0 ? "Opaco" : "Brillante");
        public string RecommendationC => (Math.Abs(DeltaChroma) < 0.1) ? "✔" : ColorimetricCalculator.GetRecommendationC_Expert(DeltaL, DeltaChroma, Math.Abs(PercentChroma), SecondaryDyeName, PrimaryDyeName);

        public string DiagnosisH => Math.Abs(DeltaHue) < 0.1 ? "✔" : ColorimetricCalculator.GetEngineeringDiagnosis("db", DeltaHue, ImpactoMatiz);
       
        // Dirección del viraje: eje dominante (|da| vs |db|)
        public string ImpactoMatiz => (Math.Abs(DeltaHue) < 0.1) ? "✔" : ColorimetricCalculator.GetHueDirection(DeltaA, DeltaB);
        public string RecomendacionMatiz => (Math.Abs(DeltaHue) < 0.1) ? "✔" : $"{(DeltaHue > 0 ? "Aumentar" : "Disminuir")} {TonerDyeName} {Math.Abs(DeltaHue):F2}%";

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
    // CALCULADORA COLORIMÉTRICA (INDUSTRIAL STANDARD ENGINE)
    // ========================================================================
    public static class ColorimetricCalculator
    {
        /// Motor de Decisión.
        /// Fórmula: (Std - Lot) / Std
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
                    resD65.MetamerismAlert = "Inconsistencia Metamérica (D65 vs TL84)";
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

            // Variaciones Relativas con límite 15% (Protocolo de Seguridad Industrial)
            decimal Limit15(decimal f) => Math.Max(-0.15m, Math.Min(0.15m, f));

            res.FactorL = Limit15(sL != 0 ? Math.Round((sL - lL) / sL, 8) : 0);
            res.FactorA = Limit15(sA != 0 ? Math.Round((sA - lA) / sA, 8) : 0);
            res.FactorB = Limit15(sB != 0 ? Math.Round((sB - lB) / sB, 8) : 0);
            res.FactorC = Limit15(sC != 0 ? Math.Round((sC - lC) / sC, 8) : 0);
            
            decimal dH_Raw = lH - sH;
            if (dH_Raw > 180) dH_Raw -= 360;
            if (dH_Raw < -180) dH_Raw += 360;
            
            // FactorH (Hue) differs conceptually, limits equivalent limit 15% applied to degrees
            res.FactorH = Limit15(Math.Round(dH_Raw, 8) / 100.0m) * 100.0m;

            // Deltas para UI y Gráficos (PARIDAD EXCEL COATS: Std - Lot)
            res.DeltaL = (double)(sL - lL);
            res.DeltaA = (double)(sA - lA);
            res.DeltaB = (double)(sB - lB);
            res.DeltaChroma = (double)(sC - lC);
            res.DeltaHue = (double)(-res.FactorH); 
            
            var cmc = report.CmcDifferences?.FirstOrDefault(c => c.Illuminant.ToUpper().Contains(illuminantName.ToUpper()));
            res.DeltaE = cmc?.DeltaCMC ?? Math.Sqrt(res.DeltaL*res.DeltaL + res.DeltaA*res.DeltaA + res.DeltaB*res.DeltaB);
            res.CmcValue = res.DeltaE;

            res.StdL = std.L; res.StdA = std.A; res.StdB = std.B; res.StdC = std.Chroma; res.StdH = std.Hue;
            res.LotL = lot.L; res.LotA = lot.A; res.LotB = lot.B; res.LotC = lot.Chroma; res.LotH = lot.Hue;

            return res;
        }

        // --- HELPERS ---

        public static (double sl, double sc, double sh) CalculateCmcSemiAxes(double L1, double C1, double h1)
        {
            double f = Math.Sqrt(Math.Pow(C1, 4) / (Math.Pow(C1, 4) + 1900.0));
            double T = (h1 >= 164.0 && h1 <= 345.0) 
                ? 0.56 + Math.Abs(0.2 * Math.Cos((Math.PI/180.0)*(h1 + 168.0)))
                : 0.36 + Math.Abs(0.4 * Math.Cos((Math.PI/180.0)*(h1 + 35.0)));
            double sl = L1 < 16.0 ? 0.511 : (0.040975 * L1) / (1.0 + 0.01765 * L1);
            double sc = (0.0638 * C1) / (1.0 + 0.0131 * C1) + 0.638;
            double sh = sc * (f * T + 1.0 - f);
            return (sl, sc, sh);
        }

        public static string GetDiagL_Expert(double dL) => Math.Abs(dL) < 0.2 ? "✔" : (Math.Abs(dL) > 0.5 ? "Desviación Crítica" : "Desviación Moderada");
        public static string GetImpactoLRecipe(double dL) => Math.Abs(dL) < 0.2 ? "✔" : (dL < 0 ? "Más Claro" : "Más Oscuro");
        public static string GetImpactoLLot(double dL) => Math.Abs(dL) < 0.2 ? "✔" : (dL < 0 ? "Brillante" : "");
        public static string GetInstLRecipe(double dL, double varL, string name) => Math.Abs(dL) < 0.2 ? "✔" : $"{(dL < 0 ? "INCREMENTAR" : "REDUCIR")} {Math.Abs(varL):F2}%";
        public static string GetInstLLot(double dL, double varL, string name) => Math.Abs(dL) < 0.2 ? "✔" : $"{(dL < 0 ? "ADICIONAR" : "REDUCIR")} {Math.Abs(varL):F2}%";
        public static string GetDiagC_Expert(double dC) => Math.Abs(dC) < 0.15 ? "✔" : (dC < 0 ? "Saturado" : "Opaco");
        public static string GetDiagH_Expert(double dH, double tol) => Math.Abs(dH) < tol ? "✔" : (dH < 0 ? "Viraje (+)" : "Viraje (-)");


        /// Determina la dirección visual del viraje de tono
        public static string GetHueDirection(double dA, double dB)
        {
            if (Math.Abs(dA) >= Math.Abs(dB))
                return dA < 0 ? " Virado al Rojo" : " Virado al Verde";
            else
                return dB < 0 ? " Virado al Amarillo" : " Virado al Azul";
        }

        // --- MÓDULO DE INGENIERÍA TEXTIL (DIAGNÓSTICO FINAL) ---
        public static string GetEngineeringDiagnosis(string eje, double delta, string impacto)
        {
            switch (eje.ToUpper())
            {
                case "DL":
                case "L":
                    if (delta < 0) 
                        return " Mas Claro ";
                    else
                        return "Mas Oscuro";

                case "DC": 
                case "DA":
                case "C":
                    if (impacto.Contains("Brillante") || delta > 0) 
                        return "Mas Brillante";
                    else
                        return "Mas Opaco";

                case "DH":
                case "DB":
                case "H":
                    if (delta < 0) 
                        return "Tendecia al  Amarillo.";
                    else 
                        return "Tendencia al AZul.";
            }

            return "DENTRO DE TOLERANCIA:";
        }

        // Lógica de Ciclo Industrial (Matriz Diagonal)
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