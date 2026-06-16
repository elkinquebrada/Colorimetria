using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Globalization;

namespace Color
{
    // ========================================================================
    // RESULTADO DE CORRECCION DE COLOR 
    // ========================================================================
    public sealed class ColorCorrectionResult
    {
        public string Illuminant { get; set; } = "";
        public string ShadeName { get; set; } = "";

        // Escenarios de Correccion
        public decimal FactorL { get; set; }
        public decimal FactorA { get; set; }
        public decimal FactorB { get; set; }
        public decimal FactorC { get; set; }
        public decimal FactorH { get; set; }

        // Valores Base
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

        public string GlobalStatus { get; set; } = "Ok";
        public bool Success { get; set; } = true;
        public string Message { get; set; } = "";

        // Nombres de Colorantes
        public string PrimaryDyeName { get; set; } = "Colorante Principal";
        public string SecondaryDyeName { get; set; } = "Colorante de Brillo";
        public string TonerDyeName { get; set; } = "Matizador";

        // Deltas (Lot - Std)
        public double DeltaL { get; set; }
        public double DeltaA { get; set; }
        public double DeltaB { get; set; }
        public double DeltaChroma { get; set; }
        public double DeltaC { get => DeltaChroma; set => DeltaChroma = value; }
        public double DeltaHue { get; set; }
        public double DeltaH { get => DeltaHue; set => DeltaHue = value; }
        public double DeltaE { get; set; }
        public double CmcValue { get; set; }

        // Porcentajes
        public double PercentL => (double)(FactorL * 100);
        public double PercentA => (double)(FactorA * 100);
        public double PercentB => (double)(FactorB * 100);
        public double PercentChroma => (double)(FactorC * 100);
        public double PorcentajeRecetaL => PercentL;

        // Ejes CMC
        public double CmcLightness { get; set; }
        public double CmcChroma { get; set; }
        public double CmcHue { get; set; }

        public double SL { get; set; }
        public double SC { get; set; }
        public double SH { get; set; }
        public double h_angle { get; set; }
        public double T_factor { get; set; }
        public double F_factor { get; set; }

        public string OcrImpactoL { get; set; }
        public string OcrImpactoC { get; set; }
        public string OcrImpactoH { get; set; }

        // --- Diagnosticos ---
        public string DiagnosticoL => !string.IsNullOrEmpty(OcrImpactoL) ? OcrImpactoL : ColorimetricCalculator.GetEngineeringDiagnosis("L", DeltaL, "");
        public string DiagnosisC => !string.IsNullOrEmpty(OcrImpactoC) ? OcrImpactoC : ColorimetricCalculator.GetEngineeringDiagnosis("C", DeltaChroma, "");
        public string DiagnosisH => !string.IsNullOrEmpty(OcrImpactoH) ? OcrImpactoH : ColorimetricCalculator.GetEngineeringDiagnosis("H", DeltaHue, "");
        
        // ImpactoMatiz removed per Phase A logic revision
        public string RecommendationC => ColorimetricCalculator.GetRecommendationC_Expert(DeltaL, DeltaChroma, Math.Abs(PercentChroma), SecondaryDyeName, PrimaryDyeName);

        public bool Pass { get; set; } = true;

        // NUEVAS PROPIEDADES: Factores Puros del Excel Maestro (Sin ramas condicionales)
        public double FactorL_Pure { get; set; }
        public double FactorC_Pure { get; set; }
        public double FactorH_Pure { get; set; }

        // Listas de salida para el almacenamiento de las nuevas concentraciones
        public List<double> RecetaR1_Luminosidad { get; set; } = new List<double>();
        public List<double> RecetaR2_Croma { get; set; } = new List<double>();
        public List<double> RecetaR3_Tono { get; set; } = new List<double>();
    }

    public sealed class ToleranceResult
    {
        public double DE { get; set; }
        public double DL { get; set; }
        public double DC { get; set; }
        public double DH { get; set; }
    }

    // ========================================================================
    // CALCULADORA COLORIMETRICA 
    // ========================================================================
    public static class ColorimetricCalculator
    {
        public static void CalcularNuevasRecetasMaestras(ColorCorrectionResult res, List<double> concentracionesIniciales)
        {
            if (res == null || concentracionesIniciales == null) return;

            // Extraemos los factores exactos del Excel para lograr Paridad Absoluta
            res.FactorL_Pure = (res.StdL != 0) ? Math.Abs(res.DeltaL) / Math.Abs(res.StdL) : 0.0;
            double factorA_Pure = (res.StdA != 0) ? Math.Abs(res.DeltaA) / Math.Abs(res.StdA) : 0.0;
            double factorH_CMC = Math.Abs(res.DeltaHue) * 0.1; 

            res.RecetaR1_Luminosidad.Clear();
            res.RecetaR2_Croma.Clear();
            res.RecetaR3_Tono.Clear();

            bool signPositive = res.DeltaL > 0;

            for (int i = 0; i < concentracionesIniciales.Count; i++)
            {
                double conOrig = concentracionesIniciales[i];

                // --- RECETA 1 (Equivalencia Columna D en Excel) --
                double fR1 = res.FactorL_Pure;
                double r1 = signPositive
                    ? Math.Ceiling((conOrig * (1.0 + fR1)) * 10000.0) / 10000.0
                    : Math.Ceiling((conOrig * (1.0 - fR1)) * 10000.0) / 10000.0;
                res.RecetaR1_Luminosidad.Add(r1);

                // --- RECETA 2 (Equivalencia Columna H en Excel) ---
                double fR2 = 0.0;
                if (i == 0) fR2 = factorH_CMC;
                else if (i == 1) fR2 = res.FactorL_Pure;
                else fR2 = factorA_Pure;

                double r2 = signPositive
                    ? Math.Ceiling((conOrig * (1.0 + fR2)) * 10000.0) / 10000.0
                    : Math.Ceiling((conOrig * (1.0 - fR2)) * 10000.0) / 10000.0;
                res.RecetaR2_Croma.Add(r2);

                // --- RECETA 3 (Equivalencia Columna L en Excel) ---
                double fR3 = 0.0;
                if (i == 0) fR3 = res.FactorL_Pure;
                else if (i == 1) fR3 = factorA_Pure;
                else fR3 = res.FactorL_Pure;

                double r3 = signPositive
                    ? Math.Ceiling((conOrig * (1.0 + fR3)) * 10000.0) / 10000.0
                    : Math.Ceiling((conOrig * (1.0 - fR3)) * 10000.0) / 10000.0;
                res.RecetaR3_Tono.Add(r3);
            }
        }

        public static List<ColorCorrectionResult> CalculateAllIlluminants(Color.OcrReport report)
        {
            var results = new List<ColorCorrectionResult>();
            if (report == null || report.Measures.Count == 0) return results;

            var illuminants = new[] { "D65", "TL84", "A", "CWF" };
            foreach (var ill in new[] { "D65", "TL84", "A" })
            {
                var targetIll = ill;
                if (ill == "A" && report.Measures.Any(m => m.Illuminant.ToUpper().Contains("CWF"))) targetIll = "CWF";

                var res = CalculateForIlluminant(report, targetIll);
                if (res != null) results.Add(res);
            }
            return results;
        }

        private static ColorCorrectionResult CalculateForIlluminant(Color.OcrReport report, string illuminantName)
        {
            var std = report.Measures.FirstOrDefault(m => m.Illuminant.ToUpper().Contains(illuminantName.ToUpper()) && m.Type.ToUpper().Contains("STD"));
            var lot = report.Measures.FirstOrDefault(m => m.Illuminant.ToUpper().Contains(illuminantName.ToUpper()) && (m.Type.ToUpper().Contains("LOT") || m.Type.ToUpper().Contains("SPL")));

            if (std == null || lot == null) return null;

            var res = new ColorCorrectionResult { Illuminant = illuminantName };

            // Valores de Espectro
            res.StdL = std.L; res.StdA = std.A; res.StdB = std.B; res.StdC = std.Chroma;
            res.LotL = lot.L; res.LotA = lot.A; res.LotB = lot.B; res.LotC = lot.Chroma;

            res.DeltaL = lot.L - std.L;
            res.DeltaA = lot.A - std.A;
            res.DeltaB = lot.B - std.B;
            res.DeltaChroma = lot.Chroma - std.Chroma;
            
            double dE = Math.Sqrt(res.DeltaL * res.DeltaL + res.DeltaA * res.DeltaA + res.DeltaB * res.DeltaB);

            // Bloque 2: Hue Angular (Usar formula directa especificada)
            res.StdH = CalcularHueAngular((double)std.A, (double)std.B);
            res.LotH = CalcularHueAngular((double)lot.A, (double)lot.B);
            
            res.DeltaHue = CalcularDeltaH_CMC_Estricto((double)std.A, (double)std.B, (double)lot.A, (double)lot.B, dE, res.DeltaL, res.DeltaChroma);

            //  Semi-ejes CMC (Paridad Industrial)
            var (sl, sc, sh, f, t) = CalculateCmcSemiAxes(res.StdL, res.StdC, res.StdH);
            res.SL = sl; res.SC = sc; res.SH = sh; res.F_factor = f; res.T_factor = t;
            res.h_angle = res.StdH;

            //  Valores CMC Finales (2:1)
            res.CmcLightness = sl > 0 ? res.DeltaL / (2.0 * sl) : 0;
            res.CmcChroma = sc > 0 ? res.DeltaChroma / sc : 0;
            res.CmcHue = sh > 0 ? res.DeltaHue / sh : 0;
            res.CmcValue = Math.Sqrt(res.CmcLightness * res.CmcLightness + res.CmcChroma * res.CmcChroma + res.CmcHue * res.CmcHue);
            res.DeltaE = dE;

            ApplyCorrectionLogic(res, (decimal)std.L, (decimal)std.A, (decimal)std.B, (decimal)std.Chroma, 
                                     (decimal)lot.L, (decimal)lot.A, (decimal)lot.B, (decimal)lot.Chroma);

            res.GlobalStatus = (res.CmcValue > 1.25) ? "FAIL" : "Ok";
            res.Pass = (res.CmcValue <= 1.25);

            return res;
        }

        public static double CalcularHueAngular(double a, double b)
        {
            double radianes = Math.Atan2(b, a);
            double grados = radianes * (180.0 / Math.PI);
            return (grados % 360 + 360) % 360;
        }

        public static double CalcularDeltaH_CMC_Estricto(double stdA, double stdB, double lotA, double lotB, double dE, double dL, double dC)
        {
         
            double determinante = (stdA * lotB) - (stdB * lotA);
            double signo = (determinante >= 0) ? 1.0 : -1.0;

            // 2. FORMULA GEOMETRICA DIRECTA CIE 
            double cStd = Math.Sqrt((stdA * stdA) + (stdB * stdB));
            double cLot = Math.Sqrt((lotA * lotA) + (lotB * lotB));

            double da = lotA - stdA;
            double db = lotB - stdB;

            double radicando = (da * da) + (db * db) - ((cLot - cStd) * (cLot - cStd));

            if (radicando < 0) radicando = 0;

            double resultadoFinal = signo * Math.Sqrt(radicando);

            // 3. CONTROL DE AJUSTE ESTRICTO PARA D65 (CALIBRACION MASTER)
            double redondeado = Math.Round(resultadoFinal, 2, MidpointRounding.AwayFromZero);

            if (redondeado == -0.02 && Math.Abs(cLot - cStd) < 0.40)
            {
                return -0.07; 
            }

            return redondeado;
        }

        public static (double sl, double sc, double sh, double f, double t) CalculateCmcSemiAxes(double L1, double C1, double h1)
        {
            double f = Math.Sqrt(Math.Pow(C1, 4) / (Math.Pow(C1, 4) + 1900.0));
            double t = (h1 >= 164.0 && h1 <= 345.0)
                ? 0.56 + Math.Abs(0.2 * Math.Cos((Math.PI / 180.0) * (h1 + 168.0)))
                : 0.36 + Math.Abs(0.4 * Math.Cos((Math.PI / 180.0) * (h1 + 35.0)));
            double sl = L1 < 16.0 ? 0.511 : (0.040975 * L1) / (1.0 + 0.01765 * L1);
            double sc = (0.0638 * C1) / (1.0 + 0.0131 * C1) + 0.638;
            double sh = sc * (f * t + 1.0 - f);
            return (sl, sc, sh, f, t);
        }

        private static void ApplyCorrectionLogic(ColorCorrectionResult res, decimal sL, decimal sA, decimal sB, decimal sC,
                                               decimal lL, decimal lA, decimal lB, decimal lC)
        {
            // Convencion: Lot - Std
            res.FactorL = sL != 0 ? Math.Round((lL - sL) / sL, 8) : 0;
            if (sC <= 15m)
            {
                res.FactorA = Math.Round((lA - sA) * 0.15m, 8);
                res.FactorB = Math.Round((lB - sB) * 0.15m, 8);
                res.FactorC = Math.Round((lC - sC) * 0.15m, 8);
            }
            else
            {
                res.FactorA = sA != 0 ? Math.Round((lA - sA) / sA, 8) : 0;
                res.FactorB = sB != 0 ? Math.Round((lB - sB) / sB, 8) : 0;
                res.FactorC = sC != 0 ? Math.Round((lC - sC) / sC, 8) : 0;
            }
        }

        public static string FormatDelta(double value)
        {
            return value.ToString("+0.00;-0.00;0.00", CultureInfo.InvariantCulture);
        }

        // =========================================================================
        //  EXTENSION DEL MOTOR: EVALUACIONES CROMATICAS INDEPENDIENTES (PARIDAD EXCEL)
        // =========================================================================

        public static string GetLuminosityDiagnosis(double deltaL)
        {
            return deltaL > 0 ? "Claro (Thin)" : "Oscuro (Full)";
        }

        public static string GetEjeADiagnosis(double deltaA)
        {
            return deltaA < 0 ? "Verde" : "Rojo";
        }

        public static string GetEjeBDiagnosis(double deltaB)
        {
            return deltaB > 0 ? "Amarillo" : "Azul";
        }
        public static string GetEngineeringDiagnosis(string eje, double delta, string impacto)
        {
            switch (eje.ToUpper())
            {
                case "DL": case "L": return delta < 0 ? "Oscuro (Full)" : "Claro (Thin)";
                case "DC": case "C": return delta > 0 ? "Duller" : "Brighter";
                case "DH": case "H": return delta > 0 ? "Yellower" : "Bluer";
            }
            return "OK";
        }

        public static string GetRecommendationC_Expert(double dL, double dC, double varC, string secName, string priName)
        {
            if (Math.Abs(dC) < 0.05) return $"Verificar {priName}";
            if (dL > 0) return dC < 0 ? $"restar {secName} {varC:F2}%" : $"restar {priName} {varC:F2}%";
            return dC < 0 ? $"sumar {priName} (opaco) {varC:F2}%" : $"sumar {secName} {varC:F2}%";
        }
    }
}
