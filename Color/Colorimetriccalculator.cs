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
        public string ShadeName { get; set; } = "";

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
        public double PercentChroma { get; set; }
        public double PercentHue { get; set; }

        public double StdHue { get; set; }
        public double LotHue { get; set; }

        // --- INSTRUCCIONES DE CORRECCIÓN PROFESIONAL ---

        // Determina la acción sobre la claridad
        public string LightnessInstruction { get; set; } = "";

        // Determina la acción para el eje a* (Rojo/Verde)
        public string CorrectionA { get; set; } = "";

        // Determina la acción para el eje b* (Amarillo/Azul)
        public string CorrectionB { get; set; } = "";

        // Determina la acción para Chroma
        public string ChromaInstruction { get; set; } = "";

        // ΔE CMC(2:1) - Recomendado para grado comercial (Elipse de tolerancia)
        public double CmcValue { get; set; }

        // --- PROPIEDADES DINÁMICAS PARA EL REPORTE (Expert System) ---
        public string DiagnosisL => ColorimetricCalculator.GetDiagL_Expert(DeltaL);
        // --- NUEVAS PROPIEDADES PARA PANELES SEPARADOS (ESTÁNDAR TEXTIL 2026) ---
        public double PorcentajeRecetaL => Math.Abs(PercentL * 100);
        
        public string ImpactoRecetaL => ColorimetricCalculator.GetImpactoLRecipe(DeltaL);
        public string ImpactoLoteL => ColorimetricCalculator.GetImpactoLLot(DeltaL);
        
        public string RecomendacionRecetaL => ColorimetricCalculator.GetInstLRecipe(DeltaL, PorcentajeRecetaL);
        public string RecomendacionLoteL => ColorimetricCalculator.GetInstLLot(DeltaL, PorcentajeRecetaL);
        
        // Diagnóstico técnico (mismo para ambos)
        public string DiagnosticoL => ColorimetricCalculator.GetDiagL_Expert(DeltaL);
        
        // Compatibilidad legacy
        public string DescripcionL => ImpactoRecetaL;
        public string RecomendacionL => RecomendacionRecetaL; 
        public string DiagnosticoLRecipe => DiagnosticoL;
        public string DiagnosticoLoteL => DiagnosticoL;

        public string RecomendacionMatiz
        {
            get
            {
                // Si Delta es positivo, significa que el Lote tiene de más -> DISMINUIR
                // Si Delta es negativo, significa que al Lote le falta -> AGREGAR
                string accionA = DeltaA > 0 ? "Disminuir" : "Agregar";
                string accionB = DeltaB > 0 ? "Disminuir" : "Agregar";
                
                string baseRec = $"{accionA} Rojo {Math.Abs(PercentA * 100):F1}% / {accionB} Amarillo {Math.Abs(PercentB * 100):F1}%";
                
                // Si el desvío es crítico, añadir instrucción experta
                if (Math.Abs(DeltaHue) > 0.4)
                {
                    string coloranteNeutralizador = DeltaHue > 0 ? "Azulado" : "Rojizo";
                    return $"AJUSTE DE MATIZ: Adicionar {Math.Abs(DeltaHue * 10):F1}% de colorante para neutralizar {coloranteNeutralizador}. " + baseRec;
                }
                return baseRec;
            }
        }

        public string ImpactoMatiz => ColorimetricCalculator.GetImpactH_Expert(DeltaHue);
        
        public string DiagnosisC => ColorimetricCalculator.GetDiagC_Expert(DeltaChroma);
        public string DescripcionC => DeltaChroma < 0 ? "Más Opaca / Sucia" : "Más Brillante / Vívida";
        public string RecommendationC => DeltaChroma > 0 ? $"DISMINUIR FUERZA {Math.Abs(PercentChroma * 100):F1}%" : $"AUMENTAR FUERZA {Math.Abs(PercentChroma * 100):F1}%";

        public string DiagnosisH => ColorimetricCalculator.GetDiagH_Expert(DeltaHue, 0.1); 

        // Impacto visual de dE
        public string ImpactoDE => ColorimetricCalculator.GetImpactDE_Expert(DeltaE);

        // Alerta de Metamerismo (se llena externamente si se evalúan múltiples iluminantes)
        public string MetamerismAlert { get; set; } = "";

        // Estado de aprobación basado en la tolerancia seleccionada
        public bool Pass { get; set; }
    }

    // ================================
    // RESULTADO CMC(2:1)
    // ================================
    public sealed class CmcResult
    {
        public string Illuminant { get; set; } = "";

        // Componentes CMC(2:1) calculadas desde ΔL, ΔChroma, ΔHue
        public double Lightness { get; set; }
        public double Chroma { get; set; }
        public double Hue { get; set; }
        public double CmcValue { get; set; }

        // Conversiones: valor absoluto × 10  (para uso en receta)
        public double ConversionLightness { get; set; }
        public double ConversionChroma { get; set; }
    }

    // ================================
    // RESULTADO DE RECETA POR ILUMINANTE
    // ================================
    public sealed class RecipeDyeResult
    {
        public string DyeName { get; set; } = "";
        public double OriginalAmount { get; set; }
        public double Calc1Normalized { get; set; }
        public double Calc2Amount { get; set; }
        public double Calc2Normalized { get; set; }
        public double Calc3Amount { get; set; }
        public double Calc3Normalized { get; set; }
    }

    public sealed class RecipeResult
    {
        public string Illuminant { get; set; } = "";
        public List<RecipeDyeResult> Dyes { get; set; } = new List<RecipeDyeResult>();

        public double TotalOriginal { get; set; }
        public double TotalCalc2 { get; set; }
        public double TotalCalc3 { get; set; }

        public double VariationLightness { get; set; }
        public double VariationChroma { get; set; }
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

        // --- CONSTANTES DE INSTRUCCIÓN PROFESIONAL ---
        public const string MSG_DECREASE_RED = "DISMINUIR ROJO / AUMENTAR VERDE";
        public const string MSG_INCREASE_RED = "AUMENTAR ROJO / DISMINUIR VERDE";
        public const string MSG_DECREASE_YELLOW = "DISMINUIR AMARILLO / AUMENTAR AZUL";
        public const string MSG_INCREASE_YELLOW = "AUMENTAR AMARILLO / DISMINUIR AZUL";
        public const string MSG_DARKEN = "OSCURECER";
        public const string MSG_LIGHTEN = "ACLARAR";
        public const string MSG_OK = "OK / DENTRO DE NORMA";
    }

    // ================================
    // EVALUACIÓN DE TOLERANCIA POR ILUMINANTE
    // ================================
    public sealed class IlluminantToleranceCheck
    {
 
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
        ///Banda de tolerancia usada.
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

                // --- CÁLCULOS TRIGONOMÉTRICOS (Propuesta Técnica) ---
                double chromaStd = Math.Sqrt(std.A * std.A + std.B * std.B);
                double chromaLot = Math.Sqrt(lot.A * lot.A + lot.B * lot.B);
                double dChroma = chromaLot - chromaStd;

                // Hue (h°) usando atan2(b*, a*)
                double hStd = Math.Atan2(std.B, std.A) * (180.0 / Math.PI);
                if (hStd < 0) hStd += 360.0;
                
                double hLot = Math.Atan2(lot.B, lot.A) * (180.0 / Math.PI);
                if (hLot < 0) hLot += 360.0;

                // Delta h (angular difference)
                double dhAngular = hLot - hStd;
                if (dhAngular > 180) dhAngular -= 360;
                if (dhAngular < -180) dhAngular += 360;

                // Delta Hue (ΔH) simplificado para reporte (Fidelidad al Estándar Excel)
                double dHueSimp = Math.Sqrt(Math.Max(0, Math.Pow(dA, 2) + Math.Pow(dB, 2) - Math.Pow(dChroma, 2)));

                // --- % REALES PARA DIAGNÓSTICO (Ratios base para multiplicar por 100 en UI) ---
                double pctL_Ratio = (std.L > 0.1) ? (dL / std.L) : 0.0;
                double pctA_Ratio = (Math.Abs(std.A) > 0.1) ? (dA / Math.Abs(std.A)) : 0.0;
                double pctB_Ratio = (Math.Abs(std.B) > 0.1) ? (dB / Math.Abs(std.B)) : 0.0;
                double pctChromaRatio = (chromaStd > 0.1) ? (dChroma / chromaStd) : 0.0;
                double pctHueRatio = (chromaStd > 0.1) ? (dHueSimp / chromaStd) : 0.0;

                // --- FACTOR ACCIONABLE (Sincronizado con Cálculos de Receta) ---
                double actPctL = Math.Abs(dL) * 10.0;
                double actPctA = Math.Abs(dA) * 10.0;
                double actPctB = Math.Abs(dB) * 10.0;

                // --- LÓGICA DE CORRECCIÓN PROFESIONAL (INVERSA AL DELTA) ---
                string lightnessInst = dL < 0 ? $"{ToleranceResult.MSG_LIGHTEN} «({actPctL:F1}%)»" 
                                     : dL > 0 ? $"{ToleranceResult.MSG_DARKEN} «({actPctL:F1}%)»" 
                                     : ToleranceResult.MSG_OK;

                // Corrección A: Si dA > 0 el lote está muy Rojo -> Quitar Rojo/Poner Verde
                string correctionA = dA > 0 ? $"{ToleranceResult.MSG_DECREASE_RED} «({actPctA:F1}%)»"
                                   : dA < 0 ? $"{ToleranceResult.MSG_INCREASE_RED} «({actPctA:F1}%)»" 
                                   : ToleranceResult.MSG_OK;

                // Corrección B: Si dB > 0 el lote está muy Amarillo -> Quitar Amarillo/Poner Azul
                string correctionB = dB > 0 ? $"{ToleranceResult.MSG_DECREASE_YELLOW} «({actPctB:F1}%)»"
                                   : dB < 0 ? $"{ToleranceResult.MSG_INCREASE_YELLOW} «({actPctB:F1}%)»" 
                                   : ToleranceResult.MSG_OK;

                string chromaInst = RecipeCorrector.ObtenerDiagnosticoChroma(dChroma, chromaStd);

                results.Add(new ColorCorrectionResult
                {
                    Illuminant = illuminant,
                    DeltaL = Math.Round(dL, 4),
                    DeltaA = Math.Round(dA, 4),
                    DeltaB = Math.Round(dB, 4),
                    AbsDeltaL = Math.Round(Math.Abs(dL), 4),
                    AbsDeltaA = Math.Round(Math.Abs(dA), 4),
                    AbsDeltaB = Math.Round(Math.Abs(dB), 4),
                    DeltaChroma = Math.Round(dChroma, 4),
                    DeltaHue = Math.Round(dHueSimp, 4),
                    DeltaE = Math.Round(dE, 4),
                    PercentL = Math.Round(pctL_Ratio, 6),
                    PercentA = Math.Round(pctA_Ratio, 6),
                    PercentB = Math.Round(pctB_Ratio, 6),
                    PercentChroma = Math.Round(pctChromaRatio, 6),
                    PercentHue = Math.Round(pctHueRatio, 6),
                    StdHue = Math.Round(hStd, 2),
                    LotHue = Math.Round(hLot, 2),
                    LightnessInstruction = lightnessInst,
                    ChromaInstruction = chromaInst,
                    CorrectionA = correctionA,
                    CorrectionB = correctionB,
                    Pass = false   
                });
            }

            // --- EVALUACIÓN DE METAMERISMO (Propuesta Técnica 3) ---
            var d65Res = results.FirstOrDefault(r => r.Illuminant.Contains("D65"));
            var tl84Res = results.FirstOrDefault(r => r.Illuminant.Contains("TL84"));
            if (d65Res != null && tl84Res != null)
            {
                double diffDE = Math.Abs(d65Res.DeltaE - tl84Res.DeltaE);
                if (diffDE > 0.3)
                {
                    string alert = $"ALTA INCONSISTENCIA: Muestra metamérica bajo luz de tienda (ΔΔE={diffDE:F2})";
                    d65Res.MetamerismAlert = alert;
                    tl84Res.MetamerismAlert = alert;
                }
            }

            return results;
        }

        // --- LÓGICA DE DIAGNÓSTICO EXPERTO (Propuesta Técnica 1, 2 y 3) ---

        public static string GetDiagL_Expert(double dL)
        {
            double absDL = Math.Abs(dL);
            if (absDL > 0.5) return "Desviación Crítica: Error de pesaje o sustrato contaminado";
            if (absDL > 0.2) return "Desviación Moderada: Revisar relación de baño y agotamiento";
            return "Luminosidad dentro de Tolerancia";
        }

        public static string GetActionL_Expert(double dL, double percentL)
        {
            return dL > 0 
                ? $"AUMENTAR RECETA {Math.Abs(percentL * 100):F1}%" 
                : $"REDUCIR RECETA {Math.Abs(percentL * 100):F1}%";
        }

        // --- NUEVA LÓGICA DE RECOMENDACIÓN DINÁMICA ---
        
        /// <summary>
        /// A. Para el Panel de RECETA (Laboratorio)
        /// </summary>
        public static string GetImpactoLRecipe(double dL) => dL > 0 ? "Más Claro" : "Más Oscuro";

        public static string GetInstLRecipe(double dL, double varL) 
        {
            string accion = dL > 0 ? "AUMENTAR" : "REDUCIR";
            return $"{accion} % TOTAL RECETA EN {Math.Abs(varL):F1}%";
        }

        /// <summary>
        /// B. Para el Panel de LOTE (Planta/Proceso)
        /// </summary>
        public static string GetImpactoLLot(double dL) => dL > 0 ? "Más Claro" : "Más Oscuro";

        public static string GetInstLLot(double dL, double pctL) 
        {
            string accion = dL > 0 ? "AUMENTAR" : "DISMINUIR";
            return $"{accion} FUERZA {Math.Abs(pctL):F1}%";
        }

        public static string GetDiagC_Expert(double dC, double tolerance = 0.4)
        {
            if (dC < -tolerance) return "Muestra Opaca/Sucia: Posible hidrólisis o exceso de sales";
            if (dC > tolerance) return "Muestra Brillante/Vivida: Revisar pureza de colorantes primarios";
            return dC < 0 ? "Muestra Opaca" : "Muestra Brillante";
        }

        public static string GetDiagH_Expert(double dH, double tolerance)
        {
            double absDH = Math.Abs(dH);
            if (absDH <= tolerance) return "Tono Estable";

            if (dH > 0)
                return absDH > 0.4 ? "Viraje CRÍTICO hacia Azulado" : "Desvío leve hacia Azulado";
            else
                return absDH > 0.4 ? "Viraje CRÍTICO hacia Rojizo" : "Desvío leve hacia Rojizo";
        }

        public static string GetImpactH_Expert(double dH)
        {
            if (dH > 0.1) return "Viraje hacia el siguiente color en espectro (más frío/azul)";
            if (dH < -0.1) return "Viraje hacia color anterior en espectro (más cálido/rojo)";
            return "Tono equilibrado";
        }

        public static string GetImpactDE_Expert(double dE)
        {
            return dE > 0.8 ? "Diferencia de color PERCEPTIBLE - Requiere corrección" : "Diferencia mínima";
        }

        // Retorna (sl, sc, sh)
        public static (double sl, double sc, double sh) CalculateCmcSemiAxes(double L1, double C1, double h1)
        {
            double f = Math.Sqrt(Math.Pow(C1, 4) / (Math.Pow(C1, 4) + 1900.0));
            double T;
            if (h1 >= 164.0 && h1 <= 345.0)
                T = 0.56 + Math.Abs(0.2 * Math.Cos(DegreeToRadian(h1 + 168.0)));
            else
                T = 0.36 + Math.Abs(0.4 * Math.Cos(DegreeToRadian(h1 + 35.0)));

            double sl = L1 < 16.0 ? 0.511 : (0.040975 * L1) / (1.0 + 0.01765 * L1);
            double sc = (0.0638 * C1) / (1.0 + 0.0131 * C1) + 0.638;
            double sh = sc * (f * T + 1.0 - f);
            return (sl, sc, sh);
        }

        // ------------------------------------------------------------------
        //    Fórmulas del Excel (hoja CALCULO RECETA): (CMC 2:1)
        // ------------------------------------------------------------------
        public static List<CmcResult> CalculateCmc(List<ColorCorrectionResult> corrections, List<ColorimetricRow> rows)
        {
            if (corrections == null || rows == null) return new List<CmcResult>();

            var results = new List<CmcResult>();
            double l = 2.0; // CMC(2:1)
            double c_val = 1.0;

            foreach (var cor in corrections)
            {
                var std = rows.FirstOrDefault(r => 
                    string.Equals(r.Illuminant, cor.Illuminant, StringComparison.OrdinalIgnoreCase) && 
                    string.Equals(r.Type, "Std", StringComparison.OrdinalIgnoreCase));
                
                if (std == null) continue;

                double L1 = std.L;
                double C1 = std.Chroma;
                double h1 = std.Hue;

                double dL = cor.DeltaL;
                double dC = cor.DeltaChroma;
                double dH = cor.DeltaHue;

                // --- Pesos CMC ---
                var axes = CalculateCmcSemiAxes(L1, C1, h1);
                double sl = axes.sl;
                double sc = axes.sc;
                double sh = axes.sh;

                // --- Diferencia CMC final ---
                double deCmc = Math.Sqrt(
                    Math.Pow(dL / (l * sl), 2) + 
                    Math.Pow(dC / (c_val * sc), 2) + 
                    Math.Pow(dH / sh, 2)
                );

                results.Add(new CmcResult
                {
                    Illuminant = cor.Illuminant,
                    Lightness = cor.DeltaL,
                    Chroma = cor.DeltaChroma,
                    Hue = cor.DeltaHue,
                    CmcValue = Math.Round(deCmc, 4),
                    ConversionLightness = Math.Round(Math.Abs(cor.DeltaL * 10.0), 4),
                    ConversionChroma = Math.Round(Math.Abs(cor.DeltaChroma * 10.0), 4)
                });
            }

            return results;
        }

        private static double DegreeToRadian(double angle) => (Math.PI / 180.0) * angle;

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
                Illuminant = illuminant,
                Lightness = cmcLightness,
                Chroma = cmcChroma,
                Hue = cmcHue,
                CmcValue = cmcValue,
                ConversionLightness = Math.Round(Math.Abs(cmcLightness * 10.0), 4),
                ConversionChroma = Math.Round(Math.Abs(cmcChroma * 10.0), 4)
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
                    DyeName = dyes[i].name,
                    OriginalAmount = dyes[i].amount,
                    Calc1Normalized = totalOriginal > 0
                        ? Math.Round(dyes[i].amount / totalOriginal, 6) : double.NaN,
                    Calc2Amount = Math.Round(calc2[i], 6),
                    Calc2Normalized = totalCalc2 > 0
                        ? Math.Round(calc2[i] / totalCalc2, 6) : double.NaN,
                    Calc3Amount = Math.Round(calc3[i], 6),
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
                Illuminant = illuminant,
                Dyes = dyeResults,
                TotalOriginal = Math.Round(totalOriginal, 6),
                TotalCalc2 = Math.Round(totalCalc2, 6),
                TotalCalc3 = Math.Round(totalCalc3, 6),
                VariationLightness = varLightness,
                VariationChroma = varChroma
            };
        }

        // ------------------------------------------------------------------
        //    Fórmula del Excel (hoja TOLERANCIA):
        // ------------------------------------------------------------------
        public static ToleranceResult CalculateTolerance(double de)
        {
            double component = Math.Sqrt((de * de) / 3.0);
            return new ToleranceResult
            {
                DE = Math.Round(de, 5),
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

                // --- PRECISIÓN CMC (ELIPSE) ---
                if (c.CmcValue > 0)
                {
                    if (c.CmcValue > band.DE) 
                        failing.Add("CMC");
                }
                else
                {
                    // Fallback a Delta E estándar si no hay CMC disponible
                    if (absDE > band.DE) failing.Add("DE");
                }

                if (absDL > band.DL) failing.Add("DL");
                if (absDC > band.DC) failing.Add("DC");
                if (absDH > band.DH) failing.Add("DH");

                bool passes = failing.Count == 0;

                // Propagar el estado de aprobación al resultado de corrección
                c.Pass = passes;

                evaluation.Checks.Add(new IlluminantToleranceCheck
                {
                    Illuminant = c.Illuminant,
                    Passes = passes,
                    MeasuredDE = Math.Round(absDE, 2),
                    MeasuredDL = Math.Round(c.DeltaL, 2),
                    MeasuredDC = Math.Round(c.DeltaChroma, 2),
                    MeasuredDH = Math.Round(c.DeltaHue, 2),
                    LimitDE = band.DE,
                    LimitDL = band.DL,
                    LimitDC = band.DC,
                    LimitDH = band.DH,
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