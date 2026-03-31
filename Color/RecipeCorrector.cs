using System;
using System.Collections.Generic;
using System.Globalization;


namespace Color
{
    // ── Modelos de entrada ──────────────────────────────────────────────────

    public class RecipeIngredientInput
    {
        public string Code { get; set; }
        public string Name { get; set; }
        public double Percentage { get; set; }   
    }

    ///Deltas CMC de un iluminante (viene del OcrReport).
    public class IlluminantDelta
    {
        public string Illuminant { get; set; }  
        public double DeltaLightness { get; set; }  
        public double DeltaChroma { get; set; }  
    }

    // ── Modelos de resultado ────────────────────────────────────────────────

    public class IngredientCorrectionDetail
    {
        public string Name { get; set; }
        public double Original { get; set; }   
        public double Calc1 { get; set; }  
        public double Calc2 { get; set; }   
        public double CalcI { get; set; }   
        public double Calc3 { get; set; }   
        public double NewPct2 { get; set; }   
        public double NewPct3 { get; set; }  
    }

    ///Resultado completo de corrección para un iluminante.
    public class IlluminantCorrectionResult
    {
        public string Illuminant { get; set; }
        public double ConvLightness { get; set; }   
        public double ConvChroma { get; set; }   

        public List<IngredientCorrectionDetail> Ingredients { get; set; }
            = new List<IngredientCorrectionDetail>();

        // Sumas parciales
        public double SumCalc1 { get; set; }
        public double SumCalc2 { get; set; }
        public double SumCalcI { get; set; }
        public double SumCalc3 { get; set; }

        // Resultados finales
        public double ResultadoLightness { get; set; }  
        public double ResultadoChroma { get; set; }   

        public override string ToString()
        {
            return string.Format(
                "[{0}] convL={1:F1}  convC={2:F1}  ResLightness={3:F4}  ResChroma={4:F4}",
                Illuminant, ConvLightness, ConvChroma, ResultadoLightness, ResultadoChroma);
        }
    }

    // ── Calculadora principal ───────────────────────────────────────────────

    public static class RecipeCorrector
    {
        /// Calcula la corrección de la receta para cada iluminante.
        public static List<IlluminantCorrectionResult> Calculate(
            List<RecipeIngredientInput> ingredients,
            List<IlluminantDelta> deltas)
        {
            var results = new List<IlluminantCorrectionResult>();

            if (ingredients == null || ingredients.Count == 0) return results;
            if (deltas == null || deltas.Count == 0) return results;

            // Total de la receta
            double totalReceta = 0;
            foreach (var ing in ingredients)
                totalReceta += ing.Percentage;

            if (totalReceta <= 0) return results;

            foreach (var delta in deltas)
            {
                var result = new IlluminantCorrectionResult
                {
                    Illuminant = delta.Illuminant,
                    ConvLightness = delta.DeltaLightness * 10.0,
                    ConvChroma = delta.DeltaChroma * 10.0
                };

                double sumCalc1 = 0, sumCalc2 = 0, sumCalc3 = 0;

                // Calcular Calc1, Calc2, Calc3 por ingrediente
                foreach (var ing in ingredients)
                {
                    double calc1 = ing.Percentage / totalReceta;
                    double calc2 = ing.Percentage * (1.0 - result.ConvLightness / 100.0);
                    double calc3 = calc2 * (1.0 - result.ConvChroma / 100.0);

                    sumCalc1 += calc1;
                    sumCalc2 += calc2;
                    sumCalc3 += calc3;

                    result.Ingredients.Add(new IngredientCorrectionDetail
                    {
                        Name = ing.Name,
                        Original = ing.Percentage,
                        Calc1 = calc1,
                        Calc2 = calc2,   
                        Calc3 = calc3    
                    });
                }

                // Resultados finales (verificados contra Excel)
                result.SumCalc1 = sumCalc1;
                result.SumCalc2 = sumCalc2;
                result.SumCalc3 = sumCalc3;

                result.ResultadoLightness = (totalReceta > 0)
                    ? (sumCalc2 / totalReceta) - 1.0 : 0;
                result.ResultadoChroma = sumCalc3 - 1.0;

                results.Add(result);
            }

            return results;
        }

        /// Convierte un OcrReport + ShadeExtractionResult en los inputs necesarios.
        public static List<RecipeIngredientInput> IngredientsFromShade(
            ShadeExtractionResult shade)
        {
            var list = new List<RecipeIngredientInput>();
            if (shade == null || shade.Recipe == null) return list;

            foreach (var item in shade.Recipe)
            {
                // Parsear porcentaje (quitar el % final)
                string pctStr = (item.Percentage ?? "").Replace("%", "").Trim();
                double pct;
                if (double.TryParse(pctStr, NumberStyles.Float,
                    CultureInfo.InvariantCulture, out pct))
                {
                    list.Add(new RecipeIngredientInput
                    {
                        Code = item.Code,
                        Name = item.Name,
                        Percentage = pct
                    });
                }
            }
            return list;
        }

        /// Convierte los CmcDifferenceRow del OcrReport en IlluminantDelta.
        public static List<IlluminantDelta> DeltasFromReport(OcrReport report)
        {
            var list = new List<IlluminantDelta>();
            if (report == null || report.CmcDifferences == null) return list;

            foreach (var cmc in report.CmcDifferences)
            {
                list.Add(new IlluminantDelta
                {
                    Illuminant = cmc.Illuminant,
                    DeltaLightness = cmc.DeltaLightness,
                    DeltaChroma = cmc.DeltaChroma
                });
            }
            return list;
        }

        /// Genera un resumen de texto de los resultados de corrección.
        public static string BuildSummaryText(List<IlluminantCorrectionResult> results)
        {
            if (results == null || results.Count == 0)
                return "Sin resultados de corrección de receta.";

            var sb = new System.Text.StringBuilder();
            sb.AppendLine("══════════════════════════════════════════════════════════════════");
            sb.AppendLine("  CORRECCIÓN DE RECETA POR ILUMINANTE");
            sb.AppendLine("══════════════════════════════════════════════════════════════════");

            foreach (var r in results)
            {
                sb.AppendLine();
                sb.AppendLine(string.Format("  ── [{0}]  CMC(2:1)  convL={1:0.0}%  convC={2:0.0}%",
                    r.Illuminant, r.ConvLightness, r.ConvChroma));
                sb.AppendLine();

                // Encabezado tabla
                sb.AppendLine(string.Format("  {0,-28} {1,10} {2,8} {3,14} {4,8} {5,14} {6,8}",
                    "", "Original%", "Calc1%",
                    "Calc2(corrL)", "Calc1%",
                    "Calc3(corrL+C)", "Calc1%"));
                sb.AppendLine("  " + new string('─', 95));

                foreach (var ing in r.Ingredients)
                {
                    sb.AppendLine(string.Format("  {0,-28} {1,10:F5} {2,8:P1} {3,14:F8} {4,8:P1} {5,14:F8} {6,8:P1}",
                        ing.Name,
                        ing.Original,
                        ing.Calc1,
                        ing.Calc2,
                        ing.Calc1,
                        ing.Calc3,
                        ing.Calc1));
                }

                sb.AppendLine("  " + new string('─', 95));
                sb.AppendLine(string.Format("  {0,-28} {1,10:F5} {2,8}  {3,14:F8} {4,8}  {5,14:F8}",
                    "resultado parcial",
                    r.SumCalc1,
                    "100%",
                    r.SumCalc2,
                    "100%",
                    r.SumCalc3));

                // Fila resultado final: etiquetas Lightness/Chroma con sus valores
                sb.AppendLine(string.Format("  {0,-28} {1,10} {2,8}  {3,-9} {4,13:P1}  {5,-9} {6,13:P1}",
                    "resultado final",
                    "", "",
                    "Lightness",
                    r.ResultadoLightness,
                    "Chroma",
                    r.ResultadoChroma));
            }

            sb.AppendLine();
            sb.AppendLine("══════════════════════════════════════════════════════════════════");
            return sb.ToString();
        }
    }
}