using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Color
{
    // ── Modelos de entrada ──────────────────────────────────────────────────

    public class RecipeIngredientInput
    {
        public string Code { get; set; }
        public string Name { get; set; }
        public double Percentage { get; set; }   
    }

    /// Deltas CMC 2:1 para un iluminante.
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
        public double Calc3 { get; set; }   
    }

    ///Resultado completo de corrección para un iluminante.
    public class IlluminantCorrectionResult
    {
        public string Illuminant { get; set; }
        public double dL { get; set; }
        public double dC { get; set; }
        public double VariacionL { get; set; }
        public double VariacionC { get; set; }
        public double TotalOriginal { get; set; } 

        public List<IngredientCorrectionDetail> Ingredients { get; set; }
            = new List<IngredientCorrectionDetail>();

        // Sumas parciales
        public double SumCalc1 { get; set; }
        public double SumCalc2 { get; set; }
        public double SumCalc3 { get; set; }

        public double ResultadoLightness { get; set; }  
        public double ResultadoChroma { get; set; }   

        public override string ToString()
        {
            return string.Format(
                "[{0}] VarL={1:F1}% VarC={2:F1}% ResL={3:F2}% ResC={4:F2}%",
                Illuminant, VariacionL, VariacionC, ResultadoLightness * 100, ResultadoChroma * 100);
        }
    }

    // ── Calculadora principal ───────────────────────────────────────────────

    public static class RecipeCorrector
    {
        /// Calcula la corrección de la receta para cada iluminante usando la fómula secuencial de Excel.
        public static List<IlluminantCorrectionResult> Calculate(
            List<RecipeIngredientInput> ingredients,
            List<IlluminantDelta> deltas)
        {
            var results = new List<IlluminantCorrectionResult>();

            if (ingredients == null || ingredients.Count == 0) return results;
            if (deltas == null || deltas.Count == 0) return results;

            double totalReceta = ingredients.Sum(i => i.Percentage);
            if (totalReceta <= 0) return results;

            foreach (var d in deltas)
            {
                // Variación porcentual exacta de Excel: dL * 10
                double varL = d.DeltaLightness * 10.0;
                double varC = d.DeltaChroma * 10.0;

                var result = new IlluminantCorrectionResult
                {
                    Illuminant = d.Illuminant,
                    dL = d.DeltaLightness,
                    dC = d.DeltaChroma,
                    VariacionL = varL,
                    VariacionC = varC,
                    TotalOriginal = totalReceta
                };

                double sumCalc2 = 0, sumCalc3 = 0;

                foreach (var ing in ingredients)
                {
                    // Calculo 1: % de participación
                    double calc1 = ing.Percentage / totalReceta;
                    
                    // Calculo 2: Corrección por Lightness (dL * 10)
                    double calc2 = ing.Percentage * (1.0 + varL / 100.0);
                    
                    // Calculo 3: Corrección por Croma (dC * 10) sobre Calc2
                    // REVERSIÓN A ESCALA UNITARIA (Sincronización Espejo)
                    double calc3 = calc2 * (1.0 + varC / 100.0);

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

                result.SumCalc2 = sumCalc2;
                result.SumCalc3 = sumCalc3;
                result.ResultadoLightness = (sumCalc2 / totalReceta) - 1.0;
                // VAR CHROMÁTICA ESPEJO: Exceso sobre la unidad (ej: 1.633 -> 63.3%)
                result.ResultadoChroma = (sumCalc3 - 1.0);

                results.Add(result);
            }

            return results;
        }

        public static List<RecipeIngredientInput> IngredientsFromShade(ShadeExtractionResult shade)
        {
            var list = new List<RecipeIngredientInput>();
            if (shade == null || shade.Recipe == null) return list;

            foreach (var item in shade.Recipe)
            {
                string pctStr = (item.Percentage ?? "").Replace("%", "").Trim();
                if (double.TryParse(pctStr, NumberStyles.Float, CultureInfo.InvariantCulture, out double pct))
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

        public static string BuildSummaryText(List<IlluminantCorrectionResult> results)
        {
            if (results == null || results.Count == 0)
                return "Sin resultados de corrección de receta.";

            var sorted = results.OrderBy(r => {
                string i = r.Illuminant.ToUpper();
                if (i.Contains("D65")) return 0;
                if (i.Contains("TL84")) return 1;
                if (i.Contains("A")) return 2;
                return 100;
            }).ThenBy(r => r.Illuminant).ToList();

            var sb = new System.Text.StringBuilder();
            sb.AppendLine("════════════════════════════════════════════════════════════════════════════════════════");
            sb.AppendLine("  CÁLCULO DE RECETA (SINCRONIZADO CON EXCEL)");
            sb.AppendLine("════════════════════════════════════════════════════════════════════════════════════════");

            foreach (var r in sorted)
            {
                bool isD65 = r.Illuminant.ToUpper().Contains("D65");
                string tag = isD65 ? "«" : "";
                string endTag = isD65 ? "»" : "";

                sb.AppendLine();
                sb.AppendLine($"  RECETA ({r.Illuminant.ToUpper()})");
                sb.AppendLine("  " + new string('─', 84));

                sb.AppendLine(string.Format("  {0,-28} {1,14} {2,12} {3,14} {4,12}",
                    "", "Original", "Calculo 1", "Calculo 2", $"Chroma ({r.Illuminant})"));
                sb.AppendLine("  " + new string('─', 84));

                foreach (var ing in r.Ingredients)
                {
                    // Se aplica la etiqueta de importancia si es D65
                    sb.AppendLine(string.Format("  {0,-28} {1,14:F5} {2,12:P1} {3,14} {4,12}",
                        ing.Name, ing.Original, ing.Calc1, 
                        tag + ing.Calc2.ToString("F3", CultureInfo.InvariantCulture) + endTag,
                        tag + ing.Calc3.ToString("F3", CultureInfo.InvariantCulture) + endTag));
                }

                sb.AppendLine("  " + new string('─', 84));
                
                sb.AppendLine(string.Format("  {0,-28} {1,14:F5} {2,12} {3,14} {4,12}",
                    "  [Total]", r.TotalOriginal, "100%", 
                    tag + r.SumCalc2.ToString("F3", CultureInfo.InvariantCulture) + endTag,
                    tag + r.SumCalc3.ToString("F3", CultureInfo.InvariantCulture) + endTag));

                sb.AppendLine(string.Format("  {0,-28} {1,14} {2,12} {3,14} {4,12}",
                    "  Variación en la []", "", "Lightness", 
                    tag + (r.ResultadoLightness * 100.0).ToString("F1", CultureInfo.InvariantCulture) + "%" + endTag,
                    tag + (r.ResultadoChroma * 100.0).ToString("F1", CultureInfo.InvariantCulture) + "%" + endTag));
                
                sb.AppendLine("  " + new string('─', 84));
            }

            return sb.ToString();
        }

        public static string BuildConsolidatedCalculationTable(List<IlluminantCorrectionResult> results)
        {
            if (results == null || results.Count == 0) return "";

            var d65 = results.FirstOrDefault(r => r.Illuminant.ToUpper().Contains("D65"));
            var tl84 = results.FirstOrDefault(r => r.Illuminant.ToUpper().Contains("TL84"));
            var a = results.FirstOrDefault(r => r.Illuminant.ToUpper().Contains("A"));

            if (d65 == null || tl84 == null || a == null) return "";

            var sb = new System.Text.StringBuilder();
            sb.AppendLine();
            sb.AppendLine("════════════════════════════════════════════════════════════════════════════════════════");
            sb.AppendLine("  RESUMEN DE CÁLCULOS REALIZADOS (CHROMA)");
            sb.AppendLine("════════════════════════════════════════════════════════════════════════════════════════");
            
            // Cabecera con separadores verticales (Expandida)
            sb.AppendLine(string.Format("  {0,-28} | {1,14} | {2,14} | {3,14} | {4,14}",
                "Ingrediente", "Chroma (D65)", "Chroma (TL84)", "Chroma (A)", "Promedio Ch."));
            sb.AppendLine("  " + new string('─', 92));

            int ingCount = d65.Ingredients.Count;
            double sumAvg = 0;
            for (int i = 0; i < ingCount; i++)
            {
                var ingD65 = d65.Ingredients[i];
                var ingTL = tl84.Ingredients[i];
                var ingA = a.Ingredients[i];
                sumAvg += (ingD65.Calc3 + ingTL.Calc3 + ingA.Calc3) / 3.0;
            }

            for (int i = 0; i < ingCount; i++)
            {
                var ingD65 = d65.Ingredients[i];
                var ingTL = tl84.Ingredients[i];
                var ingA = a.Ingredients[i];
                double avg = (ingD65.Calc3 + ingTL.Calc3 + ingA.Calc3) / 3.0;

                // Solo la columna D65 se etiqueta con « » para aplicar NEGRITA en la UI
                string fmtD65 = string.Format(CultureInfo.InvariantCulture, "«{0,6:F3} | {1,4:P0}»", ingD65.Calc3, ingD65.Calc3 / d65.SumCalc3);
                string fmtTL = string.Format(CultureInfo.InvariantCulture, "{0,6:F3} | {1,4:P0}", ingTL.Calc3, ingTL.Calc3 / tl84.SumCalc3);
                string fmtA = string.Format(CultureInfo.InvariantCulture, "{0,6:F3} | {1,4:P0}", ingA.Calc3, ingA.Calc3 / a.SumCalc3);
                string fmtAvg = string.Format(CultureInfo.InvariantCulture, "{0,6:F3} | {1,4:P0}", avg, avg / sumAvg);

                sb.AppendLine(string.Format("  {0,-28} | {1,14} | {2,14} | {3,14} | {4,14}",
                    ingD65.Name, fmtD65, fmtTL, fmtA, fmtAvg));
                
                if (i < ingCount - 1)
                     sb.AppendLine("  " + new string('┈', 92));
            }

            sb.AppendLine("  " + new string('─', 92));

            // Fila de Totales
            string totD65 = string.Format(CultureInfo.InvariantCulture, "«{0,6:F3} | {1,4}»", d65.SumCalc3, "100%");
            string totTL = string.Format(CultureInfo.InvariantCulture, "{0,6:F3} | {1,4}", tl84.SumCalc3, "100%");
            string totA = string.Format(CultureInfo.InvariantCulture, "{0,6:F3} | {1,4}", a.SumCalc3, "100%");
            string totAvg = string.Format(CultureInfo.InvariantCulture, "{0,6:F3} | {1,4}", sumAvg, "100%");

            sb.AppendLine(string.Format("  {0,-28} | {1,14} | {2,14} | {3,14} | {4,14}",
                "  [Total]", totD65, totTL, totA, totAvg));

            // Fila de Variaciones
            double varAvg = (d65.ResultadoChroma + tl84.ResultadoChroma + a.ResultadoChroma) / 3.0;
            
            sb.AppendLine(string.Format("  {0,-28} | {1,14} | {2,14} | {3,14} | {4,14}",
                "  Variación (Croma %)", 
                string.Format(CultureInfo.InvariantCulture, "«{0,8:F1}%»     ", d65.ResultadoChroma * 100.0),
                string.Format(CultureInfo.InvariantCulture, "{0,8:F1}%      ", tl84.ResultadoChroma * 100.0),
                string.Format(CultureInfo.InvariantCulture, "{0,8:F1}%      ", a.ResultadoChroma * 100.0),
                string.Format(CultureInfo.InvariantCulture, "{0,8:F1}%      ", varAvg * 100.0)));

            sb.AppendLine("  " + new string('─', 92));

            return sb.ToString();
        }
    }
}