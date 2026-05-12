using System;
using System.Collections.Generic;
using System.Linq;

namespace Color
{
    public class RecipeIngredientInput
    {
        public string Code { get; set; }
        public string Name { get; set; }
        public double Percentage { get; set; }   
    }

    // Modelo para exportación Excel y compatibilidad
    public class IlluminantCorrectionResult
    {
        public string Illuminant { get; set; }
        public double dl { get; set; }
        public double da { get; set; }
        public double Variaciondl { get; set; }
        public double Variacionda { get; set; }
        public double TotalOriginal { get; set; }
        public double DeltaE { get; set; }
    }

    public class CorrectiveIngredientDetail
    {
        public string Code { get; set; }
        public string Name { get; set; }
        public double Original { get; set; }
        
        // Los 3 Escenarios de Fase 2 (Diagonal Matrix Logic)
        public string Optiondl { get; set; }
        public string Optionda { get; set; }
        public string Optiondb { get; set; }
        
        public string Status { get; set; }
        public bool IsCritical { get; set; }
    }

    public class CorrectiveRecipeResult
    {
        public string Illuminant { get; set; }
        public List<CorrectiveIngredientDetail> Ingredients { get; set; } = new List<CorrectiveIngredientDetail>();
        public double TotalOriginal { get; set; }
        public string AlertMessage { get; set; }
        public string AlertSeverity { get; set; }
    }

    public static class RecipeCorrector
    {
        public static CorrectiveRecipeResult CalculateCorrectiveRecipe(
            List<RecipeIngredientInput> originalRecipe,
            ColorCorrectionResult analysis)
        {
            var result = new CorrectiveRecipeResult
            {
                Illuminant = analysis.Illuminant,
                TotalOriginal = originalRecipe.Sum(i => i.Percentage)
            };

            // Factores de Variación Relativa (Excel Parity)
            double fL = (double)(1.0m - analysis.FactorL);
            double fC = (double)(1.0m + analysis.FactorC);
            double fH = (double)(1.0m + (analysis.FactorH / 100.0m));

            // Ordenar por concentración para identificar roles
            var sorted = originalRecipe.OrderByDescending(i => i.Percentage).ToList();
            var primary = sorted.Count > 0 ? sorted[0] : null;
            var secondary = sorted.Count > 1 ? sorted[1] : null;
            var toner = sorted.Count > 2 ? sorted[2] : sorted.LastOrDefault();

            // Inyectar nombres reales (Solo el nombre, sin código) en el análisis para recomendaciones dinámicas
            if (analysis != null)
            {
                if (primary != null) analysis.PrimaryDyeName = primary.Name.Trim();
                if (secondary != null) analysis.SecondaryDyeName = secondary.Name.Trim();
                if (toner != null) analysis.TonerDyeName = toner.Name.Trim();
            }

            foreach (var ing in originalRecipe)
            {
                var detail = new CorrectiveIngredientDetail
                {
                    Code = ing.Code,
                    Name = ing.Name,
                    Original = ing.Percentage,
                    Status = "OK"
                };

                // ESCENARIO 1 (dl): AJUSTE GLOBAL DE CARGA (Afecta a todos simultáneamente)
                double valL = ing.Percentage * fL;
                double diffL = ((valL - ing.Percentage) / ing.Percentage) * 100.0;
                detail.Optiondl = $"{valL:F5} ({(diffL >= 0 ? "+" : "")}{diffL:F2}%)";
                if (Math.Abs(valL - ing.Percentage) / ing.Percentage > 0.15) { detail.IsCritical = true; detail.Status = "REVISAR"; }

                // ESCENARIO 2 (da): Solo secundario (Brillo)
                if (secondary != null && ing.Code == secondary.Code)
                {
                    double valC = ing.Percentage * fC;
                    double diffC = ((valC - ing.Percentage) / ing.Percentage) * 100.0;
                    detail.Optionda = $"{valC:F5} ({(diffC >= 0 ? "+" : "")}{diffC:F2}%)";
                    if (Math.Abs(valC - ing.Percentage) / ing.Percentage > 0.15) { detail.IsCritical = true; detail.Status = "REVISAR"; }
                }
                else detail.Optionda = "---";

                // ESCENARIO 3 (db): Solo toner
                if (toner != null && ing.Code == toner.Code)
                {
                    double valH = ing.Percentage * fH;
                    double diffH = ((valH - ing.Percentage) / ing.Percentage) * 100.0;
                    detail.Optiondb = $"{valH:F5} ({(diffH >= 0 ? "+" : "")}{diffH:F2}%)";
                    if (Math.Abs(valH - ing.Percentage) / ing.Percentage > 0.15) { detail.IsCritical = true; detail.Status = "REVISAR"; }
                }
                else detail.Optiondb = "---";

                result.Ingredients.Add(detail);
            }

            result.AlertMessage = result.Ingredients.Any(i => i.IsCritical) ? "ALERTA: Ajustes > 15%" : "Escenarios OK";
            result.AlertSeverity = result.Ingredients.Any(i => i.IsCritical) ? "Warning" : "None";

            return result;
        }

        public static List<RecipeIngredientInput> IngredientsFromShade(ShadeExtractionResult shadeData)
        {
            if (shadeData?.Recipe == null) return new List<RecipeIngredientInput>();
            return shadeData.Recipe.Select(r => new RecipeIngredientInput {
                Code = r.Code,
                Name = r.Name,
                Percentage = ParsePct(r.Percentage)
            }).ToList();
        }

        private static double ParsePct(string val)
        {
            if (string.IsNullOrWhiteSpace(val)) return 0;
            string clean = val.Replace("%", "").Trim().Replace(",", ".");
            if (double.TryParse(clean, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double res))
                return res;
            return 0;
        }
    }
}