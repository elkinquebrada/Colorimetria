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

            // Factores de Variación Relativa (sin factores K de sensibilidad)
            double rawFL = (double)(1.0m - analysis.FactorL);
            double rawFC = (double)(1.0m + analysis.FactorC);
            double rawFH = (double)(1.0m + ((decimal)analysis.DeltaHue / 100.0m));
            
            double fL = rawFL;
            double fC = rawFC;
            double fH = rawFH;

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

            double totalRecipePct = originalRecipe.Sum(i => i.Percentage);

            foreach (var ing in originalRecipe)
            {
                double partPct = totalRecipePct > 0 ? (ing.Percentage / totalRecipePct) * 100.0 : 0;
                bool isLowPart = partPct < 2.0;

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
                bool capL = Math.Abs(rawFL - 1.0) > 0.15;
                string flagL = (capL && isLowPart) ? "\n* Sensibilidad Limitada" : "";
                detail.Optiondl = $"{valL:F5} ({(diffL >= 0 ? "+" : "")}{diffL:F2}%){flagL}";
                if (capL) { detail.IsCritical = true; detail.Status = "REV"; }

                // ESCENARIO 2 (da): Solo secundario (Brillo)
                if (secondary != null && ing.Code == secondary.Code)
                {
                    double valC = ing.Percentage * fC;
                    double diffC = ((valC - ing.Percentage) / ing.Percentage) * 100.0;
                    bool capC = Math.Abs(rawFC - 1.0) > 0.15;
                    string flagC = (capC && isLowPart) ? "\n* Sensibilidad Limitada" : "";
                    detail.Optionda = $"{valC:F5} ({(diffC >= 0 ? "+" : "")}{diffC:F2}%){flagC}";
                    if (capC) { detail.IsCritical = true; detail.Status = "REV"; }
                }
                else detail.Optionda = "---";

                // ESCENARIO 3 (db): Solo toner
                if (toner != null && ing.Code == toner.Code)
                {
                    double valH = ing.Percentage * fH;
                    double diffH = ((valH - ing.Percentage) / ing.Percentage) * 100.0;
                    bool capH = Math.Abs(rawFH - 1.0) > 0.15;
                    string flagH = (capH && isLowPart) ? "\n* Sensibilidad Limitada" : "";
                    detail.Optiondb = $"{valH:F5} ({(diffH >= 0 ? "+" : "")}{diffH:F2}%){flagH}";
                    if (capH) { detail.IsCritical = true; detail.Status = "REV"; }
                }
                else detail.Optiondb = "---";

                result.Ingredients.Add(detail);
            }

            bool isCritical = result.Ingredients.Any(i => i.IsCritical);
            bool requiresAnyAdjustment = Math.Abs(rawFL - 1.0) > 0.001 || Math.Abs(rawFC - 1.0) > 0.001 || Math.Abs(rawFH - 1.0) > 0.001;

            if (isCritical)
            {
                result.AlertMessage = "Ajuste Requerido - Límite de Seguridad Alcanzado";
                result.AlertSeverity = "Critical";
            }
            else if (requiresAnyAdjustment)
            {
                result.AlertMessage = "Revisar Ajustes Sugeridos";
                result.AlertSeverity = "Warning";
            }
            else
            {
                result.AlertMessage = "Lote Cumple Tolerancia - Sin Ajustes";
                result.AlertSeverity = "None";
            }

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