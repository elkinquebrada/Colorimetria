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
        
        // Los 3 Escenarios de Resta Estricta
        public double R1 { get; set; }
        public double R2 { get; set; }
        public double R3 { get; set; }
        
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

            // --- LÓGICA DE RESTA ESTRICTA ---
            // 1. Obtener ajustes y ordenar por magnitud (valor absoluto)
            // Usamos FactorL, FactorA, FactorB como los subtrahendos directos
            var listaAjustes = new List<double> { (double)analysis.FactorL, (double)analysis.FactorA, (double)analysis.FactorB }
                .OrderByDescending(x => Math.Abs(x))
                .ToList();

            // REGLA NUEVA: Resta de valores absolutos (Ignora la ley de signos) - SOLO PARA TABLA DE FORMULACIÓN
            double adj1 = Math.Abs(listaAjustes[0]);
            double adj2 = Math.Abs(listaAjustes[1]);
            double adj3 = Math.Abs(listaAjustes[2]);

            // Inyectar nombres reales para compatibilidad con diagnósticos y guardado
            var sortedByConc = originalRecipe.OrderByDescending(i => i.Percentage).ToList();
            var primary = sortedByConc.Count > 0 ? sortedByConc[0] : null;
            var secondary = sortedByConc.Count > 1 ? sortedByConc[1] : null;
            var toner = sortedByConc.Count > 2 ? sortedByConc[2] : sortedByConc.LastOrDefault();

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
                    Status = "OK",
                    // REGLA NUEVA: Original - Valor Absoluto del Ajuste
                    R1 = Math.Max(0, ing.Percentage - adj1),
                    R2 = Math.Max(0, ing.Percentage - adj2),
                    R3 = Math.Max(0, ing.Percentage - adj3)
                };

                result.Ingredients.Add(detail);
            }

            bool requiresAnyAdjustment = Math.Abs(adj1) > 0.0001 || Math.Abs(adj2) > 0.0001 || Math.Abs(adj3) > 0.0001;

            if (requiresAnyAdjustment)
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