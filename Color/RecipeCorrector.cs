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

    public class CorrectiveIngredientDetail
    {
        public string Code { get; set; }
        public string Name { get; set; }
        public double Original { get; set; }
        public double FactorDL { get; set; }
        public double FactorDH { get; set; }
        public double NewConcentration { get; set; }
        public string Status { get; set; } // "OK", "SATURACIÓN", "REVISAR"
    }

    public class CorrectiveRecipeResult
    {
        public string Illuminant { get; set; }
        public List<CorrectiveIngredientDetail> Ingredients { get; set; } = new List<CorrectiveIngredientDetail>();
        public double TotalOriginal { get; set; }
        public double TotalNew { get; set; }
        public string AlertMessage { get; set; }
        public string AlertSeverity { get; set; } // "None", "Warning", "Critical", "Error"
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
        public static string ObtenerDiagnosticoChroma(double deltaC, double valorEstandarC)
        {
            // Evitar división por cero si el estándar es un gris puro
            if (valorEstandarC == 0) return "ESTABLE (0%)";
            
            double porcentaje = (deltaC / valorEstandarC) * 100;
            string instruccion = deltaC < 0 ? "AUMENTAR FUERZA / CONCENTRACIÓN" : "REDUCIR CARGA / DILUIR";
            
            return $"{instruccion} «({Math.Abs(porcentaje):F1}%)»";
        }

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
                    
                    double calc1 = ing.Percentage / totalReceta;
                    
                    // Calculo 2: Corrección por Lightness (dL * 10)
                    double calc2 = ing.Percentage * (1.0 + varL / 100.0);
                    
                    // Calculo 3: Corrección por Chroma (dC * 10) sobre Calc2
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

        // ── NUEVO MÓDULO DE RECETA CORRECTIVA (PROPUESTA COATS 2026) ──

        public static CorrectiveRecipeResult CalculateCorrectiveRecipe(
            List<RecipeIngredientInput> originalRecipe,
            ColorCorrectionResult analysis)
        {
            var result = new CorrectiveRecipeResult
            {
                Illuminant = analysis.Illuminant,
                TotalOriginal = originalRecipe.Sum(i => i.Percentage)
            };

            // 1. Factores de Entrada
            double factorDL = 1.0 + (analysis.DeltaL * 10.0 / 100.0);
            double factorDH = 1.0 + (analysis.DeltaHue * 10.0 / 100.0);
            
            // 2. Identificación del Matizador
            // Se asume que DH se aplica al colorante que necesita corregir el matiz.
            // Según propuesta: "Factor_DH se aplica exclusivamente al colorante identificado como Principal/Matizador"
            var matizador = IdentifyMatizador(originalRecipe, analysis.DeltaHue);

            bool incompatibility = (analysis.DeltaHue != 0 && matizador == null);

            foreach (var ing in originalRecipe)
            {
                bool isMatizador = (matizador != null && ing.Code == matizador.Code);
                double fDH = isMatizador ? factorDH : 1.0;
                
                double newConc = ing.Percentage * factorDL * fDH;
                string status = "OK";

                // Alerta de Saturación (Crítica)
                if (newConc > 4.5) status = "SATURACIÓN";

                result.Ingredients.Add(new CorrectiveIngredientDetail
                {
                    Code = ing.Code,
                    Name = ing.Name,
                    Original = ing.Percentage,
                    FactorDL = factorDL,
                    FactorDH = fDH,
                    NewConcentration = Math.Round(newConc, 5),
                    Status = status
                });
            }

            result.TotalNew = result.Ingredients.Sum(i => i.NewConcentration);

            // 3. Sistema de Alertas
            double combinedFactor = factorDL * factorDH;
            if (incompatibility)
            {
                result.AlertMessage = "Error: Incompatibilidad de matiz detectada";
                result.AlertSeverity = "Error";
            }
            else if (result.Ingredients.Any(i => i.Status == "SATURACIÓN"))
            {
                result.AlertMessage = "Límite de saturación excedido: Riesgo de solidez deficiente";
                result.AlertSeverity = "Critical";
            }
            else if (combinedFactor > 1.30)
            {
                result.AlertMessage = "Ajuste agresivo detectado: Verificar condiciones de máquina";
                result.AlertSeverity = "Warning";
            }
            else
            {
                result.AlertMessage = "Receta optimizada correctamente";
                result.AlertSeverity = "None";
            }

            return result;
        }

        private static RecipeIngredientInput IdentifyMatizador(List<RecipeIngredientInput> ingredients, double deltaHue)
        {
            if (Math.Abs(deltaHue) < 0.01) return null;

            // Mapeo simplificado de familias Coats
            // Si el desvío es positivo (Azulado en nuestra escala), necesitamos corregir con Rojo/Rubine
            // Si el desvío es negativo (Rojizo), necesitamos corregir con Amarillo o Azul según el eje.
            // Para esta propuesta, usaremos keywords:
            
            string[] keywords = null;
            if (deltaHue > 0) keywords = new[] { "RED", "RUBINE", "PINK", "BORDEAUX", "ROJO" };
            else keywords = new[] { "YELLOW", "GOLDEN", "ORANGE", "BLUE", "NAVY", "TURQUOISE", "AMARILLO", "AZUL" };

            foreach (var kw in keywords)
            {
                var found = ingredients.FirstOrDefault(i => i.Name.ToUpper().Contains(kw));
                if (found != null) return found;
            }

            return ingredients.OrderByDescending(i => i.Percentage).FirstOrDefault(); // Fallback al de mayor concentración
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

        public static string BuildConsolidatedLightnessTable(List<IlluminantCorrectionResult> results)
        {
            if (results == null || results.Count == 0) return "";
            var validResults = results.Where(r => r.Ingredients != null && r.Ingredients.Count > 0).ToList();
            if (validResults.Count == 0) return "";

            var sb = new System.Text.StringBuilder();
            sb.AppendLine();
            sb.AppendLine("════════════════════════════════════════════════════════════════════════════════════════");
            sb.AppendLine("  CÁLCULOS REALIZADOS (LIGHTNESS)");
            sb.AppendLine("════════════════════════════════════════════════════════════════════════════════════════");

            string formatStr = "  {0,-28} |";
            var args = new List<object> { "Ingrediente" };
            
            for (int i = 0; i < validResults.Count; i++)
            {
                formatStr += $" {{{i + 1},14}} |";
                string n = validResults[i].Illuminant.ToUpper();
                args.Add(i == 0 ? $"Lightness({n})" : $"Lightn({n})");
            }
            formatStr += $" {{{validResults.Count + 1},14}}";
            args.Add("Promedio L.");

            sb.AppendLine(string.Format(formatStr, args.ToArray()));
            int dashCount = 35 + validResults.Count * 17 + 14; 
            if (dashCount < 92) dashCount = 92;
            sb.AppendLine("  " + new string('─', dashCount));

            int ingCount = validResults[0].Ingredients.Count;
            for (int i = 0; i < ingCount; i++)
            {
                var rowArgs = new List<object> { validResults[0].Ingredients[i].Name };
                var rowFormat = "  {0,-28} |";
                
                double sumVal = 0;
                for (int j = 0; j < validResults.Count; j++)
                {
                    double val = validResults[j].Ingredients[i].Calc2;
                    double orig = validResults[j].Ingredients[i].Original;
                    double totalOrig = validResults[j].TotalOriginal;
                    
                    sumVal += val;
                    rowFormat += $" {{{j + 1},14}} |";
                    
                    if (j == 0)
                        rowArgs.Add(string.Format(CultureInfo.InvariantCulture, "«{0,6:F5} | {1,4:P0}»", val, orig / totalOrig));
                    else
                        rowArgs.Add(string.Format(CultureInfo.InvariantCulture, "{0,6:F5} | {1,4:P0}", val, orig / totalOrig));
                }
                
                double avg = sumVal / validResults.Count;
                rowFormat += $" {{{validResults.Count + 1},14}}";
                rowArgs.Add(string.Format(CultureInfo.InvariantCulture, "{0,6:P0}", avg));

                sb.AppendLine(string.Format(rowFormat, rowArgs.ToArray()));
                
                if (i < ingCount - 1)
                     sb.AppendLine("  " + new string('┈', dashCount));
            }

            sb.AppendLine("  " + new string('─', dashCount));

            // Fila de Totales
            var totArgs = new List<object> { "  [Total]" };
            var totFormat = "  {0,-28} |";
            
            double sumTotalAvg = 0;
            for (int j = 0; j < validResults.Count; j++)
            {
                double sumCalc2 = validResults[j].SumCalc2;
                sumTotalAvg += sumCalc2;
                totFormat += $" {{{j + 1},14}} |";
                if (j == 0)
                    totArgs.Add(string.Format(CultureInfo.InvariantCulture, "«{0,6:F5} | 100%»", sumCalc2));
                else
                    totArgs.Add(string.Format(CultureInfo.InvariantCulture, "{0,6:F5} | 100%", sumCalc2));
            }
            double totalAvg = sumTotalAvg / validResults.Count;
            totFormat += $" {{{validResults.Count + 1},14}}";
            totArgs.Add(string.Format(CultureInfo.InvariantCulture, "{0,6:P0}", totalAvg));

            sb.AppendLine(string.Format(totFormat, totArgs.ToArray()));

            // Fila de Variaciones
            var varArgs = new List<object> { "  Variación (Ligthness %)" };
            var varFormat = "  {0,-28} |";
            for (int j = 0; j < validResults.Count; j++)
            {
                varFormat += $" {{{j + 1},14}} |";
                double resL = validResults[j].ResultadoLightness * 100.0;
                if (j == 0)
                    varArgs.Add(string.Format(CultureInfo.InvariantCulture, "«{0,8:F1}%»     ", resL));
                else
                    varArgs.Add(string.Format(CultureInfo.InvariantCulture, "{0,8:F1}%      ", resL));
            }
            varFormat += $" {{{validResults.Count + 1},14}}";
            varArgs.Add(""); 

            sb.AppendLine(string.Format(varFormat, varArgs.ToArray()));
            sb.AppendLine("  " + new string('─', dashCount));

            return sb.ToString();
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
                        tag + ing.Calc2.ToString("F5", CultureInfo.InvariantCulture) + endTag,
                        tag + ing.Calc3.ToString("F5", CultureInfo.InvariantCulture) + endTag));
                }

                sb.AppendLine("  " + new string('─', 84));
                
                sb.AppendLine(string.Format("  {0,-28} {1,14:F5} {2,12} {3,14} {4,12}",
                    "  [Total]", r.TotalOriginal, "100%", 
                    tag + r.SumCalc2.ToString("F5", CultureInfo.InvariantCulture) + endTag,
                    tag + r.SumCalc3.ToString("F5", CultureInfo.InvariantCulture) + endTag));

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
            var validResults = results.Where(r => r.Ingredients != null && r.Ingredients.Count > 0).ToList();
            if (validResults.Count == 0) return "";

            var sb = new System.Text.StringBuilder();
            sb.AppendLine();
            sb.AppendLine("════════════════════════════════════════════════════════════════════════════════════════");
            sb.AppendLine("  CÁLCULOS REALIZADOS (CHROMA)");
            sb.AppendLine("════════════════════════════════════════════════════════════════════════════════════════");
            
            string formatStr = "  {0,-28} |";
            var args = new List<object> { "Ingrediente" };
            
            for (int i = 0; i < validResults.Count; i++)
            {
                formatStr += $" {{{i + 1},14}} |";
                string n = validResults[i].Illuminant.ToUpper();
                args.Add($"Chroma ({n})");
            }
            formatStr += $" {{{validResults.Count + 1},14}}";
            args.Add("Promedio Ch.");

            sb.AppendLine(string.Format(formatStr, args.ToArray()));
            int dashCount = 35 + validResults.Count * 17 + 14; 
            if (dashCount < 92) dashCount = 92;
            sb.AppendLine("  " + new string('─', dashCount));

            int ingCount = validResults[0].Ingredients.Count;
            double globalSumAvg = 0;
            
            for (int i = 0; i < ingCount; i++)
            {
                double sumTemp = 0;
                for (int j = 0; j < validResults.Count; j++)
                    sumTemp += validResults[j].Ingredients[i].Calc3;
                globalSumAvg += sumTemp / validResults.Count;
            }

            for (int i = 0; i < ingCount; i++)
            {
                var rowArgs = new List<object> { validResults[0].Ingredients[i].Name };
                var rowFormat = "  {0,-28} |";
                
                double sumVal = 0;
                for (int j = 0; j < validResults.Count; j++)
                {
                    double val = validResults[j].Ingredients[i].Calc3;
                    double sumCalc3 = validResults[j].SumCalc3;
                    
                    sumVal += val;
                    rowFormat += $" {{{j + 1},14}} |";
                    
                    if (j == 0)
                        rowArgs.Add(string.Format(CultureInfo.InvariantCulture, "«{0,6:F5} | {1,4:P0}»", val, val / sumCalc3));
                    else
                        rowArgs.Add(string.Format(CultureInfo.InvariantCulture, "{0,6:F5} | {1,4:P0}", val, val / sumCalc3));
                }
                
                double avg = sumVal / validResults.Count;
                rowFormat += $" {{{validResults.Count + 1},14}}";
                rowArgs.Add(string.Format(CultureInfo.InvariantCulture, "{0,6:F3} | {1,4:P0}", avg, avg / globalSumAvg));

                sb.AppendLine(string.Format(rowFormat, rowArgs.ToArray()));
                
                if (i < ingCount - 1)
                     sb.AppendLine("  " + new string('┈', dashCount));
            }

            sb.AppendLine("  " + new string('─', dashCount));

            // Fila de Totales
            var totArgs = new List<object> { "  [Total]" };
            var totFormat = "  {0,-28} |";
            
            for (int j = 0; j < validResults.Count; j++)
            {
                double sumCalc3 = validResults[j].SumCalc3;
                totFormat += $" {{{j + 1},14}} |";
                if (j == 0)
                    totArgs.Add(string.Format(CultureInfo.InvariantCulture, "«{0,6:F5} | 100%»", sumCalc3));
                else
                    totArgs.Add(string.Format(CultureInfo.InvariantCulture, "{0,6:F5} | 100%", sumCalc3));
            }
            totFormat += $" {{{validResults.Count + 1},14}}";
            totArgs.Add(string.Format(CultureInfo.InvariantCulture, "{0,6:F5} | 100%", globalSumAvg));

            sb.AppendLine(string.Format(totFormat, totArgs.ToArray()));

            // Fila de Variaciones
            var varArgs = new List<object> { "  Variación (Croma %)" };
            var varFormat = "  {0,-28} |";
            double sumResC = 0;
            for (int j = 0; j < validResults.Count; j++)
            {
                varFormat += $" {{{j + 1},14}} |";
                double resC = validResults[j].ResultadoChroma * 100.0;
                sumResC += resC;
                if (j == 0)
                    varArgs.Add(string.Format(CultureInfo.InvariantCulture, "«{0,8:F1}%»     ", resC));
                else
                    varArgs.Add(string.Format(CultureInfo.InvariantCulture, "{0,8:F1}%      ", resC));
            }
            double varAvg = sumResC / validResults.Count;
            varFormat += $" {{{validResults.Count + 1},14}}";
            varArgs.Add(string.Format(CultureInfo.InvariantCulture, "{0,8:F1}%      ", varAvg)); 

            sb.AppendLine(string.Format(varFormat, varArgs.ToArray()));
            sb.AppendLine("  " + new string('─', dashCount));

            return sb.ToString();
        }
    }
}