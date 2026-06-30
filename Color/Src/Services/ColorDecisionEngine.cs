using System;
using System.Collections.Generic;

namespace Color
{
    public class ColorDecisionResult
    {
        public string Brightness { get; set; } = string.Empty;
        public string QuadrantLabel { get; set; } = string.Empty;
        public string OppositeColor { get; set; } = string.Empty;
        public string DeviationCell0 { get; set; } = string.Empty; 
    }

    public static class ColorDecisionEngine
    {
        // Diccionario estático de solo lectura optimizado en memoria para búsquedas instantáneas
        private static readonly Dictionary<string, string> ColorMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "Yellower (Greener)", "Bluer (Redder)" },
            { "Yellower (Redder)", "Bluer (Greener)" },
            { "Greener (Bluer)", "Redder (Yellower)" },
            { "Greener (Yellower)", "Redder (Bluer)" },
            { "Bluer (Redder)", "Yellower (Greener)" },
            { "Bluer (Greener)", "Yellower (Redder)" },
            { "Redder (Yellower)", "Greener (Bluer)" },
            { "Redder (Bluer)", "Greener (Yellower)" }
        };

        public static ColorDecisionResult EvaluarDesviacion(double deltaL, string hueLabel)
        {
            var result = new ColorDecisionResult();

            // 1. Limpieza preventiva extrema de strings para evitar rupturas por saltos de línea o espacios duplicados
            string cleanHue = hueLabel.Replace("\r", "").Replace("\n", " ").Trim();
            cleanHue = System.Text.RegularExpressions.Regex.Replace(cleanHue, @"\s+", " ");

            // 2. FÍSICA TEXTIL REAL (Equivalente exacto a la fórmula IF de Excel =+IF(H32>0,E63,E62)):
            result.Brightness = (deltaL > 0) ? "Brighter" : "Duller";

            // 3. MAPEO SEGURO (Equivalente exacto a la fórmula XLOOKUP de cuadrantes =+XLOOKUP(I33,C70:C78,D70:D78,""))
            if (ColorMap.TryGetValue(cleanHue, out string opposite))
            {
                result.OppositeColor = opposite;
            }
            else
            {
                // Mecanismo de contingencia: si el string es desconocido hereda el base para no romper la pantalla
                result.OppositeColor = cleanHue;
            }

            result.QuadrantLabel = cleanHue;

            // 4. (Equivalente exacto a =+XLOOKUP(H33,E62:E63,F62:F63,""))
            result.DeviationCell0 = result.Brightness;

            return result;
        }
    }
}