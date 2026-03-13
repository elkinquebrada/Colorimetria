// OcrTextNormalizer.cs — C# 7.3
using System;
using System.Text.RegularExpressions;

namespace Colorimetria
{
    /// <summary>
    /// Normalización compartida para texto OCR (caracteres Unicode, comas/puntos, correcciones típicas).
    /// </summary>
    public static class OcrTextNormalizer
    {
        /// <summary>
        /// Normaliza un bloque grande (aplica por líneas).
        /// </summary>
        public static string NormalizeBlock(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            string[] lines = text.Split('\n');
            for (int i = 0; i < lines.Length; i++)
                lines[i] = NormalizeLine(lines[i]);
            return string.Join("\n", lines);
        }

        /// <summary>
        /// Normaliza una línea (mayúsculas conservadoras para etiquetas, arregla Unicode y confusiones frecuentes).
        /// </summary>
        public static string NormalizeLine(string s)
        {
            if (string.IsNullOrEmpty(s)) return string.Empty;

            // Unicode y puntuación
            string t = s
                .Replace('\u2212', '-') // minus unicode → '-'
                .Replace('\u2013', '-') // en dash
                .Replace('\u2014', '-') // em dash
                .Replace('\u2018', '\'')
                .Replace('\u2019', '\'')
                .Replace('\u201C', '\"')
                .Replace('\u201D', '\"');

            // OCR a veces usa coma decimal; mantenemos coma/punto para regex tolerantes,
            // y durante el parse convertiremos coma→punto.
            // Correcciones de etiquetas y acrónimos típicos en tus reportes:
            t = Regex.Replace(t, @"\b5TD\b", "STD", RegexOptions.IgnoreCase);
            t = Regex.Replace(t, @"\bTLS4\b", "TL84", RegexOptions.IgnoreCase);
            t = Regex.Replace(t, @"\bTLS3\b", "TL83", RegexOptions.IgnoreCase);
            t = Regex.Replace(t, @"\bTLS5\b", "TL85", RegexOptions.IgnoreCase);
            t = Regex.Replace(t, @"\bDO5\b", "D65", RegexOptions.IgnoreCase);
            t = Regex.Replace(t, @"\bD6O\b", "D65", RegexOptions.IgnoreCase);
            t = Regex.Replace(t, @"\bDSO\b", "D50", RegexOptions.IgnoreCase);
            t = Regex.Replace(t, @"\bD5O\b", "D50", RegexOptions.IgnoreCase);
            t = Regex.Replace(t, @"\bC\s*W\s*F\b", "CWF", RegexOptions.IgnoreCase);
            t = Regex.Replace(t, @"\bCW/F\b", "CWF", RegexOptions.IgnoreCase);
            t = Regex.Replace(t, @"\bF\s*1\s*1\b", "F11", RegexOptions.IgnoreCase);
            t = Regex.Replace(t, @"\bF\s*1\s*2\b", "F12", RegexOptions.IgnoreCase);
            t = Regex.Replace(t, @"\[\s*LLUMINANT\b", "ILLUMINANT", RegexOptions.IgnoreCase);
            t = Regex.Replace(t, @"\bL0T\b", "LOT", RegexOptions.IgnoreCase);

            // Limpieza conservadora (deja letras, dígitos, puntuación básica, símbolos frecuentes y espacios)
            t = Regex.Replace(t, @"[^\p{L}0-9\.\,\-\%\(\)\/\:_\s]", " ");
            // Compactar espacios múltiples
            t = Regex.Replace(t, @"[ \t]+", " ").Trim();

            return t;
        }

        /// <summary>
        /// Convierte comas decimales a punto (antes de double.Parse InvariantCulture).
        /// </summary>
        public static string FixDecimalSeparator(string token)
        {
            return string.IsNullOrEmpty(token) ? token : token.Replace(',', '.').Trim();
        }

        /// <summary>
        /// Inserta un punto decimal si falta y el rango objetivo es [-99,99] (para L*, a*, b*).
        /// "3224" → "32.24", "8283" → "82.83", "-1424" → "-14.24"
        /// </summary>
        public static double RestoreMeasureDecimal(string token)
        {
            if (string.IsNullOrWhiteSpace(token)) return 0;
            token = FixDecimalSeparator(token);

            // Verificar si ya tiene punto decimal
            bool hasPoint = token.Contains(".");
            double v;

            if (hasPoint)
            {
                // Ya tiene punto: devolver tal cual (confiar en el OCR)
                if (double.TryParse(token, System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out v))
                    return v;
                return 0;
            }

            // Sin punto decimal: insertar a 2 dígitos del final (8283 → 82.83)
            bool neg = token.StartsWith("-");
            string d = neg ? token.Substring(1) : token;
            if (!Regex.IsMatch(d, @"^\d+$")) return 0;
            if (d.Length <= 2) return SafeParse(token);
            string rebuilt = (neg ? "-" : "") + d.Substring(0, d.Length - 2) + "." + d.Substring(d.Length - 2);
            double val2;
            return double.TryParse(rebuilt, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out val2) ? val2 : 0;
        }

        /// <summary>
        /// Inserta punto decimal para deltas pequeños [-9.99, 9.99] si falta. "118" → "1.18"
        /// </summary>
        public static double RestoreDeltaDecimal(string token)
        {
            if (string.IsNullOrWhiteSpace(token)) return 0;
            token = FixDecimalSeparator(token);
            double v;
            if (double.TryParse(token, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out v))
            {
                if (v >= -9.99 && v <= 9.99) return v;
            }
            bool neg = token.StartsWith("-");
            string d = neg ? token.Substring(1) : token;
            if (!Regex.IsMatch(d, @"^\d+$")) return 0;
            string sign = neg ? "-" : "";
            if (d.Length == 1) return SafeParse(sign + "0." + d);
            if (d.Length == 2) return SafeParse(sign + d[0] + "." + d[1]);
            return SafeParse(sign + d.Substring(0, d.Length - 2) + "." + d.Substring(d.Length - 2));
        }

        public static double SafeParse(string s)
        {
            double v;
            return double.TryParse(FixDecimalSeparator(s), System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out v) ? v : 0d;
        }
    }
}