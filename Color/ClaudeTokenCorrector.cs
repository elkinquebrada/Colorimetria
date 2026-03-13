// ClaudeTokenCorrector.cs — C# 7.3 / .NET Framework 4.8
// Corrección de tokens numéricos erróneos via Claude API (Anthropic)
// PRIVACIDAD: Solo se envían números matemáticos — NUNCA imagen, nombre, lote ni receta.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace Color
{
    /// <summary>
    /// Token numérico con error de coherencia colorimétrica detectado localmente.
    /// Solo contiene números — sin datos industriales sensibles.
    /// </summary>
    public class ErrorToken
    {
        /// <summary>Campo: "L", "a", "b", "Chroma", "Hue"</summary>
        public string Field { get; set; }

        /// <summary>Iluminante: "D65", "TL84", "A" — solo para contexto matemático</summary>
        public string Illuminant { get; set; }

        /// <summary>Tipo: "Std" o "Lot"</summary>
        public string Type { get; set; }

        /// <summary>Valor que entregó el OCR (posiblemente erróneo)</summary>
        public double OcrValue { get; set; }

        /// <summary>Token string original del OCR</summary>
        public string OcrToken { get; set; }

        /// <summary>Valor de a* ya corregido (contexto para calcular coherencia)</summary>
        public double A { get; set; }

        /// <summary>Valor de b* ya corregido (contexto para calcular coherencia)</summary>
        public double B { get; set; }

        /// <summary>Chroma reportada por OCR</summary>
        public double Chroma { get; set; }

        /// <summary>Error de coherencia: |sqrt(a²+b²) - Chroma|</summary>
        public double CoherenceError { get; set; }
    }

    /// <summary>
    /// Resultado de corrección devuelto por Claude API.
    /// </summary>
    public class ClaudeCorrectionResult
    {
        public string Field { get; set; }
        public string Illuminant { get; set; }
        public string Type { get; set; }
        public double CorrectedValue { get; set; }
        public double NewCoherenceError { get; set; }
        public bool Accepted { get; set; }
        public string Reason { get; set; }
    }

    /// <summary>
    /// Cliente Claude API para corrección de tokens colorimétricos erróneos.
    /// Envía SOLO números — sin imagen ni datos industriales.
    /// </summary>
    public class ClaudeTokenCorrector
    {
        private readonly string _apiKey;
        private const string API_URL = "https://api.anthropic.com/v1/messages";
        private const string MODEL = "claude-haiku-4-5-20251001"; // Más rápido y económico
        private const int MAX_TOKENS = 512;
        private const double IMPROVEMENT_THRESHOLD = 0.3; // Solo aceptar si mejora >0.3

        // ── Control de uso ─────────────────────────────────────────────
        private int _callCount = 0;
        private const int MAX_CALLS_PER_SESSION = 20; // Límite de seguridad por sesión

        public int CallCount => _callCount;
        public bool IsEnabled { get; set; } = true;

        public ClaudeTokenCorrector(string apiKey)
        {
            if (string.IsNullOrWhiteSpace(apiKey))
                throw new ArgumentException("API key requerida", "apiKey");
            _apiKey = apiKey.Trim();
        }

        // ── API pública ─────────────────────────────────────────────────

        /// <summary>
        /// Detecta tokens erróneos en el reporte y solicita corrección a Claude.
        /// Solo se procesan filas donde sqrt(a²+b²) difiere de Chroma más del umbral.
        /// </summary>
        public List<ClaudeCorrectionResult> CorrectReport(OcrReport report, double chromaThreshold = 0.35)
        {
            var results = new List<ClaudeCorrectionResult>();
            if (!IsEnabled || report == null || report.Measures == null) return results;

            var errorTokens = DetectErrorTokens(report.Measures, chromaThreshold);
            if (errorTokens.Count == 0) return results;

            // Agrupar en un solo llamado para minimizar requests
            var corrections = RequestCorrections(errorTokens);
            if (corrections == null) return results;

            // Aplicar correcciones aceptadas al reporte
            foreach (var correction in corrections)
            {
                if (!correction.Accepted) continue;
                ApplyCorrection(report, correction);
                results.Add(correction);
            }

            return results;
        }

        // ── Detección de tokens erróneos ────────────────────────────────

        private List<ErrorToken> DetectErrorTokens(List<ColorimetricRow> rows, double threshold)
        {
            var tokens = new List<ErrorToken>();

            foreach (var row in rows)
            {
                double chromaCalc = Math.Sqrt(row.A * row.A + row.B * row.B);
                double chromaErr = Math.Abs(chromaCalc - row.Chroma);

                if (chromaErr <= threshold) continue; // Coherente — no necesita corrección

                // b* es el más frecuentemente erróneo
                tokens.Add(new ErrorToken
                {
                    Field = "b",
                    Illuminant = row.Illuminant,
                    Type = row.Type,
                    OcrValue = row.B,
                    OcrToken = row.B.ToString("G6", CultureInfo.InvariantCulture),
                    A = row.A,
                    B = row.B,
                    Chroma = row.Chroma,
                    CoherenceError = chromaErr
                });

                // a* también puede ser erróneo
                if (chromaErr > 2.0)
                {
                    tokens.Add(new ErrorToken
                    {
                        Field = "a",
                        Illuminant = row.Illuminant,
                        Type = row.Type,
                        OcrValue = row.A,
                        OcrToken = row.A.ToString("G6", CultureInfo.InvariantCulture),
                        A = row.A,
                        B = row.B,
                        Chroma = row.Chroma,
                        CoherenceError = chromaErr
                    });
                }

                // Hue — verificar coherencia con a*/b*
                double hueCalc = Math.Atan2(row.B, row.A) * 180.0 / Math.PI;
                if (hueCalc < 0) hueCalc += 360.0;
                double hueErr = Math.Abs(row.Hue - hueCalc);
                if (hueErr > 180) hueErr = 360 - hueErr;
                if (hueErr > 5.0)
                {
                    tokens.Add(new ErrorToken
                    {
                        Field = "Hue",
                        Illuminant = row.Illuminant,
                        Type = row.Type,
                        OcrValue = row.Hue,
                        OcrToken = row.Hue.ToString("G6", CultureInfo.InvariantCulture),
                        A = row.A,
                        B = row.B,
                        Chroma = row.Chroma,
                        CoherenceError = hueErr
                    });
                }
            }

            return tokens;
        }

        // ── Llamada a Claude API ────────────────────────────────────────

        private List<ClaudeCorrectionResult> RequestCorrections(List<ErrorToken> tokens)
        {
            if (_callCount >= MAX_CALLS_PER_SESSION)
            {
                LogWarning("Límite de sesión alcanzado (" + MAX_CALLS_PER_SESSION + " calls)");
                return null;
            }

            string prompt = BuildPrompt(tokens);
            string response = CallClaudeApi(prompt);
            if (string.IsNullOrEmpty(response)) return null;

            return ParseResponse(response, tokens);
        }

        private string BuildPrompt(List<ErrorToken> tokens)
        {
            // Construir prompt con SOLO números — sin datos industriales
            var sb = new StringBuilder();
            sb.AppendLine("You are a colorimetry expert. Correct OCR errors in CIELab measurements.");
            sb.AppendLine("Rules: Chroma = sqrt(a² + b²). Hue = atan2(b,a) * 180/PI (0-360°).");
            sb.AppendLine("OCR often misplaces decimal points: -36.0 should be -3.60, -50.6 should be -5.06.");
            sb.AppendLine();
            sb.AppendLine("For each token, find the corrected value that minimizes |sqrt(a²+b²) - Chroma|.");
            sb.AppendLine("Respond ONLY with JSON array, no explanation:");
            sb.AppendLine("[{\"idx\":0,\"corrected\":VALUE,\"reason\":\"brief\"}]");
            sb.AppendLine();
            sb.AppendLine("Tokens to correct:");

            for (int i = 0; i < tokens.Count; i++)
            {
                var t = tokens[i];
                if (t.Field == "Hue")
                    sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                        "[{0}] field={1} illuminant={2} type={3} ocr_value={4} a={5} b={6} hue_error={7:F1}deg",
                        i, t.Field, t.Illuminant, t.Type, t.OcrValue, t.A, t.B, t.CoherenceError));
                else
                    sb.AppendLine(string.Format(CultureInfo.InvariantCulture,
                        "[{0}] field={1} illuminant={2} type={3} ocr_value={4} a={5} b={6} chroma={7} coherence_error={8:F3}",
                        i, t.Field, t.Illuminant, t.Type, t.OcrValue, t.A, t.B, t.Chroma, t.CoherenceError));
            }

            return sb.ToString();
        }

        private string CallClaudeApi(string prompt)
        {
            try
            {
                // Construir JSON manualmente (C# 7.3 sin System.Text.Json)
                string escapedPrompt = prompt
                    .Replace("\\", "\\\\")
                    .Replace("\"", "\\\"")
                    .Replace("\r\n", "\\n")
                    .Replace("\n", "\\n")
                    .Replace("\r", "\\n");

                string body = string.Format(
                    "{{\"model\":\"{0}\",\"max_tokens\":{1},\"messages\":[{{\"role\":\"user\",\"content\":\"{2}\"}}]}}",
                    MODEL, MAX_TOKENS, escapedPrompt);

                byte[] bodyBytes = Encoding.UTF8.GetBytes(body);

                var request = (HttpWebRequest)WebRequest.Create(API_URL);
                request.Method = "POST";
                request.ContentType = "application/json";
                request.ContentLength = bodyBytes.Length;
                request.Timeout = 15000; // 15 segundos
                request.Headers.Add("x-api-key", _apiKey);
                request.Headers.Add("anthropic-version", "2023-06-01");

                using (var stream = request.GetRequestStream())
                    stream.Write(bodyBytes, 0, bodyBytes.Length);

                using (var response = (HttpWebResponse)request.GetResponse())
                using (var reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8))
                {
                    _callCount++;
                    return reader.ReadToEnd();
                }
            }
            catch (WebException ex)
            {
                string errorBody = "";
                if (ex.Response != null)
                    using (var r = new StreamReader(ex.Response.GetResponseStream()))
                        errorBody = r.ReadToEnd();
                LogWarning("Claude API error: " + ex.Message + " | " + errorBody);
                return null;
            }
            catch (Exception ex)
            {
                LogWarning("Claude API exception: " + ex.Message);
                return null;
            }
        }

        // ── Parseo de respuesta JSON ────────────────────────────────────

        private List<ClaudeCorrectionResult> ParseResponse(string jsonResponse, List<ErrorToken> tokens)
        {
            var results = new List<ClaudeCorrectionResult>();

            try
            {
                // Extraer el texto de la respuesta de Claude
                // Buscar "text":"..." en el JSON de respuesta
                var textMatch = Regex.Match(jsonResponse,
                    "\"text\"\\s*:\\s*\"((?:[^\"\\\\]|\\\\.)*)\"",
                    RegexOptions.Singleline);

                string claudeText = textMatch.Success
                    ? Regex.Unescape(textMatch.Groups[1].Value)
                    : jsonResponse;

                // Extraer array JSON de la respuesta de Claude
                var arrayMatch = Regex.Match(claudeText,
                    @"\[[\s\S]*?\]",
                    RegexOptions.Singleline);

                if (!arrayMatch.Success) return results;

                string jsonArray = arrayMatch.Value;

                // Parsear cada objeto del array manualmente (sin Newtonsoft ni System.Text.Json)
                var objMatches = Regex.Matches(jsonArray, @"\{[^}]+\}");
                foreach (Match objMatch in objMatches)
                {
                    string obj = objMatch.Value;

                    // Extraer idx
                    var idxMatch = Regex.Match(obj, "\"idx\"\\s*:\\s*(\\d+)");
                    if (!idxMatch.Success) continue;
                    int idx;
                    if (!int.TryParse(idxMatch.Groups[1].Value, out idx)) continue;
                    if (idx < 0 || idx >= tokens.Count) continue;

                    // Extraer corrected value
                    var valMatch = Regex.Match(obj,
                        "\"corrected\"\\s*:\\s*(-?\\d+(?:\\.\\d+)?)");
                    if (!valMatch.Success) continue;
                    double corrected;
                    if (!double.TryParse(valMatch.Groups[1].Value,
                        NumberStyles.Float, CultureInfo.InvariantCulture, out corrected))
                        continue;

                    // Extraer reason
                    var reasonMatch = Regex.Match(obj, "\"reason\"\\s*:\\s*\"([^\"]+)\"");
                    string reason = reasonMatch.Success ? reasonMatch.Groups[1].Value : "";

                    var token = tokens[idx];
                    bool accepted = EvaluateCorrection(token, corrected);

                    results.Add(new ClaudeCorrectionResult
                    {
                        Field = token.Field,
                        Illuminant = token.Illuminant,
                        Type = token.Type,
                        CorrectedValue = corrected,
                        Accepted = accepted,
                        Reason = reason,
                        NewCoherenceError = token.Field == "Hue"
                            ? 0
                            : Math.Abs(Math.Sqrt(
                                (token.Field == "a" ? corrected : token.A) *
                                (token.Field == "a" ? corrected : token.A) +
                                (token.Field == "b" ? corrected : token.B) *
                                (token.Field == "b" ? corrected : token.B)) - token.Chroma)
                    });
                }
            }
            catch (Exception ex)
            {
                LogWarning("ParseResponse error: " + ex.Message);
            }

            return results;
        }

        // ── Evaluación y aplicación ─────────────────────────────────────

        private bool EvaluateCorrection(ErrorToken token, double corrected)
        {
            // Validar rango físico
            if (token.Field == "L" && (corrected < 0 || corrected > 100)) return false;
            if ((token.Field == "a" || token.Field == "b") && Math.Abs(corrected) > 100) return false;
            if (token.Field == "Chroma" && (corrected < 0 || corrected > 200)) return false;
            if (token.Field == "Hue" && (corrected < 0 || corrected > 360)) return false;

            if (token.Field == "Hue")
            {
                double hueCalc = Math.Atan2(token.B, token.A) * 180.0 / Math.PI;
                if (hueCalc < 0) hueCalc += 360.0;
                double errNew = Math.Abs(corrected - hueCalc);
                if (errNew > 180) errNew = 360 - errNew;
                return errNew < token.CoherenceError - IMPROVEMENT_THRESHOLD;
            }

            // Para a* y b*: verificar que mejora la coherencia con Chroma
            double newA = token.Field == "a" ? corrected : token.A;
            double newB = token.Field == "b" ? corrected : token.B;
            double newErr = Math.Abs(Math.Sqrt(newA * newA + newB * newB) - token.Chroma);
            return newErr < token.CoherenceError - IMPROVEMENT_THRESHOLD;
        }

        private void ApplyCorrection(OcrReport report, ClaudeCorrectionResult correction)
        {
            foreach (var row in report.Measures)
            {
                if (!string.Equals(row.Illuminant, correction.Illuminant,
                    StringComparison.OrdinalIgnoreCase)) continue;
                if (!string.Equals(row.Type, correction.Type,
                    StringComparison.OrdinalIgnoreCase)) continue;

                switch (correction.Field)
                {
                    case "a": row.A = correction.CorrectedValue; break;
                    case "b": row.B = correction.CorrectedValue; break;
                    case "L": row.L = correction.CorrectedValue; break;
                    case "Chroma": row.Chroma = correction.CorrectedValue; break;
                    case "Hue": row.Hue = correction.CorrectedValue; break;
                }

                // Recalcular Hue desde a*/b* corregidos
                if (correction.Field == "a" || correction.Field == "b")
                {
                    double newHue = Math.Atan2(row.B, row.A) * 180.0 / Math.PI;
                    if (newHue < 0) newHue += 360.0;
                    row.Hue = Math.Round(newHue);
                }

                row.NeedsReview = false;
                break;
            }
        }

        private void LogWarning(string msg)
        {
            System.Diagnostics.Debug.WriteLine("[ClaudeCorrector] " + msg);
        }
    }
}