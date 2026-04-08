using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Globalization;

namespace Color
{
    /// Cliente para la API REST de corrección colorimétrica.
    public class ColorimetriaApiClient
    {
        private readonly string _apiBaseUrl;

        // Usa localhost:5000 si la API corre en la misma máquina
        public ColorimetriaApiClient(string apiBaseUrl = "http://localhost:5000")
        {
            _apiBaseUrl = apiBaseUrl.TrimEnd('/');
        }

        // ── Healthcheck ────────────────────────────────────────────────────────

        public bool IsApiAvailable()
        {
            try
            {
                var req = (HttpWebRequest)WebRequest.Create(_apiBaseUrl + "/api/colorimetry/health");
                req.Method = "GET";
                req.Timeout = 3000;
                using (var resp = (HttpWebResponse)req.GetResponse())
                    return resp.StatusCode == HttpStatusCode.OK;
            }
            catch { return false; }
        }

        // ── Corrección principal ───────────────────────────────────────────────
        public List<ApiCorrectionResult> CorrectReport(OcrReport report, double chromaThreshold = 0.35)
        {
            var results = new List<ApiCorrectionResult>();
            if (report?.Measures == null || report.Measures.Count == 0) return results;

            try
            {
                // Construir JSON del request
                string body = BuildRequestJson(report, chromaThreshold);
                string responseJson = PostJson(_apiBaseUrl + "/api/colorimetry/correct", body);
                if (string.IsNullOrEmpty(responseJson)) return results;

                // Parsear correcciones aceptadas y aplicar al reporte
                results = ParseCorrections(responseJson);
                foreach (var c in results)
                    ApplyToReport(report, c);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[ApiClient] Error: " + ex.Message);
            }

            return results;
        }

        // ── Helpers ────────────────────────────────────────────────────────────

        private string BuildRequestJson(OcrReport report, double threshold)
        {
            var sb = new StringBuilder();
            sb.Append("{");
            sb.AppendFormat(CultureInfo.InvariantCulture, "\"chromaThreshold\":{0},", threshold);
            sb.Append("\"measures\":[");

            for (int i = 0; i < report.Measures.Count; i++)
            {
                var r = report.Measures[i];
                if (i > 0) sb.Append(",");
                sb.AppendFormat(CultureInfo.InvariantCulture,
                    "{{\"illuminant\":\"{0}\",\"type\":\"{1}\"," +
                    "\"l\":{2},\"a\":{3},\"b\":{4},\"chroma\":{5},\"hue\":{6},\"needsReview\":{7}}}",
                    r.Illuminant, r.Type,
                    r.L, r.A, r.B, r.Chroma, r.Hue,
                    r.NeedsReview ? "true" : "false");
            }

            sb.Append("]}");
            return sb.ToString();
        }

        private string PostJson(string url, string body)
        {
            try
            {
                byte[] bodyBytes = Encoding.UTF8.GetBytes(body);
                var req = (HttpWebRequest)WebRequest.Create(url);
                req.Method = "POST";
                req.ContentType = "application/json";
                req.ContentLength = bodyBytes.Length;
                req.Timeout = 20000;

                using (var stream = req.GetRequestStream())
                    stream.Write(bodyBytes, 0, bodyBytes.Length);

                using (var resp = (HttpWebResponse)req.GetResponse())
                using (var reader = new StreamReader(resp.GetResponseStream(), Encoding.UTF8))
                    return reader.ReadToEnd();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[ApiClient] HTTP error: " + ex.Message);
                return null;
            }
        }

        private List<ApiCorrectionResult> ParseCorrections(string json)
        {
            var list = new List<ApiCorrectionResult>();
            // Extraer array "corrections"
            var corrMatch = Regex.Match(json, "\"corrections\"\\s*:\\s*(\\[[^\\]]*\\])", RegexOptions.Singleline);
            if (!corrMatch.Success) return list;

            var objMatches = Regex.Matches(corrMatch.Groups[1].Value, @"\{[^}]+\}");
            foreach (Match m in objMatches)
            {
                string obj = m.Value;
                if (!ExtractBool(obj, "accepted")) continue;

                list.Add(new ApiCorrectionResult
                {
                    Field = ExtractString(obj, "field"),
                    Illuminant = ExtractString(obj, "illuminant"),
                    Type = ExtractString(obj, "type"),
                    CorrectedValue = ExtractDouble(obj, "correctedValue"),
                    Reason = ExtractString(obj, "reason")
                });
            }
            return list;
        }

        private void ApplyToReport(OcrReport report, ApiCorrectionResult c)
        {
            foreach (var row in report.Measures)
            {
                if (!string.Equals(row.Illuminant, c.Illuminant, StringComparison.OrdinalIgnoreCase)) continue;
                if (!string.Equals(row.Type, c.Type, StringComparison.OrdinalIgnoreCase)) continue;

                switch (c.Field)
                {
                    case "a": row.A = c.CorrectedValue; break;
                    case "b": row.B = c.CorrectedValue; break;
                    case "L": row.L = c.CorrectedValue; break;
                    case "Chroma": row.Chroma = c.CorrectedValue; break;
                    case "Hue": row.Hue = c.CorrectedValue; break;
                }

                if (c.Field == "a" || c.Field == "b")
                {
                    double newHue = Math.Atan2(row.B, row.A) * 180.0 / Math.PI;
                    if (newHue < 0) newHue += 360.0;
                    row.Hue = Math.Round(newHue);
                }
                row.NeedsReview = false;
                break;
            }
        }

        // ── Regex helpers (sin Newtonsoft, compatible .NET 4.8) ────────────────
        private string ExtractString(string obj, string key) =>
            Regex.Match(obj, "\"" + key + "\"\\s*:\\s*\"([^\"]+)\"").Groups[1].Value;

        private double ExtractDouble(string obj, string key)
        {
            var m = Regex.Match(obj, "\"" + key + "\"\\s*:\\s*(-?\\d+(?:\\.\\d+)?)");
            return m.Success && double.TryParse(m.Groups[1].Value, NumberStyles.Float,
                CultureInfo.InvariantCulture, out double v) ? v : 0;
        }

        private bool ExtractBool(string obj, string key) =>
            Regex.Match(obj, "\"" + key + "\"\\s*:\\s*(true|false)").Groups[1].Value == "true";
    }

    public class ApiCorrectionResult
    {
        public string Field { get; set; }
        public string Illuminant { get; set; }
        public string Type { get; set; }
        public double CorrectedValue { get; set; }
        public string Reason { get; set; }
    }
}
