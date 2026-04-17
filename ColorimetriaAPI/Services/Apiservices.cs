using Color;
using ColorimetriaAPI.Models;

namespace ColorimetriaAPI.Services
{
    /// <summary>
    /// Servicio de corrección colorimétrica 100% local.
    /// No realiza ninguna llamada a internet ni a servicios externos.
    /// Toda la lógica es matemática pura basada en el espacio CIELab.
    /// </summary>
    public class ColorimetryService
    {
        private readonly ILogger<ColorimetryService> _logger;
        private const double IMPROVEMENT_THRESHOLD = 0.3;

        public ColorimetryService(ILogger<ColorimetryService> logger)
        {
            _logger = logger;
        }

        // ── API pública ────────────────────────────────────────────────────────

        public Task<CorrectReportResponse> CorrectReportAsync(OcrReportRequest request)
        {
            var response = new CorrectReportResponse { Success = false };

            if (request?.Measures == null || request.Measures.Count == 0)
            {
                response.Error = "No se recibieron medidas.";
                return Task.FromResult(response);
            }

            var measures = request.Measures.Select(r => new ColorimetricRow
            {
                Illuminant = r.Illuminant,
                Type = r.Type,
                L = r.L,
                A = r.A,
                B = r.B,
                Chroma = r.Chroma,
                Hue = r.Hue,
                NeedsReview = r.NeedsReview
            }).ToList();

            var errorTokens = DetectErrorTokens(measures, request.ChromaThreshold);
            response.TokensAnalyzed = errorTokens.Count;

            if (errorTokens.Count == 0)
            {
                response.Success = true;
                response.CorrectedMeasures = measures;
                return Task.FromResult(response);
            }

            foreach (var token in errorTokens)
            {
                var correction = CorrectTokenLocally(token);
                if (correction == null) continue;
                ApplyCorrection(measures, correction);
                response.Corrections.Add(correction);
                response.TokensCorrected++;
            }

            response.Success = true;
            response.CorrectedMeasures = measures;
            return Task.FromResult(response);
        }

        // ── Detección ─────────────────────────────────────────────────────────

        private List<ErrorToken> DetectErrorTokens(List<ColorimetricRow> rows, double threshold)
        {
            var tokens = new List<ErrorToken>();

            foreach (var row in rows)
            {
                double chromaCalc = Math.Sqrt(row.A * row.A + row.B * row.B);
                double chromaErr = Math.Abs(chromaCalc - row.Chroma);

                if (chromaErr <= threshold) continue;

                tokens.Add(new ErrorToken
                {
                    Field = "b",
                    Illuminant = row.Illuminant,
                    Type = row.Type,
                    OcrValue = row.B,
                    A = row.A,
                    B = row.B,
                    Chroma = row.Chroma,
                    CoherenceError = chromaErr
                });

                if (chromaErr > 2.0)
                    tokens.Add(new ErrorToken
                    {
                        Field = "a",
                        Illuminant = row.Illuminant,
                        Type = row.Type,
                        OcrValue = row.A,
                        A = row.A,
                        B = row.B,
                        Chroma = row.Chroma,
                        CoherenceError = chromaErr
                    });

                double hueCalc = Math.Atan2(row.B, row.A) * 180.0 / Math.PI;
                if (hueCalc < 0) hueCalc += 360.0;

                double hueErr = Math.Abs(row.Hue - hueCalc);
                if (hueErr > 180) hueErr = 360 - hueErr;

                if (hueErr > 5.0)
                    tokens.Add(new ErrorToken
                    {
                        Field = "Hue",
                        Illuminant = row.Illuminant,
                        Type = row.Type,
                        OcrValue = row.Hue,
                        A = row.A,
                        B = row.B,
                        Chroma = row.Chroma,
                        CoherenceError = hueErr
                    });
            }

            return tokens;
        }

        // ── Corrección matemática local ────────────────────────────────────────

        private CorrectionResult? CorrectTokenLocally(ErrorToken token)
        {
            var candidates = GenerateCandidates(token);
            CorrectionResult? best = null;
            double bestError = token.CoherenceError;

            foreach (var (candidate, reason) in candidates)
            {
                if (!IsPhysicallyValid(token.Field, candidate)) continue;

                double newError = ComputeNewError(token, candidate);

                if (newError < bestError - IMPROVEMENT_THRESHOLD)
                {
                    bestError = newError;
                    best = new CorrectionResult
                    {
                        Field = token.Field,
                        Illuminant = token.Illuminant,
                        Type = token.Type,
                        OriginalValue = token.OcrValue,
                        CorrectedValue = Math.Round(candidate, 4),
                        OriginalCoherenceError = token.CoherenceError,
                        NewCoherenceError = Math.Round(newError, 4),
                        Accepted = true,
                        Reason = reason
                    };
                }
            }

            if (best != null)
                _logger.LogInformation(
                    "[Corrección] {Field} {Illuminant}/{Type}: {Orig} → {Corr} ({Reason})",
                    best.Field, best.Illuminant, best.Type,
                    best.OriginalValue, best.CorrectedValue, best.Reason);

            return best;
        }

        private List<(double value, string reason)> GenerateCandidates(ErrorToken token)
        {
            double v = token.OcrValue;
            var list = new List<(double, string)>();

            if (token.Field == "Hue")
            {
                double hueCalc = Math.Atan2(token.B, token.A) * 180.0 / Math.PI;
                if (hueCalc < 0) hueCalc += 360.0;
                list.Add((Math.Round(hueCalc, 2), "Recalculado desde atan2(b,a)"));
                return list;
            }

            // Desplazamientos de punto decimal
            list.Add((v / 10.0, "Punto decimal ÷10"));
            list.Add((v / 100.0, "Punto decimal ÷100"));
            list.Add((v * 10.0, "Punto decimal ×10"));
            list.Add((v * 100.0, "Punto decimal ×100"));

            // Cambio de signo
            list.Add((-v, "Signo invertido"));
            list.Add((-v / 10.0, "Signo invertido + ÷10"));
            list.Add((-v / 100.0, "Signo invertido + ÷100"));
            list.Add((-v * 10.0, "Signo invertido + ×10"));

            // Resolución directa desde Chroma = sqrt(a²+b²)
            double chroma2 = token.Chroma * token.Chroma;

            if (token.Field == "b")
            {
                double a2 = token.A * token.A;
                if (chroma2 >= a2)
                {
                    double bSolved = Math.Sqrt(chroma2 - a2);
                    list.Add((bSolved, "Resuelto sqrt(Chroma²-a²)"));
                    list.Add((-bSolved, "Resuelto -sqrt(Chroma²-a²)"));
                }
            }
            else if (token.Field == "a")
            {
                double b2 = token.B * token.B;
                if (chroma2 >= b2)
                {
                    double aSolved = Math.Sqrt(chroma2 - b2);
                    list.Add((aSolved, "Resuelto sqrt(Chroma²-b²)"));
                    list.Add((-aSolved, "Resuelto -sqrt(Chroma²-b²)"));
                }
            }

            return list;
        }

        private double ComputeNewError(ErrorToken token, double candidate)
        {
            if (token.Field == "Hue")
            {
                double hueCalc = Math.Atan2(token.B, token.A) * 180.0 / Math.PI;
                if (hueCalc < 0) hueCalc += 360.0;
                double err = Math.Abs(candidate - hueCalc);
                return err > 180 ? 360 - err : err;
            }

            double newA = token.Field == "a" ? candidate : token.A;
            double newB = token.Field == "b" ? candidate : token.B;
            return Math.Abs(Math.Sqrt(newA * newA + newB * newB) - token.Chroma);
        }

        private static bool IsPhysicallyValid(string field, double value) => field switch
        {
            "L" => value is >= 0 and <= 100,
            "a" or "b" => Math.Abs(value) <= 150,
            "Chroma" => value is >= 0 and <= 200,
            "Hue" => value is >= 0 and <= 360,
            _ => true
        };

        // ── Aplicar corrección ────────────────────────────────────────────────

        private void ApplyCorrection(List<ColorimetricRow> measures, CorrectionResult correction)
        {
            foreach (var row in measures)
            {
                if (!string.Equals(row.Illuminant, correction.Illuminant, StringComparison.OrdinalIgnoreCase)) continue;
                if (!string.Equals(row.Type, correction.Type, StringComparison.OrdinalIgnoreCase)) continue;

                switch (correction.Field)
                {
                    case "a": row.A = correction.CorrectedValue; break;
                    case "b": row.B = correction.CorrectedValue; break;
                    case "L": row.L = correction.CorrectedValue; break;
                    case "Chroma": row.Chroma = correction.CorrectedValue; break;
                    case "Hue": row.Hue = correction.CorrectedValue; break;
                }

                // Recalcular Hue automáticamente si se corrigió a o b
                if (correction.Field is "a" or "b")
                {
                    double newHue = Math.Atan2(row.B, row.A) * 180.0 / Math.PI;
                    if (newHue < 0) newHue += 360.0;
                    row.Hue = Math.Round(newHue, 2);
                }

                row.NeedsReview = false;
                break;
            }
        }
    }
}