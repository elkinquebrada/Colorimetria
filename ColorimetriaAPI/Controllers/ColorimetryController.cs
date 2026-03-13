// ColorimetryController.cs — Endpoints de la API
using ColorimetriaAPI.Models;
using ColorimetriaAPI.Services;
using Microsoft.AspNetCore.Mvc;

namespace ColorimetriaAPI.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    [Produces("application/json")]
    public class ColorimetryController : ControllerBase
    {
        private readonly ClaudeService _claude;
        private readonly ILogger<ColorimetryController> _logger;

        public ColorimetryController(ClaudeService claude, ILogger<ColorimetryController> logger)
        {
            _claude = claude;
            _logger = logger;
        }

        /// <summary>
        /// Healthcheck — verifica que la API está corriendo
        /// </summary>
        [HttpGet("health")]
        public IActionResult Health() =>
            Ok(new { status = "ok", timestamp = DateTime.UtcNow });

        /// <summary>
        /// Corrige tokens colorimétricos erróneos detectados por OCR.
        /// Envía las medidas CIELab y recibe las correcciones sugeridas por Claude.
        /// </summary>
        /// <remarks>
        /// Ejemplo de body:
        /// {
        ///   "chromaThreshold": 0.35,
        ///   "measures": [
        ///     { "illuminant": "D65", "type": "Std", "l": 52.3, "a": -3.6, "b": -36.0, "chroma": 5.06, "hue": 264.2 }
        ///   ]
        /// }
        /// </remarks>
        [HttpPost("correct")]
        [ProducesResponseType(typeof(CorrectReportResponse), 200)]
        [ProducesResponseType(400)]
        [ProducesResponseType(500)]
        public async Task<IActionResult> CorrectReport([FromBody] OcrReportRequest request)
        {
            if (request == null)
                return BadRequest(new { error = "Body requerido." });

            if (request.Measures == null || request.Measures.Count == 0)
                return BadRequest(new { error = "Se requiere al menos una medida en 'measures'." });

            _logger.LogInformation("POST /correct — {Count} medidas, threshold={Threshold}",
                request.Measures.Count, request.ChromaThreshold);

            var result = await _claude.CorrectReportAsync(request);

            if (!result.Success)
                return StatusCode(500, result);

            return Ok(result);
        }

        /// <summary>
        /// Valida coherencia colorimétrica sin llamar a Claude.
        /// Útil para pre-filtrar antes de enviar al corrector.
        /// </summary>
        [HttpPost("validate")]
        [ProducesResponseType(200)]
        public IActionResult Validate([FromBody] OcrReportRequest request)
        {
            if (request?.Measures == null)
                return BadRequest(new { error = "Body requerido." });

            var issues = new List<object>();

            foreach (var row in request.Measures)
            {
                double chromaCalc = Math.Sqrt(row.A * row.A + row.B * row.B);
                double chromaErr = Math.Abs(chromaCalc - row.Chroma);

                double hueCalc = Math.Atan2(row.B, row.A) * 180.0 / Math.PI;
                if (hueCalc < 0) hueCalc += 360.0;
                double hueErr = Math.Abs(row.Hue - hueCalc);
                if (hueErr > 180) hueErr = 360 - hueErr;

                if (chromaErr > request.ChromaThreshold || hueErr > 5.0)
                    issues.Add(new
                    {
                        illuminant = row.Illuminant,
                        type = row.Type,
                        chromaError = Math.Round(chromaErr, 4),
                        hueError = Math.Round(hueErr, 2),
                        needsCorrection = true
                    });
            }

            return Ok(new
            {
                total = request.Measures.Count,
                issuesFound = issues.Count,
                coherent = request.Measures.Count - issues.Count,
                issues
            });
        }
    }
}
