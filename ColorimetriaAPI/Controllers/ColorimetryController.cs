
using ColorimetriaAPI.Models;
using ColorimetriaAPI.Services;
using System;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Http; // Único necesario para .NET 4.8
using System.Web.Http.Description;

namespace ColorimetriaAPI.Models
{
    [RoutePrefix("api/colorimetry")]
    public class IOcrReportRequest
    {
        private readonly ColorimetryService _claude;

        public ColorimetryController(ColorimetryService claude)
        {
            _claude = claude;
        }

        [HttpGet]
        [Route("health")]
        public IHttpActionResult Health()
        {
            return Ok(new { status = "ok", timestamp = DateTime.UtcNow });
        }

        [HttpPost]
        [Route("correct")]
        [ResponseType(typeof(CorrectReportResponse))]
        public async Task<IHttpActionResult> CorrectReport([FromBody] OcrReportRequest request)
        {
            if (request == null) return BadRequest("Body requerido.");

            try
            {
                var result = await _claude.CorrectReportAsync(request);
                return Ok(result);
            }
            catch (Exception ex)
            {
                return InternalServerError(ex);
            }
        }

        [HttpPost]
        [Route("validate")]
        public IHttpActionResult Validate([FromBody] OcrReportRequest request)
        {
            if (request == null || request.Measures == null)
                return BadRequest("Datos insuficientes.");

            var issues = new List<object>();
            // Lógica de validación...
            return Ok(new { total = request.Measures.Count, issuesFound = issues.Count, issues });
        }
    }
}