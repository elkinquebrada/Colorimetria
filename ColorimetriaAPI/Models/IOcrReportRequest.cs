// Models.cs — Modelos de datos de la API Colorimetría

namespace ColorimetriaAPI.Models
{
    public interface IOcrReportRequest
    {
        double ChromaThreshold { get; set; }
        List<ColorimetricRow> Measures { get; set; }
    }
}