namespace Color.Models
{
    /// <summary>
    /// Entidad de transporte para los metadatos de produccion textil del encabezado del reporte.
    /// Totalmente desacoplada del OcrReport y del motor matematico.
    /// </summary>
    public class TextileMetadata
    {
        public string ShadeName  { get; set; } = "-";
        public string DyeingClass { get; set; } = "-";
        public string Substrate  { get; set; } = "-";
        public string CountPly   { get; set; } = "-";
        public string FiberType  { get; set; } = "-";

        /// <summary>
        /// Verdadero si se extrajo al menos un campo con dato real.
        /// </summary>
        public bool IsValid => ShadeName != "-" || DyeingClass != "-"
                            || Substrate != "-"  || CountPly   != "-"
                            || FiberType != "-";
    }
}
