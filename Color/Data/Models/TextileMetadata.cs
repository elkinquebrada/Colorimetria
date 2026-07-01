namespace Color.Models
{
    /// Entidad de transporte para los metadatos de produccion textil del encabezado del reporte.
    public class TextileMetadata
    {
        public string ShadeName  { get; set; } = "-";
        public string DyeingClass { get; set; } = "-";
        public string Substrate  { get; set; } = "-";
        public string CountPly   { get; set; } = "-";
        public string FiberType  { get; set; } = "-";

        /// Verdadero si se extrajo al menos un campo con dato real.
        public bool IsValid => ShadeName != "-" || DyeingClass != "-"
                            || Substrate != "-"  || CountPly   != "-"
                            || FiberType != "-";
    }
}
