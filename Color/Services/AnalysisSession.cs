using System;
using System.Collections.Generic;
using System.Data;

namespace Color
{
    /// <summary>
    /// Centro de datos unificado para la sesión activa de análisis.
    /// Garantiza la persistencia de datos al navegar entre módulos.
    /// </summary>
    public class AnalysisSession
    {
        private static AnalysisSession _instance;
        public static AnalysisSession Instance => _instance ?? (_instance = new AnalysisSession());

        // Datos del Análisis Activo
        public object LastShadeResult { get; set; }
        public System.Collections.IEnumerable CurrentCorrections { get; set; } // List<IlluminantCorrectionResult>
        public System.Collections.IEnumerable CurrentRecipe { get; set; } // List<IlluminantCorrectionResult>
        public DataTable CurrentHistoryTable { get; set; }

        public bool HasActiveData => CurrentCorrections != null;

        // Evento para notificar cambios (Real-time update)
        public event EventHandler DataUpdated;

        public void NotifyUpdate()
        {
            DataUpdated?.Invoke(this, EventArgs.Empty);
        }

        public void Clear()
        {
            LastShadeResult = null;
            CurrentCorrections = null;
            CurrentRecipe = null;
            NotifyUpdate();
        }
    }
}
