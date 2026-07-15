namespace OfficeIMO.PowerPoint.LegacyPpt.Diagnostics {
    /// <summary>Identifies the severity of a legacy PowerPoint import diagnostic.</summary>
    public enum LegacyPptDiagnosticSeverity {
        /// <summary>Additional information that does not indicate data loss.</summary>
        Information,

        /// <summary>Content was recognized but could not be projected completely.</summary>
        Warning,

        /// <summary>The binary presentation could not be imported safely.</summary>
        Error
    }
}
