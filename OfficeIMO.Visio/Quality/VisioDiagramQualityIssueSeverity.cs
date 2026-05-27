namespace OfficeIMO.Visio {
    /// <summary>
    /// Severity of a visual quality issue found in a diagram.
    /// </summary>
    public enum VisioDiagramQualityIssueSeverity {
        /// <summary>Informational issue.</summary>
        Information,

        /// <summary>Potential visual problem.</summary>
        Warning,

        /// <summary>Likely visual defect.</summary>
        Error
    }
}
