namespace OfficeIMO.Word.LegacyDoc.Diagnostics {
    /// <summary>
    /// Severity of a legacy DOC import diagnostic.
    /// </summary>
    public enum LegacyDocDiagnosticSeverity {
        /// <summary>An informational import note that does not affect projection success.</summary>
        Info,
        /// <summary>A non-fatal import limitation or preservation boundary.</summary>
        Warning,
        /// <summary>A fatal import problem that prevents safe projection.</summary>
        Error
    }
}
