namespace OfficeIMO.Pdf;

/// <summary>
/// Collects shared PDF conversion warnings from one adapter run.
/// </summary>
public sealed class PdfConversionReport {
    private readonly List<PdfConversionWarning> _warnings = new();
    private readonly List<PdfConversionReport> _linkedReports = new();

    /// <summary>Warnings recorded by the converter in production order.</summary>
    public IReadOnlyList<PdfConversionWarning> Warnings {
        get {
            if (_linkedReports.Count == 0) {
                return _warnings;
            }

            int capacity = _warnings.Count;
            foreach (PdfConversionReport report in _linkedReports) {
                capacity += report.Warnings.Count;
            }

            var warnings = new List<PdfConversionWarning>(capacity);
            warnings.AddRange(_warnings);
            foreach (PdfConversionReport report in _linkedReports) {
                warnings.AddRange(report.Warnings);
            }

            return warnings;
        }
    }

    /// <summary>True when at least one warning was recorded.</summary>
    public bool HasWarnings {
        get {
            if (_warnings.Count > 0) {
                return true;
            }

            foreach (PdfConversionReport report in _linkedReports) {
                if (report.HasWarnings) {
                    return true;
                }
            }

            return false;
        }
    }

    /// <summary>True when at least one error-severity warning was recorded.</summary>
    public bool HasErrors => Warnings.Any(static warning => warning.Severity == PdfConversionWarningSeverity.Error);

    /// <summary>True when conversion reported an approximation, omission, or error.</summary>
    public bool HasLoss => Warnings.Any(static warning => warning.Severity != PdfConversionWarningSeverity.Information);

    /// <summary>
    /// High-level fidelity outcome derived from the structured warnings. Declared font-family and
    /// bounded typography substitutions are distinguished from other lossy or unsupported behavior.
    /// </summary>
    public PdfConversionFidelityStatus FidelityStatus {
        get {
            PdfConversionWarning[] lossWarnings = Warnings
                .Where(static warning => warning.Severity != PdfConversionWarningSeverity.Information)
                .ToArray();
            if (lossWarnings.Length == 0) {
                return PdfConversionFidelityStatus.Faithful;
            }

            return lossWarnings.All(IsDeclaredSubstitutionWarning)
                ? PdfConversionFidelityStatus.FaithfulWithSubstitutions
                : PdfConversionFidelityStatus.Degraded;
        }
    }

    /// <summary>
    /// Builds a stable count summary for proof packs, logs, wrapper routing, and user-facing diagnostics.
    /// </summary>
    public PdfConversionReportSummary Summarize() {
        return new PdfConversionReportSummary(Warnings);
    }

    /// <summary>
    /// Throws when the report contains any conversion warning; otherwise returns the current report for fluent checks.
    /// </summary>
    public PdfConversionReport RequireNoWarnings() {
        if (HasWarnings) {
            throw new InvalidOperationException(CreateFailureMessage("PDF conversion produced warnings.", Warnings));
        }

        return this;
    }

    /// <summary>
    /// Throws when conversion reported an approximation, omission, or error; informational diagnostics are allowed.
    /// </summary>
    public PdfConversionReport RequireNoLoss() {
        PdfConversionWarning[] lossWarnings = Warnings
            .Where(static warning => warning.Severity != PdfConversionWarningSeverity.Information)
            .ToArray();
        if (lossWarnings.Length > 0) {
            throw new InvalidOperationException(CreateFailureMessage("PDF conversion reported possible content loss.", lossWarnings));
        }

        return this;
    }

    /// <summary>
    /// Throws when the report contains error-severity conversion warnings; otherwise returns the current report for fluent checks.
    /// </summary>
    public PdfConversionReport RequireNoErrorWarnings() {
        PdfConversionWarning[] errors = Warnings
            .Where(static warning => warning.Severity == PdfConversionWarningSeverity.Error)
            .ToArray();
        if (errors.Length > 0) {
            throw new InvalidOperationException(CreateFailureMessage("PDF conversion produced error warnings.", errors));
        }

        return this;
    }

    /// <summary>Adds one warning to the report.</summary>
    public void Add(PdfConversionWarning warning) {
        Guard.NotNull(warning, nameof(warning));
        _warnings.Add(warning);
    }

    /// <summary>Adds all warnings from another report.</summary>
    public void AddRange(IEnumerable<PdfConversionWarning> warnings) {
        Guard.NotNull(warnings, nameof(warnings));
        foreach (PdfConversionWarning warning in warnings) {
            Add(warning);
        }
    }

    /// <summary>Adds shared text encoding diagnostics as conversion warnings.</summary>
    public void AddTextDiagnostics(IEnumerable<PdfTextEncodingDiagnostic> diagnostics, string converter = "OfficeIMO.Pdf") {
        Guard.NotNull(diagnostics, nameof(diagnostics));
        foreach (PdfTextEncodingDiagnostic diagnostic in diagnostics) {
            Add(diagnostic.ToConversionWarning(converter));
        }
    }

    /// <summary>Adds missing-glyph diagnostics from an embedded-font fallback plan as conversion warnings.</summary>
    public void AddTextFallbackPlanDiagnostics(PdfTextFallbackPlan plan, string converter = "OfficeIMO.Pdf") {
        Guard.NotNull(plan, nameof(plan));
        AddTextDiagnostics(plan.Diagnostics, converter);
    }

    /// <summary>Adds shared text shaping diagnostics as conversion warnings.</summary>
    public void AddTextShapingDiagnostics(IEnumerable<PdfTextShapingDiagnostic> diagnostics, string converter = "OfficeIMO.Pdf") {
        Guard.NotNull(diagnostics, nameof(diagnostics));
        foreach (PdfTextShapingDiagnostic diagnostic in diagnostics) {
            Add(diagnostic.ToConversionWarning(converter));
        }
    }

    /// <summary>Adds shared embedded-font diagnostics as conversion warnings.</summary>
    public void AddFontDiagnostics(IEnumerable<PdfFontEmbeddingDiagnostic> diagnostics, string converter = "OfficeIMO.Pdf") {
        Guard.NotNull(diagnostics, nameof(diagnostics));
        foreach (PdfFontEmbeddingDiagnostic diagnostic in diagnostics) {
            Add(diagnostic.ToConversionWarning(converter));
        }
    }

    /// <summary>Clears warnings from a previous conversion run.</summary>
    public void Clear() {
        _warnings.Clear();
        _linkedReports.Clear();
    }

    internal void LinkReport(PdfConversionReport report) {
        Guard.NotNull(report, nameof(report));
        if (ReferenceEquals(this, report) || _linkedReports.Contains(report)) {
            return;
        }

        _linkedReports.Add(report);
    }

    internal void ClearLinkedReports() {
        _linkedReports.Clear();
    }

    private static string CreateFailureMessage(string message, IReadOnlyList<PdfConversionWarning> warnings) {
        if (warnings.Count == 0) {
            return message;
        }

        return message + " First warning: " + warnings[0].ToString();
    }

    private static bool IsDeclaredSubstitutionWarning(PdfConversionWarning warning) =>
        warning.Code.EndsWith("FontFamilySubstituted", StringComparison.OrdinalIgnoreCase) ||
        warning.Code.EndsWith("FontSubstitution", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(warning.Code, "font-family-substitution", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(warning.Code, "unsupported-font-ligature-substitution", StringComparison.OrdinalIgnoreCase);
}
