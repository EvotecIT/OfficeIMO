namespace OfficeIMO;

/// <summary>Raised when an accessibility policy rejects a document.</summary>
public sealed class OfficeAccessibilityException : InvalidOperationException {
    internal OfficeAccessibilityException(OfficeAccessibilityReport report, bool includeWarnings)
        : base(report.Format + " accessibility inspection found " + report.ErrorCount + " error(s) and " + report.WarningCount + " warning(s).") {
        Report = report;
        IncludedWarnings = includeWarnings;
    }

    /// <summary>Report that caused the policy failure.</summary>
    public OfficeAccessibilityReport Report { get; }
    /// <summary>Whether warnings participated in the failure threshold.</summary>
    public bool IncludedWarnings { get; }
}
