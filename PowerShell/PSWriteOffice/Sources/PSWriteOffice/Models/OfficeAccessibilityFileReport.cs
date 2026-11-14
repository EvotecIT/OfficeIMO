namespace PSWriteOffice.Models;

/// <summary>Accessibility findings for one Word, Excel or PowerPoint file or open document.</summary>
public sealed class OfficeAccessibilityFileReport {
    internal OfficeAccessibilityFileReport(OfficeAccessibilityReport report, string? path) {
        Path = path;
        Report = report;
    }

    /// <summary>The inspected file, or null for an open document.</summary>
    public string? Path { get; }
    /// <summary>The inspected format: <c>Word</c>, <c>PowerPoint</c> or <c>Excel</c>.</summary>
    public string Format => Report.Format;
    /// <summary>The policy profile used.</summary>
    public OfficeAccessibilityPolicyProfile Profile => Report.Profile;
    /// <summary>Number of error findings.</summary>
    public int ErrorCount => Report.ErrorCount;
    /// <summary>Number of warning findings.</summary>
    public int WarningCount => Report.WarningCount;
    /// <summary>Whether no errors were found.</summary>
    public bool IsSuccessful => Report.IsSuccessful;
    /// <summary>Findings with their stable code, severity and location.</summary>
    public IReadOnlyList<OfficeAccessibilityFinding> Findings => Report.Findings;
    /// <summary>The engine report, with <c>ToJson()</c>, <c>ToSarif()</c> and <c>SaveSarif()</c>.</summary>
    public OfficeAccessibilityReport Report { get; }
}
