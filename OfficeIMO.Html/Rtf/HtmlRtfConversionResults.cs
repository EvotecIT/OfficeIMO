namespace OfficeIMO.Html;

/// <summary>RTF document plus structured diagnostics from HTML import.</summary>
public sealed class HtmlToRtfResult : HtmlConversionResult<RtfDocument> {
    internal HtmlToRtfResult(
        RtfDocument document,
        IEnumerable<HtmlDiagnostic> diagnostics,
        IReadOnlyList<HtmlRtfConversionDiagnostic> rtfDiagnostics,
        RtfConversionReport report) : base(document) {
        Diagnostics.AddRange(diagnostics);
        RtfDiagnostics = Array.AsReadOnly(rtfDiagnostics.ToArray());
        Report = Snapshot(report);
    }

    /// <summary>Imported RTF document.</summary>
    public RtfDocument Document => Artifact;

    /// <summary>RTF-specific diagnostics in emission order.</summary>
    public IReadOnlyList<HtmlRtfConversionDiagnostic> RtfDiagnostics { get; }

    /// <summary>Shared RTF fidelity report for this operation.</summary>
    public RtfConversionReport Report { get; }

    private static RtfConversionReport Snapshot(RtfConversionReport report) {
        var snapshot = new RtfConversionReport();
        snapshot.Merge(report);
        return snapshot;
    }
}

/// <summary>Semantic HTML plus structured diagnostics from RTF export.</summary>
public sealed class RtfToHtmlResult : HtmlConversionResult<string> {
    internal RtfToHtmlResult(
        string html,
        IEnumerable<HtmlDiagnostic> diagnostics,
        IReadOnlyList<HtmlRtfConversionDiagnostic> rtfDiagnostics,
        RtfConversionReport report) : base(html) {
        Diagnostics.AddRange(diagnostics);
        RtfDiagnostics = Array.AsReadOnly(rtfDiagnostics.ToArray());
        Report = Snapshot(report);
    }

    /// <summary>Exported semantic HTML.</summary>
    public string Html => Artifact;

    /// <summary>RTF-specific diagnostics in emission order.</summary>
    public IReadOnlyList<HtmlRtfConversionDiagnostic> RtfDiagnostics { get; }

    /// <summary>Shared RTF fidelity report for this operation.</summary>
    public RtfConversionReport Report { get; }

    private static RtfConversionReport Snapshot(RtfConversionReport report) {
        var snapshot = new RtfConversionReport();
        snapshot.Merge(report);
        return snapshot;
    }
}
