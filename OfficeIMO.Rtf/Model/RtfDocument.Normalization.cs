using OfficeIMO.Rtf.Diagnostics;

namespace OfficeIMO.Rtf;

public sealed partial class RtfDocument {
    private List<RtfConversionDiagnostic> _sourceNormalizationDiagnostics = new List<RtfConversionDiagnostic>();

    /// <summary>Writes normalized semantic RTF with an operation-scoped report of source content that cannot survive normalization.</summary>
    /// <remarks>Use the read result's lossless APIs when original syntax must survive. Call RequireNoLoss on this result before accepting strict semantic output.</remarks>
    public RtfConversionResult<string> ToRtfResult(RtfWriteOptions? options = null) {
        RtfWriteOptions resolved = options ?? new RtfWriteOptions();
        var report = new RtfConversionReport();
        foreach (RtfConversionDiagnostic diagnostic in _sourceNormalizationDiagnostics) report.Add(diagnostic);
        if (HtmlEncapsulation != null && (!resolved.IncludeHtmlEncapsulation || !IsHtmlEncapsulationCurrent)) {
            report.Add(RtfConversionSeverity.Warning, "RtfNormalizationHtmlOmitted", "Original encapsulated HTML is omitted from semantic output because it is disabled or no longer corresponds to the edited document.", RtfConversionAction.Omitted, feature: "htmltag");
        }
        return new RtfConversionResult<string>(ToRtf(resolved), report);
    }

    internal void RecordUnboundDestination(string? destination, int position) {
        destination ??= "unnamed";
        _sourceNormalizationDiagnostics.Add(new RtfConversionDiagnostic(RtfConversionSeverity.Warning,
            "RtfNormalizationDestinationOmitted", "Source destination '" + destination + "' is not represented by the semantic model and is omitted by normalized writing.",
            RtfConversionAction.Omitted, "Source/" + position.ToString(CultureInfo.InvariantCulture), destination));
    }

    internal void RecordReadNormalizationDiagnostics(IEnumerable<RtfDiagnostic> diagnostics) {
        var report = new RtfConversionReport();
        report.AddReadDiagnostics(diagnostics.Where(diagnostic => diagnostic.Code != "RTF101"));
        _sourceNormalizationDiagnostics.AddRange(report.Diagnostics);
    }
}
