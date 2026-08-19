using OfficeIMO.Drawing;
using OfficeIMO.Html;

namespace OfficeIMO.PowerPoint.Html;

/// <summary>PowerPoint HTML output and visual snapshot diagnostics from one conversion.</summary>
public sealed class PowerPointToHtmlResult : HtmlConversionResult<string> {
    internal PowerPointToHtmlResult(
        string value,
        IEnumerable<OfficeImageExportDiagnostic> imageDiagnostics,
        IEnumerable<HtmlDiagnostic> conversionDiagnostics)
        : base(value) {
        if (imageDiagnostics == null) throw new ArgumentNullException(nameof(imageDiagnostics));
        if (conversionDiagnostics == null) throw new ArgumentNullException(nameof(conversionDiagnostics));
        ImageDiagnostics = Array.AsReadOnly(imageDiagnostics.ToArray());
        AddDiagnostics(ImageDiagnostics.Select(ToHtmlDiagnostic));
        AddDiagnostics(conversionDiagnostics);
    }

    /// <summary>Visual snapshot diagnostics captured while rendering positioned review HTML.</summary>
    public IReadOnlyList<OfficeImageExportDiagnostic> ImageDiagnostics { get; }

    /// <summary>True when the visual snapshot reported an approximation or unsupported feature.</summary>
    public bool HasImageDiagnostics => ImageDiagnostics.Count > 0;

    private static HtmlDiagnostic ToHtmlDiagnostic(OfficeImageExportDiagnostic diagnostic) {
        HtmlDiagnosticSeverity severity = diagnostic.Severity switch {
            OfficeImageExportDiagnosticSeverity.Error => HtmlDiagnosticSeverity.Error,
            OfficeImageExportDiagnosticSeverity.Warning => HtmlDiagnosticSeverity.Warning,
            _ => HtmlDiagnosticSeverity.Info
        };
        OfficeConversionLossKind lossKind = diagnostic.Severity switch {
            OfficeImageExportDiagnosticSeverity.Error => OfficeConversionLossKind.Failure,
            OfficeImageExportDiagnosticSeverity.Warning => OfficeConversionLossKind.Approximation,
            _ => OfficeConversionLossKind.None
        };
        return new HtmlDiagnostic(
            "OfficeIMO.PowerPoint.Html",
            diagnostic.Code,
            diagnostic.Message,
            severity,
            diagnostic.Source,
            lossKind: lossKind);
    }
}
