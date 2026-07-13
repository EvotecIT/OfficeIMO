using OfficeIMO.Drawing;
using OfficeIMO.Html;

namespace OfficeIMO.PowerPoint.Html;

/// <summary>PowerPoint HTML output and visual snapshot diagnostics from one conversion.</summary>
public sealed class PowerPointToHtmlResult : HtmlConversionResult<string> {
    internal PowerPointToHtmlResult(string value, IEnumerable<OfficeImageExportDiagnostic> imageDiagnostics)
        : base(value) {
        if (imageDiagnostics == null) throw new ArgumentNullException(nameof(imageDiagnostics));
        ImageDiagnostics = Array.AsReadOnly(imageDiagnostics.ToArray());
        AddDiagnostics(ImageDiagnostics.Select(ToHtmlDiagnostic));
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
        HtmlConversionLossKind lossKind = diagnostic.Severity switch {
            OfficeImageExportDiagnosticSeverity.Error => HtmlConversionLossKind.Failure,
            OfficeImageExportDiagnosticSeverity.Warning => HtmlConversionLossKind.Approximation,
            _ => HtmlConversionLossKind.None
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
