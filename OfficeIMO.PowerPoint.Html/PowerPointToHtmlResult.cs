using OfficeIMO.Drawing;
using OfficeIMO.Html;

namespace OfficeIMO.PowerPoint.Html;

/// <summary>PowerPoint HTML output and visual snapshot diagnostics from one conversion.</summary>
public sealed class PowerPointToHtmlResult : HtmlConversionResult<string> {
    internal PowerPointToHtmlResult(string value, IEnumerable<OfficeImageExportDiagnostic> imageDiagnostics)
        : base(value) {
        if (imageDiagnostics == null) throw new ArgumentNullException(nameof(imageDiagnostics));
        ImageDiagnostics = Array.AsReadOnly(imageDiagnostics.ToArray());
    }

    /// <summary>Visual snapshot diagnostics captured while rendering positioned review HTML.</summary>
    public IReadOnlyList<OfficeImageExportDiagnostic> ImageDiagnostics { get; }

    /// <summary>True when the visual snapshot reported an approximation or unsupported feature.</summary>
    public bool HasImageDiagnostics => ImageDiagnostics.Count > 0;
}
