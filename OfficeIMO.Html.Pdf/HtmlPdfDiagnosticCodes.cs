namespace OfficeIMO.Html.Pdf;

/// <summary>Stable diagnostics emitted by the direct rendered HTML-to-PDF adapter.</summary>
public static class HtmlPdfDiagnosticCodes {
    /// <summary>The current PDF font slots could not represent every distinct active web-font family.</summary>
    public const string RenderedFontFamilyLimitExceeded = "HtmlPdfRenderedFontFamilyLimitExceeded";
}
