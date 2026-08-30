namespace OfficeIMO.Html.Pdf;

/// <summary>Stable diagnostics emitted by the direct rendered HTML-to-PDF adapter.</summary>
public static class HtmlPdfDiagnosticCodes {
    /// <summary>A shaped font program was painted as vector outlines while retaining logical text for extraction and accessibility.</summary>
    public const string FontProgramOutlined = "HtmlPdfFontProgramOutlined";

    /// <summary>Compatibility diagnostic retained for callers that inspect older results; named PDF font resources no longer impose a fixed family limit.</summary>
    public const string RenderedFontFamilyLimitExceeded = "HtmlPdfRenderedFontFamilyLimitExceeded";
}
