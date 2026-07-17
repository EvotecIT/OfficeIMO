namespace OfficeIMO.Pdf;

/// <summary>
/// Built-in generated-text fallback groups that OfficeIMO can enable for PDF output.
/// </summary>
[Flags]
public enum PdfTextFallbackFeatures {
    /// <summary>No built-in generated-text fallback groups are enabled.</summary>
    None = 0,

    /// <summary>Try the shared sans-serif document font fallback for generated body, header, and footer text.</summary>
    DocumentFont = 1,

    /// <summary>Try the shared monospace fallback for generated code and preformatted text.</summary>
    MonospaceFont = 2,

    /// <summary>Try installed symbol and emoji font families as embedded fallback runs.</summary>
    SymbolAndEmojiFonts = 4,

    /// <summary>Try installed multilingual font families for CJK, Arabic, and other non-Latin text.</summary>
    MultilingualFonts = 8,

    /// <summary>Enable the recommended OfficeIMO document, monospace, symbol, and emoji text fallback groups.</summary>
    Default = DocumentFont | MonospaceFont | SymbolAndEmojiFonts
}
