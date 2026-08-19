namespace OfficeIMO.Html;

/// <summary>Named RTF-to-HTML output contracts.</summary>
public enum RtfHtmlExportProfile {
    /// <summary>RTF document content as safe, accessible semantic HTML.</summary>
    SemanticDocument,

    /// <summary>Trusted editable round-trip output with private OfficeIMO metadata.</summary>
    DocumentRoundTrip,

    /// <summary>Print-oriented review output without claiming full Word layout parity.</summary>
    PrintReview
}
