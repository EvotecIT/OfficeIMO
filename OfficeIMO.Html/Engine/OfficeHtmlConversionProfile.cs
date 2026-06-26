namespace OfficeIMO.Html;

/// <summary>
/// Shared source-specific Office-to-HTML lane identifiers used before dedicated adapters need their own package API.
/// </summary>
public enum OfficeHtmlConversionProfile {
    /// <summary>Word document content as accessible semantic HTML.</summary>
    WordSemanticDocument,

    /// <summary>Word document HTML intended for editable HTML to Word to HTML roundtrip review.</summary>
    WordDocumentRoundTrip,

    /// <summary>Word document HTML intended for print-oriented review without claiming browser layout parity.</summary>
    WordPrintReview,

    /// <summary>Workbook and worksheet content as accessible semantic HTML tables.</summary>
    ExcelSemanticTables,

    /// <summary>Worksheet ranges, pages, and shapes as positioned visual-review HTML.</summary>
    ExcelVisualReview,

    /// <summary>Slides, notes, and slide content as accessible semantic HTML.</summary>
    PowerPointSemanticSlides,

    /// <summary>Slides as positioned visual-review HTML backed by shared drawing primitives.</summary>
    PowerPointVisualReview
}
