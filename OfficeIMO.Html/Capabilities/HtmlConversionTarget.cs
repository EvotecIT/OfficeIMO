namespace OfficeIMO.Html;

/// <summary>Built-in target families understood by shared HTML capability analysis.</summary>
public enum HtmlConversionTarget {
    /// <summary>Editable Word document content.</summary>
    Word,
    /// <summary>Editable Excel workbook content.</summary>
    Excel,
    /// <summary>Editable PowerPoint presentation content.</summary>
    PowerPoint,
    /// <summary>Editable offline OneNote content.</summary>
    OneNote,
    /// <summary>Markdown text and document models.</summary>
    Markdown,
    /// <summary>Semantic Rich Text Format content.</summary>
    Rtf,
    /// <summary>Rendered PDF output.</summary>
    Pdf,
    /// <summary>Rendered PNG, JPEG, TIFF, SVG, or WebP output.</summary>
    Image,
    /// <summary>Structured OfficeIMO.Reader output.</summary>
    Reader
}
