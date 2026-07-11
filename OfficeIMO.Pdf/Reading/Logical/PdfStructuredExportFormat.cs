namespace OfficeIMO.Pdf;

/// <summary>Stable structured-text interchange formats emitted from the logical PDF model.</summary>
public enum PdfStructuredExportFormat {
    /// <summary>OfficeIMO logical JSON schema.</summary>
    Json,

    /// <summary>Human-readable Markdown.</summary>
    Markdown,

    /// <summary>ALTO XML 4.4.</summary>
    AltoXml,

    /// <summary>hOCR HTML with page and line bounding boxes.</summary>
    Hocr,

    /// <summary>PAGE XML 2019-07-15.</summary>
    PageXml
}
