namespace OfficeIMO.Pdf;

/// <summary>Structured interchange exports for the first-party logical PDF model.</summary>
public static class PdfLogicalDocumentStructuredExportExtensions {
    /// <summary>Exports an already parsed logical document without rerunning extraction.</summary>
    public static string ExportStructured(this PdfLogicalDocument document, PdfStructuredExportFormat format) {
        Guard.NotNull(document, nameof(document));
        return PdfStructuredExportEngine.Export(document, format);
    }

    /// <summary>
    /// Exports one schema-valid PAGE XML document per logical page. PAGE XML is image/page scoped
    /// and does not define a multi-page root.
    /// </summary>
    public static IReadOnlyList<string> ToPageXmlDocuments(this PdfLogicalDocument document) {
        Guard.NotNull(document, nameof(document));
        return PdfStructuredExportEngine.ExportPageXmlDocuments(document);
    }
}
