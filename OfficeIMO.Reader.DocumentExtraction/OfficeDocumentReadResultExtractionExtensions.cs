using Microsoft.Extensions.Documents;

namespace OfficeIMO.Reader.DocumentExtraction;

/// <summary>Converts OfficeIMO read results to the shared extraction document model.</summary>
public static class OfficeDocumentReadResultExtractionExtensions {
    /// <summary>
    /// Maps a read result to an <see cref="ExtractionDocument"/>: headings open nested sections, paragraphs and list items
    /// become text elements, tables keep their header row, images carry alternative and OCR text, and page or slide
    /// frames are listed in <see cref="ExtractionDocument.SourcePages"/>.
    /// </summary>
    /// <param name="result">The OfficeIMO result, kept as the document's <c>RawRepresentation</c>.</param>
    /// <param name="identifier">The identifier for the resulting document.</param>
    /// <param name="options">Optional mapping settings.</param>
    public static ExtractionDocument ToExtractionDocument(
        this OfficeDocumentReadResult result,
        string identifier,
        OfficeIMODocumentExtractionOptions? options = null) =>
        ExtractionDocumentMapper.Map(result, identifier, options?.Clone());
}
