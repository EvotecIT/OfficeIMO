using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Rtf.Pdf;

/// <content>Converts the first-party logical PDF model to RTF.</content>
public static partial class RtfPdfConverterExtensions {
    /// <summary>Converts a logical PDF model into an editable RTF document.</summary>
    public static RtfDocument ToRtfDocument(
        this PdfCore.PdfLogicalDocument document,
        PdfRtfReadOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return PdfRtfConverter.Convert(document, options);
    }
}
