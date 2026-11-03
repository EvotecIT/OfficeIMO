using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Rtf.Pdf;

/// <summary>Converts parsed <see cref="RtfDocument"/> models to PDF.</summary>
public static partial class RtfPdfConverterExtensions {
    /// <summary>Converts an RTF document to PDF with operation-scoped diagnostics.</summary>
    internal static PdfCore.PdfDocumentConversionResult ConvertToPdfWithEvidence(
        this RtfDocument document,
        RtfPdfExportOptions? options = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (document == null) throw new ArgumentNullException(nameof(document));
        RtfPdfExportOptions operation = (options ?? new RtfPdfExportOptions()).CloneForConversion();
        operation.CancellationToken = cancellationToken;
        PdfCore.PdfDocument pdf = RtfPdfConverter.Convert(document, operation);
        return new PdfCore.PdfDocumentConversionResult(pdf, operation.Report);
    }

}
