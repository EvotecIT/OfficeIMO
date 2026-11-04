using PdfWord = OfficeIMO.Pdf.Word;
using OfficeIMO.Pdf.Word;
using OfficeIMO.Word.OpenDocument;
using WordPdf = OfficeIMO.Word.Pdf;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.OpenDocument.Odt.Pdf;

/// <summary>Direct, loss-aware ODT to PDF conversion through the Word semantic and PDF engines.</summary>
public static partial class OdtPdfConversionExtensions {
    /// <summary>Converts an ODT document to PDF and preserves diagnostics from both conversion stages.</summary>
    private static PdfCore.PdfDocumentConversionResult ConvertToPdfCore(
        this OdtDocument source,
        WordOpenDocumentConversionOptions? conversionOptions = null,
        WordPdf.WordPdfExportOptions? pdfOptions = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (source == null) throw new ArgumentNullException(nameof(source));
        OdfConversionResult<OfficeIMO.Word.WordDocument> conversion =
            source.ConvertToWordWithEvidence(conversionOptions);
        // The PDF model renders headers and footers from the intermediate Word
        // document when it is written, so that in-memory document must stay open
        // for the lifetime of the returned result.
        PdfCore.PdfDocumentConversionResult result =
            WordPdf.WordPdfConverterExtensions.ConvertToPdfWithEvidence(conversion.Value, pdfOptions, cancellationToken);
        return result.WithSourceConversionReport(conversion.Report);
    }

    /// <summary>Reconstructs an ODT document and preserves diagnostics from both semantic stages.</summary>
    private static PdfOdtConversionResult ConvertOdtCore(
        this PdfCore.PdfDocument source,
        PdfWord.PdfWordExportOptions? pdfOptions = null,
        WordOpenDocumentConversionOptions? openDocumentOptions = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (source == null) throw new ArgumentNullException(nameof(source));
        PdfWord.PdfWordConversionResult pdf =
            PdfWord.PdfWordConverterExtensions.ConvertToWordWithEvidence(source, pdfOptions, cancellationToken);
        using (pdf.Value) {
            OdfConversionResult<OdtDocument> odt = pdf.Value.ConvertToOpenDocumentWithEvidence(openDocumentOptions);
            cancellationToken.ThrowIfCancellationRequested();
            return new PdfOdtConversionResult(
                odt.Value,
                new PdfOdtConversionReport(pdf.Report, odt.Report));
        }
    }

    /// <summary>Reconstructs an ODT document from a logical PDF model and preserves both stage reports.</summary>
    private static PdfOdtConversionResult ConvertOdtCore(
        this PdfCore.PdfDocumentReadResult source,
        PdfWord.PdfWordExportOptions? pdfOptions = null,
        WordOpenDocumentConversionOptions? openDocumentOptions = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (source == null) throw new ArgumentNullException(nameof(source));
        PdfWord.PdfWordConversionResult pdf =
            PdfWord.PdfWordConverterExtensions.ConvertToWordWithEvidence(source, pdfOptions, cancellationToken);
        using (pdf.Value) {
            OdfConversionResult<OdtDocument> odt = pdf.Value.ConvertToOpenDocumentWithEvidence(openDocumentOptions);
            cancellationToken.ThrowIfCancellationRequested();
            return new PdfOdtConversionResult(
                odt.Value,
                new PdfOdtConversionReport(pdf.Report, odt.Report));
        }
    }

}
