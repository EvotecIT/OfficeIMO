using PdfPowerPoint = OfficeIMO.Pdf.PowerPoint;
using OfficeIMO.Pdf.PowerPoint;
using OfficeIMO.PowerPoint.OpenDocument;
using PowerPointPdf = OfficeIMO.PowerPoint.Pdf;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.OpenDocument.Odp.Pdf;

/// <summary>Direct, loss-aware ODP to PDF conversion through the PowerPoint semantic and PDF engines.</summary>
public static partial class OdpPdfConversionExtensions {
    /// <summary>Converts an ODP presentation to PDF and preserves diagnostics from both conversion stages.</summary>
    private static PdfCore.PdfDocumentConversionResult ConvertToPdfCore(this OdpPresentation source, PowerPointOpenDocumentConversionOptions? conversionOptions = null, PowerPointPdf.PowerPointPdfExportOptions? pdfOptions = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (source == null) throw new ArgumentNullException(nameof(source));
        OdfConversionResult<OfficeIMO.PowerPoint.PowerPointDocument> conversion = source.ToPowerPointDocumentResult(conversionOptions);
        using (conversion.Value) {
            PdfCore.PdfDocumentConversionResult result = PowerPointPdf.PowerPointPdfConverterExtensions.ConvertToPdfWithEvidence(conversion.Value, pdfOptions, cancellationToken);
            return result.WithSourceConversionReport(conversion.Report);
        }
    }

    /// <summary>Reconstructs an ODP presentation and preserves diagnostics from both conversion stages.</summary>
    private static PdfOdpConversionResult ConvertOdpCore(this PdfCore.PdfDocument source, PdfPowerPoint.PdfPowerPointExportOptions? pdfOptions = null, PowerPointOpenDocumentConversionOptions? openDocumentOptions = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (source == null) throw new ArgumentNullException(nameof(source));
        PdfPowerPoint.PdfPowerPointConversionResult pdf =
            PdfPowerPoint.PdfPowerPointConverterExtensions.ConvertToPowerPointWithEvidence(source, pdfOptions, cancellationToken);
        using (pdf.Value) {
            OdfConversionResult<OdpPresentation> odp = pdf.Value.ToOpenDocumentResult(openDocumentOptions);
            cancellationToken.ThrowIfCancellationRequested();
            return new PdfOdpConversionResult(
                odp.Value,
                new PdfOdpConversionReport(pdf.Report, odp.Report));
        }
    }

    /// <summary>Reconstructs an ODP presentation from a logical PDF model and preserves both stage reports.</summary>
    /// <remarks>A logical PDF model resolves Auto to editable tables and also supports explicit editable-content reconstruction. Visual and hybrid projections require an opened PDF.</remarks>
    private static PdfOdpConversionResult ConvertOdpCore(this PdfCore.PdfDocumentReadResult source, PdfPowerPoint.PdfPowerPointExportOptions? pdfOptions = null, PowerPointOpenDocumentConversionOptions? openDocumentOptions = null, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (source == null) throw new ArgumentNullException(nameof(source));
        PdfPowerPoint.PdfPowerPointConversionResult pdf =
            PdfPowerPoint.PdfPowerPointConverterExtensions.ConvertToPowerPointWithEvidence(source, pdfOptions, cancellationToken);
        using (pdf.Value) {
            OdfConversionResult<OdpPresentation> odp = pdf.Value.ToOpenDocumentResult(openDocumentOptions);
            cancellationToken.ThrowIfCancellationRequested();
            return new PdfOdpConversionResult(
                odp.Value,
                new PdfOdpConversionReport(pdf.Report, odp.Report));
        }
    }

}
