using System.Diagnostics;
using OfficeIMO.Pdf;
using OfficeIMO.Reader;

namespace OfficeIMO;

/// <summary>Converts normalized Reader content through the canonical PDF projection.</summary>
public static partial class OfficeReaderConversionExtensions {
    /// <summary>Projects normalized Reader content into a PDF model.</summary>
    public static PdfDocument ToPdf(this OfficeDocumentReadResult source, PdfProjectionOptions? options = null,
        CancellationToken cancellationToken = default) => source.ToPdf(options, out _, cancellationToken);

    /// <summary>Projects normalized content and returns source and PDF fidelity findings.</summary>
    public static PdfDocument ToPdf(this OfficeDocumentReadResult source, PdfProjectionOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        var timer = Stopwatch.StartNew();
        PdfProjectionOptions settings = (options ?? new PdfProjectionOptions()).Snapshot();
        PdfDocumentConversionResult conversion = source.ToPdfDocumentResult(settings, cancellationToken);
        return PdfConversionEvidence.CompleteModel(conversion, source.Kind.ToString(), settings.FailOnLoss,
            timer, cancellationToken, out report);
    }

    /// <summary>Returns the projected model and its completed fidelity report together.</summary>
    public static async Task<OfficeConversionResult<PdfDocument>> ToPdfAsync(this OfficeDocumentReadResult source,
        PdfProjectionOptions? options = null, CancellationToken cancellationToken = default) {
        PdfProjectionOptions settings = (options ?? new PdfProjectionOptions()).Snapshot();
        return await Task.Run(() => {
            PdfDocument value = source.ToPdf(settings, out OfficeConversionReport report, cancellationToken);
            if (cancellationToken.IsCancellationRequested) {
                value.Dispose();
                cancellationToken.ThrowIfCancellationRequested();
            }
            return new OfficeConversionResult<PdfDocument>(value, report);
        }, cancellationToken).ConfigureAwait(false);
    }
}
