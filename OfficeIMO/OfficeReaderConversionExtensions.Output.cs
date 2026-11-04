using System.Diagnostics;
using OfficeIMO.Pdf;
using OfficeIMO.Reader;

namespace OfficeIMO;

public static partial class OfficeReaderConversionExtensions {
    /// <summary>Writes normalized Reader content as PDF and returns completed fidelity findings.</summary>
    public static OfficeConversionReport SaveAsPdf(this OfficeDocumentReadResult source, string path,
        PdfProjectionOptions? options = null, CancellationToken cancellationToken = default) {
        PdfConversionEvidence.ValidateDestination(path);
        return SaveProjection(source, options, (conversion, settings, timer) => settings.FailOnLoss
            ? PdfConversionEvidence.SaveLossless(conversion, source.Kind.ToString(), path, timer, cancellationToken)
            : conversion.Save(path, cancellationToken), cancellationToken);
    }

    /// <summary>Writes normalized Reader content to a caller-owned stream.</summary>
    public static OfficeConversionReport SaveAsPdf(this OfficeDocumentReadResult source, Stream stream,
        PdfProjectionOptions? options = null, CancellationToken cancellationToken = default) {
        PdfConversionEvidence.ValidateDestination(stream);
        return SaveProjection(source, options, (conversion, settings, timer) => settings.FailOnLoss
            ? PdfConversionEvidence.SaveLossless(conversion, source.Kind.ToString(), stream, timer, cancellationToken)
            : conversion.Save(stream, cancellationToken), cancellationToken);
    }

    /// <summary>Writes normalized Reader content as PDF asynchronously.</summary>
    public static Task<OfficeConversionReport> SaveAsPdfAsync(this OfficeDocumentReadResult source, string path,
        PdfProjectionOptions? options = null, CancellationToken cancellationToken = default) {
        PdfConversionEvidence.ValidateDestination(path);
        return SaveProjectionAsync(source, options, (conversion, settings, timer) => settings.FailOnLoss
            ? PdfConversionEvidence.SaveLosslessAsync(conversion, source.Kind.ToString(), path, timer, cancellationToken)
            : conversion.SaveAsync(path, cancellationToken), cancellationToken);
    }

    /// <summary>Writes normalized Reader content to a caller-owned stream asynchronously.</summary>
    public static Task<OfficeConversionReport> SaveAsPdfAsync(this OfficeDocumentReadResult source, Stream stream,
        PdfProjectionOptions? options = null, CancellationToken cancellationToken = default) {
        PdfConversionEvidence.ValidateDestination(stream);
        return SaveProjectionAsync(source, options, (conversion, settings, timer) => settings.FailOnLoss
            ? PdfConversionEvidence.SaveLosslessAsync(conversion, source.Kind.ToString(), stream, timer, cancellationToken)
            : conversion.SaveAsync(stream, cancellationToken), cancellationToken);
    }

    private static OfficeConversionReport SaveProjection(OfficeDocumentReadResult source, PdfProjectionOptions? options,
        Func<PdfDocumentConversionResult, PdfProjectionOptions, Stopwatch, PdfSaveResult> save, CancellationToken token) {
        var timer = Stopwatch.StartNew();
        PdfProjectionOptions settings = (options ?? new PdfProjectionOptions()).Snapshot();
        PdfDocumentConversionResult conversion = source.ToPdfDocumentResult(settings, token);
        using PdfDocument model = conversion.Value;
        PdfSaveResult saved = PdfConversionEvidence.RunOutput(conversion, source.Kind.ToString(), () => save(conversion, settings, timer));
        return PdfConversionEvidence.CreateReport(source.Kind.ToString(), saved.ConversionReports, saved.OutputPath,
            timer.Elapsed, saved.Pipeline.Output?.PageCount ?? 0);
    }

    private static async Task<OfficeConversionReport> SaveProjectionAsync(OfficeDocumentReadResult source, PdfProjectionOptions? options,
        Func<PdfDocumentConversionResult, PdfProjectionOptions, Stopwatch, Task<PdfSaveResult>> save, CancellationToken token) {
        var timer = Stopwatch.StartNew();
        PdfProjectionOptions settings = (options ?? new PdfProjectionOptions()).Snapshot();
        PdfDocumentConversionResult conversion = await Task.Run(() => source.ToPdfDocumentResult(settings, token), token).ConfigureAwait(false);
        using PdfDocument model = conversion.Value;
        PdfSaveResult saved = await PdfConversionEvidence.RunOutputAsync(conversion, source.Kind.ToString(), () => save(conversion, settings, timer)).ConfigureAwait(false);
        return PdfConversionEvidence.CreateReport(source.Kind.ToString(), saved.ConversionReports, saved.OutputPath,
            timer.Elapsed, saved.Pipeline.Output?.PageCount ?? 0);
    }
}
