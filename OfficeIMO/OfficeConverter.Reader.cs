using System.Diagnostics;
using OfficeIMO.Pdf;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Email;
using OfficeIMO.Reader.Epub;
using OfficeIMO.Reader.Pdf;
using OfficeIMO.Reader.Visio;

namespace OfficeIMO;

public static partial class OfficeConverter {
    internal static OfficeDocumentReader CaptureReader(OfficeConversionOptions options) => new OfficeDocumentReaderBuilder()
        .AddEmailHandler(options.ReaderEmail).AddEpubHandler(options.ReaderEpub).AddVisioHandler(options.ReaderVisio).Build();

    private static async Task<OfficeConversionReport> ConvertReaderPdfAsync(Stream source, Stream destination,
        OfficeConversionOptions options, CancellationToken token) {
        var timer = Stopwatch.StartNew();
        OfficeDocumentReader reader = options.ReaderPipeline ?? CaptureReader(options);
        ReaderOptions reading = DocumentReaderEngine.NormalizeOptions(options.Reader);
        reading.MaxInputBytes = Math.Min(reading.MaxInputBytes ?? long.MaxValue, options.MaximumInputBytes);
        reading.OpenPassword = options.SourcePassword ?? reading.OpenPassword;
        OfficeDocumentReadResult normalized = reader.ReadDocument(source, options.SourceName!, reading, token);
        PdfDocumentConversionResult converted = normalized.ToPdfDocumentResult(options.ReaderPdf, token);
        using PdfDocument model = converted.Value;
        PdfSaveResult saved = options.FailOnLoss || options.ReaderPdf?.FailOnLoss == true
            ? await PdfConversionEvidence.SaveLosslessAsync(converted, options.SourceName!, destination, timer, token).ConfigureAwait(false)
            : await PdfConversionEvidence.RunOutputAsync(converted, options.SourceName!, () => converted.SaveAsync(destination, token)).ConfigureAwait(false);
        return PdfConversionEvidence.CreateReport(options.SourceName!, saved.ConversionReports,
            elapsed: timer.Elapsed, itemsWritten: saved.Pipeline.Output?.PageCount ?? 0);
    }
}
