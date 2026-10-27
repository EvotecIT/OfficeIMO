using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Pdf;

/// <summary>Converts ordered image sources into PDF pages.</summary>
public static partial class PdfImageConverterExtensions {
    private static PdfDocument CompleteImageModel(Func<PdfDocument> create, PdfImageImportOptions options,
        CancellationToken token, out OfficeConversionReport report) {
        token.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        using var scope = new PdfOutputDiagnosticScope();
        var conversion = new PdfDocumentConversionResult(create(), scope.Report);
        return PdfConversionEvidence.CompleteModel(conversion, "Image", options.FailOnLoss, timer, token, out report);
    }

    private static OfficeConversionReport SaveImages(Func<PdfDocument> create,
        Func<PdfDocumentConversionResult, Stopwatch, PdfSaveResult> save, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        using var scope = new PdfOutputDiagnosticScope();
        using PdfDocument model = create();
        var conversion = new PdfDocumentConversionResult(model, scope.Report);
        PdfSaveResult saved = PdfConversionEvidence.RunOutput(conversion, "Image", () => save(conversion, timer));
        return PdfConversionEvidence.CreateReport("Image", saved.ConversionReports, saved.OutputPath,
            timer.Elapsed, saved.Pipeline.Output?.PageCount ?? 0);
    }

    private static async Task<OfficeConversionReport> SaveImagesAsync(Func<PdfDocument> create,
        Func<PdfDocumentConversionResult, Stopwatch, Task<PdfSaveResult>> save, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        using var scope = new PdfOutputDiagnosticScope();
        using PdfDocument model = await Task.Run(create, token).ConfigureAwait(false);
        var conversion = new PdfDocumentConversionResult(model, scope.Report);
        PdfSaveResult saved = await PdfConversionEvidence.RunOutputAsync(conversion, "Image", () => save(conversion, timer)).ConfigureAwait(false);
        return PdfConversionEvidence.CreateReport("Image", saved.ConversionReports, saved.OutputPath,
            timer.Elapsed, saved.Pipeline.Output?.PageCount ?? 0);
    }
}
