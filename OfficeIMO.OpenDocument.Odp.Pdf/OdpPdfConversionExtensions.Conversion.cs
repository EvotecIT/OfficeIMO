using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Pdf;
using PdfPowerPoint = OfficeIMO.Pdf.PowerPoint;
using OfficeIMO.PowerPoint.OpenDocument;
using PowerPointPdf = OfficeIMO.PowerPoint.Pdf;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.OpenDocument.Odp.Pdf;

public static partial class OdpPdfConversionExtensions {
    private const string Source = "Odp";

    internal static PdfCore.PdfDocumentConversionResult ConvertToPdfWithEvidence(this OdpPresentation source, OdpPdfExportOptions? options,
        CancellationToken cancellationToken) =>
        ConvertToPdfCore(source, options?.ConversionOptions, options?.PdfOptions, cancellationToken);

    /// <summary>Converts the OpenDocument presentation to a PDF model.</summary>
    public static PdfDocument ToPdf(this OdpPresentation source, OdpPdfExportOptions? options = null,
        CancellationToken cancellationToken = default) =>
        source.ToPdf(options, out _, cancellationToken);

    /// <summary>Converts the OpenDocument presentation and returns its completed conversion findings.</summary>
    public static PdfDocument ToPdf(this OdpPresentation source, OdpPdfExportOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = source.ConvertToPdfWithEvidence(options, cancellationToken);
        return PdfConversionEvidence.CompleteModel(converted, Source, options?.FailOnLoss == true,
            timer, cancellationToken, out report);
    }

    /// <summary>Converts the OpenDocument presentation, returning the model and report together.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool so the caller is not blocked.</remarks>
    public static async Task<OfficeConversionResult<PdfDocument>> ToPdfAsync(this OdpPresentation source,
        OdpPdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        options = (options ?? new OdpPdfExportOptions()).Clone();
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = await Task.Run(
            () => source.ConvertToPdfWithEvidence(options, cancellationToken), cancellationToken).ConfigureAwait(false);
        PdfDocument value = PdfConversionEvidence.CompleteModel(converted, Source, options.FailOnLoss,
            timer, cancellationToken, out OfficeConversionReport report);
        return new OfficeConversionResult<PdfDocument>(value, report);
    }

    /// <summary>Writes the OpenDocument presentation as a PDF file and returns the complete conversion report.</summary>
    public static OfficeConversionReport SaveAsPdf(this OdpPresentation source, string path,
        OdpPdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfConversionEvidence.ValidateDestination(path);
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = source.ConvertToPdfWithEvidence(options, cancellationToken);
        using PdfDocument model = converted.Value;
        PdfSaveResult saved = options?.FailOnLoss == true
            ? PdfConversionEvidence.SaveLossless(converted, Source, path, timer, cancellationToken)
            : PdfConversionEvidence.RunOutput(converted, Source, () => converted.Save(path, cancellationToken));
        return PdfConversionEvidence.CreateReport(Source, saved.ConversionReports, saved.OutputPath, timer.Elapsed, saved.Pipeline.Output?.PageCount ?? 0);
    }

    /// <summary>Writes the OpenDocument presentation as PDF to a caller-owned stream and returns the complete conversion report.</summary>
    public static OfficeConversionReport SaveAsPdf(this OdpPresentation source, Stream stream,
        OdpPdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfConversionEvidence.ValidateDestination(stream);
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = source.ConvertToPdfWithEvidence(options, cancellationToken);
        using PdfDocument model = converted.Value;
        PdfSaveResult saved = options?.FailOnLoss == true
            ? PdfConversionEvidence.SaveLossless(converted, Source, stream, timer, cancellationToken)
            : PdfConversionEvidence.RunOutput(converted, Source, () => converted.Save(stream, cancellationToken));
        return PdfConversionEvidence.CreateReport(Source, saved.ConversionReports, elapsed: timer.Elapsed, itemsWritten: saved.Pipeline.Output?.PageCount ?? 0);
    }

    /// <summary>Asynchronously writes the OpenDocument presentation as a PDF file and returns the complete conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool so the caller is not blocked. The file is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsPdfAsync(this OdpPresentation source, string path,
        OdpPdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfConversionEvidence.ValidateDestination(path);
        options = (options ?? new OdpPdfExportOptions()).Clone();
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = await Task.Run(
            () => source.ConvertToPdfWithEvidence(options, cancellationToken), cancellationToken).ConfigureAwait(false);
        using PdfDocument model = converted.Value;
        PdfSaveResult saved = options.FailOnLoss
            ? await PdfConversionEvidence.SaveLosslessAsync(converted, Source, path, timer, cancellationToken).ConfigureAwait(false)
            : await PdfConversionEvidence.RunOutputAsync(converted, Source, () => converted.SaveAsync(path, cancellationToken)).ConfigureAwait(false);
        return PdfConversionEvidence.CreateReport(Source, saved.ConversionReports, saved.OutputPath, timer.Elapsed, saved.Pipeline.Output?.PageCount ?? 0);
    }

    /// <summary>Asynchronously writes the OpenDocument presentation as PDF to a caller-owned stream and returns the complete conversion report.</summary>
    /// <remarks>Conversion is CPU-bound and runs on the thread pool so the caller is not blocked. The stream is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsPdfAsync(this OdpPresentation source, Stream stream,
        OdpPdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfConversionEvidence.ValidateDestination(stream);
        options = (options ?? new OdpPdfExportOptions()).Clone();
        var timer = Stopwatch.StartNew();
        PdfDocumentConversionResult converted = await Task.Run(
            () => source.ConvertToPdfWithEvidence(options, cancellationToken), cancellationToken).ConfigureAwait(false);
        using PdfDocument model = converted.Value;
        PdfSaveResult saved = options.FailOnLoss
            ? await PdfConversionEvidence.SaveLosslessAsync(converted, Source, stream, timer, cancellationToken).ConfigureAwait(false)
            : await PdfConversionEvidence.RunOutputAsync(converted, Source, () => converted.SaveAsync(stream, cancellationToken)).ConfigureAwait(false);
        return PdfConversionEvidence.CreateReport(Source, saved.ConversionReports, elapsed: timer.Elapsed, itemsWritten: saved.Pipeline.Output?.PageCount ?? 0);
    }

    internal static PdfOdpConversionResult ConvertToOdpWithEvidence(this PdfCore.PdfDocument source, PdfOdpExportOptions? options,
        CancellationToken cancellationToken) =>
        ConvertOdpCore(source, options?.PdfOptions, options?.OpenDocumentOptions, cancellationToken);

    /// <summary>Reconstructs an Odp document from an opened PDF.</summary>
    public static OdpPresentation ToOdp(this PdfCore.PdfDocument source, PdfOdpExportOptions? options = null,
        CancellationToken cancellationToken = default) =>
        source.ToOdp(options, out _, cancellationToken);

    /// <summary>Reconstructs an Odp document from an opened PDF and returns the report of both stages.</summary>
    /// <remarks>The report is a <see cref="PdfOdpConversionReport"/> with the PDF and OpenDocument stage reports.</remarks>
    public static OdpPresentation ToOdp(this PdfCore.PdfDocument source, PdfOdpExportOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        PdfOdpConversionResult result = source.ConvertToOdpWithEvidence(options, cancellationToken);
        report = result.Report;
        try {
            if (options?.FailOnLoss == true) report.ThrowIfLoss();
            return result.Value;
        } catch { result.Value.Dispose(); throw; }
    }

    /// <summary>Reconstructs an Odp document from an opened PDF, returning the model and report together.</summary>
    /// <remarks>Reconstruction is CPU-bound and runs on the thread pool so the caller is not blocked.</remarks>
    public static async Task<OfficeConversionResult<OdpPresentation>> ToOdpAsync(this PdfCore.PdfDocument source,
        PdfOdpExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        options = (options ?? new PdfOdpExportOptions()).Clone();
        return await Task.Run(() => {
            OdpPresentation value = source.ToOdp(options, out OfficeConversionReport report, cancellationToken);
            return new OfficeConversionResult<OdpPresentation>(value, report);
        }, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Reconstructs and saves an Odp file from an opened PDF.</summary>
    public static OfficeConversionReport SaveAsOdp(this PdfCore.PdfDocument source, string path,
        PdfOdpExportOptions? options = null, CancellationToken cancellationToken = default) {
        var timer = Stopwatch.StartNew();
        using OdpPresentation value = source.ToOdp(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        global::OfficeIMO.Core.Internal.OfficeConversionFailures.Run(report, () => value.Save(path));
        return new OfficeConversionReport("Pdf", "Odp", Path.GetFullPath(path), 0, timer.Elapsed, report.Diagnostics);
    }

    /// <summary>Reconstructs an Odp document from an opened PDF into a caller-owned stream.</summary>
    public static OfficeConversionReport SaveAsOdp(this PdfCore.PdfDocument source, Stream stream,
        PdfOdpExportOptions? options = null, CancellationToken cancellationToken = default) {
        var timer = Stopwatch.StartNew();
        using OdpPresentation value = source.ToOdp(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        global::OfficeIMO.Core.Internal.OfficeConversionFailures.Run(report, () => value.Save(stream));
        return new OfficeConversionReport("Pdf", "Odp", elapsed: timer.Elapsed, diagnostics: report.Diagnostics);
    }

    /// <summary>Asynchronously reconstructs and saves an Odp file from an opened PDF.</summary>
    /// <remarks>Reconstruction is CPU-bound and runs on the thread pool; the file is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsOdpAsync(this PdfCore.PdfDocument source, string path,
        PdfOdpExportOptions? options = null, CancellationToken cancellationToken = default) {
        var timer = Stopwatch.StartNew();
        OfficeConversionResult<OdpPresentation> converted = await source.ToOdpAsync(options, cancellationToken).ConfigureAwait(false);
        using OdpPresentation value = converted.Value;
        await global::OfficeIMO.Core.Internal.OfficeConversionFailures.RunAsync(converted.Report,
            () => value.SaveAsync(path, cancellationToken: cancellationToken)).ConfigureAwait(false);
        return new OfficeConversionReport("Pdf", "Odp", Path.GetFullPath(path), 0, timer.Elapsed, converted.Report.Diagnostics);
    }

    /// <summary>Asynchronously reconstructs an Odp document from an opened PDF into a caller-owned stream.</summary>
    /// <remarks>Reconstruction is CPU-bound and runs on the thread pool; the stream is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsOdpAsync(this PdfCore.PdfDocument source, Stream stream,
        PdfOdpExportOptions? options = null, CancellationToken cancellationToken = default) {
        var timer = Stopwatch.StartNew();
        OfficeConversionResult<OdpPresentation> converted = await source.ToOdpAsync(options, cancellationToken).ConfigureAwait(false);
        using OdpPresentation value = converted.Value;
        await global::OfficeIMO.Core.Internal.OfficeConversionFailures.RunAsync(converted.Report,
            () => value.SaveAsync(stream, cancellationToken: cancellationToken)).ConfigureAwait(false);
        return new OfficeConversionReport("Pdf", "Odp", elapsed: timer.Elapsed, diagnostics: converted.Report.Diagnostics);
    }

    internal static PdfOdpConversionResult ConvertToOdpWithEvidence(this PdfCore.PdfDocumentReadResult source, PdfOdpExportOptions? options,
        CancellationToken cancellationToken) =>
        ConvertOdpCore(source, options?.PdfOptions, options?.OpenDocumentOptions, cancellationToken);

    /// <summary>Reconstructs an Odp document from a logical PDF model.</summary>
    public static OdpPresentation ToOdp(this PdfCore.PdfDocumentReadResult source, PdfOdpExportOptions? options = null,
        CancellationToken cancellationToken = default) =>
        source.ToOdp(options, out _, cancellationToken);

    /// <summary>Reconstructs an Odp document from a logical PDF model and returns the report of both stages.</summary>
    /// <remarks>The report is a <see cref="PdfOdpConversionReport"/> with the PDF and OpenDocument stage reports.</remarks>
    public static OdpPresentation ToOdp(this PdfCore.PdfDocumentReadResult source, PdfOdpExportOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        PdfOdpConversionResult result = source.ConvertToOdpWithEvidence(options, cancellationToken);
        report = result.Report;
        try {
            if (options?.FailOnLoss == true) report.ThrowIfLoss();
            return result.Value;
        } catch { result.Value.Dispose(); throw; }
    }

    /// <summary>Reconstructs an Odp document from a logical PDF model, returning the model and report together.</summary>
    /// <remarks>Reconstruction is CPU-bound and runs on the thread pool so the caller is not blocked.</remarks>
    public static async Task<OfficeConversionResult<OdpPresentation>> ToOdpAsync(this PdfCore.PdfDocumentReadResult source,
        PdfOdpExportOptions? options = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        options = (options ?? new PdfOdpExportOptions()).Clone();
        return await Task.Run(() => {
            OdpPresentation value = source.ToOdp(options, out OfficeConversionReport report, cancellationToken);
            return new OfficeConversionResult<OdpPresentation>(value, report);
        }, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Reconstructs and saves an Odp file from a logical PDF model.</summary>
    public static OfficeConversionReport SaveAsOdp(this PdfCore.PdfDocumentReadResult source, string path,
        PdfOdpExportOptions? options = null, CancellationToken cancellationToken = default) {
        var timer = Stopwatch.StartNew();
        using OdpPresentation value = source.ToOdp(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        global::OfficeIMO.Core.Internal.OfficeConversionFailures.Run(report, () => value.Save(path));
        return new OfficeConversionReport("Pdf", "Odp", Path.GetFullPath(path), 0, timer.Elapsed, report.Diagnostics);
    }

    /// <summary>Reconstructs an Odp document from a logical PDF model into a caller-owned stream.</summary>
    public static OfficeConversionReport SaveAsOdp(this PdfCore.PdfDocumentReadResult source, Stream stream,
        PdfOdpExportOptions? options = null, CancellationToken cancellationToken = default) {
        var timer = Stopwatch.StartNew();
        using OdpPresentation value = source.ToOdp(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        global::OfficeIMO.Core.Internal.OfficeConversionFailures.Run(report, () => value.Save(stream));
        return new OfficeConversionReport("Pdf", "Odp", elapsed: timer.Elapsed, diagnostics: report.Diagnostics);
    }

    /// <summary>Asynchronously reconstructs and saves an Odp file from a logical PDF model.</summary>
    /// <remarks>Reconstruction is CPU-bound and runs on the thread pool; the file is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsOdpAsync(this PdfCore.PdfDocumentReadResult source, string path,
        PdfOdpExportOptions? options = null, CancellationToken cancellationToken = default) {
        var timer = Stopwatch.StartNew();
        OfficeConversionResult<OdpPresentation> converted = await source.ToOdpAsync(options, cancellationToken).ConfigureAwait(false);
        using OdpPresentation value = converted.Value;
        await global::OfficeIMO.Core.Internal.OfficeConversionFailures.RunAsync(converted.Report,
            () => value.SaveAsync(path, cancellationToken: cancellationToken)).ConfigureAwait(false);
        return new OfficeConversionReport("Pdf", "Odp", Path.GetFullPath(path), 0, timer.Elapsed, converted.Report.Diagnostics);
    }

    /// <summary>Asynchronously reconstructs an Odp document from a logical PDF model into a caller-owned stream.</summary>
    /// <remarks>Reconstruction is CPU-bound and runs on the thread pool; the stream is written asynchronously.</remarks>
    public static async Task<OfficeConversionReport> SaveAsOdpAsync(this PdfCore.PdfDocumentReadResult source, Stream stream,
        PdfOdpExportOptions? options = null, CancellationToken cancellationToken = default) {
        var timer = Stopwatch.StartNew();
        OfficeConversionResult<OdpPresentation> converted = await source.ToOdpAsync(options, cancellationToken).ConfigureAwait(false);
        using OdpPresentation value = converted.Value;
        await global::OfficeIMO.Core.Internal.OfficeConversionFailures.RunAsync(converted.Report,
            () => value.SaveAsync(stream, cancellationToken: cancellationToken)).ConfigureAwait(false);
        return new OfficeConversionReport("Pdf", "Odp", elapsed: timer.Elapsed, diagnostics: converted.Report.Diagnostics);
    }
}
