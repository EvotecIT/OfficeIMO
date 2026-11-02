using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Pdf;

internal static partial class PdfConversionEvidence {
    internal static PdfSaveResult SaveLossless(PdfDocumentConversionResult conversion, string source,
        string path, Stopwatch timer, CancellationToken token) {
        string fullPath = PdfDocument.ValidateOutputPath(path);
        var staged = Stage(conversion, source, saved => AcceptLossless(conversion, source, saved, timer, fullPath), token);
        token.ThrowIfCancellationRequested();
        OfficeConversionFailures.Run(CreateReport(source, staged.Result.ConversionReports),
            () => OfficeFileCommit.WriteAllBytes(fullPath, staged.Bytes));
        return staged.Result.WithOutputPath(fullPath);
    }

    internal static PdfSaveResult SaveLossless(PdfDocumentConversionResult conversion, string source,
        Stream stream, Stopwatch timer, CancellationToken token) {
        ValidateDestination(stream);
        var staged = Stage(conversion, source, saved => AcceptLossless(conversion, source, saved, timer), token);
        token.ThrowIfCancellationRequested();
        OfficeConversionFailures.Run(CreateReport(source, staged.Result.ConversionReports),
            () => OfficeStreamWriter.WriteAllBytes(stream, staged.Bytes));
        return staged.Result;
    }

    internal static async Task<PdfSaveResult> SaveLosslessAsync(PdfDocumentConversionResult conversion, string source,
        string path, Stopwatch timer, CancellationToken token) {
        string fullPath = PdfDocument.ValidateOutputPath(path);
        var staged = await StageAsync(conversion, source,
            saved => AcceptLossless(conversion, source, saved, timer, fullPath), token).ConfigureAwait(false);
        await OfficeConversionFailures.RunAsync(CreateReport(source, staged.Result.ConversionReports),
            () => OfficeFileCommit.WriteAllBytesAsync(fullPath, staged.Bytes, cancellationToken: token)).ConfigureAwait(false);
        return staged.Result.WithOutputPath(fullPath);
    }

    internal static async Task<PdfSaveResult> SaveLosslessAsync(PdfDocumentConversionResult conversion, string source,
        Stream stream, Stopwatch timer, CancellationToken token) {
        ValidateDestination(stream);
        var staged = await StageAsync(conversion, source,
            saved => AcceptLossless(conversion, source, saved, timer), token).ConfigureAwait(false);
        await OfficeConversionFailures.RunAsync(CreateReport(source, staged.Result.ConversionReports),
            () => OfficeStreamWriter.WriteAllBytesAsync(stream, staged.Bytes, token)).ConfigureAwait(false);
        return staged.Result;
    }

    private static (byte[] Bytes, PdfSaveResult Result) Stage(PdfDocumentConversionResult conversion, string source,
        Action<PdfSaveResult?> accept, CancellationToken token) {
        try { return conversion.StageLossless(token, accept); }
        catch (Exception error) when (error is not OperationCanceledException and not OfficeConversionException and not OutOfMemoryException and not StackOverflowException) {
            throw OutputFailure(conversion, source, error);
        }
    }

    private static async Task<(byte[] Bytes, PdfSaveResult Result)> StageAsync(PdfDocumentConversionResult conversion, string source,
        Action<PdfSaveResult?> accept, CancellationToken token) {
        try { return await conversion.StageLosslessAsync(token, accept).ConfigureAwait(false); }
        catch (Exception error) when (error is not OperationCanceledException and not OfficeConversionException and not OutOfMemoryException and not StackOverflowException) {
            throw OutputFailure(conversion, source, error);
        }
    }

    internal static PdfSaveResult RunOutput(PdfDocumentConversionResult conversion, string source, Func<PdfSaveResult> save) {
        try { return save(); }
        catch (Exception error) when (error is not OperationCanceledException and not OfficeConversionException and not OutOfMemoryException and not StackOverflowException) {
            throw OutputFailure(conversion, source, error);
        }
    }

    internal static async Task<PdfSaveResult> RunOutputAsync(PdfDocumentConversionResult conversion, string source, Func<Task<PdfSaveResult>> save) {
        try { return await save().ConfigureAwait(false); }
        catch (Exception error) when (error is not OperationCanceledException and not OfficeConversionException and not OutOfMemoryException and not StackOverflowException) {
            throw OutputFailure(conversion, source, error);
        }
    }

    internal static OfficeConversionException OutputFailure(PdfDocumentConversionResult conversion, string source, Exception error) {
        PdfSaveResult failure = PdfSaveResult.FromFailure(null, error);
        OfficeConversionReport findings = CreateReport(source, conversion.ConversionReports.Concat(failure.ConversionReports));
        var report = new OfficeConversionReport(source, "Pdf", diagnostics: findings.Diagnostics.Concat(new[] {
            new OfficeDiagnostic(OfficeDiagnosticCodes.ConversionOutputFailed, OfficeDiagnosticSeverity.Error, error.Message,
                format: source, lossKind: OfficeConversionLossKind.Failure)
        }));
        return new OfficeConversionException(error.Message, report, error);
    }

    private static void AcceptLossless(PdfDocumentConversionResult conversion, string source,
        PdfSaveResult? saved, Stopwatch timer, string? path = null) =>
        CreateReport(source, saved?.ConversionReports ?? conversion.ConversionReports, path, timer.Elapsed,
            saved?.Pipeline.Output?.PageCount ?? 0).ThrowIfLoss();

    internal static void ValidateDestination(string path) => _ = PdfDocument.ValidateOutputPath(path);

    internal static void ValidateDestination(Stream stream) {
        Guard.NotNull(stream, nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(stream));
    }
}
