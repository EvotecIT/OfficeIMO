namespace OfficeIMO.Pdf;

/// <summary>
/// Result returned by file and stream output operations.
/// </summary>
public sealed class PdfSaveResult {
    private PdfSaveResult(
        string? outputPath,
        long bytesWritten,
        IReadOnlyList<string> diagnostics,
        Exception? exception,
        PdfConversionReport? report = null,
        PdfPipelineReport? pipeline = null) {
        OutputPath = outputPath;
        BytesWritten = bytesWritten;
        Diagnostics = diagnostics;
        Exception = exception;
        TextEncodingDiagnostics = PdfOutputDiagnostics.ExtractTextEncodingDiagnostics(exception);
        Report = Snapshot(report);
        Report.AddRange(PdfOutputDiagnostics.ToConversionWarnings(TextEncodingDiagnostics));
        Pipeline = pipeline ?? PdfPipelineReport.Empty();
    }

    /// <summary>True when the save operation completed.</summary>
    public bool Succeeded => Exception is null;

    /// <summary>Full output path when the operation targeted a file path.</summary>
    public string? OutputPath { get; }

    /// <summary>Number of bytes written when the operation completed; otherwise 0.</summary>
    public long BytesWritten { get; }

    /// <summary>Human-readable diagnostics explaining why the save failed.</summary>
    public IReadOnlyList<string> Diagnostics { get; }

    /// <summary>Exception captured from the save attempt, when available.</summary>
    public Exception? Exception { get; }

    /// <summary>Structured text encoding diagnostics captured from PDF generation failures.</summary>
    public IReadOnlyList<PdfTextEncodingDiagnostic> TextEncodingDiagnostics { get; }

    /// <summary>Snapshot of source-conversion and structured output warnings for this save attempt.</summary>
    public PdfConversionReport Report { get; }

    /// <summary>End-to-end create/open, mutation, and output evidence for this save attempt.</summary>
    public PdfPipelineReport Pipeline { get; }

    /// <summary>Source-conversion and structured output warnings for this save attempt.</summary>
    public IReadOnlyList<PdfConversionWarning> Warnings => Report.Warnings;

    /// <summary>True when source conversion or PDF output produced a warning.</summary>
    public bool HasWarnings => Report.HasWarnings;

    /// <summary>Returns this result or throws with diagnostics when the save failed.</summary>
    public PdfSaveResult RequireSuccess() {
        if (Succeeded) {
            return this;
        }

        string message = Diagnostics.Count == 0
            ? "PDF save did not complete."
            : "PDF save did not complete. " + string.Join(" ", Diagnostics);
        throw new InvalidOperationException(message, Exception);
    }

    /// <summary>Creates a successful save result.</summary>
    public static PdfSaveResult FromSuccess(string? outputPath, long bytesWritten) {
        return new PdfSaveResult(outputPath, bytesWritten, Array.Empty<string>(), null);
    }

    /// <summary>Creates a failed save result from an exception captured by a wrapper or adapter.</summary>
    public static PdfSaveResult FromFailure(string? outputPath, Exception exception) {
        Guard.NotNull(exception, nameof(exception));
        IReadOnlyList<string> diagnostics = PdfOutputDiagnostics.BuildExceptionDiagnostics(exception);
        return new PdfSaveResult(outputPath, 0, diagnostics, exception);
    }

    internal static PdfSaveResult Success(
        string? outputPath,
        long bytesWritten,
        PdfPipelineReport? pipeline = null) {
        return new PdfSaveResult(outputPath, bytesWritten, Array.Empty<string>(), null, pipeline: pipeline);
    }

    internal static PdfSaveResult Failed(
        string? outputPath,
        Exception exception,
        PdfPipelineReport? pipeline = null) {
        Guard.NotNull(exception, nameof(exception));
        IReadOnlyList<string> diagnostics = PdfOutputDiagnostics.BuildExceptionDiagnostics(exception);
        return new PdfSaveResult(outputPath, 0, diagnostics, exception, pipeline: pipeline);
    }

    internal PdfSaveResult WithReport(PdfConversionReport report) {
        return new PdfSaveResult(OutputPath, BytesWritten, Diagnostics, Exception, report, Pipeline);
    }

    private static PdfConversionReport Snapshot(PdfConversionReport? report) {
        var snapshot = new PdfConversionReport();
        if (report != null) snapshot.AddRange(report.Warnings);
        return snapshot;
    }
}
