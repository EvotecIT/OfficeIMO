using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Result returned by file and stream output operations.
/// </summary>
public sealed class PdfSaveResult : IOfficeOutputResult {
    private PdfSaveResult(
        string? outputPath,
        long bytesWritten,
        IReadOnlyList<string> diagnostics,
        Exception? exception,
        PdfConversionReport? report = null,
        PdfPipelineReport? pipeline = null,
        PdfSerializationReport? serialization = null,
        IReadOnlyList<IOfficeConversionReport>? sourceConversionReports = null) {
        OutputPath = outputPath;
        BytesWritten = bytesWritten;
        Diagnostics = diagnostics;
        Exception = exception;
        TextEncodingDiagnostics = PdfOutputDiagnostics.ExtractTextEncodingDiagnostics(exception);
        Report = Snapshot(report);
        IOfficeConversionReport[] sourceReports = SnapshotSourceReports(sourceConversionReports);
        ConversionReports = CreateConversionReports(sourceReports, report == null ? null : Report);
        Report.AddRange(PdfOutputDiagnostics.ToConversionWarnings(TextEncodingDiagnostics));
        Pipeline = pipeline ?? PdfPipelineReport.Empty();
        Serialization = serialization;
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

    /// <summary>Snapshot of PDF-stage structured output warnings for this save attempt.</summary>
    public PdfConversionReport Report { get; }

    /// <summary>Ordered source-stage reports followed by the PDF-stage report, when conversion preceded the save.</summary>
    public IReadOnlyList<IOfficeConversionReport> ConversionReports { get; }

    /// <summary>End-to-end create/open, mutation, and output evidence for this save attempt.</summary>
    public PdfPipelineReport Pipeline { get; }

    /// <summary>Bounded serialization evidence for a successful save, when available.</summary>
    public PdfSerializationReport? Serialization { get; }

    /// <summary>PDF-stage structured output warnings for this save attempt.</summary>
    public IReadOnlyList<PdfConversionWarning> Warnings => Report.Warnings;

    /// <summary>True when the PDF stage produced a structured output warning.</summary>
    public bool HasWarnings => Report.HasWarnings;

    /// <summary>True when any source or PDF conversion stage reported possible content loss.</summary>
    public bool HasLoss => ConversionReports.Any(static report => report.HasLoss);

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

    /// <summary>Requires a successful save without reported conversion loss.</summary>
    public PdfSaveResult RequireNoLoss() {
        RequireSuccess();
        foreach (IOfficeConversionReport report in ConversionReports) {
            report.RequireNoLoss();
        }
        return this;
    }

    /// <summary>Creates a successful save result.</summary>
    public static PdfSaveResult FromSuccess(string? outputPath, long bytesWritten) {
        return new PdfSaveResult(outputPath, bytesWritten, Array.Empty<string>(), null);
    }

    /// <summary>Creates a failed save result from an exception captured by a wrapper or adapter.</summary>
    public static PdfSaveResult FromFailure(string? outputPath, Exception exception) {
        Guard.NotNull(exception, nameof(exception));
        IReadOnlyList<string> diagnostics = PdfOutputDiagnostics.BuildExceptionDiagnostics(exception);
        return new PdfSaveResult(
            outputPath,
            0,
            diagnostics,
            exception,
            pipeline: PdfPipelineReport.FailedOutput("Save", exception));
    }

    internal static PdfSaveResult Success(
        string? outputPath,
        long bytesWritten,
        PdfPipelineReport? pipeline = null,
        PdfSerializationReport? serialization = null) {
        return new PdfSaveResult(outputPath, bytesWritten, Array.Empty<string>(), null, pipeline: pipeline, serialization: serialization);
    }

    internal static PdfSaveResult Failed(
        string? outputPath,
        Exception exception,
        PdfPipelineReport? pipeline = null) {
        Guard.NotNull(exception, nameof(exception));
        IReadOnlyList<string> diagnostics = PdfOutputDiagnostics.BuildExceptionDiagnostics(exception);
        return new PdfSaveResult(outputPath, 0, diagnostics, exception, pipeline: pipeline);
    }

    internal PdfSaveResult WithReport(
        PdfConversionReport report,
        IReadOnlyList<IOfficeConversionReport>? sourceConversionReports = null) {
        return new PdfSaveResult(
            OutputPath,
            BytesWritten,
            Diagnostics,
            Exception,
            report,
            Pipeline,
            Serialization,
            sourceConversionReports);
    }

    private static PdfConversionReport Snapshot(PdfConversionReport? report) {
        var snapshot = new PdfConversionReport();
        if (report != null) snapshot.AddRange(report.Warnings);
        return snapshot;
    }

    private static IOfficeConversionReport[] SnapshotSourceReports(
        IReadOnlyList<IOfficeConversionReport>? reports) {
        if (reports == null || reports.Count == 0) return Array.Empty<IOfficeConversionReport>();
        var snapshot = new IOfficeConversionReport[reports.Count];
        for (int i = 0; i < reports.Count; i++) {
            snapshot[i] = reports[i];
        }
        return snapshot;
    }

    private static IOfficeConversionReport[] CreateConversionReports(
        IOfficeConversionReport[] sourceReports,
        PdfConversionReport? pdfReport) {
        int count = sourceReports.Length + (pdfReport == null ? 0 : 1);
        if (count == 0) return Array.Empty<IOfficeConversionReport>();
        var reports = new IOfficeConversionReport[count];
        for (int i = 0; i < sourceReports.Length; i++) {
            reports[i] = sourceReports[i];
        }
        if (pdfReport != null) reports[reports.Length - 1] = pdfReport;
        return reports;
    }
}
