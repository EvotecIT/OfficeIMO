using System;

namespace OfficeIMO;

/// <summary>Common artifact status for document output operations.</summary>
public interface IOfficeOutputResult : IOfficeResult {
    /// <summary>Destination path, or null for caller-owned streams.</summary>
    string? OutputPath { get; }
    /// <summary>Captured output failure, or null when output succeeded.</summary>
    Exception? Exception { get; }
}

/// <summary>Pairs output status with the typed fidelity evidence from a document conversion.</summary>
/// <typeparam name="TReport">Format-specific conversion report.</typeparam>
public sealed class OfficeOutputResult<TReport> : IOfficeOutputResult where TReport : class, IOfficeConversionReport {
    private OfficeOutputResult(string? outputPath, TReport? report, Exception? exception) {
        OutputPath = outputPath;
        Report = report;
        Exception = exception;
    }

    /// <inheritdoc />
    public bool Succeeded => Exception == null;
    /// <inheritdoc />
    public string? OutputPath { get; }
    /// <inheritdoc />
    public Exception? Exception { get; }
    /// <summary>Conversion evidence produced before writing, when available.</summary>
    public TReport? Report { get; }
    /// <summary>Whether available conversion evidence reports possible content loss.</summary>
    public bool HasLoss => Report?.HasLoss == true;

    /// <summary>Returns this result after verifying that an artifact was written successfully.</summary>
    public OfficeOutputResult<TReport> RequireSuccess() {
        if (!Succeeded) throw new InvalidOperationException("Document output did not complete.", Exception);
        return this;
    }

    /// <summary>Requires successful output and conversion evidence without reported loss.</summary>
    public OfficeOutputResult<TReport> RequireNoLoss() {
        RequireSuccess();
        if (Report == null) throw new InvalidOperationException("Document output has no conversion report.");
        Report.RequireNoLoss();
        return this;
    }

    /// <summary>Creates a successful output result after the native writer has completed.</summary>
    public static OfficeOutputResult<TReport> FromSuccess(string? outputPath, TReport report) =>
        new OfficeOutputResult<TReport>(outputPath, report ?? throw new ArgumentNullException(nameof(report)), null);

    /// <summary>Creates a failed output result, preserving any report completed before the failure.</summary>
    public static OfficeOutputResult<TReport> FromFailure(string? outputPath, Exception exception, TReport? report = null) =>
        new OfficeOutputResult<TReport>(outputPath, report, exception ?? throw new ArgumentNullException(nameof(exception)));
}
