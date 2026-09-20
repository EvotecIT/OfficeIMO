using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO;

/// <summary>
/// One common, category-preserving fidelity diagnostic emitted by a conversion stage.
/// </summary>
public sealed class OfficeConversionFidelityDiagnostic {
    /// <summary>Creates a fidelity diagnostic.</summary>
    public OfficeConversionFidelityDiagnostic(
        string code,
        string message,
        OfficeConversionLossKind lossKind,
        string source,
        string? location = null) {
        if (string.IsNullOrWhiteSpace(code)) throw new ArgumentException("Diagnostic code cannot be empty.", nameof(code));
        if (string.IsNullOrWhiteSpace(source)) throw new ArgumentException("Diagnostic source cannot be empty.", nameof(source));
        Code = code;
        Message = message ?? string.Empty;
        LossKind = lossKind;
        Source = source;
        Location = location;
    }

    /// <summary>Gets the stable diagnostic code.</summary>
    public string Code { get; }

    /// <summary>Gets the human-readable diagnostic message.</summary>
    public string Message { get; }

    /// <summary>Gets the exact fidelity-loss category.</summary>
    public OfficeConversionLossKind LossKind { get; }

    /// <summary>Gets the conversion stage or projection owner that emitted the diagnostic.</summary>
    public string Source { get; }

    /// <summary>Gets an optional source part, path, span, page, sheet, or record location.</summary>
    public string? Location { get; }
}

/// <summary>Shared composition helpers for typed conversion diagnostics.</summary>
public static class OfficeConversionFidelityDiagnostics {
    /// <summary>Flattens stage diagnostics without collapsing their loss categories.</summary>
    public static IReadOnlyList<OfficeConversionFidelityDiagnostic> Flatten(
        IEnumerable<IOfficeConversionReport> reports) {
        if (reports == null) throw new ArgumentNullException(nameof(reports));
        return Array.AsReadOnly(reports.SelectMany(static report => report.FidelityDiagnostics).ToArray());
    }

    /// <summary>Creates a diagnostic from the shared native-format projection contract.</summary>
    public static OfficeConversionFidelityDiagnostic From(
        OfficeConversionDiagnostic diagnostic,
        string source) {
        if (diagnostic == null) throw new ArgumentNullException(nameof(diagnostic));
        return new OfficeConversionFidelityDiagnostic(
            diagnostic.Code,
            diagnostic.Message,
            diagnostic.LossKind,
            source,
            diagnostic.SourceLocation);
    }

    /// <summary>Creates a diagnostic from a shared compatibility finding.</summary>
    public static OfficeConversionFidelityDiagnostic From(
        OfficeCompatibilityFinding finding,
        string source) {
        if (finding == null) throw new ArgumentNullException(nameof(finding));
        return new OfficeConversionFidelityDiagnostic(
            finding.Code,
            string.IsNullOrEmpty(finding.Message) ? finding.Category : finding.Message,
            finding.LossKind,
            source,
            finding.SourceLocation);
    }

    internal static OfficeConversionLossKind GetLossKind(
        OfficeCompatibilityState state,
        bool representsLoss) => state switch {
            OfficeCompatibilityState.Approximated => OfficeConversionLossKind.Approximation,
            OfficeCompatibilityState.Rasterized => OfficeConversionLossKind.Approximation,
            OfficeCompatibilityState.Dropped => OfficeConversionLossKind.Omission,
            OfficeCompatibilityState.Blocked => OfficeConversionLossKind.Failure,
            _ => representsLoss ? OfficeConversionLossKind.Approximation : OfficeConversionLossKind.None
        };
}
