using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO;

/// <summary>Shared composition helpers for common conversion diagnostics.</summary>
internal static class OfficeConversionFidelityDiagnostics {
    /// <summary>Flattens stage diagnostics without collapsing their loss categories.</summary>
    public static IReadOnlyList<OfficeDiagnostic> Flatten(IEnumerable<OfficeConversionReport> reports) {
        ArgumentNullException.ThrowIfNull(reports);
        var diagnostics = new List<OfficeDiagnostic>();
        foreach (OfficeConversionReport report in reports) {
            if (report == null) throw new ArgumentException("Conversion reports cannot contain null entries.", nameof(reports));
            IReadOnlyList<OfficeDiagnostic> stageDiagnostics = report.Diagnostics;
            diagnostics.AddRange(stageDiagnostics);
            if (report.HasLoss && !stageDiagnostics.Any(static diagnostic =>
                    diagnostic.LossKind != OfficeConversionLossKind.None)) {
                string source = report.GetType().FullName ?? report.GetType().Name;
                diagnostics.Add(new OfficeDiagnostic(
                    global::OfficeIMO.OfficeDiagnosticCodes.ConversionReportUntypedLoss,
                    OfficeDiagnosticSeverity.Warning,
                    "The conversion stage reported aggregate loss without a typed fidelity diagnostic.",
                    null,
                    source,
                    OfficeConversionLossKind.Failure));
            }
        }
        return Array.AsReadOnly(diagnostics.ToArray());
    }

    /// <summary>Creates a diagnostic from the shared native-format projection contract.</summary>
    public static OfficeDiagnostic From(OfficeConversionDiagnostic diagnostic, string source) {
        ArgumentNullException.ThrowIfNull(diagnostic);
        return new OfficeDiagnostic(
            diagnostic.Code,
            diagnostic.Severity switch {
                OfficeConversionDiagnosticSeverity.Error => OfficeDiagnosticSeverity.Error,
                OfficeConversionDiagnosticSeverity.Warning => OfficeDiagnosticSeverity.Warning,
                _ => OfficeDiagnosticSeverity.Information
            },
            diagnostic.Message,
            diagnostic.SourceLocation,
            source,
            diagnostic.LossKind);
    }

    /// <summary>Creates a diagnostic from a shared compatibility finding.</summary>
    public static OfficeDiagnostic From(OfficeCompatibilityFinding finding, string source) {
        ArgumentNullException.ThrowIfNull(finding);
        return new OfficeDiagnostic(
            finding.Code,
            finding.LossKind == OfficeConversionLossKind.None ? OfficeDiagnosticSeverity.Information : OfficeDiagnosticSeverity.Warning,
            string.IsNullOrEmpty(finding.Message) ? finding.Category : finding.Message,
            finding.SourceLocation,
            source,
            finding.LossKind);
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
