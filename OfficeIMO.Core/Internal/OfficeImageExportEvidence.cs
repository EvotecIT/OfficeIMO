using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Drawing;

/// <summary>Freezes renderer diagnostics at the public conversion boundary without inferring severity from loss.</summary>
internal static class OfficeImageExportEvidence {
    internal static OfficeConversionReport CreateReport(string source, OfficeImageExportResult result, TimeSpan elapsed = default) =>
        new OfficeConversionReport(source, result.Format.ToString(), result.SavedPath, 1, elapsed,
            result.Diagnostics.Select(diagnostic => ProjectDiagnostic(result.Source ?? source, diagnostic)));

    internal static OfficeConversionReport CreateReport(string source, OfficeImageExportFormat format,
        IEnumerable<OfficeImageExportDiagnostic> diagnostics, int itemsWritten, TimeSpan elapsed = default, string? outputPath = null) =>
        new OfficeConversionReport(source, format.ToString(), outputPath, itemsWritten, elapsed,
            diagnostics.Select(diagnostic => ProjectDiagnostic(source, diagnostic)));

    /// <summary>Preserves severity and fidelity independently when adapting a renderer finding.</summary>
    internal static OfficeDiagnostic ProjectDiagnostic(string source, OfficeImageExportDiagnostic diagnostic) =>
        new OfficeDiagnostic(diagnostic.Code, diagnostic.Severity switch {
                OfficeImageExportDiagnosticSeverity.Error => OfficeDiagnosticSeverity.Error,
                OfficeImageExportDiagnosticSeverity.Warning => OfficeDiagnosticSeverity.Warning,
                _ => OfficeDiagnosticSeverity.Information
            }, diagnostic.Message, diagnostic.FidelityLocation ?? diagnostic.Source,
                diagnostic.FidelitySource ?? source, diagnostic.LossKind);
    /// <summary>Retains the result origin when a saved batch releases its payload and aggregates findings.</summary>
    internal static OfficeImageExportDiagnostic CaptureOrigin(OfficeImageExportDiagnostic diagnostic, string? source) =>
        diagnostic.FidelitySource != null || source == null ? diagnostic : new OfficeImageExportDiagnostic(
            diagnostic.Severity, diagnostic.Code, diagnostic.Message, diagnostic.Source, diagnostic.LossKind,
            source, diagnostic.FidelityLocation ?? diagnostic.Source);

    /// <summary>Carries source-stage evidence into an encoded result without losing its provenance.</summary>
    internal static OfficeImageExportResult AttachSourceReport(OfficeImageExportResult result, OfficeConversionReport report) =>
        result.WithDiagnostics(report.Diagnostics.Select(MapSourceDiagnostic).Concat(result.Diagnostics).ToArray());

    internal static OfficeImageExportDiagnostic MapSourceDiagnostic(OfficeDiagnostic diagnostic) =>
        new OfficeImageExportDiagnostic(diagnostic.Severity switch {
                OfficeDiagnosticSeverity.Error => OfficeImageExportDiagnosticSeverity.Error,
                OfficeDiagnosticSeverity.Warning => OfficeImageExportDiagnosticSeverity.Warning,
                _ => OfficeImageExportDiagnosticSeverity.Info
            }, diagnostic.Code, diagnostic.Message, diagnostic.Location ?? diagnostic.Format, diagnostic.LossKind,
            diagnostic.Format, diagnostic.Location);
}
