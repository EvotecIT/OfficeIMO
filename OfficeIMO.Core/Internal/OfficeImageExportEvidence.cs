using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Drawing;

/// <summary>Freezes renderer diagnostics at the public conversion boundary without inferring severity from loss.</summary>
internal static class OfficeImageExportEvidence {
    internal static OfficeConversionReport CreateReport(string source, OfficeImageExportResult result, TimeSpan elapsed = default) =>
        CreateReport(source, result.Format, result.Diagnostics, 1, elapsed, result.SavedPath);

    internal static OfficeConversionReport CreateReport(string source, OfficeImageExportFormat format,
        IEnumerable<OfficeImageExportDiagnostic> diagnostics, int itemsWritten, TimeSpan elapsed = default, string? outputPath = null) =>
        new OfficeConversionReport(source, format.ToString(), outputPath, itemsWritten, elapsed,
            diagnostics.Select(diagnostic => new OfficeDiagnostic(diagnostic.Code, diagnostic.Severity switch {
                OfficeImageExportDiagnosticSeverity.Error => OfficeDiagnosticSeverity.Error,
                OfficeImageExportDiagnosticSeverity.Warning => OfficeDiagnosticSeverity.Warning,
                _ => OfficeDiagnosticSeverity.Information
            }, diagnostic.Message, diagnostic.FidelityLocation ?? diagnostic.Source,
                diagnostic.FidelitySource ?? source, diagnostic.LossKind)));
}
