using OfficeIMO.Drawing;
using System.Text;

namespace OfficeIMO.OpenDocument;

/// <summary>
/// Maps OpenDocument conversion evidence into the shared image-export diagnostic contract.
/// </summary>
public static class OdfImageExportDiagnostics {
    /// <summary>Returns only lossy conversion mappings as image-export diagnostics.</summary>
    public static IReadOnlyList<OfficeImageExportDiagnostic> FromConversionReport(
        OdfConversionReport report) {
        if (report == null) throw new ArgumentNullException(nameof(report));
        var diagnostics = new List<OfficeImageExportDiagnostic>();
        foreach (OdfConversionMapping mapping in report.Mappings) {
            if (mapping.Status == OdfConversionMappingStatus.Converted) continue;
            OfficeImageExportLossKind lossKind =
                mapping.Status == OdfConversionMappingStatus.Approximated
                    ? OfficeImageExportLossKind.Approximation
                    : OfficeImageExportLossKind.Omission;
            string message = string.IsNullOrWhiteSpace(mapping.Message)
                ? $"{mapping.Count} {mapping.Feature} item(s) were {mapping.Status.ToString().ToLowerInvariant()} " +
                  $"while converting {report.SourceFormat} to {report.TargetFormat} for image export."
                : mapping.Message!;
            diagnostics.Add(new OfficeImageExportDiagnostic(
                OfficeImageExportDiagnosticSeverity.Warning,
                CreateCode(mapping),
                message,
                $"{report.SourceFormat}->{report.TargetFormat}:{mapping.Feature}",
                lossKind));
        }

        return diagnostics.AsReadOnly();
    }

    /// <summary>Returns an image result with OpenDocument conversion loss attached.</summary>
    public static OfficeImageExportResult Attach(
        OfficeImageExportResult result,
        OdfConversionReport report) {
        if (result == null) throw new ArgumentNullException(nameof(result));
        IReadOnlyList<OfficeImageExportDiagnostic> conversionDiagnostics =
            FromConversionReport(report);
        if (conversionDiagnostics.Count == 0) return result;
        var diagnostics = new List<OfficeImageExportDiagnostic>(
            conversionDiagnostics.Count + result.Diagnostics.Count);
        diagnostics.AddRange(conversionDiagnostics);
        diagnostics.AddRange(result.Diagnostics);
        return new OfficeImageExportResult(
            result.Format,
            result.Width,
            result.Height,
            result.Bytes,
            result.Name,
            result.Source,
            diagnostics,
            result.SavedPath);
    }

    private static string CreateCode(OdfConversionMapping mapping) {
        var builder = new StringBuilder("ODF_IMAGE_");
        AppendToken(builder, mapping.Feature);
        builder.Append('_');
        AppendToken(builder, mapping.Status.ToString());
        return builder.ToString();
    }

    private static void AppendToken(StringBuilder builder, string value) {
        bool previousUnderscore = builder.Length > 0 && builder[builder.Length - 1] == '_';
        foreach (char character in value) {
            char normalized = char.IsLetterOrDigit(character)
                ? char.ToUpperInvariant(character)
                : '_';
            if (normalized == '_' && previousUnderscore) continue;
            builder.Append(normalized);
            previousUnderscore = normalized == '_';
        }
        while (builder.Length > 0 && builder[builder.Length - 1] == '_') {
            builder.Length--;
        }
    }
}
