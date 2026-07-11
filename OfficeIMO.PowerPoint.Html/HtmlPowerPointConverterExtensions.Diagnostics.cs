using OfficeIMO.Html;

namespace OfficeIMO.PowerPoint.Html;

public static partial class HtmlPowerPointConverterExtensions {
    private const string ImportComponentName = "OfficeIMO.PowerPoint.Html";

    private static void AddImportDiagnostic(
        HtmlToPowerPointResult result,
        string code,
        string message,
        HtmlDiagnosticSeverity severity = HtmlDiagnosticSeverity.Warning,
        HtmlConversionLossKind lossKind = HtmlConversionLossKind.None,
        string? source = null,
        string? detail = null) {
        result.Diagnostics.Add(ImportComponentName, code, message, severity, source, detail, lossKind);
    }
}
