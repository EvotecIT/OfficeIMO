using OfficeIMO.Html;

namespace OfficeIMO.Excel.Html;

public static partial class HtmlExcelConverterExtensions {
    private const string ImportComponentName = "OfficeIMO.Excel.Html";

    private static void AddImportDiagnostic(
        HtmlToExcelResult result,
        string code,
        string message,
        HtmlDiagnosticSeverity severity = HtmlDiagnosticSeverity.Warning,
        HtmlConversionLossKind lossKind = HtmlConversionLossKind.None,
        string? source = null,
        string? detail = null) {
        result.Diagnostics.Add(ImportComponentName, code, message, severity, source, detail, lossKind);
    }
}
