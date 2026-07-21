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
        result.AddImportDiagnostic(new HtmlDiagnostic(ImportComponentName, code, message, severity, source, detail, lossKind));
    }

    private static bool IsSupportedExcelImage(
        HtmlImageDataUri dataUri,
        HtmlToExcelResult result,
        string? source = null) {
        if (ExcelSheet.IsSupportedImageContentType(dataUri.MediaType)) return true;
        AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ResourceTypeUnsupported,
            "An embedded worksheet image was omitted because its media type is not supported by Excel image parts.",
            lossKind: HtmlConversionLossKind.Omission,
            source: source,
            detail: "mediaType=" + dataUri.MediaType);
        return false;
    }
}
