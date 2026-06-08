namespace OfficeIMO.Pdf;

internal static class PdfOutputDiagnostics {
    public static PdfTextEncodingDiagnostic[] ExtractTextEncodingDiagnostics(Exception? exception) {
        if (exception is PdfTextEncodingPreflightException preflightException) {
            return preflightException.TextEncodingDiagnostics.ToArray();
        }

        if (exception == null ||
            !TryGetStringData(exception, "code", out string code) ||
            (code != "unsupported-text-glyph" && code != "unsupported-control-character") ||
            !TryGetIntData(exception, "index", out int index) ||
            !TryGetStringData(exception, "codePoint", out string codePoint) ||
            !TryGetBoolData(exception, "isControlCharacter", out bool isControlCharacter)) {
            return Array.Empty<PdfTextEncodingDiagnostic>();
        }

        string source = TryGetStringData(exception, "source", out string sourceValue) ? sourceValue : string.Empty;
        string text = TryGetStringData(exception, "text", out string textValue) ? textValue : string.Empty;
        string encoding = TryGetStringData(exception, "encoding", out string encodingValue) ? encodingValue : string.Empty;
        string remediation = TryGetStringData(exception, "remediation", out string remediationValue) ? remediationValue : string.Empty;
        string location = TryGetStringData(exception, "location", out string locationValue) ? locationValue : string.Empty;
        int? runIndex = TryGetIntData(exception, "runIndex", out int runIndexValue) ? runIndexValue : null;
        int? pageNumber = TryGetIntData(exception, "pageNumber", out int pageNumberValue) ? pageNumberValue : null;
        int? tableRowIndex = TryGetIntData(exception, "tableRowIndex", out int tableRowIndexValue) ? tableRowIndexValue : null;
        int? tableColumnIndex = TryGetIntData(exception, "tableColumnIndex", out int tableColumnIndexValue) ? tableColumnIndexValue : null;
        string fieldName = TryGetStringData(exception, "fieldName", out string fieldNameValue) ? fieldNameValue : string.Empty;

        return new[] {
            new PdfTextEncodingDiagnostic(source, index, codePoint, text, isControlCharacter, encoding, remediation, location, runIndex, pageNumber, tableRowIndex, tableColumnIndex, fieldName)
        };
    }

    public static PdfConversionWarning[] ToConversionWarnings(IEnumerable<PdfTextEncodingDiagnostic> diagnostics) {
        Guard.NotNull(diagnostics, nameof(diagnostics));
        return diagnostics
            .Select(diagnostic => diagnostic.ToConversionWarning())
            .ToArray();
    }

    public static IReadOnlyList<string> BuildExceptionDiagnostics(Exception exception) {
        Guard.NotNull(exception, nameof(exception));
        if (exception is PdfTextEncodingPreflightException preflightException) {
            var diagnostics = new List<string> { exception.Message };
            diagnostics.AddRange(preflightException.TextEncodingDiagnostics.Select(diagnostic => diagnostic.Message));
            return diagnostics.AsReadOnly();
        }

        return string.IsNullOrWhiteSpace(exception.Message)
            ? Array.Empty<string>()
            : new[] { exception.Message };
    }

    private static bool TryGetStringData(Exception exception, string key, out string value) {
        if (exception.Data.Contains(key) && exception.Data[key] is string text) {
            value = text;
            return true;
        }

        value = string.Empty;
        return false;
    }

    private static bool TryGetIntData(Exception exception, string key, out int value) {
        if (exception.Data.Contains(key) && exception.Data[key] is int number) {
            value = number;
            return true;
        }

        value = 0;
        return false;
    }

    private static bool TryGetBoolData(Exception exception, string key, out bool value) {
        if (exception.Data.Contains(key) && exception.Data[key] is bool flag) {
            value = flag;
            return true;
        }

        value = false;
        return false;
    }
}
