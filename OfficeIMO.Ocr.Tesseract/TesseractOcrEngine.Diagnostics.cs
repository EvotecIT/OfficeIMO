using OfficeIMO.Ocr;
using OfficeIMO.Ocr.Process;

namespace OfficeIMO.Ocr.Tesseract;

public sealed partial class TesseractOcrEngine {
    // Native Tesseract writes routine resolution estimation to stderr. Unknown messages and
    // truncated logs remain warnings so genuine failures cannot become automatic acceptance.
    internal OcrDiagnostic? CreateStandardErrorDiagnostic(OcrProcessResult process, string code) {
        if (string.IsNullOrWhiteSpace(process.StandardError) && !process.StandardErrorTruncated) return null;
        string[] lines = process.StandardError.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        bool informational = !process.StandardErrorTruncated && lines.Length > 0 && lines.All(IsResolutionEstimate);
        return new OcrDiagnostic {
            Code = code, Severity = informational ? OcrDiagnosticSeverity.Info : OcrDiagnosticSeverity.Warning,
            Message = process.StandardError, Source = Id, IsRecoverable = true,
            Attributes = new Dictionary<string, string>(StringComparer.Ordinal) {
                ["truncated"] = process.StandardErrorTruncated ? "true" : "false"
            }
        };
    }

    private static bool IsResolutionEstimate(string line) {
        const string prefix = "Estimating resolution as ";
        line = line.Trim();
        return line.StartsWith(prefix, StringComparison.Ordinal)
            && int.TryParse(line.Substring(prefix.Length), NumberStyles.None, CultureInfo.InvariantCulture, out int dpi)
            && dpi > 0;
    }
}
