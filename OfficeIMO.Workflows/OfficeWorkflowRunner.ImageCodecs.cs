using OfficeIMO.Drawing;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    private static OfficeRasterDecodeInfo? FindCodecUnavailable(Exception exception) {
        for (Exception? current = exception; current != null; current = current.InnerException) {
            if (current is OfficeRasterDecodeException decoded &&
                decoded.DecodeInfo.CodecUnavailableReason != OfficeRasterCodecUnavailableReason.None)
                return decoded.DecodeInfo;
        }
        return null;
    }

    private static void AddImageCodecDiagnostic(Exception exception, ICollection<OfficeWorkflowDiagnostic> diagnostics,
        WorkflowFailureStage stage) {
        OfficeRasterDecodeInfo? info = FindCodecUnavailable(exception);
        if (info == null || diagnostics.Any(item => item.Code == info.DiagnosticCode)) return;
        diagnostics.Add(new OfficeWorkflowDiagnostic(
            info.DiagnosticCode!, info.Diagnostic ?? exception.Message, OfficeWorkflowDiagnosticSeverity.Error,
            GetDiagnosticStage(stage), new Dictionary<string, string>(StringComparer.Ordinal) {
                ["reason"] = info.CodecUnavailableReason.ToString(),
                ["format"] = info.Format.ToString(),
                ["source"] = info.Format.ToString(),
                ["lossKind"] = OfficeConversionLossKind.Failure.ToString()
            }));
    }
}
