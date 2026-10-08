using OfficeIMO.Epub;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    private static (byte[] Bytes, bool HasLoss) ExportBookProject(byte[] input, long maximumOutputBytes,
        List<OfficeWorkflowDiagnostic> diagnostics, CancellationToken token) {
        BookProject project = BookProject.LoadProject(input, token);
        EpubWriteResult result = project.Export(new EpubWriteOptions {
            MaxOutputBytes = Math.Min(maximumOutputBytes, 128L * 1024 * 1024)
        }, token);
        bool hasLoss = false;
        foreach (OfficeConversionFidelityDiagnostic finding in project.ImportDiagnostics.Concat(result.Report.FidelityDiagnostics)) {
            token.ThrowIfCancellationRequested();
            hasLoss |= finding.LossKind != OfficeConversionLossKind.None;
            var details = new Dictionary<string, string> {
                ["source"] = finding.Source, ["lossKind"] = finding.LossKind.ToString(),
                ["importLossAcknowledged"] = project.ImportLossAcknowledged.ToString()
            };
            if (finding.Location != null) details["location"] = finding.Location;
            diagnostics.Add(new OfficeWorkflowDiagnostic(finding.Code, finding.Message,
                finding.LossKind == OfficeConversionLossKind.None ? OfficeWorkflowDiagnosticSeverity.Information : OfficeWorkflowDiagnosticSeverity.Warning,
                "book-export", details));
        }
        return (result.Bytes, hasLoss);
    }
}
