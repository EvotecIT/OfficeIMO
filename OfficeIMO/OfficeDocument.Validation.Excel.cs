using OfficeIMO.Excel;

namespace OfficeIMO;

public static partial class OfficeDocument {
    private static void ValidateExcel(ExcelDocument document, OfficeExcelValidationOptions? options,
        List<OfficeDiagnostic> diagnostics, List<object> evidence, CancellationToken cancellationToken) {
        var report = document.RunWorkbookDoctor(options?.Doctor);
        evidence.Add(report);
        foreach (var issue in report.Issues) diagnostics.Add(new OfficeDiagnostic("Excel." + issue.Category,
            (OfficeDiagnosticSeverity)(int)issue.Severity, issue.Message + (options?.IncludeRepairHints == true && issue.RepairAction != null
                ? " Repair hint: " + issue.RepairAction : string.Empty), issue.SheetName + ":" + issue.Address, "Excel"));
        if (options == null) return;
        cancellationToken.ThrowIfCancellationRequested();
        var features = document.InspectFeatures();
        if (options.IncludeFeatures) {
            evidence.Add(features);
            foreach (var feature in features.Features) diagnostics.Add(new OfficeDiagnostic("Excel.Feature", OfficeDiagnosticSeverity.Information,
                feature.Name + " (" + feature.SupportLevel + ", count=" + feature.Count + "): " + feature.Note + " " + string.Join("; ", feature.Details),
                feature.Scope, "Excel"));
        }
        var capabilities = options.Capability ?? global::OfficeIMO.Internal.EnumCompat.GetValues<ExcelPreflightCapability>();
        foreach (var capability in capabilities.Distinct()) {
            cancellationToken.ThrowIfCancellationRequested();
            bool available = features.Can(capability);
            var details = features.GetCapabilityDiagnostics(capability);
            diagnostics.Add(new OfficeDiagnostic("Excel.Capability." + capability,
                available ? OfficeDiagnosticSeverity.Information : OfficeDiagnosticSeverity.Error,
                available ? capability + " is available." : capability + ": " + string.Join("; ", details), format: "Excel"));
            if (options.IncludeRepairHints) {
                foreach (var hint in features.GetRepairHints(capability)) {
                    evidence.Add(hint);
                    diagnostics.Add(new OfficeDiagnostic("Excel.RepairHint." + capability, OfficeDiagnosticSeverity.Information,
                        hint.Action + " " + hint.Command + " " + hint.Details, hint.FeatureName, "Excel"));
                }
            }
        }
    }
}
