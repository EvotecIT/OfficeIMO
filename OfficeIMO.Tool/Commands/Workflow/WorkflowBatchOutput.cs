namespace OfficeIMO.Tool.Commands.Workflow;

/// <summary>Serializes shared batch outcomes without mixing human progress into stdout.</summary>
internal sealed class WorkflowBatchOutput(TextWriter output, TextWriter error, string? format) : IProgress<OfficeBatchItemReport> {
    private readonly object _gate = new();

    public void Report(OfficeBatchItemReport item) {
        lock (_gate) {
            if (format == "ndjson") output.WriteLine(OfficeBatchReportSerializer.SerializeItemEvent(item));
            if (item.Skipped || item.Status != OfficeBatchItemStatus.Completed) error.WriteLine(item.InputPath + ": " + item.Summary);
            foreach (OfficeDiagnostic finding in item.Diagnostics.Where(finding => finding.Severity != OfficeDiagnosticSeverity.Information))
                error.WriteLine(item.InputPath + ": " + finding.Code + (finding.Location == null ? "" : " [" + finding.Location + "]") + ": " + finding.Message);
        }
    }

    internal void Complete(OfficeBatchReport report, bool allowEmpty = false, string? recordingError = null) {
        lock (_gate) {
            OfficeBatchSummary result = report.Summary;
            bool succeeded = report.Succeeded && recordingError == null;
            if (recordingError != null) error.WriteLine("The conversions finished, but their run report could not be finalized: " + recordingError);
            if (result.FailureSummary != null) error.WriteLine("Batch stopped: " + result.FailureSummary);
            if (result.Selected == 0 && !allowEmpty && !result.Cancelled && result.FailureSummary == null)
                error.WriteLine("No files were selected for conversion. Check the inputs and executable route, or use --allow-empty for an intentionally empty job.");
            output.WriteLine(format == "ndjson" ? OfficeBatchReportSerializer.SerializeSummaryEvent(result, succeeded)
                : OfficeBatchReportSerializer.Serialize(report with { Succeeded = succeeded }));
        }
    }
}
