namespace OfficeIMO.Tool.Commands.Workflow;

/// <summary>Serializes shared batch outcomes without mixing human progress into stdout.</summary>
/// <remarks>Skipped files are counted and summarized at the end; <c>verbose</c> lists each one as it is skipped.</remarks>
internal sealed class WorkflowBatchOutput(TextWriter output, TextWriter error, string? format, bool verbose = false) : IProgress<OfficeBatchItemReport> {
    private readonly object _gate = new();
    private readonly Dictionary<string, long> _skipped = new(StringComparer.Ordinal);

    public void Report(OfficeBatchItemReport item) {
        lock (_gate) {
            if (format == "ndjson") output.WriteLine(OfficeBatchReportSerializer.SerializeItemEvent(item));
            if (item.Skipped && !verbose) {
                string reason = DescribeSkip(item.SkippedReason);
                _skipped[reason] = _skipped.TryGetValue(reason, out long count) ? count + 1 : 1;
                return;
            }
            if (item.Skipped || item.Status != OfficeBatchItemStatus.Completed) error.WriteLine(item.InputPath + ": " + item.Summary);
            foreach (OfficeDiagnostic finding in item.Diagnostics.Where(finding => finding.Severity != OfficeDiagnosticSeverity.Information))
                error.WriteLine(item.InputPath + ": " + finding.Code + (finding.Location == null ? "" : " [" + finding.Location + "]") + ": " + finding.Message);
        }
    }

    internal void Complete(OfficeBatchReport report, bool allowEmpty = false, string? recordingError = null) {
        lock (_gate) {
            OfficeBatchSummary result = report.Summary;
            bool succeeded = report.Succeeded && recordingError == null;
            foreach (KeyValuePair<string, long> skipped in _skipped.OrderBy(pair => pair.Key, StringComparer.Ordinal))
                error.WriteLine("Skipped " + skipped.Value + (skipped.Value == 1 ? " file: " : " files: ") + skipped.Key + ".");
            if (_skipped.Count > 0) error.WriteLine("The report lists each skipped file; add --verbose to print them here.");
            if (recordingError != null) error.WriteLine("The conversions finished, but their run report could not be finalized: " + recordingError);
            if (result.FailureSummary != null) error.WriteLine("Batch stopped: " + result.FailureSummary);
            if (result.Selected == 0 && !allowEmpty && !result.Cancelled && result.FailureSummary == null)
                error.WriteLine("No files were selected for conversion. Check the inputs and executable route, or use --allow-empty for an intentionally empty job.");
            output.WriteLine(format == "ndjson" ? OfficeBatchReportSerializer.SerializeSummaryEvent(result, succeeded)
                : OfficeBatchReportSerializer.Serialize(report with { Succeeded = succeeded }));
        }
    }

    private static string DescribeSkip(string? reason) => reason switch {
        "OutputExists" => "existing output left unchanged",
        "UpToDate" => "output is newer than its source",
        _ => "no conversion route to the target format, or excluded by --source-extension"
    };
}
