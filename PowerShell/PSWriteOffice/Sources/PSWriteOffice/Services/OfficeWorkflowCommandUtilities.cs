using OfficeIMO;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Workflows;
using OfficeIMO.Workflows.IWork;

namespace PSWriteOffice.Services;

/// <summary>Captures the optional first-party workflow adapters shipped by this PowerShell host.</summary>
internal static class OfficeWorkflowCommandUtilities {
    private static readonly Lazy<IReadOnlyList<OfficeWorkflowRoute>> Routes = new(() => CreateRunner().ConversionRoutes);
    internal static OfficeWorkflowRunner CreateRunner() { RegisterConversions(); return IWorkWorkflow.CreateRunner(); }
    internal static void RegisterConversions() { OfficeWorkflow.RegisterConversions(); IWorkWorkflow.RegisterConversions(); }
    internal static bool IsExecutable(string id) => Routes.Value.Any(route => route.Id == id);

    internal static ErrorRecord? GetBatchError(OfficeBatchSummary result, bool allowEmpty, string errorPrefix) {
        if (result.FailureSummary != null)
            return new ErrorRecord(new InvalidOperationException($"The batch stopped before completing discovery: {result.FailureSummary} Completed files remain available; inspect the summary and per-file evidence."),
                errorPrefix + "Failed", ErrorCategory.InvalidOperation, result);
        if (result.Failed > 0)
            return new ErrorRecord(new InvalidOperationException($"{result.Failed} conversion(s) failed. Inspect the summary or use -ItemResults for per-file evidence."),
                errorPrefix + "Failed", ErrorCategory.InvalidOperation, result);
        if (result.Selected == 0 && !result.Cancelled && !allowEmpty)
            return new ErrorRecord(new InvalidOperationException("The batch selected no executable inputs. Inspect skipped files or use -AllowEmpty for an intentional empty selection."),
                errorPrefix + "Empty", ErrorCategory.InvalidResult, result);
        return null;
    }

    /// <summary>Publishes a saved-job outcome with its summary assigned before a single combined conversion/recording error.</summary>
    internal static void WriteSavedJobOutcome(AsyncPSCmdlet command, OfficeBatchSummary result,
        OfficeBatchReport? report, string? summaryVariable, bool allowEmpty, bool itemResults) {
        if (!string.IsNullOrWhiteSpace(summaryVariable)) command.SessionState.PSVariable.Set(summaryVariable, result);
        ErrorRecord? error = GetBatchError(result, allowEmpty, "OfficeConversionJob");
        if (report?.RecordingError is { } recordingError) {
            string conversion = error?.Exception.Message ?? (result.Cancelled
                ? "The conversion attempt was cancelled." : "The conversion attempt finished.");
            error = new ErrorRecord(new IOException(conversion + " The run report could not be finalized: " + recordingError +
                " Completed outputs and the conversion summary remain available."), "OfficeConversionJobReportFailed", ErrorCategory.WriteError, report);
        }
        if (error != null) command.WriteError(error);
        if (!itemResults) command.WriteObject(result);
    }
}
