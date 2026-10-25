#if !FRAMEWORK
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Workflows;
using OfficeIMO.Workflows.IWork;

namespace PSWriteOffice.Services;

/// <summary>Captures the optional first-party workflow adapters shipped by this PowerShell host.</summary>
internal static class OfficeWorkflowCommandUtilities {
    private static readonly Lazy<IReadOnlyList<OfficeWorkflowRoute>> Routes = new(() => CreateRunner().ConversionRoutes);
    internal static OfficeWorkflowRunner CreateRunner() => IWorkWorkflow.CreateRunner();
    internal static bool IsExecutable(string id) => Routes.Value.Any(route => route.Id == id);

    internal static ErrorRecord? GetBatchError(OfficeConversionBatchResult result, bool allowEmpty, string errorPrefix) {
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
}
#endif
