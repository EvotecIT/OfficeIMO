#if !FRAMEWORK
using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Workflows;
using OfficeIMO.Workflows.IWork;

namespace PSWriteOffice.Services;

/// <summary>Captures the optional first-party workflow adapters shipped by this PowerShell host.</summary>
internal static class OfficeWorkflowCommandUtilities {
    private static readonly Lazy<IReadOnlyList<OfficeWorkflowRoute>> Routes = new(() => CreateRunner().ConversionRoutes);
    internal static OfficeWorkflowRunner CreateRunner() => IWorkWorkflow.CreateRunner();
    internal static bool IsExecutable(string id) => Routes.Value.Any(route => route.Id == id);
}
#endif
