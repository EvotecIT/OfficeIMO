namespace OfficeIMO.Word.Rtf;

/// <summary>Result of an RTF workflow routed through the OfficeIMO.Word engine.</summary>
/// <typeparam name="T">Workflow-specific result type.</typeparam>
public sealed class RtfWordWorkflowResult<T> {
    /// <summary>Creates a result-bearing RTF workflow response.</summary>
    public RtfWordWorkflowResult(RtfDocument document, T workflowResult, RtfConversionReport report) {
        Document = document ?? throw new ArgumentNullException(nameof(document));
        WorkflowResult = workflowResult;
        Report = report ?? throw new ArgumentNullException(nameof(report));
    }

    /// <summary>RTF document produced by the workflow.</summary>
    public RtfDocument Document { get; }

    /// <summary>Workflow-specific result, such as a replacement count or field update report.</summary>
    public T WorkflowResult { get; }

    /// <summary>Combined parse, bridge, and workflow diagnostics.</summary>
    public RtfConversionReport Report { get; }

    /// <summary>Throws when any stage reported fidelity loss or a blocked feature.</summary>
    public RtfWordWorkflowResult<T> RequireNoLoss() {
        Report.RequireNoLoss();
        return this;
    }
}
