using System;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Configures OfficeIMO Excel execution and validation behavior for a workbook.</summary>
/// <example>
///   <summary>Force sequential execution for a workbook.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$workbook | Set-OfficeExcelExecutionPolicy -Mode Sequential</code>
///   <para>Disables automatic parallel execution decisions for subsequent OfficeIMO operations.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelExecutionPolicy", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelExecutionPolicy")]
[OutputType(typeof(ExcelDocument))]
public sealed class SetOfficeExcelExecutionPolicyCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>Workbook whose execution policy should be updated.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Execution mode for large operations.</summary>
    [Parameter]
    public ExcelExecutionMode? Mode { get; set; }

    /// <summary>Global item threshold above which Automatic mode switches to Parallel.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int? ParallelThreshold { get; set; }

    /// <summary>Optional cap for parallel compute phases.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int? MaxDegreeOfParallelism { get; set; }

    /// <summary>Worksheet validation mode for write operations.</summary>
    [Parameter]
    public ExcelWorksheetValidationMode? WorksheetValidation { get; set; }

    /// <summary>Request diagnostics-aware validation without wiring verbose callbacks.</summary>
    [Parameter]
    public SwitchParameter Diagnostics { get; set; }

    /// <summary>Do not save worksheet parts immediately after AutoFit width/height mutations.</summary>
    [Parameter]
    public SwitchParameter DisableAutoFitImmediateSave { get; set; }

    /// <summary>Emit the workbook after updating the policy.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var document = Document ?? ExcelDslContext.Require(this).Document;
        var policy = document.Execution;

        if (Mode.HasValue) policy.Mode = Mode.Value;

        if (ParallelThreshold.HasValue)
        {
            policy.ParallelThreshold = ParallelThreshold.Value;
        }

        if (MaxDegreeOfParallelism.HasValue)
        {
            policy.MaxDegreeOfParallelism = MaxDegreeOfParallelism.Value;
        }

        if (WorksheetValidation.HasValue) policy.WorksheetValidation = WorksheetValidation.Value;

        if (Diagnostics.IsPresent)
        {
            policy.DiagnosticsRequested = true;
        }

        if (DisableAutoFitImmediateSave.IsPresent)
        {
            policy.SaveWorksheetAfterAutoFit = false;
        }

        if (PassThru.IsPresent)
        {
            WriteObject(document);
        }
    }

}
