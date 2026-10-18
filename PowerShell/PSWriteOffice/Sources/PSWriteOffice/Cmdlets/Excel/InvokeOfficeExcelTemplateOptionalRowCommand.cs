using System.Collections;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Includes or removes an optional Excel template row block.</summary>
/// <example>
///   <summary>Keep an optional discount row and bind markers in it.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Invoke-OfficeExcelTemplateOptionalRow -Path .\Invoice.xlsx -Sheet Invoice -FirstRow 10 -Value @{ Discount = '10%' }</code>
///   <para>Leaves the optional row block in place, replaces its markers, and saves the workbook.</para>
/// </example>
[Cmdlet(VerbsLifecycle.Invoke, "OfficeExcelTemplateOptionalRow", SupportsShouldProcess = true, ConfirmImpact = ConfirmImpact.Low, DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelTemplateOptionalRow", "ExcelTemplateOptionalRows")]
[OutputType(typeof(int))]
public sealed class InvokeOfficeExcelTemplateOptionalRowCommand : PSCmdlet {
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPath = "Path";

    /// <summary>Workbook path to update.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Workbook to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Worksheet name. Defaults to the current sheet inside an ExcelSheet block.</summary>
    [Parameter]
    public string? Sheet { get; set; }

    /// <summary>Worksheet index when using a workbook object or path.</summary>
    [Parameter]
    public int? SheetIndex { get; set; }

    /// <summary>1-based first row in the optional block.</summary>
    [Parameter(Mandatory = true)]
    public int FirstRow { get; set; }

    /// <summary>Number of rows in the optional block.</summary>
    [Parameter]
    public int RowCount { get; set; } = 1;

    /// <summary>Template marker values used when the optional row block is included.</summary>
    [Parameter]
    [Alias("Values")]
    public Hashtable? Value { get; set; }

    /// <summary>Removes the optional row block instead of keeping and binding it.</summary>
    [Parameter]
    public SwitchParameter Remove { get; set; }

    /// <summary>Culture name used for built-in marker format aliases such as currency and date.</summary>
    [Parameter]
    public string? CultureName { get; set; }

    /// <summary>Behavior used when a marker in the optional block is not supplied by -Value.</summary>
    [Parameter]
    public ExcelTemplateMissingValueBehavior? MissingValueBehavior { get; set; }

    /// <summary>Throws when a marker in the optional block is not supplied by -Value.</summary>
    [Parameter]
    public SwitchParameter ThrowOnMissing { get; set; }

    /// <summary>Returns the number of marker replacements.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (FirstRow < 1) {
            ThrowTerminatingError(new ErrorRecord(
                new PSArgumentOutOfRangeException(nameof(FirstRow)),
                "InvalidFirstRow",
                ErrorCategory.InvalidArgument,
                FirstRow));
        }

        if (RowCount < 1) {
            ThrowTerminatingError(new ErrorRecord(
                new PSArgumentOutOfRangeException(nameof(RowCount)),
                "InvalidRowCount",
                ErrorCategory.InvalidArgument,
                RowCount));
        }

        using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path, Document, readOnly: false);
        var sheet = ExcelWorkbookCommandService.ResolveSheet(this, workbook.Document, ParameterSetName, Sheet, SheetIndex);
        var lastRow = FirstRow + RowCount - 1;
        var action = Remove.IsPresent
            ? "Remove Excel template optional rows"
            : "Apply Excel template optional rows";
        if (!ShouldProcess($"{sheet.Name}!{FirstRow}:{lastRow}", action)) {
            return;
        }

        var replacements = Remove.IsPresent
            ? sheet.RemoveTemplateOptionalRows(FirstRow, RowCount)
            : sheet.ApplyTemplateOptionalRows(
                FirstRow,
                RowCount,
                include: true,
                Value == null ? new Dictionary<string, object?>() : ExcelTemplateValueService.ConvertValues(Value),
                ExcelTemplateValueService.CreateOptions(CultureName, MissingValueBehavior, ThrowOnMissing.IsPresent));

        workbook.SaveIfOwned();
        if (PassThru.IsPresent) {
            WriteObject(replacements);
        }
    }
}