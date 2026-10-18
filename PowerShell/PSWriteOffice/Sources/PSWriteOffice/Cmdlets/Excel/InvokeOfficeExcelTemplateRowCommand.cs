using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;
using PSWriteOffice.Services.Table;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Repeats an Excel template row for pipeline data and replaces markers in each inserted row.</summary>
/// <example>
///   <summary>Fill invoice line rows from pipeline objects.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$items | Invoke-OfficeExcelTemplateRow -Path .\Invoice.xlsx -Sheet Invoice -TemplateRow 12 -CultureName en-US</code>
///   <para>Copies the template row once per input object, applies marker values, and saves the workbook.</para>
/// </example>
[Cmdlet(VerbsLifecycle.Invoke, "OfficeExcelTemplateRow", SupportsShouldProcess = true, ConfirmImpact = ConfirmImpact.Low, DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelTemplateRow", "ExcelTemplateRows")]
[OutputType(typeof(int))]
public sealed class InvokeOfficeExcelTemplateRowCommand : PSCmdlet {
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPath = "Path";
    private readonly List<object?> _rows = [];

    /// <summary>Workbook path to update.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Workbook to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Worksheet name. Defaults to the current sheet inside an ExcelSheet block.</summary>
    [Parameter]
    public string? Sheet { get; set; }

    /// <summary>Worksheet index when using a workbook object or path.</summary>
    [Parameter]
    public int? SheetIndex { get; set; }

    /// <summary>1-based row number that contains template markers to repeat.</summary>
    [Parameter(Mandatory = true)]
    public int TemplateRow { get; set; }

    /// <summary>Pipeline data. Hashtables, dictionaries, PSCustomObjects, and typed objects are supported.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 1)]
    [Alias("Rows", "Data")]
    public object? InputObject { get; set; }

    /// <summary>Culture name used for built-in marker format aliases such as currency and date.</summary>
    [Parameter]
    public string? CultureName { get; set; }

    /// <summary>Behavior used when a marker is not supplied by each input row.</summary>
    [Parameter]
    public ExcelTemplateMissingValueBehavior? MissingValueBehavior { get; set; }

    /// <summary>Throws when a marker is not supplied by an input row.</summary>
    [Parameter]
    public SwitchParameter ThrowOnMissing { get; set; }

    /// <summary>Returns the number of marker replacements.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        TableInputCollector.AddInput(_rows, InputObject);
    }

    /// <inheritdoc />
    protected override void EndProcessing() {
        if (TemplateRow < 1) {
            ThrowTerminatingError(new ErrorRecord(
                new PSArgumentOutOfRangeException(nameof(TemplateRow)),
                "InvalidTemplateRow",
                ErrorCategory.InvalidArgument,
                TemplateRow));
        }

        if (_rows.Count == 0) {
            return;
        }

        var rows = ExcelTemplateValueService.ConvertRows(_rows);
        var options = ExcelTemplateValueService.CreateOptions(CultureName, MissingValueBehavior, ThrowOnMissing.IsPresent);
        using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path, Document, readOnly: false);
        var sheet = ExcelWorkbookCommandService.ResolveSheet(this, workbook.Document, ParameterSetName, Sheet, SheetIndex);
        if (!ShouldProcess($"{sheet.Name}!{TemplateRow}", "Apply Excel template rows")) {
            return;
        }

        var replacements = sheet.ApplyTemplateRows(TemplateRow, rows, options);
        workbook.SaveIfOwned();
        if (PassThru.IsPresent) {
            WriteObject(replacements);
        }
    }
}