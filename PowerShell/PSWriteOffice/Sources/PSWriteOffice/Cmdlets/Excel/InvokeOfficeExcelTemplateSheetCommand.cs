using System;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;
using PSWriteOffice.Services.Table;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Repeats an Excel template worksheet for pipeline data and applies markers in each generated sheet.</summary>
/// <example>
///   <summary>Create one invoice worksheet per pipeline object.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$invoices | Invoke-OfficeExcelTemplateSheet -Path .\Invoices.xlsx -TemplateSheet Template -SheetNameProperty SheetName</code>
///   <para>Uses the template sheet for the first object, copies it for later objects, binds markers, and saves the workbook.</para>
/// </example>
[Cmdlet(VerbsLifecycle.Invoke, "OfficeExcelTemplateSheet", SupportsShouldProcess = true, ConfirmImpact = ConfirmImpact.Low, DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelTemplateSheet", "ExcelTemplateSheets")]
[OutputType(typeof(int))]
public sealed class InvokeOfficeExcelTemplateSheetCommand : PSCmdlet {
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPath = "Path";
    private readonly List<object?> _items = [];

    /// <summary>Workbook path to update.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Workbook to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Template worksheet name. Defaults to the current sheet inside an ExcelSheet block or the first sheet for path/document use.</summary>
    [Parameter]
    [Alias("Sheet", "WorksheetName")]
    public string? TemplateSheet { get; set; }

    /// <summary>Pipeline data. Hashtables, dictionaries, PSCustomObjects, and typed objects are supported.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 1)]
    [Alias("Rows", "Data", "InputObject")]
    public object? Item { get; set; }

    /// <summary>Input property used as the generated worksheet name.</summary>
    [Parameter]
    [Alias("NameProperty")]
    public string? SheetNameProperty { get; set; }

    /// <summary>Culture name used for built-in marker format aliases such as currency and date.</summary>
    [Parameter]
    public string? CultureName { get; set; }

    /// <summary>Behavior used when a marker is not supplied by each input item.</summary>
    [Parameter]
    public ExcelTemplateMissingValueBehavior? MissingValueBehavior { get; set; }

    /// <summary>Throws when a marker is not supplied by an input item.</summary>
    [Parameter]
    public SwitchParameter ThrowOnMissing { get; set; }

    /// <summary>Returns the number of marker replacements.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        TableInputCollector.AddInput(_items, Item);
    }

    /// <inheritdoc />
    protected override void EndProcessing() {
        if (_items.Count == 0) {
            return;
        }

        var items = ExcelTemplateValueService.ConvertRows(_items);
        Func<IDictionary<string, object?>, int, string>? sheetNameSelector = null;
        if (!string.IsNullOrWhiteSpace(SheetNameProperty)) {
            sheetNameSelector = (values, _) => ExcelTemplateValueService.GetStringValue(values, SheetNameProperty) ?? string.Empty;
        }

        var options = ExcelTemplateValueService.CreateOptions(CultureName, MissingValueBehavior, ThrowOnMissing.IsPresent);
        using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path, Document, readOnly: false);
        var templateSheetName = ExcelWorkbookCommandService.ResolveSheetNameOrCurrent(this, workbook.Document, ParameterSetName, TemplateSheet);
        if (!ShouldProcess(templateSheetName, "Apply Excel template sheets")) {
            return;
        }

        var replacements = workbook.Document.ApplyTemplateSheets(templateSheetName, items, sheetNameSelector, options);
        workbook.SaveIfOwned();
        if (PassThru.IsPresent) {
            WriteObject(replacements);
        }
    }
}