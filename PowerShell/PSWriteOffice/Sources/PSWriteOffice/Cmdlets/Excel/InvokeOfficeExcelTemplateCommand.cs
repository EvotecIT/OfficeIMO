using System.Collections;
using System.Globalization;
using System.Management.Automation;
using System.Security;
using OfficeIMO;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Table;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Fills {{Marker}} placeholders, repeats template rows or sheets, and keeps or removes optional rows.</summary>
/// <example>
/// <summary>Fill template markers</summary>
/// <code>New-OfficeExcel -Path .\Reports\Service-Report-October.xlsx -Template .\Templates\Service-Report.xltx {
///     ExcelTemplate @{ Client = 'Northwind Traders'; Month = 'October 2026'; Availability = 0.9994 }
/// }</code>
/// <para>Creates a report from a template and replaces its {{Client}}, {{Month}} and {{Availability}} markers.</para>
/// </example>
/// <example>
/// <summary>Repeat a template row for each incident</summary>
/// <code>$incidents = Import-OfficeCsv -Path .\Data\Incidents-October.csv
/// $values = @{
///     Client       = 'Northwind Traders'
///     Month        = 'October 2026'
///     Availability = 0.9994
/// }
/// New-OfficeExcel -Path .\Reports\Service-Report-October.xlsx -Template .\Templates\Service-Report.xltx {
///     ExcelTemplate $values -Culture en-GB
///     ExcelSheet 'Report' {
///         ExcelTemplate -TemplateRow 12 -InputObject $incidents
///         ExcelTemplate -OptionalRow 18 -RowCount 3 -Remove
///     }
/// }</code>
/// <para>Fills the header markers, repeats row 12 once per incident and removes the escalation block because nothing was escalated this month.</para>
/// </example>
/// <example>
/// <summary>Create one sheet per site from a template sheet</summary>
/// <code>$sites = Import-OfficeCsv -Path .\Facilities\Sites.csv
/// Open-OfficeExcel -Path .\Facilities\Inspections-2026.xlsx {
///     ExcelTemplate -TemplateSheet Site -InputObject $sites -SheetNameProperty Name -ThrowOnMissing
/// }</code>
/// <para>Copies the Site template sheet once per site, names each copy after the site and fills its markers from the site record.</para>
/// </example>
[Cmdlet(VerbsLifecycle.Invoke, "OfficeExcelTemplate", DefaultParameterSetName = "Values")]
[Alias("ExcelTemplate")]
public sealed class InvokeOfficeExcelTemplateCommand : PSCmdlet {
    /// <summary>Caller-owned workbook; omission uses the enclosing Excel content block.</summary>
    [Parameter(ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.</summary>
    [Parameter(ParameterSetName = "Values")]
    [Parameter(ParameterSetName = "Rows")]
    [Parameter(ParameterSetName = "OptionalRows")]
    [OfficeSheetTransformation]
    public OfficeSheetSelector[]? Sheet { get; set; }

    /// <summary>Value or values consumed by the selected operation.</summary>
    [Parameter(ParameterSetName = "Values", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "OptionalRows")]
    public Hashtable Value { get; set; } = null!;

    /// <summary>Culture used for value conversion and number-format presets.</summary>
    [Parameter()]
    public CultureInfo? Culture { get; set; }

    /// <summary>Behavior used when a marker is not supplied by -Value.</summary>
    [Parameter()]
    public ExcelTemplateMissingValueBehavior? MissingValueBehavior { get; set; }

    /// <summary>Throws when a marker is not supplied by -Value.</summary>
    [Parameter()]
    public SwitchParameter ThrowOnMissing { get; set; }

    /// <summary>Emits the native result after the operation.</summary>
    [Parameter()]
    public SwitchParameter PassThru { get; set; }

    /// <summary>1-based row number that contains template markers to repeat.</summary>
    [Parameter(ParameterSetName = "Rows", Mandatory = true)]
    [ValidateRange(1, 1048576)]
    public int? TemplateRow { get; set; }

    /// <summary>Objects, dictionaries or native tabular data supplied directly or through the pipeline.</summary>
    [Parameter(ParameterSetName = "Rows", Mandatory = true, Position = 1, ValueFromPipeline = true)]
    [Parameter(ParameterSetName = "Sheets", Mandatory = true, Position = 1, ValueFromPipeline = true)]
    public object InputObject { get; set; } = null!;

    /// <summary>Template worksheet name. Defaults to the current sheet inside an ExcelSheet block or the first sheet for path/document use.</summary>
    [Parameter(ParameterSetName = "Sheets", Mandatory = true)]
    [OfficeSheetTransformation]
    public OfficeSheetSelector TemplateSheet { get; set; } = null!;

    /// <summary>Input property used as the generated worksheet name.</summary>
    [Parameter(ParameterSetName = "Sheets")]
    public string? SheetNameProperty { get; set; }

    /// <summary>Template row blocks whose visibility or inclusion is controlled by the supplied values.</summary>
    [Parameter(ParameterSetName = "OptionalRows", Mandatory = true)]
    [ValidateRange(1, 1048576)]
    public int? OptionalRow { get; set; }

    /// <summary>Number of rows in the optional block.</summary>
    [Parameter(ParameterSetName = "OptionalRows")]
    [ValidateRange(1, 1048576)]
    public int RowCount { get; set; } = 1;

    /// <summary>Removes the optional row block instead of keeping and binding it.</summary>
    [Parameter(ParameterSetName = "OptionalRows")]
    public SwitchParameter Remove { get; set; }

    private readonly List<object?> _rows = new();

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (ParameterSetName is "Rows" or "Sheets") {
            TableInputCollector.AddInput(_rows, InputObject);
            return;
        }
        ApplyTemplate();
    }

    /// <inheritdoc />
    protected override void EndProcessing() {
        if (ParameterSetName is "Rows" or "Sheets") ApplyTemplate();
    }

    private void ApplyTemplate() {
        var document = ExcelCommandTargets.Document(Document);
        var options = ExcelTemplateValueService.CreateOptions(Culture?.Name, MissingValueBehavior, ThrowOnMissing);
        int replacements = 0;
        if (ParameterSetName == "Sheets") {
            var rows = ExcelTemplateValueService.ConvertRows(_rows);
            Func<IDictionary<string, object?>, int, string>? getName = SheetNameProperty == null ? null : (values, _) => ExcelTemplateValueService.GetStringValue(values, SheetNameProperty) ?? string.Empty;
            replacements = document.ApplyTemplateSheets(TemplateSheet!.Resolve(document.Sheets, s => s.Name).Name, rows, getName, options);
        } else {
            var selectors = Sheet;
            foreach (var sheet in ExcelSheetResolver.ResolveMany(document, selectors, Document == null ? ExcelDslContext.Current?.CurrentSheet : null)) {
                if (ParameterSetName == "Rows") replacements += sheet.ApplyTemplateRows(TemplateRow!.Value, ExcelTemplateValueService.ConvertRows(_rows), options);
                else if (ParameterSetName == "OptionalRows") replacements += Remove ? sheet.RemoveTemplateOptionalRows(OptionalRow!.Value, RowCount) : sheet.ApplyTemplateOptionalRows(OptionalRow!.Value, RowCount, true, Value == null ? new Dictionary<string, object?>() : ExcelTemplateValueService.ConvertValues(Value), options);
                else replacements += sheet.ApplyTemplate(ExcelTemplateValueService.ConvertValues(Value!), options);
            }
        }
        if (PassThru) WriteObject(replacements);
    }

}
