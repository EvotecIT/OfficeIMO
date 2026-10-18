using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;
using PSWriteOffice.Services.Table;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Builds a worksheet dashboard from tabular data using OfficeIMO dashboard defaults.</summary>
/// <example>
///   <summary>Create a dashboard table and chart.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$rows | New-OfficeExcelDashboard -Title 'Sales Dashboard' -TableName Sales -ChartPreset CompactComparison</code>
///   <para>Writes a table and chart into the current Excel DSL worksheet.</para>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeExcelDashboard", DefaultParameterSetName = ParameterSetContext, SupportsShouldProcess = true)]
[Alias("ExcelDashboard")]
[OutputType(typeof(PSObject))]
public sealed class NewOfficeExcelDashboardCommand : PSCmdlet {
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPath = "Path";

    private readonly List<object?> _items = new();

    /// <summary>Workbook path to update.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetPath)]
    [Alias("InputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Workbook to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Worksheet name when using Path or Document.</summary>
    [Parameter(ParameterSetName = ParameterSetPath)]
    [Parameter(ParameterSetName = ParameterSetDocument)]
    [Alias("SheetName", "Worksheet")]
    public string? Sheet { get; set; }

    /// <summary>Worksheet index (0-based) when using Path or Document.</summary>
    [Parameter(ParameterSetName = ParameterSetPath)]
    [Parameter(ParameterSetName = ParameterSetDocument)]
    public int? SheetIndex { get; set; }

    /// <summary>Rows to render in the dashboard table.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("Data", "DataTable")]
    public object? InputObject { get; set; }

    /// <summary>Dashboard title.</summary>
    [Parameter]
    public string? Title { get; set; }

    /// <summary>Dashboard subtitle.</summary>
    [Parameter]
    public string? Subtitle { get; set; }

    /// <summary>Name for the generated table.</summary>
    [Parameter]
    public string? TableName { get; set; }

    /// <summary>Top-left row for the generated table.</summary>
    [Parameter]
    public int TableRow { get; set; } = 3;

    /// <summary>Top-left column for the generated table.</summary>
    [Parameter]
    public int TableColumn { get; set; } = 1;

    /// <summary>Built-in table style.</summary>
    [Parameter]
    public ExcelTableStyle TableStyle { get; set; } = ExcelTableStyle.TableStyleMedium9;

    /// <summary>Disable AutoFilter dropdowns on the generated table.</summary>
    [Parameter]
    public SwitchParameter NoAutoFilter { get; set; }

    /// <summary>Disable auto-fit for generated table columns.</summary>
    [Parameter]
    public SwitchParameter NoAutoFit { get; set; }

    /// <summary>Do not create a chart.</summary>
    [Parameter]
    public SwitchParameter NoChart { get; set; }

    /// <summary>Dashboard chart preset.</summary>
    [Parameter]
    public ExcelDashboardChartPreset ChartPreset { get; set; } = ExcelDashboardChartPreset.Comparison;

    /// <summary>Chart title. Defaults to Title when omitted.</summary>
    [Parameter]
    public string? ChartTitle { get; set; }

    /// <summary>Top-left chart row.</summary>
    [Parameter]
    public int? ChartRow { get; set; }

    /// <summary>Top-left chart column.</summary>
    [Parameter]
    public int? ChartColumn { get; set; }

    /// <summary>Emit dashboard build metadata.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        TableInputCollector.AddInput(_items, InputObject, preserveTabularInput: true);
    }

    /// <inheritdoc />
    protected override void EndProcessing() {
        var rows = TableInputCollector.RequireRows(_items, nameof(InputObject));
        var table = ExcelTabularInputService.ToDataTable(rows, TableName);
        var style = TableStyle;

        using var workbook = ExcelWorkbookCommandService.ResolveWorkbook(this, ParameterSetName, Path, Document, readOnly: false);

        if (!ExcelShouldProcessService.ShouldProcessWorkbook(this, workbook.Document, Path, "Update Excel workbook")) {

            return;

        }

        ExcelSheet sheet = ExcelWorkbookCommandService.ResolveSheet(this, workbook.Document, ParameterSetName, Sheet, SheetIndex);
        ExcelDashboardResult result = sheet.AddDashboard(table, new ExcelDashboardOptions {
            Title = Title,
            Subtitle = Subtitle,
            TableRow = TableRow,
            TableColumn = TableColumn,
            TableName = TableName,
            TableStyle = style,
            IncludeAutoFilter = !NoAutoFilter.IsPresent,
            AutoFit = !NoAutoFit.IsPresent,
            AddChart = !NoChart.IsPresent,
            ChartPreset = ChartPreset,
            ChartTitle = ChartTitle,
            ChartRow = ChartRow,
            ChartColumn = ChartColumn
        });
        workbook.SaveIfOwned();

        if (PassThru.IsPresent) {
            var output = new PSObject();
            output.Properties.Add(new PSNoteProperty("TableRange", result.TableRange));
            output.Properties.Add(new PSNoteProperty("TableName", result.TableName));
            output.Properties.Add(new PSNoteProperty("ChartTitle", result.Chart?.Title));
            output.Properties.Add(new PSNoteProperty("ChartType", result.Chart?.ChartType.ToString()));
            WriteObject(output);
        }
    }
}
