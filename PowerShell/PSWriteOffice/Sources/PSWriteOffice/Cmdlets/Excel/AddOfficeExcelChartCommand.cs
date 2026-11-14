using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Reports;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Excel;
using PSWriteOffice.Services.Table;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds a native Excel chart from a range, named table or input rows.</summary>
/// <para>Table columns are selected by header. With no value selection, every numeric column except the category is plotted. Input rows are written as a table below existing worksheet data. Output is quiet unless PassThru is specified.</para>
/// <example>
/// <summary>Chart a table</summary>
/// <code>New-OfficeExcel -Path .\Reports\Regions.xlsx {
///     ExcelSheet 'Summary' {
///         ExcelTable (Import-OfficeCsv -Path .\Data\Region-Revenue.csv) -Name RegionSummary
///         ExcelChart -Table RegionSummary -Category Region -Value Revenue -Address F2 -Type Column -Title 'Revenue by region'
///     }
/// }</code>
/// <para>Adds a themed column chart of the RegionSummary table, anchored at F2.</para>
/// </example>
/// <example>
/// <summary>Chart pipeline data directly</summary>
/// <code>$trend = Import-OfficeCsv -Path .\Data\Ticket-Trend.csv
/// New-OfficeExcel -Path .\Reports\Helpdesk-Trend.xlsx {
///     ExcelSheet 'Trend' {
///         ExcelChart -InputObject $trend -Category Week -Value Opened, Closed -Type Line -Address A1 -Title 'Tickets per week' -Width 18cm -Height 9cm
///     }
/// }</code>
/// <para>Writes the weekly ticket trend into the sheet and draws a line chart of opened and closed tickets from it in one step.</para>
/// </example>
/// <example>
/// <summary>Chart a fixed range</summary>
/// <code>Open-OfficeExcel -Path .\Finance\Budget-2027.xlsx {
///     ExcelSheet 'Summary' {
///         ExcelChart B3:C15 -Address E3 -Type Bar -Title 'Budget by month'
///     }
/// }</code>
/// <para>Charts the monthly budget in B3:C15 of an existing sheet as a clustered bar chart.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelChart", DefaultParameterSetName = "Range")]
[Alias("ExcelChart")]
[OutputType(typeof(ExcelChart))]
public sealed partial class AddOfficeExcelChartCommand : PSCmdlet {
    private const string RangeSet = "Range";
    private const string TableSet = "Table";
    private const string DataSet = "Data";
    private readonly List<object?> _items = new();
    private bool _explicitInput;

    /// <summary>Workbook to modify outside an Excel composition block.</summary>
    [Parameter(ValueFromPipeline = true)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Sheet name, one-based position or an existing sheet belonging to Document.</summary>
    [Parameter()]
    [OfficeSheetTransformation]
    public OfficeSheetSelector? Sheet { get; set; }

    /// <summary>Bounded A1 range containing categories followed by value columns.</summary>
    [Parameter(ParameterSetName = "Range", Mandatory = true, Position = 0)]
    [ValidateNotNullOrEmpty]
    public string Range { get; set; } = string.Empty;

    /// <summary>Name or display name of a table on the selected worksheet.</summary>
    [Parameter(ParameterSetName = "Table", Mandatory = true)]
    [ValidateNotNullOrEmpty]
    public string Table { get; set; } = string.Empty;

    /// <summary>Objects or tabular input to write as a table and chart.</summary>
    [Parameter(ParameterSetName = "Data", Mandatory = true, ValueFromPipeline = true)]
    [ValidateNotNull]
    public object? InputObject { get; set; }

    /// <summary>Category column; defaults to the first table column.</summary>
    [Parameter(ParameterSetName = "Table")]
    [Parameter(ParameterSetName = "Data", Mandatory = true)]
    [ValidateNotNullOrEmpty]
    public string? Category { get; set; }

    /// <summary>Value columns in series order; defaults to all numeric table columns except Category.</summary>
    [Parameter(ParameterSetName = "Table")]
    [Parameter(ParameterSetName = "Data", Mandatory = true)]
    [ValidateNotNullOrEmpty]
    public string[]? Value { get; set; }

    /// <summary>Single A1 cell anchoring the top-left corner of the chart.</summary>
    [Parameter(Mandatory = true)]
    [ValidateNotNullOrEmpty]
    public string Address { get; set; } = string.Empty;

    /// <summary>Chart width; defaults to 640 pixels. Bare numbers mean points.</summary>
    [Parameter()]
    [OfficeLengthTransformation]
    public OfficeLength Width { get; set; } = OfficeLength.Pixels(640);

    /// <summary>Chart height; defaults to 360 pixels. Bare numbers mean points.</summary>
    [Parameter()]
    [OfficeLengthTransformation]
    public OfficeLength Height { get; set; } = OfficeLength.Pixels(360);

    /// <summary>Shared chart family. Bubble requires the separate native bubble-data contract.</summary>
    [Parameter()]
    public OfficeChartType Type { get; set; } = OfficeChartType.ColumnClustered;

    /// <summary>Chart title.</summary>
    [Parameter()]
    public string? Title { get; set; }

    /// <summary>Dashboard appearance preset from the native Excel chart owner.</summary>
    [Parameter()]
    public ExcelDashboardChartPreset Preset { get; set; } = ExcelDashboardChartPreset.Comparison;

    /// <summary>Built-in chart style identifier.</summary>
    [Parameter()]
    [ValidateRange(1, int.MaxValue)]
    public int? Style { get; set; }

    /// <summary>Built-in chart color-style identifier.</summary>
    [Parameter()]
    [ValidateRange(1, int.MaxValue)]
    public int? ColorStyle { get; set; }

    /// <summary>Treat the first row of Range as data rather than column headers.</summary>
    [Parameter(ParameterSetName = "Range")]
    public SwitchParameter NoHeader { get; set; }

    /// <summary>Omit cached chart values. Not available for scatter charts of non-adjacent table columns.</summary>
    [Parameter()]
    public SwitchParameter NoCachedData { get; set; }

    /// <summary>Return the native ExcelChart.</summary>
    [Parameter()]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void BeginProcessing() {
        _explicitInput = MyInvocation.BoundParameters.ContainsKey(nameof(InputObject));
    }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (ParameterSetName == DataSet) {
            if (_explicitInput) _items.Clear();
            TableInputCollector.AddInput(_items, InputObject, preserveTabularInput: true);
            if (_explicitInput) AddChart();
            return;
        }

        AddChart();
    }

    /// <inheritdoc />
    protected override void EndProcessing() {
        if (ParameterSetName == DataSet && !_explicitInput) AddChart();
    }

    private void AddChart() {
        var placement = ResolvePlacement();
        var sheet = ResolveSheet();
        var plan = ParameterSetName == DataSet ? PlanInput(sheet) : PlanWorksheet(sheet);
        var kind = ResolveChartKind();
        ValidateChartData(plan.Data, kind);
        if (NoCachedData.IsPresent && plan.SourceRange == null && plan.InputTable == null && !plan.UsesSeriesRanges(kind)) {
            throw new PSArgumentException("-NoCachedData is not available for scatter charts of non-adjacent columns. The native snapshot chart API always writes cached data.", nameof(NoCachedData));
        }

        // All mapping, source, dimensions and family validation precede worksheet writes.
        if (plan.InputTable != null) {
            sheet.InsertDataTableAsTable(plan.InputTable, startRow: plan.StartRow, startColumn: 1, includeHeaders: true);
        }

        ExcelChart chart;
        if (plan.SourceRange != null) {
            chart = sheet.AddDashboardChart(new ExcelDashboardChartOptions {
                Range = plan.SourceRange,
                Row = placement.Row,
                Column = placement.Column,
                Width = OfficeLength.Pixels(placement.Width),
                Height = OfficeLength.Pixels(placement.Height),
                ChartType = kind,
                HasHeaders = plan.HasHeaders,
                Title = Title,
                IncludeCachedData = !NoCachedData.IsPresent,
                Preset = Preset,
                StyleId = Style,
                ColorStyleId = ColorStyle
            });
        } else if (plan.UsesSeriesRanges(kind)) {
            chart = AddRangeSeriesChart(sheet, kind, plan, new ExcelChartOptions { Row = placement.Row, Column = placement.Column, Width = OfficeLength.Pixels(placement.Width), Height = OfficeLength.Pixels(placement.Height), Title = Title }, !NoCachedData.IsPresent);
            chart.ApplyDashboardPreset(Preset, Style, ColorStyle);
        } else {
            chart = sheet.AddChart(kind, plan.Data, new ExcelChartOptions { Row = placement.Row, Column = placement.Column, Width = OfficeLength.Pixels(placement.Width), Height = OfficeLength.Pixels(placement.Height), Title = Title });
            chart.ApplyDashboardPreset(Preset, Style, ColorStyle);
        }

        if (PassThru.IsPresent) WriteObject(chart);
    }

    private ExcelSheet ResolveSheet() {
        return ExcelSheetResolver.ResolveTarget(this, Document, Sheet);
    }
}
