using System;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Excel;
using PSWriteOffice.Services.Table;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds a native Excel chart from a range, named table or input rows.</summary>
/// <para>Table columns are selected by header. With no value selection, every numeric column except the category is plotted. Input rows are written as a table below existing worksheet data. Output is quiet unless PassThru is specified.</para>
/// <example>
///   <summary>Chart selected table columns.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Sales' { ExcelChart -Table Sales -Category Region -Value Revenue,Profit -Address F2 -Type ColumnClustered }</code>
///   <para>Places a chart at F2 using only the named table's data rows.</para>
/// </example>
/// <example>
///   <summary>Write input rows and chart them.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$rows | Add-OfficeExcelChart -Document $book -Sheet Sales -Category Region -Value Revenue -Address F2 -Width 18cm -Height 9cm</code>
///   <para>Writes a table on Sales and adds a native chart with length-based dimensions.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelChart", DefaultParameterSetName = RangeSet)]
[Alias("ExcelChart")]
[OutputType(typeof(ExcelChart))]
public sealed partial class AddOfficeExcelChartCommand : PSCmdlet {
    private const string RangeSet = "Range";
    private const string TableSet = "Table";
    private const string DataSet = "Data";
    private readonly List<object?> _items = new();
    private bool _explicitInput;

    /// <summary>Workbook to modify outside an Excel composition block.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = RangeSet)]
    [Parameter(ValueFromPipeline = true, ParameterSetName = TableSet)]
    [Parameter(ValueFromPipeline = true, ParameterSetName = DataSet)]
    public ExcelDocument? Document { get; set; }

    /// <summary>Sheet name, one-based position or an existing sheet belonging to Document.</summary>
    [Parameter]
    [OfficeSheetTransformation]
    public OfficeSheetSelector? Sheet { get; set; }

    /// <summary>Bounded A1 range containing categories followed by value columns.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = RangeSet)]
    [ValidateNotNullOrEmpty]
    public string Range { get; set; } = string.Empty;

    /// <summary>Name or display name of a table on the selected worksheet.</summary>
    [Parameter(Mandatory = true, ParameterSetName = TableSet)]
    [ValidateNotNullOrEmpty]
    public string Table { get; set; } = string.Empty;

    /// <summary>Objects or tabular input to write as a table and chart.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = DataSet)]
    [ValidateNotNull]
    public object? InputObject { get; set; }

    /// <summary>Category column; defaults to the first table column.</summary>
    [Parameter(ParameterSetName = TableSet)]
    [Parameter(Mandatory = true, ParameterSetName = DataSet)]
    [ValidateNotNullOrEmpty]
    public string? Category { get; set; }

    /// <summary>Value columns in series order; defaults to all numeric table columns except Category.</summary>
    [Parameter(ParameterSetName = TableSet)]
    [Parameter(Mandatory = true, ParameterSetName = DataSet)]
    [ValidateNotNullOrEmpty]
    public string[]? Value { get; set; }

    /// <summary>Single A1 cell anchoring the top-left corner of the chart.</summary>
    [Parameter(Mandatory = true)]
    [ValidateNotNullOrEmpty]
    public string Address { get; set; } = string.Empty;

    /// <summary>Chart width; defaults to 640 pixels. Bare numbers mean points.</summary>
    [Parameter]
    [OfficeLengthTransformation]
    public OfficeLength Width { get; set; } = OfficeLength.Pixels(640);

    /// <summary>Chart height; defaults to 360 pixels. Bare numbers mean points.</summary>
    [Parameter]
    [OfficeLengthTransformation]
    public OfficeLength Height { get; set; } = OfficeLength.Pixels(360);

    /// <summary>Shared chart family. Bubble requires the separate native bubble-data contract.</summary>
    [Parameter]
    public OfficeChartKind Type { get; set; } = OfficeChartKind.ColumnClustered;

    /// <summary>Chart title.</summary>
    [Parameter]
    public string? Title { get; set; }

    /// <summary>Dashboard appearance preset from the native Excel chart owner.</summary>
    [Parameter]
    public ExcelDashboardChartPreset Preset { get; set; } = ExcelDashboardChartPreset.Comparison;

    /// <summary>Built-in chart style identifier.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int? Style { get; set; }

    /// <summary>Built-in chart color-style identifier.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int? ColorStyle { get; set; }

    /// <summary>Treat the first row of Range as data rather than column headers.</summary>
    [Parameter(ParameterSetName = RangeSet)]
    public SwitchParameter NoHeader { get; set; }

    /// <summary>Omit cached chart values. Requires a contiguous category/value source block.</summary>
    [Parameter]
    public SwitchParameter NoCachedData { get; set; }

    /// <summary>Return the native ExcelChart.</summary>
    [Parameter]
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
        if (NoCachedData.IsPresent && plan.SourceRange == null && plan.InputTable == null) {
            throw new PSArgumentException("-NoCachedData requires contiguous category/value columns. The native snapshot chart API always writes cached data.", nameof(NoCachedData));
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
                WidthPixels = placement.Width,
                HeightPixels = placement.Height,
                ChartType = ToExcelChartType(kind),
                HasHeaders = plan.HasHeaders,
                Title = Title,
                IncludeCachedData = !NoCachedData.IsPresent,
                Preset = Preset,
                StyleId = Style,
                ColorStyleId = ColorStyle
            });
        } else {
            chart = sheet.AddChart(kind, plan.Data, placement.Row, placement.Column, placement.Width, placement.Height, Title);
            chart.ApplyDashboardPreset(Preset, Style, ColorStyle);
        }

        if (PassThru.IsPresent) WriteObject(chart);
    }

    private ExcelSheet ResolveSheet() {
        if (Document == null && Sheet?.Sheet is ExcelSheet explicitSheet && ExcelDslContext.Current == null) return explicitSheet;
        var document = Document ?? ExcelDslContext.Require(this).Document;
        if (Sheet != null) return Sheet.Resolve(document.Sheets, sheet => sheet.Name);
        return Document != null ? ExcelSheetResolver.Resolve(document, null, null) : ExcelDslContext.Require(this).RequireSheet();
    }
}
