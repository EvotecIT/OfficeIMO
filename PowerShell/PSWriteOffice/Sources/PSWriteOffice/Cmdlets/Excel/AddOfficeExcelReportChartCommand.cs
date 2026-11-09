using System.Collections;
using System.Globalization;
using System.Management.Automation;
using System.Security;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;
using OfficeIMO.Excel.Reports;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds a chart of a report table, using a dashboard preset or chart type.</summary>
/// <example>
/// <summary>Chart the table above</summary>
/// <code>New-OfficeExcel -Path .\Reports\Regions.xlsx {
///     ExcelReportSheet 'Summary' {
///         ExcelReportTable (Import-OfficeCsv -Path .\Data\Region-Revenue.csv) 'Revenue by region'
///         ExcelReportChart -Category Region -Value Revenue
///     }
/// }</code>
/// <para>Charts the previous report table and places the chart beside it.</para>
/// </example>
/// <example>
/// <summary>Show a trend with a dashboard preset</summary>
/// <code>New-OfficeExcel -Path .\Reports\Helpdesk-Review.xlsx {
///     ExcelReportSheet 'Trend' {
///         ExcelReportTitle 'Helpdesk trend' 'Last 12 weeks'
///         ExcelReportTable (Import-OfficeCsv -Path .\Data\Ticket-Trend.csv) 'Tickets per week'
///         ExcelReportChart -Preset Trend -Category Week -Value Opened, Closed -Title 'Opened and closed tickets'
///     }
/// }</code>
/// <para>Adds a weekly ticket table and charts opened against closed tickets with the trend preset.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelReportChart", DefaultParameterSetName = "Default")]
[Alias("ExcelReportChart")]
public sealed class AddOfficeExcelReportChartCommand : PSCmdlet {
    /// <summary>Native Excel table used to resolve the source range or header.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string? Table { get; set; }

    /// <summary>Shared chart family. Bubble requires the separate native bubble-data contract.</summary>
    [Parameter(ParameterSetName = "Default")]
    public OfficeChartType? Type { get; set; }

    /// <summary>Dashboard appearance preset from the native Excel chart owner.</summary>
    [Parameter(ParameterSetName = "Default")]
    public ExcelDashboardChartPreset? Preset { get; set; }

    /// <summary>Chart title.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string? Title { get; set; }

    /// <summary>Category column; defaults to the first table column.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string? Category { get; set; }

    /// <summary>Value or values consumed by the selected operation.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string[]? Value { get; set; }

    /// <summary>Chart width; defaults to 640 pixels. Bare numbers mean points.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeLengthTransformation]
    public OfficeLength? Width { get; set; }

    /// <summary>Chart height; defaults to 360 pixels. Bare numbers mean points.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeLengthTransformation]
    public OfficeLength? Height { get; set; }

    /// <summary>Emits the native result after the operation.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var context = ExcelDslContext.Require(this);
        var sheet = context.RequireSheet();
        string table = Table ?? context.LastReportTable?.NameOrRange ?? throw new PSArgumentException("Supply -Table or add an ExcelReportTable first.");
        var source = sheet.GetTables().SingleOrDefault(item => item.Name == table || item.DisplayName == table) ?? throw new PSArgumentException("The report table does not exist.");
        var (row, _, _, column) = ExcelA1.ParseRange(source.Range);
        var plan = AddOfficeExcelChartCommand.ReportTablePlan(sheet, table, Category, Value);
        var type = Type ?? (Preset switch {
            ExcelDashboardChartPreset.Trend => OfficeChartType.Line,
            ExcelDashboardChartPreset.Contribution => OfficeChartType.Doughnut,
            ExcelDashboardChartPreset.CompactComparison => OfficeChartType.BarClustered,
            _ => OfficeChartType.ColumnClustered
        });
        AddOfficeExcelChartCommand.ValidateChartData(plan.Data, type);
        ExcelChart chart;
        if (plan.SourceRange != null) chart = sheet.AddDashboardChart(new ExcelDashboardChartOptions {
            Range = plan.SourceRange, HasHeaders = plan.HasHeaders, ChartType = type,
            Row = row, Column = column + 2, Width = Width, Height = Height, Title = Title,
            Preset = Preset ?? ExcelDashboardChartPreset.Comparison
        });
        else {
            chart = sheet.AddChart(type, plan.Data, new ExcelChartOptions { Row = row, Column = column + 2, Width = Width ?? OfficeLength.Pixels(640), Height = Height ?? OfficeLength.Pixels(360), Title = Title });
            chart.ApplyDashboardPreset(Preset ?? ExcelDashboardChartPreset.Comparison);
        }
        if (PassThru) WriteObject(chart);
    }

}
