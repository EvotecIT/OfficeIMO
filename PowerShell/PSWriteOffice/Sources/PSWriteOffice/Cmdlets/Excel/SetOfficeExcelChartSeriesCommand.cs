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

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Styles a chart series or one of its data points.</summary>
/// <example>
/// <summary>Colour a series</summary>
/// <code>New-OfficeExcel -Path .\Reports\Helpdesk-Trend.xlsx {
///     ExcelSheet 'Trend' {
///         ExcelChart -InputObject (Import-OfficeCsv -Path .\Data\Ticket-Trend.csv) -Category Week -Value Opened, Closed -Type Line -Address A1 -PassThru | ExcelChartSeries -Index 2 -LineColor Success
///     }
/// }</code>
/// <para>Draws the second series, Closed, in the theme's success colour.</para>
/// </example>
/// <example>
/// <summary>Highlight one bar</summary>
/// <code>New-OfficeExcel -Path .\Reports\Regions.xlsx {
///     ExcelSheet 'Summary' {
///         $chart = ExcelChart -InputObject (Import-OfficeCsv -Path .\Data\Region-Revenue.csv) -Category Region -Value Revenue -Address A1 -Type Column -PassThru
///         ExcelChartSeries -Chart $chart -Name Revenue -Point 3 -FillColor Danger
///     }
/// }</code>
/// <para>Colours the bar of the region that missed its target in the danger colour and leaves the others on the theme palette.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelChartSeries", DefaultParameterSetName = "Index")]
[Alias("ExcelChartSeries")]
public sealed class SetOfficeExcelChartSeriesCommand : PSCmdlet {
    /// <summary>Chart to update.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    public ExcelChart Chart { get; set; } = null!;

    /// <summary>One-based index identifying the target item.</summary>
    [Parameter(ParameterSetName = "Index", Mandatory = true)]
    [ValidateRange(1, int.MaxValue)]
    public int Index { get; set; }

    /// <summary>One-based point index within the selected chart series.</summary>
    [Parameter()]
    [ValidateRange(1, uint.MaxValue)]
    public uint? Point { get; set; }

    /// <summary>Series fill color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter()]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? FillColor { get; set; }

    /// <summary>Series line color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter()]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? LineColor { get; set; }

    /// <summary>Physical line width; numbers are interpreted as points.</summary>
    [Parameter()]
    [OfficeLengthTransformation]
    public OfficeLength? LineWidth { get; set; }

    /// <summary>Marker style name.</summary>
    [Parameter()]
    public OfficeChartMarkerShape? MarkerStyle { get; set; }

    /// <summary>Marker size.</summary>
    [Parameter()]
    public int? MarkerSize { get; set; }

    /// <summary>Marker fill color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter()]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? MarkerFillColor { get; set; }

    /// <summary>Marker line color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter()]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? MarkerLineColor { get; set; }

    /// <summary>Physical width of the marker outline; omission preserves the existing width.</summary>
    [Parameter()]
    [OfficeLengthTransformation]
    public OfficeLength? MarkerLineWidth { get; set; }

    /// <summary>Emits the native result after the operation.</summary>
    [Parameter()]
    public SwitchParameter PassThru { get; set; }

    /// <summary>Named range name.</summary>
    [Parameter(ParameterSetName = "Name", Mandatory = true)]
    public string Name { get; set; } = null!;

    /// <summary>Matches names and text using ordinal case sensitivity. Omission ignores case.</summary>
    [Parameter(ParameterSetName = "Name")]
    public SwitchParameter CaseSensitive { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (Point.HasValue) {
            if (Name != null) Chart.SetDataPointColor(Name, Point.Value - 1, FillColor, LineColor, LineWidth, !CaseSensitive);
            else Chart.SetDataPointColor(Index - 1, Point.Value - 1, FillColor, LineColor, LineWidth);
        } else {
            if (FillColor.HasValue) {
                if (Name != null) Chart.SetSeriesFillColor(Name, FillColor.Value, !CaseSensitive);
                else Chart.SetSeriesFillColor(Index - 1, FillColor.Value);
            }
            if (LineColor.HasValue || LineWidth.HasValue) {
                if (!LineColor.HasValue) throw new PSArgumentException("LineColor is required when changing LineWidth.");
                if (Name != null) Chart.SetSeriesLineColor(Name, LineColor.Value, LineWidth, !CaseSensitive);
                else Chart.SetSeriesLineColor(Index - 1, LineColor.Value, LineWidth);
            }
            if (MarkerStyle.HasValue || MarkerSize.HasValue || MarkerFillColor.HasValue || MarkerLineColor.HasValue || MarkerLineWidth.HasValue) {
                OfficeLength? size = MarkerSize.HasValue ? OfficeLength.Points(MarkerSize.Value) : (OfficeLength?)null;
                if (Name != null) Chart.SetSeriesMarkerByName(Name, MarkerStyle, size, MarkerFillColor, MarkerLineColor, MarkerLineWidth, !CaseSensitive);
                else Chart.SetSeriesMarker(Index - 1, MarkerStyle, size, MarkerFillColor, MarkerLineColor, MarkerLineWidth);
            }
        }
        if (PassThru) WriteObject(Chart);
    }

}
