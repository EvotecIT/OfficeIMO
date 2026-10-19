using System;
using System.Management.Automation;
using OfficeIMO.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using OfficeIMO.Excel;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds or replaces an Excel chart series trendline.</summary>
/// <example>
///   <summary>Add a polynomial trendline.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$chart | Set-OfficeExcelChartTrendline -SeriesIndex 0 -Type Polynomial -Order 2 -DisplayEquation -DisplayRSquared</code>
///   <para>Adds a polynomial trendline to the first series.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelChartTrendline", DefaultParameterSetName = ParameterSetIndex)]
[Alias("ExcelChartTrendline")]
[OutputType(typeof(ExcelChart))]
public sealed class SetOfficeExcelChartTrendlineCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    private const string ParameterSetIndex = "Index";
    private const string ParameterSetSeriesName = "Name";

    /// <summary>Chart to update.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public ExcelChart Chart { get; set; } = null!;

    /// <summary>Zero-based series index.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetIndex)]
    public int SeriesIndex { get; set; }

    /// <summary>Series name.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetSeriesName)]
    public string SeriesName { get; set; } = string.Empty;

    /// <summary>Ignore case when matching series name.</summary>
    [Parameter(ParameterSetName = ParameterSetSeriesName)]
    public bool IgnoreCase { get; set; } = true;

    /// <summary>Trendline type.</summary>
    [Parameter(Mandatory = true)]
    public OfficeChartTrendlineType Type { get; set; }

    /// <summary>Polynomial order.</summary>
    [Parameter]
    public int? Order { get; set; }

    /// <summary>Moving-average period.</summary>
    [Parameter]
    public int? Period { get; set; }

    /// <summary>Forward forecast units.</summary>
    [Parameter]
    public double? Forward { get; set; }

    /// <summary>Backward forecast units.</summary>
    [Parameter]
    public double? Backward { get; set; }

    /// <summary>Trendline intercept.</summary>
    [Parameter]
    public double? Intercept { get; set; }

    /// <summary>Display the trendline equation.</summary>
    [Parameter]
    public SwitchParameter DisplayEquation { get; set; }

    /// <summary>Display the R-squared value.</summary>
    [Parameter]
    public SwitchParameter DisplayRSquared { get; set; }

    /// <summary>Trendline line color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation(UseExcelArgb = true)]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? LineColor { get; set; }

    /// <summary>Trendline line width in points.</summary>
    [Parameter]
    public double? LineWidthPoints { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        try {
            if (ParameterSetName == ParameterSetSeriesName) {
                Chart.SetSeriesTrendline(SeriesName, Type, Order, Period, Forward, Backward, Intercept, DisplayEquation.IsPresent, DisplayRSquared.IsPresent, LineColor, LineWidthPoints, IgnoreCase);
            } else {
                Chart.SetSeriesTrendline(SeriesIndex, Type, Order, Period, Forward, Backward, Intercept, DisplayEquation.IsPresent, DisplayRSquared.IsPresent, LineColor, LineWidthPoints);
            }

            WritePassThru(Chart);
        } catch (Exception ex) {
            WriteError(new ErrorRecord(ex, "ExcelChartTrendlineFailed", ErrorCategory.InvalidOperation, Chart));
        }
    }
}
