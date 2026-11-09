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

/// <summary>Adds or replaces a series trendline.</summary>
/// <example>
/// <summary>Add a linear trendline</summary>
/// <code>New-OfficeExcel -Path .\Reports\Growth.xlsx {
///     ExcelSheet 'Users' {
///         ExcelChart -InputObject (Import-OfficeCsv -Path .\Data\Monthly-Users.csv) -Category Month -Value ActiveUsers -Type Line -Address A1 -PassThru | ExcelChartTrendline -Index 1 -Type Linear
///     }
/// }</code>
/// <para>Adds a straight trendline to the first series.</para>
/// </example>
/// <example>
/// <summary>Project storage growth</summary>
/// <code>New-OfficeExcel -Path .\Reports\Storage-Forecast.xlsx {
///     ExcelSheet 'Storage' {
///         $chart = ExcelChart -InputObject (Import-OfficeCsv -Path .\Data\Storage-Used.csv) -Category Month -Value UsedTB -Type Line -Address A1 -Title 'Storage used (TB)' -PassThru
///         ExcelChartTrendline -Chart $chart -Name UsedTB -Type Linear -Forward 3 -DisplayRSquared
///     }
/// }</code>
/// <para>Fits a trendline to monthly storage use, projects it three months forward and shows the R-squared value on the chart.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelChartTrendline", DefaultParameterSetName = "Index")]
[Alias("ExcelChartTrendline")]
public sealed class SetOfficeExcelChartTrendlineCommand : PSCmdlet {
    /// <summary>Chart to update.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    public ExcelChart Chart { get; set; } = null!;

    /// <summary>One-based index identifying the target item.</summary>
    [Parameter(ParameterSetName = "Index", Mandatory = true)]
    [ValidateRange(1, int.MaxValue)]
    public int Index { get; set; }

    /// <summary>Trendline type.</summary>
    [Parameter()]
    public OfficeChartTrendlineType Type { get; set; } = OfficeChartTrendlineType.Linear;

    /// <summary>Polynomial order.</summary>
    [Parameter()]
    public int? Order { get; set; }

    /// <summary>Moving-average period.</summary>
    [Parameter()]
    public int? Period { get; set; }

    /// <summary>Forward forecast units.</summary>
    [Parameter()]
    public double? Forward { get; set; }

    /// <summary>Backward forecast units.</summary>
    [Parameter()]
    public double? Backward { get; set; }

    /// <summary>Trendline intercept.</summary>
    [Parameter()]
    public double? Intercept { get; set; }

    /// <summary>Display the trendline equation.</summary>
    [Parameter()]
    public SwitchParameter DisplayEquation { get; set; }

    /// <summary>Display the R-squared value.</summary>
    [Parameter()]
    public SwitchParameter DisplayRSquared { get; set; }

    /// <summary>Trendline line color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter()]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? LineColor { get; set; }

    /// <summary>Physical line width; numbers are interpreted as points.</summary>
    [Parameter()]
    [OfficeLengthTransformation]
    public OfficeLength? LineWidth { get; set; }

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
        if (Name != null) Chart.SetSeriesTrendline(Name, Type, Order, Period, Forward, Backward, Intercept, DisplayEquation, DisplayRSquared, LineColor, LineWidth, !CaseSensitive);
        else Chart.SetSeriesTrendline(Index - 1, Type, Order, Period, Forward, Backward, Intercept, DisplayEquation, DisplayRSquared, LineColor, LineWidth);
        if (PassThru) WriteObject(Chart);
    }

}
