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

/// <summary>Shows and styles chart data labels.</summary>
/// <example>
/// <summary>Show values on a chart</summary>
/// <code>New-OfficeExcel -Path .\Reports\Regions.xlsx {
///     ExcelSheet 'Summary' {
///         ExcelTable (Import-OfficeCsv -Path .\Data\Region-Revenue.csv) -Name RegionSummary
///         ExcelChart -Table RegionSummary -Address F2 -Type Column -PassThru | ExcelChartDataLabel -ShowValue
///     }
/// }</code>
/// <para>Shows each bar's value on the chart.</para>
/// </example>
/// <example>
/// <summary>Label a share chart with percentages</summary>
/// <code>New-OfficeExcel -Path .\Reports\Cloud-Spend.xlsx {
///     ExcelSheet 'Spend' {
///         $chart = ExcelChart -InputObject (Import-OfficeCsv -Path .\Data\Spend-By-Service.csv) -Category Service -Value Cost -Address A1 -Type Pie -Title 'Cloud spend by service' -PassThru
///         ExcelChartDataLabel -Chart $chart -ShowCategoryName -ShowPercent -Position OutsideEnd -NumberFormat '0%'
///     }
/// }</code>
/// <para>Labels each slice of a pie chart with its category and percentage outside the slice.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelChartDataLabel", DefaultParameterSetName = "Default")]
[Alias("ExcelChartDataLabel")]
public sealed class SetOfficeExcelChartDataLabelCommand : PSCmdlet {
    /// <summary>Chart to update.</summary>
    [Parameter(ParameterSetName = "Default", Mandatory = true, Position = 0, ValueFromPipeline = true)]
    public ExcelChart Chart { get; set; } = null!;

    /// <summary>Show values in labels.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter ShowValue { get; set; }

    /// <summary>Show category names in labels.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter ShowCategoryName { get; set; }

    /// <summary>Show series names in labels.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter ShowSeriesName { get; set; }

    /// <summary>Show percentages in labels.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter ShowPercent { get; set; }

    /// <summary>Show legend keys in labels.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter ShowLegendKey { get; set; }

    /// <summary>Optional data label position.</summary>
    [Parameter(ParameterSetName = "Default")]
    public OfficeChartDataLabelPosition? Position { get; set; }

    /// <summary>Optional number format code.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string? NumberFormat { get; set; }

    /// <summary>Keep number formatting linked to the source cells.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter SourceLinked { get; set; }

    /// <summary>Font family name.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string? Font { get; set; }

    /// <summary>Font size, interpreted as points for numbers; explicit physical units are accepted.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeLengthTransformation]
    public OfficeLength? Size { get; set; }

    /// <summary>Text or rule colour expressed as a literal OfficeColor or theme role.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? Color { get; set; }

    /// <summary>Applies bold text. An explicit false value clears bold; omission preserves the native value.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter Bold { get; set; }

    /// <summary>Applies italic text. An explicit false value clears italic; omission preserves the native value.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter Italic { get; set; }

    /// <summary>Optional label fill color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? FillColor { get; set; }

    /// <summary>Remove label fill.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter NoFill { get; set; }

    /// <summary>Optional label line color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? LineColor { get; set; }

    /// <summary>Physical line width; numbers are interpreted as points.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeLengthTransformation]
    public OfficeLength? LineWidth { get; set; }

    /// <summary>Remove label border.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter NoLine { get; set; }

    /// <summary>Emits the native result after the operation.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (Position.HasValue || NumberFormat != null || new[] { nameof(ShowValue), nameof(ShowCategoryName), nameof(ShowSeriesName), nameof(ShowLegendKey), nameof(ShowPercent), nameof(SourceLinked) }.Any(MyInvocation.BoundParameters.ContainsKey))
            Chart.SetDataLabels(OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(ShowValue)), OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(ShowCategoryName)), OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(ShowSeriesName)), OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(ShowLegendKey)), OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(ShowPercent)), Position, NumberFormat, OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(SourceLinked)));
        if (Size.HasValue || Color.HasValue || Font != null || MyInvocation.BoundParameters.ContainsKey(nameof(Bold)) || MyInvocation.BoundParameters.ContainsKey(nameof(Italic)))
            Chart.SetDataLabelTextStyle(Size, OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(Bold)), OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(Italic)), Color, Font);
        if (FillColor.HasValue || LineColor.HasValue || LineWidth.HasValue || NoFill || NoLine)
            Chart.SetDataLabelShapeStyle(FillColor, LineColor, LineWidth, NoFill, NoLine);
        if (PassThru) WriteObject(Chart);
    }

}
