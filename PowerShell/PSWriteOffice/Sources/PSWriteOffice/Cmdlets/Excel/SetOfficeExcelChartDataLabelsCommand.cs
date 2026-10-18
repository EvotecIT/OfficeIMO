using System;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Configures data labels and optional styling for an Excel chart.</summary>
/// <example>
///   <summary>Add outside-end value labels.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$chart | Set-OfficeExcelChartDataLabels -ShowValue $true -Position OutsideEnd</code>
///   <para>Adds value labels to the chart and returns the chart for further formatting.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelChartDataLabels")]
[OutputType(typeof(ExcelChart))]
public sealed class SetOfficeExcelChartDataLabelsCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    /// <summary>Chart to update.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public ExcelChart Chart { get; set; } = null!;

    /// <summary>Show values in labels.</summary>
    [Parameter]
    public bool ShowValue { get; set; } = true;

    /// <summary>Show category names in labels.</summary>
    [Parameter]
    public bool ShowCategoryName { get; set; }

    /// <summary>Show series names in labels.</summary>
    [Parameter]
    public bool ShowSeriesName { get; set; }

    /// <summary>Show legend keys in labels.</summary>
    [Parameter]
    public bool ShowLegendKey { get; set; }

    /// <summary>Show percentages in labels.</summary>
    [Parameter]
    public bool ShowPercent { get; set; }

    /// <summary>Optional data label position.</summary>
    [Parameter]
    public OfficeChartDataLabelPosition? Position { get; set; }

    /// <summary>Optional number format code.</summary>
    [Parameter]
    public string? NumberFormat { get; set; }

    /// <summary>Keep number formatting linked to the source cells.</summary>
    [Parameter]
    public bool SourceLinked { get; set; }

    /// <summary>Optional label font size in points.</summary>
    [Parameter]
    public double? FontSizePoints { get; set; }

    /// <summary>Optional bold setting for label text.</summary>
    [Parameter]
    public bool? Bold { get; set; }

    /// <summary>Optional italic setting for label text.</summary>
    [Parameter]
    public bool? Italic { get; set; }

    /// <summary>Optional label text color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation(UseExcelArgb = true)]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? Color { get; set; }

    /// <summary>Optional label font name.</summary>
    [Parameter]
    public string? FontName { get; set; }

    /// <summary>Optional label fill color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation(UseExcelArgb = true)]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? FillColor { get; set; }

    /// <summary>Optional label line color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation(UseExcelArgb = true)]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? LineColor { get; set; }

    /// <summary>Optional label border width in points.</summary>
    [Parameter]
    public double? LineWidthPoints { get; set; }

    /// <summary>Remove label fill.</summary>
    [Parameter]
    public SwitchParameter NoFill { get; set; }

    /// <summary>Remove label border.</summary>
    [Parameter]
    public SwitchParameter NoLine { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        try {
            Chart.SetDataLabels(ShowValue, ShowCategoryName, ShowSeriesName, ShowLegendKey, ShowPercent, Position, NumberFormat, SourceLinked);

            if (FontSizePoints.HasValue || Bold.HasValue || Italic.HasValue ||
                !string.IsNullOrWhiteSpace(Color) || !string.IsNullOrWhiteSpace(FontName)) {
                Chart.SetDataLabelTextStyle(FontSizePoints, Bold, Italic, Color, FontName);
            }

            if (!string.IsNullOrWhiteSpace(FillColor) || !string.IsNullOrWhiteSpace(LineColor) ||
                LineWidthPoints.HasValue || NoFill.IsPresent || NoLine.IsPresent) {
                Chart.SetDataLabelShapeStyle(FillColor, LineColor, LineWidthPoints, NoFill.IsPresent, NoLine.IsPresent);
            }

            WritePassThru(Chart);
        } catch (Exception ex) {
            WriteError(new ErrorRecord(ex, "ExcelChartDataLabelsFailed", ErrorCategory.InvalidOperation, Chart));
        }
    }

}
