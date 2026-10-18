using System;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Configures legend visibility and styling for an Excel chart.</summary>
/// <example>
///   <summary>Move the legend to the right.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$chart | Set-OfficeExcelChartLegend -Position Right</code>
///   <para>Shows the legend on the right side of the chart and returns the chart for chaining.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelChartLegend")]
[OutputType(typeof(ExcelChart))]
public sealed class SetOfficeExcelChartLegendCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    /// <summary>Chart to update.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public ExcelChart Chart { get; set; } = null!;

    /// <summary>Legend position.</summary>
    [Parameter]
    public OfficeChartLegendPosition Position { get; set; } = OfficeChartLegendPosition.Right;

    /// <summary>Overlay the legend on the chart area.</summary>
    [Parameter]
    public bool Overlay { get; set; }

    /// <summary>Hide the legend instead of positioning it.</summary>
    [Parameter]
    public SwitchParameter Hide { get; set; }

    /// <summary>Optional legend font size in points.</summary>
    [Parameter]
    public double? FontSizePoints { get; set; }

    /// <summary>Optional bold setting for legend text.</summary>
    [Parameter]
    public bool? Bold { get; set; }

    /// <summary>Optional italic setting for legend text.</summary>
    [Parameter]
    public bool? Italic { get; set; }

    /// <summary>Optional legend text color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation(UseExcelArgb = true)]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? Color { get; set; }

    /// <summary>Optional legend font name.</summary>
    [Parameter]
    public string? FontName { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        try {
            if (Hide.IsPresent) {
                Chart.HideLegend();
            } else {
                Chart.SetLegend(Position, Overlay);
            }

            if (FontSizePoints.HasValue || Bold.HasValue || Italic.HasValue ||
                !string.IsNullOrWhiteSpace(Color) || !string.IsNullOrWhiteSpace(FontName)) {
                Chart.SetLegendTextStyle(FontSizePoints, Bold, Italic, Color, FontName);
            }

            WritePassThru(Chart);
        } catch (Exception ex) {
            WriteError(new ErrorRecord(ex, "ExcelChartLegendFailed", ErrorCategory.InvalidOperation, Chart));
        }
    }

}
