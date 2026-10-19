using System;
using System.Management.Automation;
using OfficeIMO.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Configures fill and line styling for a single Excel chart data point.</summary>
/// <example>
///   <summary>Highlight one data point.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$chart | Set-OfficeExcelChartPoint -SeriesName 'Revenue' -PointIndex 1 -FillColor '#C00000'</code>
///   <para>Applies a point-specific fill override to the second point in the Revenue series.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelChartPoint", DefaultParameterSetName = ParameterSetIndex)]
[Alias("ExcelChartPoint")]
[OutputType(typeof(ExcelChart))]
public sealed class SetOfficeExcelChartPointCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
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

    /// <summary>Zero-based data point index within the series.</summary>
    [Parameter(Mandatory = true)]
    public uint PointIndex { get; set; }

    /// <summary>Point fill color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation(UseExcelArgb = true)]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? FillColor { get; set; }

    /// <summary>Point line color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation(UseExcelArgb = true)]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? LineColor { get; set; }

    /// <summary>Point line width in points.</summary>
    [Parameter]
    public double? LineWidthPoints { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        try {
            bool hasFill = !string.IsNullOrWhiteSpace(FillColor);
            bool hasLine = !string.IsNullOrWhiteSpace(LineColor);
            if (!hasFill && !hasLine && !LineWidthPoints.HasValue) {
                throw new PSArgumentException("Specify FillColor, LineColor, or LineWidthPoints to style the chart point.");
            }
            if (LineWidthPoints.HasValue && !hasLine) {
                throw new PSArgumentException("LineColor is required when LineWidthPoints is used.");
            }

            if (ParameterSetName == ParameterSetSeriesName) {
                Chart.SetDataPointColor(SeriesName, PointIndex, FillColor, LineColor, LineWidthPoints, IgnoreCase);
            } else {
                Chart.SetDataPointColor(SeriesIndex, PointIndex, FillColor, LineColor, LineWidthPoints);
            }

            WritePassThru(Chart);
        } catch (Exception ex) {
            WriteError(new ErrorRecord(ex, "ExcelChartPointFailed", ErrorCategory.InvalidOperation, Chart));
        }
    }
}
