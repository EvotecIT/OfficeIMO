using System;
using System.Management.Automation;
using OfficeIMO.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using OfficeIMO.Excel;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Configures Excel chart series colors, line style, and markers.</summary>
/// <example>
///   <summary>Format a series by name.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$chart | Set-OfficeExcelChartSeries -SeriesName 'Revenue' -FillColor '#4472C4' -LineColor '#1F4E79' -MarkerStyle Circle -MarkerSize 6</code>
///   <para>Applies fill, line, and marker settings to the Revenue series.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelChartSeries", DefaultParameterSetName = ParameterSetIndex)]
[Alias("ExcelChartSeries")]
[OutputType(typeof(ExcelChart))]
public sealed class SetOfficeExcelChartSeriesCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
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

    /// <summary>Series fill color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation(UseExcelArgb = true)]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? FillColor { get; set; }

    /// <summary>Series line color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation(UseExcelArgb = true)]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? LineColor { get; set; }

    /// <summary>Series line width in points.</summary>
    [Parameter]
    public double? LineWidthPoints { get; set; }

    /// <summary>Marker style name.</summary>
    [Parameter]
    public OfficeChartMarkerShape? MarkerStyle { get; set; }

    /// <summary>Marker size.</summary>
    [Parameter]
    public int? MarkerSize { get; set; }

    /// <summary>Marker fill color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation(UseExcelArgb = true)]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? MarkerFillColor { get; set; }

    /// <summary>Marker line color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation(UseExcelArgb = true)]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? MarkerLineColor { get; set; }

    /// <summary>Marker line width in points.</summary>
    [Parameter]
    public double? MarkerLineWidthPoints { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        try {
            var fillColor = FillColor;
            if (!string.IsNullOrWhiteSpace(fillColor)) {
                if (ParameterSetName == ParameterSetSeriesName) {
                    Chart.SetSeriesFillColor(SeriesName, fillColor!, IgnoreCase);
                } else {
                    Chart.SetSeriesFillColor(SeriesIndex, fillColor!);
                }
            }

            var lineColor = LineColor;
            if (!string.IsNullOrWhiteSpace(lineColor) || LineWidthPoints.HasValue) {
                if (string.IsNullOrWhiteSpace(lineColor)) {
                    throw new PSArgumentException("LineColor is required when LineWidthPoints is used because the current OfficeIMO chart API applies series line width together with a line color.");
                }

                if (ParameterSetName == ParameterSetSeriesName) {
                    Chart.SetSeriesLineColor(SeriesName, lineColor!, LineWidthPoints, IgnoreCase);
                } else {
                    Chart.SetSeriesLineColor(SeriesIndex, lineColor!, LineWidthPoints);
                }
            }

            var markerStyle = MarkerStyle ?? OfficeChartMarkerShape.Circle;
            bool markerRequested = MarkerStyle.HasValue || MarkerSize.HasValue ||
                !string.IsNullOrWhiteSpace(MarkerFillColor) || !string.IsNullOrWhiteSpace(MarkerLineColor) || MarkerLineWidthPoints.HasValue;
            if (markerRequested) {
                if (ParameterSetName == ParameterSetSeriesName) {
                    Chart.SetSeriesMarker(SeriesName, markerStyle, MarkerSize, MarkerFillColor, MarkerLineColor, MarkerLineWidthPoints, IgnoreCase);
                } else {
                    Chart.SetSeriesMarker(SeriesIndex, markerStyle, MarkerSize, MarkerFillColor, MarkerLineColor, MarkerLineWidthPoints);
                }
            }

            WritePassThru(Chart);
        } catch (Exception ex) {
            WriteError(new ErrorRecord(ex, "ExcelChartSeriesFailed", ErrorCategory.InvalidOperation, Chart));
        }
    }
}
