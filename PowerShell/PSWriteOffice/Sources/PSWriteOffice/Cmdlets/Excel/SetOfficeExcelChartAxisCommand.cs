using System;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Configures common Excel chart axis titles, formats, scale, and gridlines.</summary>
/// <example>
///   <summary>Format value axis and gridlines.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$chart | Set-OfficeExcelChartAxis -CategoryTitle 'Month' -ValueTitle 'Revenue' -ValueNumberFormat '$#,##0' -ValueMinimum 0 -ValueMajorUnit 100 -ShowValueMajorGridlines</code>
///   <para>Sets axis titles, value formatting, scale, and major value gridlines.</para>
/// </example>
/// <example>
///   <summary>Use a scaled date/category axis for a line chart.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$chart |
///     Set-OfficeExcelChartAxis -CategoryTitle 'Week' -CategoryNumberFormat 'yyyy-mm-dd' -CategoryMinimum 46000 -CategoryMaximum 46090 -CategoryMajorUnit 14 -CategoryMinorUnit 7</code>
///   <para>Configures the chart's X axis through OfficeIMO's category/date-axis scale support while keeping the workbook chart object on the pipeline.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelChartAxis")]
[Alias("ExcelChartAxis")]
[OutputType(typeof(ExcelChart))]
public sealed class SetOfficeExcelChartAxisCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    /// <summary>Chart to update.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    public ExcelChart Chart { get; set; } = null!;

    /// <summary>Axis group to configure.</summary>
    [Parameter]
    public OfficeChartAxisGroup AxisGroup { get; set; } = OfficeChartAxisGroup.Primary;

    /// <summary>Category axis title.</summary>
    [Parameter]
    public string? CategoryTitle { get; set; }

    /// <summary>Value axis title.</summary>
    [Parameter]
    public string? ValueTitle { get; set; }

    /// <summary>Category axis number format code.</summary>
    [Parameter]
    public string? CategoryNumberFormat { get; set; }

    /// <summary>Value axis number format code.</summary>
    [Parameter]
    public string? ValueNumberFormat { get; set; }

    /// <summary>Keep axis number formats linked to source cells.</summary>
    [Parameter]
    public bool SourceLinked { get; set; }

    /// <summary>Value axis minimum.</summary>
    [Parameter]
    public double? ValueMinimum { get; set; }

    /// <summary>Value axis maximum.</summary>
    [Parameter]
    public double? ValueMaximum { get; set; }

    /// <summary>Value axis major unit.</summary>
    [Parameter]
    public double? ValueMajorUnit { get; set; }

    /// <summary>Value axis minor unit.</summary>
    [Parameter]
    public double? ValueMinorUnit { get; set; }

    /// <summary>Category/date axis minimum.</summary>
    [Parameter]
    public double? CategoryMinimum { get; set; }

    /// <summary>Category/date axis maximum.</summary>
    [Parameter]
    public double? CategoryMaximum { get; set; }

    /// <summary>Category/date axis major unit.</summary>
    [Parameter]
    public double? CategoryMajorUnit { get; set; }

    /// <summary>Category/date axis minor unit.</summary>
    [Parameter]
    public double? CategoryMinorUnit { get; set; }

    /// <summary>Show category major gridlines.</summary>
    [Parameter]
    public SwitchParameter ShowCategoryMajorGridlines { get; set; }

    /// <summary>Show category minor gridlines.</summary>
    [Parameter]
    public SwitchParameter ShowCategoryMinorGridlines { get; set; }

    /// <summary>Show value major gridlines.</summary>
    [Parameter]
    public SwitchParameter ShowValueMajorGridlines { get; set; }

    /// <summary>Show value minor gridlines.</summary>
    [Parameter]
    public SwitchParameter ShowValueMinorGridlines { get; set; }

    /// <summary>Optional category gridline color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation(UseExcelArgb = true)]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? CategoryGridlineColor { get; set; }

    /// <summary>Optional value gridline color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation(UseExcelArgb = true)]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? ValueGridlineColor { get; set; }

    /// <summary>Optional gridline width in points.</summary>
    [Parameter]
    public double? GridlineWidthPoints { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        try {
            var categoryTitle = CategoryTitle;
            if (!string.IsNullOrWhiteSpace(categoryTitle)) {
                Chart.SetCategoryAxisTitle(categoryTitle!, AxisGroup);
            }

            var valueTitle = ValueTitle;
            if (!string.IsNullOrWhiteSpace(valueTitle)) {
                Chart.SetValueAxisTitle(valueTitle!, AxisGroup);
            }

            var categoryNumberFormat = CategoryNumberFormat;
            if (!string.IsNullOrWhiteSpace(categoryNumberFormat)) {
                Chart.SetCategoryAxisNumberFormat(categoryNumberFormat!, SourceLinked, AxisGroup);
            }

            var valueNumberFormat = ValueNumberFormat;
            if (!string.IsNullOrWhiteSpace(valueNumberFormat)) {
                Chart.SetValueAxisNumberFormat(valueNumberFormat!, SourceLinked, AxisGroup);
            }

            if (ValueMinimum.HasValue || ValueMaximum.HasValue || ValueMajorUnit.HasValue || ValueMinorUnit.HasValue) {
                Chart.SetValueAxisScale(ValueMinimum, ValueMaximum, ValueMajorUnit, ValueMinorUnit, axisGroup: AxisGroup);
            }

            if (CategoryMinimum.HasValue || CategoryMaximum.HasValue || CategoryMajorUnit.HasValue || CategoryMinorUnit.HasValue) {
                Chart.SetCategoryAxisScale(CategoryMinimum, CategoryMaximum, CategoryMajorUnit, CategoryMinorUnit, axisGroup: AxisGroup);
            }

            var categoryGridlineColor = CategoryGridlineColor;
            var valueGridlineColor = ValueGridlineColor;

            bool categoryGridlinesRequested = ShowCategoryMajorGridlines.IsPresent ||
                ShowCategoryMinorGridlines.IsPresent ||
                !string.IsNullOrWhiteSpace(categoryGridlineColor);
            bool valueGridlinesRequested = ShowValueMajorGridlines.IsPresent ||
                ShowValueMinorGridlines.IsPresent ||
                !string.IsNullOrWhiteSpace(valueGridlineColor);
            bool widthOnlyRequest = GridlineWidthPoints.HasValue && !categoryGridlinesRequested && !valueGridlinesRequested;

            bool categoryStyleRequested = !string.IsNullOrWhiteSpace(categoryGridlineColor) ||
                (GridlineWidthPoints.HasValue && (categoryGridlinesRequested || widthOnlyRequest));
            if (categoryGridlinesRequested || widthOnlyRequest) {
                bool showMajor = ShowCategoryMajorGridlines.IsPresent || ShowCategoryMinorGridlines.IsPresent || categoryStyleRequested;
                Chart.SetCategoryAxisGridlines(showMajor,
                    ShowCategoryMinorGridlines.IsPresent, categoryGridlineColor, GridlineWidthPoints, AxisGroup);
            }

            bool valueStyleRequested = !string.IsNullOrWhiteSpace(valueGridlineColor) ||
                (GridlineWidthPoints.HasValue && (valueGridlinesRequested || widthOnlyRequest));
            if (valueGridlinesRequested || widthOnlyRequest) {
                bool showMajor = ShowValueMajorGridlines.IsPresent || ShowValueMinorGridlines.IsPresent || valueStyleRequested;
                Chart.SetValueAxisGridlines(showMajor,
                    ShowValueMinorGridlines.IsPresent, valueGridlineColor, GridlineWidthPoints, AxisGroup);
            }

            WritePassThru(Chart);
        } catch (Exception ex) {
            WriteError(new ErrorRecord(ex, "ExcelChartAxisFailed", ErrorCategory.InvalidOperation, Chart));
        }
    }
}
