using System;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Configures common Excel chart axis titles, formats, scale, and gridlines.</summary>
/// <example>
/// <summary>Title the value axis</summary>
/// <code>New-OfficeExcel -Path .\Reports\Regions.xlsx {
///     ExcelSheet 'Summary' {
///         ExcelTable (Import-OfficeCsv -Path .\Data\Region-Revenue.csv) -Name RegionSummary
///         ExcelChart -Table RegionSummary -Address F2 -Type Column -PassThru | ExcelChartAxis -ValueTitle 'Revenue (USD)'
///     }
/// }</code>
/// <para>Adds a value axis title to a new chart.</para>
/// </example>
/// <example>
/// <summary>Fix the scale of an availability chart</summary>
/// <code>New-OfficeExcel -Path .\Reports\Availability.xlsx {
///     ExcelSheet 'Uptime' {
///         $chart = ExcelChart -InputObject (Import-OfficeCsv -Path .\Data\Site-Uptime.csv) -Category Site -Value Uptime -Address A1 -Type Column -PassThru
///         ExcelChartAxis -Chart $chart -ValueMinimum 0.95 -ValueMaximum 1 -ValueMajorUnit 0.01 -ValueNumberFormat '0%' -CategoryTitle 'Site'
///     }
/// }</code>
/// <para>Scales the value axis from 95 to 100 percent with a percentage format, so small differences in site uptime are visible.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelChartAxis", DefaultParameterSetName = "Default")]
[Alias("ExcelChartAxis")]
[OutputType(typeof(ExcelChart))]
public sealed class SetOfficeExcelChartAxisCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    /// <summary>Chart to update.</summary>
    [Parameter(ParameterSetName = "Default", Mandatory = true, Position = 0, ValueFromPipeline = true)]
    public ExcelChart Chart { get; set; } = null!;

    /// <summary>Axis group to configure.</summary>
    [Parameter(ParameterSetName = "Default")]
    public OfficeChartAxisGroup AxisGroup { get; set; } = OfficeChartAxisGroup.Primary;

    /// <summary>Category axis title.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string? CategoryTitle { get; set; }

    /// <summary>Value axis title.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string? ValueTitle { get; set; }

    /// <summary>Category axis number format code.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string? CategoryNumberFormat { get; set; }

    /// <summary>Value axis number format code.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string? ValueNumberFormat { get; set; }

    /// <summary>Keep axis number formats linked to source cells.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter SourceLinked { get; set; }

    /// <summary>Value axis minimum.</summary>
    [Parameter(ParameterSetName = "Default")]
    public double? ValueMinimum { get; set; }

    /// <summary>Value axis maximum.</summary>
    [Parameter(ParameterSetName = "Default")]
    public double? ValueMaximum { get; set; }

    /// <summary>Value axis major unit.</summary>
    [Parameter(ParameterSetName = "Default")]
    public double? ValueMajorUnit { get; set; }

    /// <summary>Value axis minor unit.</summary>
    [Parameter(ParameterSetName = "Default")]
    public double? ValueMinorUnit { get; set; }

    /// <summary>Category/date axis minimum.</summary>
    [Parameter(ParameterSetName = "Default")]
    public double? CategoryMinimum { get; set; }

    /// <summary>Category/date axis maximum.</summary>
    [Parameter(ParameterSetName = "Default")]
    public double? CategoryMaximum { get; set; }

    /// <summary>Category/date axis major unit.</summary>
    [Parameter(ParameterSetName = "Default")]
    public double? CategoryMajorUnit { get; set; }

    /// <summary>Category/date axis minor unit.</summary>
    [Parameter(ParameterSetName = "Default")]
    public double? CategoryMinorUnit { get; set; }

    /// <summary>Show category major gridlines.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter ShowCategoryMajorGridlines { get; set; }

    /// <summary>Show category minor gridlines.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter ShowCategoryMinorGridlines { get; set; }

    /// <summary>Show value major gridlines.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter ShowValueMajorGridlines { get; set; }

    /// <summary>Show value minor gridlines.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter ShowValueMinorGridlines { get; set; }

    /// <summary>Optional category gridline color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? CategoryGridlineColor { get; set; }

    /// <summary>Optional value gridline color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor? ValueGridlineColor { get; set; }

    /// <summary>Optional gridline width in points.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeLengthTransformation]
    public OfficeLength? GridlineWidth { get; set; }

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

            bool categoryRequest = MyInvocation.BoundParameters.ContainsKey(nameof(ShowCategoryMajorGridlines)) ||
                MyInvocation.BoundParameters.ContainsKey(nameof(ShowCategoryMinorGridlines)) || CategoryGridlineColor.HasValue;
            bool valueRequest = MyInvocation.BoundParameters.ContainsKey(nameof(ShowValueMajorGridlines)) ||
                MyInvocation.BoundParameters.ContainsKey(nameof(ShowValueMinorGridlines)) || ValueGridlineColor.HasValue;
            bool widthOnly = GridlineWidth.HasValue && !categoryRequest && !valueRequest;
            if (categoryRequest || widthOnly)
                Chart.SetCategoryAxisGridlines(OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(ShowCategoryMajorGridlines)),
                    OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(ShowCategoryMinorGridlines)), CategoryGridlineColor, GridlineWidth, AxisGroup);
            if (valueRequest || widthOnly)
                Chart.SetValueAxisGridlines(OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(ShowValueMajorGridlines)),
                    OfficeSwitch.Get(MyInvocation.BoundParameters, nameof(ShowValueMinorGridlines)), ValueGridlineColor, GridlineWidth, AxisGroup);

            WritePassThru(Chart);
        } catch (Exception ex) {
            WriteError(new ErrorRecord(ex, "ExcelChartAxisFailed", ErrorCategory.InvalidOperation, Chart));
        }
    }
}
