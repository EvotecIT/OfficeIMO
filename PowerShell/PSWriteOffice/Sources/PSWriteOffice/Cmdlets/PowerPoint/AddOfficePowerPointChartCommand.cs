using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Adds a chart to a PowerPoint slide.</summary>
/// <para>Supports default chart data or object-based category/series mappings for standard and scatter charts.</para>
/// <example>
///   <summary>Add a clustered column chart from objects.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$rows = @(
///     [pscustomobject]@{ Month = 'Jan'; Sales = 42; Profit = 9 }
///     [pscustomobject]@{ Month = 'Feb'; Sales = 55; Profit = 13 }
/// )
/// New-OfficePowerPoint -Path .\Examples\Documents\PowerPointChart.pptx {
///     $slide = Add-OfficePowerPointSlide -PassThru
///     Add-OfficePowerPointChart -Slide $slide -InputObject $rows -Category Month -Value Sales,Profit -Title 'Monthly performance'
/// }</code>
///   <para>Creates a clustered column chart using Month for categories and Sales/Profit as series.</para>
/// </example>
/// <example>
///   <summary>Add a scatter chart from numeric properties.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$rows = @(
///     [pscustomobject]@{ Quarter = 1; Revenue = 20 }
///     [pscustomobject]@{ Quarter = 2; Revenue = 34 }
/// )
/// New-OfficePowerPoint -Path .\Examples\Documents\PowerPointScatter.pptx {
///     $slide = Add-OfficePowerPointSlide -PassThru
///     Add-OfficePowerPointChart -Slide $slide -Type Scatter -InputObject $rows -Category Quarter -Value Revenue -Title 'Revenue trend'
/// }</code>
///   <para>Creates a scatter chart using Quarter on the X axis and Revenue on the Y axis.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficePowerPointChart", DefaultParameterSetName = ParameterSetSample)]
[Alias("PowerPointChart")]
[OutputType(typeof(PowerPointChart))]
public sealed class AddOfficePowerPointChartCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    private const string ParameterSetSample = "Sample";
    private const string ParameterSetInputObject = "InputObject";
    private readonly List<object> _input = new();

    /// <summary>Target slide that will receive the chart (optional inside DSL).</summary>
    [Parameter, PowerPointSlideTransformation]
    public PowerPointSlideSelector? Slide { get; set; }

    /// <summary>Chart type to create.</summary>
    [Parameter(ParameterSetName = ParameterSetSample)]
    [Parameter(ParameterSetName = ParameterSetInputObject)]
    public OfficeChartType Type { get; set; } = OfficeChartType.Column;

    /// <summary>Source objects used to build chart data.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetInputObject)]
    public object[] InputObject { get; set; } = Array.Empty<object>();

    /// <summary>Property used for category labels, or numeric X values on scatter charts.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetInputObject)]
    public string Category { get; set; } = string.Empty;

    /// <summary>Properties used as numeric series, or Y series on scatter charts.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetInputObject)]
    public string[] Value { get; set; } = Array.Empty<string>();

    /// <summary>Left offset from the slide origin; numbers mean points.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength X { get; set; } = OfficeLength.Points(40);

    /// <summary>Top offset from the slide origin; numbers mean points.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength Y { get; set; } = OfficeLength.Points(120);

    /// <summary>Chart width; numbers mean points.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength Width { get; set; } = OfficeLength.Points(420);

    /// <summary>Chart height; numbers mean points.</summary>
    [Parameter, OfficeLengthTransformation]
    public OfficeLength Height { get; set; } = OfficeLength.Points(240);

    /// <summary>One-based start column in the native twelve-column grid; explicit X overrides it.</summary>
    [Parameter, ValidateRange(1, 12)]
    public int? Column { get; set; }

    /// <summary>Number of grid columns; defaults to twelve, and explicit Width overrides it.</summary>
    [Parameter, ValidateRange(1, 12)]
    public int? Span { get; set; }

    /// <summary>Optional chart title.</summary>
    [Parameter]
    public string? Title { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (ParameterSetName == ParameterSetInputObject && InputObject != null) _input.AddRange(InputObject);
    }

    /// <inheritdoc />
    protected override void EndProcessing() {
        InputObject = _input.ToArray();
        PowerPointChart? chart = null;
        try {
            if (Width <= OfficeLength.Zero) {
                throw new ArgumentOutOfRangeException(nameof(Width), "Width must be greater than 0.");
            }

            if (Height <= OfficeLength.Zero) {
                throw new ArgumentOutOfRangeException(nameof(Height), "Height must be greater than 0.");
            }

            var target = PowerPointManualPlacement.Resolve(Slide);
            var scope = target.Scope;
            var slide = target.Slide;
            var horizontal = PowerPointManualPlacement.Horizontal(slide,
                MyInvocation.BoundParameters.ContainsKey(nameof(X)) ? X : null,
                MyInvocation.BoundParameters.ContainsKey(nameof(Width)) ? Width : null, Column, Span);
            var data = ParameterSetName == ParameterSetSample ? BuildSampleData() : BuildChartData();
            chart = slide.AddChart(Type, data, new PowerPointChartOptions {
                Left = horizontal.X ?? X, Top = Y, Width = horizontal.Width ?? Width, Height = Height
            });

            if (!string.IsNullOrWhiteSpace(Title)) {
                chart.SetTitle(Title!);
            }

            if (scope != null && ReferenceEquals(scope.NativeSlide, slide)) {
                if (Slide != null || Column.HasValue || Span.HasValue || PowerPointManualPlacement.HasExplicitBounds(MyInvocation.BoundParameters.Keys)) scope.TrackFixedFigure(chart);
                else {
                    if (!chart.TryGetOfficeSnapshot(out var snapshot)) throw new InvalidOperationException("The created chart has no supported data snapshot.");
                    scope.AddFigure(new PowerPointLayoutChart(snapshot.ChartKind, snapshot.Data) { NativeChart = chart }, chart);
                }
            }

            WritePassThru(chart);
        } catch (Exception ex) {
            chart?.Remove();
            if (ex is PipelineStoppedException or ActionPreferenceStopException) throw;
            if (PowerPointContentScope.Current != null) ThrowTerminatingError(OfficeError.Create(ex, "invalid-argument", InputObject, ErrorCategory.InvalidArgument));
            WriteError(new ErrorRecord(ex, "PowerPointAddChartFailed", ErrorCategory.InvalidOperation, Slide));
        }
    }

    private OfficeChartData BuildSampleData() => Type == OfficeChartType.Scatter
        ? new OfficeChartData(new[] { "1", "2", "3" },
            new[] { new OfficeSeries("Series 1", new[] { 12d, 18d, 14d }, new[] { 1d, 2d, 3d }) })
        : new OfficeChartData(new[] { "Q1", "Q2", "Q3" },
            new[] { new OfficeSeries("Series 1", new[] { 12d, 18d, 14d }) });

    private OfficeChartData BuildChartData() {
        var items = EnsureData();
        bool scatter = Type == OfficeChartType.Scatter;
        var xValues = scatter
            ? items.Select(item => ConvertToDouble(GetPropertyValue(item, Category), Category)).ToArray()
            : null;
        var categories = xValues != null
            ? xValues.Select(value => value.ToString(CultureInfo.InvariantCulture)).ToArray()
            : items.Select(item => ConvertToString(GetPropertyValue(item, Category), Category)).ToArray();
        var series = Value.Select(property => new OfficeSeries(property,
            items.Select(item => ConvertToDouble(GetPropertyValue(item, property), property)).ToArray(), xValues)).ToArray();
        return new OfficeChartData(categories, series);
    }

    private object[] EnsureData() {
        if (InputObject == null || InputObject.Length == 0) {
            throw new PSArgumentException("Provide at least one data item.", nameof(InputObject));
        }

        return InputObject;
    }

    private static object? GetPropertyValue(object item, string propertyName) {
        if (item == null) {
            throw new PSArgumentException("Chart data items cannot be null.");
        }

        if (string.IsNullOrWhiteSpace(propertyName)) {
            throw new PSArgumentException("Property name cannot be empty.", nameof(propertyName));
        }

        if (item is IDictionary dictionary) {
            foreach (DictionaryEntry entry in dictionary) {
                if (entry.Key is string key && string.Equals(key, propertyName, StringComparison.OrdinalIgnoreCase)) {
                    return UnwrapValue(entry.Value);
                }
            }
        }

        var property = PSObject.AsPSObject(item).Properties[propertyName];
        if (property == null) {
            throw new PSArgumentException($"Property '{propertyName}' was not found on chart data item.");
        }

        return UnwrapValue(property.Value);
    }

    private static object? UnwrapValue(object? value) {
        // Select-Object calculated properties can retain a PSObject around a scalar.
        while (value is PSObject wrapper && !ReferenceEquals(value, wrapper.BaseObject)) {
            value = wrapper.BaseObject;
        }

        return value;
    }

    private static string ConvertToString(object? value, string propertyName) {
        if (value == null) {
            throw new PSArgumentException($"Property '{propertyName}' cannot be null.");
        }

        var text = Convert.ToString(value, CultureInfo.InvariantCulture);
        if (string.IsNullOrWhiteSpace(text)) {
            throw new PSArgumentException($"Property '{propertyName}' cannot be empty.");
        }

        return text;
    }

    private static double ConvertToDouble(object? value, string propertyName) {
        if (value == null) {
            throw new PSArgumentException($"Property '{propertyName}' cannot be null.");
        }

        try {
            return Convert.ToDouble(value, CultureInfo.InvariantCulture);
        } catch (Exception) {
            throw new PSArgumentException($"Property '{propertyName}' must be numeric.", propertyName);
        }
    }
}
