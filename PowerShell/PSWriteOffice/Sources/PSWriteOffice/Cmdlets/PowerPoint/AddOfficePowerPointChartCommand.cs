using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.PowerPoint;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Chart types supported by Add-OfficePowerPointChart.</summary>
public enum PowerPointChartType {
    /// <summary>Clustered column chart.</summary>
    ClusteredColumn,
    /// <summary>Line chart.</summary>
    Line,
    /// <summary>Pie chart.</summary>
    Pie,
    /// <summary>Doughnut chart.</summary>
    Doughnut,
    /// <summary>Scatter chart.</summary>
    Scatter
}

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
///     Add-OfficePowerPointChart -Slide $slide -Type Scatter -InputObject $rows -XProperty Quarter -YProperty Revenue -Title 'Revenue trend'
/// }</code>
///   <para>Creates a scatter chart using Quarter on the X axis and Revenue on the Y axis.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficePowerPointChart", DefaultParameterSetName = ParameterSetDefault)]
[Alias("PowerPointChart")]
[OutputType(typeof(PowerPointChart))]
public sealed class AddOfficePowerPointChartCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    private const string ParameterSetDefault = "Default";
    private const string ParameterSetCategorical = "Categorical";
    private const string ParameterSetScatter = "Scatter";

    /// <summary>Target slide that will receive the chart (optional inside DSL).</summary>
    [Parameter(ValueFromPipeline = true)]
    public PowerPointSlide? Slide { get; set; }

    /// <summary>Chart type to create.</summary>
    [Parameter]
    public PowerPointChartType Type { get; set; } = PowerPointChartType.ClusteredColumn;

    /// <summary>Source objects used to build chart data.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetCategorical)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetScatter)]
    [Alias("Data")]
    public object[] InputObject { get; set; } = Array.Empty<object>();

    /// <summary>Property name used for category labels on standard charts.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetCategorical)]
    public string Category { get; set; } = string.Empty;

    /// <summary>Property names used as numeric series on standard charts.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetCategorical)]
    public string[] Value { get; set; } = Array.Empty<string>();

    /// <summary>Property name used for the X axis on scatter charts.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetScatter)]
    public string XProperty { get; set; } = string.Empty;

    /// <summary>Property names used as numeric Y series on scatter charts.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetScatter)]
    public string[] YProperty { get; set; } = Array.Empty<string>();

    /// <summary>Left offset in points from the slide origin.</summary>
    [Parameter]
    public double X { get; set; } = 40;

    /// <summary>Top offset in points from the slide origin.</summary>
    [Parameter]
    public double Y { get; set; } = 120;

    /// <summary>Chart width in points.</summary>
    [Parameter]
    public double Width { get; set; } = 420;

    /// <summary>Chart height in points.</summary>
    [Parameter]
    public double Height { get; set; } = 240;

    /// <summary>Optional chart title.</summary>
    [Parameter]
    public string? Title { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        try {
            if (Width <= 0) {
                throw new ArgumentOutOfRangeException(nameof(Width), "Width must be greater than 0.");
            }

            if (Height <= 0) {
                throw new ArgumentOutOfRangeException(nameof(Height), "Height must be greater than 0.");
            }

            var scope = Slide == null ? PowerPointContentScope.Current : null;
            var slide = Slide ?? scope?.NativeSlide ?? PowerPointDslContext.Require(this).RequireSlide();
            var chart = ParameterSetName switch {
                ParameterSetCategorical => AddCategoricalChart(slide),
                ParameterSetScatter => AddScatterChart(slide),
                _ => AddDefaultChart(slide)
            };

            if (!string.IsNullOrWhiteSpace(Title)) {
                chart.SetTitle(Title!);
            }

            if (scope != null) {
                if (!chart.TryGetOfficeSnapshot(out var snapshot)) throw new InvalidOperationException("The created chart has no supported data snapshot.");
                scope.AddFigure(new PowerPointLayoutChart(snapshot.ChartKind, snapshot.Data) { NativeChart = chart }, chart);
            }

            WritePassThru(chart);
        } catch (Exception ex) {
            if (PowerPointContentScope.Current != null) ThrowTerminatingError(OfficeError.Create(ex, "invalid-argument", InputObject, ErrorCategory.InvalidArgument));
            WriteError(new ErrorRecord(ex, "PowerPointAddChartFailed", ErrorCategory.InvalidOperation, Slide));
        }
    }

    private PowerPointChart AddDefaultChart(PowerPointSlide slide) {
        OfficeChartType kind = GetOfficeChartKind(Type);
        OfficeChartData data = kind == OfficeChartType.Scatter
            ? new OfficeChartData(
                new[] { "1", "2", "3" },
                new[] { new OfficeSeries("Series 1", new[] { 12d, 18d, 14d }, new[] { 1d, 2d, 3d }) })
            : new OfficeChartData(
                new[] { "Q1", "Q2", "Q3" },
                new[] { new OfficeSeries("Series 1", new[] { 12d, 18d, 14d }) });

        return slide.AddChart(kind, data, new OfficeIMO.PowerPoint.PowerPointChartOptions { Left = global::OfficeIMO.OfficeLength.Points(X), Top = global::OfficeIMO.OfficeLength.Points(Y), Width = global::OfficeIMO.OfficeLength.Points(Width), Height = global::OfficeIMO.OfficeLength.Points(Height) });
    }

    private PowerPointChart AddCategoricalChart(PowerPointSlide slide) {
        if (Type == PowerPointChartType.Scatter) {
            throw new PSArgumentException("Use -XProperty and -YProperty when -Type Scatter is selected.", nameof(Type));
        }

        if (Value == null || Value.Length == 0) {
            throw new PSArgumentException("Provide at least one -Value.", nameof(Value));
        }

        if ((Type == PowerPointChartType.Pie || Type == PowerPointChartType.Doughnut) && Value.Length > 1) {
            throw new PSArgumentException("Pie and Doughnut charts support only one series property.", nameof(Value));
        }

        return slide.AddChart(GetOfficeChartKind(Type), BuildChartData(), new PowerPointChartOptions { Left = global::OfficeIMO.OfficeLength.Points(X), Top = global::OfficeIMO.OfficeLength.Points(Y), Width = global::OfficeIMO.OfficeLength.Points(Width), Height = global::OfficeIMO.OfficeLength.Points(Height) });
    }

    private PowerPointChart AddScatterChart(PowerPointSlide slide) {
        if (Type != PowerPointChartType.Scatter) {
            throw new PSArgumentException("Use -Category/-Value for non-scatter charts.", nameof(Type));
        }

        if (YProperty == null || YProperty.Length == 0) {
            throw new PSArgumentException("Provide at least one -YProperty.", nameof(YProperty));
        }

        return slide.AddChart(OfficeChartType.Scatter, BuildScatterChartData(), new OfficeIMO.PowerPoint.PowerPointChartOptions { Left = global::OfficeIMO.OfficeLength.Points(X), Top = global::OfficeIMO.OfficeLength.Points(Y), Width = global::OfficeIMO.OfficeLength.Points(Width), Height = global::OfficeIMO.OfficeLength.Points(Height) });
    }

    private OfficeChartData BuildChartData() {
        var items = EnsureData();
        var categories = items.Select(item => ConvertToString(GetPropertyValue(item, Category), Category)).ToArray();
        var series = Value.Select(property =>
            new OfficeSeries(property, items.Select(item => ConvertToDouble(GetPropertyValue(item, property), property)).ToArray()))
            .ToArray();

        return new OfficeChartData(categories, series);
    }

    private OfficeChartData BuildScatterChartData() {
        var items = EnsureData();
        var xValues = items.Select(item => ConvertToDouble(GetPropertyValue(item, XProperty), XProperty)).ToArray();
        var series = YProperty.Select(property =>
            new OfficeSeries(
                property,
                items.Select(item => ConvertToDouble(GetPropertyValue(item, property), property)).ToArray(),
                xValues))
            .ToArray();

        return new OfficeChartData(
            xValues.Select(value => value.ToString(CultureInfo.InvariantCulture)),
            series);
    }

    private static OfficeChartType GetOfficeChartKind(PowerPointChartType type) => type switch {
        PowerPointChartType.Line => OfficeChartType.Line,
        PowerPointChartType.Pie => OfficeChartType.Pie,
        PowerPointChartType.Doughnut => OfficeChartType.Doughnut,
        PowerPointChartType.Scatter => OfficeChartType.Scatter,
        _ => OfficeChartType.ColumnClustered
    };

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
