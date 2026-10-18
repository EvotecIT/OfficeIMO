using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Management.Automation;
using DocumentFormat.OpenXml.Drawing.Charts;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Chart types supported by Add-OfficeWordChart.</summary>
public enum WordChartType
{
    /// <summary>Pie chart.</summary>
    Pie,
    /// <summary>Bar chart.</summary>
    Bar,
    /// <summary>Line chart.</summary>
    Line,
    /// <summary>Area chart.</summary>
    Area
}

/// <summary>Adds a chart to a Word document.</summary>
/// <para>Creates a Word chart from object data using one category property and one or more numeric series properties.</para>
/// <example>
///   <summary>Add pie and line charts to a report.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$rows = @(
///     [pscustomobject]@{ Region = 'North America'; Revenue = 125000; Profit = 42000 }
///     [pscustomobject]@{ Region = 'EMEA'; Revenue = 98000; Profit = 31000 }
///     [pscustomobject]@{ Region = 'APAC'; Revenue = 143000; Profit = 52000 }
/// )
/// New-OfficeWord -Path .\RegionalReport.docx {
///     Add-OfficeWordParagraph -Text 'Regional revenue'
///     Add-OfficeWordChart -Type Pie -InputObject $rows -CategoryProperty Region -SeriesProperty Revenue -Title 'Revenue mix' -FitToPageWidth -WidthFraction 0.75
///     Add-OfficeWordChart -Type Bar -InputObject $rows -CategoryProperty Region -SeriesProperty Revenue, Profit -Legend -XAxisTitle 'Region' -YAxisTitle 'Amount'
/// }</code>
///   <para>Creates a report with a pie chart and a multi-series bar chart from PowerShell objects.</para>
/// </example>
/// <example>
///   <summary>Add a line chart to an open document.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$trend = @(
///     [pscustomobject]@{ Month = 'Jan'; Sales = 10; Profit = 4 }
///     [pscustomobject]@{ Month = 'Feb'; Sales = 12; Profit = 5 }
///     [pscustomobject]@{ Month = 'Mar'; Sales = 15; Profit = 7 }
/// )
/// $doc = New-OfficeWord -Path .\Trend.docx -NoSave
/// Add-OfficeWordChart -Document $doc -Type Line -InputObject $trend -CategoryProperty Month -SeriesProperty Sales, Profit -Legend -Title 'Quarter trend'
/// Save-OfficeWord -Document $doc</code>
///   <para>Creates a multi-series line chart on the document and shows a legend.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordChart", DefaultParameterSetName = ParameterSetContext)]
[Alias("WordChart")]
[OutputType(typeof(WordChart))]
public sealed class AddOfficeWordChartCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetParagraph = "Paragraph";

    private static readonly string[] DefaultSeriesPalette =
    {
        "#1f77b4",
        "#ff7f0e",
        "#2ca02c",
        "#d62728",
        "#9467bd",
        "#8c564b"
    };

    /// <summary>Target document that will receive the chart.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public WordDocument? Document { get; set; }

    /// <summary>Target paragraph used as the chart anchor.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = ParameterSetParagraph)]
    public WordParagraph? Paragraph { get; set; }

    /// <summary>Chart type to create.</summary>
    [Parameter]
    public WordChartType Type { get; set; } = WordChartType.Pie;

    /// <summary>Source objects used to build chart data.</summary>
    [Parameter(Mandatory = true)]
    [Alias("Data")]
    public object[] InputObject { get; set; } = Array.Empty<object>();

    /// <summary>Property name used for category labels.</summary>
    [Parameter(Mandatory = true)]
    public string CategoryProperty { get; set; } = string.Empty;

    /// <summary>Property names used as numeric series.</summary>
    [Parameter(Mandatory = true)]
    public string[] SeriesProperty { get; set; } = Array.Empty<string>();

    /// <summary>Chart width in pixels.</summary>
    [Parameter]
    public int WidthPixels { get; set; } = 600;

    /// <summary>Chart height in pixels.</summary>
    [Parameter]
    public int HeightPixels { get; set; } = 360;

    /// <summary>Optional chart title.</summary>
    [Parameter]
    public string? Title { get; set; }

    /// <summary>Color values applied to the series in order.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string[] SeriesColor { get; set; } = Array.Empty<string>();

    /// <summary>Add a legend to the chart.</summary>
    [Parameter]
    public SwitchParameter Legend { get; set; }

    /// <summary>Legend position when <c>-Legend</c> is used.</summary>
    [Parameter]
    public OfficeChartLegendPosition LegendPosition { get; set; } = OfficeChartLegendPosition.Right;

    /// <summary>Optional X axis title for non-pie charts.</summary>
    [Parameter]
    public string? XAxisTitle { get; set; }

    /// <summary>Optional Y axis title for non-pie charts.</summary>
    [Parameter]
    public string? YAxisTitle { get; set; }

    /// <summary>Scale the chart width to the page content width.</summary>
    [Parameter]
    public SwitchParameter FitToPageWidth { get; set; }

    /// <summary>Fraction of the page content width to use when <c>-FitToPageWidth</c> is specified.</summary>
    [Parameter]
    public double WidthFraction { get; set; } = 1.0;

    /// <summary>Emit the created chart.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (InputObject == null || InputObject.Length == 0)
        {
            throw new PSArgumentException("Provide at least one data item.", nameof(InputObject));
        }

        if (SeriesProperty == null || SeriesProperty.Length == 0)
        {
            throw new PSArgumentException("Provide at least one -SeriesProperty.", nameof(SeriesProperty));
        }

        if (Type == WordChartType.Pie && SeriesProperty.Length > 1)
        {
            throw new PSArgumentException("Pie charts support only one series property.", nameof(SeriesProperty));
        }

        if (WidthPixels <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(WidthPixels), "WidthPixels must be greater than 0.");
        }

        if (HeightPixels <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(HeightPixels), "HeightPixels must be greater than 0.");
        }

        if (FitToPageWidth.IsPresent && (WidthFraction <= 0 || WidthFraction > 1))
        {
            throw new ArgumentOutOfRangeException(nameof(WidthFraction), "WidthFraction must be between 0 and 1 when FitToPageWidth is used.");
        }

        var items = InputObject.Where(item => item != null).ToArray();
        if (items.Length == 0)
        {
            throw new PSArgumentException("Chart data items cannot all be null.", nameof(InputObject));
        }

        var chart = CreateChartTarget();
        PopulateChart(chart, items);
        ApplyFormatting(chart);

        if (PassThru.IsPresent)
        {
            WriteObject(chart);
        }
    }

    private WordChart CreateChartTarget()
    {
        var title = Title ?? string.Empty;

        if (Paragraph != null)
        {
            return Paragraph.AddChart(title, width: WidthPixels, height: HeightPixels);
        }

        if (Document != null)
        {
            return Document.AddChart(title, width: WidthPixels, height: HeightPixels);
        }

        var context = WordDslContext.Require(this);
        if (context.CurrentParagraph != null)
        {
            return context.CurrentParagraph.AddChart(title, width: WidthPixels, height: HeightPixels);
        }

        var paragraph = context.RequireParagraphHost().AddParagraph();
        return paragraph.AddChart(title, width: WidthPixels, height: HeightPixels);
    }

    private void PopulateChart(WordChart chart, IReadOnlyList<object> items)
    {
        if (Type == WordChartType.Pie)
        {
            var seriesName = SeriesProperty[0];
            foreach (var item in items)
            {
                var category = ConvertToString(GetPropertyValue(item, CategoryProperty), CategoryProperty);
                var value = ConvertToDouble(GetPropertyValue(item, seriesName), seriesName);
                chart.AddPie(category, value);
            }

            return;
        }

        var categories = items.Select(item => ConvertToString(GetPropertyValue(item, CategoryProperty), CategoryProperty)).ToList();
        chart.AddChartAxisX(categories);

        for (var index = 0; index < SeriesProperty.Length; index++)
        {
            var seriesName = SeriesProperty[index];
            var values = items.Select(item => ConvertToDouble(GetPropertyValue(item, seriesName), seriesName)).ToList();
            var color = ResolveSeriesColor(index);

            switch (Type)
            {
                case WordChartType.Line:
                    chart.AddLine(seriesName, values, color);
                    break;
                case WordChartType.Area:
                    chart.AddArea(seriesName, values, color);
                    break;
                default:
                    chart.AddBar(seriesName, values, color);
                    break;
            }
        }
    }

    private void ApplyFormatting(WordChart chart)
    {
        if (FitToPageWidth.IsPresent)
        {
            chart.SetWidthToPageContent(WidthFraction, HeightPixels);
        }

        if (Type != WordChartType.Pie)
        {
            if (!string.IsNullOrWhiteSpace(XAxisTitle))
            {
                chart.SetXAxisTitle(XAxisTitle!);
            }

            if (!string.IsNullOrWhiteSpace(YAxisTitle))
            {
                chart.SetYAxisTitle(YAxisTitle!);
            }
        }

        if (Legend.IsPresent || SeriesProperty.Length > 1)
        {
            chart.AddLegend(LegendPosition);
        }
    }

    private OfficeColor ResolveSeriesColor(int index)
    {
        var colorValue = index < SeriesColor.Length && !string.IsNullOrWhiteSpace(SeriesColor[index])
            ? SeriesColor[index]
            : DefaultSeriesPalette[index % DefaultSeriesPalette.Length];

        return OfficeColor.Parse(colorValue);
    }

    private static object? GetPropertyValue(object item, string propertyName)
    {
        if (string.IsNullOrWhiteSpace(propertyName))
        {
            throw new PSArgumentException("Property name cannot be empty.", nameof(propertyName));
        }

        if (item is IDictionary dictionary)
        {
            foreach (DictionaryEntry entry in dictionary)
            {
                if (entry.Key is string key && string.Equals(key, propertyName, StringComparison.OrdinalIgnoreCase))
                {
                    return entry.Value;
                }
            }
        }

        var property = PSObject.AsPSObject(item).Properties[propertyName];
        if (property == null)
        {
            throw new PSArgumentException($"Property '{propertyName}' was not found on chart data item.");
        }

        return property.Value;
    }

    private static string ConvertToString(object? value, string propertyName)
    {
        if (value == null)
        {
            throw new PSArgumentException($"Property '{propertyName}' cannot be null.");
        }

        var text = Convert.ToString(value, CultureInfo.InvariantCulture);
        if (string.IsNullOrWhiteSpace(text))
        {
            throw new PSArgumentException($"Property '{propertyName}' cannot be empty.");
        }

        return text;
    }

    private static double ConvertToDouble(object? value, string propertyName)
    {
        if (value == null)
        {
            throw new PSArgumentException($"Property '{propertyName}' cannot be null.");
        }

        try
        {
            return Convert.ToDouble(value, CultureInfo.InvariantCulture);
        }
        catch (Exception)
        {
            throw new PSArgumentException($"Property '{propertyName}' must be numeric.", propertyName);
        }
    }
}
