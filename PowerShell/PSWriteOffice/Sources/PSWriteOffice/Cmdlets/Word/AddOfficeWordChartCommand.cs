using PSWriteOffice.Foundation;
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
///     Add-OfficeWordChart -Type Pie -InputObject $rows -Category Region -Value Revenue -Title 'Revenue mix' -FitToPageWidth -WidthFraction 0.75
///     Add-OfficeWordChart -Type Bar -InputObject $rows -Category Region -Value Revenue, Profit -Legend -XAxisTitle 'Region' -YAxisTitle 'Amount'
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
/// Add-OfficeWordChart -Document $doc -Type Line -InputObject $trend -Category Month -Value Sales, Profit -Legend -Title 'Quarter trend'
/// Save-OfficeWord -Document $doc</code>
///   <para>Creates a multi-series line chart on the document and shows a legend.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeWordChart", DefaultParameterSetName = ParameterSetContext)]
[OutputType(typeof(WordChart))]
[Alias("WordChart") ]
public sealed class AddOfficeWordChartCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetParagraph = "Paragraph";

    private readonly List<object> _input = new();

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
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetDocument)]
    public WordDocument? Document { get; set; }

    /// <summary>Target paragraph used as the chart anchor.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetParagraph)]
    public WordParagraph? Paragraph { get; set; }

    /// <summary>Chart type to create.</summary>
    [Parameter]
    public OfficeChartType Type { get; set; } = OfficeChartType.Pie;

    /// <summary>Source objects used to build chart data.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    [Alias("Data")]
    public object[] InputObject { get; set; } = Array.Empty<object>();

    /// <summary>Property name used for category labels.</summary>
    [Parameter(Mandatory = true)]
    public string Category { get; set; } = string.Empty;

    /// <summary>Property names used as numeric series.</summary>
    [Parameter(Mandatory = true)]
    public string[] Value { get; set; } = Array.Empty<string>();

    /// <summary>Chart width; numbers mean points.</summary>
    [Parameter]
    [OfficeLengthTransformation]
    public OfficeIMO.OfficeLength Width { get; set; } = OfficeIMO.OfficeLength.Pixels(600);

    /// <summary>Chart height; numbers mean points.</summary>
    [Parameter]
    [OfficeLengthTransformation]
    public OfficeIMO.OfficeLength Height { get; set; } = OfficeIMO.OfficeLength.Pixels(360);

    /// <summary>Optional chart title.</summary>
    [Parameter]
    public string? Title { get; set; }

    /// <summary>Color values applied to the series in order.</summary>
    [Parameter]
    [OfficeColorTransformation]
    [ArgumentCompleter(typeof(OfficeColorCompleter))]
    public OfficeColor[] SeriesColor { get; set; } = Array.Empty<OfficeColor>();

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
        if (InputObject != null) _input.AddRange(InputObject);
    }

    /// <inheritdoc />
    protected override void EndProcessing()
    {
        InputObject = _input.ToArray();
        if (InputObject == null || InputObject.Length == 0)
        {
            throw new PSArgumentException("Provide at least one data item.", nameof(InputObject));
        }

        if (Value == null || Value.Length == 0)
        {
            throw new PSArgumentException("Provide at least one -Value.", nameof(Value));
        }

        if (Type == OfficeChartType.Pie && Value.Length > 1)
        {
            throw new PSArgumentException("Pie charts support only one series property.", nameof(Value));
        }

        if (Width <= OfficeIMO.OfficeLength.Zero)
        {
            throw new ArgumentOutOfRangeException(nameof(Width), "Width must be greater than 0.");
        }

        if (Height <= OfficeIMO.OfficeLength.Zero)
        {
            throw new ArgumentOutOfRangeException(nameof(Height), "Height must be greater than 0.");
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
            return Paragraph.AddChart(title,width: Width,height: Height);
        }

        if (Document != null)
        {
            return Document.AddChart(title,width: Width,height: Height);
        }

        var context = WordDslContext.Require(this);
        if (context.CurrentParagraph != null)
        {
            return context.CurrentParagraph.AddChart(title,width: Width,height: Height);
        }

        var paragraph = context.RequireParagraphHost().AddParagraph();
        return paragraph.AddChart(title,width: Width,height: Height);
    }

    private void PopulateChart(WordChart chart, IReadOnlyList<object> items)
    {
        OfficeChartData source = PSWriteOffice.Services.OfficeChartDataBuilder.Build(items.ToArray(), Category, Value, Type);
        WordDocument document = Paragraph?._document ?? Document ?? WordDslContext.Require(this).Document;
        var series = source.Series.Select((item, index) => new OfficeSeries(item.Name, item.Values, item.XValues,
            OfficeColor.ParseHex(WordRunColor.Create(ResolveSeriesColor(index), document).Val!.Value!))).ToArray();
        chart.SetData(Type, new OfficeChartData(source.Categories, series));
    }

    private void ApplyFormatting(WordChart chart)
    {
        if (FitToPageWidth.IsPresent)
        {
            chart.SetWidthToPageContent(WidthFraction,Height);
        }

        if (Type != OfficeChartType.Pie)
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

        if (Legend.IsPresent || Value.Length > 1)
        {
            chart.AddLegend(LegendPosition);
        }
    }

    private OfficeColor ResolveSeriesColor(int index) => index < SeriesColor.Length
        ? SeriesColor[index] : OfficeColor.Parse(DefaultSeriesPalette[index % DefaultSeriesPalette.Length]);
}
