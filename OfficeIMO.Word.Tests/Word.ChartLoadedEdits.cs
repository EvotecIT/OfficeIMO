using OfficeIMO.Word;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class WordChartLoadedEditTests {
    [Fact]
    public void AxisTitleEditsOnLoadedChartKeepTheOtherSavedTitle() {
        using var reloaded = SaveAndLoad(document => {
            var chart = CreateBarChart(document);
            chart.SetXAxisTitle("Quarter");
            chart.SetYAxisTitle("Revenue");
            var plot = chart.ChartPart!.ChartSpace!.Descendants<C.PlotArea>().Single();
            plot.GetFirstChild<C.CategoryAxis>()!.AddChild(new C.MajorGridlines(), true);
            plot.GetFirstChild<C.ValueAxis>()!.AddChild(new C.MinorGridlines(), true);
            var savedTitle = plot.GetFirstChild<C.ValueAxis>()!.GetFirstChild<C.Title>()!;
            var rich = savedTitle.ChartText!.RichText!;
            var run = rich.Descendants<A.Run>().Single();
            (run.RunProperties ?? run.PrependChild(new A.RunProperties())).Bold = true;
            rich.Append(new A.Paragraph(new A.Run(new A.Text("(EUR)"))));
        });
        WordChart chart = Assert.Single(reloaded.Charts);

        chart.SetXAxisTitle("Period");
        Assert.Equal(("Period", "Revenue(EUR)"), ReadAxisTitles(chart));
        var categoryAxis = chart.ChartPart!.ChartSpace!.Descendants<C.CategoryAxis>().Single();
        Assert.True(categoryAxis.ChildElements.ToList().IndexOf(categoryAxis.GetFirstChild<C.Title>()!) >
            categoryAxis.ChildElements.ToList().IndexOf(categoryAxis.GetFirstChild<C.MajorGridlines>()!));

        chart.SetAxisTitleFormat("Arial", OfficeLength.Points(14), OfficeColor.Parse("#C0392B"));
        Assert.Equal(("Period", "Revenue(EUR)"), ReadAxisTitles(chart));
        var valueTitle = chart.ChartPart.ChartSpace.Descendants<C.ValueAxis>().Single().GetFirstChild<C.Title>()!;
        Assert.Equal(2, valueTitle.ChartText!.RichText!.Elements<A.Paragraph>().Count());
        var runs = valueTitle.Descendants<A.Run>().ToList();
        Assert.True(runs[0].RunProperties!.Bold!.Value);
        Assert.All(runs, item => {
            Assert.Equal("Arial", item.RunProperties!.GetFirstChild<A.LatinFont>()!.Typeface!.Value);
            Assert.Equal(1400, item.RunProperties.FontSize!.Value);
            Assert.Equal("C0392B", item.RunProperties.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value);
        });
        Assert.Empty(reloaded.ValidateDocument());
    }

    [Fact]
    public void LoadedChartReadsSavedTitleAndCategoriesForLaterEdits() {
        using var reloaded = SaveAndLoad(document => CreateBarChart(document));
        WordChart chart = Assert.Single(reloaded.Charts);
        var title = chart.ChartPart!.ChartSpace!.Descendants<C.Title>().First();
        title.Descendants<A.Run>().Single().InsertAfterSelf(new A.Run(new A.Text(" by region")));
        Assert.Equal("Sales by region", chart.Title);

        chart.AddBar("Target", new List<int> { 3, 4 }, OfficeColor.Parse("#16A085"));
        var added = chart.ChartPart.ChartSpace.Descendants<C.BarChartSeries>().Last();
        Assert.Equal(new[] { "a", "b" }, added.GetFirstChild<C.CategoryAxisData>()!.Descendants<C.StringPoint>()
            .Select(point => point.NumericValue!.Text));
        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));
        Assert.Equal(new[] { 3D, 4D }, snapshot.Data.Series[1].Values);

        chart.ChartPart.ChartSpace.GetFirstChild<C.RoundedCorners>()?.Remove();
        chart.RoundedCorners = false;
        Assert.False(chart.RoundedCorners);
        Assert.False(chart.ChartPart.ChartSpace.GetFirstChild<C.RoundedCorners>()!.Val!.Value);
    }

    private static WordChart CreateBarChart(WordDocument document) {
        var chart = document.AddChart("Sales");
        chart.AddCategories(new List<string> { "a", "b" });
        chart.AddBar("Actual", new List<int> { 1, 2 }, OfficeColor.Parse("#4472C4"));
        return chart;
    }

    private static WordDocument SaveAndLoad(Action<WordDocument> build) {
        using var bytes = new MemoryStream();
        using (var document = WordDocument.Create()) {
            build(document);
            document.Save(bytes);
        }
        return WordDocument.Load(new MemoryStream(bytes.ToArray()));
    }

    private static (string? X, string? Y) ReadAxisTitles(WordChart chart) {
        var plot = chart.ChartPart!.ChartSpace!.Descendants<C.PlotArea>().Single();
        static string? Text(C.Title? title) => title == null ? null
            : string.Concat(title.Descendants<A.Text>().Select(text => text.Text));
        return (Text(plot.GetFirstChild<C.CategoryAxis>()?.GetFirstChild<C.Title>()),
            Text(plot.GetFirstChild<C.ValueAxis>()?.GetFirstChild<C.Title>()));
    }
}
