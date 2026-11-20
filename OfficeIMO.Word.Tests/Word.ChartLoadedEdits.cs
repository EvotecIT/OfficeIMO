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
        });
        WordChart chart = Assert.Single(reloaded.Charts);

        chart.SetXAxisTitle("Period");
        Assert.Equal(("Period", "Revenue"), ReadAxisTitles(chart));

        chart.SetAxisTitleFormat("Arial", OfficeLength.Points(14), OfficeColor.Parse("#C0392B"));
        Assert.Equal(("Period", "Revenue"), ReadAxisTitles(chart));
        var valueTitle = chart.ChartPart!.ChartSpace!.Descendants<C.ValueAxis>().Single().GetFirstChild<C.Title>()!;
        Assert.Equal("Arial", valueTitle.Descendants<A.LatinFont>().Single().Typeface!.Value);
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
