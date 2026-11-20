using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordChartEnumerationTests {
    [Fact]
    public void AllChartsListsBodyHeaderFooterAndFootnoteChartsOfALoadedDocument() {
        using var bytes = new MemoryStream();
        using (var document = WordDocument.Create()) {
            AddChart(document.AddParagraph(), "Body");
            AddChart(document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0], "Cell");
            AddChart(document.Sections[0].PageSetup.Header.AddParagraph(), "Header");
            AddChart(document.Sections[0].PageSetup.Footer.AddParagraph(), "Footer");
            WordFootNote note = document.AddParagraph("Revenue").AddFootNote("Source");
            AddChart(note.Paragraphs!.Last(), "Note");
            document.Save(bytes);
        }

        using var loaded = WordDocument.Load(new MemoryStream(bytes.ToArray()));
        Assert.Equal(new[] { "Body", "Cell" }, loaded.Charts.Select(chart => chart.Title));
        Assert.Equal(new[] { "Body", "Cell", "Header", "Footer", "Note" }, loaded.AllCharts.Select(chart => chart.Title));
        Assert.All(loaded.AllCharts, chart =>
            Assert.Equal(OfficeImageExportFormat.Png, chart.ToImage().Format));
    }

    private static void AddChart(WordParagraph paragraph, string title) {
        WordChart chart = paragraph.AddChart(OfficeChartType.ColumnClustered, new OfficeChartData(new[] { "A", "B" },
            new[] { new OfficeSeries("Value", new[] { 1D, 2D }) }), new OfficeChartOptions { Title = title });
        Assert.Equal(title, chart.Title);
    }
}
