using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordChartStoryPartTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NativeChart_IsOwnedByItsHeaderOrFooterAndSurvivesReopen(bool footer) {
        using var document = WordDocument.Create();
        var bodyChart = document.AddChart(OfficeChartKind.ColumnClustered, Data(1));
        WordHeaderFooter region = footer ? document.Sections[0].GetOrCreateFooter(WordHeaderFooterType.Default) : document.Sections[0].GetOrCreateHeader(WordHeaderFooterType.Default);
        var paragraph = region.AddParagraph();
        var chart = paragraph.AddChart(OfficeChartKind.ColumnClustered, Data(9));
        var owner = footer ? (DocumentFormat.OpenXml.Packaging.OpenXmlPart)region._footer!.FooterPart! : region._header!.HeaderPart!;
        Assert.Single(owner.Parts, part => part.OpenXmlPart is DocumentFormat.OpenXml.Packaging.ChartPart);
        const string collisionId = "rIdChartCollision";
        document.MainDocumentPartRoot.ChangeIdOfPart(bodyChart.ChartPart!, collisionId);
        bodyChart.Drawing!.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartReference>().Single().Id = collisionId;
        var regionPart = owner.Parts.Single(part => part.OpenXmlPart is DocumentFormat.OpenXml.Packaging.ChartPart).OpenXmlPart;
        owner.ChangeIdOfPart(regionPart, collisionId);
        chart.Drawing!.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartReference>().Single().Id = collisionId;
        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));
        Assert.Equal(9d, snapshot.Data.Series.Single().Values.Single());
        chart.SetData(OfficeChartKind.ColumnClustered, Data(11));
        Assert.True(bodyChart.TryGetOfficeSnapshot(out var bodySnapshot));
        Assert.Equal(1d, bodySnapshot.Data.Series.Single().Values.Single());
        Assert.Empty(document.ValidateDocument());
        using var stream = new MemoryStream(); document.Save(stream); stream.Position = 0;
        using var reopened = WordDocument.Load(stream);
        WordHeaderFooter reopenedRegion = footer ? reopened.Sections[0].Footer.Default! : reopened.Sections[0].Header.Default!;
        var reopenedChart = reopenedRegion.Paragraphs.SelectMany(p => p.GetRuns()).Single(p => p.IsChart).Chart!;
        Assert.True(reopenedChart.TryGetOfficeSnapshot(out snapshot));
        Assert.Equal(11d, snapshot.Data.Series.Single().Values.Single());
        Assert.Empty(reopened.ValidateDocument());
    }

    private static OfficeChartData Data(double value) => new(new[] { "A" }, new[] { new OfficeChartSeries("Values", new[] { value }) });
}
