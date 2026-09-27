using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordChartAlternateContentTests {
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ChartDiscoveryAndVisibleSegmentsSelectOnlyTheEffectiveBranch(bool supportedChoice) {
        using var document = WordDocument.Create();
        document.AddChart(OfficeChartKind.ColumnClustered, new OfficeChartData(new[] { "A" }, new[] { new OfficeChartSeries("Values", new[] { 3d }) }));
        var paragraph = document.Paragraphs.Single().GetRuns().Single();
        var run = paragraph._run!;
        var drawing = run.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Drawing>()!;
        drawing.Remove();
        var choice = new AlternateContentChoice { Requires = supportedChoice ? "wps" : "unsupported" };
        var fallback = new AlternateContentFallback();
        if (supportedChoice) { choice.Append(drawing); fallback.Append(drawing.CloneNode(true)); }
        else { choice.Append(new DocumentFormat.OpenXml.Wordprocessing.Text("Unselected")); fallback.Append(drawing); }
        run.Append(new AlternateContent(choice, fallback));
        Assert.True(paragraph.IsChart);
        Assert.True(paragraph.Chart!.TryGetOfficeSnapshot(out var snapshot));
        Assert.Equal(3d, snapshot.Data.Series.Single().Values.Single());
        var segments = WordEquation.GetVisibleContentSegments(run, Array.Empty<WordEquationOccurrence>());
        Assert.Single(segments, segment => segment.IsRunArtifact);
        Assert.DoesNotContain(segments, segment => segment.Text == "Unselected");
    }
}
