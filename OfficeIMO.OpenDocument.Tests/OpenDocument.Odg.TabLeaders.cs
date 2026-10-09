using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentOdgTabLeaderTests {
    [Fact]
    public void NativeExportRetainsTextualLeadersAcrossProjectionAndBothContainers() {
        var source = OdgDocument.LoadFlatXml(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Drawing", "libreoffice-tab-leaders.fodg"));
        string before = source.GetXml("content.xml").ToString();
        using var package = new MemoryStream(); source.Save(package); package.Position = 0;
        using var flat = new MemoryStream(); source.SaveFlatXml(flat); flat.Position = 0;
        foreach (var document in new[] { source, OdgDocument.Load(package), OdgDocument.LoadFlatXml(flat) }) {
            var result = document.Pages[0].ToDrawing();
            var paragraphs = result.Value.Elements.OfType<OfficeDrawingRichText>().SelectMany(t => t.Paragraphs).ToArray();
            Assert.Equal(8, paragraphs.Length);
            Assert.Equal(new[] { ".", "-", "_", ".", ".", ".", "_", "-" }, paragraphs.Select(p => p.TabStops!.Stops[0].LeaderText));
            Assert.Equal(3, paragraphs[5].TabStops!.Stops.Count);
            Assert.Contains(paragraphs[6].Runs, r => r.Text == ",45" && r.Bold && Math.Abs(r.FontSize - 20) < .01);
            Assert.DoesNotContain(result.Report.Mappings, m => m.Feature.EndsWith(":tab-leaders", StringComparison.Ordinal) && m.Status == OdfConversionMappingStatus.Unsupported);
        }
        Assert.Equal(before, source.GetXml("content.xml").ToString());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TypedLeadersSurviveAllAlignmentRoundTripsAndProjectionLeavesSourceUntouched(bool flat) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var paragraph = page.Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 12, 8), "Leaders").AddParagraph("A\tB\tC\t123,45");
        var original = new OdfTabStop(OdfLength.Points(40));
        paragraph.SetTabStops(new[] { original.WithLeader("."), new OdfTabStop(OdfLength.Points(90), "center").WithLeader("-"),
            new OdfTabStop(OdfLength.Points(140), "right").WithLeader("_"), new OdfTabStop(OdfLength.Points(200), "char", ",").WithLeader(".") });
        Assert.Null(original.LeaderText);
        using var stream = new MemoryStream(); if (flat) document.SaveFlatXml(stream); else document.Save(stream); stream.Position = 0;
        var reopened = flat ? OdgDocument.LoadFlatXml(stream) : OdgDocument.Load(stream);
        Assert.All(reopened.GetXml("content.xml").Descendants(OdfNamespaces.Style + "tab-stop"),
            stop => Assert.Equal("solid", (string?)stop.Attribute(OdfNamespaces.Style + "leader-style")));
        string before = reopened.GetXml("content.xml").ToString();
        var result = reopened.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
        var text = Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>());
        Assert.Equal(new[] { ".", "-", "_", "." }, text.Paragraphs[0].TabStops!.Stops.Select(s => s.LeaderText));
        Assert.Equal("A\tB\tC\t123,45", text.PlainText);
        Assert.Contains(result.Report.Mappings, m => m.Feature.EndsWith(":tab-leaders", StringComparison.Ordinal) && m.Status == OdfConversionMappingStatus.Approximated);
        Assert.Contains("...", OfficeDrawingSvgExporter.ToSvg(result.Value));
        Assert.Equal(before, reopened.GetXml("content.xml").ToString());
    }

    [Theory]
    [InlineData(".", false, true)]
    [InlineData(" ", false, true)]
    [InlineData("..", false, false)]
    [InlineData(".", true, false)]
    public void TextTakesPrecedenceOverLinePatternsAndUnsupportedTextRetainsSpacing(string leader, bool separateStyle, bool supported) {
        var document = OdgDocument.Create(); var paragraph = document.AddPage().Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 12, 8), "Leader").AddParagraph("A\tB");
        paragraph.SetTabStops(new[] { new OdfTabStop(OdfLength.Points(50)) });
        var stop = paragraph.EnsureStyle().Element.Descendants(OdfNamespaces.Style + "tab-stop").Single();
        stop.SetAttributeValue(OdfNamespaces.Style + "leader-text", leader);
        stop.SetAttributeValue(OdfNamespaces.Style + "leader-style", "wave");
        if (separateStyle) stop.SetAttributeValue(OdfNamespaces.Style + "leader-text-style", "StyledLeader");
        string before = document.GetXml("content.xml").ToString(); var result = document.Pages[0].ToDrawing();
        var projected = Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>()).Paragraphs[0].TabStops!.Stops.Single();
        Assert.Equal(50, projected.Position); Assert.Equal(supported ? leader : null, projected.LeaderText);
        Assert.Equal(!supported, result.Report.HasSkippedOrUnsupported);
        Assert.Equal(before, document.GetXml("content.xml").ToString());
    }
}
