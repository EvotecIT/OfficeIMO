using System;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed partial class OpenDocumentOdgLineLabelTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void PercentageListFontsUseTheFirstDisplayedCharacterWhenBodyFormattingMovesIntoASpan(bool lineCaption, bool ordered) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var shape = lineCaption ? page.Shapes.AddConnector(new OfficePoint(100, 100), new OfficePoint(180, 100)) :
            page.Shapes.AddTextBox(new OdfRect(OdfLength.Points(100), OdfLength.Points(100), OdfLength.Points(180), OdfLength.Points(100)), string.Empty);
        shape.TextRoot.RemoveNodes();
        shape.FontFamily = "Arial";
        Graphic(document, shape).SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Fo + "wrap-option", "no-wrap");
        var list = shape.AddList(ordered); var paragraph = list.AddItem().Paragraphs[0]; paragraph.FontSize = OdfLength.Points(20);
        var first = paragraph.AddRun("First"); first.FontSize = OdfLength.Points(10); paragraph.AddRun(" later").FontSize = OdfLength.Points(30);
        var level = document.GetXml("content.xml").Descendants(OdfNamespaces.Text + "list-style").Single().Elements().Single();
        level.Add(new XElement(OdfNamespaces.Style + "text-properties", new XAttribute(OdfNamespaces.Fo + "font-size", "50%")));
        foreach (var read in RoundTrips(document)) {
            var projected = Assert.Single(Frames(read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported).Value)).Text.Paragraphs[0];
            Assert.Equal(5, projected.Label!.Run.FontSize);
            Assert.Equal(10, projected.Runs[0].FontSize);
        }
        level.Element(OdfNamespaces.Style + "text-properties")!.SetAttributeValue(OdfNamespaces.Fo + "font-size", "7pt");
        Assert.Equal(7, Assert.Single(Frames(page.ToDrawing().Value)).Text.Paragraphs[0].Label!.Run.FontSize);
        if (!ordered) {
            level.SetAttributeValue(OdfNamespaces.Text + "bullet-relative-size", "70%");
            Assert.Equal(14, Assert.Single(Frames(page.ToDrawing().Value)).Text.Paragraphs[0].Label!.Run.FontSize);
        }
    }

    [Fact]
    public void EmptyListItemCaptionRetainsItsVisibleMarkerAndReservesItsLargerFontHeight() {
        var document = OdgDocument.Create(); var shape = document.AddPage().Shapes.AddConnector(new OfficePoint(120, 100), new OfficePoint(140, 100));
        shape.FontFamily = "Arial";
        Graphic(document, shape).SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Fo + "wrap-option", "no-wrap");
        var list = shape.AddList(); list.AddItem().Paragraphs[0].FontSize = OdfLength.Points(10);
        var level = document.GetXml("content.xml").Descendants(OdfNamespaces.Text + "list-style").Single().Elements().Single();
        level.Add(new XElement(OdfNamespaces.Style + "text-properties", new XAttribute(OdfNamespaces.Fo + "font-size", "400%")));
        foreach (var read in RoundTrips(document)) {
            var result = read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
            var frame = Assert.Single(Frames(result.Value));
            Assert.Equal("• ", frame.Text.PlainText);
            Assert.Equal(40, frame.Text.Paragraphs[0].Label!.Run.FontSize);
            Assert.True(frame.Text.Height >= 40);
            Assert.Contains("•", OfficeDrawingSvgExporter.ToSvg(result.Value));
        }
    }

    [Theory]
    [InlineData("bullet", false)]
    [InlineData("bullet", true)]
    [InlineData("numbered", false)]
    [InlineData("numbered", true)]
    [InlineData("left", false)]
    [InlineData("left", true)]
    [InlineData("center", false)]
    [InlineData("center", true)]
    [InlineData("right", false)]
    [InlineData("right", true)]
    [InlineData("char", false)]
    [InlineData("char", true)]
    public void LineCaptionsRetainListMarkersAndMeasuredTabFieldsAcrossBothContainers(string profile, bool line) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        OdgShape shape = line ? page.Shapes.AddLine(OdfLength.Points(100), OdfLength.Points(100), OdfLength.Points(120), OdfLength.Points(120)) :
            page.Shapes.AddConnector(new OfficePoint(100, 100), new OfficePoint(120, 120));
        shape.Name = "Structured caption"; shape.FontFamily = "Arial";
        Graphic(document, shape).SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Fo + "wrap-option", "no-wrap");
        bool listed = profile is "bullet" or "numbered";
        if (listed) {
            var list = shape.AddList(profile == "numbered"); list.AddItem("First item"); list.AddItem("Second item");
            var level = document.GetXml("content.xml").Descendants(OdfNamespaces.Text + "list-style")
                .Single(e => (string?)e.Attribute(OdfNamespaces.Style + "name") == list.StyleName).Elements().Single();
            level.Add(new XElement(OdfNamespaces.Style + "text-properties", new XAttribute(OdfNamespaces.Fo + "font-size", "100%")));
        } else {
            foreach (string value in new[] { profile == "char" ? "123,45" : "Body", "Next" }) {
                var paragraph = shape.AddParagraph("A\t" + value);
                paragraph.SetTabStops(new[] { new OdfTabStop(OdfLength.Points(50), profile, profile == "char" ? "," : null) });
            }
        }
        foreach (var paragraph in shape.Paragraphs) { paragraph.FontSize = OdfLength.Points(10); paragraph.TextAlign = "left"; }
        string before = document.GetXml("content.xml").ToString();
        foreach (var read in RoundTrips(document)) {
            var result = read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
            var frame = Assert.Single(Frames(result.Value));
            Assert.False(frame.Text.WrapText);
            Assert.Equal(110, Center(frame).X, 6); Assert.Equal(110, Center(frame).Y, 6);
            Assert.Equal(line ? Math.PI / 4 : 0, Math.Atan2(frame.Transform.M12, frame.Transform.M11), 6);
            if (listed) {
                Assert.Equal(profile == "numbered" ? new[] { "1.", "2." } : new[] { "•", "•" }, frame.Text.Paragraphs.Select(p => p.Label!.Run.Text));
                Assert.All(frame.Text.Paragraphs, p => Assert.Equal(10, p.Label!.Run.FontSize));
                Assert.Equal(profile == "numbered" ? "1. First item\n2. Second item" : "• First item\n• Second item", frame.Text.PlainText);
            } else {
                Assert.Equal("A\t" + (profile == "char" ? "123,45" : "Body") + "\nA\tNext", frame.Text.PlainText);
                Assert.All(frame.Text.Paragraphs, p => Assert.Equal(50, Assert.Single(p.TabStops!.Stops).Position));
            }
            Assert.Contains(listed ? "Second item" : "Next", OfficeDrawingSvgExporter.ToSvg(result.Value));
            Assert.NotEmpty(OfficeDrawingRasterRenderer.ToPng(result.Value));
        }
        Assert.Equal(before, document.GetXml("content.xml").ToString());
    }
}
