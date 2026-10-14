using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentOdgAlignedTabTests {
    private static OdgDocument Document(string alignment, string type = "left", double indent = 0, bool hardBreaks = false) {
        var doc = OdgDocument.Create();
        var paragraph = doc.AddPage().Shapes.AddTextBox(OdfRect.FromCentimeters(1, 1, 15, 5), "").Paragraphs[0];
        paragraph.Text = hardBreaks ? "Before\nA\t123,45\nAfter" : "A\t123,45";
        paragraph.TextAlign = alignment; paragraph.FontSize = OdfLength.Points(12);
        paragraph.EnsureStyle().TextIndent = OdfLength.Points(indent);
        paragraph.SetTabStops(new[] { new OdfTabStop(OdfLength.Points(220), type, ",").WithLeader(".") });
        return doc;
    }

    [Theory]
    [InlineData("center", "left")]
    [InlineData("center", "center")]
    [InlineData("center", "right")]
    [InlineData("center", "char")]
    [InlineData("right", "left")]
    [InlineData("right", "center")]
    [InlineData("right", "right")]
    [InlineData("right", "char")]
    [InlineData("justify", "left")]
    [InlineData("justify", "center")]
    [InlineData("justify", "right")]
    [InlineData("justify", "char")]
    [InlineData("end", "left")]
    public void SupportedAlignedTabParagraphsPassStrictConversionAndPreserveBothContainers(string alignment, string type) {
        var document = Document(alignment, type);
        using var flat = new MemoryStream(); document.SaveFlatXml(flat); flat.Position = 0;
        foreach (var read in new[] { document, OdgDocument.Load(new MemoryStream(document.ToBytes())), OdgDocument.LoadFlatXml(flat) }) {
            string before = read.GetXml("content.xml").ToString();
            var result = read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
            var paragraph = Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>()).Paragraphs[0];
            Assert.Equal(alignment == "center" ? OfficeTextAlignment.Center : alignment == "justify" ? OfficeTextAlignment.Justify : OfficeTextAlignment.Right, paragraph.Alignment);
            Assert.Equal("A\t123,45", string.Concat(paragraph.Runs.Select(r => r.Text)));
            Assert.Equal(220, Assert.Single(paragraph.TabStops!.Stops).Position);
            Assert.True(paragraph.TabStops.AlignWithParagraph);
            Assert.Contains(result.Report.Mappings, m => m.Feature.EndsWith(":tab-paragraph-alignment", StringComparison.Ordinal) && m.Status == OdfConversionMappingStatus.Approximated);
            Assert.False(result.Report.HasSkippedOrUnsupported); Assert.Equal(before, read.GetXml("content.xml").ToString());
        }
    }

    [Theory]
    [InlineData("center")]
    [InlineData("right")]
    public void ParagraphAlignmentMovesCompleteTabbedLinesWhileRetainingFieldSpacingAndPlainLineAlignment(string alignment) {
        static double X(XDocument svg, string text) => double.Parse(Assert.Single(svg.Descendants(), e => e.Name.LocalName == "text" && e.Value == text).Attribute("x")!.Value, System.Globalization.CultureInfo.InvariantCulture);
        var left = XDocument.Parse(OfficeDrawingSvgExporter.ToSvg(Document("left", hardBreaks: true).Pages[0].ToDrawing().Value));
        var aligned = XDocument.Parse(OfficeDrawingSvgExporter.ToSvg(Document(alignment, hardBreaks: true).Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported).Value));
        Assert.True(X(aligned, "123,45") > X(left, "123,45"));
        Assert.InRange(Math.Abs((X(left, "123,45") - X(left, "A")) - (X(aligned, "123,45") - X(aligned, "A"))), 0, .002);
        Assert.True(X(aligned, "Before") > X(left, "Before")); Assert.True(X(aligned, "After") > X(left, "After"));
    }

    [Theory]
    [InlineData("center")]
    [InlineData("right")]
    [InlineData("end")]
    public void CenteredOrRightHangingIndentRemainsExplicitlyUnqualified(string alignment) {
        var document = Document(alignment, indent: -25);
        document.Pages[0].Shapes[0].Paragraphs[0].MarginLeft = OdfLength.Points(25);
        string before = document.GetXml("content.xml").ToString();
        var result = document.Pages[0].ToDrawing();
        Assert.Contains(result.Report.Mappings, m => m.Feature.EndsWith(":tab-paragraph-alignment", StringComparison.Ordinal) && m.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Equal("A\t123,45", Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>()).PlainText);
        Assert.Throws<OdfConversionLossException>(() => document.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported)); Assert.Equal(before, document.GetXml("content.xml").ToString());
    }

    [Theory]
    [InlineData("center")]
    [InlineData("right")]
    [InlineData("end")]
    [InlineData("justify")]
    public void PositiveFirstLineIndentUsesTheQualifiedTabAlignment(string alignment) {
        var result = Document(alignment, indent: 25).Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
        var p = Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>()).Paragraphs[0];
        Assert.Equal(25, p.Indent.FirstLineOffset); Assert.False(result.Report.HasSkippedOrUnsupported);
    }
    [Theory]
    [InlineData("center")]
    [InlineData("right")]
    [InlineData("justify")]
    public void NonLeftListBodyTabsRetainTheUnqualifiedSelectionReport(string alignment) {
        var document = OdgDocument.Create(); var box = document.AddPage().Shapes.AddTextBox(OdfRect.FromCentimeters(1, 1, 15, 5), "");
        var p = box.AddList(ordered: true).AddItem("A\t123,45").Paragraphs[0]; p.TextAlign = alignment;
        p.SetTabStops(new[] { new OdfTabStop(OdfLength.Points(220)) });
        string before = document.GetXml("content.xml").ToString(); var result = document.Pages[0].ToDrawing();
        Assert.Contains(result.Report.Mappings, m => m.Feature.EndsWith(":tab-paragraph-alignment", StringComparison.Ordinal) && m.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => document.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
        Assert.Contains("123,45", Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>()).PlainText); Assert.Equal(before, document.GetXml("content.xml").ToString());
    }

    [Fact]
    public void NativeResaveRetainsTheQualifiedAndUnqualifiedAlignmentCasesWithoutMutatingXml() {
        var document = OdgDocument.LoadFlatXml(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Drawing", "libreoffice-aligned-tabs.fodg"));
        string before = document.GetXml("content.xml").ToString(); var result = document.Pages[0].ToDrawing();
        Assert.Equal(26, result.Value.Elements.OfType<OfficeDrawingRichText>().Count(t => t.PlainText.Contains('\t')));
        Assert.Equal(21, result.Report.Mappings.Count(m => m.Feature.EndsWith(":tab-paragraph-alignment", StringComparison.Ordinal) && m.Status == OdfConversionMappingStatus.Approximated));
        Assert.DoesNotContain(result.Report.Mappings, m => m.Feature.EndsWith(":tab-paragraph-alignment", StringComparison.Ordinal) && m.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Equal(before, document.GetXml("content.xml").ToString());
    }

}
