using System;
using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed partial class OpenDocumentOdgLineLabelTests {
    [Theory]
    [InlineData("left", false, 150)]
    [InlineData("center", false, 150)]
    [InlineData("right", false, 150)]
    [InlineData("left", true, 150)]
    [InlineData("center", true, 150)]
    [InlineData("right", true, 150)]
    [InlineData("left", false, 20)]
    [InlineData("center", false, 20)]
    [InlineData("right", false, 20)]
    [InlineData("left", true, 20)]
    [InlineData("center", true, 20)]
    [InlineData("right", true, 20)]
    public void CaptionAreaAlignmentPositionsTheIntrinsicBlockIndependentlyOfItsParagraphs(
        string area, bool line, double length) {
        var metrics = new OfficeRasterCanvas(new OfficeRasterImage(1, 1));
        double measuredWidth = metrics.MeasureText("Second item", 10, "Arial");
        foreach (string paragraphAlignment in new[] { "left", "center", "right" }) {
            var document = AreaCaption(area, paragraphAlignment, line, length);
            string source = document.GetXml("content.xml").ToString();
            foreach (var read in RoundTrips(document)) {
                var drawing = read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported).Value;
                var frame = Assert.Single(Frames(drawing));
                Assert.Equal(measuredWidth, frame.Text.Width, 6);
                Assert.Equal("First item\nSecond item", frame.Text.PlainText);
                var expectedAlignment = paragraphAlignment switch {
                    "center" => OfficeTextAlignment.Center, "right" => OfficeTextAlignment.Right, _ => OfficeTextAlignment.Left
                };
                Assert.All(frame.Text.Paragraphs, p => Assert.Equal(expectedAlignment, p.Alignment));
                double x = area switch { "left" => 0, "right" => length - measuredWidth, _ => (length - measuredWidth) / 2 };
                double angle = line ? Math.Atan2(.8, .6) : 0;
                var expected = OfficeTransform.RotateDegrees(angle * 180 / Math.PI)
                    .Then(OfficeTransform.Translate(100, 100)).TransformPoint(new OfficePoint(x, 0));
                var actual = frame.Transform.TransformPoint(new OfficePoint(frame.Text.X, frame.Text.Y + frame.Text.Height / 2));
                Assert.Equal(expected.X, actual.X, 6); Assert.Equal(expected.Y, actual.Y, 6);
                Assert.Contains("Second item", OfficeDrawingSvgExporter.ToSvg(drawing));
            }
            Assert.Equal(source, document.GetXml("content.xml").ToString());
        }
    }

    [Theory]
    [InlineData("left")]
    [InlineData("center")]
    [InlineData("right")]
    [InlineData("justify")]
    public void ZeroWidthAreaRetainsVisibleCaptionsAndReportsNativeOffPagePlacement(string area) {
        foreach (string paragraphAlignment in new[] { "left", "center", "right" }) {
            var document = AreaCaption(area, paragraphAlignment, false, 0);
            foreach (var read in RoundTrips(document)) {
                var result = read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
                var frame = Assert.Single(Frames(result.Value));
                Assert.InRange(frame.X, 0, 200); Assert.InRange(frame.Y, 0, 200);
                bool differs = area is "left" or "right" || area == "justify" && paragraphAlignment != "center";
                Assert.Equal(differs, result.Report.Mappings.Any(m => m.Feature.EndsWith(":label-degenerate-area", StringComparison.Ordinal) &&
                    m.Status == OdfConversionMappingStatus.Approximated));
                Assert.Equal("First item\nSecond item", frame.Text.PlainText);
                Assert.Contains("Second item", OfficeDrawingSvgExporter.ToSvg(result.Value));
            }
        }
    }

    private static OdgDocument AreaCaption(string area, string paragraphAlignment, bool line, double length) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var shape = line ? page.Shapes.AddLine(OdfLength.Points(100), OdfLength.Points(100),
            OdfLength.Points(100 + length * .6), OdfLength.Points(100 + length * .8)) :
            page.Shapes.AddConnector(new OfficePoint(100, 100), new OfficePoint(100 + length, length == 0 ? 200 : 100));
        foreach (string text in new[] { "First item", "Second item" }) {
            var paragraph = shape.AddParagraph(text); paragraph.FontFamily = "Arial";
            paragraph.FontSize = OdfLength.Points(10); paragraph.LineHeight = OdfLength.Points(11);
            paragraph.TextAlign = paragraphAlignment;
        }
        var graphic = Graphic(document, shape);
        graphic.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Fo + "wrap-option", "no-wrap");
        graphic.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "textarea-horizontal-align", area);
        graphic.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "textarea-vertical-align", "middle");
        return document;
    }
}
