using System;
using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed partial class OpenDocumentOdgLineLabelTests {
    [Theory]
    [InlineData(true, 100, 0, "top", 7.5, -11)]
    [InlineData(true, 100, 0, "middle", 7.5, -4)]
    [InlineData(true, 100, 0, "bottom", 7.5, 3)]
    [InlineData(true, 60, 80, "top", 7.5, -11)]
    [InlineData(true, 60, 80, "middle", 7.5, -4)]
    [InlineData(true, 60, 80, "bottom", 7.5, 3)]
    [InlineData(false, 100, 0, "top", 7.5, -11)]
    [InlineData(false, 100, 0, "middle", 7.5, -4)]
    [InlineData(false, 100, 0, "bottom", 7.5, 3)]
    [InlineData(false, 100, 5, "top", 7.5, -1.5)]
    [InlineData(false, 100, 5, "middle", 7.5, -3.25)]
    [InlineData(false, 100, 5, "bottom", 7.5, -5)]
    [InlineData(false, 100, 12, "top", 7.5, 2)]
    [InlineData(false, 100, 12, "middle", 7.5, -4)]
    [InlineData(false, 100, 12, "bottom", 7.5, -10)]
    [InlineData(false, 100, 40, "top", 7.5, 3)]
    [InlineData(false, 100, 40, "middle", 7.5, -4)]
    [InlineData(false, 100, 40, "bottom", 7.5, -11)]
    [InlineData(false, 10, 0, "middle", 6.25, -4)]
    [InlineData(false, 0, 85, "middle", 7.5, -4)]
    public void AsymmetricCaptionInsetsUseNativeRouteAnchorsAcrossReopening(bool line,
        double routeWidth, double routeHeight, string alignment, double along, double across) {
        var plain = CreateCaption(false); var padded = CreateCaption(true);
        string content = padded.GetXml("content.xml").ToString();
        string styles = padded.GetXml("styles.xml").ToString();
        var plainFrame = Assert.Single(Frames(plain.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported).Value));
        OfficePoint baseline = TextTop(plainFrame);
        double angle = line ? Math.Atan2(routeHeight, routeWidth) : 0;
        var delta = OfficeTransform.RotateDegrees(angle * 180 / Math.PI).TransformPoint(new OfficePoint(along, across));
        foreach (var reopened in RoundTrips(padded)) {
            var drawing = reopened.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported).Value;
            var frame = Assert.Single(Frames(drawing)); var top = TextTop(frame);
            Assert.Equal(baseline.X + delta.X, top.X, 6);
            Assert.Equal(baseline.Y + delta.Y, top.Y, 6);
            Assert.Equal(new OfficeTextPadding(20, 3, 5, 11), frame.Text.Padding);
            Assert.Equal("First item\nSecond item", frame.Text.PlainText);
            Assert.False(frame.Text.WrapText);
            Assert.All(frame.Text.Paragraphs.SelectMany(p => p.Runs), r => Assert.Equal(10, r.FontSize));
            string svg = OfficeDrawingSvgExporter.ToSvg(drawing);
            Assert.Contains("First item", svg); Assert.Contains("Second item", svg);
            Assert.NotEmpty(OfficeDrawingRasterRenderer.ToPng(drawing));
        }
        Assert.Equal(content, padded.GetXml("content.xml").ToString());
        Assert.Equal(styles, padded.GetXml("styles.xml").ToString());

        OdgDocument CreateCaption(bool inset) {
            var document = OdgDocument.Create(); var page = document.AddPage();
            var shape = line ? page.Shapes.AddLine(OdfLength.Points(100), OdfLength.Points(100),
                OdfLength.Points(100 + routeWidth), OdfLength.Points(100 + routeHeight)) :
                page.Shapes.AddConnector(new OfficePoint(100, 100), new OfficePoint(100 + routeWidth, 100 + routeHeight));
            foreach (string text in new[] { "First item", "Second item" }) {
                var paragraph = shape.AddParagraph(text); paragraph.FontFamily = "Arial";
                paragraph.FontSize = OdfLength.Points(10); paragraph.LineHeight = OdfLength.Points(11);
                paragraph.TextAlign = "center";
            }
            var graphic = Graphic(document, shape);
            graphic.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Fo + "wrap-option", "no-wrap");
            graphic.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "textarea-vertical-align", alignment);
            if (inset) foreach (var side in new[] { ("left", "20pt"), ("top", "3pt"), ("right", "5pt"), ("bottom", "11pt") })
                graphic.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Fo + "padding-" + side.Item1, side.Item2);
            return document;
        }

        static OfficePoint TextTop((OfficeDrawingRichText Text, double X, double Y, OfficeTransform Transform) frame) {
            var text = frame.Text; double innerHeight = text.Height - text.Padding.Vertical;
            // Two explicit 11-point lines make this a logical placement assertion, independent of font baselines.
            double y = OfficeTextPlacement.ResolveTop(text.Y + text.Padding.Top, innerHeight, 22, text.VerticalAlignment);
            return frame.Transform.TransformPoint(new OfficePoint(text.X + text.Padding.Left +
                (text.Width - text.Padding.Horizontal) / 2, y));
        }
    }
}
