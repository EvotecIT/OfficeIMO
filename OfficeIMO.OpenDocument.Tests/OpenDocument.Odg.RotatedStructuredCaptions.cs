using System;
using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed partial class OpenDocumentOdgLineLabelTests {
    [Theory]
    [InlineData("justify", false)]
    [InlineData("left", false)]
    [InlineData("center", false)]
    [InlineData("right", false)]
    [InlineData("justify", true)]
    [InlineData("left", true)]
    [InlineData("center", true)]
    [InlineData("right", true)]
    public void RotatedNumberedBodyTabsRetainTheirAreaAnchorAndFields(string area, bool line) {
        foreach (string declaration in new[] {
            "rotate(0.5235987755982988) translate(220pt 250pt)",
            "rotate(-0.5235987755982988) translate(220pt 250pt)",
            "matrix(0.8660254037844386 -0.5 0.5 0.8660254037844386 220pt 250pt)"
        }) {
            double degrees = declaration.StartsWith("rotate(-", StringComparison.Ordinal) ? 30 : -30;
            var transform = OfficeTransform.RotateDegrees(degrees).Then(OfficeTransform.Translate(220, 250));
            var start = transform.TransformPoint(new OfficePoint(80, 40));
            var end = transform.TransformPoint(new OfficePoint(230, 40));
            var document = OdgDocument.Create(); var page = document.AddPage();
            var shape = line ? page.Shapes.AddLine(OdfLength.Points(80), OdfLength.Points(40), OdfLength.Points(230), OdfLength.Points(40)) :
                page.Shapes.AddConnector(new OfficePoint(80, 40), new OfficePoint(230, 40));
            ConfigureNumberedTabBody(document, shape);
            shape.Transform = declaration;
            foreach (var paragraph in shape.Paragraphs) paragraph.Color = OdfColor.Parse("#000000");
            document.GetXml("content.xml").Descendants(OdfNamespaces.Text + "list-style").Single()
                .Descendants(OdfNamespaces.Style + "text-properties").Single().SetAttributeValue(OdfNamespaces.Fo + "color", "#000000");
            var graphic = Graphic(document, shape);
            graphic.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "textarea-horizontal-align", area);
            graphic.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "textarea-vertical-align", "middle");
            foreach (var side in new[] { ("left", 20D), ("top", 3D), ("right", 5D), ("bottom", 11D) })
                graphic.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Fo + "padding-" + side.Item1, OdfLength.Points(side.Item2).ToString());
            var metrics = new OfficeRasterCanvas(new OfficeRasterImage(1, 1));
            double stop = OdfLength.Centimeters(1.5).ToPoints(), bodyStart = OdfLength.Centimeters(.6).ToPoints();
            double intrinsic = Math.Max(stop - metrics.MeasureText("123", 10, "Arial") + metrics.MeasureText("123,45", 10, "Arial"),
                bodyStart + metrics.MeasureText("A", 10, "Arial") + metrics.MeasureText("Next", 10, "Arial"));
            double routeWidth = line ? 150 : Math.Abs(end.X - start.X);
            double innerWidth = area == "justify" ? routeWidth - 25 : intrinsic;
            double expected = area switch { "left" => 20, "right" => routeWidth - 5 - innerWidth, _ => 20 + (routeWidth - 25 - innerWidth) / 2 };
            double angle = line ? degrees * Math.PI / 180 : 0;
            string content = document.GetXml("content.xml").ToString(), styles = document.GetXml("styles.xml").ToString();
            foreach (var read in RoundTrips(document)) {
                var readShape = Assert.Single(read.Pages[0].Shapes);
                Assert.Equal(declaration, readShape.Transform);
                var drawing = read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported).Value;
                var frame = Assert.Single(Frames(drawing));
                Assert.Equal(innerWidth + 25, frame.Text.Width, 6);
                var origin = frame.Transform.TransformPoint(new OfficePoint(frame.Text.X + frame.Text.Padding.Left, frame.Text.Y));
                double along = line ? (origin.X - start.X) * Math.Cos(angle) + (origin.Y - start.Y) * Math.Sin(angle) : origin.X - Math.Min(start.X, end.X);
                Assert.Equal(expected, along, 6);
                Assert.Equal(Math.Cos(angle), frame.Transform.M11, 6); Assert.Equal(Math.Sin(angle), frame.Transform.M12, 6);
                Assert.Equal("1. A\t123,45\n2. A\tNext", frame.Text.PlainText);
                Assert.All(frame.Text.Paragraphs, p => {
                    Assert.Equal(10, p.Label!.Run.FontSize); Assert.Equal(OfficeColor.Black, p.Label.Run.Color);
                    Assert.All(p.Runs.Where(r => r.Text.Length > 0), r => Assert.Equal(10, r.FontSize));
                });
                AssertNumberedTabFields(drawing, angle, stop, bodyStart, metrics);
            }
            Assert.Equal(content, document.GetXml("content.xml").ToString());
            Assert.Equal(styles, document.GetXml("styles.xml").ToString());
        }
    }
}
