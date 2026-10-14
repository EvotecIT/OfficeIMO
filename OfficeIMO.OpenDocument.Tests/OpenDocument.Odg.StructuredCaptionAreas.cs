using System;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
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
    public void ListBodyTabsKeepTheParagraphOriginWhenAreaAlignmentAndRouteInsetsCombine(string area, bool line) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var shape = line ? page.Shapes.AddLine(OdfLength.Points(100), OdfLength.Points(100), OdfLength.Points(190), OdfLength.Points(220)) :
            page.Shapes.AddConnector(new OfficePoint(100, 100), new OfficePoint(250, 100));
        ConfigureNumberedTabBody(document, shape);
        var graphic = Graphic(document, shape);
        graphic.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "textarea-horizontal-align", area);
        foreach (var side in new[] { ("left", 20D), ("top", 3D), ("right", 5D), ("bottom", 11D) })
            graphic.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Fo + "padding-" + side.Item1, OdfLength.Points(side.Item2).ToString());
        var metrics = new OfficeRasterCanvas(new OfficeRasterImage(1, 1));
        double stop = OdfLength.Centimeters(1.5).ToPoints();
        double bodyStart = OdfLength.Centimeters(.6).ToPoints();
        double current = bodyStart + metrics.MeasureText("A", 10, "Arial");
        // A character-aligned field cannot move backward over its preceding body.
        double intrinsicWidth = Math.Max(Math.Max(current, stop - metrics.MeasureText("123", 10, "Arial")) + metrics.MeasureText("123,45", 10, "Arial"),
            Math.Max(current, stop - metrics.MeasureText("Next", 10, "Arial")) + metrics.MeasureText("Next", 10, "Arial"));
        double angle = line ? Math.Atan2(.8, .6) : 0;
        string before = document.GetXml("content.xml").ToString();
        foreach (var read in RoundTrips(document)) {
            var drawing = read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported).Value;
            var frame = Assert.Single(Frames(drawing));
            double innerWidth = area == "justify" ? 125 : intrinsicWidth;
            Assert.Equal(innerWidth + 25, frame.Text.Width, 6);
            var origin = frame.Transform.TransformPoint(new OfficePoint(frame.Text.X + frame.Text.Padding.Left, frame.Text.Y));
            double along = (origin.X - 100) * Math.Cos(angle) + (origin.Y - 100) * Math.Sin(angle);
            double expected = area switch { "left" => 20, "right" => 145 - innerWidth, _ => 20 + (125 - innerWidth) / 2 };
            Assert.Equal(expected, along, 6);
            Assert.Equal("1. A\t123,45\n2. A\tNext", frame.Text.PlainText);
            AssertNumberedTabFields(drawing, angle, stop, bodyStart, metrics);
        }
        Assert.Equal(before, document.GetXml("content.xml").ToString());
    }

    [Fact]
    public void OrdinaryShapeListTabsAlsoKeepTheOriginBeforeGeneratedBodyIndent() {
        var document = OdgDocument.Create(); var shape = document.AddPage().Shapes.AddTextBox(
            new OdfRect(OdfLength.Points(100), OdfLength.Points(100), OdfLength.Points(180), OdfLength.Points(100)), string.Empty);
        ConfigureNumberedTabBody(document, shape);
        foreach (var read in RoundTrips(document)) AssertNumberedTabFields(
            read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported).Value, 0,
            OdfLength.Centimeters(1.5).ToPoints(), OdfLength.Centimeters(.6).ToPoints(), new OfficeRasterCanvas(new OfficeRasterImage(1, 1)));
    }

    [Theory]
    [InlineData("space", false)]
    [InlineData("nothing", false)]
    [InlineData("listtab", false)]
    [InlineData("space", true)]
    [InlineData("nothing", true)]
    [InlineData("listtab", true)]
    public void ModernListsResolveTheBodyTabMarginForLabelsAndContinuationParagraphs(string separator, bool line) {
        foreach (double? declaredMargin in new double?[] { null, 12 }) foreach (bool relative in new[] { true, false }) {
            var document = OdgDocument.Create(); var page = document.AddPage();
            var shape = line ? page.Shapes.AddLine(OdfLength.Points(100), OdfLength.Points(100), OdfLength.Points(190), OdfLength.Points(220)) :
                page.Shapes.AddTextBox(new OdfRect(OdfLength.Points(100), OdfLength.Points(100), OdfLength.Points(180), OdfLength.Points(100)), string.Empty);
            ConfigureNumberedTabBody(document, shape);
            shape.Paragraphs[1].Text = "C\t123,45";
            var first = shape.TextRoot.Descendants(OdfNamespaces.Text + "p").First();
            var continuation = new XElement(first); first.AddAfterSelf(continuation);
            new OdfTextParagraph(document, continuation, shape.Element).Text = "B\t123,45";
            foreach (var paragraph in shape.Paragraphs) {
                paragraph.MarginLeft = declaredMargin.HasValue ? OdfLength.Points(declaredMargin.Value) : null;
                paragraph.SetTabStops(new[] { new OdfTabStop(OdfLength.Points(80), "char", ",") });
            }
            var properties = document.GetXml("content.xml").Descendants(OdfNamespaces.Style + "list-level-properties").Single();
            properties.ReplaceAttributes(new XAttribute(OdfNamespaces.Text + "list-level-position-and-space-mode", "label-alignment"));
            properties.Add(new XElement(OdfNamespaces.Style + "list-level-label-alignment",
                new XAttribute(OdfNamespaces.Fo + "margin-left", "0.6cm"), new XAttribute(OdfNamespaces.Fo + "text-indent", "0pt"),
                new XAttribute(OdfNamespaces.Text + "label-followed-by", separator), new XAttribute(OdfNamespaces.Text + "list-tab-stop-position", "0.6cm")));
            document.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "styles")!.Add(new XElement(OdfNamespaces.Style + "default-style",
                new XAttribute(OdfNamespaces.Style + "family", "paragraph"), new XElement(OdfNamespaces.Style + "paragraph-properties",
                    new XAttribute(OdfNamespaces.Text + "relative-tab-stop-position", relative ? "true" : "false"))));
            double margin = declaredMargin ?? OdfLength.Centimeters(.6).ToPoints();
            double angle = line ? Math.Atan2(.8, .6) : 0;
            double expected = (relative ? margin : 0) + 80 - new OfficeRasterCanvas(new OfficeRasterImage(1, 1)).MeasureText("123", 10, "Arial");
            string before = document.GetXml("content.xml").ToString();
            foreach (var read in RoundTrips(document)) {
                var result = read.Pages[0].ToDrawing(); var frame = Assert.Single(Frames(result.Value));
                Assert.Equal(new[] { "1.", null, "2." }, frame.Text.Paragraphs.Select(p => p.Label?.Run.Text));
                Assert.All(frame.Text.Paragraphs, p => Assert.Equal(margin, p.Margins.Left, 6));
                if (separator == "listtab") {
                    Assert.EndsWith(":list-tab-stops", Assert.Single(result.Report.Mappings, m => m.Status == OdfConversionMappingStatus.Unsupported).Feature);
                    Assert.Throws<OdfConversionLossException>(() => read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
                } else read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
                var origin = frame.Transform.TransformPoint(new OfficePoint(frame.Text.X + frame.Text.Padding.Left, frame.Text.Y));
                XNamespace svg = "http://www.w3.org/2000/svg";
                var fields = XDocument.Parse(OfficeDrawingSvgExporter.ToSvg(result.Value)).Descendants(svg + "text").Where(n => n.Value == "123,45").ToArray();
                Assert.Equal(3, fields.Length);
                foreach (var field in fields) {
                    var actual = StructuredTextOrigin(field);
                    double along = (actual.X - origin.X) * Math.Cos(angle) + (actual.Y - origin.Y) * Math.Sin(angle);
                    Assert.InRange(Math.Abs(along - expected), 0, .02);
                }
            }
            Assert.Equal(before, document.GetXml("content.xml").ToString());
        }
    }

    private static void ConfigureNumberedTabBody(OdgDocument document, OdgShape shape) {
        shape.TextRoot.RemoveNodes(); var list = shape.AddList(true);
        list.AddItem("A\t123,45"); list.AddItem("A\tNext");
        foreach (var paragraph in shape.Paragraphs) {
            paragraph.FontFamily = "Arial"; paragraph.FontSize = OdfLength.Points(10); paragraph.LineHeight = OdfLength.Points(11);
            paragraph.TextAlign = "left";
            paragraph.SetTabStops(new[] { new OdfTabStop(OdfLength.Centimeters(1.5), "char", ",") });
        }
        var level = document.GetXml("content.xml").Descendants(OdfNamespaces.Text + "list-style").Single().Elements().Single();
        level.Element(OdfNamespaces.Style + "list-level-properties")!.SetAttributeValue(OdfNamespaces.Text + "min-label-width", "0.6cm");
        level.Add(new XElement(OdfNamespaces.Style + "text-properties", new XAttribute(OdfNamespaces.Fo + "font-family", "Arial"),
            new XAttribute(OdfNamespaces.Fo + "font-size", "100%")));
        Graphic(document, shape).SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Fo + "wrap-option", "no-wrap");
    }

    private static void AssertNumberedTabFields(OfficeDrawing drawing, double angle, double stop, double bodyStart, OfficeRasterCanvas metrics) {
        XNamespace svg = "http://www.w3.org/2000/svg";
        var nodes = XDocument.Parse(OfficeDrawingSvgExporter.ToSvg(drawing)).Descendants(svg + "text").ToArray();
        var starts = nodes.Where(n => n.Value == "A").Select(StructuredTextOrigin).ToArray(); Assert.Equal(2, starts.Length);
        foreach (var field in new[] { ("123,45", 0, "123"), ("Next", 1, "Next") }) {
            var actual = StructuredTextOrigin(Assert.Single(nodes, n => n.Value == field.Item1));
            double along = (actual.X - starts[field.Item2].X) * Math.Cos(angle) + (actual.Y - starts[field.Item2].Y) * Math.Sin(angle);
            double expected = Math.Max(metrics.MeasureText("A", 10, "Arial"), stop - bodyStart - metrics.MeasureText(field.Item3, 10, "Arial"));
            Assert.InRange(Math.Abs(along - expected), 0, .01);
        }
    }

    private static OfficePoint StructuredTextOrigin(XElement node) {
        var point = new OfficePoint(double.Parse(node.Attribute("x")!.Value, CultureInfo.InvariantCulture), double.Parse(node.Attribute("y")!.Value, CultureInfo.InvariantCulture));
        foreach (var parent in node.Ancestors()) if (parent.Attribute("transform") is { } value) {
            var numbers = value.Value.Substring(7, value.Value.Length - 8).Split(new[] { ' ', ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(s => double.Parse(s, CultureInfo.InvariantCulture)).ToArray();
            point = new OfficeTransform(numbers[0], numbers[1], numbers[2], numbers[3], numbers[4], numbers[5]).TransformPoint(point);
        }
        return point;
    }
}
