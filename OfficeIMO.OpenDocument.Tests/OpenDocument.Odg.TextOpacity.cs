using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentOdgTextOpacityTests {
    private static readonly XNamespace LoExt = "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0";

    [Theory]
    [InlineData("0%", 0)]
    [InlineData("50%", 128)]
    [InlineData("100%", 255)]
    public void ImportedBodyOpacityRetainsTextAndDeclarationsThroughBothContainers(string value, byte alpha) {
        var document = OdgDocument.LoadFlatXml(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Drawing", "libreoffice-transparent-text.fodg"));
        var run = document.Pages[0].Shapes[0].Paragraphs[0].Runs.Single();
        run.EnsureStyle().SetProperty(OdfNamespaces.Style + "text-properties", LoExt + "opacity", value);
        string before = document.GetXml("content.xml").ToString();
        foreach (var read in RoundTrips(document)) {
            var result = read.Pages[0].ToDrawing();
            var text = Text(result.Value);
            Assert.Equal("asdf", text.PlainText);
            Assert.Equal(OfficeColor.FromRgba(255, 0, 0, alpha), text.Paragraphs[0].Runs.Single().Color);
            Assert.Equal(value, read.Pages[0].Shapes[0].Paragraphs[0].Runs.Single().Styles.Select(s =>
                (string?)s.TextProperties?.Attribute(LoExt + "opacity")).First(v => v != null));
            Assert.DoesNotContain(result.Report.Mappings, m => m.Feature.EndsWith(":text-opacity", StringComparison.Ordinal));
            var paint = XDocument.Parse(OfficeDrawingSvgExporter.ToSvg(result.Value)).Descendants(XName.Get("text", "http://www.w3.org/2000/svg"));
            if (alpha == 0) { Assert.Empty(paint); continue; }
            var rendered = paint.Single();
            Assert.Equal("asdf", rendered.Value);
            Assert.Equal(OfficeSvgFormatting.FormatNumber(alpha / 255D), (string?)rendered.Attribute("fill-opacity") ?? "1");
        }
        Assert.Equal(before, document.GetXml("content.xml").ToString());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void BodyOpacityUsesTheNearestStyleAcrossNamespaceAliasesAndNestedLinks(bool alternate) {
        var document = OdgDocument.Create(); var shape = Shape(document);
        var parent = document.Styles.CreateNamed("Parent", OdfStyleFamily.Paragraph);
        parent.Color = OdfColor.Parse("#FF0000"); parent.FontSize = OdfLength.Points(14);
        parent.SetProperty(OdfNamespaces.Style + "text-properties", (alternate ? LoExt : OdfNamespaces.Draw) + "opacity", "50%");
        var paragraph = shape.AddParagraph("P"); paragraph.StyleName = parent.Name;
        var first = paragraph.AddRun("A"); Set(first, alternate ? OdfNamespaces.Draw : LoExt, "25%");
        var nested = first.AddRun("B"); Set(nested, alternate ? LoExt : OdfNamespaces.Draw, "100%");
        var link = paragraph.AddHyperlink("C", "https://example.invalid/"); Set(link, LoExt, "0%");
        foreach (var read in RoundTrips(document)) {
            var result = read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
            var text = Text(result.Value); Assert.Equal("PABC", text.PlainText);
            Assert.Equal(new byte[] { 128, 64, 255, 0 }, text.Paragraphs[0].Runs.Select(r => r.Color.A));
            Assert.All(text.Paragraphs[0].Runs, r => Assert.Equal((byte)255, r.Color.R));
        }
    }

    [Theory]
    [InlineData("loext", "101%")]
    [InlineData("draw", "0.5")]
    [InlineData("loext", "NaN%")]
    [InlineData("loext", "1e1%")]
    [InlineData("foreign", "25%")]
    public void InvalidOrUnknownOpacityKeepsOpaqueTextWithAnExplicitLoss(string prefix, string value) {
        var document = OdgDocument.Create(); var shape = Shape(document);
        var p = shape.AddParagraph("Retained"); p.Color = OdfColor.Parse("#FF0000");
        XNamespace ns = prefix == "loext" ? LoExt : prefix == "draw" ? OdfNamespaces.Draw : "urn:example:unqualified";
        Set(p, ns, value);
        foreach (var read in RoundTrips(document)) {
            var result = read.Pages[0].ToDrawing();
            Assert.Equal(OfficeColor.Red, Text(result.Value).Paragraphs[0].Runs.Single().Color);
            Assert.Contains(result.Report.Mappings, m => m.Feature.EndsWith(":text-opacity", StringComparison.Ordinal) && m.Status == OdfConversionMappingStatus.Unsupported);
            Assert.Throws<OdfConversionLossException>(() => read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
            Assert.Equal(value, read.Pages[0].Shapes[0].Paragraphs[0].Styles.Select(s => (string?)s.TextProperties?.Attribute(ns + "opacity")).First(v => v != null));
        }
    }

    [Fact]
    public void ConflictingAliasesKeepTheLossWhileEquivalentValuesAreResolved() {
        foreach (string draw in new[] { "25%", "50.0%" }) {
            var document = OdgDocument.Create(); var shape = Shape(document);
            var p = shape.AddParagraph("Aliases"); p.Color = OdfColor.Parse("#FF0000");
            Set(p, LoExt, "50%"); Set(p, OdfNamespaces.Draw, draw);
            foreach (var read in RoundTrips(document)) {
                var result = read.Pages[0].ToDrawing();
                Assert.Equal(draw == "25%" ? (byte)255 : (byte)128, Text(result.Value).Paragraphs[0].Runs.Single().Color.A);
                if (draw == "25%") Assert.Throws<OdfConversionLossException>(() => read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
                else read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
            }
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OpacityWithoutAQualifiedForegroundRetainsItsColorFallback(bool windowColor) {
        var document = OdgDocument.Create(); var shape = Shape(document);
        var p = shape.AddParagraph("Defaults"); Set(p, LoExt, "25%");
        if (windowColor) {
            p.Color = OdfColor.Parse("#FF0000");
            p.EnsureStyle().SetProperty(OdfNamespaces.Style + "text-properties", OdfNamespaces.Style + "use-window-font-color", "true");
        }
        foreach (var read in RoundTrips(document)) {
            var result = read.Pages[0].ToDrawing();
            Assert.Equal(windowColor ? OfficeColor.Red : OfficeColor.Black, Text(result.Value).Paragraphs[0].Runs.Single().Color);
            Assert.Contains(result.Report.Mappings, m => m.Feature.EndsWith(":text-opacity", StringComparison.Ordinal) && m.Status == OdfConversionMappingStatus.Unsupported);
            Assert.Throws<OdfConversionLossException>(() => read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
        }
    }

    [Theory]
    [InlineData("paragraph")]
    [InlineData("level")]
    [InlineData("named")]
    [InlineData("window")]
    public void ExplicitListForegroundKeepsAnOpaqueMarkerAboveTransparentParagraphDefaults(string foreground) {
        var document = OdgDocument.Create(); var shape = Shape(document);
        shape.AddList(true).AddItem("Body");
        var p = shape.Paragraphs.Single(); p.Color = OdfColor.Parse("#000000"); Set(p, LoExt, "0%");
        p.Text = string.Empty; var body = p.AddRun("Body"); Set(body, LoExt, "100%");
        var level = document.GetXml("content.xml").Descendants(OdfNamespaces.Text + "list-style").Single().Elements().Single();
        if (foreground == "named") {
            var named = document.Styles.CreateNamed("LabelForeground", OdfStyleFamily.Text);
            named.Color = OdfColor.Parse("#000000");
            named.SetProperty(OdfNamespaces.Style + "text-properties", LoExt + "opacity", "50%");
            level.SetAttributeValue(OdfNamespaces.Text + "style-name", named.Name);
        }
        level.Add(new XElement(OdfNamespaces.Style + "text-properties",
            foreground is "level" or "window" ? new XAttribute(OdfNamespaces.Fo + "color", "#000000") : null,
            foreground == "window" ? new XAttribute(OdfNamespaces.Style + "use-window-font-color", "true") : null,
            new XAttribute(OdfNamespaces.Fo + "font-family", "Arial"), new XAttribute(OdfNamespaces.Fo + "font-size", "100%")));
        foreach (var read in RoundTrips(document)) {
            var result = read.Pages[0].ToDrawing(); var paragraph = Text(result.Value).Paragraphs.Single();
            Assert.Equal(OfficeColor.Black, paragraph.Label!.Run.Color); Assert.Equal(OfficeColor.Black, paragraph.Runs.Single().Color);
            Assert.Equal("1.", paragraph.Label.Run.Text);
            if (foreground == "level") {
                Assert.DoesNotContain(result.Report.Mappings, m => m.Feature.EndsWith(":text-opacity", StringComparison.Ordinal));
                read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
            } else {
                Assert.Contains(result.Report.Mappings, m => m.Feature.EndsWith(":text-opacity", StringComparison.Ordinal) && m.Status == OdfConversionMappingStatus.Unsupported);
                Assert.Throws<OdfConversionLossException>(() => read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
            }
        }
    }

    [Theory]
    [InlineData("0%", 0)]
    [InlineData("25%", 64)]
    [InlineData("50%", 128)]
    [InlineData("100%", 255)]
    public void ExplicitListForegroundIsIndependentOfBodyAndLevelOpacity(string opacity, byte bodyAlpha) {
        foreach (bool levelOpacity in new[] { false, true }) {
            var document = OdgDocument.Create(); var shape = Shape(document);
            shape.AddList(true).AddItem("Body"); var p = shape.Paragraphs.Single(); p.Color = OdfColor.Parse("#000000");
            p.Text = string.Empty; var body = p.AddRun("Body"); Set(body, LoExt, levelOpacity ? "100%" : opacity);
            var level = document.GetXml("content.xml").Descendants(OdfNamespaces.Text + "list-style").Single().Elements().Single();
            level.Add(new XElement(OdfNamespaces.Style + "text-properties", new XAttribute(OdfNamespaces.Fo + "color", "#000000"),
                new XAttribute(OdfNamespaces.Fo + "font-family", "Arial"), new XAttribute(OdfNamespaces.Fo + "font-size", "100%"),
                levelOpacity ? new XAttribute(LoExt + "opacity", opacity) : null));
            string before = document.GetXml("content.xml").ToString();
            foreach (var read in RoundTrips(document)) {
                var result = read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
                var paragraph = Text(result.Value).Paragraphs.Single();
                Assert.Equal(OfficeColor.Black, paragraph.Label!.Run.Color);
                Assert.Equal(levelOpacity ? (byte)255 : bodyAlpha, paragraph.Runs.Single().Color.A);
                var paint = XDocument.Parse(OfficeDrawingSvgExporter.ToSvg(result.Value)).Descendants(XName.Get("text", "http://www.w3.org/2000/svg"));
                Assert.Equal("1.", paint.First().Value); Assert.Null(paint.First().Attribute("fill-opacity"));
                if (levelOpacity) Assert.Equal(opacity, (string?)read.GetXml("content.xml")
                    .Descendants(OdfNamespaces.Text + "list-style").Single().Elements().Single()
                    .Element(OdfNamespaces.Style + "text-properties")!.Attribute(LoExt + "opacity"));
            }
            Assert.Equal(before, document.GetXml("content.xml").ToString());
        }
    }

    private static void Set(OdfTextContent content, XNamespace ns, string value) =>
        content.EnsureStyle().SetProperty(OdfNamespaces.Style + "text-properties", ns + "opacity", value);
    private static OdgShape Shape(OdgDocument document) {
        var shape = document.AddPage().Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 12, 8));
        shape.FontFamily = "Arial"; shape.FontSize = OdfLength.Points(14);
        shape.EnsureGraphicStyle().SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Fo + "wrap-option", "no-wrap");
        return shape;
    }
    private static OfficeDrawingRichText Text(OfficeDrawing drawing) => Assert.Single(drawing.Elements.OfType<OfficeDrawingRichText>());
    private static OdgDocument[] RoundTrips(OdgDocument document) {
        using var flat = new MemoryStream(); document.SaveFlatXml(flat); flat.Position = 0;
        return new[] { document, OdgDocument.Load(new MemoryStream(document.ToBytes(new OdfSaveOptions { CompatibilityProfile = OdfCompatibilityProfile.PreserveSource }))), OdgDocument.LoadFlatXml(flat) };
    }
}
