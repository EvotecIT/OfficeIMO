using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed partial class OpenDocumentOdgImportTests {
    [Theory]
    [InlineData(false, "paragraph")]
    [InlineData(true, "paragraph")]
    [InlineData(false, "graphic")]
    public void FontAliasesDoNotPromoteDefaultFontAboveExplicitGraphicFont(bool explicitReference, string defaultFamily) {
        var source = OdgDocument.Create(); var page = source.AddPage(); var shape = page.Shapes.AddTextBox(Rect(0, 0, 200, 60), "Font aliases");
        source.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "font-face-decls")!.Add(new XElement(OdfNamespaces.Style + "font-face",
            new XAttribute(OdfNamespaces.Style + "name", "Face"), new XAttribute(OdfNamespaces.Svg + "font-family", explicitReference ? "Arial" : "DejaVu Sans")));
        GraphicStyle(source, shape).SetProperty(OdfNamespaces.Style + "text-properties", explicitReference ? OdfNamespaces.Style + "font-name" : OdfNamespaces.Fo + "font-family", explicitReference ? "Face" : "Arial");
        AddTextDefault(source, defaultFamily, new XAttribute(explicitReference ? OdfNamespaces.Fo + "font-family" : OdfNamespaces.Style + "font-name", explicitReference ? "DejaVu Sans" : "Face"));
        var target = OdgDocument.Create(); var imported = target.ImportPage(source, 0);
        Assert.All(DrawingRuns(page), run => Assert.Equal("Arial", run.FontFamily)); Assert.Equal(Svg(page), Svg(imported));
        foreach (var read in RoundTrips(target)) Assert.Equal(Svg(page), Svg(read.Pages[0]));
    }

    [Theory]
    [InlineData("graphic", 30)]
    [InlineData("inline", 30)]
    [InlineData("nested", 45)]
    public void RelativeFontSizesRetainTheirSourceDefaultBase(string mode, double expected) {
        var source = OdgDocument.Create(); var page = source.AddPage(); var shape = page.Shapes.AddTextBox(Rect(0, 0, 200, 60), "Relative size");
        XElement paragraph = shape.Element.Descendants(OdfNamespaces.Text + "p").Single();
        if (mode == "graphic") shape.FontSize = OdfLength.Parse("150%");
        else {
            source.Styles.CreateNamed("Relative", OdfStyleFamily.Text).FontSize = OdfLength.Parse("150%");
            XElement span = new XElement(OdfNamespaces.Text + "span", new XAttribute(OdfNamespaces.Text + "style-name", "Relative"), "Relative size");
            if (mode == "nested") span = new XElement(OdfNamespaces.Text + "span", new XAttribute(OdfNamespaces.Text + "style-name", "Relative"), span);
            paragraph.RemoveNodes(); paragraph.Add(span); source.MarkPartDirty("content.xml");
        }
        foreach (string family in new[] { "paragraph", "text" }) AddTextDefault(source, family, new XAttribute(OdfNamespaces.Fo + "font-size", "20pt"));
        var target = OdgDocument.Create(); var imported = target.ImportPage(source, 0);
        Assert.All(DrawingRuns(page), run => Assert.Equal(expected, run.FontSize)); Assert.Equal(Svg(page), Svg(imported));
        foreach (var read in RoundTrips(target)) Assert.Equal(Svg(page), Svg(read.Pages[0]));
    }

    [Fact]
    public void PercentageWithExplicitParagraphBaseRemainsRelativeAndEditable() {
        var source = OdgDocument.Create(); var page = source.AddPage(); var shape = page.Shapes.AddTextBox(Rect(0, 0, 200, 60), "Explicit base");
        source.Styles.CreateNamed("Paragraph", OdfStyleFamily.Paragraph).FontSize = OdfLength.Points(20);
        source.Styles.CreateNamed("Relative", OdfStyleFamily.Text).FontSize = OdfLength.Parse("150%");
        XElement paragraph = shape.Element.Descendants(OdfNamespaces.Text + "p").Single(); paragraph.SetAttributeValue(OdfNamespaces.Text + "style-name", "Paragraph");
        paragraph.RemoveNodes(); paragraph.Add(new XElement(OdfNamespaces.Text + "span", new XAttribute(OdfNamespaces.Text + "style-name", "Relative"), "Explicit base")); source.MarkPartDirty("content.xml");
        var target = OdgDocument.Create(); var imported = target.ImportPage(source, 0); Assert.Equal(Svg(page), Svg(imported));
        XElement importedParagraph = imported.Shapes[0].Element.Descendants(OdfNamespaces.Text + "p").Single();
        target.Styles.FindInPart(OdfStyleFamily.Paragraph, (string)importedParagraph.Attribute(OdfNamespaces.Text + "style-name")!, "content.xml")!.FontSize = OdfLength.Points(24);
        Assert.All(DrawingRuns(imported), run => Assert.Equal(36, run.FontSize));
        foreach (var read in RoundTrips(target)) Assert.All(DrawingRuns(read.Pages[0]), run => Assert.Equal(36, run.FontSize));
    }

    [Theory]
    [InlineData("underline")]
    [InlineData("line-through")]
    public void DisabledDecorationTypeIsNotReenabledBySourceDefaultStyle(string decoration) {
        var source = OdgDocument.Create(); var page = source.AddPage(); var shape = page.Shapes.AddTextBox(Rect(0, 0, 200, 60), "No decoration");
        GraphicStyle(source, shape).SetProperty(OdfNamespaces.Style + "text-properties", OdfNamespaces.Style + ("text-" + decoration + "-type"), "none");
        AddTextDefault(source, "paragraph", new XAttribute(OdfNamespaces.Style + ("text-" + decoration + "-style"), "solid"));
        var target = OdgDocument.Create(); var imported = target.ImportPage(source, 0);
        Assert.Equal(Svg(page), Svg(imported)); foreach (var read in RoundTrips(target)) Assert.Equal(Svg(page), Svg(read.Pages[0]));
    }

    [Fact]
    public void CoupledPropertiesFromTheSameDefaultAreCopiedTogetherInEitherFamily() {
        foreach (string family in new[] { "paragraph", "graphic" }) {
            var source = OdgDocument.Create(); var page = source.AddPage(); page.Shapes.AddTextBox(Rect(0, 0, 200, 60), "Grouped defaults");
            source.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "font-face-decls")!.Add(new XElement(OdfNamespaces.Style + "font-face",
                new XAttribute(OdfNamespaces.Style + "name", "Face"), new XAttribute(OdfNamespaces.Svg + "font-family", "DejaVu Sans")));
            AddTextDefault(source, family, new XAttribute(OdfNamespaces.Style + "font-name", "Face"), new XAttribute(OdfNamespaces.Fo + "font-family", "Arial"),
                new XAttribute(OdfNamespaces.Style + "text-underline-style", "solid"), new XAttribute(OdfNamespaces.Style + "text-underline-type", "none"));
            var target = OdgDocument.Create(); var imported = target.ImportPage(source, 0);
            Assert.Equal(Svg(page), Svg(imported)); foreach (var read in RoundTrips(target)) Assert.Equal(Svg(page), Svg(read.Pages[0]));
        }
    }

    [Fact]
    public void PercentageFontCombinedWithRelativeSizeChangeIsRejectedBeforeAttachment() {
        var source = OdgDocument.Create(); var shape = source.AddPage().Shapes.AddTextBox(Rect(0, 0, 200, 60), "Relative change");
        shape.FontSize = OdfLength.Parse("150%"); GraphicStyle(source, shape).SetProperty(OdfNamespaces.Style + "text-properties", OdfNamespaces.Style + "font-size-rel", "2pt");
        var target = OdgDocument.Create(); string[] before = Parts(target);
        Assert.Throws<System.NotSupportedException>(() => target.ImportPage(source, 0)); Assert.Equal(before, Parts(target));
    }

    private static OdfStyle GraphicStyle(OdgDocument document, OdgShape shape) => document.Styles.FindInPart(OdfStyleFamily.Graphic,
        (string)shape.Element.Attribute(OdfNamespaces.Draw + "style-name")!, "content.xml")!;
    private static void AddTextDefault(OdgDocument document, string family, params XAttribute[] attributes) {
        document.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "styles")!.Add(new XElement(OdfNamespaces.Style + "default-style",
            new XAttribute(OdfNamespaces.Style + "family", family), new XElement(OdfNamespaces.Style + "text-properties", attributes)));
        document.MarkPartDirty("styles.xml");
    }
    private static System.Collections.Generic.IEnumerable<OfficeRichTextRun> DrawingRuns(OdgPage page) => page.ToDrawing().Value.Elements.OfType<OfficeDrawingRichText>()
        .SelectMany(text => text.Paragraphs).SelectMany(paragraph => paragraph.Runs);
}
