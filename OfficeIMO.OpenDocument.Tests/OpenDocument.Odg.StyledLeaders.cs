using OfficeIMO.Drawing;
using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentOdgStyledLeaderTests {
    private static OdfRect Bounds() => OdfRect.FromCentimeters(1, 1, 15, 5);
    private static string[] Parts(OdgDocument document) => new[] { "content.xml", "styles.xml" }.Select(p => document.GetXml(p).ToString()).ToArray();
    private static OdgDocument[] Copies(OdgDocument document) {
        using var flat = new MemoryStream(); document.SaveFlatXml(flat); flat.Position = 0;
        return new[] { document, OdgDocument.Load(new MemoryStream(document.ToBytes())), OdgDocument.LoadFlatXml(flat) };
    }
    private static OfficeDrawingRichText Text(OdgPage page) => Assert.Single(page.ToDrawing().Value.Elements.OfType<OfficeDrawingRichText>());

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TypedStyleBindingsSurviveContainersCloneAndImportAndCompileParentRelativeSize(bool automatic) {
        var source = OdgDocument.Create(); var page = source.AddPage();
        var parent = source.Styles.CreateNamed("LeaderBase", OdfStyleFamily.Text); parent.FontSize = OdfLength.Points(20); parent.Bold = true;
        var style = automatic ? source.Styles.CreateAutomatic(OdfStyleFamily.Text, parentStyleName: parent.Name) : source.Styles.CreateNamed("Leader", OdfStyleFamily.Text, parent.Name);
        style.FontSize = OdfLength.Parse("150%"); style.Color = OdfColor.Parse("#0000ff");
        var paragraph = page.Shapes.AddTextBox(Bounds(), "A\tB").Paragraphs[0];
        var stop = new OdfTabStop(OdfLength.Points(100)).WithLeader(".").WithLeaderTextStyle(style);
        paragraph.SetTabStops(new[] { stop }); Assert.Null(stop.WithLeaderTextStyle(null).LeaderTextStyle);
        Assert.Same(style, stop.WithLeader("_").WithLineLeader(new OdfTabLineLeader()).LeaderTextStyle);
        string[] before = Parts(source);
        var reads = Copies(source); Assert.Equal(before, Parts(source));
        foreach (var read in reads) {
            string[] readBefore = Parts(read);
            var text = Text(read.Pages[0]); var leader = text.Paragraphs[0].TabStops!.Stops[0];
            Assert.Equal("A\tB", text.PlainText); Assert.Equal(30, leader.LeaderStyle!.FontSizePoints); Assert.True(leader.LeaderStyle.Bold);
            Assert.Equal(OfficeColor.Blue, leader.LeaderStyle.Color);
            var svg = XDocument.Parse(OfficeDrawingSvgExporter.ToSvg(read.Pages[0].ToDrawing().Value));
            var glyph = Assert.Single(svg.Descendants(), e => e.Name.LocalName == "text" && e.Value.Length > 0 && e.Value.All(c => c == '.'));
            Assert.Equal(OfficeColor.Blue, OfficeColor.Parse((string)glyph.Attribute("fill")!));
            Assert.Equal(readBefore, Parts(read));
            var cloned = read.ClonePage(0, "Clone"); Assert.Equal(OfficeColor.Blue, Text(cloned).Paragraphs[0].TabStops!.Stops[0].LeaderStyle!.Color);
            var target = OdgDocument.Create(); target.AddPage(); target.Styles.CreateNamed("Leader", OdfStyleFamily.Text).Color = OdfColor.Parse("#ffff00");
            var imported = target.ImportPage(read, 0); Assert.Equal(OfficeColor.Blue, Text(imported).Paragraphs[0].TabStops!.Stops[0].LeaderStyle!.Color);
        }
    }

    [Fact]
    public void PartialColorAndOpacityOverridesUseEachActualTabRun() {
        var doc = OdgDocument.Create(); var p = doc.AddPage().Shapes.AddTextBox(Bounds(), "").Paragraphs[0];
        var style = doc.Styles.CreateNamed("BlueLeader", OdfStyleFamily.Text); style.Color = OdfColor.Parse("#0000ff");
        var first = p.AddRun("A\t"); first.FontSize = OdfLength.Points(24); first.Bold = true; first.Color = OdfColor.Parse("#008800"); first.TextOpacity = .5;
        var second = p.AddRun("B\tC"); second.FontSize = OdfLength.Points(12); second.Color = OdfColor.Parse("#008800"); second.TextOpacity = .25;
        p.SetTabStops(new[] { new OdfTabStop(OdfLength.Points(100)).WithLeader(".").WithLeaderTextStyle(style), new OdfTabStop(OdfLength.Points(200)).WithLeader(".").WithLeaderTextStyle(style) });
        string[] before = Parts(doc); var result = doc.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
        var svg = XDocument.Parse(OfficeDrawingSvgExporter.ToSvg(result.Value));
        var glyphs = svg.Descendants().Where(e => e.Name.LocalName == "text" && e.Value.Length > 0 && e.Value.All(c => c == '.')).ToArray();
        Assert.Equal(2, glyphs.Length); Assert.Equal(new[] { "24", "12" }, glyphs.Select(e => (string?)e.Attribute("font-size")));
        Assert.Equal(new[] { OfficeSvgFormatting.FormatNumber(128 / 255D), OfficeSvgFormatting.FormatNumber(64 / 255D) }, glyphs.Select(e => (string?)e.Attribute("fill-opacity")));
        Assert.All(glyphs, e => Assert.Equal("#0000FF", (string?)e.Attribute("fill"))); Assert.Equal(before, Parts(doc));
    }

    [Theory]
    [InlineData("page-automatic", "#FF0000")]
    [InlineData("master-automatic", "#008800")]
    [InlineData("paragraph-common", "#0000FF")]
    [InlineData("graphic-common", "#0000FF")]
    [InlineData("graphic-automatic", "#FF0000")]
    [InlineData("paragraph-default", "#0000FF")]
    [InlineData("graphic-default", "#0000FF")]
    public void ProjectionPaintsTheOriginalLeaderBindingDespiteAutomaticAndCommonNameCollisions(string origin, string expected) {
        var doc = OdgDocument.Create(); var page = doc.AddPage();
        doc.Styles.CreateNamed("SharedLeader", OdfStyleFamily.Text).Color = OdfColor.Parse("#0000ff");
        foreach (string part in new[] { "content.xml", "styles.xml" }) {
            doc.GetXml(part).Root!.Element(OdfNamespaces.Office + "automatic-styles")!.Add(new XElement(OdfNamespaces.Style + "style",
                new XAttribute(OdfNamespaces.Style + "name", "SharedLeader"), new XAttribute(OdfNamespaces.Style + "family", "text"),
                new XElement(OdfNamespaces.Style + "text-properties", new XAttribute(OdfNamespaces.Fo + "color", part == "content.xml" ? "#ff0000" : "#008800"))));
            doc.MarkPartDirty(part);
        }
        var shape = (origin == "master-automatic" ? page.MasterShapes : page.Shapes).AddTextBox(Bounds(), "A\tB");
        var properties = new XElement(OdfNamespaces.Style + "paragraph-properties", new XElement(OdfNamespaces.Style + "tab-stops",
            new XElement(OdfNamespaces.Style + "tab-stop", new XAttribute(OdfNamespaces.Style + "position", "100pt"),
                new XAttribute(OdfNamespaces.Style + "leader-text", "."), new XAttribute(OdfNamespaces.Style + "leader-text-style", "SharedLeader"))));
        if (origin.EndsWith("default", StringComparison.Ordinal)) {
            doc.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "styles")!.Add(new XElement(OdfNamespaces.Style + "default-style",
                new XAttribute(OdfNamespaces.Style + "family", origin.StartsWith("graphic", StringComparison.Ordinal) ? "graphic" : "paragraph"), properties));
            doc.MarkPartDirty("styles.xml");
        } else {
            bool graphic = origin.StartsWith("graphic", StringComparison.Ordinal);
            OdfStyle owner = origin.EndsWith("common", StringComparison.Ordinal)
                ? doc.Styles.CreateNamed("Owner", graphic ? OdfStyleFamily.Graphic : OdfStyleFamily.Paragraph)
                : graphic ? shape.EnsureGraphicStyle() : shape.Paragraphs[0].EnsureStyle();
            owner.Element.Element(OdfNamespaces.Style + "paragraph-properties")?.Remove(); owner.Element.Add(properties); doc.MarkPartDirty(owner.PartPath);
            if (graphic) shape.Element.SetAttributeValue(OdfNamespaces.Draw + "style-name", owner.Name);
            else shape.Paragraphs[0].StyleName = owner.Name;
        }
        foreach (var read in Copies(doc)) {
            string[] before = Parts(read);
            var svg = XDocument.Parse(OfficeDrawingSvgExporter.ToSvg(read.Pages[0].ToDrawing().Value));
            var glyph = Assert.Single(svg.Descendants(), e => e.Name.LocalName == "text" && e.Value.Length > 0 && e.Value.All(c => c == '.'));
            Assert.Equal(OfficeColor.Parse(expected), OfficeColor.Parse((string)glyph.Attribute("fill")!)); Assert.Equal(before, Parts(read));
            var destination = OdgDocument.Create(); destination.AddPage(); destination.Styles.CreateNamed("SharedLeader", OdfStyleFamily.Text).Color = OdfColor.Parse("#ffff00");
            var imported = destination.ImportPage(read, 0);
            var importedSvg = XDocument.Parse(OfficeDrawingSvgExporter.ToSvg(imported.ToDrawing().Value));
            var importedGlyph = Assert.Single(importedSvg.Descendants(), e => e.Name.LocalName == "text" && e.Value.Length > 0 && e.Value.All(c => c == '.'));
            Assert.Equal(OfficeColor.Parse(expected), OfficeColor.Parse((string)importedGlyph.Attribute("fill")!));
        }
    }

    [Fact]
    public void InvalidTypedBindingsRejectBeforeMutatingTheParagraphOrPackage() {
        var doc = OdgDocument.Create(); var p = doc.AddPage().Shapes.AddTextBox(Bounds(), "A\tB").Paragraphs[0];
        var foreign = OdgDocument.Create().Styles.CreateNamed("Foreign", OdfStyleFamily.Text);
        var wrongFamily = doc.Styles.CreateNamed("WrongFamily", OdfStyleFamily.Paragraph);
        string[] before = Parts(doc);
        Assert.Throws<ArgumentException>(() => p.SetTabStops(new[] { new OdfTabStop(OdfLength.Points(100)).WithLeader(".").WithLeaderTextStyle(foreign) }));
        Assert.Throws<ArgumentException>(() => new OdfTabStop(OdfLength.Points(100)).WithLeaderTextStyle(wrongFamily));
        Assert.Equal(before, Parts(doc));
    }

    [Fact]
    public void TabsAddedAfterFontEditingSerializeParagraphPropertiesBeforeTextProperties() {
        var doc = OdgDocument.Create(); var p = doc.AddPage().Shapes.AddTextBox(Bounds(), "A\tB").Paragraphs[0];
        p.FontSize = OdfLength.Points(14);
        p.SetTabStops(new[] { new OdfTabStop(OdfLength.Points(100)).WithLeader(".") });
        using var output = new MemoryStream(); doc.SaveFlatXml(output); output.Position = 0;
        var xml = XDocument.Load(output);
        var style = Assert.Single(xml.Descendants(OdfNamespaces.Style + "style"), e => (string?)e.Attribute(OdfNamespaces.Style + "name") == p.StyleName);
        var properties = style.Elements().Select(e => e.Name).ToArray();
        Assert.True(Array.IndexOf(properties, OdfNamespaces.Style + "paragraph-properties") < Array.IndexOf(properties, OdfNamespaces.Style + "text-properties"));
    }
}
