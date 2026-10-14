using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentLeaderTextStyleTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FlatAliasesRetainPartLocalLeaderBindingsAndCommonParagraphOrigins(bool commonShadow) {
        var source = OdgDocument.Create(); var page = source.AddPage();
        var body = page.Shapes.AddTextBox(Bounds(), "Body", "Body").Paragraphs[0];
        var master = page.MasterShapes.AddTextBox(Bounds(), "Master", "Master").Paragraphs[0];
        if (commonShadow) source.Styles.CreateNamed("Leader", OdfStyleFamily.Text).Color = OdfColor.Parse("0000ff");
        AddAutomaticTextStyle(source, "content.xml", "Leader", "#ff0000");
        AddAutomaticTextStyle(source, "styles.xml", "Leader", "#00ff00");
        SetAutomaticLeader(body, "Leader"); SetAutomaticLeader(master, "Leader");
        if (commonShadow) {
            var common = source.Styles.CreateNamed("CommonParagraph", OdfStyleFamily.Paragraph);
            common.Element.Add(LeaderProperties("Leader")); source.MarkPartDirty("styles.xml");
            page.Shapes.AddTextBox(Bounds(), "Common", "Common").Paragraphs[0].StyleName = common.Name;
        }
        string[] before = Parts(source);
        foreach (var read in RoundTrips(source)) {
            Assert.Equal("#FF0000", LeaderColor(read, read.Pages[0].Shapes[0].Paragraphs[0]));
            Assert.Equal("#00FF00", LeaderColor(read, read.Pages[0].MasterShapes[0].Paragraphs[0]));
            if (commonShadow) Assert.Equal("#0000FF", LeaderColor(read, read.Pages[0].Shapes[1].Paragraphs[0]));
        }
        Assert.Equal(before, Parts(source));
    }

    [Theory]
    [InlineData("page-automatic", "#FF0000")]
    [InlineData("master-automatic", "#00FF00")]
    [InlineData("common", "#0000FF")]
    [InlineData("default", "#0000FF")]
    public void PageImportRemapsLeaderStylesFromTheirOriginalAutomaticCommonOrDefaultScope(string origin, string expected) {
        var source = OdgDocument.Create(); var page = source.AddPage();
        source.Styles.CreateNamed("Leader", OdfStyleFamily.Text).Color = OdfColor.Parse("0000ff");
        AddAutomaticTextStyle(source, "content.xml", "Leader", "#ff0000");
        AddAutomaticTextStyle(source, "styles.xml", "Leader", "#00ff00");
        var paragraph = (origin == "master-automatic" ? page.MasterShapes : page.Shapes)
            .AddTextBox(Bounds(), "Row", "Row").Paragraphs[0];
        if (origin.EndsWith("automatic", StringComparison.Ordinal)) SetAutomaticLeader(paragraph, "Leader");
        else if (origin == "common") {
            var common = source.Styles.CreateNamed("CommonParagraph", OdfStyleFamily.Paragraph);
            common.Element.Add(LeaderProperties("Leader")); paragraph.StyleName = common.Name; source.MarkPartDirty("styles.xml");
        } else {
            source.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "styles")!.Add(
                new XElement(OdfNamespaces.Style + "default-style", new XAttribute(OdfNamespaces.Style + "family", "paragraph"), LeaderProperties("Leader")));
            source.MarkPartDirty("styles.xml");
        }
        string[] before = Parts(source);
        var target = OdgDocument.Create(); target.AddPage("Existing");
        target.Styles.CreateNamed("Leader", OdfStyleFamily.Text).Color = OdfColor.Parse("ffff00");
        var imported = target.ImportPage(source, 0);
        Assert.Equal(expected, LeaderColor(target, (origin == "master-automatic" ? imported.MasterShapes : imported.Shapes)[0].Paragraphs[0]));
        foreach (var read in RoundTrips(target)) {
            var readParagraph = (origin == "master-automatic" ? read.Pages[1].MasterShapes : read.Pages[1].Shapes)[0].Paragraphs[0];
            Assert.Equal(expected, LeaderColor(read, readParagraph));
        }
        Assert.Equal(before, Parts(source));
    }

    [Theory]
    [InlineData("automatic", "#FF0000")]
    [InlineData("common", "#0000FF")]
    [InlineData("default", "#0000FF")]
    public void GraphicParagraphPropertiesRetainLeaderStyleOriginsAcrossBothContainersAndPageImport(string origin, string expected) {
        var source = OdgDocument.Create(); var page = source.AddPage();
        source.Styles.CreateNamed("Leader", OdfStyleFamily.Text).Color = OdfColor.Parse("0000ff");
        AddAutomaticTextStyle(source, "content.xml", "Leader", "#ff0000");
        AddAutomaticTextStyle(source, "styles.xml", "Leader", "#00ff00");
        var shape = page.Shapes.AddTextBox(Bounds(), "Row", "Row");
        if (origin == "default") {
            source.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "styles")!.Add(
                new XElement(OdfNamespaces.Style + "default-style", new XAttribute(OdfNamespaces.Style + "family", "graphic"), LeaderProperties("Leader")));
            source.MarkPartDirty("styles.xml");
        } else {
            OdfStyle graphic = origin == "automatic" ? shape.EnsureGraphicStyle() : source.Styles.CreateNamed("CommonGraphic", OdfStyleFamily.Graphic);
            graphic.Element.Element(OdfNamespaces.Style + "paragraph-properties")?.Remove();
            graphic.Element.Add(LeaderProperties("Leader")); source.MarkPartDirty(graphic.PartPath);
            shape.Element.SetAttributeValue(OdfNamespaces.Draw + "style-name", graphic.Name); source.MarkPartDirty("content.xml");
        }
        string[] before = Parts(source);
        var sources = new[] { source }.Concat(RoundTrips(source)).ToArray();
        foreach (var read in sources) Assert.Equal(expected, LeaderColor(read, read.Pages[0].Shapes[0]));
        foreach (var read in sources) {
            string[] readBefore = Parts(read);
            var target = OdgDocument.Create(); target.AddPage("Existing");
            target.Styles.CreateNamed("Leader", OdfStyleFamily.Text).Color = OdfColor.Parse("ffff00");
            AddAutomaticTextStyle(target, "content.xml", "Leader", "#ff00ff");
            var imported = target.ImportPage(read, 0);
            Assert.Equal(expected, LeaderColor(target, imported.Shapes[0]));
            foreach (var reopened in RoundTrips(target)) Assert.Equal(expected, LeaderColor(reopened, reopened.Pages[1].Shapes[0]));
            Assert.Equal(readBefore, Parts(read));
        }
        Assert.Equal(before, Parts(source));
    }

    [Theory]
    [InlineData(OdfStyleFamily.Paragraph)]
    [InlineData(OdfStyleFamily.Graphic)]
    public void MissingCommonLeaderStyleRejectsImportBeforeMutatingEitherDocument(OdfStyleFamily family) {
        var source = OdgDocument.Create(); var page = source.AddPage();
        // Common paragraph properties cannot bind an automatic shadow when their common target is absent.
        AddAutomaticTextStyle(source, "styles.xml", "MissingCommon", "#00ff00");
        var common = source.Styles.CreateNamed("CommonStyle", family);
        common.Element.Add(LeaderProperties("MissingCommon")); source.MarkPartDirty("styles.xml");
        var shape = page.Shapes.AddTextBox(Bounds(), "Row");
        if (family == OdfStyleFamily.Paragraph) shape.Paragraphs[0].StyleName = common.Name;
        else { shape.Element.SetAttributeValue(OdfNamespaces.Draw + "style-name", common.Name); source.MarkPartDirty("content.xml"); }
        var target = OdgDocument.Create(); target.AddPage("Existing").Shapes.AddTextBox(Bounds(), "Existing text");
        string[] sourceBefore = Parts(source), targetBefore = Parts(target);
        Assert.Throws<InvalidDataException>(() => target.ImportPage(source, 0));
        Assert.Equal(sourceBefore, Parts(source)); Assert.Equal(targetBefore, Parts(target));
        Assert.Single(target.Pages);
    }

    [Fact]
    public void LeaderOutsideParagraphPropertiesRejectsImportBeforeMutatingEitherDocument() {
        var source = OdgDocument.Create(); var page = source.AddPage();
        source.Styles.CreateNamed("Leader", OdfStyleFamily.Text).Color = OdfColor.Parse("0000ff");
        var shape = page.Shapes.AddTextBox(Bounds(), "Row");
        OdfStyle graphic = shape.EnsureGraphicStyle();
        // The binding has a captured origin, but tab stops are outside their supported property context.
        graphic.Element.Element(OdfNamespaces.Style + "graphic-properties")!.Add(LeaderProperties("Leader").Elements());
        source.MarkPartDirty(graphic.PartPath);
        var target = OdgDocument.Create(); target.AddPage("Existing");
        string[] sourceBefore = Parts(source), targetBefore = Parts(target);
        Assert.Throws<NotSupportedException>(() => target.ImportPage(source, 0));
        Assert.Equal(sourceBefore, Parts(source)); Assert.Equal(targetBefore, Parts(target));
        Assert.Single(target.Pages);
    }

    private static void SetAutomaticLeader(OdfTextParagraph paragraph, string name) {
        paragraph.SetTabStops(new[] { new OdfTabStop(OdfLength.Points(80)).WithLeader(".") });
        OdfStyle style = paragraph.EnsureStyle();
        style.Element.Descendants(OdfNamespaces.Style + "tab-stop").Single().SetAttributeValue(OdfNamespaces.Style + "leader-text-style", name);
        paragraph.Document.MarkPartDirty(style.PartPath);
    }

    private static XElement LeaderProperties(string name) => new(OdfNamespaces.Style + "paragraph-properties",
        new XElement(OdfNamespaces.Style + "tab-stops", new XElement(OdfNamespaces.Style + "tab-stop",
            new XAttribute(OdfNamespaces.Style + "position", "80pt"), new XAttribute(OdfNamespaces.Style + "leader-text", "."),
            new XAttribute(OdfNamespaces.Style + "leader-text-style", name))));

    private static void AddAutomaticTextStyle(OdgDocument document, string part, string name, string color) {
        document.GetXml(part).Root!.Element(OdfNamespaces.Office + "automatic-styles")!.Add(
            new XElement(OdfNamespaces.Style + "style", new XAttribute(OdfNamespaces.Style + "name", name),
                new XAttribute(OdfNamespaces.Style + "family", "text"), new XElement(OdfNamespaces.Style + "text-properties",
                    new XAttribute(OdfNamespaces.Fo + "color", color))));
        document.MarkPartDirty(part);
    }

    private static string? LeaderColor(OdgDocument document, OdfTextParagraph paragraph) {
        OdfStyle? bound = paragraph.StyleName == null ? null : document.Styles.FindInPart(OdfStyleFamily.Paragraph, paragraph.StyleName, paragraph.PartPath);
        return LeaderColor(document, document.Styles.ResolveWithDefault(bound, OdfStyleFamily.Paragraph));
    }

    private static string? LeaderColor(OdgDocument document, OdgShape shape) =>
        LeaderColor(document, OdfTextStyleResolver.Resolve(document.Styles, shape.Paragraphs[0].Element, shape.Element, shape.PartPath));

    private static string? LeaderColor(OdgDocument document, IEnumerable<OdfStyle> styles) {
        foreach (OdfStyle style in styles) {
            XElement? stop = style.ParagraphProperties?.Element(OdfNamespaces.Style + "tab-stops")?.Elements(OdfNamespaces.Style + "tab-stop").SingleOrDefault();
            if (stop == null) continue;
            string name = (string)stop.Attribute(OdfNamespaces.Style + "leader-text-style")!;
            OdfStyle text = style.IsAutomatic
                ? document.Styles.FindInPart(OdfStyleFamily.Text, name, style.PartPath) ?? throw new InvalidDataException("The persisted automatic leader binding is missing.")
                : document.Styles.Named.Single(candidate => candidate.Family == OdfStyleFamily.Text && candidate.Name == name);
            return text.Color?.ToString();
        }
        throw new InvalidDataException("The persisted paragraph has no leader binding.");
    }

    private static OdgDocument[] RoundTrips(OdgDocument document) {
        using var flat = new MemoryStream(); document.SaveFlatXml(flat); flat.Position = 0;
        return new[] { OdgDocument.Load(new MemoryStream(document.ToBytes())), OdgDocument.LoadFlatXml(flat) };
    }

    private static string[] Parts(OdgDocument document) => new[] { "content.xml", "styles.xml" }
        .Select(part => document.GetXml(part).ToString(SaveOptions.DisableFormatting)).ToArray();
    private static OdfRect Bounds() => OdfRect.FromCentimeters(1, 1, 15, 3);
}
