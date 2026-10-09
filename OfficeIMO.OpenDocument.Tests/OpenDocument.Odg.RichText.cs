using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public class OpenDocumentOdgRichTextTests {
    [Fact]
    public void TargetedRunEditInIndependentDrawPreservesGeometryAndNativeStyles() {
        var document = OdgDocument.LoadFlatXml(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Drawing", "libreoffice-transparent-text.fodg"));
        var shape = document.Pages[0].Shapes[0];
        XElement geometry = new XElement(shape.ToXml().Element(OdfNamespaces.Draw + "enhanced-geometry")!);
        var styles = document.Styles.Named.Concat(document.Styles.Automatic).ToDictionary(s => (s.Family, s.Name), s => StyleProperties(s.Element));
        XElement metadata = Payload(document, "meta.xml");
        var run = Assert.Single(Assert.Single(shape.Paragraphs).Runs);
        string? originalStyle = run.StyleName;
        Assert.Equal("asdf", run.Text); Assert.NotNull(run.FontSize);
        run.Text = "A  native\tedit\nżółć";
        foreach (var read in RoundTrips(document)) {
            var reopened = read.Pages[0].Shapes[0];
            Assert.Equal(run.Text, reopened.Text);
            Assert.Equal(originalStyle, Assert.Single(Assert.Single(reopened.Paragraphs).Runs).StyleName);
            Assert.True(XNode.DeepEquals(geometry, reopened.ToXml().Element(OdfNamespaces.Draw + "enhanced-geometry")));
            foreach (var style in styles) Assert.True(XNode.DeepEquals(style.Value, StyleProperties(read.Styles.Find(style.Key.Family, style.Key.Name)!.Element)));
            Assert.Equal(run.FontFamily, reopened.Paragraphs.Single().Runs.Single().FontFamily);
            Assert.True(XNode.DeepEquals(metadata, Payload(read, "meta.xml")));
        }
    }

    [Fact]
    public void TargetedEditsRetainLinksFieldsNestedListsAndOpaqueInlineContent() {
        var document = OdgDocument.Create(); var shape = document.AddPage().Shapes.AddTextBox(OdfRect.FromCentimeters(1, 1, 15, 10), string.Empty);
        shape.Element.Element(OdfNamespaces.Draw + "text-box")!.Elements().Remove();
        var paragraph = shape.AddParagraph("Before ");
        var run = paragraph.AddRun("old"); run.Bold = true;
        var link = paragraph.AddHyperlink("", "../manual.html#anchor"); link.TargetFrameName = "_blank";
        link.AddRun("Link").Italic = true;
        paragraph.AddText(" "); var field = paragraph.AddField(OdfTextFieldKind.Date, "2026-10-05"); field.IsFixed = true;
        var native = paragraph.Element;
        var bookmark = new XElement(OdfNamespaces.Text + "bookmark", new XAttribute(OdfNamespaces.Text + "name", "anchor"));
        var annotation = new XElement(OdfNamespaces.Office + "annotation", new XElement(OdfNamespaces.Text + "p", new XElement(OdfNamespaces.Text + "span", "Hidden note")));
        native.Add(bookmark, annotation); document.MarkPartDirty("content.xml");
        var list = shape.AddList(true); var item = list.AddItem("First"); item.AddParagraph("Second paragraph");
        var nested = new XElement(OdfNamespaces.Text + "list", new XElement(OdfNamespaces.Text + "list-item", new XElement(OdfNamespaces.Text + "p", "Nested")));
        shape.Element.Element(OdfNamespaces.Draw + "text-box")!.Elements(OdfNamespaces.Text + "list").Single().Element(OdfNamespaces.Text + "list-item")!.Add(nested);
        run.Text = "new  words"; link.Href = "https://example.invalid/help"; field.DisplayText = "Fixed\tdate";
        Assert.Equal(2, paragraph.Runs.Count); Assert.Single(paragraph.Fields);
        Assert.Equal(new[] { "Before ", "new  words", "Link", " ", "Fixed\tdate", "", "" }, paragraph.InlineNodes.Select(n => n.Text));
        Assert.Equal(OdfTextNodeKind.Field, paragraph.InlineNodes[4].Kind);
        Assert.Equal("Link", paragraph.InlineNodes[2].Children.Single().Text);
        Assert.True(item.Lists.Single().Items.Single().Paragraphs.Single().Text == "Nested");
        foreach (var read in RoundTrips(document)) {
            var text = read.Pages[0].Shapes[0]; var p = text.Paragraphs[0];
            Assert.Equal("Before new  wordsLink Fixed\tdate\nFirst\nSecond paragraph\nNested", text.Text);
            Assert.Equal("https://example.invalid/help", p.Hyperlinks.Single().Href);
            Assert.Equal("_blank", p.Hyperlinks.Single().TargetFrameName);
            Assert.True(p.Runs[0].Bold); Assert.True(p.Runs[1].Italic);
            Assert.True(p.Fields.Single().IsFixed); Assert.False(p.Fields.Single().ToXml().HasElements);
            Assert.True(XNode.DeepEquals(bookmark, p.ToXml().Element(OdfNamespaces.Text + "bookmark")));
            Assert.True(XNode.DeepEquals(annotation, p.ToXml().Element(OdfNamespaces.Office + "annotation")));
            Assert.True(text.Lists.Single().IsOrdered);
            Assert.Contains(read.Pages[0].ToDrawing().Report.Mappings, m => m.Feature.EndsWith(":text", StringComparison.Ordinal) && m.Status == OdfConversionMappingStatus.Approximated);
        }
    }

    [Fact]
    public void EffectiveFormattingResolvesNestedParagraphAndShapeDefaultsAndEditsCopyOnWrite() {
        var document = OdgDocument.Create(); var shape = document.AddPage().Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 8, 4));
        shape.FontFamily = "Liberation Sans"; shape.FontSize = OdfLength.Points(18);
        var paragraphStyle = document.Styles.CreateNamed("Paragraph", OdfStyleFamily.Paragraph); paragraphStyle.Color = OdfColor.Parse("#123456"); paragraphStyle.Bold = true;
        var spanStyle = document.Styles.CreateAutomatic(OdfStyleFamily.Text); spanStyle.Italic = true;
        var paragraph = shape.AddParagraph(); paragraph.StyleName = paragraphStyle.Name;
        var parent = paragraph.AddRun(); parent.StyleName = spanStyle.Name;
        var nested = parent.AddRun("Target"); var sibling = paragraph.AddRun("Sibling"); sibling.StyleName = spanStyle.Name;
        Assert.Equal("Liberation Sans", nested.FontFamily); Assert.Equal(18, nested.FontSize?.ToPoints());
        Assert.Equal(paragraphStyle.Color, nested.Color); Assert.True(nested.Bold); Assert.True(nested.Italic);
        sibling.Bold = false;
        Assert.True(parent.Bold); Assert.False(sibling.Bold); Assert.NotEqual(spanStyle.Name, sibling.StyleName); Assert.Null(spanStyle.Bold);
        paragraph.TextAlign = "center"; paragraph.MarginLeft = OdfLength.Centimeters(.25);
        Assert.Null(paragraphStyle.TextAlign);
        var link = parent.AddHyperlink("Highlight", "#anchor"); parent.BackgroundColor = OdfColor.Parse("#FFCC00");
        link.StyleName = document.Styles.CreateNamed("Transparent", OdfStyleFamily.Text).Name;
        document.Styles.Find(OdfStyleFamily.Text, link.StyleName!)!.SetProperty(OdfNamespaces.Style + "text-properties", OdfNamespaces.Fo + "background-color", "transparent");
        Assert.Null(link.BackgroundColor);
        foreach (var read in RoundTrips(document)) {
            var p = read.Pages[0].Shapes[0].Paragraphs.Single();
            Assert.Equal("center", p.TextAlign); Assert.Equal(.25, p.MarginLeft?.ToCentimeters()); Assert.False(p.Runs.Last().Bold);
            Assert.True(p.Runs[1].Italic); Assert.Null(p.Hyperlinks.Single().BackgroundColor);
        }
    }

    [Fact]
    public void InvalidEditsDoNotMutateNativeTextAndCaseTransformationKeepsStoriesSeparate() {
        var document = OdgDocument.Create(); var page = document.AddPage(); var shape = page.Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 8, 4));
        var paragraph = shape.AddParagraph("Abc "); var link = paragraph.AddHyperlink("Link", "#a"); var nested = link.AddRun("Nested");
        var count = paragraph.AddField(OdfTextFieldKind.PageCount, "1");
        string before = document.GetXml("content.xml").ToString();
        Assert.Throws<NotSupportedException>(() => link.AddHyperlink("Bad", "#b"));
        Assert.Throws<NotSupportedException>(() => nested.AddHyperlink("Bad", "#b"));
        Assert.Throws<ArgumentException>(() => link.Href = " "); Assert.Throws<ArgumentOutOfRangeException>(() => paragraph.AddField((OdfTextFieldKind)99));
        Assert.Throws<NotSupportedException>(() => count.IsFixed = true);
        Assert.Equal(before, document.GetXml("content.xml").ToString());
        paragraph.Element.Add(new XElement(OdfNamespaces.Office + "annotation", new XElement(OdfNamespaces.Text + "p", "Hidden")));
        paragraph.TransformTextCase(OfficeTextCase.Uppercase);
        Assert.Equal("ABC LINKNESTED1", paragraph.Text); Assert.Equal("Hidden", paragraph.ToXml().Element(OdfNamespaces.Office + "annotation")!.Value);
        Assert.Equal("#a", link.Href); Assert.Single(paragraph.Fields);
        var group = page.Shapes.AddGroup(); Assert.Throws<NotSupportedException>(() => group.AddParagraph("Bad")); Assert.Throws<NotSupportedException>(() => group.AddList());
    }

    [Fact]
    public void ShapeTextAndInlineViewsShareAnAggregateDecodedCharacterLimit() {
        var document = OdgDocument.Create(); var shape = document.AddPage().Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 8, 4));
        for (int i = 0; i < 2; i++) shape.AddParagraph().Element.Add(new XElement(OdfNamespaces.Text + "s", new XAttribute(OdfNamespaces.Text + "c", OdfTextCodec.MaximumDecodedCharacters / 2)));
        Assert.Throws<InvalidDataException>(() => shape.Text);
        var paragraph = shape.Paragraphs[0]; paragraph.AddRun().Element.Add(new XElement(OdfNamespaces.Text + "s", new XAttribute(OdfNamespaces.Text + "c", OdfTextCodec.MaximumDecodedCharacters / 2 + 1)));
        Assert.Throws<InvalidDataException>(() => paragraph.InlineNodes);
    }

    [Theory]
    [InlineData(OdfTextFieldKind.PageNumber)]
    [InlineData(OdfTextFieldKind.PageCount)]
    [InlineData(OdfTextFieldKind.Date)]
    [InlineData(OdfTextFieldKind.Time)]
    public void BasicFieldCacheEditPreservesNativeAttributesAndTextOnlyContent(OdfTextFieldKind kind) {
        var document = OdgDocument.Create(); var shape = document.AddPage().Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 8, 4));
        var paragraph = shape.AddParagraph(); var field = paragraph.AddField(kind, "cached");
        XElement element = paragraph.Element.Elements().Single();
        if (kind == OdfTextFieldKind.Date) element.SetAttributeValue(OdfNamespaces.Text + "date-value", "2026-10-05");
        if (kind == OdfTextFieldKind.Time) element.SetAttributeValue(OdfNamespaces.Text + "time-value", "PT12H30M");
        var attributes = element.Attributes().Select(a => (a.Name, a.Value)).ToArray();
        field.DisplayText = "  updated\tvalue\n";
        foreach (var read in RoundTrips(document)) {
            var imported = read.Pages[0].Shapes[0].Paragraphs.Single().Fields.Single();
            Assert.Equal(kind, imported.Kind); Assert.Equal(field.DisplayText, imported.DisplayText);
            Assert.False(imported.ToXml().HasElements); Assert.Equal(attributes, imported.ToXml().Attributes().Select(a => (a.Name, a.Value)));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AppendingAndReplacingCustomShapeTextKeepsEnhancedGeometryLast(bool replace) {
        var document = OdgDocument.LoadFlatXml(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Drawing", "libreoffice-transparent-text.fodg"));
        var shape = document.Pages[0].Shapes[0]; var geometry = new XElement(shape.ToXml().Element(OdfNamespaces.Draw + "enhanced-geometry")!);
        if (replace) shape.Text = "One\nTwo";
        else { shape.AddParagraph("Two"); shape.AddList(true).AddItem("Three"); }
        foreach (var read in RoundTrips(document)) {
            var xml = read.Pages[0].Shapes[0].ToXml(); Assert.Equal(OdfNamespaces.Draw + "enhanced-geometry", xml.Elements().Last().Name);
            Assert.True(XNode.DeepEquals(geometry, xml.Elements().Last())); Assert.Equal(replace ? 2 : 3, read.Pages[0].Shapes[0].Paragraphs.Count);
        }
    }

    [Fact]
    public void ParagraphStyleEditsDoNotChangeOtherParagraphsUsingShapeInheritance() {
        var document = OdgDocument.Create(); var shape = document.AddPage().Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 8, 4));
        var shared = document.Styles.CreateAutomatic(OdfStyleFamily.Paragraph); shared.Bold = true; shared.FontSize = OdfLength.Points(18);
        shape.Element.SetAttributeValue(OdfNamespaces.Draw + "text-style-name", shared.Name);
        var explicitParagraph = shape.AddParagraph("Explicit"); explicitParagraph.StyleName = shared.Name;
        var inheritedParagraph = shape.AddParagraph("Inherited"); Assert.True(inheritedParagraph.Bold);
        explicitParagraph.Bold = false; explicitParagraph.FontSize = OdfLength.Points(22);
        Assert.True(shared.Bold); Assert.True(inheritedParagraph.Bold); Assert.Equal(18, inheritedParagraph.FontSize?.ToPoints());
        foreach (var read in RoundTrips(document)) {
            var paragraphs = read.Pages[0].Shapes[0].Paragraphs; Assert.False(paragraphs[0].Bold); Assert.True(paragraphs[1].Bold);
            Assert.Equal(22, paragraphs[0].FontSize?.ToPoints()); Assert.Equal(18, paragraphs[1].FontSize?.ToPoints());
        }
    }

    [Fact]
    public void ExplicitGraphicPropertiesPrecedeFinalParagraphDefaults() {
        var document = OdgDocument.Create(); var shape = document.AddPage().Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 8, 4));
        document.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "styles")!.Add(new XElement(OdfNamespaces.Style + "default-style",
            new XAttribute(OdfNamespaces.Style + "family", "paragraph"), new XElement(OdfNamespaces.Style + "text-properties", new XAttribute(OdfNamespaces.Fo + "font-size", "12pt"))));
        shape.FontSize = OdfLength.Points(28); var paragraph = shape.AddParagraph(); var run = paragraph.AddRun("Inherited");
        Assert.Equal(28, paragraph.FontSize?.ToPoints()); Assert.Equal(28, run.FontSize?.ToPoints());
        paragraph.FontSize = OdfLength.Points(32); Assert.Equal(32, run.FontSize?.ToPoints());
        shape.FontSize = null; paragraph.FontSize = null; Assert.Equal(12, run.FontSize?.ToPoints());
        document.GetXml("styles.xml").Root!.Element(OdfNamespaces.Office + "styles")!.Add(new XElement(OdfNamespaces.Style + "default-style",
            new XAttribute(OdfNamespaces.Style + "family", "text"), new XElement(OdfNamespaces.Style + "text-properties", new XAttribute(OdfNamespaces.Fo + "font-size", "10pt"))));
        run.Bold = true; Assert.Equal(10, run.FontSize?.ToPoints());
    }

    private static OdgDocument[] RoundTrips(OdgDocument document) {
        using var flat = new MemoryStream(); document.SaveFlatXml(flat); flat.Position = 0;
        return new[] { OdgDocument.Load(new MemoryStream(document.ToBytes())), OdgDocument.LoadFlatXml(flat) };
    }
    private static XElement Payload(OdgDocument document, string part) => new XElement("payload", document.GetXml(part).Root!.Elements());
    private static XElement StyleProperties(XElement element) {
        var copy = new XElement(element);
        // Flat save remaps font-face IDs when package parts contain duplicate declarations.
        copy.DescendantsAndSelf().Attributes().Where(a => a.IsNamespaceDeclaration || a.Name == OdfNamespaces.Style + "font-name" || a.Name == OdfNamespaces.Style + "font-name-asian" || a.Name == OdfNamespaces.Style + "font-name-complex").Remove();
        return copy;
    }
}
