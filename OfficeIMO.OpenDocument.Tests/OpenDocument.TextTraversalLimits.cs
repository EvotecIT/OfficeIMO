using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentTextTraversalLimitsTests {
    [Fact]
    public void NativeDrawParagraphsKeepListHeadersAndNestedItemsInStoryOrder() {
        var document = OdgDocument.Create();
        var shape = CreateTextBox(document);
        XElement root = shape.Element.Element(OdfNamespaces.Draw + "text-box")!;
        root.Add(Paragraph("Before"),
            new XElement(OdfNamespaces.Text + "list",
                new XElement(OdfNamespaces.Text + "list-header", new XElement(OdfNamespaces.Text + "h", "Header")),
                new XElement(OdfNamespaces.Text + "list-item", Paragraph("First"),
                    new XElement(OdfNamespaces.Text + "list",
                        new XElement(OdfNamespaces.Text + "list-item", Paragraph("Nested"))),
                    Paragraph("Last in item"))),
            new XElement(OdfNamespaces.Office + "annotation", NestedList(65)),
            new XElement(OdfNamespaces.Text + "note", Paragraph("Footnote")),
            new XElement(OdfNamespaces.Draw + "object", Paragraph("Embedded")),
            new XElement(OdfNamespaces.Table + "table", Paragraph("Other story")),
            Paragraph("After"));

        string[] expected = { "Before", "Header", "First", "Nested", "Last in item", "After" };
        Assert.Equal(expected, shape.Paragraphs.Select(paragraph => paragraph.Text));
        Assert.Equal(string.Join("\n", expected), shape.Text);
    }

    [Fact]
    public void NativeInlineCollectionsKeepVisibleOrderAndDoNotEnterOtherStories() {
        var document = OdgDocument.Create();
        var shape = CreateTextBox(document);
        var paragraph = shape.AddParagraph();
        paragraph.Element.Add(
            Span("outer", new XElement(OdfNamespaces.Text + "a", new XAttribute(OdfNamespaces.XLink + "href", "#first"),
                Span("nested", new XElement(OdfNamespaces.Text + "date", "Date")))),
            Span("sibling", "Sibling"),
            new XElement(OdfNamespaces.Text + "a", new XAttribute(OdfNamespaces.XLink + "href", "#second"),
                new XElement(OdfNamespaces.Text + "page-number", "2")),
            new XElement(OdfNamespaces.Office + "annotation", NestedSpans(129)),
            new XElement(OdfNamespaces.Text + "note", Span("footnote", new XElement(OdfNamespaces.Text + "time", "Time"))),
            new XElement(OdfNamespaces.Draw + "object", Span("embedded", "Embedded")));

        Assert.Equal(new[] { "outer", "nested", "sibling" }, paragraph.Runs.Select(run => run.StyleName));
        Assert.Equal(new[] { "#first", "#second" }, paragraph.Hyperlinks.Select(link => link.Href));
        Assert.Equal(new[] { OdfTextFieldKind.Date, OdfTextFieldKind.PageNumber }, paragraph.Fields.Select(field => field.Kind));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Exactly128VisibleContainersRemainAccessible(bool inline) {
        var document = OdgDocument.Create();
        var shape = CreateTextBox(document);
        if (inline) {
            var paragraph = shape.AddParagraph();
            paragraph.Element.Add(NestedSpans(128));
            Assert.Equal(128, paragraph.Runs.Count);
            Assert.Equal("Leaf", Assert.Single(paragraph.Fields).DisplayText);
        } else {
            shape.Element.Element(OdfNamespaces.Draw + "text-box")!.Add(NestedList(64));
            Assert.Equal("Leaf", Assert.Single(shape.Paragraphs).Text);
            Assert.Equal("Leaf", shape.Text);
        }
    }

    [Fact]
    public void BoundedTextDecodingAndCaseTransformsPreserveWhitespaceAndOtherStories() {
        var document = OdgDocument.Create();
        var shape = CreateTextBox(document);
        var paragraph = shape.AddParagraph();
        XNode nested = new XText("mIxed");
        for (int index = 0; index < 128; index++) nested = new XElement(OdfNamespaces.Text + "span", nested);
        var annotation = new XElement(OdfNamespaces.Office + "annotation", NestedSpans(129));
        paragraph.Element.Add(nested,
            new XElement(OdfNamespaces.Text + "s", new XAttribute(OdfNamespaces.Text + "c", 2)),
            new XElement(OdfNamespaces.Text + "tab"), new XElement(OdfNamespaces.Text + "span", "Tail"),
            new XElement(OdfNamespaces.Text + "line-break"), annotation);
        string originalAnnotation = annotation.ToString();

        Assert.Equal("mIxed  \tTail\n", paragraph.Text);
        Assert.Equal(paragraph.Text, shape.Text);
        paragraph.TransformTextCase(OfficeTextCase.Lowercase);
        Assert.Equal("mixed  \ttail\n", paragraph.Text);
        Assert.Equal(originalAnnotation, annotation.ToString());
        Assert.Equal(129, paragraph.Runs.Count);
    }

    [Theory]
    [InlineData(65)]
    [InlineData(2048)]
    public void ParsedDeepListsAreRejectedByNativeViewsAndReportedByProjection(int listLevels) {
        var document = OdgDocument.Create();
        var shape = CreateTextBox(document);
        shape.Element.Element(OdfNamespaces.Draw + "text-box")!.Add(NestedList(listLevels));
        document.MarkPartDirty("content.xml");
        var read = OdgDocument.Load(new MemoryStream(document.ToBytes()),
            new OdfLoadOptions { MaxXmlDepth = listLevels * 2 + 20 });
        var reopened = read.Pages[0].Shapes[0];

        Assert.Contains("128-container", Assert.Throws<NotSupportedException>(() => _ = reopened.Paragraphs).Message);
        Assert.Contains("128-container", Assert.Throws<NotSupportedException>(() => _ = reopened.Text).Message);
        var result = read.Pages[0].ToDrawing();
        Assert.Empty(result.Value.Elements.OfType<OfficeDrawingRichText>());
        Assert.Contains(result.Report.Mappings, mapping => mapping.Status == OdfConversionMappingStatus.Skipped &&
            mapping.Message?.Contains("128-container") == true);
    }

    [Theory]
    [InlineData(129)]
    [InlineData(2048)]
    public void ParsedDeepSpansAreRejectedByAllNativeInlineCollections(int spanLevels) {
        var document = OdgDocument.Create();
        var shape = CreateTextBox(document);
        shape.AddParagraph("early lowercase ").Element.Add(NestedSpans(spanLevels));
        document.MarkPartDirty("content.xml");
        var read = OdgDocument.Load(new MemoryStream(document.ToBytes()),
            new OdfLoadOptions { MaxXmlDepth = spanLevels + 20 });
        var paragraph = Assert.Single(read.Pages[0].Shapes[0].Paragraphs);

        Assert.Contains("128-container", Assert.Throws<NotSupportedException>(() => _ = paragraph.Runs).Message);
        Assert.Contains("128-container", Assert.Throws<NotSupportedException>(() => _ = paragraph.Hyperlinks).Message);
        Assert.Contains("128-container", Assert.Throws<NotSupportedException>(() => _ = paragraph.Fields).Message);
        Assert.Contains("128-container", Assert.Throws<NotSupportedException>(() => _ = paragraph.Text).Message);
        Assert.Contains("128-container", Assert.Throws<NotSupportedException>(() => _ = read.Pages[0].Shapes[0].Text).Message);
        string before = read.GetXml("content.xml").ToString();
        Assert.Throws<NotSupportedException>(() => paragraph.TransformTextCase(OfficeTextCase.Uppercase));
        Assert.Equal(before, read.GetXml("content.xml").ToString());
        var result = read.Pages[0].ToDrawing();
        Assert.Empty(result.Value.Elements.OfType<OfficeDrawingRichText>());
        Assert.Contains(result.Report.Mappings, mapping => mapping.Status == OdfConversionMappingStatus.Skipped &&
            mapping.Message?.Contains("128-container") == true);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void VeryWideNativeTextCollectionsStopAtTheirElementBudget(bool inline) {
        var document = OdgDocument.Create();
        var shape = CreateTextBox(document);
        var paragraph = inline ? shape.AddParagraph() : null;
        XElement root = paragraph?.Element ?? shape.Element.Element(OdfNamespaces.Draw + "text-box")!;
        XName name = OdfNamespaces.Text + (inline ? "span" : "p");
        for (int index = 0; index <= 100_000; index++) root.Add(new XElement(name));

        NotSupportedException exception = inline
            ? Assert.Throws<NotSupportedException>(() => _ = paragraph!.Runs)
            : Assert.Throws<NotSupportedException>(() => _ = shape.Paragraphs);
        Assert.Contains("100000-element", exception.Message);
        if (paragraph != null) {
            Assert.Contains("100000-node", Assert.Throws<NotSupportedException>(() => _ = paragraph.Text).Message);
            string before = paragraph.ToXml().ToString();
            Assert.Throws<NotSupportedException>(() => paragraph.TransformTextCase(OfficeTextCase.Uppercase));
            Assert.Equal(before, paragraph.ToXml().ToString());
        }
    }

    private static OdgShape CreateTextBox(OdgDocument document) {
        var shape = document.AddPage().Shapes.AddTextBox(OdfRect.FromCentimeters(1, 1, 12, 8), string.Empty);
        shape.Element.Element(OdfNamespaces.Draw + "text-box")!.RemoveNodes();
        return shape;
    }

    private static XElement Paragraph(string text) => new(OdfNamespaces.Text + "p", text);

    private static XElement Span(string name, object content) => new(OdfNamespaces.Text + "span",
        new XAttribute(OdfNamespaces.Text + "style-name", name), content);

    private static XElement NestedList(int levels) {
        XElement content = Paragraph("Leaf");
        for (int index = 0; index < levels; index++)
            content = new XElement(OdfNamespaces.Text + "list", new XElement(OdfNamespaces.Text + "list-item", content));
        return content;
    }

    private static XElement NestedSpans(int levels) {
        XElement content = new(OdfNamespaces.Text + "date", "Leaf");
        for (int index = 0; index < levels; index++) content = new XElement(OdfNamespaces.Text + "span", content);
        return content;
    }
}
