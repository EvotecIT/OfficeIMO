using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;
using static OfficeIMO.Xps.Tests.XpsLogicalStructureTests;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsPdfLogicalStructureTests {
    [Theory]
    [InlineData(XpsFormat.Xps)]
    [InlineData(XpsFormat.OpenXps)]
    public void NativeReadingOrderChangesTagsAndSearchTextWithoutChangingPaint(XpsFormat format) {
        var doc = Create(format, new[] { "Alpha", "Beta" }); var ns = doc.StructureNamespace;
        doc.Pages[0].ReplaceStoryFragmentsMarkup(Fragments(doc, null, null, Paragraph(ns, "Beta", "Alpha")));
        byte[] tagged = doc.ToPdf(), plain = doc.ToPdf(preserveLogicalStructure: false);
        var read = PdfReadDocument.Open(tagged);
        Assert.Equal("BetaAlpha", TreeText(read));
        Assert.Contains(read.TaggedContent!.StructureElements, e => e.StructureType == "P");
        Assert.Null(PdfReadDocument.Open(plain).TaggedContent);
        var expected = OfficeDrawingRasterRenderer.Render(PdfDocument.Load(plain).Render.Drawing(1));
        var actual = OfficeDrawingRasterRenderer.Render(PdfDocument.Load(tagged).Render.Drawing(1));
        Assert.Equal(expected.Width, actual.Width); Assert.Equal(expected.Height, actual.Height);
        for (int y = 0; y < expected.Height; y++) for (int x = 0; x < expected.Width; x++)
            Assert.Equal(expected.GetPixel(x, y), actual.GetPixel(x, y));
    }

    [Fact]
    public void CrossPageParagraphFollowsDeclaredStoryOrderAndHasOneSharedParent() {
        var doc = Create(XpsFormat.OpenXps, new[] { "Alpha" }, new[] { "Beta" }); var ns = doc.StructureNamespace;
        doc.Pages[0].ReplaceStoryFragmentsMarkup(Fragments(doc, "body", "one", Paragraph(ns, "Alpha")));
        doc.Pages[1].ReplaceStoryFragmentsMarkup(Fragments(doc, "body", "two", Paragraph(ns, "Beta")));
        doc.Documents[0].ReplaceDocumentStructureMarkup(Story(doc, "body", (2, "two"), (1, "one")));
        var read = PdfReadDocument.Open(doc.ToPdf());
        Assert.Equal("BetaAlpha", TreeText(read));
        var paragraph = Assert.Single(read.TaggedContent!.StructureElements, e => e.StructureType == "P");
        Assert.Null(paragraph.PageObjectNumber);
        Assert.Equal(2, paragraph.ChildElementObjectNumbers.SelectMany(id =>
            read.TaggedContent.StructureElements.Single(e => e.ObjectNumber == id).MarkedContentReferences).Select(r => r.PageObjectNumber).Distinct().Count());
    }

    [Fact]
    public void IndependentStoriesRetainTheirDeclaredRootOrder() {
        var doc = Create(XpsFormat.Xps, new[] { "Alpha" }, new[] { "Beta" }); var ns = doc.StructureNamespace;
        doc.Pages[0].ReplaceStoryFragmentsMarkup(Fragments(doc, "alpha", null, Paragraph(ns, "Alpha")));
        doc.Pages[1].ReplaceStoryFragmentsMarkup(Fragments(doc, "beta", null, Paragraph(ns, "Beta")));
        var structure = Story(doc, "beta", (2, null));
        structure.Add(Story(doc, "alpha", (1, null)).Elements()); doc.Documents[0].ReplaceDocumentStructureMarkup(structure);
        Assert.Equal("BetaAlpha", TreeText(PdfReadDocument.Open(doc.ToPdf())));
    }

    [Fact]
    public void ListsAndTablesRetainLabelsAndCellSpans() {
        var doc = Create(XpsFormat.OpenXps, new[] { "Marker", "Body", "Cell" }); var ns = doc.StructureNamespace;
        doc.Pages[0].ReplaceStoryFragmentsMarkup(Fragments(doc, null, null,
            new XElement(ns + "ListStructure", new XElement(ns + "ListItemStructure", new XAttribute("Marker", "Marker"), Paragraph(ns, "Body"))),
            new XElement(ns + "TableStructure", new XElement(ns + "TableRowGroupStructure", new XElement(ns + "TableRowStructure",
                new XElement(ns + "TableCellStructure", new XAttribute("RowSpan", "2"), new XAttribute("ColumnSpan", "3"), Paragraph(ns, "Cell")))))));
        var read = PdfReadDocument.Open(doc.ToPdf());
        Assert.Equal("MarkerBodyCell", TreeText(read));
        foreach (string role in new[] { "L", "LI", "Lbl", "LBody", "Table", "TBody", "TR", "TD" })
            Assert.Contains(read.TaggedContent!.StructureElements, e => e.StructureType == role);
        var cell = Assert.Single(read.TaggedContent!.StructureElements, e => e.StructureType == "TD");
        var raw = read.RawStructure().Objects.Single(o => o.ObjectNumber == cell.ObjectNumber).Value;
        Assert.Equal(2D, raw.Entries["A"].Entries["RowSpan"].Number);
        Assert.Equal(3D, raw.Entries["A"].Entries["ColSpan"].Number);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void GraphicFiguresOwnTheirActualPaintIncludingEffectGroups(bool effect) {
        var doc = XpsDocument.Create(); var page = doc.AddPage(100, 100); page.AddPath("M10,10H50V40H10Z", "#FF336699");
        var xml = page.GetMarkup(); var path = xml.Elements().Single(); path.SetAttributeValue("Name", "Graphic");
        if (effect) { path.Remove(); xml.Add(new XElement(xml.Name.Namespace + "Canvas", new XAttribute("Opacity", ".5"), new XAttribute("Name", "Group"), path)); }
        page.ReplaceMarkup(xml); var ns = doc.StructureNamespace;
        page.ReplaceStoryFragmentsMarkup(Fragments(doc, null, null, new XElement(ns + "FigureStructure",
            new XElement(ns + "NamedElement", new XAttribute("NameReference", effect ? "Group" : "Graphic")))));
        var read = PdfReadDocument.Open(doc.ToPdf());
        var figure = Assert.Single(read.TaggedContent!.StructureElements, e => e.StructureType == "Figure");
        Assert.Null(figure.AlternateText);
        Assert.NotEmpty(figure.ChildElementObjectNumbers);
        Assert.True(figure.ChildElementObjectNumbers.Sum(id => read.TaggedContent.StructureElements.Single(e => e.ObjectNumber == id).MarkedContentReferenceCount) > 0);
    }

    [Fact]
    public void UnsupportedOrRepeatedSemanticsFailExplicitlyButPaintExportRemainsAvailable() {
        var doc = Create(XpsFormat.OpenXps, new[] { "Alpha" }); var ns = doc.StructureNamespace;
        doc.Pages[0].ReplaceStoryFragmentsMarkup(Fragments(doc, null, null, Paragraph(ns, "Alpha", "Alpha")));
        Assert.Throws<NotSupportedException>(() => doc.ToPdf()); Assert.NotEmpty(doc.ToPdf(preserveLogicalStructure: false));
        doc.Pages[0].ReplaceStoryFragmentsMarkup(Fragments(doc, null, null, new XElement(XName.Get("Semantic", "urn:extension"))));
        Assert.Throws<NotSupportedException>(() => doc.ToPdf()); Assert.NotEmpty(doc.ToPdf(preserveLogicalStructure: false));
    }

    [Fact]
    public void UnpaintedGlyphsKeepTheFollowingRunsNativeIdentityAndUnreferencedText() {
        var doc = Create(XpsFormat.Xps, new[] { "Zero", "Alpha", "Extra" }); var page = doc.Pages[0]; var xml = page.GetMarkup();
        xml.Elements().First().SetAttributeValue("FontRenderingEmSize", "0"); page.ReplaceMarkup(xml);
        page.ReplaceStoryFragmentsMarkup(Fragments(doc, null, null, Paragraph(doc.StructureNamespace, "Alpha")));
        var read = PdfReadDocument.Open(doc.ToPdf());
        Assert.Equal("AlphaExtra", TreeText(read));
        Assert.DoesNotContain("Zero", read.ExtractText());
        Assert.Contains(read.TaggedContent!.StructureElements, e => e.StructureType == "Div");
    }

    [Fact]
    public void EmptyTableCellsRetainTheirPositionWithoutInventedContent() {
        var doc = Create(XpsFormat.OpenXps, new[] { "Cell" }); var ns = doc.StructureNamespace;
        doc.Pages[0].ReplaceStoryFragmentsMarkup(Fragments(doc, null, null, new XElement(ns + "TableStructure",
            new XElement(ns + "TableRowGroupStructure", new XElement(ns + "TableRowStructure",
                new XElement(ns + "TableCellStructure"), new XElement(ns + "TableCellStructure", Paragraph(ns, "Cell")))))));
        var read = PdfReadDocument.Open(doc.ToPdf()); var tags = read.TaggedContent!;
        var row = Assert.Single(tags.StructureElements, e => e.StructureType == "TR");
        Assert.Equal(2, row.ChildElementObjectNumbers.Count);
        var first = tags.StructureElements.Single(e => e.ObjectNumber == row.ChildElementObjectNumbers[0]);
        Assert.Equal("TD", first.StructureType); Assert.Empty(first.ChildElementObjectNumbers); Assert.Empty(first.MarkedContentReferences);
        Assert.Equal("Cell", TreeText(read));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NestedGraphicsCannotHaveConflictingSemanticOwners(bool effect) {
        var doc = XpsDocument.Create(); var page = doc.AddPage(100, 100).AddPath("M10,10H50V40H10Z", "#FF336699");
        var xml = page.GetMarkup(); var path = xml.Elements().Single(); path.SetAttributeValue("Name", "Child"); path.Remove();
        xml.Add(new XElement(xml.Name.Namespace + "Canvas", new XAttribute("Name", "Parent"), effect ? new XAttribute("Opacity", ".5") : null, path));
        page.ReplaceMarkup(xml); var ns = doc.StructureNamespace;
        page.ReplaceStoryFragmentsMarkup(Fragments(doc, null, null, new XElement(ns + "FigureStructure",
            new XElement(ns + "NamedElement", new XAttribute("NameReference", "Parent")),
            new XElement(ns + "NamedElement", new XAttribute("NameReference", "Child")))));
        Assert.Throws<NotSupportedException>(() => doc.ToPdf()); Assert.NotEmpty(doc.ToPdf(preserveLogicalStructure: false));
    }

    [Fact]
    public void HeaderTextRemainsSearchableAsAnArtifactOutsideTheBodyTree() {
        var doc = Create(XpsFormat.Xps, new[] { "Header", "Body" }); var ns = doc.StructureNamespace;
        var markup = Fragments(doc, null, null, Paragraph(ns, "Header"));
        markup.Elements().Single().SetAttributeValue("FragmentType", "Header");
        markup.Add(Fragments(doc, null, null, Paragraph(ns, "Body")).Elements());
        doc.Pages[0].ReplaceStoryFragmentsMarkup(markup);
        var read = PdfReadDocument.Open(doc.ToPdf(), new PdfLoadOptions { IncludeArtifactText = true });
        Assert.Equal("Body", TreeText(read));
        var spans = read.Pages[0].GetTextSpans();
        Assert.Equal("Header", string.Concat(spans.Where(s => s.IsArtifactContent).Select(s => s.Text)));
        Assert.Equal("Body", string.Concat(spans.Where(s => !s.IsArtifactContent).Select(s => s.Text)));
    }

    [Fact]
    public void NonEmittingReferencesRetainAuthoredContainersAndCellSpans() {
        var doc = Create(XpsFormat.OpenXps, new[] { "Zero", "Visible" }); var page = doc.Pages[0]; var ns = doc.StructureNamespace;
        var xml = page.GetMarkup(); xml.Elements().First().SetAttributeValue("FontRenderingEmSize", "0"); page.ReplaceMarkup(xml);
        page.ReplaceStoryFragmentsMarkup(Fragments(doc, null, null, new XElement(ns + "TableStructure",
            new XElement(ns + "TableRowGroupStructure", new XElement(ns + "TableRowStructure",
                new XElement(ns + "TableCellStructure", new XAttribute("ColumnSpan", "2"), Paragraph(ns, "Zero")),
                new XElement(ns + "TableCellStructure", Paragraph(ns, "Visible")))))));
        var read = PdfReadDocument.Open(doc.ToPdf()); var tags = read.TaggedContent!;
        var row = Assert.Single(tags.StructureElements, e => e.StructureType == "TR");
        Assert.Equal(2, row.ChildElementObjectNumbers.Count);
        var first = tags.StructureElements.Single(e => e.ObjectNumber == row.ChildElementObjectNumbers[0]);
        var paragraph = tags.StructureElements.Single(e => e.ObjectNumber == Assert.Single(first.ChildElementObjectNumbers));
        Assert.Equal("P", paragraph.StructureType); Assert.Empty(paragraph.MarkedContentReferences);
        Assert.Equal(2D, read.RawStructure().Objects.Single(o => o.ObjectNumber == first.ObjectNumber).Value.Entries["A"].Entries["ColSpan"].Number);
        Assert.Equal("Visible", TreeText(read));
        // Existing compiled callers and two-parameter method groups keep their entrypoint.
        Func<XpsDocument, System.Threading.CancellationToken, byte[]> convert = XpsPdfExtensions.ToPdf;
        Assert.NotEmpty(convert(doc, default)); Assert.NotEmpty(doc.ToPdf(default));
    }

    [Theory]
    [InlineData("Header", "Content")]
    [InlineData("Header", "Footer")]
    public void ArtifactGraphicsParticipateInSemanticOwnershipValidation(string firstKind, string secondKind) {
        var doc = XpsDocument.Create(); var page = doc.AddPage(100, 100).AddPath("M10,10H50V40H10Z", "#FF336699");
        var xml = page.GetMarkup(); xml.Elements().Single().SetAttributeValue("Name", "Graphic"); page.ReplaceMarkup(xml); var ns = doc.StructureNamespace;
        XElement Fragment(string kind) => new(ns + "StoryFragment", new XAttribute("FragmentType", kind),
            new XElement(ns + "FigureStructure", new XElement(ns + "NamedElement", new XAttribute("NameReference", "Graphic"))));
        page.ReplaceStoryFragmentsMarkup(new XElement(ns + "StoryFragments", Fragment(firstKind), Fragment(secondKind)));
        Assert.Throws<NotSupportedException>(() => doc.ToPdf()); Assert.NotEmpty(doc.ToPdf(preserveLogicalStructure: false));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ArtifactCanvasAndBodyPathCannotOwnOverlappingPaint(bool effect) {
        var doc = XpsDocument.Create(); var page = doc.AddPage(100, 100).AddPath("M10,10H50V40H10Z", "#FF336699");
        var xml = page.GetMarkup(); var path = xml.Elements().Single(); path.SetAttributeValue("Name", "Child"); path.Remove();
        xml.Add(new XElement(xml.Name.Namespace + "Canvas", new XAttribute("Name", "Parent"), effect ? new XAttribute("Opacity", ".5") : null, path));
        page.ReplaceMarkup(xml); var ns = doc.StructureNamespace;
        XElement Fragment(string kind, string name) => new(ns + "StoryFragment", new XAttribute("FragmentType", kind),
            new XElement(ns + "FigureStructure", new XElement(ns + "NamedElement", new XAttribute("NameReference", name))));
        page.ReplaceStoryFragmentsMarkup(new XElement(ns + "StoryFragments", Fragment("Header", "Parent"), Fragment("Content", "Child")));
        Assert.Throws<NotSupportedException>(() => doc.ToPdf()); Assert.NotEmpty(doc.ToPdf(preserveLogicalStructure: false));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NonEmittingMarkersRetainTheirAuthoredLabelContainer(bool emptyCanvas) {
        var doc = Create(XpsFormat.OpenXps, new[] { "Marker", "Body" }); var page = doc.Pages[0]; var ns = doc.StructureNamespace;
        var xml = page.GetMarkup(); var marker = xml.Elements().First();
        if (emptyCanvas) marker.ReplaceWith(new XElement(xml.Name.Namespace + "Canvas", new XAttribute("Name", "Marker")));
        else marker.SetAttributeValue("FontRenderingEmSize", "0");
        page.ReplaceMarkup(xml);
        page.ReplaceStoryFragmentsMarkup(Fragments(doc, null, null, new XElement(ns + "ListStructure",
            new XElement(ns + "ListItemStructure", new XAttribute("Marker", "Marker"), Paragraph(ns, "Body")))));
        var read = PdfReadDocument.Open(doc.ToPdf()); var tags = read.TaggedContent!;
        var item = Assert.Single(tags.StructureElements, e => e.StructureType == "LI");
        Assert.Equal(new[] { "Lbl", "LBody" }, item.ChildElementObjectNumbers.Select(id => tags.StructureElements.Single(e => e.ObjectNumber == id).StructureType));
        var label = Assert.Single(tags.StructureElements, e => e.StructureType == "Lbl");
        Assert.Empty(label.ChildElementObjectNumbers); Assert.Empty(label.MarkedContentReferences);
        Assert.Equal("Body", TreeText(read));
    }

    internal static string TreeText(PdfReadDocument read) {
        var tagged = read.TaggedContent!; var elements = tagged.StructureElements.ToDictionary(e => e.ObjectNumber);
        var spans = read.Pages.ToDictionary(p => p.ObjectNumber, p => p.GetTextSpans());
        var raw = read.RawStructure().Objects.ToDictionary(o => o.ObjectNumber, o => o.Value);
        var pageStreams = read.Pages.ToDictionary(p => p.ObjectNumber, p => raw[p.ObjectNumber].Entries["Contents"].ReferenceObjectNumber);
        string Visit(int id) {
            var element = elements[id];
            return string.Concat(element.ChildElementObjectNumbers.Select(Visit)) +
                string.Concat(element.MarkedContentReferences.SelectMany(r => spans[r.PageObjectNumber!.Value]
                    .Where(s => s.MarkedContentId == r.MarkedContentId && s.ContentStreamObjectNumber == (r.ContentStreamObjectNumber ?? pageStreams[r.PageObjectNumber.Value]))).Select(s => s.Text));
        }
        return string.Concat(tagged.RootElementObjectNumbers.Select(Visit));
    }
}
