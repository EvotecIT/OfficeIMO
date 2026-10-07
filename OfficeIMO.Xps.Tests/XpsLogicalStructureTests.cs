using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsLogicalStructureTests {
    [Theory]
    [InlineData(XpsFormat.Xps)]
    [InlineData(XpsFormat.OpenXps)]
    public void NativeReferencesDetermineReadingOrderAndRemainDetached(XpsFormat format) {
        var doc = Create(format, new[] { "A", "B" }); var page = doc.Pages[0]; var ns = doc.StructureNamespace;
        page.ReplaceStoryFragmentsMarkup(Fragments(doc, null, null, Paragraph(ns, "B", "A")));
        var native = page.GetStoryFragmentsMarkup()!;
        native.Descendants(ns + "NamedElement").First().SetAttributeValue("NameReference", "A");
        var result = XpsDocument.Load(doc.Save()).ReadLogicalStructure();
        Assert.True(result.HasNativeStructure); Assert.True(result.IsComplete);
        Assert.Equal("BA", Assert.Single(Assert.Single(result.Stories).Blocks).Text);
        Assert.Equal("B", Assert.Single(result.Stories).Blocks[0].Children[0].Content!.Name);
        Assert.Equal(0, Assert.Single(result.Stories).Blocks[0].Children[0].Content!.PageIndex);
        Assert.Equal("A\nB", page.ExtractText());
    }

    [Fact]
    public void DeclaredStoryOrderCanDifferFromPageOrderAndBreaksStopContinuation() {
        var doc = Create(XpsFormat.OpenXps, new[] { "A" }, new[] { "B" }); var ns = doc.StructureNamespace;
        doc.Pages[0].ReplaceStoryFragmentsMarkup(Fragments(doc, "body", "one", Paragraph(ns, "A")));
        doc.Pages[1].ReplaceStoryFragmentsMarkup(Fragments(doc, "body", "two", Paragraph(ns, "B")));
        doc.Documents[0].ReplaceDocumentStructureMarkup(Story(doc, "body", (2, "two"), (1, "one")));
        var logical = doc.ReadLogicalStructure();
        Assert.True(logical.IsComplete);
        Assert.Equal("BA", Assert.Single(Assert.Single(logical.Stories).Blocks).Text);
        Assert.Equal(new[] { 1, 0 }, Assert.Single(logical.Stories).Fragments.Select(f => f.PageIndex));
        var markup = doc.Pages[0].GetStoryFragmentsMarkup()!;
        markup.Element(ns + "StoryFragment")!.AddFirst(new XElement(ns + "StoryBreak"));
        doc.Pages[0].ReplaceStoryFragmentsMarkup(markup);
        Assert.Equal(new[] { "B", "A" }, Assert.Single(doc.ReadLogicalStructure().Stories).Blocks.Select(b => b.Text));
    }

    [Fact]
    public void EcmaExample165MergesOnlyTheBoundaryRowAndAlignsItsCells() {
        // ECMA-388 example 16-5: two two-row fragments reconstruct three rows.
        // The empty cell in the trailing boundary row continues the leading cell.
        var doc = Create(XpsFormat.OpenXps, new[] { "Block2", "Block3", "Block4", "Block5", "Block6", "Block7" },
            new[] { "Block10", "Block11", "Block12", "Block13" }); var ns = doc.StructureNamespace;
        XElement Cell(params string[] names) => new(ns + "TableCellStructure", names.Length == 0 ? null : Paragraph(ns, names));
        XElement Row(params XElement[] cells) => new(ns + "TableRowStructure", cells);
        XElement Table(params XElement[] rows) => new(ns + "SectionStructure",
            new XElement(ns + "TableStructure", new XElement(ns + "TableRowGroupStructure", rows)));
        doc.Pages[0].ReplaceStoryFragmentsMarkup(Fragments(doc, "body", "first",
            Table(Row(Cell("Block2", "Block3"), Cell("Block4")), Row(Cell("Block5", "Block6"), Cell("Block7")))));
        doc.Pages[1].ReplaceStoryFragmentsMarkup(Fragments(doc, "body", "second",
            Table(Row(Cell(), Cell("Block10", "Block11")), Row(Cell("Block12"), Cell("Block13")))));
        doc.Documents[0].ReplaceDocumentStructureMarkup(Story(doc, "body", (1, "first"), (2, "second")));
        var result = XpsDocument.Load(doc.Save()).ReadLogicalStructure();
        Assert.True(result.IsComplete);
        var rows = Assert.Single(result.Stories).Blocks[0].Children[0].Children[0].Children;
        Assert.Equal(3, rows.Count);
        Assert.Equal(new[] { "Block2Block3", "Block4", "Block5Block6", "Block7Block10Block11", "Block12", "Block13" },
            rows.SelectMany(row => row.Children).Select(cell => cell.Text));
    }

    [Fact]
    public void NativeListsExposeMarkersAndCellsExposeSpansWithoutInventedText() {
        var doc = Create(XpsFormat.Xps, new[] { "Bullet", "Text", "Cell" }); var ns = doc.StructureNamespace;
        var list = new XElement(ns + "ListStructure", new XElement(ns + "ListItemStructure",
            new XAttribute("Marker", "Bullet"), Paragraph(ns, "Text")));
        var table = new XElement(ns + "TableStructure", new XElement(ns + "TableRowGroupStructure",
            new XElement(ns + "TableRowStructure", new XElement(ns + "TableCellStructure",
                new XAttribute("RowSpan", "2"), new XAttribute("ColumnSpan", "3"), Paragraph(ns, "Cell")))));
        doc.Pages[0].ReplaceStoryFragmentsMarkup(Fragments(doc, null, null, list, table));
        var blocks = Assert.Single(doc.ReadLogicalStructure().Stories).Blocks;
        Assert.Equal("Bullet", blocks[0].Children[0].Marker!.Text); Assert.Equal("Text", blocks[0].Text);
        var cell = blocks[1].Children[0].Children[0].Children[0];
        Assert.Equal(2, cell.RowSpan); Assert.Equal(3, cell.ColumnSpan); Assert.Equal("Cell", cell.Text);
    }

    [Fact]
    public void CanvasReferencesExcludeBrushResourceAndForeignNamespaceGlyphs() {
        var doc = Create(XpsFormat.Xps, new[] { "A", "B" }); var page = doc.Pages[0];
        var markup = page.GetMarkup(); var ns = markup.Name.Namespace; var glyphs = markup.Elements().ToArray();
        glyphs.Remove();
        markup.Add(new XElement(ns + "Canvas", new XAttribute("Name", "Group"), glyphs,
            new XElement(XName.Get("Glyphs", "urn:extension"), new XAttribute("Name", "Foreign"), new XAttribute("UnicodeString", "hidden")),
            new XElement(ns + "Canvas.Resources", new XElement(ns + "ResourceDictionary",
                new XElement(ns + "VisualBrush", new XElement(ns + "VisualBrush.Visual",
                    new XElement(glyphs[0])))))));
        page.ReplaceMarkup(markup);
        page.ReplaceStoryFragmentsMarkup(Fragments(doc, null, null, Paragraph(doc.StructureNamespace, "Group")));
        Assert.Equal("AB", Assert.Single(page.ReadContentStructure().Fragments).Blocks[0].Text);
    }

    [Fact]
    public void RepeatedPagesHaveDistinctLogicalOccurrenceAddresses() {
        var doc = Create(XpsFormat.OpenXps, new[] { "A" }); var page = doc.Pages[0];
        page.ReplaceStoryFragmentsMarkup(Fragments(doc, "body", "fragment", Paragraph(doc.StructureNamespace, "A")));
        doc.Documents[0].InsertPage(1, page);
        doc.Documents[0].ReplaceDocumentStructureMarkup(Story(doc, "body", (1, "fragment"), (2, "fragment")));
        var result = doc.ReadLogicalStructure(); Assert.True(result.IsComplete);
        Assert.Equal(new[] { 0, 1 }, Assert.Single(result.Stories).Fragments.Select(f => f.PageIndex));
        Assert.Equal("AA", Assert.Single(Assert.Single(result.Stories).Blocks).Text);
        Assert.Equal(0, Assert.Single(page.ReadContentStructure().Fragments).PageIndex);
    }

    [Fact]
    public void MissingAndUnknownSemanticsAreExplicitAndCancellationIsObserved() {
        var doc = Create(XpsFormat.OpenXps, new[] { "A" });
        Assert.False(doc.ReadLogicalStructure().HasNativeStructure); Assert.Empty(doc.ReadLogicalStructure().Stories);
        var native = Fragments(doc, null, null, Paragraph(doc.StructureNamespace, "A"));
        native.Element(doc.StructureNamespace + "StoryFragment")!.Add(new XElement(XName.Get("Semantic", "urn:extension"), "opaque"));
        doc.Pages[0].ReplaceStoryFragmentsMarkup(native);
        var reopened = XpsDocument.Load(doc.Save()); var result = reopened.ReadLogicalStructure();
        Assert.False(result.IsComplete); Assert.Contains(result.Diagnostics, d => d.Contains("Unsupported structure element"));
        Assert.Equal("opaque", reopened.Pages[0].GetStoryFragmentsMarkup()!.Descendants(XName.Get("Semantic", "urn:extension")).Single().Value);
        Assert.Throws<OperationCanceledException>(() => reopened.ReadLogicalStructure(new CancellationToken(true)));
    }

    internal static XpsDocument Create(XpsFormat format, params string[][] pageNames) {
        var doc = XpsDocument.Create(format);
        string font = doc.AddFont(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "RobotoFlex.ttf")), false);
        foreach (var names in pageNames) {
            var page = doc.AddPage(240, 180);
            for (int i = 0; i < names.Length; i++) page.AddText(names[i], font, 14, 10, 25 + i * 18);
            var markup = page.GetMarkup();
            for (int i = 0; i < names.Length; i++) markup.Elements().ElementAt(i).SetAttributeValue("Name", names[i]);
            page.ReplaceMarkup(markup);
        }
        return doc;
    }
    internal static XElement Paragraph(XNamespace ns, params string[] names) => new(ns + "ParagraphStructure",
        names.Select(name => new XElement(ns + "NamedElement", new XAttribute("NameReference", name))));
    internal static XElement Fragments(XpsDocument doc, string? story, string? fragment, params XElement[] blocks) =>
        new(doc.StructureNamespace + "StoryFragments", new XElement(doc.StructureNamespace + "StoryFragment",
            new XAttribute("FragmentType", "Content"), story == null ? null : new XAttribute("StoryName", story),
            fragment == null ? null : new XAttribute("FragmentName", fragment), blocks));
    internal static XElement Story(XpsDocument doc, string name, params (int Page, string? Fragment)[] references) =>
        new(doc.StructureNamespace + "DocumentStructure", new XElement(doc.StructureNamespace + "Story", new XAttribute("StoryName", name),
            references.Select(reference => new XElement(doc.StructureNamespace + "StoryFragmentReference", new XAttribute("Page", reference.Page),
                reference.Fragment == null ? null : new XAttribute("FragmentName", reference.Fragment)))));
}
