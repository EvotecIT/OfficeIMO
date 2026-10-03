using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfNoteStructureTests {
    [Fact]
    public void PositionedNoteLinkRemainsInsideItsNoteBeforeFollowingText() {
        var bytes = PdfDocument.Create().TaggedPdfCatalogMarkers()
            .Canvas(canvas => canvas.Structure(PdfCanvasStructureRole.Note, note => note
                .Text(new[] { PdfTextRun.Link("1", "https://example.test/note") }, 20, 20, 20, 20)
                .Text("Note body", 40, 20, 100, 20))).ToBytes();
        var tagged = Assert.IsType<PdfTaggedContentInfo>(PdfInspector.Inspect(bytes).TaggedContent);
        var note = Assert.Single(tagged.StructureElements, element => element.StructureType == "Note");
        var link = Assert.Single(tagged.StructureElements, element => element.StructureType == "Link");
        var marker = tagged.StructureElements.Single(element => element.ObjectNumber == link.ParentObjectNumber);
        Assert.Contains(marker.ObjectNumber, note.ChildElementObjectNumbers);
        Assert.NotEqual(marker.ObjectNumber, note.ChildElementObjectNumbers.Last());
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void WholeCellLinkRetainsInlineChildren(bool column, bool described) {
        var cells = new[] { new[] { PdfTableCell.RichTextCell(new[] {
            PdfTextRun.Normal("Before "),
            PdfTextRun.Inline(new PdfInlineBox(12, 10, background: PdfColor.Black,
                alternativeText: described ? "Status" : null)),
            PdfTextRun.Normal(" after")
        }, linkUri: "https://example.test/cell") } };
        var document = PdfDocument.Create().TaggedPdfCatalogMarkers();
        if (column) document.Compose(compose => compose.Page(page => page.Content(content => content.Row(row =>
            row.PercentColumn(100, nested => nested.Table(cells))))));
        else document.Table(cells);
        var tagged = Assert.IsType<PdfTaggedContentInfo>(PdfInspector.Inspect(document.ToBytes()).TaggedContent);
        var link = Assert.Single(tagged.StructureElements, element => element.StructureType == "Link");
        Assert.Equal(1, link.ObjectReferenceCount);
        var children = link.ChildElementObjectNumbers.Select(number =>
            tagged.StructureElements.Single(element => element.ObjectNumber == number)).ToArray();
        Assert.Equal(described ? new[] { "Span", "Figure", "Span" } : new[] { "Span", "Span" },
            children.Select(element => element.StructureType));
        Assert.All(children, child => Assert.Equal(1, child.MarkedContentReferenceCount));
    }

    [Theory]
    [InlineData(false, 0)]
    [InlineData(false, 1)]
    [InlineData(false, 2)]
    [InlineData(true, 0)]
    [InlineData(true, 1)]
    [InlineData(true, 2)]
    public void MixedRunsKeepLogicalContentInSourceOrder(bool inNote, int variant) {
        var runs = variant == 0
            ? new[] { PdfTextRun.Link("LINK", "https://example.test"), PdfTextRun.Normal(" AFTER") }
            : variant == 1
                ? new[] { PdfTextRun.Normal("BEFORE "), PdfTextRun.Link("LINK", "https://example.test"), PdfTextRun.Normal(" AFTER") }
                : new[] { PdfTextRun.Link("FIRST", "https://example.test/1"), PdfTextRun.Link("SECOND", "https://example.test/2"), PdfTextRun.Normal(" AFTER") };
        var document = PdfDocument.Create().TaggedPdfCatalogMarkers();
        if (inNote) document.Canvas(canvas => canvas.Structure(PdfCanvasStructureRole.Note,
            note => note.Text(runs, 20, 20, 300, 20)));
        else document.Canvas(canvas => canvas.Text(runs, 20, 20, 300, 20));
        var tagged = Assert.IsType<PdfTaggedContentInfo>(PdfInspector.Inspect(document.ToBytes()).TaggedContent);
        var root = Assert.Single(tagged.StructureElements, element => element.StructureType == "Document");
        var ordered = new List<int>();
        void Visit(PdfStructureElementInfo element) {
            ordered.AddRange(element.MarkedContentReferences.Select(reference => reference.MarkedContentId));
            foreach (int child in element.ChildElementObjectNumbers)
                Visit(tagged.StructureElements.Single(candidate => candidate.ObjectNumber == child));
        }
        Visit(root);
        // This single-line fixture paints runs in source order. Every MCID must be visited
        // once, in that same order, rather than moving text after a link before it.
        Assert.True(ordered.Count >= 3);
        Assert.Equal(ordered.OrderBy(value => value).Distinct(), ordered);
        Assert.All(tagged.StructureElements.Where(element => element.StructureType == "Link"),
            link => Assert.Equal(1, link.ObjectReferenceCount));
    }

    [Theory]
    [InlineData(PdfObjectSerializationMode.Buffered)]
    [InlineData(PdfObjectSerializationMode.ForwardOnly)]
    public void NotesHaveUniqueIdsResolvedByTheStructureIdTree(PdfObjectSerializationMode serialization) {
        var document = PdfDocument.Create(new PdfOptions {
            TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers,
            FileVersion = PdfFileVersion.Pdf17,
            ObjectSerializationMode = serialization
        }).Canvas(canvas => canvas
            .Structure(PdfCanvasStructureRole.Note, note => note.Text("First note", 20, 20, 100, 20))
            .Structure(PdfCanvasStructureRole.Note, note => note.Text("Second note", 20, 60, 100, 20)));
        var raw = PdfReadDocument.Open(document.ToBytes()).RawStructure();
        var notes = raw.Objects.Where(item => item.Value.Entries.TryGetValue("S", out var role) && role.Text == "Note").ToArray();
        Assert.Equal(2, notes.Length);
        var ids = notes.Select(note => note.Value.Entries["ID"].Text).ToArray();
        Assert.All(ids, id => Assert.False(string.IsNullOrEmpty(id)));
        Assert.Equal(2, ids.Distinct().Count());
        var root = Assert.Single(raw.Objects, item => item.Value.Entries.TryGetValue("Type", out var type) && type.Text == "StructTreeRoot");
        var names = raw.GetObject(root.Value.Entries["IDTree"].ReferenceObjectNumber!.Value)!.Value.Entries["Names"].Items;
        Assert.Equal(4, names.Count);
        for (int index = 0; index < names.Count; index += 2) {
            var target = raw.GetObject(names[index + 1].ReferenceObjectNumber!.Value)!;
            Assert.Equal(names[index].Text, target.Value.Entries["ID"].Text);
            Assert.Equal("Note", target.Value.Entries["S"].Text);
        }
    }
}
