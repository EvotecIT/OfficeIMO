namespace OfficeIMO.OneNote.Tests;

public sealed class NotebookOrderingWriterTests {
    [Fact]
    public void ReadWritePreservesInterleavedSectionAndGroupOrder() {
        var entries = new[] {
            new OneNoteTocWriteEntry(Guid.NewGuid(), "Group A", 0, null),
            new OneNoteTocWriteEntry(Guid.NewGuid(), "Middle.one", 1, 0xFF123456U),
            new OneNoteTocWriteEntry(Guid.NewGuid(), "Group C", 2, null)
        };
        OneNoteWriteGraph sourceGraph = new OneNoteWriteGraphBuilder().BuildTableOfContents(
            Guid.NewGuid(),
            Guid.Empty,
            "Open Notebook.onetoc2",
            entries,
            null,
            true);
        byte[] source = OneNoteRevisionStoreWriter.Write(sourceGraph);
        OneNoteNotebook loaded = OneNoteNotebookReader.Read(new MemoryStream(source));

        byte[] rewritten = OneNoteTableOfContentsWriter.Write(
            loaded,
            new OneNoteWriterOptions { PreserveUnknownData = false });
        OneNoteTocData result = OneNoteTocMapper.Map(OneNoteRevisionStoreReader.Read(new MemoryStream(rewritten)));

        Assert.Equal(
            new[] { "Group A", "Middle.one", "Group C" },
            result.Entries.OrderBy(item => item.Order).Select(item => item.Name));
    }
}
