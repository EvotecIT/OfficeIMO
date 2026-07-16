namespace OfficeIMO.OneNote.Tests;

public sealed class NotebookWriterTests {
    [Fact]
    public void TableOfContentsUsesMsoTransactionChecksum() {
        byte[] data = OneNoteTableOfContentsWriter.Write(CreateNotebook());
        OneNoteFileHeader header = OneNoteFileProbe.ReadHeader(new MemoryStream(data));
        OneNoteFileChunkReference transaction = Assert.IsType<OneNoteFileChunkReference>(header.TransactionLog);
        int offset = checked((int)transaction.Offset);
        int sentinelOffset = offset;
        while (BitConverter.ToUInt32(data, sentinelOffset) != 1U) sentinelOffset += 8;

        uint expected = OneNoteCrc32.ComputeMso(data.Skip(offset).Take(sentinelOffset - offset).ToArray());
        Assert.Equal(expected, BitConverter.ToUInt32(data, sentinelOffset + 4));
        OneNoteRevisionStoreReader.Read(new MemoryStream(data));
    }

    [Fact]
    public void TableOfContentsEmitsDependencyOverridesAfterObjectDeclarations() {
        OneNoteRevisionStore store = OneNoteRevisionStoreReader.Read(
            new MemoryStream(OneNoteTableOfContentsWriter.Write(CreateNotebook())));
        OneNoteFileNodeList revision = Assert.Single(store.FileNodeLists, list =>
            list.Nodes.Any(node => node.Id == OneNoteFileNodeId.RevisionManifestStart4));

        int lastDeclaration = revision.Nodes
            .Select((node, index) => new { node, index })
            .Where(item => item.node.Id == OneNoteFileNodeId.ObjectDeclarationWithRefCount || item.node.Id == OneNoteFileNodeId.ObjectDeclarationWithRefCount2)
            .Max(item => item.index);
        int dependencies = revision.Nodes
            .Select((node, index) => new { node, index })
            .Single(item => item.node.Id == OneNoteFileNodeId.ObjectInfoDependencyOverrides)
            .index;
        int firstRoot = revision.Nodes
            .Select((node, index) => new { node, index })
            .Where(item => item.node.Id == OneNoteFileNodeId.RootObjectReference2)
            .Min(item => item.index);

        Assert.True(lastDeclaration < dependencies);
        Assert.True(dependencies < firstRoot);
    }

    [Fact]
    public void TableOfContentsRoundTripsRootHierarchy() {
        OneNoteNotebook original = CreateNotebook();

        byte[] data = OneNoteTableOfContentsWriter.Write(original);
        OneNoteNotebook result = OneNoteNotebookReader.Read(new MemoryStream(data));

        OneNoteFileHeader header = OneNoteFileProbe.ReadHeader(new MemoryStream(data));
        Assert.Equal(OneNoteFileKind.TableOfContents, header.FileKind);
        Assert.Equal(OneNoteStorageFormat.RevisionStore, header.StorageFormat);
        Assert.Equal("Root", Assert.Single(result.Sections).Name);
        Assert.Equal("Group", Assert.Single(result.SectionGroups).Name);
    }

    [Fact]
    public void PackageRoundTripsNestedNotebookAndUsesManagedCabinet() {
        OneNoteNotebook original = CreateNotebook();

        byte[] data = OneNotePackageWriter.Write(original);
        OneNoteNotebook result = OneNotePackageReader.Read(new MemoryStream(data), "Writer.onepkg");

        Assert.Equal(new byte[] { (byte)'M', (byte)'S', (byte)'C', (byte)'F' }, data.Take(4));
        Assert.Equal("Root page", Assert.Single(Assert.Single(result.Sections).Pages).Title);
        OneNoteSectionGroup group = Assert.Single(result.SectionGroups);
        Assert.Equal("Nested page", Assert.Single(Assert.Single(group.Sections).Pages).Title);
        Assert.Empty(result.Diagnostics);
    }

    [Fact]
    public void DirectoryWriterCreatesNativeHierarchyAndRefusesExistingContent() {
        string root = Path.Combine(Path.GetTempPath(), "OfficeIMO-OneNote-" + Guid.NewGuid().ToString("N"));
        try {
            OneNoteNotebookWriter.Write(CreateNotebook(), root);

            OneNoteNotebook result = OneNoteNotebookReader.Read(Path.Combine(root, "Open Notebook.onetoc2"));
            Assert.Equal("Root page", Assert.Single(Assert.Single(result.Sections).Pages).Title);
            Assert.True(File.Exists(Path.Combine(root, "Group", "Open Notebook.onetoc2")));
            Assert.Throws<IOException>(() => OneNoteNotebookWriter.Write(CreateNotebook(), root));
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, true);
        }
    }

    private static OneNoteNotebook CreateNotebook() {
        var notebook = new OneNoteNotebook { Name = "Writer", ColorArgb = 0xFF123456U, HistoryEnabled = true };
        notebook.Sections.Add(CreateSection("Root", "Root page"));
        var group = new OneNoteSectionGroup { Name = "Group" };
        group.Sections.Add(CreateSection("Nested", "Nested page"));
        notebook.SectionGroups.Add(group);
        return notebook;
    }

    private static OneNoteSection CreateSection(string name, string title) {
        var section = new OneNoteSection { Name = name };
        var page = new OneNotePage { Title = title };
        var paragraph = new OneNoteParagraph();
        paragraph.Runs.Add(new OneNoteTextRun { Text = title + " content" });
        page.DirectContent.Add(paragraph);
        section.Pages.Add(page);
        return section;
    }
}
