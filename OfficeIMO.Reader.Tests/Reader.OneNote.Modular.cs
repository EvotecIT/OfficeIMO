using OfficeIMO.Reader;
using OfficeIMO.Reader.OneNote;
using OfficeIMO.OneNote;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderOneNoteModularTests {
    [Fact]
    public void ReaderInputKind_OneNote_AppendsWithoutChangingExistingValues() {
        Assert.Equal(19, (int)ReaderInputKind.Email);
        Assert.Equal(20, (int)ReaderInputKind.OneNote);
    }

    [Fact]
    public void OneNoteOptionsClone_PreservesTransactionChecksumCompatibilitySetting() {
        var options = new ReaderOneNoteOptions {
            OneNoteOptions = new OneNoteReaderOptions { ValidateTransactionChecksums = false }
        };

        ReaderOneNoteOptions clone = ReaderOneNoteOptionsCloner.CloneOrDefault(options);

        Assert.NotSame(options.OneNoteOptions, clone.OneNoteOptions);
        Assert.False(clone.OneNoteOptions!.ValidateTransactionChecksums);
    }

    [Fact]
    public void OneNoteAdapter_ReadDocument_ProjectsPageTextAndMetadata() {
        OfficeDocumentReadResult result = OneNoteReaderAdapter.ReadDocument(FixturePath("testOneNote2016.one"));

        Assert.Equal(ReaderInputKind.OneNote, result.Kind);
        Assert.Equal("testOneNote2016", result.Source.Title);
        Assert.Contains("officeimo.onenote.native", result.CapabilitiesUsed);
        Assert.Contains("officeimo.reader.onenote.offline", result.CapabilitiesUsed);
        ReaderChunk chunk = Assert.Single(result.Chunks);
        Assert.Equal(1, chunk.Location.Page);
        Assert.Equal("testOneNote2016 > So good", chunk.Location.HeadingPath);
        Assert.Contains("This is one note 2016", chunk.Text, StringComparison.Ordinal);
        Assert.Contains("# So good", chunk.Markdown, StringComparison.Ordinal);
        OfficeDocumentPage page = Assert.Single(result.Pages);
        Assert.Equal("So good", page.Name);
        Assert.Contains(result.Metadata, item => item.Name == "PageCount" && item.Value == "1");
        Assert.False(string.IsNullOrWhiteSpace(chunk.SourceHash));
        Assert.False(string.IsNullOrWhiteSpace(chunk.ChunkHash));
    }

    [Fact]
    public void OneNoteAdapter_ReadDocument_ProjectsEmbeddedAssetWithOptionalPayload() {
        OfficeDocumentReadResult metadataOnly = OneNoteReaderAdapter.ReadDocument(FixturePath("testOneNoteEmbeddedWordDoc.one"));
        OfficeDocumentAsset metadataAsset = Assert.Single(metadataOnly.Assets, asset => asset.Kind == "embedded-file");
        Assert.Equal("Dude this is a super cool embedded doc.docx", metadataAsset.FileName);
        Assert.True(metadataAsset.LengthBytes > 4);
        Assert.False(string.IsNullOrWhiteSpace(metadataAsset.PayloadHash));
        Assert.Null(metadataAsset.PayloadBytes);
        Assert.Contains(
            "[Dude this is a super cool embedded doc.docx](onenote-page-0001-asset-0001)",
            Assert.Single(metadataOnly.Chunks).Markdown,
            StringComparison.Ordinal);

        OfficeDocumentReadResult withPayload = OneNoteReaderAdapter.ReadDocument(
            FixturePath("testOneNoteEmbeddedWordDoc.one"),
            oneNoteOptions: new ReaderOneNoteOptions { IncludeAssetPayloads = true });
        OfficeDocumentAsset payloadAsset = Assert.Single(withPayload.Assets, asset => asset.Kind == "embedded-file");
        Assert.NotNull(payloadAsset.PayloadBytes);
        Assert.Equal((byte)'P', payloadAsset.PayloadBytes![0]);
        Assert.Equal((byte)'K', payloadAsset.PayloadBytes[1]);
        Assert.Same(payloadAsset, Assert.Single(withPayload.Pages.SelectMany(page => page.Assets), asset => asset.Id == payloadAsset.Id));
    }

    [Fact]
    public void OneNoteAdapter_BuilderDispatchesPathAndStream() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddOneNoteHandler().Build();
        string path = FixturePath("testOneNote2016.one");

        Assert.Equal(ReaderInputKind.OneNote, reader.DetectKind(path));
        Assert.Contains("This is one note 2016", Assert.Single(reader.Read(path)).Text, StringComparison.Ordinal);
        using FileStream stream = File.OpenRead(path);
        OfficeDocumentReadResult result = reader.ReadDocument(stream, "registered.one");
        Assert.Equal(ReaderInputKind.OneNote, result.Kind);
        Assert.Equal("registered.one", result.Source.Path);
        ReaderHandlerCapability capability = Assert.Single(reader.GetCapabilities(), item => item.Id == OfficeDocumentReaderBuilderOneNoteExtensions.HandlerId);
        Assert.True(capability.SupportsPath);
        Assert.True(capability.SupportsStream);
        Assert.True(capability.SupportsDocumentPath);
        Assert.True(capability.SupportsDocumentStream);
        Assert.False(capability.SupportsAsyncPath);
        Assert.False(capability.SupportsAsyncStream);
    }

    [Fact]
    public void OneNoteAdapter_NonSeekableStream_EnforcesGlobalInputLimit() {
        byte[] bytes = File.ReadAllBytes(FixturePath("testOneNote2016.one"));
        using var stream = new NonSeekableReadStream(bytes);

        IOException exception = Assert.Throws<IOException>(() => OneNoteReaderAdapter.ReadDocument(
            stream,
            "bounded.one",
            new ReaderOptions { MaxInputBytes = 32 }));

        Assert.Contains("maximum size", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void OneNoteAdapter_SeekableStream_RestoresCallerPosition() {
        byte[] bytes = File.ReadAllBytes(FixturePath("testOneNote2016.one"));
        using var stream = new MemoryStream(bytes);
        stream.Position = 7;

        OfficeDocumentReadResult result = OneNoteReaderAdapter.ReadDocument(stream, "position.one");

        Assert.Equal(7, stream.Position);
        Assert.Equal(ReaderInputKind.OneNote, result.Kind);
    }

    [Fact]
    public void OneNoteAdapter_ReadDocument_UsesSameProjectionForOffice365PackageStore() {
        OfficeDocumentReadResult result = OneNoteReaderAdapter.ReadDocument(FixturePath("testOneNoteFromOffice365.one"));

        Assert.Equal(ReaderInputKind.OneNote, result.Kind);
        Assert.Equal(new[] { "Section1Page1", "Section1Page2" }, result.Pages.Select(page => page.Name).ToArray());
        Assert.Contains(result.Chunks, chunk => chunk.Text.Contains("Section1Page1Content", StringComparison.Ordinal));
        Assert.Contains(result.Chunks, chunk => chunk.Text.Contains("Section1Page2Content", StringComparison.Ordinal));
        Assert.Contains(result.Metadata, item => item.Name == "VersionPageCount" && item.Value == "1");
    }

    [Fact]
    public void OneNoteAdapter_ProjectsStructuredTablesLinksAndStableAssetTargets() {
        var section = new OneNoteSection { Name = "Structured section" };
        var page = new OneNotePage { Title = "Structured page" };
        var table = new OneNoteTable { BordersVisible = true };
        var row = new OneNoteTableRow();
        var textCell = new OneNoteTableCell();
        var paragraph = new OneNoteParagraph();
        paragraph.Runs.Add(new OneNoteTextRun {
            Text = "Example",
            Hyperlink = "https://example.test/a path"
        });
        textCell.Content.Add(paragraph);
        var fileCell = new OneNoteTableCell();
        fileCell.Content.Add(new OneNoteEmbeddedFile {
            FileName = "inside.bin",
            Payload = OneNoteBinaryPayload.FromBytes(new byte[] { 1, 2, 3 })
        });
        row.Cells.Add(textCell);
        row.Cells.Add(fileCell);
        table.Rows.Add(row);
        page.DirectContent.Add(table);
        page.DirectContent.Add(new OneNoteImage {
            AltText = "Preview",
            FileName = "preview.png",
            MediaType = "image/png",
            Hyperlink = "https://example.test/image target",
            Payload = OneNoteBinaryPayload.FromBytes(new byte[] { 4, 5, 6 })
        });
        section.Pages.Add(page);

        OfficeDocumentReadResult result = OneNoteReaderAdapter.ReadDocument(section, "structured.one");

        ReaderTable projectedTable = Assert.Single(result.Tables);
        Assert.Equal(new[] { "Column 1", "Column 2" }, projectedTable.Columns);
        Assert.Equal(new[] { "Example", "[Embedded file: inside.bin]" }, Assert.Single(projectedTable.Rows));
        Assert.Collection(result.Links,
            link => {
                Assert.Equal("https://example.test/a path", link.Uri);
                Assert.Equal("Example", link.Text);
                Assert.Equal(1, link.Location.Page);
            },
            link => {
                Assert.Equal("https://example.test/image target", link.Uri);
                Assert.Equal("Preview", link.Text);
                Assert.Equal(1, link.Location.Page);
            });
        Assert.Equal(
            new[] { "onenote-page-0001-asset-0001", "onenote-page-0001-asset-0002" },
            result.Assets.Select(asset => asset.Id).ToArray());
        string markdown = Assert.Single(result.Chunks).Markdown!;
        Assert.Contains("[Example](https://example.test/a%20path)", markdown, StringComparison.Ordinal);
        Assert.Contains("[inside.bin](onenote-page-0001-asset-0001)", markdown, StringComparison.Ordinal);
        Assert.Contains("[![Preview](onenote-page-0001-asset-0002)](https://example.test/image%20target)", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public async Task OneNoteAdapter_AsyncFacadeReadsPathAndPreservesCallerStream() {
        string path = FixturePath("testOneNote2016.one");
        OfficeDocumentReadResult pathResult = await OneNoteReaderAdapter.ReadDocumentAsync(path);
        Assert.Equal("So good", Assert.Single(pathResult.Pages).Name);

        byte[] bytes = File.ReadAllBytes(path);
        using var stream = new MemoryStream(bytes);
        stream.Position = 11;
        OfficeDocumentReadResult streamResult = await OneNoteReaderAdapter.ReadDocumentAsync(stream, "async.one");
        Assert.Equal(11, stream.Position);
        Assert.Equal("async.one", streamResult.Source.Path);

        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddOneNoteHandler().Build();
        OfficeDocumentReadResult registeredResult = await reader.ReadDocumentAsync(path);
        Assert.Equal(ReaderInputKind.OneNote, registeredResult.Kind);
    }

    [Fact]
    public void OneNoteAdapter_ProjectsNotebookGroupAndSectionHierarchy() {
        var notebook = new OneNoteNotebook { Name = "Research" };
        var group = new OneNoteSectionGroup { Name = "Archive" };
        var section = new OneNoteSection { Name = "Experiments" };
        var parent = new OneNotePage { Title = "Alpha", Level = 0 };
        var parentText = new OneNoteParagraph();
        parentText.Runs.Add(new OneNoteTextRun { Text = "Parent observation" });
        parent.DirectContent.Add(parentText);
        var child = new OneNotePage { Title = "Follow-up", Level = 1 };
        var childText = new OneNoteParagraph();
        childText.Runs.Add(new OneNoteTextRun { Text = "Child observation" });
        child.DirectContent.Add(childText);
        section.Pages.Add(parent);
        section.Pages.Add(child);
        group.Sections.Add(section);
        notebook.SectionGroups.Add(group);

        OfficeDocumentReadResult result = OneNoteReaderAdapter.ReadDocument(notebook, "research.onepkg");

        Assert.Equal(
            new[] {
                "Research > Archive > Experiments > Alpha",
                "Research > Archive > Experiments > Alpha > Follow-up"
            },
            result.Chunks.Select(chunk => chunk.Location.HeadingPath).ToArray());
        Assert.Contains("officeimo.onenote.notebook", result.CapabilitiesUsed);
        Assert.Contains(result.Metadata, item => item.Name == "NotebookSectionCount" && item.Value == "1");
        Assert.Contains(result.Metadata, item => item.Name == "SectionGroupCount" && item.Value == "1");
    }

    [Fact]
    public void OneNoteAdapter_ContentDetectionDispatchesExtensionlessSection() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddOneNoteHandler().Build();
        using FileStream stream = File.OpenRead(FixturePath("testOneNote2016.one"));

        OfficeDocumentReadResult result = reader.ReadDocument(stream, "extensionless.bin");

        Assert.Equal(ReaderInputKind.OneNote, result.Kind);
        Assert.Contains("This is one note 2016", Assert.Single(result.Chunks).Text, StringComparison.Ordinal);
    }

    private static string FixturePath(string fileName) => Path.Combine(AppContext.BaseDirectory, "OneNoteFixtures", fileName);
}
