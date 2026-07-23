using OfficeIMO.Reader;
using OfficeIMO.Reader.OneNote;
using OfficeIMO.OneNote;
using System.Globalization;
using System.Text;
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
            IncludeConflictPages = true,
            IncludeVersionHistory = true,
            OneNoteOptions = new OneNoteReaderOptions {
                ValidateTransactionChecksums = false,
                MaxPageGraphNodes = 321,
                MaxPageRelationshipDepth = 17
            }
        };

        ReaderOneNoteOptions clone = ReaderOneNoteOptionsCloner.CloneOrDefault(options);

        Assert.NotSame(options.OneNoteOptions, clone.OneNoteOptions);
        Assert.False(clone.OneNoteOptions!.ValidateTransactionChecksums);
        Assert.Equal(321, clone.OneNoteOptions.MaxPageGraphNodes);
        Assert.Equal(17, clone.OneNoteOptions.MaxPageRelationshipDepth);
        Assert.True(clone.IncludeConflictPages);
        Assert.True(clone.IncludeVersionHistory);
    }

    [Fact]
    public void OneNoteAdapter_RelatedPagesRemainOptInAndRetainExplicitHierarchyAndAssets() {
        var section = new OneNoteSection { Name = "History" };
        var current = new OneNotePage { Title = "Current" };
        current.DirectContent.Add(Paragraph("Current text"));
        var conflict = new OneNotePage { Title = "Conflict copy", IsConflictPage = true };
        conflict.DirectContent.Add(Paragraph("Conflict text"));
        var version = new OneNotePage { Title = "Earlier copy", IsVersionHistoryPage = true };
        version.DirectContent.Add(Paragraph("Earlier text"));
        version.DirectContent.Add(new OneNoteImage {
            FileName = "historical.png",
            MediaType = "image/png",
            Payload = OneNoteBinaryPayload.FromBytes(new byte[] { 1, 2, 3 })
        });
        current.ConflictPages.Add(conflict);
        current.VersionHistory.Add(version);
        section.Pages.Add(current);

        OfficeDocumentReadResult currentOnly = OneNoteReaderAdapter.ReadDocument(section);
        Assert.Single(currentOnly.Pages);
        Assert.Empty(currentOnly.Assets);
        Assert.DoesNotContain("officeimo.onenote.conflict-pages", currentOnly.CapabilitiesUsed);
        Assert.DoesNotContain("officeimo.onenote.version-history", currentOnly.CapabilitiesUsed);
        Assert.Contains(currentOnly.Metadata, item => item.Name == "ConflictPageCount" && item.Value == "1");
        Assert.Contains(currentOnly.Metadata, item => item.Name == "VersionPageCount" && item.Value == "1");

        var options = new ReaderOneNoteOptions {
            IncludeConflictPages = true,
            IncludeVersionHistory = true
        };
        OfficeDocumentReadResult withHistory = OneNoteReaderAdapter.ReadDocument(section, oneNoteOptions: options);

        Assert.Equal(new[] { "Current", "Conflict copy", "Earlier copy" }, withHistory.Pages.Select(page => page.Name).ToArray());
        Assert.Equal(
            new[] {
                "History > Current",
                "History > Current > Conflict: Conflict copy",
                "History > Current > Version: Earlier copy"
            },
            withHistory.Chunks.Select(chunk => chunk.Location.HeadingPath).ToArray());
        Assert.Single(withHistory.Assets, asset => asset.FileName == "historical.png");
        Assert.Contains("officeimo.onenote.conflict-pages", withHistory.CapabilitiesUsed);
        Assert.Contains("officeimo.onenote.version-history", withHistory.CapabilitiesUsed);
    }

    [Fact]
    public void OneNoteAdapterRejectsRecursiveModelGraphsBeforeProjectionTraversal() {
        var contentSection = new OneNoteSection { Name = "Content" };
        var contentPage = new OneNotePage { Title = "Cycle" };
        var outline = new OneNoteOutline();
        outline.Children.Add(outline);
        contentPage.Outlines.Add(outline);
        contentSection.Pages.Add(contentPage);
        OneNoteFormatException contentError = Assert.Throws<OneNoteFormatException>(() => OneNoteReaderAdapter.ReadDocument(contentSection));
        Assert.Equal("ONENOTE_PROJECTION_CONTENT_CYCLE", contentError.Code);

        var relatedSection = new OneNoteSection { Name = "Related" };
        var relatedPage = new OneNotePage { Title = "Cycle" };
        relatedPage.VersionHistory.Add(relatedPage);
        relatedSection.Pages.Add(relatedPage);
        OneNoteFormatException relatedError = Assert.Throws<OneNoteFormatException>(() => OneNoteReaderAdapter.ReadDocument(relatedSection));
        Assert.Equal("ONENOTE_PROJECTION_PAGE_CYCLE", relatedError.Code);

        var notebook = new OneNoteNotebook { Name = "Notebook" };
        var group = new OneNoteSectionGroup { Name = "Cycle" };
        group.SectionGroups.Add(group);
        notebook.SectionGroups.Add(group);
        OneNoteFormatException groupError = Assert.Throws<OneNoteFormatException>(() => OneNoteReaderAdapter.ReadDocument(notebook));
        Assert.Equal("ONENOTE_PROJECTION_GROUP_CYCLE", groupError.Code);
    }

    [Fact]
    public void OneNoteAdapter_EmitsUnresolvedImageMetadataAndConvertsHalfInchesToPixels() {
        var section = new OneNoteSection { Name = "Images" };
        var page = new OneNotePage { Title = "Page" };
        page.DirectContent.Add(new OneNoteImage {
            AltText = "Unavailable",
            FileName = "unresolved.png",
            MediaType = "image/png",
            WidthHalfInches = 2.5,
            HeightHalfInches = 1.5
        });
        page.DirectContent.Add(new OneNoteImage {
            AltText = "Available",
            FileName = "available.png",
            MediaType = "image/png",
            Payload = OneNoteBinaryPayload.FromBytes(new byte[] { 4, 5, 6 })
        });
        section.Pages.Add(page);

        OfficeDocumentReadResult result = OneNoteReaderAdapter.ReadDocument(section);

        Assert.Collection(result.Assets,
            unresolved => {
                Assert.Equal("unresolved.png", unresolved.FileName);
                Assert.Equal(120, unresolved.Width);
                Assert.Equal(72, unresolved.Height);
                Assert.Null(unresolved.LengthBytes);
                Assert.Null(unresolved.PayloadBytes);
            },
            available => Assert.Equal("available.png", available.FileName));
        string markdown = Assert.Single(result.Chunks).Markdown!;
        Assert.Contains("[Image: Unavailable]", markdown, StringComparison.Ordinal);
        Assert.Contains("onenote-page-0001-asset-0002", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void OneNoteAdapter_NormalizesUnsupportedProjectionCharactersWithoutMutatingSource() {
        const string raw = "Alpha\vBeta\u0001\uFDDF";
        const string rawTitle = "Page\vName\uFDDF";
        const string rawLink = "https://example.test/\uFDDF";
        var section = new OneNoteSection { Name = "Controls" };
        var page = new OneNotePage { Title = rawTitle };
        page.DirectContent.Add(Paragraph(raw));
        ((OneNoteParagraph)page.DirectContent[0]).Runs[0].Hyperlink = rawLink;
        section.Pages.Add(page);

        OfficeDocumentReadResult result = OneNoteReaderAdapter.ReadDocument(section);
        ReaderChunk chunk = Assert.Single(result.Chunks);

        Assert.Contains("Alpha\nBeta??", chunk.Text, StringComparison.Ordinal);
        Assert.Contains("Alpha<br>Beta??", chunk.Markdown, StringComparison.Ordinal);
        Assert.DoesNotContain('\v', chunk.Text);
        Assert.DoesNotContain('\uFDDF', chunk.Text);
        Assert.Equal("Page\nName?", Assert.Single(result.Pages).Name);
        Assert.Equal("https://example.test/?", Assert.Single(result.Links).Uri);
        Assert.Equal(raw, ((OneNoteParagraph)page.DirectContent[0]).Runs[0].Text);
        Assert.Equal(rawLink, ((OneNoteParagraph)page.DirectContent[0]).Runs[0].Hyperlink);
        Assert.Equal(rawTitle, page.Title);
    }

    [Fact]
    public void OneNoteAdapter_BoundsCallerSuppliedListIndentation() {
        var section = new OneNoteSection { Name = "Lists" };
        var page = new OneNotePage { Title = "Extreme list" };
        var paragraph = new OneNoteParagraph {
            List = new OneNoteListInfo { Level = int.MaxValue }
        };
        paragraph.Runs.Add(new OneNoteTextRun { Text = "Bounded item" });
        page.DirectContent.Add(paragraph);
        section.Pages.Add(page);

        ReaderChunk chunk = Assert.Single(OneNoteReaderAdapter.ReadDocument(section).Chunks);
        string expectedPrefix = new string(' ', OneNoteListInfo.MaxLevel * 2) + "- Bounded item";

        Assert.Contains(expectedPrefix, chunk.Markdown!, StringComparison.Ordinal);
        Assert.Contains("Bounded item", chunk.Text, StringComparison.Ordinal);
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
    public void OneNoteAdapter_ChunksTextAndMarkdownAtSharedSemanticBoundaries() {
        var section = new OneNoteSection { Name = "Chunked" };
        var page = new OneNotePage { Title = "Chunked page" };
        page.DirectContent.Add(BoldParagraph("ALPHA " + new string('a', 32)));
        page.DirectContent.Add(BoldParagraph("BETA " + new string('b', 32)));
        section.Pages.Add(page);

        OfficeDocumentReadResult result = OneNoteReaderAdapter.ReadDocument(
            section,
            readerOptions: new ReaderOptions { MaxChars = 60 });

        Assert.Equal(2, result.Chunks.Count);
        Assert.All(result.Chunks, chunk => {
            Assert.True(chunk.Text.Length <= 60);
            Assert.True(chunk.Markdown!.Length <= 60);
            Assert.Equal(
                chunk.Text.Contains("ALPHA", StringComparison.Ordinal),
                chunk.Markdown.Contains("ALPHA", StringComparison.Ordinal));
            Assert.Equal(
                chunk.Text.Contains("BETA", StringComparison.Ordinal),
                chunk.Markdown.Contains("BETA", StringComparison.Ordinal));
        });
    }

    [Fact]
    public void OneNoteAdapter_SplitsOversizedNonTextUnitsAtMaxChars() {
        string title = "Heading " + new string('h', 40);
        string altText = "Preview " + new string('i', 40);
        var section = new OneNoteSection { Name = "Oversized" };
        var page = new OneNotePage { Title = title };
        page.DirectContent.Add(new OneNoteImage {
            AltText = altText,
            FileName = "preview.png",
            MediaType = "image/png",
            Payload = OneNoteBinaryPayload.FromBytes(new byte[] { 1, 2, 3 })
        });
        section.Pages.Add(page);

        OfficeDocumentReadResult result = OneNoteReaderAdapter.ReadDocument(
            section,
            readerOptions: new ReaderOptions { MaxChars = 16 });

        Assert.True(result.Chunks.Count > 2);
        Assert.All(result.Chunks, chunk => {
            Assert.True(chunk.Text.Length <= 16, chunk.Text);
            Assert.True(chunk.Markdown!.Length <= 16, chunk.Markdown);
        });
        Assert.Equal(title + "[Image: " + altText + "]",
            string.Concat(result.Chunks.Select(chunk => chunk.Text)));
        Assert.Equal("# " + title + "![" + altText + "](onenote-page-0001-asset-0001)",
            string.Concat(result.Chunks.Select(chunk => chunk.Markdown)));
    }

    [Theory]
    [InlineData("one two")]
    [InlineData("one  two")]
    [InlineData("one\ttwo")]
    public void OneNoteAdapter_PreservesWhitespaceAcrossRunChunkBoundaries(string sourceText) {
        var section = new OneNoteSection { Name = "Chunked" };
        var page = new OneNotePage { Title = "P" };
        var paragraph = new OneNoteParagraph();
        paragraph.Runs.Add(new OneNoteTextRun { Text = sourceText });
        page.DirectContent.Add(paragraph);
        section.Pages.Add(page);

        OfficeDocumentReadResult result = OneNoteReaderAdapter.ReadDocument(
            section,
            readerOptions: new ReaderOptions { MaxChars = 5 });
        ReaderChunk[] contentChunks = result.Chunks.Skip(1).ToArray();

        Assert.NotEmpty(contentChunks);
        Assert.Equal(sourceText, string.Concat(contentChunks.Select(chunk => chunk.Text)));
        Assert.Equal(sourceText, string.Concat(contentChunks.Select(chunk => chunk.Markdown)));
        Assert.All(result.Chunks, chunk => {
            Assert.True(chunk.Text.Length <= 5);
            Assert.True(chunk.Markdown!.Length <= 5);
        });
    }

    [Fact]
    public void OneNoteAdapter_PreservesSupplementaryUnicodeAcrossChunkBoundaries() {
        const string sourceText = "\U0001F642X";
        var section = new OneNoteSection { Name = "Unicode" };
        var page = new OneNotePage { Title = "P" };
        page.DirectContent.Add(Paragraph(sourceText));
        section.Pages.Add(page);

        OfficeDocumentReadResult result = OneNoteReaderAdapter.ReadDocument(
            section,
            readerOptions: new ReaderOptions { MaxChars = 1 });
        ReaderChunk[] contentChunks = result.Chunks.Skip(1).ToArray();

        Assert.Equal(sourceText, string.Concat(contentChunks.Select(chunk => chunk.Text)));
        Assert.Equal(sourceText, string.Concat(contentChunks.Select(chunk => chunk.Markdown)));
        Assert.DoesNotContain('?', string.Concat(contentChunks.Select(chunk => chunk.Text)));
        Assert.Equal("\U0001F642", contentChunks[0].Text);
        Assert.Contains(contentChunks[0].Warnings!, warning => warning.Contains("preserved as one chunk", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void OneNoteAdapter_MarkdownEscapesLiteralCodeAndStrikethroughDelimiters() {
        var section = new OneNoteSection { Name = "Literal delimiters" };
        var page = new OneNotePage { Title = "Literal `title` and ~~strike~~" };
        var paragraph = new OneNoteParagraph();
        paragraph.Runs.Add(new OneNoteTextRun { Text = "Literal `code` and ~~deleted~~" });
        page.DirectContent.Add(paragraph);
        section.Pages.Add(page);

        string markdown = Assert.Single(OneNoteReaderAdapter.ReadDocument(section).Chunks).Markdown!;

        Assert.Contains("# Literal \\`title\\` and \\~\\~strike\\~\\~", markdown, StringComparison.Ordinal);
        Assert.Contains("Literal \\`code\\` and \\~\\~deleted\\~\\~", markdown, StringComparison.Ordinal);
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
    public void OneNoteAdapter_UnknownLengthPayloadsRespectAggregateMaterializationLimit() {
        var section = new OneNoteSection { Name = "Bounded assets" };
        var page = new OneNotePage { Title = "Assets" };
        page.DirectContent.Add(new OneNoteEmbeddedFile {
            FileName = "first.bin",
            Payload = OneNoteBinaryPayload.FromStreamFactory(() => new MemoryStream(new byte[] { 1, 2, 3, 4 }))
        });
        page.DirectContent.Add(new OneNoteEmbeddedFile {
            FileName = "second.bin",
            Payload = OneNoteBinaryPayload.FromStreamFactory(() => new MemoryStream(new byte[] { 5, 6, 7, 8 }))
        });
        section.Pages.Add(page);

        OfficeDocumentReadResult result = OneNoteReaderAdapter.ReadDocument(
            section,
            oneNoteOptions: new ReaderOneNoteOptions {
                IncludeAssetPayloads = true,
                OneNoteOptions = new OneNoteReaderOptions {
                    MaxAssetBytes = 4,
                    MaxTotalAssetBytes = 6
                }
            });

        Assert.Collection(result.Assets,
            first => Assert.Equal(new byte[] { 1, 2, 3, 4 }, first.PayloadBytes),
            second => Assert.Null(second.PayloadBytes));
        Assert.Equal(4, result.Assets.Sum(asset => asset.PayloadBytes?.LongLength ?? 0));
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
    public void OneNoteAdapter_ProjectsNotebookInTableOfContentsOrderAtEveryLevel() {
        var notebook = new OneNoteNotebook { Name = "Ordered" };
        OneNoteSection middle = SectionWithPage("Middle", "Middle page");
        middle.TableOfContentsOrder = 1;
        notebook.Sections.Add(middle);

        var last = new OneNoteSectionGroup { Name = "Last", TableOfContentsOrder = 2 };
        last.Sections.Add(SectionWithPage("Last section", "Last page"));
        notebook.SectionGroups.Add(last);

        var first = new OneNoteSectionGroup { Name = "First", TableOfContentsOrder = 0 };
        OneNoteSection nestedMiddle = SectionWithPage("Nested middle", "Nested middle page");
        nestedMiddle.TableOfContentsOrder = 1;
        first.Sections.Add(nestedMiddle);
        var nestedFirst = new OneNoteSectionGroup { Name = "Nested first", TableOfContentsOrder = 0 };
        nestedFirst.Sections.Add(SectionWithPage("Nested first section", "Nested first page"));
        first.SectionGroups.Add(nestedFirst);
        notebook.SectionGroups.Add(first);

        OfficeDocumentReadResult result = OneNoteReaderAdapter.ReadDocument(notebook, "ordered.onepkg");

        Assert.Equal(
            new[] {
                "Ordered > First > Nested first > Nested first section > Nested first page",
                "Ordered > First > Nested middle > Nested middle page",
                "Ordered > Middle > Middle page",
                "Ordered > Last > Last section > Last page"
            },
            result.Chunks.Select(chunk => chunk.Location.HeadingPath).ToArray());
    }

    [Fact]
    public void OneNoteAdapter_BoundsUntrustedPageHierarchyDepth() {
        var section = new OneNoteSection { Name = "Bounded hierarchy" };
        section.Pages.Add(new OneNotePage { Title = "Deep page", Level = int.MaxValue });

        ReaderChunk chunk = Assert.Single(OneNoteReaderAdapter.ReadDocument(section).Chunks);
        string[] hierarchy = chunk.Location.HeadingPath!.Split(new[] { " > " }, StringSplitOptions.None);

        Assert.StartsWith("###### Deep page", chunk.Markdown!, StringComparison.Ordinal);
        Assert.Equal("Bounded hierarchy", hierarchy[0]);
        Assert.Equal("Deep page", hierarchy[hierarchy.Length - 1]);
        Assert.True(hierarchy.Length <= 34);
    }

    [Fact]
    public void OneNoteAdapter_ContentDetectionDispatchesExtensionlessSection() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddOneNoteHandler().Build();
        using FileStream stream = File.OpenRead(FixturePath("testOneNote2016.one"));

        OfficeDocumentReadResult result = reader.ReadDocument(stream, "extensionless.bin");

        Assert.Equal(ReaderInputKind.OneNote, result.Kind);
        Assert.Contains("This is one note 2016", Assert.Single(result.Chunks).Text, StringComparison.Ordinal);
    }

    [Fact]
    public void OneNoteAdapter_ContentDetectionDispatchesExtensionlessPackage() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddOneNoteHandler().Build();
        using var stream = new MemoryStream(CreateOneNotePackage());

        ReaderDetectionResult detection = reader.Detect(stream, "upload.bin");
        OfficeDocumentReadResult result = reader.ReadDocument(stream, "upload.bin");

        Assert.Equal(ReaderInputKind.OneNote, detection.Kind);
        Assert.True(detection.ContainerInspected);
        Assert.Contains("container:onenote-package", detection.Evidence);
        Assert.Equal(ReaderInputKind.OneNote, result.Kind);
        Assert.Equal("Packaged page", Assert.Single(result.Pages).Name);
        Assert.Contains("Offline package content", Assert.Single(result.Chunks).Text, StringComparison.Ordinal);
    }

    [Fact]
    public void OneNoteAdapter_ContentDetectionFindsRootTocBeforeLargePackageEntryLimit() {
        var notebook = new OneNoteNotebook { Name = "Large package" };
        for (int index = 0; index < ReaderOptions.DefaultDetectionMaxContainerEntries; index++) {
            notebook.Sections.Add(new OneNoteSection { Name = "Section " + index.ToString("D4", CultureInfo.InvariantCulture) });
        }
        byte[] package = OneNotePackageWriter.Write(
            notebook,
            new OneNoteWriterOptions { ValidateRoundTrip = false });
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddOneNoteHandler().Build();
        using var stream = new MemoryStream(package);

        ReaderDetectionResult detection = reader.Detect(stream, "upload.bin");

        Assert.Equal(ReaderInputKind.OneNote, detection.Kind);
        Assert.True(detection.ContainerInspected);
        Assert.Contains("container:onenote-package", detection.Evidence);
    }

    [Fact]
    public async Task OneNoteAdapter_ContentDetectionDispatchesExtensionlessPackageAsync() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddOneNoteHandler().Build();
        using var stream = new MemoryStream(CreateOneNotePackage());

        OfficeDocumentReadResult result = await reader.ReadDocumentAsync(stream, "upload.bin");

        Assert.Equal(ReaderInputKind.OneNote, result.Kind);
        Assert.Equal("Packaged page", Assert.Single(result.Pages).Name);
        Assert.Contains("Offline package content", Assert.Single(result.Chunks).Text, StringComparison.Ordinal);
    }

    [Fact]
    public void OneNoteAdapter_ContentDetectionDoesNotTreatGenericCabinetAsOneNote() {
        byte[] cabinet = CreateOneNotePackage();
        byte[] tocExtension = Encoding.ASCII.GetBytes(".onetoc2");
        bool replaced = false;
        for (int index = 0; index <= cabinet.Length - tocExtension.Length; index++) {
            bool match = true;
            for (int markerIndex = 0; markerIndex < tocExtension.Length; markerIndex++) {
                if (cabinet[index + markerIndex] != tocExtension[markerIndex]) {
                    match = false;
                    break;
                }
            }
            if (!match) continue;
            cabinet[index + tocExtension.Length - 1] = (byte)'x';
            replaced = true;
        }
        Assert.True(replaced);

        ReaderDetectionResult detection = OfficeIMO.Reader.Tests.ReaderTestReaders.All.Detect(cabinet, "archive.bin");

        Assert.Equal(ReaderInputKind.Unknown, detection.Kind);
        Assert.Equal(ReaderInputKind.Unknown, detection.ContentKind);
        Assert.True(detection.ContainerInspected);
        Assert.Contains("container:cabinet-generic", detection.Evidence);
    }

    private static byte[] CreateOneNotePackage() {
        var notebook = new OneNoteNotebook { Name = "Packaged notebook" };
        var section = new OneNoteSection { Name = "Packaged section" };
        var page = new OneNotePage { Title = "Packaged page" };
        var paragraph = new OneNoteParagraph();
        paragraph.Runs.Add(new OneNoteTextRun { Text = "Offline package content" });
        page.DirectContent.Add(paragraph);
        section.Pages.Add(page);
        notebook.Sections.Add(section);
        return OneNotePackageWriter.Write(notebook);
    }

    private static OneNoteSection SectionWithPage(string sectionName, string pageTitle) {
        var section = new OneNoteSection { Name = sectionName };
        section.Pages.Add(new OneNotePage { Title = pageTitle });
        return section;
    }

    private static OneNoteParagraph BoldParagraph(string text) {
        var paragraph = new OneNoteParagraph();
        var run = new OneNoteTextRun { Text = text };
        run.Style.Bold = true;
        paragraph.Runs.Add(run);
        return paragraph;
    }

    private static OneNoteParagraph Paragraph(string text) {
        var paragraph = new OneNoteParagraph();
        paragraph.Runs.Add(new OneNoteTextRun { Text = text });
        return paragraph;
    }

    private static string FixturePath(string fileName) => Path.Combine(AppContext.BaseDirectory, "OneNoteFixtures", fileName);
}
