namespace OfficeIMO.OneNote.Tests;

public sealed class WriterModelSafetyTests {
    [Fact]
    public void WriterRejectsCyclicConflictPagesBeforeDescending() {
        var section = new OneNoteSection { Name = "Cycle" };
        var page = new OneNotePage { Title = "Current" };
        page.ConflictPages.Add(page);
        section.Pages.Add(page);

        OneNoteFormatException exception = Assert.Throws<OneNoteFormatException>(() =>
            OneNoteSectionWriter.Write(section, NoRoundTrip()));

        Assert.Equal("ONENOTE_WRITE_PAGE_CYCLE", exception.Code);
    }

    [Fact]
    public void WriterRejectsRelatedPageDepthPastTheConfiguredLimit() {
        OneNoteSection section = CreateConflictChain(4);

        OneNoteFormatException exception = Assert.Throws<OneNoteFormatException>(() =>
            OneNoteSectionWriter.Write(section, new OneNoteWriterOptions {
                MaxPageRelationshipDepth = 3,
                ValidateRoundTrip = false
            }));

        Assert.Equal("ONENOTE_WRITE_PAGE_DEPTH", exception.Code);
    }

    [Fact]
    public void WriterAllowsRelatedPageDepthAtTheConfiguredLimit() {
        byte[] data = OneNoteSectionWriter.Write(CreateConflictChain(3), new OneNoteWriterOptions {
            MaxPageRelationshipDepth = 3
        });

        OneNotePage page = Assert.Single(OneNoteSectionReader.Read(new MemoryStream(data)).Pages);
        Assert.Equal("Page 2", Assert.Single(Assert.Single(page.ConflictPages).ConflictPages).Title);
    }

    [Fact]
    public void PackageWriterPropagatesRelatedPageDepthLimits() {
        var notebook = new OneNoteNotebook { Name = "Bounded" };
        notebook.Sections.Add(CreateConflictChain(3));

        OneNoteFormatException exception = Assert.Throws<OneNoteFormatException>(() =>
            OneNotePackageWriter.Write(notebook, new OneNoteWriterOptions {
                MaxPageRelationshipDepth = 2,
                ValidateRoundTrip = false
            }));

        Assert.Equal("ONENOTE_WRITE_PAGE_DEPTH", exception.Code);
    }

    [Fact]
    public void WriterRejectsIndirectContentCyclesBeforeDescending() {
        var section = new OneNoteSection { Name = "Content cycle" };
        var page = new OneNotePage { Title = "Current" };
        var outline = new OneNoteOutline();
        var paragraph = new OneNoteParagraph();
        outline.Children.Add(paragraph);
        paragraph.Children.Add(outline);
        page.Outlines.Add(outline);
        section.Pages.Add(page);

        OneNoteFormatException exception = Assert.Throws<OneNoteFormatException>(() =>
            OneNoteSectionWriter.Write(section, NoRoundTrip()));

        Assert.Equal("ONENOTE_WRITE_CONTENT_CYCLE", exception.Code);
    }

    [Fact]
    public void WriterRejectsSharedContentAcrossTableCells() {
        var section = new OneNoteSection { Name = "Shared content" };
        var page = new OneNotePage { Title = "Current" };
        var table = new OneNoteTable();
        var row = new OneNoteTableRow();
        var first = new OneNoteTableCell();
        var second = new OneNoteTableCell();
        var paragraph = new OneNoteParagraph();
        paragraph.Runs.Add(new OneNoteTextRun { Text = "Shared" });
        first.Content.Add(paragraph);
        second.Content.Add(paragraph);
        row.Cells.Add(first);
        row.Cells.Add(second);
        table.Rows.Add(row);
        page.DirectContent.Add(table);
        section.Pages.Add(page);

        OneNoteFormatException exception = Assert.Throws<OneNoteFormatException>(() =>
            OneNoteSectionWriter.Write(section, NoRoundTrip()));

        Assert.Equal("ONENOTE_WRITE_SHARED_CONTENT", exception.Code);
    }

    [Fact]
    public void WriterRejectsSharedPageInstances() {
        var section = new OneNoteSection { Name = "Shared page" };
        var page = new OneNotePage { Title = "Current" };
        section.Pages.Add(page);
        section.Pages.Add(page);

        OneNoteFormatException exception = Assert.Throws<OneNoteFormatException>(() =>
            OneNoteSectionWriter.Write(section, NoRoundTrip()));

        Assert.Equal("ONENOTE_WRITE_SHARED_PAGE", exception.Code);
    }

    [Fact]
    public void WriterRejectsContentDepthPastTheConfiguredLimit() {
        OneNoteSection section = CreateOutlineChain(4);

        OneNoteFormatException exception = Assert.Throws<OneNoteFormatException>(() =>
            OneNoteSectionWriter.Write(section, new OneNoteWriterOptions {
                MaxContentDepth = 3,
                ValidateRoundTrip = false
            }));

        Assert.Equal("ONENOTE_WRITE_CONTENT_DEPTH", exception.Code);
    }

    [Fact]
    public void WriterAllowsContentDepthAtTheConfiguredLimit() {
        byte[] data = OneNoteSectionWriter.Write(CreateOutlineChain(3), new OneNoteWriterOptions {
            MaxContentDepth = 3
        });

        OneNotePage page = Assert.Single(OneNoteSectionReader.Read(new MemoryStream(data)).Pages);
        OneNoteOutline first = Assert.Single(page.Outlines);
        OneNoteOutline second = Assert.IsType<OneNoteOutline>(Assert.Single(first.Children));
        Assert.IsType<OneNoteOutline>(Assert.Single(second.Children));
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(255)]
    [InlineData(int.MaxValue)]
    public void WriterRejectsListLevelsOutsideTheNativeRange(int level) {
        OneNoteSection section = CreateListedSection(level);

        OneNoteFormatException exception = Assert.Throws<OneNoteFormatException>(() =>
            OneNoteSectionWriter.Write(section, NoRoundTrip()));

        Assert.Equal("ONENOTE_WRITE_LIST_LEVEL", exception.Code);
        Assert.Single(Assert.Single(section.Pages).DirectContent);
    }

    [Fact]
    public void WriterRoundTripsTheMaximumNativeListLevel() {
        byte[] data = OneNoteSectionWriter.Write(CreateListedSection(OneNoteListInfo.MaxLevel));

        OneNotePage page = Assert.Single(OneNoteSectionReader.Read(new MemoryStream(data)).Pages);
        OneNoteParagraph paragraph = Assert.IsType<OneNoteParagraph>(
            Assert.Single(Assert.Single(page.Outlines).Children));
        Assert.Equal(OneNoteListInfo.MaxLevel, paragraph.List!.Level);
    }

    [Fact]
    public void WriterReusesASharedListDescriptorWithinAnObjectSpace() {
        var section = new OneNoteSection { Name = "Lists" };
        var page = new OneNotePage { Title = "Shared list" };
        var list = new OneNoteListInfo { Ordered = true, Level = 1 };
        for (int index = 1; index <= 2; index++) {
            var paragraph = new OneNoteParagraph { List = list };
            paragraph.Runs.Add(new OneNoteTextRun { Text = "Item " + index });
            page.DirectContent.Add(paragraph);
        }
        section.Pages.Add(page);

        byte[] data = OneNoteSectionWriter.Write(section);

        OneNoteOutline outline = Assert.Single(Assert.Single(OneNoteSectionReader.Read(new MemoryStream(data)).Pages).Outlines);
        OneNoteParagraph[] paragraphs = outline.Children.Cast<OneNoteParagraph>().ToArray();
        Assert.Equal(2, paragraphs.Length);
        Assert.All(paragraphs, paragraph => Assert.Equal(1, paragraph.List!.Level));
        Assert.Equal(new[] { "Item 1", "Item 2" }, paragraphs.Select(paragraph => Assert.Single(paragraph.Runs).Text));
    }

    [Fact]
    public void WriterSeparatesEditedListDescriptorsThatRetainANativeIdentity() {
        var section = new OneNoteSection { Name = "Lists" };
        var page = new OneNotePage { Title = "Edited lists" };
        var retainedId = new OneNoteExtendedGuid(Guid.NewGuid(), 1, 17);
        var first = new OneNoteParagraph {
            List = new OneNoteListInfo { ObjectId = retainedId, Ordered = true, Format = 3, Level = 1, FontFamily = "Aptos" }
        };
        first.Runs.Add(new OneNoteTextRun { Text = "Numbered" });
        var second = new OneNoteParagraph {
            List = new OneNoteListInfo { ObjectId = retainedId, Ordered = false, Level = 2, FontFamily = "Calibri" }
        };
        second.Runs.Add(new OneNoteTextRun { Text = "Bullet" });
        page.DirectContent.Add(first);
        page.DirectContent.Add(second);
        section.Pages.Add(page);

        byte[] data = OneNoteSectionWriter.Write(section);

        OneNoteParagraph[] paragraphs = Assert.Single(Assert.Single(OneNoteSectionReader.Read(new MemoryStream(data)).Pages).Outlines)
            .Children.Cast<OneNoteParagraph>().ToArray();
        Assert.True(paragraphs[0].List!.Ordered);
        Assert.Equal(3U, paragraphs[0].List!.Format);
        Assert.Equal("Aptos", paragraphs[0].List!.FontFamily);
        Assert.False(paragraphs[1].List!.Ordered);
        Assert.Equal(2, paragraphs[1].List!.Level);
        Assert.Equal("Calibri", paragraphs[1].List!.FontFamily);
        Assert.NotEqual(paragraphs[0].List!.ObjectId, paragraphs[1].List!.ObjectId);
    }

    [Fact]
    public void WriterSeparatesEditedParagraphStylesThatRetainANativeIdentity() {
        var section = new OneNoteSection { Name = "Styles" };
        var page = new OneNotePage { Title = "Edited styles" };
        var retainedId = new OneNoteExtendedGuid(Guid.NewGuid(), 1, 17);
        var first = new OneNoteParagraph();
        first.Style.ObjectId = retainedId;
        first.Style.Alignment = OneNoteParagraphAlignment.Left;
        first.Runs.Add(new OneNoteTextRun { Text = "Left" });
        var second = new OneNoteParagraph();
        second.Style.ObjectId = retainedId;
        second.Style.Alignment = OneNoteParagraphAlignment.Right;
        second.Runs.Add(new OneNoteTextRun { Text = "Right" });
        page.DirectContent.Add(first);
        page.DirectContent.Add(second);
        section.Pages.Add(page);

        byte[] data = OneNoteSectionWriter.Write(section);

        OneNoteParagraph[] paragraphs = Assert.Single(Assert.Single(OneNoteSectionReader.Read(new MemoryStream(data)).Pages).Outlines)
            .Children.Cast<OneNoteParagraph>().ToArray();
        Assert.Equal(OneNoteParagraphAlignment.Left, paragraphs[0].Style.Alignment);
        Assert.Equal(OneNoteParagraphAlignment.Right, paragraphs[1].Style.Alignment);
        Assert.NotEqual(paragraphs[0].Style.ObjectId, paragraphs[1].Style.ObjectId);
    }

    [Fact]
    public void WriterRejectsTraversalOptionsAboveTheHardSafetyLimit() {
        OneNoteSection section = CreateOutlineChain(1);

        Assert.Throws<ArgumentOutOfRangeException>(() => OneNoteSectionWriter.Write(
            section,
            new OneNoteWriterOptions {
                MaxContentDepth = OneNoteWriterOptions.MaximumTraversalDepth + 1
            }));
    }

    private static OneNoteSection CreateConflictChain(int pageCount) {
        var section = new OneNoteSection { Name = "Conflicts" };
        var root = new OneNotePage { Title = "Page 0" };
        OneNotePage parent = root;
        for (int index = 1; index < pageCount; index++) {
            var child = new OneNotePage { Title = "Page " + index, IsConflictPage = true };
            parent.ConflictPages.Add(child);
            parent = child;
        }
        section.Pages.Add(root);
        return section;
    }

    private static OneNoteSection CreateOutlineChain(int elementCount) {
        var section = new OneNoteSection { Name = "Content" };
        var page = new OneNotePage { Title = "Page" };
        var root = new OneNoteOutline();
        OneNoteOutline parent = root;
        for (int index = 1; index < elementCount; index++) {
            var child = new OneNoteOutline();
            parent.Children.Add(child);
            parent = child;
        }
        page.Outlines.Add(root);
        section.Pages.Add(page);
        return section;
    }

    private static OneNoteSection CreateListedSection(int level) {
        var section = new OneNoteSection { Name = "Lists" };
        var page = new OneNotePage { Title = "Page" };
        var paragraph = new OneNoteParagraph {
            List = new OneNoteListInfo { Level = level }
        };
        paragraph.Runs.Add(new OneNoteTextRun { Text = "Item" });
        page.DirectContent.Add(paragraph);
        section.Pages.Add(page);
        return section;
    }

    private static OneNoteWriterOptions NoRoundTrip() => new OneNoteWriterOptions {
        ValidateRoundTrip = false
    };
}
