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
