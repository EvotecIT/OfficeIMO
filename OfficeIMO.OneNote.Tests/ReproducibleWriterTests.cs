namespace OfficeIMO.OneNote.Tests;

public sealed class ReproducibleWriterTests {
    [Theory]
    [InlineData(OneNoteStorageFormat.RevisionStore)]
    [InlineData(OneNoteStorageFormat.FileSynchronizationPackage)]
    public void ExplicitIdentityNamespaceAndTimeReproduceIndependentAndRepeatedSections(OneNoteStorageFormat storageFormat) {
        var options = Options(storageFormat);
        using OneNoteSection section = CreateSection();
        byte[] first = section.ToBytes(options);
        Assert.Equal(first, section.ToBytes(options));
        using OneNoteSection independent = CreateSection();
        Assert.Equal(first, independent.ToBytes(options));

        var anotherNamespace = Options(storageFormat);
        anotherNamespace.IdentitySeed = Guid.Parse("10000000-0000-0000-0000-000000000002");
        using OneNoteSection separate = CreateSection();
        Assert.NotEqual(first, separate.ToBytes(anotherNamespace));

        using OneNoteSection reopened = OneNoteSection.Load(new MemoryStream(first));
        OneNotePage page = Assert.Single(reopened.Pages);
        Assert.Equal(section.Id, reopened.Id);
        Assert.Equal(section.Pages[0].Id, page.Id);
        Assert.Equal("Reproducible page", page.Title);
        Assert.Equal(options.GeneratedTimestampUtc, page.CreatedUtc);
        OneNoteParagraph paragraph = Assert.Single(page.Outlines.SelectMany(outline => outline.Children)
            .Concat(page.DirectContent).OfType<OneNoteParagraph>());
        Assert.Equal("Stable content", Assert.Single(paragraph.Runs).Text);
    }

    [Theory]
    [InlineData(OneNoteStorageFormat.RevisionStore)]
    [InlineData(OneNoteStorageFormat.FileSynchronizationPackage)]
    public void ExplicitIdentityNamespaceReproducesNotebookHierarchy(OneNoteStorageFormat storageFormat) {
        OneNoteNotebook Create() {
            var notebook = new OneNoteNotebook { Name = "Notebook" };
            notebook.Sections.Add(CreateSection());
            var group = new OneNoteSectionGroup { Name = "Group" };
            group.Sections.Add(new OneNoteSection { Name = "Nested" });
            notebook.SectionGroups.Add(group);
            return notebook;
        }
        OneNoteNotebook notebook = Create();
        byte[] first = OneNoteTableOfContentsWriter.Write(notebook, Options(storageFormat));
        Assert.Equal(first, OneNoteTableOfContentsWriter.Write(notebook, Options(storageFormat)));
        Assert.Equal(first, OneNoteTableOfContentsWriter.Write(Create(), Options(storageFormat)));
        OneNoteNotebook reopened = OneNoteNotebookReader.Read(new MemoryStream(first));
        Assert.Equal("Reproducible section", Assert.Single(reopened.Sections).Name);
        Assert.Equal("Group", Assert.Single(reopened.SectionGroups).Name);
        Assert.Equal(notebook.Id, reopened.Id);
    }

    [Theory]
    [InlineData(OneNoteStorageFormat.RevisionStore)]
    [InlineData(OneNoteStorageFormat.FileSynchronizationPackage)]
    public void RichCreatedSectionsReproduceIndependentPayloadsAndRepeatedWrites(OneNoteStorageFormat storageFormat) {
        using OneNoteSection section = SectionWriterTests.CreateRichSection();
        byte[] first = section.ToBytes(Options(storageFormat));
        Assert.Equal(first, section.ToBytes(Options(storageFormat)));
        using OneNoteSection independent = SectionWriterTests.CreateRichSection();
        Assert.Equal(first, independent.ToBytes(Options(storageFormat)));
        using OneNoteSection reopened = OneNoteSection.Load(new MemoryStream(first));
        var outline = Assert.Single(Assert.Single(reopened.Pages).Outlines);
        Assert.Equal(new byte[] { 1, 2, 3, 4 }, Assert.Single(outline.Children.OfType<OneNoteImage>()).Payload!.ToArray(16));
        Assert.Equal(new byte[] { 5, 6, 7 }, Assert.Single(outline.Children.OfType<OneNoteEmbeddedFile>()).Payload!.ToArray(16));
        Assert.Equal("Cell", Assert.Single(Assert.Single(Assert.Single(outline.Children.OfType<OneNoteTable>()).Rows).Cells).Content.OfType<OneNoteParagraph>().Single().Runs.Single().Text);
    }

    [Theory]
    [InlineData(OneNoteStorageFormat.RevisionStore)]
    [InlineData(OneNoteStorageFormat.FileSynchronizationPackage)]
    public void SeededSectionsRetainDistinctIdentitiesAfterInsertingContent(OneNoteStorageFormat storageFormat) {
        using var section = SectionWriterTests.CreateRichSection();
        var options = Options(storageFormat);
        section.ToBytes(options);
        var original = Assert.IsType<OneNoteParagraph>(section.Pages[0].Outlines[0].Children[0]);
        var originalId = original.ContentObjectId;
        var inserted = new OneNoteParagraph(); inserted.Runs.Add(new OneNoteTextRun { Text = "Inserted" });
        section.Pages[0].Outlines[0].Children.Insert(0, inserted);
        byte[] changed = section.ToBytes(options);
        Assert.Equal(originalId, original.ContentObjectId);
        Assert.NotEqual(original.ContentObjectId, inserted.ContentObjectId);
        Assert.Equal(changed, section.ToBytes(options));
        using var reopened = OneNoteSection.Load(new MemoryStream(changed));
        var paragraphs = reopened.Pages[0].Outlines[0].Children.OfType<OneNoteParagraph>().SelectMany(paragraph => paragraph.Runs).Select(run => run.Text).ToArray();
        Assert.Contains("Inserted", paragraphs); Assert.Contains("List item", paragraphs);
    }

    [Theory]
    [InlineData(OneNoteStorageFormat.RevisionStore)]
    [InlineData(OneNoteStorageFormat.FileSynchronizationPackage)]
    public void SeededNotebookEditsKeepRetainedAndNewFileIdentitiesDistinct(OneNoteStorageFormat storageFormat) {
        var notebook = new OneNoteNotebook { Name = "Notebook" };
        notebook.Sections.Add(new OneNoteSection { Name = "A" });
        var retained = new OneNoteSection { Name = "B" }; notebook.Sections.Add(retained);
        var options = Options(storageFormat);
        OneNoteTableOfContentsWriter.Write(notebook, options);
        var retainedId = retained.Id;
        notebook.Sections.RemoveAt(0);
        var added = new OneNoteSection { Name = "C" }; notebook.Sections.Add(added);
        byte[] changed = OneNoteTableOfContentsWriter.Write(notebook, options);
        Assert.Equal(retainedId, retained.Id); Assert.NotEqual(retained.Id, added.Id);
        Assert.Equal(changed, OneNoteTableOfContentsWriter.Write(notebook, options));
        var reopened = OneNoteNotebookReader.Read(new MemoryStream(changed));
        Assert.Equal(new[] { "B", "C" }, reopened.Sections.Select(section => section.Name));
        Assert.Equal(2, reopened.Sections.Select(section => section.Id).Distinct().Count());
    }

    [Fact]
    public void ReproducibleWritingRetainsCallerAssignedIdentitiesAndTimes() {
        using OneNoteSection section = CreateSection();
        Guid sectionId = Guid.Parse("20000000-0000-0000-0000-000000000001");
        var pageId = new OneNoteExtendedGuid(Guid.Parse("20000000-0000-0000-0000-000000000002"), 1, 17);
        DateTime created = new DateTime(2012, 3, 4, 5, 6, 7, DateTimeKind.Utc);
        section.Id = sectionId;
        section.Pages[0].Id = pageId;
        section.Pages[0].CreatedUtc = created;
        using OneNoteSection reopened = OneNoteSection.Load(new MemoryStream(section.ToBytes(Options(OneNoteStorageFormat.RevisionStore))));
        Assert.Equal(sectionId, reopened.Id);
        Assert.Equal(pageId, reopened.Pages[0].Id);
        Assert.Equal(created, reopened.Pages[0].CreatedUtc);
    }

    [Fact]
    public void ReproducibleWritingRequiresExplicitUtcClockInputBeforeWriting() {
        using OneNoteSection section = CreateSection();
        using var output = new MemoryStream();
        var options = new OneNoteSaveOptions { IdentitySeed = Guid.Empty };
        Assert.Throws<ArgumentException>(() => section.Save(output, options));
        Assert.Equal(0, output.Length);
        options.GeneratedTimestampUtc = new DateTime(2000, 1, 1, 0, 0, 0, DateTimeKind.Local);
        Assert.Throws<ArgumentException>(() => section.Save(output, options));
        Assert.Equal(0, output.Length);
    }

    private static OneNoteSaveOptions Options(OneNoteStorageFormat format) => new OneNoteSaveOptions {
        StorageFormat = format,
        IdentitySeed = Guid.Parse("10000000-0000-0000-0000-000000000001"),
        GeneratedTimestampUtc = new DateTime(2000, 1, 1, 0, 0, 0, DateTimeKind.Utc)
    };

    private static OneNoteSection CreateSection() {
        var section = new OneNoteSection { Name = "Reproducible section" };
        var page = new OneNotePage { Title = "Reproducible page" };
        var paragraph = new OneNoteParagraph();
        paragraph.Runs.Add(new OneNoteTextRun { Text = "Stable content" });
        page.DirectContent.Add(paragraph);
        section.Pages.Add(page);
        return section;
    }
}
