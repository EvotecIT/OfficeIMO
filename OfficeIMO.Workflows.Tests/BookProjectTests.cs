using OfficeIMO.Epub;
using OfficeIMO.Html;
using System.Xml.Linq;

namespace OfficeIMO.Workflows.Tests;

public sealed class BookProjectTests {
    [Theory]
    [InlineData("{\"Version\":2,\"Diagnostics\":[]}", typeof(NotSupportedException))]
    [InlineData("{\"Version\":1,\"Diagnostics\":[],\"Unexpected\":true}", typeof(System.Text.Json.JsonException))]
    [InlineData("{\"Version\":1,\"Diagnostics\":[{\"Code\":\"TEST\",\"Message\":\"Finding\",\"Source\":\"HTML\",\"LossKind\":99}]}", typeof(InvalidDataException))]
    public void InvalidProjectReviewRecordsAreRejected(string review, Type expectedException) {
        using var bytes = new MemoryStream();
        using (var zip = new System.IO.Compression.ZipArchive(bytes, System.IO.Compression.ZipArchiveMode.Create, leaveOpen: true)) {
            using (var publication = zip.CreateEntry("publication.epub").Open()) publication.Write(BookProject.Create("Book").Export().Bytes);
            using var record = new StreamWriter(zip.CreateEntry("project.json").Open()); record.Write(review);
        }
        Assert.IsType(expectedException, Record.Exception(() => BookProject.LoadProject(bytes.ToArray())));
    }
    [Fact]
    public void ExistingProjectStylesheetPathsAreRetainedAndForeignIdentifierCollisionsAreAtomic() {
        var publication = EpubPublication.Create("Book", "en");
        publication.AddStylesheet("book-project-style", "EPUB/styles/custom style.css", "p{color:red}");
        publication.AddChapter("chapter", "EPUB/text/one.xhtml", "One", "<p>Body</p>");
        var project = BookProject.FromEpub(publication.Write().Bytes);
        var content = project.Publication.GetContentXml("chapter");
        XNamespace html = "http://www.w3.org/1999/xhtml";
        content.Root!.Element(html + "head")!.Add(new XElement(html + "base", new XAttribute("href", "../")));
        project.Publication.SetContentXml("chapter", content);
        project.SetStylesheet("p{color:blue}");
        Assert.Contains("custom%20style.css", project.Publication.GetContentXml("chapter").ToString());
        project.Export().Report.RequireNoLoss();
        var foreign = BookProject.Create("Other");
        foreign.Publication.AddResource("book-project-style", "EPUB/data.bin", "application/octet-stream", [1, 2, 3]);
        byte[] before = foreign.Export().Bytes;
        Assert.Throws<InvalidDataException>(() => foreign.SetStylesheet("p{color:red}"));
        Assert.Equal(before, foreign.Export().Bytes);
    }
    [Fact]
    public void ChapterInsertionRemovalAndUndoRetainTheBookAndStyles() {
        var project = BookProject.Create("Book");
        project.SetStylesheet("p{color:navy}");
        string id = project.AddChapter("Second");
        Assert.Equal(2, project.Publication.Spine.Count);
        Assert.Contains("project.css", project.Publication.GetContentXml(id).ToString());
        project.RemoveChapter(1);
        Assert.Single(project.Publication.Spine);
        project.Undo();
        Assert.Equal(2, project.Publication.Spine.Count);
        Assert.Equal("Second", project.Publication.Read().Chapters[1].Title);
        project.Redo();
        Assert.Single(project.Publication.Spine);
        Assert.Throws<InvalidOperationException>(() => project.RemoveChapter(0));
    }
    [Fact]
    public void EditorBatchIsAtomicAndDeletionCannotBreakRetainedChapterLinks() {
        var project = Source();
        byte[] before = project.Export().Bytes;
        Assert.Throws<InvalidDataException>(() => project.RemoveChapter(1));
        Assert.Equal(before, project.Export().Bytes);
        Assert.ThrowsAny<Exception>(() => project.ApplyEdits(new BookProjectEdits {
            Title = "Changed", ChapterBodies = new Dictionary<string, string> { ["chapter-2"] = "<!DOCTYPE body [<!ENTITY x 'Unsafe'>]><body xmlns='http://www.w3.org/1999/xhtml'>&x;</body>" }
        }));
        Assert.Equal(before, project.Export().Bytes);
        project.ApplyEdits(new BookProjectEdits { Title = "Updated", ChapterTitles = new Dictionary<string, string> { ["chapter-1"] = "First" } });
        Assert.Equal("Updated", project.Publication.Title);
        Assert.Equal("First", project.Publication.Read().TableOfContents[0].Label);
        project.Undo();
        Assert.Equal("Book", project.Publication.Title);
    }
    [Fact]
    public void ProjectSaveReopensImportedPublicationAndRetainsReviewAcceptance() {
        var imported = EpubManuscript.ImportHtml(HtmlConversionDocument.Parse("<title>Book</title><h1>One</h1><p>Text</p><script>ignored()</script><h1>Two</h1><p>Final</p>"));
        BookProject project = BookProject.FromImport(imported);
        Assert.False(project.CanExport);
        Assert.Throws<InvalidOperationException>(() => project.Export());
        project.AcknowledgeImportLoss();
        BookProject restored = BookProject.LoadProject(project.ToProjectBytes());
        Assert.True(restored.CanExport);
        Assert.True(restored.ImportLossAcknowledged);
        Assert.Contains(restored.ImportDiagnostics, item => item.Code == "EPUB_IMPORT_ACTIVE_CONTENT_OMITTED");
        Assert.Equal(2, EpubDocument.Load(new MemoryStream(restored.Export().Bytes)).Chapters.Count);
    }

    [Fact]
    public void ChapterChangesReorderNavigationAndRejectInvalidEditsWithoutChangingTheBook() {
        var project = Source();
        project.MoveChapter(1, 0);
        project.RenameChapter(0, "New title");
        var reopened = EpubDocument.Load(new MemoryStream(project.Export().Bytes));
        Assert.Equal(new[] { "New title", "One" }, reopened.Chapters.Select(item => item.Title));
        Assert.Equal(new[] { "New title", "One" }, reopened.TableOfContents.Select(item => item.Label));
        byte[] before = project.Export().Bytes;
        Assert.ThrowsAny<Exception>(() => project.MoveChapter(0, 99));
        Assert.Equal(before, project.Export().Bytes);
        Assert.Throws<OperationCanceledException>(() => project.RenameChapter(0, "Cancelled", new CancellationToken(true)));
        Assert.Equal(before, project.Export().Bytes);
    }

    [Fact]
    public void ProjectStyleAndCoverRemainDeclaredAfterSaveAndExport() {
        var project = Source();
        project.SetStylesheet("body{font-family:serif;line-height:1.6}");
        const string png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";
        project.SetCoverImage(Convert.FromBase64String(png), "image/png");
        var restored = BookProject.LoadProject(project.ToProjectBytes());
        var reopened = restored.Publication.Read(new EpubReadOptions { IncludeRawHtml = true, IncludeResourceData = true });
        Assert.Contains(reopened.Resources, item => (item.Properties ?? string.Empty).Split(' ').Contains("cover-image"));
        Assert.Contains(restored.Publication.Manifest, item => item.Id == "book-project-style");
        Assert.All(restored.Publication.Spine, chapter => Assert.Contains("project.css", restored.Publication.GetContentXml(chapter.ManifestId).ToString()));
        restored.Export().Report.RequireNoLoss();
    }

    [Fact]
    public void FailedImportCannotBeAcceptedAsReadyForExport() {
        var project = BookProject.FromImport(EpubManuscript.ImportHtml(HtmlConversionDocument.Parse("<title>Broken</title><h1>One</h1><a href='#missing'>Missing</a>")));
        Assert.False(project.CanExport);
        Assert.Throws<InvalidOperationException>(project.AcknowledgeImportLoss);
        Assert.Throws<InvalidOperationException>(() => project.Export());
    }
    private static BookProject Source() => BookProject.FromImport(EpubManuscript.ImportHtml(HtmlConversionDocument.Parse(
        "<title>Book</title><h1 id='one'>One</h1><p><a href='#two'>Next</a></p><h1 id='two'>Two</h1><p><a href='#one'>Back</a></p>")));
}
