using System.Text;
using OfficeIMO.Epub;
using OfficeIMO.Html;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Tests;

public sealed class BookWorkbenchTests {
    [Fact]
    public async Task OpeningAProjectWithStylesDoesNotCreateDraftEdits() {
        using var storage = new StudioStorageAccess();
        var project = BookProject.Create("Styled book");
        project.SetStylesheet("p{color:navy}");
        var file = new TestStorageFile("content://books/styled-project", project.ToProjectBytes(), "book.oibook");
        await storage.RegisterAsync(file.Item, default);
        using var book = new BookWorkbenchViewModel(() => new Dialogs(), storage, null, () => Task.FromResult(UnsavedChangesDecision.Cancel));
        await book.OpenLocationAsync(file.Location.AbsoluteUri, default);
        Assert.Equal("p{color:navy}", book.Stylesheet);
        Assert.False(book.IsDirty);
    }
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImplicitDraftApplicationRefreshesChapterLabelsAfterSaveOrExport(bool export) {
        using var storage = new StudioStorageAccess();
        var source = new TestStorageFile("content://books/draft-source", Encoding.UTF8.GetBytes("# One\n\nFirst.\n\n# Two\n\nSecond."), "book.md");
        var output = new TestStorageFile("content://books/draft-output", [], export ? "book.epub" : "book.oibook");
        foreach (var file in new[] { source, output }) await storage.RegisterAsync(file.Item, default);
        var dialogs = new Dialogs { Save = output.Location.AbsoluteUri };
        using var book = new BookWorkbenchViewModel(() => dialogs, storage, null, () => Task.FromResult(UnsavedChangesDecision.Discard));
        await book.OpenLocationAsync(source.Location.AbsoluteUri, default);
        book.ChapterTitle = "Renamed opening";
        if (export) await book.ExportBookCommand.ExecuteAsync(null);
        else await book.SaveProjectCommand.ExecuteAsync(null);
        Assert.Equal(1, output.Writes);
        Assert.Equal("Renamed opening", book.Chapters[0].Title);
        book.SelectedChapter = book.Chapters[1];
        book.SelectedChapter = book.Chapters[0];
        Assert.Equal("Renamed opening", book.ChapterTitle);
        Assert.Equal("Renamed opening", book.Project!.Publication.Read().Chapters[0].Title);
    }
    [Fact]
    public async Task ChapterDraftsSurviveSelectionAndStylesheetUndoRefreshesTheEditor() {
        using var storage = new StudioStorageAccess();
        using var book = new BookWorkbenchViewModel(() => new Dialogs(), storage, null, () => Task.FromResult(UnsavedChangesDecision.Discard));
        await book.NewBookCommand.ExecuteAsync(null);
        await book.AddChapterCommand.ExecuteAsync(null);
        book.ChapterTitle = "Draft title";
        book.ChapterBody = "<body xmlns='http://www.w3.org/1999/xhtml'><h1>Draft title</h1><p>Draft body</p></body>";
        book.SelectedChapter = book.Chapters[0];
        book.SelectedChapter = book.Chapters[1];
        Assert.Equal("Draft title", book.ChapterTitle);
        Assert.Contains("Draft body", book.ChapterBody);
        await book.ApplyEditsCommand.ExecuteAsync(null);
        Assert.Equal("Draft title", book.Project!.Publication.Read().Chapters[1].Title);
        book.ChapterTitle = "Styled draft title";
        book.Stylesheet = "p{color:navy}";
        await book.ApplyEditsCommand.ExecuteAsync(null);
        await book.UndoEditCommand.ExecuteAsync(null);
        Assert.Equal(string.Empty, book.Stylesheet);
        Assert.Equal("Draft title", book.ChapterTitle);
        Assert.True(book.CanRedo);
        await book.RedoEditCommand.ExecuteAsync(null);
        Assert.Equal("p{color:navy}", book.Stylesheet);
        Assert.Equal("Styled draft title", book.ChapterTitle);
    }
    [Fact]
    public async Task ProviderManuscriptCanBeEditedSavedExportedAndReopened() {
        using var storage = new StudioStorageAccess();
        var source = new TestStorageFile("content://books/manuscript", Encoding.UTF8.GetBytes("---\ntitle: My book\n---\n# One\n\nFirst.\n\n# Two\n\nSecond."), "book.md");
        var saved = new TestStorageFile("content://books/project", [], "book.oibook");
        var output = new TestStorageFile("content://books/export", [], "book.epub");
        foreach (var file in new[] { source, saved, output }) await storage.RegisterAsync(file.Item, default);
        var dialogs = new Dialogs { Save = saved.Location.AbsoluteUri };
        using var book = new BookWorkbenchViewModel(() => dialogs, storage, null, () => Task.FromResult(UnsavedChangesDecision.Cancel));
        await book.OpenLocationAsync(source.Location.AbsoluteUri, default);
        Assert.Equal("My book", book.BookTitle); Assert.Equal(2, book.Chapters.Count);
        book.BookTitle = "Edited book";
        book.SelectedChapter = book.Chapters[1];
        await book.MoveUpCommand.ExecuteAsync(null);
        await book.SaveProjectCommand.ExecuteAsync(null);
        Assert.False(book.IsDirty, book.Status); Assert.Equal(1, saved.Writes);
        var restored = BookProject.LoadProject(saved.Bytes);
        Assert.Equal("Edited book", restored.Publication.Title);
        Assert.Equal("Two", restored.Publication.Read().Chapters[0].Title);
        dialogs.Save = output.Location.AbsoluteUri;
        await book.ExportBookCommand.ExecuteAsync(null);
        Assert.Equal(1, output.Writes);
        Assert.Equal("Edited book", EpubDocument.Load(new MemoryStream(output.Bytes)).Title);
        Assert.Equal(0, source.Writes);
    }

    [Fact]
    public async Task ChangedProjectAndSourceDestinationsAreProtectedBeforeWriting() {
        using var storage = new StudioStorageAccess();
        byte[] original = BookProject.Create("Original").ToProjectBytes();
        var file = new TestStorageFile("content://books/project", original, "book.oibook");
        await storage.RegisterAsync(file.Item, default);
        var dialogs = new Dialogs { Save = file.Location.AbsoluteUri };
        using var book = new BookWorkbenchViewModel(() => dialogs, storage, null, () => Task.FromResult(UnsavedChangesDecision.Cancel));
        await book.OpenLocationAsync(file.Location.AbsoluteUri, default);
        book.BookTitle = "Local edits";
        file.Bytes = BookProject.Create("External edits").ToProjectBytes();
        await book.SaveProjectCommand.ExecuteAsync(null);
        Assert.Equal(0, file.Writes); Assert.True(book.IsDirty);
        Assert.Equal("External edits", BookProject.LoadProject(file.Bytes).Publication.Title);
        await book.ExportBookCommand.ExecuteAsync(null);
        Assert.Equal(0, file.Writes);
    }

    [Fact]
    public async Task ReviewCannotOverrideFailedImportsAndInvalidDraftRetainsCurrentBook() {
        using var storage = new StudioStorageAccess();
        var file = new TestStorageFile("content://books/failure", Encoding.UTF8.GetBytes("<title>Broken</title><h1>One</h1><img alt='Missing' src='missing.png'>"), "book.html");
        await storage.RegisterAsync(file.Item, default);
        using var book = new BookWorkbenchViewModel(() => new Dialogs(), storage, null, () => Task.FromResult(UnsavedChangesDecision.Cancel));
        await book.OpenLocationAsync(file.Location.AbsoluteUri, default);
        Assert.False(book.CanExport); Assert.False(book.AcceptImportLossCommand.CanExecute(null));
        Assert.Contains(book.Diagnostics, item => item.StartsWith("EPUB_IMPORT_"));
        await book.NewBookCommand.ExecuteAsync(null);
        Assert.Equal("Broken", book.BookTitle);
        using var valid = new BookWorkbenchViewModel(() => new Dialogs(), storage, null, () => Task.FromResult(UnsavedChangesDecision.Discard));
        await valid.NewBookCommand.ExecuteAsync(null);
        byte[] before = valid.Project!.Export().Bytes;
        valid.ChapterBody = "<body xmlns='http://www.w3.org/1999/xhtml'><script>bad()</script></body>";
        await valid.ApplyEditsCommand.ExecuteAsync(null);
        Assert.Equal(before, valid.Project.Export().Bytes); Assert.True(valid.IsDirty);
    }

    [Fact]
    public async Task HomeWorkspaceBookPreventsUnconfirmedWindowCloseAndProtectsItsInput() {
        var services = TestAppBuilder.CreateTestServices();
        using var host = new StudioDocumentTabHost(open => new MainWindowViewModel(_ => Task.FromResult<string?>(null),
            openDocumentInTab: open, services: services, confirmBookChanges: () => Task.FromResult(UnsavedChangesDecision.Cancel)), _ => { });
        await host.ActiveDocument.BookWorkbench.NewBookCommand.ExecuteAsync(null);
        Assert.True(host.HasDirtyDocuments);
        Assert.False(await host.RequestCloseAllAsync());
        Assert.True(host.ActiveDocument.BookWorkbench.HasBook);
    }

    private sealed class Dialogs : IStudioFileDialogs {
        internal string? Save;
        public Task<string?> PickOpenFileAsync(string title, StudioFileType type, CancellationToken token) => Task.FromResult<string?>(null);
        public Task<string?> PickSaveFileAsync(string title, string suggestedName, StudioFileType type, CancellationToken token) => Task.FromResult(Save);
        public Task<string?> PickFolderAsync(string title, CancellationToken token) => Task.FromResult<string?>(null);
    }
}
