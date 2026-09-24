using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Sign;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioDocumentStructureTests {
    [Fact]
    public async Task PropertiesBookmarksAndAttachmentsAreRecordedAsUndoableEdits() {
        string root = CreateRoot();
        string path = CreateDocument(root);
        string attachmentSource = Path.Combine(root, "notes.txt");
        File.WriteAllText(attachmentSource, "Review notes");
        try {
            var dialogs = new RecordingFileDialogs();
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null)) { FileDialogs = dialogs };
            await model.OpenDocumentAsync(path);

            model.PropertyTitle = "Quarterly report";
            model.PropertyAuthor = "Studio";
            Assert.True(model.HasPropertyChanges);
            await model.ApplyPropertiesCommand.ExecuteAsync(null);
            Assert.Null(model.ErrorMessage);
            Assert.False(model.HasPropertyChanges);
            Assert.Equal("Quarterly report", model.PropertyTitle);
            Assert.True(model.IsDirty);
            Assert.Equal("Document properties updated", model.OperationStatus);

            model.SelectedPage = model.Pages[1];
            await model.AddBookmarkCommand.ExecuteAsync(null);
            var added = Assert.Single(model.Bookmarks);
            Assert.Equal(2, added.PageNumber);
            Assert.Equal(added, model.SelectedBookmark);

            model.BookmarkTitleDraft = "Results";
            await model.RenameBookmarkCommand.ExecuteAsync(null);
            Assert.Equal("Results", Assert.Single(model.Bookmarks).Title);

            model.SelectedPage = model.Pages[0];
            await model.AddBookmarkCommand.ExecuteAsync(null);
            Assert.Equal(["Results", "Page 1"], model.Bookmarks.Select(bookmark => bookmark.Title));
            model.SelectedBookmark = model.Bookmarks[1];
            await model.IndentBookmarkCommand.ExecuteAsync(null);
            Assert.Equal(model.Bookmarks[0].Level + 1, model.Bookmarks[1].Level);
            Assert.Equal(model.Bookmarks[0].Id, model.Bookmarks[1].ParentId);

            dialogs.OpenFile = attachmentSource;
            await model.AddAttachmentCommand.ExecuteAsync(null);
            var attachment = Assert.Single(model.DocumentAttachments);
            Assert.Equal("notes.txt", attachment.FileName);

            string saved = Path.Combine(root, "saved-notes.txt");
            dialogs.SaveFile = saved;
            await model.SaveAttachmentCommand.ExecuteAsync(attachment);
            Assert.Equal("Review notes", File.ReadAllText(saved));

            await model.RemoveAttachmentCommand.ExecuteAsync(attachment);
            Assert.Empty(model.DocumentAttachments);
            await model.UndoCommand.ExecuteAsync(null);
            Assert.Single(model.DocumentAttachments);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task HeaderFooterTextIsStampedAndExportedWithTheDocument() {
        string root = CreateRoot();
        string path = CreateDocument(root);
        try {
            var dialogs = new RecordingFileDialogs();
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null)) { FileDialogs = dialogs };
            await model.OpenDocumentAsync(path);

            model.HeaderLeft = "{file}";
            model.FooterCenter = "Sheet {page} of {pages}";
            await model.ApplyHeaderFooterCommand.ExecuteAsync(null);
            Assert.Null(model.ErrorMessage);
            Assert.Equal("Header and footer added", model.OperationStatus);
            Assert.True(model.CanUndoOperationStatus);

            string text = Path.Combine(root, "structure.txt");
            dialogs.SaveFile = text;
            await model.ExportDocumentCommand.ExecuteAsync("Text");
            string exported = File.ReadAllText(text);
            Assert.Contains("Sheet 1 of 2", exported);
            Assert.Contains("Sheet 2 of 2", exported);
            Assert.Contains("structure", exported);
            Assert.Equal("Exported to structure.txt", model.OperationStatus);
            Assert.True(model.CanUndo);
            Assert.False(model.CanUndoOperationStatus);

            string markdown = Path.Combine(root, "structure.md");
            dialogs.SaveFile = markdown;
            await model.ExportDocumentCommand.ExecuteAsync("Markdown");
            Assert.Contains("Second page body", File.ReadAllText(markdown));
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task TypedSignatureIsSavedForReuseAndPlacedWithOneClick() {
        string root = CreateRoot();
        string path = CreateDocument(root);
        try {
            using var session = TestAppBuilder.StartSession();
            await session.Dispatch(async () => {
                var dialog = new SignatureDialog(StudioSignatureKind.Signature, StudioLocalization.Current);
                dialog.NameBox.Text = "Ada Lovelace";
                byte[] png = dialog.CreateImage()!;
                Assert.NotNull(png);
                Assert.Equal(0x89, png[0]);

                var services = StudioApplicationServices.Create(new StudioDataPaths(Path.Combine(root, "profile")));
                using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services) {
                    CreateSignatureDialog = _ => Task.FromResult<StudioSignatureDraft?>(new StudioSignatureDraft(png, Remember: true))
                };
                await model.OpenDocumentAsync(path);
                await model.CreateSignatureCommand.ExecuteAsync("Signature");
                Assert.Equal(StudioDocumentMode.Forms, model.DocumentMode);
                Assert.True(model.IsPlacingFillSignItem);
                Assert.Equal(PdfEditorTool.AddImage, model.ActiveEditorTool);
                Assert.Single(model.SavedSignatures);
                Assert.Single(services.Signatures.List(StudioSignatureKind.Signature));

                // A plain click arrives as the canvas default box; the placement keeps the image proportions.
                model.Pages[1].CompleteEditorGesture(new PdfEditorGesture(2, 60D, 300D, 220D, 400D, [new PdfEditorVisualPoint(60D, 300D)]));
                using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(10));
                while (!model.CanUndo) await Task.Delay(10, timeout.Token);
                Assert.False(model.HasError, model.ErrorMessage);
                Assert.False(model.IsPlacingFillSignItem);
                Assert.Equal(PdfEditorTool.Select, model.ActiveEditorTool);
                Assert.Equal("Signature added", model.OperationStatus);

                model.PlaceDateCommand.Execute(null);
                Assert.Equal(PdfEditorTool.AddText, model.ActiveEditorTool);
                model.CancelFillSignPlacementCommand.Execute(null);
                Assert.False(model.IsPlacingFillSignItem);
                Assert.Equal(PdfEditorTool.Select, model.ActiveEditorTool);
                return true;
            }, CancellationToken.None);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task OversizeSignatureStoreFailureStillBeginsOneTimePlacement() {
        string root = CreateRoot();
        string path = CreateDocument(root);
        try {
            using var session = TestAppBuilder.StartSession();
            await session.Dispatch(async () => {
                byte[] png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAMAAAABCAYAAAAb4BS0AAAAEElEQVR4nGNgYGD4z8DAwAAABQABnEX0RwAAAABJRU5ErkJggg==");
                Array.Resize(ref png, 4 * 1024 * 1024 + 1);
                var services = StudioApplicationServices.Create(new StudioDataPaths(Path.Combine(root, "profile")));
                using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services) {
                    CreateSignatureDialog = _ => Task.FromResult<StudioSignatureDraft?>(new StudioSignatureDraft(png, Remember: true))
                };
                await model.OpenDocumentAsync(path);
                await model.CreateSignatureCommand.ExecuteAsync("Signature");
                Assert.True(model.IsPlacingFillSignItem);
                Assert.Empty(model.SavedSignatures);
                Assert.NotNull(model.ErrorMessage);
                return true;
            }, CancellationToken.None);
        } finally { Directory.Delete(root, recursive: true); }
    }

    [Fact]
    public async Task ScannedPageShortcutRequiresSavedSourceBeforeOpeningOcr() {
        string root = CreateRoot();
        string path = CreateDocument(root);
        try {
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
            await model.OpenDocumentAsync(path);
            model.SetOrganizerSelection([model.OrganizerPages[0]]);
            await model.DuplicateSelectedCommand.ExecuteAsync(null);
            Assert.True(model.IsDirty);

            await model.MakeSearchableCommand.ExecuteAsync(null);
            Assert.True(model.IsPdfWorkspaceMode);
            Assert.Contains("Save your current changes", model.ErrorMessage);
        } finally { Directory.Delete(root, recursive: true); }
    }

    [Fact]
    public async Task LockedSavedSignatureRemainsVisibleWhenDeleteFails() {
        if (!OperatingSystem.IsWindows()) return;
        string root = CreateRoot();
        string path = CreateDocument(root);
        try {
            using var session = TestAppBuilder.StartSession();
            await session.Dispatch(async () => {
                byte[] png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAMAAAABCAYAAAAb4BS0AAAAEElEQVR4nGNgYGD4z8DAwAAABQABnEX0RwAAAABJRU5ErkJggg==");
                var services = StudioApplicationServices.Create(new StudioDataPaths(Path.Combine(root, "profile")));
                using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services) {
                    CreateSignatureDialog = _ => Task.FromResult<StudioSignatureDraft?>(new StudioSignatureDraft(png, Remember: true))
                };
                await model.OpenDocumentAsync(path);
                await model.CreateSignatureCommand.ExecuteAsync("Signature");
                var saved = Assert.Single(model.SavedSignatures);
                using (var locked = new FileStream(saved.Saved.Path, FileMode.Open, FileAccess.Read, FileShare.None)) {
                    model.DeleteSavedSignatureCommand.Execute(saved);
                    Assert.Single(model.SavedSignatures);
                    Assert.NotNull(model.ErrorMessage);
                    Assert.True(File.Exists(saved.Saved.Path));
                }
                model.DeleteSavedSignatureCommand.Execute(saved);
                Assert.Empty(model.SavedSignatures);
                Assert.False(File.Exists(saved.Saved.Path));
                return true;
            }, CancellationToken.None);
        } finally { Directory.Delete(root, recursive: true); }
    }

    [Fact]
    public void SavedSignatureAndShapeAreOwnerOnlyOnUnix() {
        if (OperatingSystem.IsWindows()) return;
        string root = CreateRoot();
        try {
            var store = new StudioSignatureStore(Path.Combine(root, "signatures"));
            var saved = store.Save(StudioSignatureKind.Signature, [1, 2, 3], text: "Ada");
            string shape = Path.ChangeExtension(saved.Path, ".json");
            UnixFileMode ownerOnly = UnixFileMode.UserRead | UnixFileMode.UserWrite;
            Assert.Equal(ownerOnly, File.GetUnixFileMode(saved.Path));
            Assert.Equal(ownerOnly, File.GetUnixFileMode(shape));
            File.SetUnixFileMode(saved.Path, ownerOnly | UnixFileMode.GroupRead | UnixFileMode.OtherRead);
            File.SetUnixFileMode(shape, ownerOnly | UnixFileMode.GroupRead | UnixFileMode.OtherRead);
            Assert.Single(store.List(StudioSignatureKind.Signature));
            Assert.Equal(ownerOnly, File.GetUnixFileMode(saved.Path));
            Assert.Equal(ownerOnly, File.GetUnixFileMode(shape));
        } finally { Directory.Delete(root, recursive: true); }
    }

    [Fact]
    public async Task DrawnSignatureOnAFormOnlyPdfIsPlacedAsInk() {
        string root = CreateRoot();
        string path = Path.Combine(root, "form.pdf");
        PdfDocument.Create(compose => compose.Page(page => page.Content(content =>
            content.Item(item => item.TextField("Name", value: "Ada"))))).Save(path);
        try {
            using var session = TestAppBuilder.StartSession();
            await session.Dispatch(async () => {
                var services = StudioApplicationServices.Create(new StudioDataPaths(Path.Combine(root, "profile")));
                IReadOnlyList<IReadOnlyList<Avalonia.Point>> strokes = [[new(0.1, 0.8), new(0.4, 0.2), new(0.9, 0.7)]];
                byte[] png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAMAAAABCAYAAAAb4BS0AAAAEElEQVR4nGNgYGD4z8DAwAAABQABnEX0RwAAAABJRU5ErkJggg==");
                using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services) {
                    CreateSignatureDialog = _ => Task.FromResult<StudioSignatureDraft?>(new StudioSignatureDraft(png, Remember: false, Strokes: strokes))
                };
                await model.OpenDocumentAsync(path);
                Assert.False(model.CanEditPageContent);
                Assert.True(model.CanFillAndSign);

                await model.CreateSignatureCommand.ExecuteAsync("Signature");
                model.Pages[0].CompleteEditorGesture(new PdfEditorGesture(1, 60D, 300D, 220D, 400D, [new PdfEditorVisualPoint(60D, 300D)]));
                using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(10));
                while (!model.CanUndo) await Task.Delay(10, timeout.Token);
                Assert.False(model.HasError, model.ErrorMessage);
                Assert.False(model.IsPlacingFillSignItem);
                return true;
            }, CancellationToken.None);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    private static string CreateRoot() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-structure-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        return root;
    }

    private static string CreateDocument(string root) {
        string path = Path.Combine(root, "structure.pdf");
        PdfDocument.Create(compose => {
            compose.Page(page => page.Size(400D, 500D).Content(content => content.Text("First page body")));
            compose.Page(page => page.Size(400D, 500D).Content(content => content.Text("Second page body")));
        }).Save(path);
        return path;
    }

    private sealed class RecordingFileDialogs : IStudioFileDialogs {
        public string? OpenFile { get; set; }
        public string? SaveFile { get; set; }
        public string? Folder { get; set; }

        public Task<string?> PickOpenFileAsync(string title, StudioFileType type, CancellationToken cancellationToken) => Task.FromResult(OpenFile);
        public Task<string?> PickSaveFileAsync(string title, string suggestedName, StudioFileType type, CancellationToken cancellationToken) => Task.FromResult(SaveFile);
        public Task<string?> PickFolderAsync(string title, CancellationToken cancellationToken) => Task.FromResult(Folder);
    }
}
