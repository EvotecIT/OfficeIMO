using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.VisualTree;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Workspace;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioCommentIdentityTests {
    [Fact]
    public async Task RestrictedDocumentAllowsReviewingThreadsButRejectsCommentMutations() {
        await WithDocument(async model => {
            Assert.True(model.HasDocument, model.ErrorMessage);
            Assert.Equal(2, model.CommentThreads.Count);
            model.NextUnresolvedCommentCommand.Execute(null);
            model.CommentReplyText = "A draft must not bypass document permissions.";
            Assert.False(model.ReplyToCommentCommand.CanExecute(null));
            Assert.False(model.ResolveCommentCommand.CanExecute(null));
            await model.ReplyToCommentCommand.ExecuteAsync(null);
            await model.ResolveCommentCommand.ExecuteAsync(null);
            Assert.False(model.IsDirty);
            Assert.Equal(2, model.SelectedCommentThread!.Entries.Count);
        }, restricted: true);
    }

    [Fact]
    public async Task DuplicateAnnotationNamesKeepSeparateReplyDrafts() {
        await WithDocument(async model => {
            var first = model.CommentThreads[0];
            var second = model.CommentThreads[1];
            model.SelectedCommentThread = first;
            model.CommentReplyText = "First draft";
            model.SelectedCommentThread = second;
            Assert.Empty(model.CommentReplyText);
            model.CommentReplyText = "Second draft";
            model.SelectedCommentThread = first;
            Assert.Equal("First draft", model.CommentReplyText);
        }, duplicateNames: true);
    }

    [Fact]
    public async Task UnnamedThreadKeepsItsDraftAfterContentEditAndPageReorder() {
        await WithDocument(async model => {
            model.SelectedCommentThread = model.CommentThreads[0];
            model.CommentReplyText = "Keep this draft";
            var annotation = model.SelectedCommentThread.Annotation;
            model.Pages[0].SelectObject(new(PdfEditorSelectionKind.Annotation, 1, new(70, 700, 88, 718), ObjectNumber: annotation.ObjectNumber, Subtype: "Text"));
            model.SelectedAnnotationContents = "Updated delivery comment";
            await model.UpdateSelectedAnnotationCommand.ExecuteAsync(null);
            Assert.Null(model.ErrorMessage);
            Assert.Equal("Updated delivery comment", model.SelectedCommentThread?.Contents);
            Assert.Equal("Keep this draft", model.CommentReplyText);
            await model.ReorderByDropAsync(2, 1);
            Assert.Null(model.ErrorMessage);
            Assert.Equal("Updated delivery comment", model.SelectedCommentThread?.Contents);
            Assert.Equal(2, model.SelectedCommentThread!.Annotation.PageNumber);
            Assert.Equal("Keep this draft", model.CommentReplyText);
        });
    }

    [Fact]
    public async Task RemovedCommentDraftRemainsVisibleAndCanBeRestoredDeliberately() {
        await WithDocument(async model => {
            model.ShowAnnotateModeCommand.Execute(null);
            model.SelectedCommentThread = model.CommentThreads[1];
            model.CommentReplyText = "Retain this reply even when the comment is deleted.";
            var annotation = model.SelectedCommentThread.Annotation;
            model.Pages[1].SelectObject(new(PdfEditorSelectionKind.Annotation, 2, new(70, 700, 88, 718), ObjectNumber: annotation.ObjectNumber, Subtype: "Text"));
            await model.DeleteSelectedObjectCommand.ExecuteAsync(null);
            Assert.Null(model.ErrorMessage);
            Assert.Single(model.UnassignedCommentDrafts);
            model.SelectedCommentThread = model.CommentThreads[0];
            var window = new Window { Width = 960, Height = 620, Content = new DocumentWorkspaceView { DataContext = model } };
            try {
                window.Show(); window.UpdateLayout();
                var restore = window.GetVisualDescendants().OfType<Button>().Single(button => ReferenceEquals(button.Command, model.RestoreCommentDraftCommand));
                restore.BringIntoView(); window.UpdateLayout();
                Capture(window, "comments-retained-draft.png");
                Assert.True(restore.IsEnabled);
                model.RestoreCommentDraftCommand.Execute(null);
                Assert.Equal("Retain this reply even when the comment is deleted.", model.CommentReplyText);
                Assert.Empty(model.UnassignedCommentDrafts);
            } finally { window.Close(); }
        });
    }

    [Fact]
    public async Task NextUnresolvedRevealsTheSameAnchorAfterScrollingAway() {
        await WithDocument(async model => {
            model.ShowAnnotateModeCommand.Execute(null);
            model.SelectedCommentThread = model.CommentThreads[1];
            await model.ResolveCommentCommand.ExecuteAsync(null);
            model.NextUnresolvedCommentCommand.Execute(null);
            var window = new Window { Width = 960, Height = 620, Content = new DocumentWorkspaceView { DataContext = model } };
            try {
                window.Show(); window.UpdateLayout();
                await model.SelectedPage!.EnsureRenderedAsync();
                await Avalonia.Threading.Dispatcher.UIThread.InvokeAsync(() => window.UpdateLayout(), Avalonia.Threading.DispatcherPriority.Background);
                var canvas = window.GetVisualDescendants().OfType<OfficeIMO.Studio.Features.Reader.PdfPageCanvas>()
                    .Single(canvas => canvas.DataContext is PdfPageViewModel && canvas.Scene?.PageNumber == 1);
                var scroll = canvas.GetVisualAncestors().OfType<ScrollViewer>().First();
                scroll.Offset = default;
                window.UpdateLayout();
                await Avalonia.Threading.Dispatcher.UIThread.InvokeAsync(() => window.UpdateLayout(), Avalonia.Threading.DispatcherPriority.Background);
                Assert.Equal(0, scroll.Offset.Y);
                model.NextUnresolvedCommentCommand.Execute(null);
                await Avalonia.Threading.Dispatcher.UIThread.InvokeAsync(() => window.UpdateLayout(), Avalonia.Threading.DispatcherPriority.Background);
                Assert.True(scroll.Offset.Y > 100, $"Comment anchor was not revealed: scroll offset {scroll.Offset.Y}.");
                Capture(window, "comments-repeat-anchor.png");
            } finally { window.Close(); }
        });
    }

    private static void Capture(Window window, string name) {
        using var frame = window.CaptureRenderedFrame();
        Assert.NotNull(frame);
        string? path = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(path)) return;
        Directory.CreateDirectory(path);
        frame.Save(Path.Combine(path, name), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }

    [Fact]
    public async Task TypingAReplyHasABoundedAllocationCostAfterOpeningTheDocument() {
        await WithDocument(model => {
            model.NextUnresolvedCommentCommand.Execute(null);
            model.ReplyToCommentCommand.CanExecuteChanged += (_, _) => _ = model.ReplyToCommentCommand.CanExecute(null);
            model.ResolveCommentCommand.CanExecuteChanged += (_, _) => _ = model.ResolveCommentCommand.CanExecute(null);
            model.ReopenCommentCommand.CanExecuteChanged += (_, _) => _ = model.ReopenCommentCommand.CanExecute(null);
            model.CommentReplyText = "Warm up";
            long start = GC.GetAllocatedBytesForCurrentThread();
            for (int index = 0; index < 32; index++) model.CommentReplyText = "A review draft " + index;
            long allocated = GC.GetAllocatedBytesForCurrentThread() - start;
            Assert.True(allocated < 5 * 1024 * 1024, $"Typing 32 draft updates allocated {allocated:N0} bytes.");
            return Task.CompletedTask;
        });
    }

    [Fact]
    public async Task NextUnresolvedReturnsToTheAlreadySelectedThread() {
        await WithDocument(async model => {
            model.SelectedCommentThread = model.CommentThreads[1];
            await model.ResolveCommentCommand.ExecuteAsync(null);
            model.NextUnresolvedCommentCommand.Execute(null);
            var selected = model.SelectedCommentThread;
            Assert.Equal(1, model.SelectedPage!.PageNumber);
            model.NavigateToOrganizerPage(2);
            model.NextUnresolvedCommentCommand.Execute(null);
            Assert.Same(selected, model.SelectedCommentThread);
            Assert.Equal(1, model.SelectedPage!.PageNumber);
        });
    }

    private static async Task WithDocument(Func<MainWindowViewModel, Task> action, bool duplicateNames = false, bool restricted = false) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "identity.pdf");
            byte[] bytes = StudioCommentReviewTests.CreateSource();
            if (restricted) bytes = PdfDocument.Load(bytes).Security.Encrypt(new("reader") {
                OwnerPassword = "owner", AllowedPermissions = PdfStandardPermissions.Print | PdfStandardPermissions.CopyContents | PdfStandardPermissions.Accessibility
            }).Pdf;
            if (duplicateNames) {
                for (int page = 1; page <= 2; page++) {
                    var document = PdfDocument.Load(bytes);
                    int id = document.Inspect().Annotations.First(annotation => annotation.PageNumber == page && annotation.Review?.IsReply != true && annotation.Subtype == "Text").ObjectNumber!.Value;
                    bytes = document.Annotations.Update(id, new() { Name = "duplicate-name" }).Bytes;
                }
            }
            File.WriteAllBytes(source, bytes);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services,
                promptPdfPassword: (_, _, _) => Task.FromResult<string?>("reader"));
            await model.OpenDocumentAsync(source);
            await action(model);
            return true;
        }, CancellationToken.None);
    }
}
