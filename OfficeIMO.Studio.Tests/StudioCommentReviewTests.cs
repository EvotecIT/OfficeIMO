using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Styling;
using Avalonia.VisualTree;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workspace;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioCommentReviewTests {
    [Theory]
    [InlineData(960, false)]
    [InlineData(1280, true)]
    public async Task ReviewThreadsReplyResolveUndoAndSaveInRenderedWorkspace(int width, bool dark) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            Application.Current!.RequestedThemeVariant = dark ? ThemeVariant.Dark : ThemeVariant.Light;
            var services = ((App)Application.Current).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "comments.pdf");
            string saved = Path.Combine(services.Paths.Root, "reviewed.pdf");
            File.WriteAllBytes(source, CreateSource());
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null),
                pickSavePdf: _ => Task.FromResult<string?>(saved), services: services);
            await model.OpenDocumentAsync(source);
            model.ShowAnnotateModeCommand.Execute(null);
            var window = new Window { Width = width, Height = 740, Content = new DocumentWorkspaceView { DataContext = model } };
            try {
                window.Show(); window.UpdateLayout();
                Assert.Equal(2, model.CommentThreads.Count);
                model.NextUnresolvedCommentCommand.Execute(null);
                Assert.Equal("Check the proposed delivery date.", model.SelectedCommentThread!.Contents);
                model.CommentReplyText = "Delivery confirmed for Monday.\nPlease keep the revised schedule.";
                model.EditorAuthor = "Morgan";
                await model.ReplyToCommentCommand.ExecuteAsync(null);
                Assert.Null(model.ErrorMessage);
                Assert.NotNull(model.SelectedCommentThread);
                Assert.Equal(3, model.SelectedCommentThread.Entries.Count);
                Assert.Equal(string.Empty, model.CommentReplyText);
                await model.ResolveCommentCommand.ExecuteAsync(null);
                Assert.True(model.SelectedCommentThread!.IsResolved);
                await model.UndoCommand.ExecuteAsync(null);
                Assert.False(model.SelectedCommentThread!.IsResolved);
                await model.RedoCommand.ExecuteAsync(null);
                Assert.True(model.SelectedCommentThread!.IsResolved);
                window.UpdateLayout();
                await model.SelectedPage!.EnsureRenderedAsync();
                Assert.NotNull(model.SelectedPage.Scene);
                Assert.NotNull(model.SelectedPage.CommentAnchorObjectNumber);
                await Avalonia.Threading.Dispatcher.UIThread.InvokeAsync(() => window.UpdateLayout(), Avalonia.Threading.DispatcherPriority.Background);
                Capture(window, $"comments-{width}-{(dark ? "dark" : "light")}.png");
                var reopen = window.GetVisualDescendants().OfType<Button>().Single(button => ReferenceEquals(button.Command, model.ReopenCommentCommand));
                reopen.BringIntoView(); window.UpdateLayout();
                Capture(window, $"comments-actions-{width}.png");
                await model.SaveAsCommand.ExecuteAsync(null);
                Assert.Null(model.ErrorMessage);
                var catalog = PdfAnnotationReviewCatalog.Read(File.ReadAllBytes(saved));
                var thread = Assert.Single(catalog.Threads, t => t.Root.Annotation.Contents == "Check the proposed delivery date.");
                Assert.Equal(PdfAnnotationReviewState.Completed, thread.Root.Annotation.Review!.StandardState);
                Assert.Contains(thread.Root.Replies, entry => entry.Annotation.Title == "Morgan" && entry.Annotation.Contents!.Contains("Delivery confirmed"));
                string? artifacts = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
                if (!string.IsNullOrWhiteSpace(artifacts)) File.Copy(saved, Path.Combine(artifacts, "comments-reviewed.pdf"), true);
                await model.ReopenCommentCommand.ExecuteAsync(null);
                Assert.False(model.SelectedCommentThread!.IsResolved);
                model.CommentAuthorFilter = "Alex"; // A reply author also matches its root thread.
                Assert.Single(model.CommentThreads);
                model.SelectedCommentStatus = model.CommentStatuses.Single(choice => choice.Id == "resolved");
                Assert.Empty(model.CommentThreads);
                Assert.False(model.HasCommentThread);
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task ReplyDraftsFollowThreadsAndClearOnlyTheSubmittedDraft() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "drafts.pdf");
            File.WriteAllBytes(source, CreateSource());
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services);
            await model.OpenDocumentAsync(source);
            model.NextUnresolvedCommentCommand.Execute(null);
            var first = model.SelectedCommentThread;
            model.CommentReplyText = "Draft for the first thread";
            model.NextUnresolvedCommentCommand.Execute(null);
            Assert.NotSame(first, model.SelectedCommentThread);
            Assert.Empty(model.CommentReplyText);
            model.CommentReplyText = "Second draft";
            model.SelectedCommentThread = first;
            Assert.Equal("Draft for the first thread", model.CommentReplyText);
            System.ComponentModel.PropertyChangedEventHandler cancel = (_, args) => {
                if (args.PropertyName == nameof(model.IsWorkspaceBusy) && model.IsWorkspaceBusy) model.CancelOperationCommand.Execute(null);
            };
            model.PropertyChanged += cancel;
            await model.ReplyToCommentCommand.ExecuteAsync(null);
            model.PropertyChanged -= cancel;
            Assert.Equal("Draft for the first thread", model.CommentReplyText);
            Assert.Equal(2, model.SelectedCommentThread!.Entries.Count);
            Assert.False(model.IsDirty);
            await model.ReplyToCommentCommand.ExecuteAsync(null);
            Assert.Null(model.ErrorMessage);
            Assert.Empty(model.CommentReplyText);
            model.NextUnresolvedCommentCommand.Execute(null);
            Assert.Equal("Second draft", model.CommentReplyText);
            Assert.Equal(2, model.SelectedPage!.PageNumber);
            return true;
        }, CancellationToken.None);
    }

    private static byte[] CreateSource() {
        var document = PdfDocument.Create(compose => {
            compose.Page(page => page.Size(600, 800).Canvas(canvas => canvas.TextAnnotation("Check the proposed delivery date.", 70, 700)));
            compose.Page(page => page.Size(600, 800).Canvas(canvas => canvas.TextAnnotation("Confirm the support contact.", 70, 700)));
        });
        using var stream = new MemoryStream(); document.Save(stream);
        byte[] bytes = stream.ToArray();
        var first = PdfDocument.Load(bytes).Inspect().GetAnnotationsBySubtype("Text").First();
        return PdfAnnotationReviewEditor.AddReply(bytes, first.ObjectNumber!.Value, "Please verify with the team.",
            new() { Author = "Alex", CreatePopup = true }).Bytes;
    }

    private static void Capture(Window window, string name) {
        using var frame = window.CaptureRenderedFrame();
        Assert.NotNull(frame);
        string? path = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(path)) return;
        Directory.CreateDirectory(path);
        frame.Save(Path.Combine(path, name), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
