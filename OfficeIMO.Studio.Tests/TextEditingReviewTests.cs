using Avalonia;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workspace;
using PdfPageCanvas = OfficeIMO.Studio.Features.Reader.PdfPageCanvas;

namespace OfficeIMO.Studio.Tests;

public sealed class TextEditingReviewTests {
    [Fact]
    public async Task InlineDraftRequiresRenderedPreviewAndInvalidatesWhenChanged() {
        using var files = new Files();
        PdfDocument document = CreateDocument();
        document.Save(files.Source);
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), pickSavePdf: _ => Task.FromResult<string?>(files.Output));
            await model.OpenDocumentAsync(files.Source);
            model.ShowEditModeCommand.Execute(null);
            PdfTextMatch match = document.Text.Find("Account")[0];
            model.Pages[0].SelectObject(Selection(match));
            await WaitUntilAsync(() => model.TextEditDraft?.IsReady == true);
            Assert.Same(model.TextEditDraft, model.Pages[0].InlineTextDraft);
            model.SelectedObjectText = "Record";
            await model.ApplyReviewedTextEditCommand.ExecuteAsync(null);
            Assert.False(model.IsDirty);
            await model.PreviewTextEditCommand.ExecuteAsync(null);
            Assert.True(model.HasTextPreview, model.ErrorMessage);
            Assert.False(model.IsDirty);
            Assert.NotNull(model.TextPreviewAfter);
            Assert.NotNull(model.TextPreviewRegion);
            Assert.InRange(model.TextPreviewRegion!.Value.Width, 0.01, 0.99);
            model.TextPreviewFullPage = true;
            Assert.Null(model.TextPreviewRegion);
            model.TextPreviewFullPage = false;
            Assert.NotNull(model.TextPreviewRegion);
            model.TextEditDraft!.FontSize = 11;
            Assert.False(model.HasTextPreview);
            await model.PreviewTextEditCommand.ExecuteAsync(null);
            Assert.True(model.HasTextPreview, model.ErrorMessage);
            await model.ApplyReviewedTextEditCommand.ExecuteAsync(null);
            Assert.True(model.IsDirty, model.ErrorMessage);
            Assert.Null(model.TextEditDraft);
            Assert.Null(model.Pages[0].InlineTextDraft);
            await model.SaveAsCommand.ExecuteAsync(null);
            Assert.Contains("Record", PdfDocument.Load(files.Output).Read().Text);
            Assert.Equal(3, PdfDocument.Load(files.Source).Text.Find("Account").Count);
            return true;
        }, default);
    }

    [Fact]
    public async Task IndividualReplacementChoicesAreAppliedTogetherAndMutationClearsPreview() {
        using var files = new Files();
        CreateDocument().Save(files.Source);
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), pickSavePdf: _ => Task.FromResult<string?>(files.Output));
            await model.OpenDocumentAsync(files.Source);
            model.ShowEditModeCommand.Execute(null);
            model.ReplaceAllFindText = "Account";
            model.ReplaceAllReplacementText = "Record";
            await model.FindTextReplacementsCommand.ExecuteAsync(null);
            Assert.Equal(3, model.ReplacementMatches.Count);
            model.SelectedReplacementMatch = model.ReplacementMatches[2];
            Assert.Equal(2, model.SelectedPage!.PageNumber);
            Assert.NotNull(model.Pages[1].ActiveSearchHighlight);
            Assert.Null(model.Pages[0].ActiveSearchHighlight);
            model.ReplacementMatches[1].IsIncluded = false;
            await model.PreviewTextEditCommand.ExecuteAsync(null);
            Assert.True(model.HasTextPreview, model.ErrorMessage);
            await model.ApplyReviewedTextEditCommand.ExecuteAsync(null);
            await model.SaveAsCommand.ExecuteAsync(null);
            PdfDocument result = PdfDocument.Load(files.Output);
            Assert.Equal(2, result.Text.Find("Record").Count);
            Assert.Single(result.Text.Find("Account"));
            model.ReplaceAllFindText = "Record";
            await model.FindTextReplacementsCommand.ExecuteAsync(null);
            await model.PreviewTextEditCommand.ExecuteAsync(null);
            Assert.True(model.HasTextPreview, model.ErrorMessage);
            model.SetOrganizerSelection([model.OrganizerPages[0]]);
            await model.DuplicateSelectedCommand.ExecuteAsync(null);
            Assert.False(model.HasTextPreview);
            Assert.Empty(model.ReplacementMatches);
            return true;
        }, default);
    }

    [Fact]
    public async Task PreparedTextBytesCannotBeAppliedAfterAnotherWorkspaceMutation() {
        using var files = new Files();
        CreateDocument().Save(files.Source);
        using PdfWorkspace workspace = await PdfWorkspace.OpenAsync(files.Source, default);
        PreparedTextEdit prepared = await workspace.PreviewTextReplacementsAsync("Account", "Record", false, false,
            [0], new PdfTextEditOptions(), [1], default);
        await workspace.ReplaceAllTextAsync("Account", "Entry", false, false, default);
        await Assert.ThrowsAsync<InvalidOperationException>(() => workspace.ApplyPreparedTextEditAsync(prepared, default));
        Assert.Equal(3, PdfDocument.Load(workspace.CopyBytes()).Text.Find("Entry").Count);
    }

    [Theory]
    [InlineData(0)] [InlineData(1)] [InlineData(2)] [InlineData(3)]
    [InlineData(4)] [InlineData(5)] [InlineData(6)] [InlineData(7)]
    public void ImageHandlesPreserveAspectRatio(int handle) {
        Rect changed = PdfPageCanvas.ResizeObjectBounds(new Rect(30, 40, 120, 60), new Vector(25, 18), handle, true);
        Assert.Equal(2, changed.Width / changed.Height, precision: 8);
        Assert.True(changed.Width > 0 && changed.Height > 0);
    }

    [Theory]
    [InlineData(120, 60)]
    [InlineData(60, 120)]
    [InlineData(1000, 5)]
    [InlineData(5, 1000)]
    public void EveryImageHandleKeepsBothDimensionsEditable(double width, double height) {
        var original = new Rect(30, 40, width, height);
        for (int handle = 0; handle < 8; handle++) {
            bool left = handle is 0 or 6 or 7, top = handle is 0 or 1 or 2;
            Rect changed = PdfPageCanvas.ResizeObjectBounds(original,
                new Vector(left ? 10000 : -10000, top ? 10000 : -10000), handle, true);
            Assert.InRange(changed.Width, 4, double.MaxValue);
            Assert.InRange(changed.Height, 4, double.MaxValue);
            Assert.Equal(width / height, changed.Width / changed.Height, precision: 8);
            Assert.Equal(left ? original.Right : handle is 2 or 3 or 4 ? original.Left : original.Center.X,
                left ? changed.Right : handle is 2 or 3 or 4 ? changed.Left : changed.Center.X, precision: 8);
            Assert.Equal(top ? original.Bottom : handle is 4 or 5 or 6 ? original.Top : original.Center.Y,
                top ? changed.Bottom : handle is 4 or 5 or 6 ? changed.Top : changed.Center.Y, precision: 8);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task SupersededInspectionCompletesWhileCpuWorkerIsStillOccupied(bool dispose) {
        using var files = new Files();
        PdfDocument document = CreateDocument();
        document.Save(files.Source);
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
            await model.OpenDocumentAsync(files.Source);
            model.ShowEditModeCommand.Execute(null);
            using PdfWorkspace blocker = await PdfWorkspace.OpenAsync(files.Source, default);
            var acquired = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var release = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            Task<bool> holding = blocker.RunNonDetachableCpuWorkAsync(() => {
                acquired.SetResult();
                release.Task.GetAwaiter().GetResult();
                return true;
            }, default);
            try {
                await acquired.Task;
                Task first = model.BeginInlineTextEditAsync(Selection(document.Text.Find("Account")[0]));
                PdfTextDraftViewModel firstDraft = Assert.IsType<PdfTextDraftViewModel>(model.TextEditDraft);
                Assert.False(first.IsCompleted);
                Task second = model.BeginInlineTextEditAsync(Selection(document.Text.Find("Account")[1]));
                await first.WaitAsync(TimeSpan.FromSeconds(3));
                Assert.False(firstDraft.IsReady);
                Assert.False(second.IsCompleted);
                if (dispose) model.Dispose(); else model.CancelTextEditCommand.Execute(null);
                await second.WaitAsync(TimeSpan.FromSeconds(3));
                Assert.Null(model.TextEditDraft);
                Assert.False(holding.IsCompleted);
            } finally {
                release.TrySetResult();
                await holding;
            }
            return true;
        }, default);
    }

    internal static PdfDocument CreateDocument() => PdfDocument.Create(compose => {
        compose.Page(page => page.Content(content => content.Item(item => item.Paragraph(text => text.Text("Account alpha Account beta")))));
        compose.Page(page => page.Content(content => content.Item(item => item.Paragraph(text => text.Text("Account gamma")))));
    });

    internal static PdfEditorSelection Selection(PdfTextMatch match) => new(PdfEditorSelectionKind.Text, match.PageNumber,
        new PdfEditorVisualBounds(match.VisualBounds.TopLeft.X, match.VisualBounds.TopLeft.Y,
            match.VisualBounds.BottomRight.X, match.VisualBounds.BottomRight.Y), Text: match.Text);

    internal static async Task WaitUntilAsync(Func<bool> condition) {
        for (int attempt = 0; attempt < 200 && !condition(); attempt++) await Task.Delay(10);
        Assert.True(condition());
    }

    internal sealed class Files : IDisposable {
        private readonly string _root = Path.Combine(Path.GetTempPath(), "officeimo-text-review-" + Guid.NewGuid().ToString("N"));
        internal Files() => Directory.CreateDirectory(_root);
        internal string Source => Path.Combine(_root, "source.pdf");
        internal string Output => Path.Combine(_root, "output.pdf");
        public void Dispose() => Directory.Delete(_root, true);
    }
}
