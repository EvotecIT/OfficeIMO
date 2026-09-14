using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Workspace;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;

namespace OfficeIMO.Studio.Tests;

public sealed class WatermarkPreviewTests {
    [Fact]
    public async Task PreviewFollowsTargetRangeAndCannotReviewAnUnchangedPage() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            byte[] source = PdfDocument.Create(document => {
                document.Page(page => page.Content(content => content.Text("First page")));
                document.Page(page => page.Content(content => content.Text("Second page")));
            }).ToBytes();
            await WithDocument(async workspace => {
                var settings = new PdfWatermarkOptions { Text = "SECOND PAGE MARK", TargetPages = PdfPageSelector.Parse("2") };
                await Assert.ThrowsAsync<ArgumentException>(() => workspace.PrepareWatermarkAsync(settings, 1, CancellationToken.None));
                using var model = new Features.Editor.WatermarkPreviewViewModel(2, 1,
                    ((App)Application.Current!).Services.Localizer, workspace.PrepareWatermarkAsync,
                    _ => Task.FromResult<byte[]?>(null));
                var dialog = new Features.Editor.WatermarkDialog(model) { Width = 1000, Height = 740 };
                try {
                    dialog.Show();
                    await TextEditingReviewTests.WaitUntilAsync(() => model.CanApply);
                    model.Text = settings.Text;
                    model.PageRange = "2";
                    Assert.False(model.CanApply);
                    await model.PreviewCommand.ExecuteAsync(null);
                    Assert.True(model.CanApply, model.ErrorMessage);
                    Assert.Equal(2, model.PreviewPage);
                    var preview = model.Prepared!;
                    Assert.Equal(new[] { 2 }, preview.Pages);
                    Assert.NotEmpty(preview.PageImage);
                    var candidate = PdfDocument.Load(preview.DocumentBytes);
                    var displayOptions = new PdfPageDisplayOptions {
                        Scale = Math.Min(1.5D, 1000D / Math.Max(preview.PageWidth, preview.PageHeight)),
                        MaximumOutputBytes = 8 * 1024 * 1024
                    };
                    Assert.Equal(candidate.Render.DisplayPage(2, displayOptions).Bytes, preview.PageImage);
                    Assert.NotEqual(candidate.Render.DisplayPage(1, displayOptions).Bytes, preview.PageImage);
                    Assert.Equal(source, workspace.CopyBytes());
                    dialog.UpdateLayout();
                    using var frame = dialog.CaptureRenderedFrame();
                    Assert.NotNull(frame);
                    string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
                    if (!string.IsNullOrWhiteSpace(output)) {
                        Directory.CreateDirectory(output);
                        frame.Save(Path.Combine(output, "watermark-target-page.png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
                    }
                    await workspace.ApplyWatermarkAsync(preview, CancellationToken.None);
                    Assert.Equal(preview.DocumentBytes, workspace.CopyBytes());
                    Assert.Equal(new[] { 2 }, Assert.Single(PdfDocument.Load(workspace.CopyBytes()).Stamp.ReadWatermarks()).TargetPages!.Resolve(2));
                } finally { dialog.Close(); }
            }, source);
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task ReopeningSingleWatermarkRestoresSettingsAndRevisionSupportsUndo() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            await WithDocument(async workspace => {
                var settings = new PdfWatermarkOptions { Text = "ORIGINAL MARK", X = 80, Y = 150, Opacity = .65, RotationDegrees = 20 };
                var first = await workspace.PrepareWatermarkAsync(settings, 1, CancellationToken.None);
                await workspace.ApplyWatermarkAsync(first, CancellationToken.None);
                byte[] previous = workspace.CopyBytes();
                var existing = await workspace.ReadWatermarksAsync(CancellationToken.None);
                using var model = new Features.Editor.WatermarkPreviewViewModel(1, 1,
                    ((App)Application.Current!).Services.Localizer, workspace.PrepareWatermarkAsync,
                    _ => Task.FromResult<byte[]?>(null), existing);
                Assert.Same(model.Watermarks[1], model.SelectedWatermark);
                Assert.Equal("ORIGINAL MARK", model.Text);
                Assert.Equal(80, model.X);
                Assert.Equal(150, model.Y);
                Assert.Equal(65, model.Opacity);
                Assert.Equal(20, model.Rotation);
                await model.PreviewCommand.ExecuteAsync(null);
                Assert.True(model.CanApply, model.ErrorMessage);
                Assert.Equal(settings.Id, Assert.Single(PdfDocument.Load(model.Prepared!.DocumentBytes).Stamp.ReadWatermarks()).Id);
                model.Text = "REVISED MARK";
                await model.PreviewCommand.ExecuteAsync(null);
                Assert.True(model.CanApply, model.ErrorMessage);
                await workspace.ApplyWatermarkAsync(model.Prepared!, CancellationToken.None);
                var result = PdfDocument.Load(workspace.CopyBytes());
                string text = System.Text.RegularExpressions.Regex.Replace(result.Read().Text, @"\s+", " ");
                Assert.DoesNotContain("ORIGINAL MARK", text);
                Assert.Contains("REVISED MARK", text);
                Assert.Equal(settings.Id, Assert.Single(result.Stamp.ReadWatermarks()).Id);
                await workspace.UndoAsync(CancellationToken.None);
                Assert.Equal(previous, workspace.CopyBytes());
            });
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData(760, 580, 0, 1)]
    [InlineData(1000, 740, 0, 1)]
    [InlineData(760, 580, 90, 1)]
    [InlineData(1000, 740, 0, 2)]
    [InlineData(1000, 740, 270, 2)]
    public async Task PreviewDragUsesPageCoordinatesAtDifferentWindowSizes(int width, int height, int rotation, int userUnit) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            await WithDocument(async workspace => {
                var localizer = ((App)Application.Current!).Services.Localizer;
                using var model = new Features.Editor.WatermarkPreviewViewModel(1, 1, localizer,
                    workspace.PrepareWatermarkAsync, _ => Task.FromResult<byte[]?>(null));
                var dialog = new Features.Editor.WatermarkDialog(model) { Width = width, Height = height };
                try {
                    dialog.Show();
                    await TextEditingReviewTests.WaitUntilAsync(() => model.CanApply);
                    dialog.UpdateLayout();
                    var surface = dialog.FindControl<Grid>("PreviewSurface")!;
                    var preview = model.Prepared!;
                    var layout = PdfDocument.Load(workspace.CopyBytes()).GetPageLayouts()[0];
                    Assert.Equal(layout.VisualWidth, preview.PageWidth);
                    Assert.Equal(layout.VisualHeight, preview.PageHeight);
                    double scale = Math.Min(surface.Bounds.Width / preview.PageWidth, surface.Bounds.Height / preview.PageHeight);
                    Point start = surface.TranslatePoint(new Point(surface.Bounds.Width / 2, surface.Bounds.Height / 2), dialog)!.Value;
                    Point end = new(start.X + 20 * scale, start.Y + 30 * scale);
                    dialog.MouseDown(start, MouseButton.Left);
                    dialog.MouseMove(end, RawInputModifiers.LeftMouseButton);
                    Assert.True(dialog.FindControl<Border>("PlacementOutline")!.IsVisible);
                    dialog.MouseUp(end, MouseButton.Left);
                    await TextEditingReviewTests.WaitUntilAsync(() => model.CanApply);
                    Assert.Equal((preview.PageWidth - (double)model.Width) / 2 + 20, (double)model.X!.Value, 1);
                    Assert.Equal((preview.PageHeight - (double)model.Height) / 2 + 30, (double)model.Y!.Value, 1);
                    Assert.False(workspace.IsDirty);
                    dialog.UpdateLayout();
                    using var frame = dialog.CaptureRenderedFrame();
                    Assert.NotNull(frame);
                    string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
                    if (!string.IsNullOrWhiteSpace(output)) {
                        Directory.CreateDirectory(output);
                        frame.Save(Path.Combine(output, $"watermark-placement-{width}x{height}-r{rotation}-u{userUnit}.png"),
                            Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
                    }
                } finally { dialog.Close(); }
            }, GeometrySource(rotation, userUnit));
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task ChangingSettingsRequiresAnotherPreview() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            await WithDocument(async workspace => {
                var localizer = ((App)Avalonia.Application.Current!).Services.Localizer;
                using var model = new Features.Editor.WatermarkPreviewViewModel(1, 1, localizer,
                    workspace.PrepareWatermarkAsync, _ => Task.FromResult<byte[]?>(null));
                await model.PreviewCommand.ExecuteAsync(null);
                Assert.True(model.CanApply, model.ErrorMessage);
                Assert.NotNull(model.PreviewImage);
                model.Rotation = 15;
                Assert.False(model.CanApply);
                Assert.Null(model.Prepared);
                Assert.Null(model.PreviewImage);
                await model.PreviewCommand.ExecuteAsync(null);
                Assert.True(model.CanApply, model.ErrorMessage);
                Assert.False(workspace.IsDirty);
            });
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task ClosingPreviewIgnoresLateCompletion() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            await WithDocument(async workspace => {
                var localizer = ((App)Avalonia.Application.Current!).Services.Localizer;
                var completed = await workspace.PrepareWatermarkAsync(new(), 1, CancellationToken.None);
                var pending = new TaskCompletionSource<PdfWatermarkPreview>(TaskCreationOptions.RunContinuationsAsynchronously);
                using var model = new Features.Editor.WatermarkPreviewViewModel(1, 1, localizer,
                    (_, _, _) => pending.Task, _ => Task.FromResult<byte[]?>(null));
                Task running = model.PreviewCommand.ExecuteAsync(null);
                Assert.True(model.IsBusy);
                model.Dispose();
                pending.SetResult(completed);
                await running;
                Assert.Null(model.PreviewImage);
                Assert.False(model.CanApply);
                Assert.False(workspace.IsDirty);
            });
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task PreviewIsNonDestructiveAndApplyUsesReviewedBytesWithUndo() {
        await WithDocument(async workspace => {
            byte[] original = workspace.CopyBytes();
            long revision = workspace.Revision;
            var preview = await workspace.PrepareWatermarkAsync(new() { Text = "REVIEWED" }, 1, CancellationToken.None);
            Assert.Equal(original, workspace.CopyBytes());
            Assert.Equal(revision, workspace.Revision);
            Assert.False(workspace.IsDirty);
            Assert.NotEmpty(preview.PageImage);
            await workspace.ApplyWatermarkAsync(preview, CancellationToken.None);
            Assert.Equal(preview.DocumentBytes, workspace.CopyBytes());
            Assert.True(workspace.IsDirty);
            Assert.Contains("REVIEWED", PdfDocument.Load(workspace.CopyBytes()).Read().Text);
            await workspace.UndoAsync(CancellationToken.None);
            Assert.Equal(original, workspace.CopyBytes());
            Assert.False(workspace.IsDirty);
        });
    }

    [Fact]
    public async Task StalePreviewCannotOverwriteAnotherEdit() {
        await WithDocument(async workspace => {
            var preview = await workspace.PrepareWatermarkAsync(new(), 1, CancellationToken.None);
            await workspace.ApplyWatermarkAsync("ANOTHER EDIT", CancellationToken.None);
            byte[] changed = workspace.CopyBytes();
            long revision = workspace.Revision;
            await Assert.ThrowsAsync<InvalidOperationException>(() => workspace.ApplyWatermarkAsync(preview, CancellationToken.None));
            Assert.Equal(changed, workspace.CopyBytes());
            Assert.Equal(revision, workspace.Revision);
        });
    }

    [Fact]
    public async Task PreviewCannotBeAppliedToAnotherDocument() {
        await WithDocument(async first => {
            var preview = await first.PrepareWatermarkAsync(new(), 1, CancellationToken.None);
            await WithDocument(async second => {
                byte[] original = second.CopyBytes();
                await Assert.ThrowsAsync<InvalidOperationException>(() => second.ApplyWatermarkAsync(preview, CancellationToken.None));
                Assert.Equal(original, second.CopyBytes());
                Assert.False(second.IsDirty);
            });
        });
    }

    [Fact]
    public async Task CancelledApplyLeavesDocumentUntouched() {
        await WithDocument(async workspace => {
            byte[] original = workspace.CopyBytes();
            var preview = await workspace.PrepareWatermarkAsync(new(), 1, CancellationToken.None);
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => workspace.ApplyWatermarkAsync(preview, cancellation.Token));
            Assert.Equal(original, workspace.CopyBytes());
            Assert.False(workspace.IsDirty);
        });
    }

    private static byte[] GeometrySource(int rotation, int userUnit) {
        const string content = "BT /F1 12 Tf 50 450 Td (Original content) Tj ET";
        return System.Text.Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7", "1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj",
            "2 0 obj << /Type /Pages /Count 1 /Kids [3 0 R] >> endobj",
            $"3 0 obj << /Type /Page /Parent 2 0 R /MediaBox [0 0 400 500] /UserUnit {userUnit} /Rotate {rotation} /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >> endobj",
            $"4 0 obj << /Length {content.Length} >> stream", content, "endstream endobj",
            "5 0 obj << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> endobj",
            "trailer << /Root 1 0 R /Size 6 >>", "%%EOF"
        }));
    }

    private static async Task WithDocument(Func<PdfWorkspace, Task> action, byte[]? bytes = null) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-watermark-preview-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string path = Path.Combine(root, "source.pdf");
            if (bytes is null) PdfDocument.Create(document => document.Page(page => page.Content(content => content.Text("Original content")))).Save(path);
            else File.WriteAllBytes(path, bytes);
            using var workspace = await PdfWorkspace.OpenAsync(path, CancellationToken.None);
            await action(workspace);
        } finally { Directory.Delete(root, recursive: true); }
    }
}
