using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.VisualTree;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Organizer;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioImportPreviewTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ReviewedRangesAndSourceOrderUseCapturedContentsAndOneUndoStep(bool changeTarget) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string target = Path.Combine(services.Paths.Root, "target.pdf");
            string output = Path.Combine(services.Paths.Root, "output.pdf");
            CreatePdf(150).Save(target);
            var first = new TestStorageFile("content://import/first", CreatePdf(200, 210, 220).ToBytes(), "First.pdf");
            var second = new TestStorageFile("content://import/second", CreatePdf(300, 310).ToBytes(), "Second.pdf");
            var selected = await services.Storage.RegisterManyAsync([first.Item, second.Item], default);
            MainWindowViewModel? model = null;
            using (model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services,
                pickImportPdfs: _ => Task.FromResult(selected), pickSavePdf: _ => Task.FromResult<string?>(output),
                reviewPageImport: async preview => {
                    Assert.Equal(5, preview.ImportedPageCount);
                    preview.InsertBeforePage = "0"; Assert.False(preview.CanApply);
                    preview.InsertBeforePage = "1";
                    preview.Sources[0].PageRange = "1-999999999"; Assert.False(preview.CanApply);
                    preview.Sources[0].PageRange = "3,1";
                    preview.SelectedSource = preview.Sources[1]; preview.MoveSourceUpCommand.Execute(null);
                    Assert.Equal("Second.pdf", preview.Sources[0].Name);
                    Assert.Equal(4, preview.ImportedPageCount); Assert.True(preview.CanApply);
                    preview.Sources[0].IsIncluded = false; Assert.Equal(2, preview.ImportedPageCount);
                    preview.Sources[0].IsIncluded = true; Assert.Equal(4, preview.ImportedPageCount);
                    Array.Clear(first.Bytes); first.DenyRead = true; second.DenyRead = true;
                    if (changeTarget) await model!.RotateRightCommand.ExecuteAsync(null);
                    return true;
                })) {
                await model.OpenDocumentAsync(target);
                model.SetOrganizerSelection([model.OrganizerPages[0]]);
                await model.ImportPagesCommand.ExecuteAsync(null);
                Assert.Equal(first.Reads, first.ClosedReads); Assert.Equal(second.Reads, second.ClosedReads);
                Assert.Equal(0, first.Writes + second.Writes);
                if (changeTarget) {
                    Assert.Single(model.Pages); Assert.NotNull(model.ErrorMessage); Assert.Empty(services.Jobs.Entries);
                } else {
                    Assert.Null(model.ErrorMessage); Assert.Equal(5, model.SelectedPage!.PageNumber);
                    Assert.Equal(new[] { 1, 2, 3, 4 }, model.OrganizerPages.Where(page => page.IsSelected).Select(page => page.PageNumber));
                    await model.SaveAsCommand.ExecuteAsync(null);
                    Assert.Equal(new[] { 300D, 310D, 220D, 200D, 150D }, PdfDocument.Load(File.ReadAllBytes(output)).Inspect().Pages.Select(page => page.Width));
                    await model.UndoCommand.ExecuteAsync(null);
                    Assert.Single(model.Pages); Assert.False(model.CanUndo);
                    Assert.False(Assert.Single(services.Jobs.Entries).IsActive);
                }
            }
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task EncryptedImportUsesPasswordRetryOrCancelsWithoutMutation(bool cancelPassword) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string target = Path.Combine(services.Paths.Root, "target.pdf");
            string source = Path.Combine(services.Paths.Root, "encrypted.pdf");
            CreatePdf(150).Save(target);
            File.WriteAllBytes(source, CreatePdf(200, 210).Security.Encrypt(new PdfStandardEncryptionOptions("open") { OwnerPassword = "owner" }).Pdf);
            var invalid = new List<bool>();
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services,
                pickImportPdfs: _ => Task.FromResult<IReadOnlyList<string>>([source]), reviewPageImport: _ => Task.FromResult(true),
                promptPdfPassword: (name, failed, _) => {
                    Assert.Equal("encrypted.pdf", name); invalid.Add(failed);
                    return Task.FromResult<string?>(cancelPassword ? null : invalid.Count == 1 ? "wrong" : "owner");
                });
            await model.OpenDocumentAsync(target);
            await model.ImportPagesCommand.ExecuteAsync(null);
            Assert.Null(model.ErrorMessage);
            Assert.Equal(cancelPassword ? new[] { false } : new[] { false, true }, invalid);
            Assert.Equal(cancelPassword ? 1 : 3, model.Pages.Count);
            Assert.Equal(!cancelPassword, model.IsDirty);
            Assert.Equal(cancelPassword ? 0 : 1, services.Jobs.Entries.Count);
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData(380, 440, false)]
    [InlineData(600, 590, true)]
    public async Task RenderedImportReviewEditsRangesAndAppliesTheSelectedPages(int width, int height, bool dark) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            services.Preferences.Update(current => current with { Theme = dark ? StudioThemePreference.Dark : StudioThemePreference.Light });
            Directory.CreateDirectory(services.Paths.Root);
            string target = Path.Combine(services.Paths.Root, "target.pdf");
            CreatePdf(150).Save(target);
            string sourceName = "Quarterly reporting and supporting schedules " + new string('a', 100) + ".pdf";
            var sourceFile = new TestStorageFile("content://import/render-source", CreatePdf(200, 210, 220).ToBytes(), sourceName);
            string source = await services.Storage.RegisterAsync(sourceFile.Item, default);
            var owner = new Window { Width = 960, Height = 620 };
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services,
                pickImportPdfs: _ => Task.FromResult<IReadOnlyList<string>>([source]),
                reviewPageImport: preview => new PageImportDialog(preview) { Width = width, Height = height }.ShowDialog<bool>(owner));
            try {
                owner.Show(); await model.OpenDocumentAsync(target);
                Task pending = model.ImportPagesCommand.ExecuteAsync(null);
                PageImportDialog? dialog = null;
                var deadline = DateTime.UtcNow.AddSeconds(10);
                while (DateTime.UtcNow < deadline && (dialog = owner.OwnedWindows.OfType<PageImportDialog>().SingleOrDefault()) is null) await Task.Delay(10);
                Assert.NotNull(dialog);
                var preview = Assert.IsType<PageImportPreviewViewModel>(dialog.DataContext);
                await Layout(dialog);
                var range = dialog.GetVisualDescendants().OfType<TextBox>().Single(box => Avalonia.Automation.AutomationProperties.GetName(box) == services.Localizer.Format("Organizer.ImportRangeFor", sourceName));
                range.Text = "bad"; await Layout(dialog);
                Assert.False(preview.CanApply); Capture(dialog, $"import-invalid-{width}-{dark}");
                range.Text = "3,1"; await Layout(dialog);
                Assert.True(preview.CanApply); Assert.Equal(2, preview.ImportedPageCount);
                var edge = range.TranslatePoint(new Point(range.Bounds.Width, range.Bounds.Height), dialog)!.Value;
                var list = dialog.FindControl<ListBox>("SourcesList")!;
                var listEdge = list.TranslatePoint(new Point(list.Bounds.Width, list.Bounds.Height), dialog)!.Value;
                Assert.InRange(edge.X, 0, listEdge.X); Assert.InRange(edge.Y, 0, listEdge.Y);
                Capture(dialog, $"import-preview-{width}-{dark}");
                var apply = dialog.GetVisualDescendants().OfType<Button>().Single(button => Equals(button.Content, services.Localizer.Get("Organizer.ImportApply")));
                Assert.True(apply.IsEnabled); apply.Focus();
                apply.RaiseEvent(new KeyEventArgs { RoutedEvent = InputElement.KeyDownEvent, Key = Key.Enter });
                await pending;
                Assert.Null(model.ErrorMessage); Assert.Equal(3, model.Pages.Count);
                Assert.True(model.IsDirty);
            } finally { foreach (var dialog in owner.OwnedWindows.ToArray()) dialog.Close(); owner.Close(); }
            return true;
        }, CancellationToken.None);
    }
    private static PdfDocument CreatePdf(params int[] widths) => PdfDocument.Create(document => {
        foreach (int width in widths) document.Page(page => page.Size(width, 300));
    });
    private static async Task Layout(Window window) { window.UpdateLayout(); await Task.Delay(30); window.UpdateLayout(); }
    private static void Capture(Window window, string name) {
        string? root = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_IMPORT_CAPTURE_DIR");
        if (string.IsNullOrEmpty(root)) return;
        Directory.CreateDirectory(root);
        using var frame = window.CaptureRenderedFrame(); Assert.NotNull(frame);
        frame.Save(Path.Combine(root, name + ".png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
