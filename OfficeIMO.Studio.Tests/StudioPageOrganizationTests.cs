using Avalonia;
using Avalonia.Automation;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.VisualTree;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Organizer;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workspace;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioPageOrganizationTests {
    [Fact]
    public async Task RangeSelectionRejectsInvalidInputAndNoOpMovesDoNotCreateUndoEntries() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-organizer-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string path = Path.Combine(root, "pages.pdf");
        try {
            CreateDocument().Save(path);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
            await model.OpenDocumentAsync(path);
            model.OrganizerPageRange = "1-2,2,5";
            model.SelectPageRangeCommand.Execute(null);
            Assert.Equal(new[] { 1, 2, 5 }, Selected(model));
            model.OrganizerPageRange = "1-999999999";
            model.SelectPageRangeCommand.Execute(null);
            Assert.NotNull(model.OrganizerRangeError);
            Assert.Equal(new[] { 1, 2, 5 }, Selected(model));
            model.SelectAllPagesCommand.Execute(null);
            await model.MoveSelectedUpCommand.ExecuteAsync(null);
            Assert.False(model.CanUndo);
            Assert.False(model.IsDirty);
            model.OrganizerPageRange = "2-3,5";
            model.SelectPageRangeCommand.Execute(null);
            Assert.Null(model.OrganizerRangeError);
            model.SelectedPage = model.Pages[1];
            await model.MoveSelectedDownCommand.ExecuteAsync(null);
            Assert.Equal(new[] { 3, 4, 6 }, Selected(model));
            Assert.Equal(3, model.SelectedPage!.PageNumber);
            await model.SaveCommand.ExecuteAsync(null);
            Assert.Equal(new[] { "Page 1", "Page 4", "Page 2", "Page 3", "Page 6", "Page 5" }, ReadPages(path));
        } finally { Directory.Delete(root, recursive: true); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ReviewedMoveAppliesExactOrderOrRejectsAChangedDocument(bool changeDuringPreview) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-organizer-preview-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string path = Path.Combine(root, "pages.pdf");
        try {
            CreateDocument().Save(path);
            MainWindowViewModel? model = null;
            using (model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), reviewPageMove: async preview => {
                preview.InsertBeforePage = "1";
                Assert.Equal(new[] { 2, 4, 1, 3, 5, 6 }, preview.Plan!.SourcePageNumbers);
                if (changeDuringPreview) await model!.RotateRightCommand.ExecuteAsync(null);
                return true;
            })) {
                await model.OpenDocumentAsync(path);
                model.OrganizerPageRange = "2,4";
                model.SelectPageRangeCommand.Execute(null);
                model.SelectedPage = model.Pages[3];
                await model.MoveSelectedToCommand.ExecuteAsync(null);
                if (changeDuringPreview) {
                    Assert.NotNull(model.ErrorMessage);
                    Assert.Equal(new[] { 2, 4 }, Selected(model));
                } else {
                    Assert.Null(model.ErrorMessage);
                    Assert.Equal(new[] { 1, 2 }, Selected(model));
                    Assert.Equal(2, model.SelectedPage!.PageNumber);
                }
                await model.SaveCommand.ExecuteAsync(null);
                Assert.Equal(changeDuringPreview
                    ? new[] { "Page 1", "Page 2", "Page 3", "Page 4", "Page 5", "Page 6" }
                    : new[] { "Page 2", "Page 4", "Page 1", "Page 3", "Page 5", "Page 6" }, ReadPages(path));
            }
        } finally { Directory.Delete(root, recursive: true); }
    }

    [Theory]
    [InlineData(960, 620, false)]
    [InlineData(1280, 800, true)]
    public async Task KeyboardReorderAndMoveDialogExposeTheActualResult(int width, int height, bool dark) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            services.Preferences.Update(current => current with { Theme = dark ? StudioThemePreference.Dark : StudioThemePreference.Light });
            Directory.CreateDirectory(services.Paths.Root);
            string path = Path.Combine(services.Paths.Root, "organization.pdf");
            CreateDocument().Save(path);
            var window = new MainWindow(services) { Width = width, Height = height };
            try {
                window.Show();
                var model = window.ViewModel;
                await model.OpenDocumentAsync(path);
                model.DocumentMode = StudioDocumentMode.Pages;
                model.IsOrganizerRangeExpanded = true;
                await Layout(window);
                var input = window.GetVisualDescendants().OfType<TextBox>().Single(control =>
                    AutomationProperties.GetName(control) == services.Localizer.Get("Organizer.PageRange"));
                input.Text = "2-3,5";
                await Layout(window);
                Assert.Equal("2-3,5", model.OrganizerPageRange);
                Click(window, services.Localizer.Get("Organizer.SelectRange"));
                Assert.Equal(new[] { 2, 3, 5 }, Selected(model));
                Assert.False(model.IsOrganizerRangeExpanded);
                model.SelectedPage = model.Pages[1];
                var organizer = window.GetVisualDescendants().OfType<DocumentWorkspaceView>().Single().OrganizerListControl;
                organizer.Focus();
                organizer.RaiseEvent(new KeyEventArgs { RoutedEvent = InputElement.KeyDownEvent, Key = Key.Down, KeyModifiers = KeyModifiers.Alt });
                await model.MoveSelectedDownCommand.ExecutionTask!;
                Assert.Equal(new[] { 3, 4, 6 }, Selected(model));
                Assert.Equal(3, model.SelectedPage!.PageNumber);
                await Layout(window);
                await model.SelectedPage.EnsureRenderedAsync();
                await Layout(window);
                var deadline = DateTime.UtcNow.AddSeconds(5);
                while (model.OrganizerPages.Any(page => page.IsLoading) && DateTime.UtcNow < deadline) await Task.Delay(10);
                Assert.DoesNotContain(model.OrganizerPages, page => page.IsLoading);
                await Layout(window);
                Capture(window, $"organizer-keyboard-{width}-{dark}");
                Assert.True(organizer.Bounds.Height >= 190);
                Click(window, services.Localizer.Get("Organizer.MoveTo"));
                var dialog = Assert.Single(window.OwnedWindows.OfType<PageMoveDialog>());
                var preview = Assert.IsType<PageMovePreviewViewModel>(dialog.DataContext);
                var destination = dialog.GetVisualDescendants().OfType<TextBox>().Single();
                await Layout(dialog);
                foreach (string invalid in new[] { "abc", "99", "1.5", "", "-1", "99999999999999999999" }) {
                    destination.Text = invalid;
                    await Layout(dialog);
                    Assert.False(preview.CanApply);
                    Assert.Null(preview.Plan);
                    Assert.False(dialog.GetVisualDescendants().OfType<Button>().Single(button => Equals(button.Content, services.Localizer.Get("Organizer.ApplyMove"))).IsEffectivelyEnabled);
                }
                Capture(dialog, $"organizer-invalid-{width}-{dark}");
                destination.Text = "1";
                await Layout(dialog);
                Assert.True(preview.CanApply);
                Assert.Equal(new[] { 3, 4, 6, 1, 2, 5 }, preview.Plan!.SourcePageNumbers);
                Capture(dialog, $"organizer-preview-{width}-{dark}");
                Click(dialog, services.Localizer.Get("Organizer.ApplyMove"));
                await model.MoveSelectedToCommand.ExecutionTask!;
                Assert.Equal(1, model.SelectedPage!.PageNumber);
                Assert.Equal(new[] { 1, 2, 3 }, Selected(model));
                await model.SaveCommand.ExecuteAsync(null);
                Assert.Equal(new[] { "Page 2", "Page 3", "Page 5", "Page 1", "Page 4", "Page 6" }, ReadPages(path));
            } finally {
                foreach (Window dialog in window.OwnedWindows.ToArray()) dialog.Close(false);
                window.Close();
            }
            return true;
        }, CancellationToken.None);
    }

    private static PdfDocument CreateDocument() {
        return PdfDocument.Create(compose => {
            for (int page = 1; page <= 6; page++) {
                int number = page;
                compose.Page(builder => builder.Content(content => content.Item(item => item.Text("Page " + number))));
            }
        });
    }
    private static string[] ReadPages(string path) => PdfReadDocument.Open(path).Pages.Select(page => page.ExtractText().Trim()).ToArray();
    private static int[] Selected(MainWindowViewModel model) => model.OrganizerPages.Where(page => page.IsSelected).Select(page => page.PageNumber).ToArray();
    private static async Task Layout(Window window) => await Avalonia.Threading.Dispatcher.UIThread.InvokeAsync(
        () => window.UpdateLayout(), Avalonia.Threading.DispatcherPriority.Background);
    private static void Click(Window window, string text) {
        var button = window.GetVisualDescendants().OfType<Button>().Single(button => Equals(button.Content, text));
        Assert.True(button.IsEffectivelyEnabled);
        button.Focus();
        button.RaiseEvent(new KeyEventArgs { RoutedEvent = InputElement.KeyDownEvent, Key = Key.Enter });
    }
    private static void Capture(Window window, string name) {
        string? directory = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrEmpty(directory)) return;
        Directory.CreateDirectory(directory);
        using var bitmap = window.CaptureRenderedFrame();
        Assert.NotNull(bitmap);
        bitmap.Save(Path.Combine(directory, name + ".png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
