using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Input.Raw;
using Avalonia.Styling;
using Avalonia.Threading;
using Avalonia.VisualTree;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Mobile;
using OfficeIMO.Studio.Infrastructure;
using Xunit;

namespace OfficeIMO.Studio.Tests;

public sealed partial class StudioMobileWorkspaceTests {
    [Theory]
    [InlineData("Save")]
    [InlineData("Discard")]
    [InlineData("Cancel")]
    [InlineData("Save fails")]
    public async Task MobileCloseOffersAnExplicitDecisionAndProtectsTheSavedCopy(string choice) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-mobile-close-" + Guid.NewGuid().ToString("N"));
        try {
            using var session = TestAppBuilder.StartSession();
            await session.Dispatch(async () => {
                var services = StudioApplicationServices.Create(new StudioDataPaths(Path.Combine(root, "data")),
                    new StudioLocalDocumentRoot(Path.Combine(root, "Documents")));
                using var controller = new MobileDocumentController(services,
                    _ => Task.FromResult<Avalonia.Platform.Storage.IStorageFile?>(null), _ => Task.CompletedTask);
                var view = new MobileWorkspaceView();
                view.Connect(controller);
                view.ShareDocumentAsync = controller.ShareAsync;
                var window = new Window { Content = view };
                try {
                    window.Show();
                    Layout(window, 390, 844);
                    await controller.OpenSampleAsync();
                    string path = controller.Document.DocumentPath!;
                    byte[] original = File.ReadAllBytes(path);
                    controller.Document.EditorText = "Keep or discard this edit explicitly";
                    await controller.Document.ApplyPageMarkupAsync(PdfEditorTool.Note, new PdfEditorGesture(1, 24, 24, 48, 48, []));
                    Assert.True(controller.Document.IsDirty, controller.Document.ErrorMessage);
                    if (choice == "Save fails") File.AppendAllText(path, "\n% External change\n");
                    byte[] savedBeforeClose = File.ReadAllBytes(path);
                    Task close = controller.Tabs.CloseSelectedTabAsync();
                    Layout(window, 390, 844);
                    Assert.False(close.IsCompleted);
                    Assert.True(view.FindControl<ScrollViewer>("CloseScroll")!.IsEffectivelyVisible);
                    Assert.False(view.FindControl<Border>("HeaderBar")!.IsEffectivelyEnabled);
                    Assert.Same(view.FindControl<Button>("CloseCancel"), window.FocusManager!.GetFocusedElement());
                    Assert.Contains("Welcome to Studio.pdf", view.FindControl<TextBlock>("CloseDescription")!.Text);
                    Capture(window, "mobile-close-phone");
                    Click(view, choice == "Save fails" ? "Save" : choice);
                    await close;
                    Assert.False(view.FindControl<Border>("SheetScrim")!.IsVisible);
                    if (choice == "Save fails") {
                        Assert.Single(controller.Tabs.Tabs);
                        Assert.True(controller.Document.IsDirty);
                        Assert.True(controller.Document.HasError);
                        Assert.Equal(savedBeforeClose, File.ReadAllBytes(path));
                    } else if (choice == "Cancel") {
                        Assert.Single(controller.Tabs.Tabs);
                        Assert.True(controller.Document.IsDirty);
                        Assert.Equal(original, File.ReadAllBytes(path));
                        // Escape and detaching the presenter are safe cancellation paths too.
                        close = controller.Tabs.CloseSelectedTabAsync();
                        window.KeyPress(Key.Escape, RawInputModifiers.None, PhysicalKey.None, null);
                        await close;
                        Assert.Single(controller.Tabs.Tabs);
                        close = controller.Tabs.CloseSelectedTabAsync();
                        window.Close();
                        await close;
                        Assert.Single(controller.Tabs.Tabs);
                    } else {
                        Assert.Empty(controller.Tabs.Tabs);
                        Assert.Equal(choice == "Discard", original.SequenceEqual(File.ReadAllBytes(path)));
                        await controller.Tabs.ReopenClosedTabAsync();
                        Assert.Single(controller.Tabs.Tabs);
                        Assert.False(controller.Document.IsDirty);
                    }
                } finally { window.Close(); }
                return true;
            }, CancellationToken.None);
        } finally { if (Directory.Exists(root)) Directory.Delete(root, true); }
    }

    [Fact]
    public async Task MobileSearchDocumentChooserAndSaveUseTheSharedDocumentCommands() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-mobile-workflow-" + Guid.NewGuid().ToString("N"));
        try {
            using var session = TestAppBuilder.StartSession();
            await session.Dispatch(async () => {
                var services = StudioApplicationServices.Create(new StudioDataPaths(Path.Combine(root, "data")),
                    new StudioLocalDocumentRoot(Path.Combine(root, "Documents")));
                using var controller = new MobileDocumentController(services,
                    _ => Task.FromResult<Avalonia.Platform.Storage.IStorageFile?>(null), _ => Task.CompletedTask);
                var view = new MobileWorkspaceView();
                view.Connect(controller);
                view.ShareDocumentAsync = controller.ShareAsync;
                var window = new Window { Content = view };
                try {
                    window.Show();
                    Layout(window, 1024, 768);
                    await controller.OpenSampleAsync();
                    var first = controller.Tabs.SelectedTab!;
                    first.Document.SearchQuery = "page";
                    await first.Document.SearchCommand.ExecuteAsync(null);
                    Assert.True(first.Document.SearchResults.Count > 1);
                    Click(view, "Search");
                    var search = view.FindControl<TextBox>("MobileSearchBox")!;
                    search.Focus();
                    window.KeyPress(Key.Enter, RawInputModifiers.None, PhysicalKey.None, null);
                    Assert.Same(first.Document.SearchResults[1], first.Document.SelectedSearchResult);
                    window.KeyPress(Key.Enter, RawInputModifiers.Shift, PhysicalKey.None, null);
                    Assert.Same(first.Document.SearchResults[0], first.Document.SelectedSearchResult);
                    Click(view, "Previous match");
                    Assert.Same(first.Document.SearchResults.Last(), first.Document.SelectedSearchResult);
                    Click(view, "Next match");
                    Assert.Same(first.Document.SearchResults[0], first.Document.SelectedSearchResult);
                    window.KeyPress(Key.Escape, RawInputModifiers.None, PhysicalKey.None, null);
                    Assert.Empty(first.Document.SearchResults);
                    Assert.All(first.Document.Pages, page => Assert.Empty(page.SearchHighlights));
                    Assert.Same(view.FindControl<ScrollViewer>("PageScroll"), window.FocusManager!.GetFocusedElement());

                    await controller.OpenSampleAsync();
                    var second = controller.Tabs.SelectedTab!;
                    Layout(window, 390, 844);
                    Click(view, "Open documents");
                    Layout(window, 390, 844);
                    var list = view.FindControl<ListBox>("DocumentList")!;
                    Capture(window, "mobile-documents-phone");
                    var item = Assert.IsType<ListBoxItem>(list.ContainerFromIndex(0));
                    Point point = item.TranslatePoint(new Point(60, 24), window)!.Value;
                    window.MouseDown(point, MouseButton.Left);
                    window.MouseUp(point, MouseButton.Left);
                    Assert.Same(first, controller.Tabs.SelectedTab);
                    Assert.False(view.FindControl<Border>("SheetScrim")!.IsVisible);
                    window.KeyPress(Key.Tab, RawInputModifiers.Control, PhysicalKey.None, null);
                    Assert.Same(second, controller.Tabs.SelectedTab);
                    second.Document.EditorText = "An edit saved through the mobile toolbar";
                    await second.Document.ApplyPageMarkupAsync(PdfEditorTool.Note, new PdfEditorGesture(1, 24, 24, 48, 48, []));
                    Assert.Contains("Unsaved", view.FindControl<TextBlock>("DocumentSubtitle")!.Text);
                    Click(view, "Save");
                    using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(30));
                    while (second.Document.SaveCommand.IsRunning) await Task.Delay(25, timeout.Token);
                    Assert.False(second.Document.IsDirty, second.Document.ErrorMessage);
                    Assert.Equal("Saved on this device", view.FindControl<TextBlock>("DocumentSubtitle")!.Text);
                    Layout(window, 390, 844);
                    while (second.Document.SelectedPage!.IsRendering) await Task.Delay(25, timeout.Token);
                    Capture(window, "mobile-polished-phone");
                    Layout(window, 844, 390);
                    Click(view, "Open documents");
                    Capture(window, "mobile-documents-landscape");
                    Click(view, "Done");
                    Layout(window, 1024, 768);
                    while (second.Document.SelectedPage!.IsRendering) await Task.Delay(25, timeout.Token);
                    Capture(window, "mobile-polished-ipad");
                    window.RequestedThemeVariant = ThemeVariant.Dark;
                    Layout(window, 1024, 768);
                    Capture(window, "mobile-polished-ipad-dark");
                } finally { window.Close(); }
                return true;
            }, CancellationToken.None);
        } finally { if (Directory.Exists(root)) Directory.Delete(root, true); }
    }
}
