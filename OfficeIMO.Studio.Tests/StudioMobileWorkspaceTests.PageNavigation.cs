using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Input.Raw;
using Avalonia.Platform.Storage;
using Avalonia.Styling;
using Avalonia.Threading;
using Avalonia.VisualTree;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Mobile;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Workspace;
using OfficeIMO.Studio.Infrastructure;

namespace OfficeIMO.Studio.Tests;

public sealed partial class StudioMobileWorkspaceTests {
    [Theory]
    [InlineData(1366, 1024)]
    [InlineData(390, 844)]
    [InlineData(844, 390)]
    [InlineData(320, 568)]
    public async Task PageChooserValidatesInputRestoresFocusAndKeepsDocumentIdentity(int width, int height) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-page-navigation-" + Guid.NewGuid().ToString("N"));
        try {
            using var session = TestAppBuilder.StartSession();
            await session.Dispatch(async () => {
                var services = StudioApplicationServices.Create(new StudioDataPaths(Path.Combine(root, "data")),
                    new StudioLocalDocumentRoot(Path.Combine(root, "documents")));
                using var controller = new MobileDocumentController(services, _ => Task.FromResult<IStorageFile?>(null), _ => Task.CompletedTask);
                var view = new MobileWorkspaceView();
                view.Connect(controller);
                var window = new Window { Content = view };
                try {
                    window.Show();
                    await controller.Tabs.OpenDocumentAsync(CreateNavigationDocument(root));
                    var document = controller.Document;
                    Layout(window, width, height);
                    var opener = view.FindControl<Button>("GoToPageButton")!;
                    Assert.True(opener.Bounds.Height >= 44);
                    Assert.True(opener.Bounds.Width >= 44);
                    await OpenPageChooserAsync(view, window, width, height);
                    Layout(window, width, height);
                    var dialog = Assert.IsType<PageNavigationDialogContent>(view.FindControl<ContentControl>("DialogContent")!.Content);
                    var input = dialog.FindControl<TextBox>("PageInput")!;
                    Assert.Same(input, window.FocusManager!.GetFocusedElement());
                    Assert.Equal("1", input.SelectedText);
                    Assert.False(view.FindControl<SplitView>("ApplicationNavigation")!.IsEnabled);
                    foreach (string invalid in new[] { "", "0", "4", "1.5", "word", "999999999999" }) {
                        input.Text = invalid;
                        Layout(window, width, height);
                        Assert.False(dialog.FindControl<Button>("GoButton")!.IsEnabled);
                        window.KeyPress(Key.Enter, RawInputModifiers.None, PhysicalKey.None, null);
                        Assert.Equal(1, document.SelectedPage!.PageNumber);
                        Assert.True(view.FindControl<Border>("DialogScrim")!.IsVisible);
                    }
                    Capture(window, $"page-chooser-invalid-{width}");
                    input.Text = "3";
                    Layout(window, width, height);
                    Capture(window, $"page-chooser-{width}");
                    window.KeyPress(Key.Enter, RawInputModifiers.None, PhysicalKey.None, null);
                    await WaitForPageChooserAsync(view);
                    Layout(window, width, height);
                    Assert.Equal(3, document.SelectedPage!.PageNumber);
                    Assert.Same(opener, window.FocusManager.GetFocusedElement());
                    Assert.True(view.FindControl<SplitView>("ApplicationNavigation")!.IsEnabled);

                    await OpenPageChooserAsync(view, window, width, height);
                    Layout(window, width, height);
                    dialog = (PageNavigationDialogContent)view.FindControl<ContentControl>("DialogContent")!.Content!;
                    dialog.FindControl<TextBox>("PageInput")!.Text = "2";
                    window.KeyPress(Key.Escape, RawInputModifiers.None, PhysicalKey.None, null);
                    await WaitForPageChooserAsync(view);
                    Assert.Equal(3, document.SelectedPage.PageNumber);

                    await OpenPageChooserAsync(view, window, width, height);
                    Layout(window, width, height);
                    dialog = (PageNavigationDialogContent)view.FindControl<ContentControl>("DialogContent")!.Content!;
                    dialog.FindControl<TextBox>("PageInput")!.Text = "2";
                    dialog.FindControl<Button>("CancelButton")!.Focus();
                    window.KeyPress(Key.Enter, RawInputModifiers.None, PhysicalKey.None, null);
                    await WaitForPageChooserAsync(view);
                    Assert.Equal(3, document.SelectedPage.PageNumber);

                    await OpenPageChooserAsync(view, window, width, height);
                    Layout(window, width, height);
                    dialog = (PageNavigationDialogContent)view.FindControl<ContentControl>("DialogContent")!.Content!;
                    await controller.OpenSampleAsync();
                    dialog.FindControl<TextBox>("PageInput")!.Text = "2";
                    dialog.FindControl<Button>("GoButton")!.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));
                    await WaitForPageChooserAsync(view);
                    Assert.Equal(1, controller.Document.SelectedPage!.PageNumber);
                    Assert.Equal(3, document.SelectedPage.PageNumber);

                    window.RequestedThemeVariant = ThemeVariant.Dark;
                    await OpenPageChooserAsync(view, window, width, height);
                    Layout(window, width, height);
                    Capture(window, $"page-chooser-dark-{width}");
                    window.Close();
                    await WaitForPageChooserAsync(view);
                } finally { window.Close(); }
                return true;
            }, CancellationToken.None);
        } finally { if (Directory.Exists(root)) Directory.Delete(root, true); }
    }

    [Fact]
    public async Task DesktopPageEntryKeepsInvalidInputAvailableForCorrection() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-desktop-page-" + Guid.NewGuid().ToString("N"));
        try {
            using var session = TestAppBuilder.StartSession();
            await session.Dispatch(async () => {
                var services = StudioApplicationServices.Create(new StudioDataPaths(Path.Combine(root, "data")),
                    new StudioLocalDocumentRoot(Path.Combine(root, "documents")));
                using var controller = new MobileDocumentController(services, _ => Task.FromResult<IStorageFile?>(null), _ => Task.CompletedTask);
                await controller.Tabs.OpenDocumentAsync(CreateNavigationDocument(root));
                var document = controller.Document;
                var view = new DocumentWorkspaceView { DataContext = document };
                var window = new Window { Content = view };
                try {
                    window.Show();
                    Layout(window, 1280, 900);
                    var input = view.FindControl<TextBox>("PageNumberBox")!;
                    input.Focus();
                    input.Text = "4";
                    window.KeyPress(Key.Enter, RawInputModifiers.None, PhysicalKey.None, null);
                    Assert.Same(input, window.FocusManager!.GetFocusedElement());
                    Assert.True(view.FindControl<Border>("PageNumberErrorBanner")!.IsVisible);
                    Assert.Equal("4", input.Text);
                    Assert.Equal(1, document.SelectedPage!.PageNumber);
                    Layout(window, 1280, 900);
                    Capture(window, "desktop-page-invalid");
                    input.Text = "3";
                    window.KeyPress(Key.Enter, RawInputModifiers.None, PhysicalKey.None, null);
                    Assert.Equal(3, document.SelectedPage.PageNumber);
                    Assert.False(view.FindControl<Border>("PageNumberErrorBanner")!.IsVisible);
                    Assert.False(input.IsFocused);
                    input.Focus();
                    input.Text = "0";
                    window.KeyPress(Key.Enter, RawInputModifiers.None, PhysicalKey.None, null);
                    window.KeyPress(Key.Escape, RawInputModifiers.None, PhysicalKey.None, null);
                    Assert.False(view.FindControl<Border>("PageNumberErrorBanner")!.IsVisible);
                    Assert.Equal("3", input.Text);
                    Assert.Equal(3, document.SelectedPage.PageNumber);
                } finally { window.Close(); }
                return true;
            }, CancellationToken.None);
        } finally { if (Directory.Exists(root)) Directory.Delete(root, true); }
    }

    private static async Task OpenPageChooserAsync(MobileWorkspaceView view, Window window, int width, int height) {
        var button = view.FindControl<Button>("GoToPageButton")!;
        Point point = await StudioHeadlessInput.WaitForTargetAsync(window, button, () => Layout(window, width, height));
        window.MouseDown(point, MouseButton.Left);
        window.MouseUp(point, MouseButton.Left);
    }

    private static async Task WaitForPageChooserAsync(MobileWorkspaceView view) {
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(5));
        while (view.FindControl<Border>("DialogScrim")!.IsVisible) await Task.Delay(10, timeout.Token);
        await Avalonia.Threading.Dispatcher.UIThread.InvokeAsync(() => { }, Avalonia.Threading.DispatcherPriority.Background);
    }

    private static string CreateNavigationDocument(string root) {
        Directory.CreateDirectory(root);
        string path = Path.Combine(root, "Three pages.pdf");
        File.WriteAllBytes(path, PdfDocument.Create(pdf => {
            for (int page = 1; page <= 3; page++) {
                string label = "Page " + page;
                pdf.Page(item => item.Content(content => content.Text(label)));
            }
        }).ToBytes());
        return path;
    }
}
