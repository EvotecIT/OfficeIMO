using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Input.Raw;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Avalonia.VisualTree;
using OfficeIMO.Studio.Features.Mobile;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Infrastructure;
using Xunit;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioMobileWorkspaceTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task TouchWorkspaceAdaptsAndRetainsSharedEdits(bool recoveryUnavailable) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-mobile-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            using var session = TestAppBuilder.StartSession();
            await session.Dispatch(async () => {
                string path = Path.Combine(root, "Review.pdf");
                File.Copy(Path.Combine(AppContext.BaseDirectory, "Fixtures", "openpreserve-pdfa1b-text.pdf"), path);
                var services = StudioApplicationServices.Create(new StudioDataPaths(Path.Combine(root, "data")));
                Directory.CreateDirectory(services.Paths.Root);
                if (recoveryUnavailable) File.WriteAllText(services.Paths.RecoveryRoot, "Unavailable recovery folder");
                using var document = new MainWindowViewModel(_ => Task.FromResult<string?>(path), services: services);
                var view = new MobileWorkspaceView { DataContext = document };
                var window = new Window { Content = view };
                try {
                    window.Show();
                    await document.OpenCommand.ExecuteAsync(null);
                    Assert.True(document.HasDocument, document.ErrorMessage);
                    Layout(window, 1024, 768);
                    Assert.Equal(OfficeIMO.Studio.Features.Reader.ReaderLayoutMode.SinglePage, document.ReaderLayout);
                    Assert.True(view.FindControl<Border>("PageSidebar")!.IsEffectivelyVisible);
                    Click(view, "Note");
                    var draft = view.FindControl<TextBox>("NoteText")!;
                    draft.Text = "Mobile review note";
                    bool? draftEnabledDuringSubmission = null;
                    document.PropertyChanged += (_, e) => {
                        if (e.PropertyName == nameof(document.IsWorkspaceBusy) && document.IsWorkspaceBusy)
                            draftEnabledDuringSubmission = draft.IsEffectivelyEnabled;
                    };
                    Click(view, "Add note");
                    Assert.Equal(false, draftEnabledDuringSubmission);
                    if (!draft.IsEffectivelyEnabled) {
                        window.KeyPress(Key.Escape, RawInputModifiers.None, PhysicalKey.None, null);
                        Assert.True(view.FindControl<Border>("SheetScrim")!.IsVisible);
                    }
                    await WaitForSubmissionAsync(view);
                    if (recoveryUnavailable) {
                        Assert.True(document.HasError);
                        Assert.True(view.FindControl<TextBlock>("SheetError")!.IsEffectivelyVisible);
                        Assert.Equal("Mobile review note", draft.Text);
                        Assert.False(document.IsDirty);
                        Layout(window, 390, 844);
                        Capture(window, "mobile-note-error");
                        File.Delete(services.Paths.RecoveryRoot);
                        Click(view, "Add note");
                        await WaitForSubmissionAsync(view);
                    }
                    Assert.False(document.HasError, document.ErrorMessage);
                    Assert.True(document.IsDirty);
                    Assert.True(document.CanUndo);
                    Assert.Equal(string.Empty, draft.Text);
                    Assert.False(view.FindControl<Border>("SheetScrim")!.IsVisible);

                    Layout(window, 390, 844);
                    Assert.False(view.FindControl<Border>("PageSidebar")!.IsEffectivelyVisible);
                    Button pages = view.GetVisualDescendants().OfType<Button>().Single(button => Equals(button.Content, "Pages"));
                    pages.Focus();
                    pages.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                    Assert.True(view.FindControl<Border>("SheetScrim")!.IsVisible);
                    Assert.False(view.FindControl<Grid>("WorkspaceGrid")!.IsEnabled);
                    window.KeyPress(Key.Escape, RawInputModifiers.None, PhysicalKey.None, null);
                    Assert.False(view.FindControl<Border>("SheetScrim")!.IsVisible);
                    Assert.Same(pages, window.FocusManager!.GetFocusedElement());
                    Assert.True(document.IsDirty);
                    await document.UndoCommand.ExecuteAsync(null);
                    Assert.False(document.IsDirty);
                    Assert.True(document.CanRedo);
                    double fitZoom = document.Zoom;
                    Click(view, "+");
                    Assert.True(document.Zoom > fitZoom);
                    Click(view, "Fit");
                    using var renderTimeout = new CancellationTokenSource(TimeSpan.FromSeconds(20));
                    while (document.SelectedPage!.IsRendering) await Task.Delay(25, renderTimeout.Token);
                    Assert.False(document.SelectedPage.HasRenderError, document.SelectedPage.RenderError);
                    Assert.True(document.SelectedPage.HasScene);
                    Layout(window, 390, 844);
                    Capture(window, "mobile-phone");
                    Layout(window, 1024, 768);
                    while (document.SelectedPage.IsRendering) await Task.Delay(25, renderTimeout.Token);
                    Layout(window, 1024, 768);
                    Capture(window, "mobile-ipad");
                } finally { window.Close(); }
                return true;
            }, CancellationToken.None);
        } finally { Directory.Delete(root, true); }
    }

    private static void Click(MobileWorkspaceView view, string label) {
        var window = (Window)TopLevel.GetTopLevel(view)!;
        window.UpdateLayout();
        Dispatcher.UIThread.RunJobs();
        Button button = view.GetVisualDescendants().OfType<Button>().Single(button => Equals(button.Content, label) && button.IsEffectivelyVisible);
        Assert.True(button.IsEffectivelyEnabled);
        Point point = button.TranslatePoint(new Point(button.Bounds.Width / 2, button.Bounds.Height / 2), window)!.Value;
        window.MouseDown(point, MouseButton.Left);
        window.MouseUp(point, MouseButton.Left);
    }

    private static async Task WaitForSubmissionAsync(MobileWorkspaceView view) {
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(20));
        while (!view.FindControl<StackPanel>("NotePanel")!.IsEnabled) await Task.Delay(25, timeout.Token);
    }

    private static void Layout(Window window, double width, double height) {
        window.Width = width;
        window.Height = height;
        window.Measure(new Size(width, height));
        window.Arrange(new Rect(0, 0, width, height));
        window.UpdateLayout();
        Dispatcher.UIThread.RunJobs();
    }

    private static void Capture(Window window, string name) {
        string? root = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(root)) return;
        Directory.CreateDirectory(root);
        using var frame = window.CaptureRenderedFrame();
        Assert.NotNull(frame);
        frame.Save(Path.Combine(root, name + ".png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
