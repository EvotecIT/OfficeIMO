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

public sealed partial class StudioMobileWorkspaceTests {
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
                        if (draftEnabledDuringSubmission is null && e.PropertyName == nameof(document.IsWorkspaceBusy) && document.IsWorkspaceBusy) {
                            draftEnabledDuringSubmission = draft.IsEffectivelyEnabled;
                            window.KeyPress(Key.Escape, RawInputModifiers.None, PhysicalKey.None, null);
                            Assert.True(view.FindControl<Border>("SheetScrim")!.IsVisible);
                        }
                    };
                    Click(view, "Add note");
                    Assert.Equal(false, draftEnabledDuringSubmission);
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
                    Button pages = view.GetVisualDescendants().OfType<Button>().Single(button => Avalonia.Automation.AutomationProperties.GetName(button) == "Pages");
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

    [Fact]
    public async Task TabsRetainIndependentPagesEditsAndRecovery() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-mobile-tabs-" + Guid.NewGuid().ToString("N"));
        try {
            using var session = TestAppBuilder.StartSession();
            await session.Dispatch(async () => {
                var paths = new StudioDataPaths(Path.Combine(root, "data"));
                var local = new StudioLocalDocumentRoot(Path.Combine(root, "Documents"));
                var services = StudioApplicationServices.Create(paths, local);
                byte[] original = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "openpreserve-pdfa1b-text.pdf"));
                var file = new TestStorageFile("content://files/review.pdf", original, "Review.pdf");
                using (var controller = new MobileDocumentController(services, _ => Task.FromResult<Avalonia.Platform.Storage.IStorageFile?>(file.Item), _ => Task.CompletedTask)) {
                    var view = new MobileWorkspaceView();
                    view.Connect(controller);
                    var window = new Window { Content = view };
                    try {
                        window.Show();
                        Layout(window, 1024, 768);
                        await controller.OpenSampleAsync();
                        var first = controller.Tabs.SelectedTab!;
                        Assert.Equal(3, first.Document.Pages.Count);
                        first.Document.SelectedPage = first.Document.Pages[1];
                        first.Document.SetTouchZoom(1.25);
                        Click(view, "Note");
                        view.FindControl<TextBox>("NoteText")!.Text = "A draft for the first document";
                        Click(view, "Done");
                        await controller.Document.OpenCommand.ExecuteAsync(null);
                        var second = controller.Tabs.SelectedTab!;
                        Assert.Equal(2, controller.Tabs.Tabs.Count);
                        Assert.NotSame(first, second);
                        Assert.Equal(string.Empty, view.FindControl<TextBox>("NoteText")!.Text);
                        second.Document.EditorText = "Review this second document";
                        await second.Document.ApplyPageMarkupAsync(OfficeIMO.Studio.Features.Editor.PdfEditorTool.Note,
                            new OfficeIMO.Studio.Features.Editor.PdfEditorGesture(1, 24, 24, 48, 48, []));
                        Assert.True(second.IsDirty, second.Document.ErrorMessage);
                        controller.Tabs.SelectedTab = first;
                        Assert.Same(first.Document, view.Document);
                        Assert.Equal(2, first.Document.SelectedPage!.PageNumber);
                        Assert.Equal(1.25, first.Document.Zoom);
                        Assert.Equal("A draft for the first document", view.FindControl<TextBox>("NoteText")!.Text);
                        Assert.False(first.IsDirty);
                        Layout(window, 1024, 768);
                        first.Document.FitPageCommand.Execute(null);
                        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(30));
                        while (first.Document.SelectedPage.IsRendering || first.Document.OrganizerPages.Any(page => page.IsLoading))
                            await Task.Delay(25, timeout.Token);
                        Layout(window, 1024, 768);
                        Capture(window, "mobile-tabs-ipad");
                        Layout(window, 390, 844);
                        Capture(window, "mobile-tabs-phone");
                        first.Document.SetTouchZoom(1.25);
                        controller.Suspend();
                    } finally { window.Close(); }
                }
                using var restored = new MobileDocumentController(StudioApplicationServices.Create(paths, local),
                    _ => Task.FromResult<Avalonia.Platform.Storage.IStorageFile?>(null), _ => Task.CompletedTask);
                await restored.RestoreAsync();
                Assert.Equal(2, restored.Tabs.Tabs.Count);
                Assert.Same(restored.Tabs.Tabs[0], restored.Tabs.SelectedTab);
                Assert.Equal(2, restored.Document.SelectedPage!.PageNumber);
                Assert.Equal(1.25, restored.Document.Zoom);
                var dirtyTab = restored.Tabs.Tabs[1];
                Assert.True(dirtyTab.IsDirty, dirtyTab.Document.ErrorMessage);
                // A missing presenter must keep unsaved work open, never silently save or discard it.
                await dirtyTab.CloseCommand.ExecuteAsync(null);
                Assert.Equal(2, restored.Tabs.Tabs.Count);
                restored.ConfirmUnsavedChangesAsync = _ => Task.FromResult(UnsavedChangesDecision.Save);
                await dirtyTab.CloseCommand.ExecuteAsync(null);
                Assert.Single(restored.Tabs.Tabs);
                await restored.Tabs.CloseSelectedTabAsync();
                Assert.Empty(services.DocumentHistory.RestartSession.Load().Documents);
                return true;
            }, CancellationToken.None);
        } finally { if (Directory.Exists(root)) Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task WorkspaceReactivatesPresentationAfterViewportOrLifetimeChange(bool resumeSheet) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-mobile-transition-" + Guid.NewGuid().ToString("N"));
        try {
            using var session = TestAppBuilder.StartSession();
            await session.Dispatch(async () => {
                var services = StudioApplicationServices.Create(new StudioDataPaths(Path.Combine(root, "data")),
                    new StudioLocalDocumentRoot(Path.Combine(root, "Documents")));
                using var controller = new MobileDocumentController(services,
                    _ => Task.FromResult<Avalonia.Platform.Storage.IStorageFile?>(null), _ => Task.CompletedTask);
                var view = new MobileWorkspaceView();
                view.Connect(controller);
                var window = new Window { Content = view };
                try {
                    window.Show();
                    Layout(window, 1024, 768);
                    await controller.OpenSampleAsync();
                    var first = controller.Tabs.SelectedTab!;
                    first.Document.SelectedPage = first.Document.Pages[1];
                    if (resumeSheet) {
                        Layout(window, 390, 844);
                        Click(view, "Pages");
                        controller.Suspend();
                        controller.Resume();
                        Layout(window, 390, 844);
                        var pages = Assert.IsType<ListBox>(view.FindControl<ContentControl>("SheetPages")!.Content);
                        Assert.Equal(3, pages.ItemCount);
                        Assert.Equal(2, Assert.IsType<OfficeIMO.Studio.Features.Organizer.PdfOrganizerPageViewModel>(pages.SelectedItem).PageNumber);
                        Assert.True(pages.IsEffectivelyVisible);
                        Click(view, "Done");
                    } else {
                        await controller.OpenSampleAsync();
                        Layout(window, 390, 844);
                        var tabs = view.FindControl<TabControl>("MobileTabs")!;
                        Assert.Equal(Avalonia.Automation.Peers.AutomationControlType.Tab,
                            Avalonia.Automation.Peers.ControlAutomationPeer.CreatePeerForElement(tabs)!.GetAutomationControlType());
                        var container = Assert.IsType<TabItem>(tabs.ContainerFromIndex(0));
                        container.BringIntoView();
                        Layout(window, 390, 844);
                        Point point = container.TranslatePoint(new Point(40, 22), window)!.Value;
                        window.MouseDown(point, MouseButton.Left);
                        window.MouseUp(point, MouseButton.Left);
                        Assert.Same(first, controller.Tabs.SelectedTab);
                        first.Document.FitPageCommand.Execute(null);
                        Layout(window, 390, 844);
                        var viewport = view.FindControl<ScrollViewer>("PageScroll")!;
                        Assert.True(first.Document.SelectedPage!.DisplayWidth < viewport.Bounds.Width,
                            "The restored tab must fit the current viewport, not the viewport from before rotation.");
                        Assert.True(first.Document.SelectedPage.DisplayHeight < viewport.Bounds.Height);
                    }
                } finally { window.Close(); }
                return true;
            }, CancellationToken.None);
        } finally { if (Directory.Exists(root)) Directory.Delete(root, true); }
    }

    private static void Click(MobileWorkspaceView view, string label) {
        var window = (Window)TopLevel.GetTopLevel(view)!;
        window.UpdateLayout();
        Dispatcher.UIThread.RunJobs();
        Button button = view.GetVisualDescendants().OfType<Button>().Single(button => (Equals(button.Content, label) || Avalonia.Automation.AutomationProperties.GetName(button) == label) && button.IsEffectivelyVisible && button.IsEffectivelyEnabled);
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
