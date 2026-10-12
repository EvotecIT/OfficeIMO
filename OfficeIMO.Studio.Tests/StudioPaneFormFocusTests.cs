using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Styling;
using Avalonia.Threading;
using Avalonia.VisualTree;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workspace;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioPaneFormFocusTests {
    [Theory]
    [InlineData(1600)]
    [InlineData(390)]
    public async Task ClickingSecondRepeatedWidgetKeepsItsFocusAndSharedDraftThroughSave(int width) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            Application.Current!.RequestedThemeVariant = width == 390 ? ThemeVariant.Dark : ThemeVariant.Light;
            var services = ((App)Application.Current!).Services;
            string source = CreateFixture(services.Paths.Root);
            var window = new MainWindow(services) { Width = width, Height = 900 };
            try {
                window.Show(); await window.TabHost.OpenDocumentAsync(source);
                var document = window.ViewModel;
                document.ShowFormsModeCommand.Execute(null);
                document.SelectedFormField = Assert.Single(document.FormFields);
                var panes = window.TabHost.Panes;
                panes.OpenSecondPane(window.TabHost.SelectedTab);
                var left = panes.Left!; var right = panes.Right!;
                await Render(window, left, right);
                HideInspector(window);
                document.SelectedFormField = null;
                await Flush(window);
                Assert.Null(left.SelectedPage!.FormAnchorObjectNumber);
                Assert.Null(right.SelectedPage!.FormAnchorObjectNumber);
                var rightView = PaneView(window, right);
                var second = Editor(rightView, 7);
                Click(window, second);
                await Flush(window);
                Capture(window, $"pane-form-second-widget-{width}");
                Assert.True(second.IsFocused, $"Canonical widget={document.SelectedPage!.FormAnchorObjectNumber}; pane widget={right.SelectedPage!.FormAnchorObjectNumber}");
                Assert.Equal(7, right.SelectedPage!.FormAnchorObjectNumber);
                Assert.Equal(7, left.SelectedPage!.FormAnchorObjectNumber);
                Assert.Equal(7, document.SelectedPage!.FormAnchorObjectNumber);
                Assert.Null(document.SelectedPage.Scene);
                window.KeyTextInput("Pane shared draft");
                Assert.Equal("Pane shared draft", Editor(rightView, 6).Text);
                Assert.Equal("Pane shared draft", document.SelectedFormField!.TextValue);
                if (width == 1600) Assert.Equal("Pane shared draft", Editor(PaneView(window, left), 7).Text);
                Capture(window, $"pane-form-shared-draft-{width}");
                await document.SaveCommand.ExecuteAsync(null);
                Assert.Null(document.ErrorMessage); Assert.False(document.IsDirty);
                var saved = Assert.Single(PdfDocument.Load(source).Inspect().FormFields);
                Assert.Equal("Pane shared draft", saved.Value); Assert.Equal(3, saved.WidgetCount);
                await Render(window, left, right);
                await DismissToast(window);
                Assert.Same(left.Document, right.Document); Assert.Single(window.TabHost.Tabs);
                right.ActivatePage(2);
                await Render(window, right);
                HideInspector(window);
                var rotated = Editor(PaneView(window, right), 8);
                rotated.BringIntoView(); await Flush(window);
                Click(window, rotated); await Flush(window);
                Capture(window, $"pane-form-rotated-widget-{width}");
                Assert.True(rotated.IsFocused);
                Assert.Equal(8, right.SelectedPage!.FormAnchorObjectNumber);
                Assert.Null(left.SelectedPage!.FormAnchorObjectNumber);
                Assert.Equal("Pane shared draft", rotated.Text);
            } finally { window.Close(); }
            return true;
        }, default);
    }

    [Theory]
    [InlineData(1600, false)]
    [InlineData(390, false)]
    [InlineData(1600, true)]
    [InlineData(390, true)]
    public async Task F6CancelsQueuedFormFocusInThePaneThatLosesActivation(int width, bool separateDocuments) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            string source = CreateFixture(services.Paths.Root);
            Application.Current!.RequestedThemeVariant = width == 390 ? ThemeVariant.Dark : ThemeVariant.Light;
            var window = new MainWindow(services) { Width = width, Height = 900 };
            try {
                window.Show(); await window.TabHost.OpenDocumentAsync(source);
                window.ViewModel.ShowFormsModeCommand.Execute(null);
                window.ViewModel.SelectedFormField = Assert.Single(window.ViewModel.FormFields);
                var panes = window.TabHost.Panes; panes.OpenSecondPane(window.TabHost.SelectedTab);
                if (separateDocuments) {
                    string other = Path.Combine(services.Paths.Root, "other-shared-widgets.pdf");
                    File.Copy(source, other, true);
                    await window.TabHost.OpenDocumentAsync(other);
                    window.ViewModel.ShowFormsModeCommand.Execute(null);
                    window.ViewModel.SelectedFormField = Assert.Single(window.ViewModel.FormFields);
                }
                var left = panes.Left!; var right = panes.Right!;
                await Render(window, left, right); HideInspector(window);
                var editor = Editor(PaneView(window, right), 6);
                Click(window, editor); await Flush(window);
                Assert.Same(right, panes.ActivePane);
                // Dispatch the two routed keys before Loaded callbacks run, as queued native keys can.
                editor.RaiseEvent(new KeyEventArgs { RoutedEvent = InputElement.KeyDownEvent, Key = Key.Tab });
                Assert.True(right.SelectedPage!.FocusInlineFormEditorRequested);
                editor.RaiseEvent(new KeyEventArgs { RoutedEvent = InputElement.KeyDownEvent, Key = Key.F6 });
                Assert.Same(left, panes.ActivePane);
                bool inactiveRequest = right.SelectedPage.FocusInlineFormEditorRequested;
                Capture(window, $"pane-form-f6-pending-focus-{width}-{separateDocuments}");
                Assert.False(inactiveRequest);
                await Flush(window);
                Assert.Same(left, panes.ActivePane);
                Capture(window, $"pane-form-f6-focus-settled-{width}-{separateDocuments}");
                var leftView = PaneView(window, left);
                var reader = leftView.FindControl<ListBox>("PanePages")!;
                var firstFocus = window.FocusManager!.GetFocusedElement();
                bool firstReaderFocused = ReferenceEquals(reader, firstFocus);
                Assert.DoesNotContain(PaneView(window, right).GetVisualDescendants().OfType<TextBox>(), editor => editor.IsFocused);
                window.KeyPress(Key.F6, RawInputModifiers.None, PhysicalKey.None, null);
                await Flush(window);
                Capture(window, $"pane-form-f6-return-focus-{width}-{separateDocuments}");
                Assert.Same(right, panes.ActivePane);
                var rightReader = PaneView(window, right).FindControl<ListBox>("PanePages")!;
                Assert.True(firstReaderFocused,
                    $"First switch focus={firstFocus}; returning focus={window.FocusManager.GetFocusedElement()}");
                Assert.Same(rightReader, window.FocusManager.GetFocusedElement());
                Assert.NotNull(right.Document.SelectedFormField); // Remembering the field does not request editor focus.
                Assert.False(right.Document.SelectedPage!.FocusInlineFormEditorRequested);
                Assert.False(right.SelectedPage!.FocusInlineFormEditorRequested);
            } finally { window.Close(); }
            return true;
        }, default);
    }

    [Theory]
    [InlineData(1600, false)]
    [InlineData(390, false)]
    [InlineData(1600, true)]
    [InlineData(390, true)]
    public async Task SwitchButtonFocusesReaderAfterFormsContextRestoration(int width, bool separateDocuments) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            string source = CreateFixture(services.Paths.Root);
            Application.Current!.RequestedThemeVariant = width == 390 ? ThemeVariant.Dark : ThemeVariant.Light;
            var window = new MainWindow(services) { Width = width, Height = 900 };
            try {
                window.Show(); await window.TabHost.OpenDocumentAsync(source);
                window.ViewModel.ShowFormsModeCommand.Execute(null);
                window.ViewModel.SelectedFormField = Assert.Single(window.ViewModel.FormFields);
                var panes = window.TabHost.Panes; panes.OpenSecondPane(window.TabHost.SelectedTab);
                if (separateDocuments) {
                    string other = Path.Combine(services.Paths.Root, "other-switch-widgets.pdf");
                    File.Copy(source, other, true);
                    await window.TabHost.OpenDocumentAsync(other);
                    window.ViewModel.ShowFormsModeCommand.Execute(null);
                    window.ViewModel.SelectedFormField = Assert.Single(window.ViewModel.FormFields);
                }
                var left = panes.Left!; var right = panes.Right!;
                await Render(window, left, right); HideInspector(window);
                Click(window, Editor(PaneView(window, right), 7)); await Flush(window);
                var workspace = window.FindControl<DocumentWorkspaceView>("DocumentWorkspace")!;
                var surface = workspace.FindControl<IndependentDocumentPanesView>("IndependentPanes")!;
                var button = surface.FindControl<Button>("SwitchPaneButton")!;
                Click(window, button); await Flush(window);
                Capture(window, $"pane-form-switch-button-{width}-{separateDocuments}");
                Assert.Same(left, panes.ActivePane);
                if (width == 390) Assert.False(workspace.FindControl<Grid>("InspectorPane")!.IsVisible);
                Assert.Same(PaneView(window, left).FindControl<ListBox>("PanePages"), window.FocusManager!.GetFocusedElement());
                Click(window, button); await Flush(window);
                Capture(window, $"pane-form-switch-button-return-{width}-{separateDocuments}");
                Assert.Same(right, panes.ActivePane);
                Assert.Same(PaneView(window, right).FindControl<ListBox>("PanePages"), window.FocusManager.GetFocusedElement());
                Assert.NotNull(right.Document.SelectedFormField);
                Assert.False(right.Document.SelectedPage!.FocusInlineFormEditorRequested);
            } finally { window.Close(); }
            return true;
        }, default);
    }

    [Fact]
    public async Task ConsumedFormFocusIsNotReplayedByPaneZoom() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            string source = CreateFixture(services.Paths.Root);
            var window = new MainWindow(services) { Width = 1600, Height = 900 };
            try {
                window.Show(); await window.TabHost.OpenDocumentAsync(source);
                window.ViewModel.ShowFormsModeCommand.Execute(null);
                var panes = window.TabHost.Panes; panes.OpenSecondPane(window.TabHost.SelectedTab);
                var right = panes.Right!;
                await Render(window, panes.Left!, right); HideInspector(window);
                var view = PaneView(window, right);
                Click(window, Editor(view, 6)); await Flush(window);
                var zoom = view.GetVisualDescendants().OfType<Button>()
                    .Single(button => ReferenceEquals(button.Command, right.ZoomOutCommand));
                Assert.True(zoom.Focus(NavigationMethod.Tab));
                Assert.True(zoom.IsFocused);
                zoom.RaiseEvent(new KeyEventArgs { RoutedEvent = InputElement.KeyDownEvent, Key = Key.Enter });
                await Flush(window);
                Capture(window, "pane-form-zoom-keeps-focus");
                Assert.True(zoom.IsFocused, "A consumed form-focus request stole focus from the pane zoom button.");
                Assert.False(window.ViewModel.SelectedPage!.FocusInlineFormEditorRequested);
                right.SelectedLayout = right.LayoutChoices.Single(choice => choice.Mode == ReaderLayoutMode.Grid);
                await Render(window, right);
                window.KeyPress(Key.F6, RawInputModifiers.None, PhysicalKey.None, null);
                window.KeyPress(Key.F6, RawInputModifiers.None, PhysicalKey.None, null);
                Assert.Same(right, panes.ActivePane);
                Assert.Same(view.FindControl<ListBox>("PaneGridPages"), window.FocusManager!.GetFocusedElement());
                Capture(window, "pane-form-grid-focus");
            } finally { window.Close(); }
            return true;
        }, default);
    }

    private static string CreateFixture(string root) {
        Directory.CreateDirectory(root);
        string path = Path.Combine(root, "shared-widgets.pdf");
        File.WriteAllText(path, StudioFormWidgetFixture.SharedTextWidgets(), System.Text.Encoding.ASCII);
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (!string.IsNullOrWhiteSpace(output)) { Directory.CreateDirectory(output); File.Copy(path, Path.Combine(output, "pane-shared-widgets.pdf"), true); }
        return path;
    }
    private static IndependentDocumentPaneView PaneView(Window window, StudioDocumentPaneViewModel pane) =>
        window.GetVisualDescendants().OfType<IndependentDocumentPaneView>().Single(view => ReferenceEquals(view.DataContext, pane));
    private static TextBox Editor(IndependentDocumentPaneView view, int objectNumber) =>
        view.GetVisualDescendants().OfType<PdfInlineFormWidgetView>()
            .Single(editor => editor.DataContext is PdfInlineFormWidgetViewModel widget && widget.ObjectNumber == objectNumber)
            .FindControl<TextBox>("InlineFormText")!;
    private static void Click(Window window, Control control) {
        control.BringIntoView(); window.UpdateLayout();
        Point point = control.TranslatePoint(new Point(control.Bounds.Width / 2, control.Bounds.Height / 2), window)!.Value;
        window.MouseDown(point, MouseButton.Left); window.MouseUp(point, MouseButton.Left);
    }
    private static void HideInspector(MainWindow window) {
        var view = window.FindControl<DocumentWorkspaceView>("DocumentWorkspace")!;
        var toggle = view.FindControl<ToggleButton>("InspectorToggle")!;
        if (toggle.IsChecked == true) Click(window, toggle);
    }
    private static async Task DismissToast(MainWindow window) {
        var toast = window.FindControl<Border>("OperationToast")!;
        if (!toast.IsVisible) return;
        Click(window, toast.GetVisualDescendants().OfType<Button>().Single(button => Grid.GetColumn(button) == 3));
        for (int pass = 0; pass < 50 && toast.IsVisible; pass++) {
            await Task.Delay(20); await Flush(window);
        }
        Assert.False(toast.IsVisible);
    }
    private static async Task Render(Window window, params StudioDocumentPaneViewModel[] panes) {
        for (int pass = 0; pass < 2; pass++) {
            await Flush(window);
            foreach (var pane in panes) await pane.SelectedPage!.EnsureRenderedAsync();
        }
        await Flush(window);
    }
    private static async Task Flush(Window window) {
        await Dispatcher.UIThread.InvokeAsync(window.UpdateLayout, DispatcherPriority.Background);
        AvaloniaHeadlessPlatform.ForceRenderTimerTick();
    }
    private static void Capture(Window window, string name) {
        using var frame = window.CaptureRenderedFrame(); Assert.NotNull(frame);
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(output)) return;
        Directory.CreateDirectory(output); frame.Save(Path.Combine(output, name + ".png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
