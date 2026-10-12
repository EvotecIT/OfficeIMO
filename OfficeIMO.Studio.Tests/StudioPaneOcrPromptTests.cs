using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Styling;
using Avalonia.VisualTree;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workspace;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioPaneOcrPromptTests {
    [Theory]
    [InlineData(1600)]
    [InlineData(390)]
    public async Task ContextualPromptFollowsRenderedActivePageAndKeepsDocumentDismissalAndSceneLifetime(int width) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            Application.Current!.RequestedThemeVariant = width == 390 ? ThemeVariant.Dark : ThemeVariant.Light;
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string first = Path.Combine(services.Paths.Root, "First mixed scan.pdf");
            string second = Path.Combine(services.Paths.Root, "Second mixed scan.pdf");
            CreateMixedScan(first); CreateMixedScan(second);
            string? evidence = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
            if (!string.IsNullOrWhiteSpace(evidence)) {
                Directory.CreateDirectory(evidence); File.Copy(first, Path.Combine(evidence, "ocr-mixed-scan.pdf"), true);
            }
            var view = new DocumentWorkspaceView();
            using var host = new StudioDocumentTabHost(open => new MainWindowViewModel(
                _ => Task.FromResult<string?>(null), openDocumentInTab: open, services: services), document => view.DataContext = document);
            view.Panes = host.Panes;
            var window = new Window { Width = width, Height = 900, Content = view };
            try {
                window.Show(); await host.OpenDocumentAsync(first);
                var firstDocument = host.ActiveDocument;
                firstDocument.SelectedReaderLayoutChoice = firstDocument.ReaderLayoutChoices.Single(choice => choice.Mode == ReaderLayoutMode.SinglePage);
                await Render(window, firstDocument.SelectedPage!);
                Assert.False(firstDocument.ShowOcrPrompt);
                firstDocument.NextPageCommand.Execute(null);
                await Render(window, firstDocument.SelectedPage!);
                Assert.True(firstDocument.SelectedPage!.IsImageOnly);
                Assert.True(firstDocument.ShowOcrPrompt);
                Capture(window, $"ocr-prompt-ordinary-{width}");
                firstDocument.PreviousPageCommand.Execute(null);
                await Render(window, firstDocument.SelectedPage!);
                Assert.False(firstDocument.ShowOcrPrompt);

                host.Panes.OpenSecondPane(host.SelectedTab);
                var left = host.Panes.Left!; var right = host.Panes.Right!;
                right.ActivatePage(2);
                await Render(window, left.SelectedPage!, right.SelectedPage!);
                Assert.True(right.SelectedPage!.IsImageOnly);
                Assert.Null(firstDocument.Pages[1].Scene);
                Capture(window, $"ocr-prompt-pane-active-{width}");
                Assert.True(firstDocument.ShowOcrPrompt);
                var prompt = view.FindControl<Border>("OcrPrompt")!;
                Assert.True(prompt.IsEffectivelyVisible);
                var startOcr = prompt.GetVisualDescendants().OfType<Button>().Single(button => ReferenceEquals(button.Command, firstDocument.MakeSearchableCommand));
                Assert.True(startOcr.IsEffectivelyVisible); Assert.True(startOcr.IsEnabled);
                left.Activate();
                Assert.False(firstDocument.ShowOcrPrompt); // The inactive scanned pane must not suggest OCR for page one.
                right.Activate(); await Render(window, right.SelectedPage!); Assert.True(firstDocument.ShowOcrPrompt);
                right.PreviousCommand.Execute(null);
                await Render(window, right.SelectedPage!);
                Assert.False(firstDocument.ShowOcrPrompt);
                right.NextCommand.Execute(null);
                await Render(window, right.SelectedPage!);
                Assert.True(firstDocument.ShowOcrPrompt);

                var oldPage = right.SelectedPage!;
                firstDocument.ShowEditModeCommand.Execute(null);
                firstDocument.SelectedEditorToolChoice = firstDocument.EditorTools.Single(choice => choice.Tool == PdfEditorTool.AddText);
                firstDocument.EditorText = "Now selectable text";
                right.SelectedPage.CompleteEditorGesture(new(2, 60, 60, 200, 90, []));
                for (int attempt = 0; attempt < 200 && firstDocument.IsWorkspaceBusy; attempt++) await Task.Delay(10);
                Assert.False(firstDocument.IsWorkspaceBusy); Assert.Null(firstDocument.ErrorMessage);
                Assert.NotSame(oldPage, right.SelectedPage);
                Assert.False(firstDocument.ShowOcrPrompt);
                Assert.Null(oldPage.Scene);
                await Render(window, right.SelectedPage!);
                Assert.False(right.SelectedPage!.IsImageOnly);
                Assert.False(firstDocument.ShowOcrPrompt);
                await firstDocument.UndoCommand.ExecuteAsync(null);
                firstDocument.ShowViewModeCommand.Execute(null);
                await Render(window, right.SelectedPage!);
                Assert.True(firstDocument.ShowOcrPrompt);
                Assert.Null(firstDocument.Pages[1].Scene);

                window.Content = null;
                Assert.Null(right.SelectedPage!.Scene);
                Assert.False(firstDocument.ShowOcrPrompt);
                window.Content = view;
                await Render(window, right.SelectedPage!);
                Assert.True(firstDocument.ShowOcrPrompt);
                Capture(window, $"ocr-prompt-pane-reattached-{width}");
                host.Panes.ClosePane(left);
                await Render(window, firstDocument.SelectedPage!);
                Assert.True(firstDocument.SelectedPage!.IsImageOnly);
                Assert.True(firstDocument.ShowOcrPrompt); // Closing panes returns to the ordinary reader's presentation.

                host.Panes.OpenSecondPane(host.SelectedTab);
                left = host.Panes.Left!; right = host.Panes.Right!;
                await Render(window, right.SelectedPage!);
                Dismiss(window, view, firstDocument);
                Assert.False(firstDocument.ShowOcrPrompt);
                left.ActivatePage(1); right.ActivatePage(2);
                Assert.False(firstDocument.ShowOcrPrompt);
                await host.OpenDocumentAsync(second);
                right = host.Panes.Right!;
                var secondDocument = right.Document;
                Assert.NotSame(firstDocument, secondDocument);
                left.ActivatePage(2); right.ActivatePage(2);
                await Render(window, left.SelectedPage!, right.SelectedPage!);
                Capture(window, $"ocr-prompt-two-documents-{width}");
                Assert.Equal(2, right.SelectedPage!.PageNumber);
                Assert.True(right.SelectedPage.IsImageOnly, $"Active={ReferenceEquals(host.Panes.ActivePane, right)}; scene={right.SelectedPage.Scene is not null}; error={right.SelectedPage.RenderError}; document={right.Document.DocumentPath}");
                if (width == 1600) Assert.True(left.SelectedPage!.IsImageOnly);
                Assert.True(secondDocument.ShowOcrPrompt);
                Assert.Null(secondDocument.SelectedPage!.Scene);
                left.Activate(); Assert.Same(firstDocument, host.ActiveDocument); Assert.False(firstDocument.ShowOcrPrompt);
                right.Activate(); await Render(window, right.SelectedPage!); Assert.Same(secondDocument, host.ActiveDocument); Assert.True(secondDocument.ShowOcrPrompt);
                right.PreviousCommand.Execute(null);
                await Render(window, right.SelectedPage!);
                Assert.False(secondDocument.ShowOcrPrompt);
                right.NextCommand.Execute(null);
                await Render(window, right.SelectedPage!);
                Assert.True(secondDocument.ShowOcrPrompt);
                Dismiss(window, view, secondDocument);
                left.Activate(); right.Activate();
                Assert.False(secondDocument.ShowOcrPrompt);
            } finally { window.Close(); }
            return true;
        }, default);
    }

    private static void Dismiss(Window window, DocumentWorkspaceView view, MainWindowViewModel document) {
        window.UpdateLayout();
        var prompt = view.FindControl<Border>("OcrPrompt")!;
        Assert.True(prompt.IsEffectivelyVisible);
        var dismiss = prompt.GetVisualDescendants().OfType<Button>().Single(button => ReferenceEquals(button.Command, document.DismissOcrPromptCommand));
        Assert.True(dismiss.IsEffectivelyVisible);
        Point point = dismiss.TranslatePoint(new Point(dismiss.Bounds.Width / 2, dismiss.Bounds.Height / 2), window)!.Value;
        Assert.InRange(point.X, 0, window.Width); Assert.InRange(point.Y, 0, window.Height);
        window.MouseDown(point, MouseButton.Left); window.MouseUp(point, MouseButton.Left);
    }

    private static void CreateMixedScan(string path) {
        byte[] image = PdfDocument.Create(compose => compose.Page(page => page.Size(300, 400)
            .Content(content => content.Text("SCANNED DOCUMENT"))))
            .Render.Pages(PdfPageSelection.From(1), new PdfPageRenderOptions { Format = PdfPageRenderFormat.Png, Dpi = 72 }).Single().Bytes!;
        PdfDocument.Create(compose => {
            compose.Page(page => page.Size(300, 400).Content(content => content.Text("This page has selectable text.")));
            compose.Page(page => page.Size(300, 400).Content(content => content.Image(image, 150, 200)));
        }).Save(path);
    }
    private static async Task Render(Window window, params PdfPageViewModel[] pages) {
        window.Measure(new Size(window.Width, window.Height)); window.Arrange(new Rect(0, 0, window.Width, window.Height));
        for (int pass = 0; pass < 2; pass++) {
            // Let newly assigned pane bindings attach their page controls before requesting the scene.
            await Task.Yield();
            window.UpdateLayout();
            foreach (var page in pages) await page.EnsureRenderedAsync();
        }
        for (int attempt = 0; attempt < 200 && pages.Any(page => page.IsRendering); attempt++) await Task.Delay(10);
        Assert.DoesNotContain(pages, page => page.IsRendering);
        window.UpdateLayout();
    }
    private static void Capture(Window window, string name) {
        using var frame = window.CaptureRenderedFrame(); Assert.NotNull(frame);
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(output)) return;
        Directory.CreateDirectory(output); frame.Save(Path.Combine(output, name + ".png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
