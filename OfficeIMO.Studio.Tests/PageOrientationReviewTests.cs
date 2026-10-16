using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.VisualTree;
using OfficeIMO.Ocr;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workspace;

namespace OfficeIMO.Studio.Tests;

public sealed class PageOrientationReviewTests {
    [Theory]
    [InlineData(960)]
    [InlineData(1280)]
    public async Task SuggestionsRequirePreviewAndExplicitAcceptanceAndUndoTogether(int width) {
        using var files = new TextEditingReviewTests.Files();
        PdfDocument.Create(compose => {
            for (int page = 0; page < 2; page++) compose.Page(value => value.Size(300, 500).Content(content => content.Text("Orientation review")));
        }).Save(files.Source);
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var engine = new DelegateOcrEngine("orientation-ui", (_, _) => Task.FromResult(new OcrResult {
                Orientation = new() { ClockwiseRotationDegrees = 90, Confidence = .95 }
            }), new OcrEngineCapabilities { SupportsOrientationDetection = true });
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null),
                orientationEngineFactory: _ => Task.FromResult<IOcrEngine>(engine));
            var view = new DocumentWorkspaceView { DataContext = model };
            var window = new Window { Content = view, Width = width, Height = 720 };
            try {
                window.Show(); await model.OpenDocumentAsync(files.Source);
                model.ShowPagesModeCommand.Execute(null); model.SetOrganizerSelection(model.OrganizerPages.ToArray());
                view.ApplyResponsiveLayout(width, 720);
                var inspector = view.FindControl<Avalonia.Controls.Primitives.ToggleButton>("InspectorToggle")!;
                inspector.IsChecked = true; inspector.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));
                window.UpdateLayout();
                await model.Commands["DetectPageOrientation"].ExecuteAsync();
                Assert.False(model.HasError, model.ErrorMessage); Assert.True(model.HasOrientationReview);
                Assert.All(model.OrientationChoices, choice => { Assert.False(choice.IsAccepted); Assert.False(choice.CanAccept); });
                Assert.False(model.CanApplyOrientationReview); Assert.False(model.IsDirty);
                await model.PreviewPageOrientationCommand.ExecuteAsync(null);
                Assert.True(model.HasOrientationPreview, model.ErrorMessage); Assert.True(model.SelectedOrientationChoice!.CanAccept);
                model.SelectedOrientationChoice.IsAccepted = true;
                model.SelectedOrientationChoice = model.OrientationChoices[1];
                Assert.False(model.HasOrientationPreview);
                await model.PreviewPageOrientationCommand.ExecuteAsync(null);
                model.SelectedOrientationChoice.IsAccepted = true;
                window.UpdateLayout();
                var previewButton = Assert.Single(view.GetVisualDescendants().OfType<Button>(),
                    button => ReferenceEquals(button.Command, model.PreviewPageOrientationCommand));
                Assert.True(previewButton.IsEffectivelyVisible);
                previewButton.BringIntoView(); window.UpdateLayout();
                Capture(window, width);
                var applyButton = Assert.Single(view.GetVisualDescendants().OfType<Button>(),
                    button => ReferenceEquals(button.Command, model.ApplyPageOrientationReviewCommand));
                applyButton.BringIntoView(); window.UpdateLayout();
                Capture(window, width, "accepted");
                await model.ApplyPageOrientationReviewCommand.ExecuteAsync(null);
                Assert.False(model.HasError, model.ErrorMessage); Assert.False(model.HasOrientationReview);
                Assert.True(model.IsDirty); Assert.All(model.Pages, page => Assert.True(page.DisplayWidth > page.DisplayHeight));
                await model.UndoCommand.ExecuteAsync(null);
                Assert.False(model.IsDirty); Assert.All(model.Pages, page => Assert.True(page.DisplayWidth < page.DisplayHeight));
                await model.RedoCommand.ExecuteAsync(null);
                Assert.True(model.IsDirty); Assert.All(model.Pages, page => Assert.True(page.DisplayWidth > page.DisplayHeight));
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    private static void Capture(Window window, int width, string state = "preview") {
        using var frame = window.CaptureRenderedFrame(); Assert.NotNull(frame);
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(output)) return;
        Directory.CreateDirectory(output); frame.Save(Path.Combine(output, $"orientation-review-{width}-{state}.png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
