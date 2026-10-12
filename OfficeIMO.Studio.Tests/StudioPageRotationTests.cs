using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.VisualTree;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioPageRotationTests {
    [Theory]
    [InlineData(960)]
    [InlineData(1280)]
    public async Task Tagged_page_rotation_is_available_and_updates_the_rendered_workspace(int width) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var window = new MainWindow();
            try {
                window.Show();
                await window.TabHost.OpenDocumentAsync(Path.Combine(AppContext.BaseDirectory, "Fixtures", "openpreserve-pdfa1b-text.pdf"));
                var model = Assert.Single(window.TabHost.Tabs).Document;
                model.ActualSizeCommand.Execute(null);
                double savedWidth = model.Pages[0].DisplayWidth;
                double savedHeight = model.Pages[0].DisplayHeight;
                await model.Commands["RotateViewRight"].ExecuteAsync();
                Assert.Equal(savedHeight, model.Pages[0].DisplayWidth);
                Assert.Equal(savedWidth, model.Pages[0].DisplayHeight);
                Assert.False(model.IsDirty);
                Assert.False(model.CanUndo);
                model.FitPageCommand.Execute(null);
                Layout(window, width);
                await model.Pages[0].EnsureRenderedAsync();
                await TextEditingReviewTests.WaitUntilAsync(() => window.GetVisualDescendants().OfType<OfficeIMO.Studio.Features.Reader.PdfPageCanvas>()
                    .Any(canvas => canvas.IsEffectivelyVisible && canvas.Scene is not null));
                Layout(window, width);
                Capture(window, width, "view-rotated");
                await model.Commands["ResetViewRotation"].ExecuteAsync();
                Assert.Equal(0, model.ViewRotationDegrees);
                model.ShowPagesModeCommand.Execute(null);
                Assert.False(model.CanMutatePages);
                Assert.True(model.CanRotatePages);
                Assert.False(model.Commands["RotatePagesRight"].IsAvailable);
                model.SetOrganizerSelection([model.OrganizerPages[0]]);
                Assert.True(model.Commands["RotatePagesRight"].IsAvailable);
                Assert.False(model.Commands["DuplicatePages"].IsAvailable);
                Layout(window, width);
                Button rotate = Assert.Single(window.GetVisualDescendants().OfType<Button>(),
                    button => ReferenceEquals(button.Command, model.Commands["RotatePagesRight"]));
                Assert.True(rotate.IsEffectivelyEnabled);
                await TextEditingReviewTests.WaitUntilAsync(() => model.OrganizerPages[0].Scene is not null);
                Capture(window, width, "before");

                await model.Commands["RotatePagesRight"].ExecuteAsync();
                Assert.False(model.HasError, model.ErrorMessage);
                Assert.True(model.IsDirty);
                Assert.True(model.Pages[0].DisplayWidth > model.Pages[0].DisplayHeight);
                Layout(window, width);
                await TextEditingReviewTests.WaitUntilAsync(() => model.OrganizerPages[0].Scene is not null);
                Assert.True(model.OrganizerPages[0].ThumbnailWidth > model.OrganizerPages[0].ThumbnailHeight);
                Capture(window, width, "rotated");
                await model.Commands["RotatePagesRight"].ExecuteAsync();
                Assert.True(model.Pages[0].DisplayWidth < model.Pages[0].DisplayHeight);
                await model.Commands["Undo"].ExecuteAsync();
                Assert.True(model.Pages[0].DisplayWidth > model.Pages[0].DisplayHeight);
                await model.Commands["Undo"].ExecuteAsync();
                Assert.True(model.Pages[0].DisplayWidth < model.Pages[0].DisplayHeight);
                await model.Commands["Redo"].ExecuteAsync();
                Assert.True(model.Pages[0].DisplayWidth > model.Pages[0].DisplayHeight);
                Assert.False(model.Commands["DuplicatePages"].IsAvailable);
                await model.Commands["Undo"].ExecuteAsync();
                Assert.False(model.IsDirty);
                model.SetOrganizerSelection([model.OrganizerPages[0]]);
                model.ActualSizeCommand.Execute(null);
                Assert.True(model.CanCropSelection);
                Layout(window, width);
                var workspaceView = Assert.Single(window.GetVisualDescendants().OfType<OfficeIMO.Studio.Features.Workspace.DocumentWorkspaceView>());
                var inspectorToggle = workspaceView.FindControl<Avalonia.Controls.Primitives.ToggleButton>("InspectorToggle")!;
                inspectorToggle.IsChecked = true;
                inspectorToggle.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));
                Layout(window, width);
                Button crop = Assert.Single(window.GetVisualDescendants().OfType<Button>(),
                    button => ReferenceEquals(button.Command, model.CropSelectedCommand));
                Button resetCrop = Assert.Single(window.GetVisualDescendants().OfType<Button>(),
                    button => ReferenceEquals(button.Command, model.ResetCropSelectedCommand));
                Assert.True(crop.IsEffectivelyEnabled);
                Assert.True(resetCrop.IsEffectivelyEnabled);
                double originalWidth = model.Pages[0].DisplayWidth;
                model.CropMargin = 24;
                await model.CropSelectedCommand.ExecuteAsync(null);
                Assert.False(model.HasError, model.ErrorMessage);
                Assert.Equal(originalWidth - 48, model.Pages[0].DisplayWidth, 2);
                Layout(window, width);
                await TextEditingReviewTests.WaitUntilAsync(() => model.OrganizerPages[0].Scene is not null);
                Capture(window, width, "cropped");
                await model.ResetCropSelectedCommand.ExecuteAsync(null);
                Assert.False(model.HasError, model.ErrorMessage);
                Assert.Equal(originalWidth, model.Pages[0].DisplayWidth);
                Layout(window, width);
                await TextEditingReviewTests.WaitUntilAsync(() => model.OrganizerPages[0].Scene is not null);
                Capture(window, width, "crop-reset");
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    private static void Layout(MainWindow window, int width) {
        window.Width = width;
        window.Height = 620;
        window.ApplyResponsiveLayout(width);
        window.Measure(new Size(width, 620));
        window.Arrange(new Rect(0, 0, width, 620));
        window.UpdateLayout();
    }

    private static void Capture(Window window, int width, string state) {
        using var frame = window.CaptureRenderedFrame();
        Assert.NotNull(frame);
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(output)) return;
        Directory.CreateDirectory(output);
        frame.Save(Path.Combine(output, $"tagged-rotation-{width}-{state}.png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
