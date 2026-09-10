using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.VisualTree;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Tests;

public sealed class RedactionReviewVisualTests {
    [Theory]
    [InlineData(960, 620, false)]
    [InlineData(1280, 900, true)]
    public async Task ReviewControlsApplyOnlyIncludedMarks(int width, int height, bool dark) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-redaction-visual-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string source = Path.Combine(root, "source.pdf");
            PdfDocument.Create(compose => {
                compose.Page(page => page.Content(content => content.Item(item => item.Paragraph(text => text.Text("Private account 123")))));
                compose.Page(page => page.Content(content => content.Item(item => item.Paragraph(text => text.Text("Private account 456")))));
            }).Save(source);
            using var app = TestAppBuilder.StartSession();
            await app.Dispatch(async () => {
                var services = ((App)Application.Current!).Services;
                services.Preferences.Update(current => current with { Theme = dark ? StudioThemePreference.Dark : StudioThemePreference.Light });
                var window = new MainWindow(services) { Width = width, Height = height };
                try {
                    window.Show();
                    await window.TabHost.OpenDocumentAsync(source);
                    var model = window.ViewModel;
                    model.ShowProtectModeCommand.Execute(null);
                    model.RedactionSearchText = "Private account";
                    await model.SearchRedactionsCommand.ExecuteAsync(null);
                    window.UpdateLayout();
                    var inspector = Assert.Single(window.GetVisualDescendants().OfType<RedactionInspectorView>());
                    var marks = Assert.Single(inspector.GetVisualDescendants().OfType<ListBox>());
                    Assert.Equal(2, marks.ItemCount);
                    var included = inspector.GetVisualDescendants().OfType<CheckBox>()
                        .Where(check => check.DataContext is PdfRedactionMarkViewModel).ToArray();
                    Assert.Equal(2, included.Length);
                    included[1].IsChecked = false;
                    Assert.False(model.RedactionMarks[1].IsIncluded);
                    var review = inspector.GetVisualDescendants().OfType<Button>()
                        .Single(button => ReferenceEquals(button.Command, model.ReviewRedactionsCommand));
                    Click(window, review);
                    if (model.ReviewRedactionsCommand.ExecutionTask is { } reviewTask) await reviewTask;
                    var apply = inspector.GetVisualDescendants().OfType<Button>()
                        .Single(button => ReferenceEquals(button.Command, model.ApplyPendingRedactionCommand));
                    Assert.True(apply.IsEnabled, model.ErrorMessage ?? model.PendingRedactionSummary);
                    apply.BringIntoView();
                    window.UpdateLayout();
                    Point position = apply.TranslatePoint(default, window)!.Value;
                    Assert.InRange(position.X, 0, window.Bounds.Width - apply.Bounds.Width + 1);
                    Assert.InRange(position.Y, 0, window.Bounds.Height - apply.Bounds.Height + 1);
                    Capture(window, $"redaction-reviewed-{width}-{dark}.png");
                    Click(window, apply);
                    if (model.ApplyPendingRedactionCommand.ExecutionTask is { } applyTask) await applyTask;
                    Assert.False(model.HasError, model.ErrorMessage);
                    Assert.True(model.HasRedactionEvidence);
                    Assert.Equal(1, model.LastRedactionSummary!.AreaCount);
                    Assert.Empty(model.RedactionMarks);
                    window.UpdateLayout();
                    Capture(window, $"redaction-applied-{width}-{dark}.png");
                } finally {
                    foreach (var tab in window.TabHost.Tabs.ToArray()) tab.Document.CompletePreparedClose();
                    window.Close();
                    window.TabHost.Dispose();
                }
                return true;
            }, CancellationToken.None);
        } finally { Directory.Delete(root, recursive: true); }
    }

    private static void Click(Window window, Button button) {
        button.BringIntoView();
        window.UpdateLayout();
        Point point = button.TranslatePoint(new Point(button.Bounds.Width / 2, button.Bounds.Height / 2), window)!.Value;
        window.MouseDown(point, MouseButton.Left);
        window.MouseUp(point, MouseButton.Left);
    }

    private static void Capture(Window window, string name) {
        using var frame = window.CaptureRenderedFrame();
        Assert.NotNull(frame);
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(output)) return;
        Directory.CreateDirectory(output);
        frame.Save(Path.Combine(output, name), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
