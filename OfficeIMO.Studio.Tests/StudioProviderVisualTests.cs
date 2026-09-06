using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Avalonia.VisualTree;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioProviderVisualTests {
    [Theory]
    [InlineData(960, 620, false)]
    [InlineData(1280, 820, true)]
    public async Task ProviderSaveExplainsGuaranteesAndFailedWriteRetainsVisibleEdits(int width, int height, bool dark) {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            services.Preferences.Update(current => current with { Theme = dark ? StudioThemePreference.Dark : StudioThemePreference.Light });
            var provider = new TestStorageFile("content://documents/opaque-source-id", StudioProviderDocumentTests.CreatePdf(),
                "Quarterly review with comments and proposed changes from the document provider.pdf") { FailWrite = true };
            string location = await services.Storage.RegisterAsync(provider.Item, default);
            var window = new MainWindow(services) { Width = width, Height = height };
            try {
                window.Show();
                await window.TabHost.OpenDocumentAsync(location);
                var model = window.ViewModel;
                Assert.True(model.HasDocument, model.ErrorMessage);
                Assert.Equal(provider.Name, model.DocumentName);
                model.SetOrganizerSelection([model.OrganizerPages[0]]);
                await model.DuplicateSelectedCommand.ExecuteAsync(null);
                Task save = model.SaveCommand.ExecuteAsync(null);
                var dialog = Assert.Single(window.OwnedWindows.OfType<ProviderSaveDialog>());
                dialog.UpdateLayout();
                AssertContained(dialog);
                Capture(dialog, $"provider-save-warning-{width}-{(dark ? "dark" : "light")}.png");
                var saveButton = dialog.GetVisualDescendants().OfType<Button>()
                    .Single(button => Equals(button.Content, services.Localizer.Get("Common.Save")));
                saveButton.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                await save;
                await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
                Assert.True(model.HasError);
                Assert.True(model.IsDirty);
                Assert.True(window.GetVisualDescendants().OfType<Button>()
                    .Single(button => ReferenceEquals(button.Command, model.Commands["Save"])).IsEnabled);
                Assert.Equal(2, model.Pages.Count);
                Assert.Equal(1, provider.Writes);
                Assert.Equal(provider.Reads, provider.ClosedReads);
                window.UpdateLayout();
                var error = window.GetVisualDescendants().OfType<TextBlock>()
                    .First(text => text.IsEffectivelyVisible && text.Text == model.ErrorMessage);
                AssertContained(window, error);
                Capture(window, $"provider-save-failure-{width}-{(dark ? "dark" : "light")}.png");
            } finally {
                foreach (Window dialog in window.OwnedWindows.ToArray()) dialog.Close(false);
                foreach (var tab in window.TabHost.Tabs.ToArray()) tab.Document.CompletePreparedClose();
                window.Close();
                window.TabHost.Dispose();
            }
            return true;
        }, CancellationToken.None);
    }

    private static void AssertContained(Window window) {
        foreach (Control control in window.GetVisualDescendants().OfType<Control>()
                     .Where(control => control.IsEffectivelyVisible && control is TextBlock or Button)) AssertContained(window, control);
    }

    private static void AssertContained(Window window, Control control) {
        Point point = control.TranslatePoint(default, window)!.Value;
        Assert.InRange(point.X, 0, window.Bounds.Width - control.Bounds.Width + 1);
        Assert.InRange(point.Y, 0, window.Bounds.Height - control.Bounds.Height + 1);
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
