using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.VisualTree;
using Avalonia.Threading;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioSourceConflictVisualTests {
    [Theory]
    [InlineData(960, 620, false)]
    [InlineData(1280, 820, true)]
    public async Task SaveConflictExplainsTheNextActionAndRetainsTheEditedDocument(int width, int height, bool dark) {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            services.Preferences.Update(current => current with { Theme = dark ? StudioThemePreference.Dark : StudioThemePreference.Light });
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "Report edited in another application.pdf");
            PdfDocument.Create(document => document.Page(page => page.Size(200, 300))).Save(source);
            var window = new MainWindow(services) { Width = width, Height = height };
            try {
                window.Show();
                await window.TabHost.OpenDocumentAsync(source);
                var model = window.ViewModel;
                model.SetOrganizerSelection([model.OrganizerPages[0]]);
                await model.DuplicateSelectedCommand.ExecuteAsync(null);
                byte[] external = PdfDocument.Create(document => document.Page(page => page.Size(300, 400))).ToBytes();
                File.WriteAllBytes(source, external);
                await model.SaveCommand.ExecuteAsync(null);
                await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
                Assert.Contains("changed since", model.ErrorMessage);
                Assert.Equal("Operation failed", model.OperationStatus);
                Assert.True(model.IsDirty);
                Assert.Equal(2, model.Pages.Count);
                Assert.Equal(external, File.ReadAllBytes(source));
                window.UpdateLayout();
                var message = window.GetVisualDescendants().OfType<TextBlock>().First(text => text.Text == model.ErrorMessage && text.IsEffectivelyVisible);
                Point point = message.TranslatePoint(default, window)!.Value;
                Assert.InRange(point.X, 0, window.Bounds.Width - message.Bounds.Width);
                Assert.InRange(point.Y, 0, window.Bounds.Height - message.Bounds.Height);
                using var frame = window.CaptureRenderedFrame();
                Assert.NotNull(frame);
                string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
                if (!string.IsNullOrWhiteSpace(output)) {
                    Directory.CreateDirectory(output);
                    frame.Save(Path.Combine(output, $"save-conflict-{width}-{(dark ? "dark" : "light")}.png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
                }
                await model.UndoCommand.ExecuteAsync(null);
            } finally {
                foreach (Window dialog in window.OwnedWindows.ToArray()) dialog.Close(false);
                window.Close();
                window.TabHost.Dispose();
            }
            return true;
        }, CancellationToken.None);
    }
}
