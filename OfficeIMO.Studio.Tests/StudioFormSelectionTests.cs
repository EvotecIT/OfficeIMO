using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.VisualTree;
using OfficeIMO.Pdf;
using PdfPageCanvas = OfficeIMO.Studio.Features.Reader.PdfPageCanvas;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioFormSelectionTests {
    [Theory]
    [InlineData(960, 620, false)]
    [InlineData(1280, 800, true)]
    public async Task SelectingAFieldRevealsItsRotatedPageAndKeyboardSelectionReturnsToTheInspector(int width, int height, bool dark) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            services.Preferences.Update(current => current with { Theme = dark ? StudioThemePreference.Dark : StudioThemePreference.Light });
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "selection.pdf");
            byte[] bytes = PdfDocument.Create(compose => {
                compose.Page(page => page.Size(300, 400));
                compose.Page(page => page.Size(300, 400));
            }).Pages.Rotate(90, 2).Forms.Edit(edit => edit
                .Create(new() { Name = "First", PageNumber = 1, X = 20, Y = 50, Value = "First page" })
                .Create(new() { Name = "Rotated", PageNumber = 2, X = 20, Y = 50, Value = "Second page" })
                .Create(new() { Name = "Next", PageNumber = 2, X = 20, Y = 150, Value = "Next field" }))
                .ToBytes();
            File.WriteAllBytes(source, bytes);
            var window = new MainWindow(services) { Width = width, Height = height };
            try {
                window.Show();
                var model = window.ViewModel;
                await model.OpenDocumentAsync(source);
                model.DocumentMode = StudioDocumentMode.Forms;
                model.SelectedFormField = model.FormFields.Single(field => field.Name == "Rotated");
                Assert.Equal(2, model.SelectedPage!.PageNumber);
                await Avalonia.Threading.Dispatcher.UIThread.InvokeAsync(() => window.UpdateLayout(), Avalonia.Threading.DispatcherPriority.Background);
                await model.SelectedPage.EnsureRenderedAsync();
                await Avalonia.Threading.Dispatcher.UIThread.InvokeAsync(() => window.UpdateLayout(), Avalonia.Threading.DispatcherPriority.Background);
                var canvas = window.GetVisualDescendants().OfType<PdfPageCanvas>().Single(canvas => canvas.Scene?.PageNumber == 2 && canvas.DataContext is OfficeIMO.Studio.Features.Reader.PdfPageViewModel);
                Assert.Equal("Rotated", canvas.FormAnchorFieldName);
                Assert.Single(canvas.FormAnchorBounds);
                Assert.All(model.Pages.Where(page => page.PageNumber != 2), page => Assert.Null(page.FormAnchorFieldName));
                canvas.Focus();
                canvas.RaiseEvent(new KeyEventArgs { RoutedEvent = InputElement.KeyDownEvent, Key = Key.End });
                Assert.Equal("Next", model.SelectedFormField!.Name);
                Assert.Equal("Next", canvas.FormAnchorFieldName);
                string? folder = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
                if (!string.IsNullOrEmpty(folder)) {
                    Directory.CreateDirectory(folder);
                    window.UpdateLayout();
                    using var image = window.CaptureRenderedFrame();
                    Assert.NotNull(image);
                    image.Save(Path.Combine(folder, $"forms-selection-{width}-{dark}.png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
                }
                Assert.False(model.IsDirty);
                Assert.Equal(bytes, File.ReadAllBytes(source));
                model.DocumentMode = StudioDocumentMode.View;
                Assert.Null(canvas.FormAnchorFieldName);
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }
}
