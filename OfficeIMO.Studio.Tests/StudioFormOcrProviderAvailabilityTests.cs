using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Threading;
using Avalonia.VisualTree;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Tests {
    public sealed class StudioFormOcrProviderAvailabilityTests {
        [Theory]
        [InlineData(360, 760)]
        [InlineData(680, 920)]
        public async Task FormRecognitionIsOfferedForExistingFields(int width, int height) {
            using var session = TestAppBuilder.StartSession();
            await session.Dispatch(async () => {
                var services = ((App)Application.Current!).Services;
                Directory.CreateDirectory(services.Paths.Root);
                string source = Path.Combine(services.Paths.Root, "provider-availability-form.pdf");
                File.Copy(Path.Combine(AppContext.BaseDirectory, "Fixtures", "PdfFormOcr", "reportlab-scanned-form.pdf"), source);
                byte[] original = File.ReadAllBytes(source);
                using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services);
                await model.OpenDocumentAsync(source);
                Assert.True(model.HasFormFields);
                Assert.False(model.HasFormDrafts);
                Assert.False(model.IsOpening);
                Assert.False(model.IsWorkspaceBusy);

                var view = new FormsInspectorView { DataContext = model };
                var window = new Window { Width = width, Height = height, Content = new ScrollViewer { Content = view } };
                try {
                    window.Show();
                    await Render(window);
                    view.GetVisualDescendants().OfType<Expander>().First().IsExpanded = true;
                    await Render(window);
                    var recognize = view.GetVisualDescendants().OfType<Button>()
                        .Single(button => ReferenceEquals(button.Command, model.RecognizeFormValuesCommand));
                    Assert.True(recognize.IsEffectivelyVisible);
                    Capture(window, width);

                    Assert.True(model.CanRecognizeFormValues);
                    Assert.True(recognize.IsEnabled);
                    Assert.False(model.IsDirty);
                    Assert.Equal(original, File.ReadAllBytes(source));
                } finally {
                    window.Close();
                }
                return true;
            }, CancellationToken.None);
        }

        private static async Task Render(Window window) {
            await Dispatcher.UIThread.InvokeAsync(window.UpdateLayout, DispatcherPriority.Background);
            AvaloniaHeadlessPlatform.ForceRenderTimerTick();
        }

        private static void Capture(Window window, int width) {
            using var frame = window.CaptureRenderedFrame();
            Assert.NotNull(frame);
            string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
            if (string.IsNullOrWhiteSpace(output)) return;
            Directory.CreateDirectory(output);
            frame.Save(Path.Combine(output, $"form-ocr-provider-{width}.png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
        }
    }
}
