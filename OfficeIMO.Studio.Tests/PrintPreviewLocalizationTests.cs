using System.Globalization;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.VisualTree;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Localization;
using OfficeIMO.Studio.Infrastructure.Preferences;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Tests;

public sealed class PrintPreviewLocalizationTests {
    [Theory]
    [InlineData("pl", 960, 640, false)]
    [InlineData("de", 960, 640, false)]
    [InlineData("fr", 1280, 850, true)]
    public async Task ChoicesAndPreparedPreviewUseCultureResources(string culture, int width, int height, bool dark) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            IStudioLocalizer previous = StudioLocalization.Current;
            var previousTheme = Application.Current!.RequestedThemeVariant;
            CultureInfo previousCulture = CultureInfo.CurrentCulture, previousUi = CultureInfo.CurrentUICulture;
            CultureInfo? previousDefault = CultureInfo.DefaultThreadCurrentCulture, previousDefaultUi = CultureInfo.DefaultThreadCurrentUICulture;
            var paths = new StudioDataPaths(Path.Combine(Path.GetTempPath(), "officeimo-print-localized-" + Guid.NewGuid().ToString("N")));
            new JsonStudioPreferencesStore(paths.PreferencesPath).Save(new StudioPreferences {
                UiCulture = culture, Theme = dark ? StudioThemePreference.Dark : StudioThemePreference.Light
            });
            var services = StudioApplicationServices.Create(paths);
            StudioLocalization.Configure(services.Localizer);
            Application.Current.RequestedThemeVariant = dark ? Avalonia.Styling.ThemeVariant.Dark : Avalonia.Styling.ThemeVariant.Light;
            var window = new MainWindow(services) { Width = width, Height = height };
            try {
                IStudioLocalizer localizer = services.Localizer;
                string source = Path.Combine(paths.Root, "localized-print.pdf");
                PdfDocument.Create(document => document.Page(page => page.Size(200, 300).Margin(12)
                    .Background(PdfColor.LightGray).Content(content => content.Item(item =>
                        item.Paragraph(text => text.Text("Print pixel review")))))).Save(source);
                byte[] original = File.ReadAllBytes(source);
                window.Show();
                await window.TabHost.OpenDocumentAsync(source);
                window.ViewModel.ShowPrintPreviewCommand.Execute(null);
                var print = window.ViewModel.OutputWorkbench.PrintPreview;
                Assert.Equal(localizer.Get("PrintPreview.Status.Ready"), print.Status);
                Assert.Equal(localizer.Get("PrintPreview.Summary.Empty"), print.Summary);
                foreach (var choice in print.OrientationChoices)
                    Assert.Equal(localizer.Get("PrintPreview.Orientation." + choice.Value), choice.Label);
                foreach (var choice in print.ScaleChoices) {
                    Assert.Equal(localizer.Get("PrintPreview.Scale." + choice.Value + ".Label"), choice.Label);
                    Assert.Equal(localizer.Get("PrintPreview.Scale." + choice.Value + ".Description"), choice.Description);
                }
                string[] pageKeys = ["One", "Two", "Four", "Six", "Nine"];
                Assert.Equal(pageKeys.Select(key => localizer.Get("PrintPreview.PagesPerSheet." + key)),
                    print.PagesPerSheetChoices.Select(choice => choice.Label));
                window.UpdateLayout();
                ComboBox Choice(object choices) => window.GetVisualDescendants().OfType<ComboBox>()
                    .Single(control => ReferenceEquals(control.ItemsSource, choices));
                Assert.Contains(Choice(print.OrientationChoices).GetVisualDescendants().OfType<TextBlock>(),
                    text => text.Text == print.SelectedOrientation.Label);
                Assert.Contains(Choice(print.ScaleChoices).GetVisualDescendants().OfType<TextBlock>(),
                    text => text.Text == print.SelectedScale.Label);
                Choice(print.ScaleChoices).BringIntoView();
                window.UpdateLayout();
                Capture(window, culture + "-print-choices-" + width);

                Choice(print.ColorChoices).SelectedItem = print.ColorChoices.Single(choice => choice.Value == PdfPrintColorMode.Grayscale);
                await print.BuildPreviewCommand.ExecuteAsync(null);
                Assert.True(print.HasPreview, print.Status);
                Assert.Equal(localizer.Get("PrintPreview.Status.Completed"), print.Status);
                Assert.Equal(localizer.Format("PrintPreview.Summary", 1, 1), print.Summary);
                Assert.Equal(PdfPrintOrientation.Automatic, print.SelectedOrientation.Value);
                Assert.Equal(PdfPrintScaleMode.Fit, print.SelectedScale.Value);
                Assert.Single(print.Sheets);
                Assert.Single(print.Sheets[0].Placements);
                Assert.True(print.HasRenderingNotices);
                PrintRenderingNotice font = Assert.Single(print.RenderingNotices,
                    notice => notice.Code == "render.resource.font-substitution");
                Assert.Equal(localizer.Get("PrintPreview.Rendering.FontSubstitution"), font.Message);
                Assert.StartsWith(font.Code + ":", font.Details);
                window.UpdateLayout();
                Assert.Contains(window.GetVisualDescendants().OfType<TextBlock>(), text => text.IsEffectivelyVisible && text.Text == font.Message);
                Capture(window, culture + "-print-ready-" + width);
                Assert.Equal(original, File.ReadAllBytes(source));
                print.SelectedScale = print.ScaleChoices.Single(choice => choice.Value == PdfPrintScaleMode.Custom);
                Assert.False(print.HasPreview);
                Assert.False(print.HasRenderingNotices);
                Assert.Empty(print.RenderingNotices);
            } finally {
                window.Close();
                Application.Current!.RequestedThemeVariant = previousTheme;
                StudioLocalization.Configure(previous);
                CultureInfo.CurrentCulture = previousCulture; CultureInfo.CurrentUICulture = previousUi;
                CultureInfo.DefaultThreadCurrentCulture = previousDefault; CultureInfo.DefaultThreadCurrentUICulture = previousDefaultUi;
                Directory.Delete(paths.Root, recursive: true);
            }
            return true;
        }, CancellationToken.None);
    }

    private static void Capture(Window window, string name) {
        using var image = window.CaptureRenderedFrame();
        Assert.NotNull(image);
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(output)) return;
        Directory.CreateDirectory(output);
        image.Save(Path.Combine(output, name + ".png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
