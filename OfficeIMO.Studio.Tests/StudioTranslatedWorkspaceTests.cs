using System.Globalization;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.VisualTree;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Settings;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Localization;
using OfficeIMO.Studio.Infrastructure.Preferences;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioTranslatedWorkspaceTests {
    [Theory]
    [InlineData("pl", 960, 620, false)]
    [InlineData("de", 960, 620, false)]
    [InlineData("fr", 1280, 800, true)]
    public async Task TranslatedSettingsAndPrintControlsRemainReachable(string culture, int width, int height, bool dark) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            IStudioLocalizer previous = StudioLocalization.Current;
            CultureInfo previousCulture = CultureInfo.CurrentCulture, previousUiCulture = CultureInfo.CurrentUICulture;
            CultureInfo? previousDefault = CultureInfo.DefaultThreadCurrentCulture, previousDefaultUi = CultureInfo.DefaultThreadCurrentUICulture;
            var previousTheme = Application.Current!.RequestedThemeVariant;
            var paths = new StudioDataPaths(Path.Combine(Path.GetTempPath(), "officeimo-translated-" + Guid.NewGuid().ToString("N")));
            new JsonStudioPreferencesStore(paths.PreferencesPath).Save(new StudioPreferences {
                UiCulture = culture, Theme = dark ? StudioThemePreference.Dark : StudioThemePreference.Light
            });
            var services = StudioApplicationServices.Create(paths);
            StudioLocalization.Configure(services.Localizer);
            Application.Current.RequestedThemeVariant = dark ? Avalonia.Styling.ThemeVariant.Dark : Avalonia.Styling.ThemeVariant.Light;
            var window = new MainWindow(services) { Width = width, Height = height };
            try {
                Assert.Equal(culture, services.Localizer.Culture.Name);
                window.Show();
                window.ViewModel.ShowSettingsCommand.Execute(null);
                window.UpdateLayout();
                var settings = window.GetVisualDescendants().OfType<SettingsView>().Single(view => view.IsEffectivelyVisible);
                var setup = window.ViewModel.Settings.OcrSetup;
                Button check = settings.GetVisualDescendants().OfType<Button>().Single(button => ReferenceEquals(button.Command, setup.CheckCommand));
                check.BringIntoView(); window.UpdateLayout();
                AssertReachable(check, window, width, height);
                settings.GetVisualDescendants().OfType<ScrollViewer>().First().Offset = new Vector(0, 370);
                window.UpdateLayout();
                Capture(window, culture + "-settings-" + width);
                string source = Path.Combine(paths.Root, "translation-layout.pdf");
                PdfDocument.Create(document => document.Page(page => page.Size(200, 300).Content(content =>
                    content.Item(item => item.Paragraph(text => text.Text("Translated interface print layout")))))).Save(source);
                await window.TabHost.OpenDocumentAsync(source);
                window.ViewModel.ShowPrintPreviewCommand.Execute(null);
                var print = window.ViewModel.OutputWorkbench.PrintPreview;
                window.UpdateLayout();
                ComboBox choice = window.GetVisualDescendants().OfType<ComboBox>().Single(control => ReferenceEquals(control.ItemsSource, print.PagesPerSheetChoices));
                choice.SelectedItem = print.PagesPerSheetChoices.Single(item => item.Value == 9);
                ComboBox color = window.GetVisualDescendants().OfType<ComboBox>().Single(control => ReferenceEquals(control.ItemsSource, print.ColorChoices));
                color.SelectedItem = print.ColorChoices.Single(item => item.Value == PdfPrintColorMode.Grayscale);
                Button prepare = window.GetVisualDescendants().OfType<Button>().Single(button => ReferenceEquals(button.Command, print.BuildPreviewCommand));
                prepare.BringIntoView(); window.UpdateLayout();
                AssertReachable(prepare, window, width, height);
                await print.BuildPreviewCommand.ExecuteAsync(null);
                Assert.True(print.HasPreview, print.Status);
                Assert.Equal(9, print.SelectedPagesPerSheet!.Value);
                Assert.Equal(PdfPrintColorMode.Grayscale, print.SelectedColor!.Value);
                Capture(window, culture + "-print-" + width);
            } finally {
                window.Close();
                StudioLocalization.Configure(previous);
                Application.Current!.RequestedThemeVariant = previousTheme;
                CultureInfo.CurrentCulture = previousCulture; CultureInfo.CurrentUICulture = previousUiCulture;
                CultureInfo.DefaultThreadCurrentCulture = previousDefault; CultureInfo.DefaultThreadCurrentUICulture = previousDefaultUi;
                Directory.Delete(paths.Root, recursive: true);
            }
            return true;
        }, CancellationToken.None);
    }

    private static void AssertReachable(Control control, Window window, int width, int height) {
        Assert.True(control.IsEffectivelyVisible); Assert.True(control.IsEffectivelyEnabled);
        Point point = control.TranslatePoint(new Point(), window)!.Value;
        Assert.InRange(point.X, 0, width - control.Bounds.Width);
        Assert.InRange(point.Y, 0, height - control.Bounds.Height);
    }

    private static void Capture(Window window, string name) {
        using var image = window.CaptureRenderedFrame(); Assert.NotNull(image);
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(output)) return;
        Directory.CreateDirectory(output);
        image.Save(Path.Combine(output, name + ".png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
