using Avalonia;
using Avalonia.Headless;
using Avalonia.Styling;
using System.Globalization;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Localization;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Tests;

public sealed class ComparisonRangeReportTests {
    [Theory]
    [InlineData(960, 640, false)]
    [InlineData(1280, 800, true)]
    public async Task SelectedLongDocumentPairsNavigateOriginalPagesAndExportCurrentReport(int width, int height, bool dark) {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var appServices = ((App)Application.Current!).Services;
            string culture = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_COMPARISON_CULTURE") ?? "en";
            var paths = new StudioDataPaths(Path.Combine(appServices.Paths.Root, "comparison-localized"));
            new JsonStudioPreferencesStore(paths.PreferencesPath).Save(new StudioPreferences { UiCulture = culture });
            var services = StudioApplicationServices.Create(paths);
            IStudioLocalizer previousLocalizer = StudioLocalization.Current;
            CultureInfo previousCulture = CultureInfo.CurrentCulture, previousUi = CultureInfo.CurrentUICulture;
            CultureInfo? previousDefault = CultureInfo.DefaultThreadCurrentCulture, previousDefaultUi = CultureInfo.DefaultThreadCurrentUICulture;
            StudioLocalization.Configure(services.Localizer);
            ThemeVariant? previous = Application.Current.RequestedThemeVariant;
            Application.Current.RequestedThemeVariant = dark ? ThemeVariant.Dark : ThemeVariant.Light;
            string source = Path.Combine(services.Paths.Root, "large-current.pdf"), actual = Path.Combine(services.Paths.Root, "large-comparison.pdf");
            Create(source, false); Create(actual, true);
            string output = Path.Combine(services.Paths.Root, "report.html");
            var window = new MainWindow(services) { Width = width, Height = height };
            try {
                window.Show(); await window.TabHost.OpenDocumentAsync(source);
                var model = window.ViewModel;
                await model.OpenComparisonDocumentAsync(actual);
                model.ComparisonExpectedRange = "110,2"; model.ComparisonActualRange = "109,2,108";
                await model.ComparePagesCommand.ExecuteAsync(null);
                Assert.Equal(110, model.Pages.Count); Assert.Equal(110, model.ComparisonPages.Count);
                Assert.Equal(2, model.ComparisonDifferences.Count);
                Assert.Equal(110, model.SelectedPage!.PageNumber); Assert.Equal(109, model.ComparisonSelectedPage!.PageNumber);
                Assert.True(model.CanExportComparisonReport);
                model.SelectedPage = model.Pages[1];
                Assert.Equal(2, model.ComparisonSelectedPage!.PageNumber); Assert.Null(model.SelectedComparisonDifference);
                model.SelectedPage = model.Pages[109]; Assert.Equal(109, model.ComparisonSelectedPage!.PageNumber);
                model.NextComparisonDifferenceCommand.Execute(null);
                Assert.Equal(108, model.ComparisonSelectedPage!.PageNumber); Assert.Null(model.SelectedPage);
                model.PreviousComparisonDifferenceCommand.Execute(null);
                Assert.Equal(110, model.SelectedPage!.PageNumber); Assert.Equal(109, model.ComparisonSelectedPage!.PageNumber);
                window.UpdateLayout();
                for (int attempt = 0; attempt < 100 && (model.SelectedPage!.IsRendering || model.ComparisonSelectedPage!.IsRendering); attempt++)
                    await Task.Delay(20);
                window.UpdateLayout(); Capture(window, "comparison-range-" + culture + "-" + width + (dark ? "-dark" : "-light"));
                // An editable scope is a live input: changing it invalidates the existing gallery and navigation evidence.
                model.ComparisonActualRange = "109";
                Assert.False(model.CanExportComparisonReport); Assert.Empty(model.ComparisonDifferences);
            } finally {
                window.Close(); Application.Current.RequestedThemeVariant = previous;
                StudioLocalization.Configure(previousLocalizer);
                CultureInfo.CurrentCulture = previousCulture; CultureInfo.CurrentUICulture = previousUi;
                CultureInfo.DefaultThreadCurrentCulture = previousDefault; CultureInfo.DefaultThreadCurrentUICulture = previousDefaultUi;
            }
            using var exportModel = new MainWindowViewModel(_ => Task.FromResult<string?>(null),
                pickSaveComparisonReport: _ => Task.FromResult<string?>(output), services: TestAppBuilder.CreateTestServices());
            await exportModel.OpenDocumentAsync(source); await exportModel.OpenComparisonDocumentAsync(actual);
            exportModel.ComparisonExpectedRange = "110,2"; exportModel.ComparisonActualRange = "109,2,108";
            await exportModel.ComparePagesCommand.ExecuteAsync(null); await exportModel.ExportComparisonReportCommand.ExecuteAsync(null);
            Assert.True(File.Exists(output), exportModel.ErrorMessage);
            string html = File.ReadAllText(output); Assert.Contains("Expected 110 / actual 109", html);
            Assert.Contains("Unmatched actual page 108", html);
            string? artifactOutput = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
            if (!string.IsNullOrWhiteSpace(artifactOutput)) File.Copy(output, Path.Combine(artifactOutput, "comparison-report-" + width + ".html"), true);
            return true;
        }, default);
    }

    [Fact]
    public async Task ExportRejectsAChangedScopeOrSourceDuringDestinationSelection() {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            string source = Path.Combine(services.Paths.Root, "source.pdf"), actual = Path.Combine(services.Paths.Root, "actual.pdf");
            Create(source, false); Create(actual, true);
            var selected = new TaskCompletionSource<string?>(TaskCreationOptions.RunContinuationsAsynchronously);
            string output = Path.Combine(services.Paths.Root, "stale.html");
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null),
                pickSaveComparisonReport: _ => selected.Task, services: services);
            await model.OpenDocumentAsync(source); await model.OpenComparisonDocumentAsync(actual);
            model.ComparisonExpectedRange = "110"; model.ComparisonActualRange = "109";
            await model.ComparePagesCommand.ExecuteAsync(null);
            Task exporting = model.ExportComparisonReportCommand.ExecuteAsync(null);
            model.ComparisonExpectedRange = "1";
            selected.SetResult(output); await exporting; Assert.False(File.Exists(output));
            model.ComparisonExpectedRange = "110"; await model.ComparePagesCommand.ExecuteAsync(null);
            Create(actual, false);
            await model.ExportComparisonReportCommand.ExecuteAsync(null);
            Assert.False(File.Exists(output)); Assert.Contains("source PDF changed", model.ErrorMessage!, StringComparison.OrdinalIgnoreCase);
            return true;
        }, default);
    }

    private static void Create(string path, bool comparison) => PdfDocument.Create(document => {
        for (int i = 1; i <= 110; i++) document.Page(page => page.Size(150, 180).Margin(12).Content(content => {
            if (i == (comparison ? 109 : 110)) content.Item(item => item.Paragraph(text => text.Text(comparison ? "Revised last section" : "Original last section")));
        }));
    }).Save(path);
    private static void Capture(Avalonia.Controls.Window window, string name) {
        using var frame = window.CaptureRenderedFrame(); Assert.NotNull(frame);
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(output)) return;
        Directory.CreateDirectory(output); frame.Save(Path.Combine(output, name + ".png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
