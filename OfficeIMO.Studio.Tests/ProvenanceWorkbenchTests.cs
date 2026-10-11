using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.VisualTree;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Localization;
using OfficeIMO.Studio.Infrastructure.Preferences;
using OfficeIMO.Workflows;
using System.Globalization;
using System.Text.Json;
using Xunit;

namespace OfficeIMO.Studio.Tests;

public sealed class ProvenanceWorkbenchTests {
    [Theory]
    [InlineData("en", 840, 600)]
    [InlineData("pl", 840, 600)]
    [InlineData("de", 840, 600)]
    [InlineData("fr", 840, 600)]
    [InlineData("en", 1280, 800)]
    public async Task AggregateAiDeclarationRemainsVisibleWhenItsActionIsOutsideTheDisplayedTimeline(string culture, int width, int height) {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(() => {
            var previous = StudioLocalization.Current;
            var previousCulture = CultureInfo.CurrentCulture;
            var previousUi = CultureInfo.CurrentUICulture;
            var previousDefault = CultureInfo.DefaultThreadCurrentCulture;
            var previousDefaultUi = CultureInfo.DefaultThreadCurrentUICulture;
            var paths = ((App)Application.Current!).Services.Paths;
            new JsonStudioPreferencesStore(paths.PreferencesPath).Save(new StudioPreferences { UiCulture = culture });
            var services = StudioApplicationServices.Create(paths);
            StudioLocalization.Configure(services.Localizer);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services);
            model.ShowProvenanceCommand.Execute(null);
            var actions = Enumerable.Range(0, 64).Select(index => new ProvenanceActionDto("c2pa.edited", "Editor " + index,
                null, "DigitalCapture", null)).ToArray();
            model.ProvenanceWorkbench.Credentials.Add(new("PNG/caBX", new ProvenanceManifestDto("active", "Generator", "image.png", "image/png",
                actions, [], null, null, 1, true)));
            var view = new ProvenanceWorkbenchView { DataContext = model };
            var window = new Window { Width = width, Height = height, Content = view };
            try {
                window.Show(); window.UpdateLayout();
                Capture(window, culture + "-provenance-aggregate-ai-" + width);
                var indicator = Assert.Single(view.GetVisualDescendants().OfType<TextBlock>(),
                    text => text.Text == services.Localizer.Get("Provenance.SourceGenerativeAi") && text.IsEffectivelyVisible);
                indicator.BringIntoView(); window.UpdateLayout();
                Assert.InRange(indicator.TranslatePoint(new Point(0, 0), window)!.Value.Y, 0, height - indicator.Bounds.Height);
                Capture(window, culture + "-provenance-aggregate-ai-" + width);
                return Task.FromResult(true);
            } finally {
                window.Close(); StudioLocalization.Configure(previous);
                CultureInfo.CurrentCulture = previousCulture; CultureInfo.CurrentUICulture = previousUi;
                CultureInfo.DefaultThreadCurrentCulture = previousDefault; CultureInfo.DefaultThreadCurrentUICulture = previousDefaultUi;
            }
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData("en", 840, 600)]
    [InlineData("en", 1280, 800)]
    [InlineData("pl", 840, 600)]
    [InlineData("de", 840, 600)]
    [InlineData("fr", 840, 600)]
    public async Task CredentialPanelShowsRecordedActionsAndClearsForTheNextInput(string culture, int width, int height) {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            IStudioLocalizer previous = StudioLocalization.Current;
            var previousTheme = Application.Current!.RequestedThemeVariant;
            CultureInfo previousCulture = CultureInfo.CurrentCulture, previousUi = CultureInfo.CurrentUICulture;
            CultureInfo? previousDefault = CultureInfo.DefaultThreadCurrentCulture, previousDefaultUi = CultureInfo.DefaultThreadCurrentUICulture;
            Window? window = null;
            try {
                var paths = ((App)Application.Current!).Services.Paths;
                new JsonStudioPreferencesStore(paths.PreferencesPath).Save(new StudioPreferences {
                    UiCulture = culture, Theme = width >= 1000 ? StudioThemePreference.Dark : StudioThemePreference.Light
                });
                var services = StudioApplicationServices.Create(paths);
                StudioLocalization.Configure(services.Localizer);
                Application.Current.RequestedThemeVariant = width >= 1000 ? Avalonia.Styling.ThemeVariant.Dark : Avalonia.Styling.ThemeVariant.Light;
                IStudioLocalizer localizer = services.Localizer;
                Directory.CreateDirectory(services.Paths.Root);
                string source = Path.Combine(services.Paths.Root, "credential.png");
                File.Copy(Path.Combine(AppContext.BaseDirectory, "Fixtures", "unsigned-credential-12-actions.png"), source);
                using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services);
                model.ShowProvenanceCommand.Execute(null);
                var workspace = model.ProvenanceWorkbench;
                workspace.InputPath = source;
                workspace.OutputFolder = services.Paths.Root;
                await workspace.AssessCommand.ExecuteAsync(null);
                var credential = Assert.Single(workspace.Credentials);
                Assert.True(workspace.HasCredentials, workspace.Status);
                Assert.Equal(12, credential.Actions.Count);
                Assert.Equal("Generator", credential.Generator);
                Assert.Equal("Unverified subject", credential.CertificateSubject);
                Assert.Equal(localizer.Get("Provenance.TimeNotRecorded"), credential.Actions[1].Time);
                Assert.Equal(localizer.Get("Provenance.ActionWatermarked"), credential.Actions[1].Action);
                Assert.True(credential.Actions[1].HasWatermark);
                Assert.Contains("Verification: NotConfigured", workspace.Checks);
                var view = new ProvenanceWorkbenchView { DataContext = model };
                window = new Window { Width = width, Height = height, Content = view };
                window.Show(); window.UpdateLayout();
                Assert.Contains(view.GetVisualDescendants().OfType<TextBlock>(), text => text.Text == localizer.Get("Provenance.Timeline"));
                TextBlock subject = view.GetVisualDescendants().OfType<TextBlock>().Single(text => text.Text == localizer.Format("Provenance.CertificateSubjectFormat", "Unverified subject"));
                subject.BringIntoView(); window.UpdateLayout();
                Assert.True(subject.IsEffectivelyVisible);
                Capture(window, culture + "-provenance-credentials-" + width);
                TextBlock watermark = view.GetVisualDescendants().OfType<TextBlock>().Single(text => text.IsVisible && text.Text == localizer.Get("Provenance.WatermarkHint"));
                watermark.BringIntoView(); window.UpdateLayout();
                Capture(window, culture + "-provenance-timeline-" + width);
                TextBlock finalAction = view.GetVisualDescendants().OfType<TextBlock>().Single(text => text.Text == "Editor 11");
                finalAction.BringIntoView(); window.UpdateLayout();
                AvaloniaHeadlessPlatform.ForceRenderTimerTick();
                var finalPosition = finalAction.TranslatePoint(new Point(0, 0), window);
                Assert.NotNull(finalPosition);
                Assert.InRange(finalPosition.Value.Y, 0, height - finalAction.Bounds.Height);
                Capture(window, culture + "-provenance-timeline-last-" + width);
                await workspace.ExportReportCommand.ExecuteAsync(null);
                using var report = JsonDocument.Parse(await File.ReadAllTextAsync(workspace.ReportPath));
                Assert.Equal(12, report.RootElement.GetProperty("assessment").GetProperty("structural").GetProperty("evidence")[0].GetProperty("manifest").GetProperty("actions").GetArrayLength());
                workspace.RemoveManifests = true;
                await workspace.CreateCopyCommand.ExecuteAsync(null);
                Assert.True(File.Exists(workspace.OutputPath), workspace.Status);
                Assert.False(workspace.HasCredentials);
                workspace.InputPath = source + ".other";
                Assert.Empty(workspace.Credentials);
                Assert.False(workspace.CanExportReport);
                Assert.False(workspace.CanCreateCopy);
            } finally {
                window?.Close();
                Application.Current!.RequestedThemeVariant = previousTheme;
                StudioLocalization.Configure(previous);
                CultureInfo.CurrentCulture = previousCulture; CultureInfo.CurrentUICulture = previousUi;
                CultureInfo.DefaultThreadCurrentCulture = previousDefault; CultureInfo.DefaultThreadCurrentUICulture = previousDefaultUi;
            }
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData(840, 600)]
    [InlineData(1280, 800)]
    public async Task ReviewCopyAndExportRemainAccessibleWithoutChangingTheSource(int width, int height) {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            services.Preferences.Update(current => current with { Theme = width >= 1000 ? StudioThemePreference.Dark : StudioThemePreference.Light });
            string source = Path.Combine(services.Paths.Root, "source.html");
            Directory.CreateDirectory(services.Paths.Root);
            string original = "<!doctype html><html><head><link rel=\"c2pa-manifest\" href=\"claim.c2pa\"></head><body>review\u202Ethis</body></html>";
            await File.WriteAllTextAsync(source, original);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services,
                pickProvenanceFile: _ => Task.FromResult<string?>(source), pickOutputFolder: _ => Task.FromResult<string?>(services.Paths.Root));
            model.ShowProvenanceCommand.Execute(null);
            Assert.True(model.IsProvenanceMode);
            Assert.Contains(model.Commands.Items, item => item.Id == "Provenance");
            var workspace = model.ProvenanceWorkbench;
            await workspace.ChooseInputCommand.ExecuteAsync(null);
            await workspace.ChooseFolderCommand.ExecuteAsync(null);
            Assert.False(workspace.CanCreateCopy);
            await workspace.AssessCommand.ExecuteAsync(null);
            Assert.Contains(workspace.Findings, item => item.Contains("U+202E"));
            Assert.Contains("Verification: NotConfigured", workspace.Checks);
            workspace.RemoveReferences = true;
            Assert.True(workspace.CanCreateCopy);
            var view = new ProvenanceWorkbenchView { DataContext = model };
            var window = new Window { Width = width, Height = height, Content = view };
            try {
                window.Show(); window.UpdateLayout();
                Capture(window, "provenance-review-" + width);
                var provider = view.GetVisualDescendants().OfType<Expander>().Single();
                provider.IsExpanded = true; provider.BringIntoView(); window.UpdateLayout();
                Capture(window, "provenance-provider-" + width);
                var providerButton = view.GetVisualDescendants().OfType<Button>().Single(button => ReferenceEquals(button.Command, workspace.CheckProviderCommand));
                providerButton.BringIntoView(); window.UpdateLayout();
                Assert.True(providerButton.IsEffectivelyVisible);
                provider.IsExpanded = false;
                var copyButton = view.GetVisualDescendants().OfType<Button>().Single(button => ReferenceEquals(button.Command, workspace.CreateCopyCommand));
                copyButton.BringIntoView(); window.UpdateLayout();
                Assert.True(copyButton.IsEffectivelyVisible); Assert.True(copyButton.IsEnabled);
                Capture(window, "provenance-options-" + width);
                await workspace.CreateCopyCommand.ExecuteAsync(null);
                Assert.True(File.Exists(workspace.OutputPath), workspace.Status);
                Assert.Equal(original, await File.ReadAllTextAsync(source));
                Assert.NotEmpty(workspace.Changes);
                Assert.Contains("‮", await File.ReadAllTextAsync(workspace.OutputPath));
                await workspace.ExportReportCommand.ExecuteAsync(null);
                var exported = services.Jobs.Entries[0];
                Assert.True(exported.IsSucceeded, exported.Summary);
                Assert.Equal(workspace.ReportPath, exported.OutputPath);
                using var report = JsonDocument.Parse(await File.ReadAllTextAsync(workspace.ReportPath));
                Assert.Equal("officeimo.provenance.result.v2", report.RootElement.GetProperty("schema").GetString());
                Assert.Equal(workspace.OutputHash, report.RootElement.GetProperty("outputSha256").GetString());
                copyButton.BringIntoView(); window.UpdateLayout(); Capture(window, "provenance-copy-" + width);
                var jobs = new StudioJobsView { DataContext = model.Jobs };
                window.Content = jobs; window.UpdateLayout();
                Assert.Contains(jobs.GetVisualDescendants().OfType<Button>(), button =>
                    ReferenceEquals(button.Command, model.Jobs.OpenOutputCommand) &&
                    ReferenceEquals(button.CommandParameter, exported) && button.IsEnabled);
                Capture(window, "provenance-jobs-" + width);
                workspace.C2paToolPath = Path.Combine(services.Paths.Root, "missing-c2patool");
                Assert.False(workspace.CanCreateCopy); Assert.False(workspace.CanExportReport);
                await workspace.CheckProviderCommand.ExecuteAsync(null);
                Assert.Contains("Unavailable", workspace.ProviderStatus);
                Assert.False(workspace.IsBusy);
                workspace.InputPath = Path.Combine(services.Paths.Root, "other.html");
                Assert.False(workspace.CanCreateCopy); Assert.False(workspace.CanExportReport);
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }
    private static void Capture(Window window, string name) {
        AvaloniaHeadlessPlatform.ForceRenderTimerTick();
        using var frame = window.CaptureRenderedFrame(); Assert.NotNull(frame);
        string? folder = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(folder)) return;
        Directory.CreateDirectory(folder); frame.Save(Path.Combine(folder, name + ".png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
