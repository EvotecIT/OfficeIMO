using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.VisualTree;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Studio.Infrastructure.Preferences;
using System.Text.Json;
using Xunit;

namespace OfficeIMO.Studio.Tests;

public sealed class ProvenanceWorkbenchTests {
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
