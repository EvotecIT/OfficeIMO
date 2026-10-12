using OfficeIMO.Ocr.Tesseract;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Features.Workflows;
using Xunit;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioDistributionPolicyTests {
    [Fact]
    public async Task StoreChannelBlocksInvoiceStandardsBeforeRuleDiscovery() {
        using var model = new InvoiceWorkbenchViewModel(_ => Task.FromResult<string?>(null), _ => Task.FromResult<string?>(null)) {
            InputPath = "invoice.xml", RequireStandards = true,
            RuleBundlePath = "missing-rule-bundle", SaxonJarPath = "missing-saxon.jar", JavaExecutable = "missing-java"
        };
        Assert.Equal(StudioDistributionPolicy.ExternalToolsAllowed, model.IsStandardsAvailable);
#if OFFICEIMO_MAC_APP_STORE
        await model.RunCommand.ExecuteAsync(null);
        Assert.Contains("Mac App Store edition", model.Status);
        Assert.False(model.HasOutput);
#else
        Assert.Contains("Java/Saxon", model.StandardsHint);
        await Task.CompletedTask;
#endif
    }

    [Fact]
    public async Task StoreChannelBlocksExternalOcrBeforeToolDiscovery() {
#if OFFICEIMO_MAC_APP_STORE
        Assert.True(StudioDistributionPolicy.IsMacAppStore);
        var error = await Assert.ThrowsAsync<NotSupportedException>(() => StudioOcrProvider.CreateSessionAsync(
            new TesseractOcrSessionOptions(), CancellationToken.None));
        Assert.Contains("Mac App Store edition", error.Message);
#else
        Assert.False(StudioDistributionPolicy.IsMacAppStore);
        Assert.True(StudioDistributionPolicy.ExternalToolsAllowed);
        await Task.CompletedTask;
#endif
    }

    [Fact]
    public async Task StoreChannelDoesNotDiscoverSystemPrinterTools() {
#if OFFICEIMO_MAC_APP_STORE
        var error = await Assert.ThrowsAsync<NotSupportedException>(() => StudioPrinterService.Create().GetPrintersAsync());
        Assert.Contains("Save the prepared PDF", error.Message);
#else
        Assert.True(StudioDistributionPolicy.ExternalToolsAllowed);
        await Task.CompletedTask;
#endif
    }

    [Fact]
    public void StoreChannelOffersOnlyWorkThatReachesPickedLocations() {
        // Sandboxed builds reach user files only through a picker: no external verifier, no home-folder login import,
        // no folder-wide batch export and no output written next to a picked file.
        bool picked = StudioDistributionPolicy.RequiresPickedLocations;
        Assert.Equal(StudioDistributionPolicy.IsMacAppStore || OperatingSystem.IsIOS() || OperatingSystem.IsAndroid(), picked);
        using var provenance = new ProvenanceWorkbenchViewModel(_ => Task.FromResult<string?>(null), _ => Task.FromResult<string?>(null));
        Assert.Equal(StudioDistributionPolicy.ExternalToolsAllowed, provenance.CanConfigureProvider);
        Assert.Equal(!picked, Features.Assistant.StudioAiConnections.CodexImportSupported);
        using var invoice = new InvoiceWorkbenchViewModel(_ => Task.FromResult<string?>(null), _ => Task.FromResult<string?>(null)) {
            InputPath = Path.Combine(Path.GetTempPath(), "invoice.xml")
        };
        invoice.SelectedOperation = invoice.Operations.Single(o => o.Value == OfficeIMO.Workflows.OfficeInvoiceWorkflowOperation.Convert);
        Assert.Equal(!picked, invoice.CanRun);
    }

    [Fact]
    public async Task LocationBoxesAreReadOnlyWhereOnlyPickedLocationsAreReachable() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(() => {
            // The application applies the policy of this build; a window can apply the sandboxed policy explicitly.
            var current = new Avalonia.Controls.TextBox { Classes = { "path" } };
            var location = new Avalonia.Controls.TextBox { Classes = { "path" } };
            var other = new Avalonia.Controls.TextBox();
            var window = new Avalonia.Controls.Window { Content = new Avalonia.Controls.StackPanel { Children = { current } } };
            var sandboxed = new Avalonia.Controls.Window { Content = new Avalonia.Controls.StackPanel { Children = { location, other } } };
            App.ApplyLocationPolicy(sandboxed.Styles, requiresPickedLocations: true);
            try {
                window.Show(); sandboxed.Show();
                Assert.Equal(StudioDistributionPolicy.RequiresPickedLocations, current.IsReadOnly);
                Assert.True(location.IsReadOnly);
                Assert.False(other.IsReadOnly);
            } finally { window.Close(); sandboxed.Close(); }
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public void StoreChannelKeepsDataInThePlatformApplicationDataLocation() {
        string? previous = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_DATA_ROOT");
        string overridePath = Path.Combine(Path.GetTempPath(), "outside-app-container");
        try {
            Environment.SetEnvironmentVariable("OFFICEIMO_STUDIO_DATA_ROOT", overridePath);
#if OFFICEIMO_MAC_APP_STORE
            Assert.Equal(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OfficeIMO", "Studio"), StudioDataPaths.CreateDefault().Root);
#else
            Assert.Equal(overridePath, StudioDataPaths.CreateDefault().Root);
#endif
        } finally { Environment.SetEnvironmentVariable("OFFICEIMO_STUDIO_DATA_ROOT", previous); }
    }
}
