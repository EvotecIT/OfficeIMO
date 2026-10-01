using OfficeIMO.Ocr.Tesseract;
using OfficeIMO.Studio.Infrastructure;
using Xunit;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioDistributionPolicyTests {
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
