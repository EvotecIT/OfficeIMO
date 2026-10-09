using System.Text.Json;
using OfficeIMO.Ocr;
using OfficeIMO.Tool.Commands.Pdf;
using Xunit;

namespace OfficeIMO.Tool.Tests;

public sealed class PdfOcrProviderCommandTests {
    [Fact]
    public async Task ExplicitInstalledProviderDeploymentIsDiscoveredWithoutExecutingTesseract() {
        using var output = new StringWriter(); using var error = new StringWriter();
        int exit = await PdfCommand.RunAsync(["providers", "--ocr-provider-assembly",
            typeof(OfficeIMO.Ocr.Tesseract.TesseractOcrEngineProvider).Assembly.Location], output, error);
        Assert.True(exit == 0, error.ToString());
        using JsonDocument json = JsonDocument.Parse(output.ToString());
        Assert.Equal("tesseract-cli", Assert.Single(json.RootElement.EnumerateArray()).GetProperty("id").GetString());
        Assert.Equal(string.Empty, error.ToString());
        var provider = new OfficeIMO.Ocr.Tesseract.TesseractOcrEngineProvider();
        Assert.Throws<ArgumentException>(() => provider.Create(new Dictionary<string, string> { ["arbitraryArguments"] = "--help" }));
        Assert.Throws<ArgumentException>(() => provider.Create(new Dictionary<string, string> { ["pageSegmentationMode"] = "14" }));
    }
    [Fact]
    public async Task ProviderDiscoveryListsInjectedOptionalProviders() {
        var catalog = new OcrEngineCatalog().Register(new FixtureProvider());
        using var output = new StringWriter();
        using var error = new StringWriter();

        int exitCode = await PdfCommand.RunAsync(
            ["redact", "providers"],
            output,
            error,
            ocrCatalog: catalog);

        Assert.Equal((int)OfficeImoToolExitCode.Success, exitCode);
        using JsonDocument json = JsonDocument.Parse(output.ToString());
        JsonElement provider = Assert.Single(json.RootElement.EnumerateArray());
        Assert.Equal("fixture-cli", provider.GetProperty("id").GetString());
        Assert.True(provider.GetProperty("capabilities").GetProperty("supportsWordSpans").GetBoolean());
        Assert.Equal(string.Empty, error.ToString());
    }

    [Fact]
    public void RedactionArgumentsCaptureSelectedProviderAndBoundedScalarConfiguration() {
        PdfArguments parsed = PdfArguments.Parse([
            "redact", "plan", "source.pdf",
            "--recipe", "recipe.json",
            "--evidence", "plan.json",
            "--ocr-provider", "fixture-cli",
            "--ocr-language", "en",
            "--ocr-min-confidence", "0.75",
            "--ocr-option", "model=document"
        ]);

        Assert.Equal("fixture-cli", parsed.OcrProviderId);
        Assert.Equal("en", parsed.OcrLanguage);
        Assert.Equal(0.75D, parsed.OcrMinimumConfidence);
        Assert.Equal("document", parsed.OcrProviderOptions["model"]);
    }

    [Fact]
    public void ProviderAssemblyLoadingRejectsAnUnboundedPathCollectionBeforeLoading() {
        string[] paths = Enumerable.Repeat("provider.dll", 33).ToArray();

        ArgumentException exception = Assert.Throws<ArgumentException>(() =>
            PdfOcrProviderLoader.LoadExplicitAssemblies(new OcrEngineCatalog(), paths));

        Assert.Contains("cannot exceed 32 entries", exception.Message, StringComparison.Ordinal);
    }

    private sealed class FixtureProvider : IOcrEngineProvider {
        public string Id => "fixture-cli";
        public string DisplayName => "Fixture CLI OCR";
        public OcrEngineCapabilities Capabilities => new() { SupportsWordSpans = true };
        public IOcrEngine Create(IReadOnlyDictionary<string, string> options) =>
            new DelegateOcrEngine(Id, (_, _) => Task.FromResult(new OcrResult()), Capabilities);
    }
}
