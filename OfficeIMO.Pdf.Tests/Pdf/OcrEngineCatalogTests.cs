using OfficeIMO.Ocr;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class OcrEngineCatalogTests {
    [Fact]
    public void CatalogDiscoversAndCreatesExplicitProvidersFromBoundedOptions() {
        var provider = new RecordingProvider();
        var catalog = new OcrEngineCatalog().Register(provider);

        OcrEngineDescriptor descriptor = Assert.Single(catalog.Discover());
        IOcrEngine engine = catalog.Create("FIXTURE", new Dictionary<string, string> { ["model"] = "fast" });

        Assert.Equal("fixture", descriptor.Id);
        Assert.Equal("Fixture OCR", descriptor.DisplayName);
        Assert.True(descriptor.Capabilities.SupportsWordSpans);
        Assert.Equal("fixture", engine.Id);
        Assert.Equal("fast", provider.Options!["model"]);
    }

    [Fact]
    public void CatalogRejectsDuplicateIdsAndOversizedConfiguration() {
        var catalog = new OcrEngineCatalog().Register(new RecordingProvider());
        Assert.Throws<ArgumentException>(() => catalog.Register(new RecordingProvider()));
        var options = Enumerable.Range(0, 129).ToDictionary(index => "k" + index, _ => "v", StringComparer.Ordinal);
        Assert.Throws<ArgumentException>(() => catalog.Create("fixture", options));
    }

    private sealed class RecordingProvider : IOcrEngineProvider {
        internal IReadOnlyDictionary<string, string>? Options { get; private set; }
        public string Id => "fixture";
        public string DisplayName => "Fixture OCR";
        public OcrEngineCapabilities Capabilities => new() { SupportsWordSpans = true };
        public IOcrEngine Create(IReadOnlyDictionary<string, string> options) {
            Options = options;
            return new DelegateOcrEngine(Id, (_, _) => Task.FromResult(new OcrResult()), Capabilities);
        }
    }
}
