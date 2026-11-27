using BenchmarkDotNet.Attributes;

namespace OfficeIMO.PowerPoint.Benchmarks;

/// <summary>
/// Binary PowerPoint conversion: importing a .ppt into the editable model, and writing a presentation (the same deck
/// converted to .pptx first) with the native .ppt writer. Fixtures come from the legacy .ppt test corpus.
/// </summary>
[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
public class PowerPointLegacyPptBenchmarks {
    // The corpus decks carry content the native writer reports as lossy (notes drawings, connectors).
    private static readonly PowerPointSaveOptions WriteOptions = new() {
        Format = PowerPointFileFormat.Ppt,
        LossPolicy = OfficeConversionLossPolicy.Allow
    };

    private byte[] _ppt = Array.Empty<byte>();
    private byte[] _pptx = Array.Empty<byte>();

    [Params("PicturePowerPoint.ppt", "BasicPowerPoint.ppt", "RichTextPowerPoint.ppt", "apache-poi-testdata/bug61881.ppt")]
    public string Fixture { get; set; } = string.Empty;

    [GlobalSetup]
    public void Setup() {
        _ppt = File.ReadAllBytes(Path.Combine(FindCorpus(), Fixture));
        using PowerPointDocument presentation = PowerPointDocument.Load(new MemoryStream(_ppt, writable: false));
        if (presentation.Slides.Count == 0) throw new InvalidOperationException("The fixture has no slides.");
        _pptx = presentation.ToBytes();
        using PowerPointDocument converted = PowerPointDocument.Load(new MemoryStream(_pptx, writable: false));
        if (converted.ToBytes(WriteOptions).Length == 0) {
            throw new InvalidOperationException("The native writer produced nothing.");
        }
    }

    [Benchmark]
    public int Import() {
        using PowerPointDocument presentation = PowerPointDocument.Load(new MemoryStream(_ppt, writable: false));
        return presentation.Slides.Count;
    }

    [Benchmark]
    public byte[] WriteNative() {
        using PowerPointDocument presentation = PowerPointDocument.Load(new MemoryStream(_pptx, writable: false));
        return presentation.ToBytes(WriteOptions);
    }

    private static string FindCorpus() {
        string? directory = AppContext.BaseDirectory;
        while (directory != null) {
            string candidate = Path.Combine(directory, "OfficeIMO.TestAssets", "Documents", "LegacyPptCorpus");
            if (Directory.Exists(candidate)) return candidate;
            directory = Path.GetDirectoryName(directory);
        }
        throw new DirectoryNotFoundException("The legacy .ppt corpus was not found above the benchmark output.");
    }
}
