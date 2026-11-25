using BenchmarkDotNet.Attributes;

namespace OfficeIMO.PowerPoint.Benchmarks;

/// <summary>
/// Loads the saved 200-slide deck and counts its slides, or extracts its text as Markdown chunks: once with the Open
/// XML SDK package opened, as every load did before, and once without it. Setup checks that the second never opens
/// the SDK and that both give the same count and chunks.
/// </summary>
[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
[GroupBenchmarksBy(BenchmarkDotNet.Configs.BenchmarkLogicalGroupRule.ByCategory)]
public class PowerPointTextWithoutSdkBenchmarks {
    private byte[] _deck = Array.Empty<byte>();

    [Params(200)]
    public int Slides { get; set; }

    [GlobalSetup]
    public void Setup() {
        _deck = PowerPointUnchangedSaveWorkload.ReadDeck(Slides);
        using (PowerPointDocument presentation = PowerPointDocument.Load(new MemoryStream(_deck, writable: false))) {
            presentation.ExtractMarkdownChunks().ToList();
            if (presentation.IsSdkLoaded) throw new InvalidOperationException("The extraction opened the Open XML SDK.");
        }
        if (CountWithSdk() != CountWithoutSdk()) throw new InvalidOperationException("The slide counts differ.");
        List<PowerPointExtractChunk> expected = ExtractWithSdk(), actual = ExtractWithoutSdk();
        if (!expected.Select(chunk => chunk.Id + chunk.Markdown).SequenceEqual(actual.Select(chunk => chunk.Id + chunk.Markdown))) {
            throw new InvalidOperationException("The extracted chunks differ.");
        }
    }

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("Count")]
    public int CountWithSdk() {
        using PowerPointDocument presentation = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        _ = presentation.OpenXmlDocument;
        return presentation.Slides.Count;
    }

    [Benchmark]
    [BenchmarkCategory("Count")]
    public int CountWithoutSdk() {
        using PowerPointDocument presentation = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        return presentation.Slides.Count;
    }

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("Extract")]
    public List<PowerPointExtractChunk> ExtractWithSdk() {
        using PowerPointDocument presentation = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        _ = presentation.OpenXmlDocument;
        return presentation.ExtractMarkdownChunks().ToList();
    }

    [Benchmark]
    [BenchmarkCategory("Extract")]
    public List<PowerPointExtractChunk> ExtractWithoutSdk() {
        using PowerPointDocument presentation = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        return presentation.ExtractMarkdownChunks().ToList();
    }
}
