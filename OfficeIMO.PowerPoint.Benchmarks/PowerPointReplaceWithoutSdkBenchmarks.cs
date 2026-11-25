using BenchmarkDotNet.Attributes;

namespace OfficeIMO.PowerPoint.Benchmarks;

/// <summary>
/// Loads the saved 200-slide deck, replaces a word that appears on every slide with
/// <see cref="PowerPointDocument.ReplaceAll"/> and saves it: once with the Open XML SDK package opened, as every load
/// did before, and once without it. Setup checks that the second never opens the SDK and that both replace the same
/// number of matches and write the same slide text.
/// </summary>
[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
public class PowerPointReplaceWithoutSdkBenchmarks {
    private byte[] _deck = Array.Empty<byte>();

    [Params(200)]
    public int Slides { get; set; }

    [GlobalSetup]
    public void Setup() {
        _deck = PowerPointUnchangedSaveWorkload.ReadDeck(Slides);
        int sdkCount, ownedCount;
        byte[] sdk, owned;
        using (PowerPointDocument presentation = PowerPointDocument.Load(new MemoryStream(_deck, writable: false))) {
            _ = presentation.OpenXmlDocument;
            sdkCount = presentation.ReplaceAll("point", "item");
            sdk = presentation.ToBytes();
        }
        using (PowerPointDocument presentation = PowerPointDocument.Load(new MemoryStream(_deck, writable: false))) {
            ownedCount = presentation.ReplaceAll("point", "item");
            owned = presentation.ToBytes();
            if (presentation.IsSdkLoaded) throw new InvalidOperationException("The replacement opened the Open XML SDK.");
        }
        if (sdkCount != ownedCount || sdkCount == 0) throw new InvalidOperationException("The replacement counts differ or found nothing.");
        if (Text(sdk) != Text(owned)) throw new InvalidOperationException("The saved slide text differs.");
    }

    [Benchmark(Baseline = true)]
    public byte[] OpenXmlSdk() {
        using PowerPointDocument presentation = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        _ = presentation.OpenXmlDocument;
        presentation.ReplaceAll("point", "item");
        return presentation.ToBytes();
    }

    [Benchmark]
    public byte[] WithoutSdk() {
        using PowerPointDocument presentation = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        presentation.ReplaceAll("point", "item");
        return presentation.ToBytes();
    }

    private static string Text(byte[] package) {
        using PowerPointDocument presentation = PowerPointDocument.Load(new MemoryStream(package, writable: false));
        return string.Concat(presentation.ExtractMarkdownChunks().Select(chunk => chunk.Text));
    }
}
