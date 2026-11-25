using BenchmarkDotNet.Attributes;

namespace OfficeIMO.PowerPoint.Benchmarks;

/// <summary>
/// Loads the saved 200-slide deck, removes every tenth slide, moves ten slides from the end to the front and saves it:
/// once with the Open XML SDK package opened, as every load did before, and once without it. Setup checks that the
/// second never opens the SDK and that both save the same slides in the same order.
/// </summary>
[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
public class PowerPointSlideOrderWithoutSdkBenchmarks {
    private byte[] _deck = Array.Empty<byte>();

    [Params(200)]
    public int Slides { get; set; }

    [GlobalSetup]
    public void Setup() {
        _deck = PowerPointUnchangedSaveWorkload.ReadDeck(Slides);
        byte[] sdk, owned;
        using (PowerPointDocument presentation = PowerPointDocument.Load(new MemoryStream(_deck, writable: false))) {
            _ = presentation.OpenXmlDocument;
            Edit(presentation);
            sdk = presentation.ToBytes();
        }
        using (PowerPointDocument presentation = PowerPointDocument.Load(new MemoryStream(_deck, writable: false))) {
            Edit(presentation);
            owned = presentation.ToBytes();
            if (presentation.IsSdkLoaded) throw new InvalidOperationException("The slide edits opened the Open XML SDK.");
        }
        if (Text(sdk) != Text(owned)) throw new InvalidOperationException("The saved slides differ.");
    }

    [Benchmark(Baseline = true)]
    public byte[] OpenXmlSdk() {
        using PowerPointDocument presentation = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        _ = presentation.OpenXmlDocument;
        Edit(presentation);
        return presentation.ToBytes();
    }

    [Benchmark]
    public byte[] WithoutSdk() {
        using PowerPointDocument presentation = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        Edit(presentation);
        return presentation.ToBytes();
    }

    private void Edit(PowerPointDocument presentation) {
        for (int index = Slides - 10; index >= 0; index -= 10) presentation.RemoveSlide(index);
        int count = presentation.Slides.Count;
        for (int moved = 0; moved < 10; moved++) presentation.MoveSlide(count - 1, 0);
    }

    private static string Text(byte[] package) {
        using PowerPointDocument presentation = PowerPointDocument.Load(new MemoryStream(package, writable: false));
        return string.Concat(presentation.ExtractMarkdownChunks().Select(chunk => chunk.Text));
    }
}
