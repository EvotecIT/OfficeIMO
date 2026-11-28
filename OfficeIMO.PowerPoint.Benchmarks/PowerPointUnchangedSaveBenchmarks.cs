using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Engines;
using System.IO.Compression;

namespace OfficeIMO.PowerPoint.Benchmarks;

/// <summary>
/// Loads a saved 200-slide deck with <see cref="PowerPointDocument.Load(Stream, PowerPointLoadOptions?)"/> and saves it
/// unchanged with <see cref="PowerPointDocument.ToBytes"/>: once with the Open XML SDK package opened, as every load
/// did before, and once without it. Setup checks that the second never opens the SDK and that both write the same
/// parts with the same content.
/// </summary>
[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
public class PowerPointUnchangedSaveBenchmarks {
    private byte[] _deck = Array.Empty<byte>();

    [Params(200)]
    public int Slides { get; set; }

    [GlobalSetup]
    public void Setup() {
        _deck = PowerPointUnchangedSaveWorkload.ReadDeck(Slides);
        PowerPointUnchangedSaveWorkload.Check(_deck, OpenXmlSdk(), WithoutSdk());
    }

    [Benchmark(Baseline = true)]
    public byte[] OpenXmlSdk() {
        using PowerPointDocument presentation = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        _ = presentation.OpenXmlDocument;
        return presentation.ToBytes();
    }

    [Benchmark]
    public byte[] WithoutSdk() {
        using PowerPointDocument presentation = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        return presentation.ToBytes();
    }
}

/// <summary>
/// The first load and save in a fresh process, for each path. The deck comes from a file the steady-state run writes,
/// so setup warms neither path.
/// </summary>
[SimpleJob(RunStrategy.ColdStart, launchCount: 15, warmupCount: 0, iterationCount: 1)]
[BenchmarkCategory("Performance")]
public class PowerPointUnchangedSaveFirstCallBenchmarks {
    private byte[] _deck = Array.Empty<byte>();

    [Params(200)]
    public int Slides { get; set; }

    [GlobalSetup]
    public void Setup() => _deck = PowerPointUnchangedSaveWorkload.ReadDeck(Slides);

    [Benchmark(Baseline = true)]
    public byte[] OpenXmlSdk() {
        using PowerPointDocument presentation = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        _ = presentation.OpenXmlDocument;
        return presentation.ToBytes();
    }

    // Setup cannot check the path without warming it, so the run checks it afterwards.
    [Benchmark]
    public byte[] WithoutSdk() {
        using PowerPointDocument presentation = PowerPointDocument.Load(new MemoryStream(_deck, writable: false));
        byte[] saved = presentation.ToBytes();
        if (presentation.IsSdkLoaded) throw new InvalidOperationException("The unchanged save opened the Open XML SDK.");
        return saved;
    }
}

internal static class PowerPointUnchangedSaveWorkload {
    private static string DeckPath(int slides) => Path.Combine(Path.GetTempPath(), "officeimo-unchanged-save-v2-" + slides + ".pptx");

    // The deck the model benchmarks write, saved once through PowerPointDocument so its document properties are
    // complete as PowerPoint writes them, and kept in a file so a cold run reads bytes only.
    internal static byte[] ReadDeck(int slides) {
        string path = DeckPath(slides);
        if (!File.Exists(path)) {
            using PowerPointDocument presentation = PowerPointDocument.Load(new MemoryStream(PowerPointModelWorkload.CreateWithSdk(slides), writable: false));
            _ = presentation.OpenXmlDocument;
            File.WriteAllBytes(path, presentation.ToBytes());
        }
        return File.ReadAllBytes(path);
    }

    internal static void Check(byte[] deck, byte[] throughSdk, byte[] withoutSdk) {
        using (PowerPointDocument presentation = PowerPointDocument.Load(new MemoryStream(deck, writable: false))) {
            presentation.ToBytes();
            if (presentation.IsSdkLoaded) throw new InvalidOperationException("The unchanged save opened the Open XML SDK.");
        }
        Dictionary<string, long> expected = Parts(throughSdk);
        Dictionary<string, long> actual = Parts(withoutSdk);
        if (!expected.Keys.OrderBy(name => name, StringComparer.Ordinal).SequenceEqual(actual.Keys.OrderBy(name => name, StringComparer.Ordinal))) {
            throw new InvalidOperationException("The two saves wrote different parts.");
        }
        using var reopened = PowerPointDocument.Load(new MemoryStream(withoutSdk, writable: false));
        using var reference = PowerPointDocument.Load(new MemoryStream(throughSdk, writable: false));
        if (reopened.Slides.Count != reference.Slides.Count) {
            throw new InvalidOperationException("The two saves hold different slides.");
        }
    }

    private static Dictionary<string, long> Parts(byte[] package) {
        using var archive = new ZipArchive(new MemoryStream(package, writable: false), ZipArchiveMode.Read);
        return archive.Entries.ToDictionary(entry => entry.FullName, entry => entry.Length, StringComparer.Ordinal);
    }
}
