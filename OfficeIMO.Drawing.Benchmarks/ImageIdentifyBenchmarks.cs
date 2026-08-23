using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Drawing.Benchmarks;

/// <summary>Measures metadata identification separately from full raster decoding.</summary>
[MemoryDiagnoser]
public class ImageIdentifyBenchmarks {
    private byte[] _encoded = null!;

    [Params("LogoPng", "PhotoJpeg", "AnimationGif", "SaturnTiff", "SnailBmp")]
    public string Asset { get; set; } = string.Empty;

    [GlobalSetup]
    public void Setup() {
        ImageBenchmarkAsset asset = ImageBenchmarkCorpus.Get(Asset);
        _encoded = asset.ReadBytes();
        ImageBenchmarkCorpus.AssertIdentified(_encoded, asset.Format, asset.Width, asset.Height, Asset);
    }

    [Benchmark(Baseline = true)]
    public OfficeImageInfo IdentifyBytes() => OfficeImageReader.Identify(_encoded);

    [Benchmark]
    public OfficeImageInfo IdentifySeekableStream() {
        using var stream = new MemoryStream(_encoded, writable: false);
        if (!OfficeImageReader.TryIdentify(stream, fileName: null, out OfficeImageInfo info)) {
            throw new InvalidOperationException("The benchmark image could not be identified.");
        }
        return info;
    }
}
