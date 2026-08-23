using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Drawing.Benchmarks;

/// <summary>Measures full managed decode into OfficeIMO's RGBA raster contract.</summary>
[MemoryDiagnoser]
public class ImageDecodeBenchmarks {
    private byte[] _encoded = null!;

    [Params("LogoPng", "PhotoJpeg", "AnimationGif", "SaturnTiff", "GeneratedBmp24")]
    public string Asset { get; set; } = string.Empty;

    [GlobalSetup]
    public void Setup() {
        ImageBenchmarkAsset? asset = ImageBenchmarkCorpus.All.SingleOrDefault(candidate => candidate.Id == Asset);
        _encoded = Asset == "GeneratedBmp24" ? ImageBenchmarkCorpus.CreateBmp24() : asset!.ReadBytes();
        OfficeRasterImage decoded = ImageBenchmarkCorpus.Decode(_encoded, Asset);
        int expectedWidth = asset?.Width ?? 256;
        int expectedHeight = asset?.Height ?? 256;
        if (decoded.Width != expectedWidth || decoded.Height != expectedHeight) {
            throw new InvalidOperationException($"{Asset} decoded to {decoded.Width}x{decoded.Height}.");
        }
    }

    [Benchmark]
    public OfficeRasterImage Decode() => ImageBenchmarkCorpus.Decode(_encoded, Asset);
}
