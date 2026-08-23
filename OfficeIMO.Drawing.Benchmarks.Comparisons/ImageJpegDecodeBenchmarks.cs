using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Drawing.Benchmarks.Comparisons;

/// <summary>Decodes one identical photographic JPEG completely in every engine.</summary>
[MemoryDiagnoser]
public class ImageJpegDecodeBenchmarks {
    private byte[] _encoded = null!;

    [GlobalSetup]
    public void Setup() {
        _encoded = ImageBenchmarkCorpus.Photo.ReadBytes();
        ImageComparisonValidation.MeasureJpegDecodeError(_encoded, 2048, 1619);
    }

    [Benchmark(Baseline = true)]
    public int OfficeIMO() => ImageComparisonAdapters.DecodeOfficeImo(_encoded);

    [Benchmark]
    public int SkiaSharp() => ImageComparisonAdapters.DecodeSkia(_encoded);

    [Benchmark]
    public int MagickNET() => ImageComparisonAdapters.DecodeMagick(_encoded);

    [Benchmark]
    public int StbImageSharp() => ImageComparisonAdapters.DecodeStb(_encoded);
}
