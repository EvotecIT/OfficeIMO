using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Drawing.Benchmarks.Comparisons;

/// <summary>Decodes one identical PNG completely in every engine.</summary>
[MemoryDiagnoser]
public class ImagePngDecodeBenchmarks {
    private byte[] _encoded = null!;

    [GlobalSetup]
    public void Setup() {
        _encoded = ImageBenchmarkCorpus.Logo.ReadBytes();
        ImageComparisonValidation.ValidatePngDecode(_encoded);
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
