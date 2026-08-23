using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Drawing.Benchmarks.Comparisons;

/// <summary>Decodes one identical PNG into a managed RGBA buffer in every engine.</summary>
[MemoryDiagnoser]
public class ImagePngDecodeBenchmarks {
    private byte[] _encoded = null!;

    [GlobalSetup]
    public void Setup() {
        _encoded = ImageBenchmarkCorpus.Logo.ReadBytes();
        ImageComparisonValidation.ValidatePngDecode(_encoded);
    }

    [Benchmark(Baseline = true)]
    public byte[] OfficeIMO() => ImageComparisonAdapters.DecodeOfficeImoRgba(_encoded);

    [Benchmark]
    public byte[] SkiaSharp() => ImageComparisonAdapters.DecodeSkiaRgba(_encoded);

    [Benchmark]
    public byte[] MagickNET() => ImageComparisonAdapters.DecodeMagickRgba(_encoded);

    [Benchmark]
    public byte[] StbImageSharp() => ImageComparisonAdapters.DecodeStbRgba(_encoded);
}
