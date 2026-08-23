using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Drawing.Benchmarks.Comparisons;

/// <summary>Decodes one identical photographic JPEG into a managed RGBA buffer in every engine.</summary>
[MemoryDiagnoser]
public class ImageJpegDecodeBenchmarks {
    private byte[] _encoded = null!;

    [GlobalSetup]
    public void Setup() {
        _encoded = ImageBenchmarkCorpus.Photo.ReadBytes();
        ImageComparisonValidation.MeasureJpegDecodeError(_encoded, 2048, 1619);
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
