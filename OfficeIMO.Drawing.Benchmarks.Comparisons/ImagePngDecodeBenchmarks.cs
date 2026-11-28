using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Drawing.Benchmarks.Comparisons;

/// <summary>Decodes one identical PNG into a managed RGBA buffer in every engine.</summary>
// No [MemoryDiagnoser]: SkiaSharp and Magick.NET allocate in native memory, which BenchmarkDotNet does not measure.
public class ImagePngDecodeBenchmarks {
    private byte[] _encoded = null!;

    [GlobalSetup]
    public void Setup() {
        _encoded = ImageBenchmarkCorpus.Logo.ReadBytes();
        ImageComparisonValidation.ValidatePngDecode(_encoded);
    }

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("library:OfficeIMO", "case:Default")]
    public byte[] OfficeIMO() => ImageComparisonAdapters.DecodeOfficeImoRgba(_encoded);

    [Benchmark]
    [BenchmarkCategory("library:SkiaSharp", "case:Default")]
    public byte[] SkiaSharp() => ImageComparisonAdapters.DecodeSkiaRgba(_encoded);

    [Benchmark]
    [BenchmarkCategory("library:Magick.NET", "case:Default")]
    public byte[] MagickNET() => ImageComparisonAdapters.DecodeMagickRgba(_encoded);

    [Benchmark]
    [BenchmarkCategory("library:StbImageSharp", "case:Default")]
    public byte[] StbImageSharp() => ImageComparisonAdapters.DecodeStbRgba(_encoded);
#if NET10_0_OR_GREATER

    [Benchmark]
    [BenchmarkCategory("library:OneImage", "case:Default")]
    public byte[] OneImage() => OneImageComparisonAdapters.DecodeRgba(_encoded);
#endif
}
