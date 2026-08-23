using BenchmarkDotNet.Attributes;
using ImageMagick;
using SkiaSharp;

namespace OfficeIMO.Drawing.Benchmarks.Comparisons;

/// <summary>Encodes one identical 512x512 RGBA buffer as lossless PNG.</summary>
[MemoryDiagnoser]
public class ImagePngEncodeBenchmarks {
    private OfficeRasterImage _officeImage = null!;
    private SKBitmap _skiaImage = null!;
    private MagickImage _magickImage = null!;

    [GlobalSetup]
    public void Setup() {
        _officeImage = ImageBenchmarkCorpus.CreatePattern();
        byte[] rgba = _officeImage.GetPixels();
        _skiaImage = ImageComparisonAdapters.CreateSkiaBitmap(rgba, _officeImage.Width, _officeImage.Height);
        _magickImage = ImageComparisonAdapters.CreateMagickImage(rgba, _officeImage.Width, _officeImage.Height);
        ImageComparisonValidation.ValidatePngEncode(_officeImage, OfficeIMO(), SkiaSharp(), MagickNET());
    }

    [GlobalCleanup]
    public void Cleanup() {
        _skiaImage.Dispose();
        _magickImage.Dispose();
    }

    [Benchmark(Baseline = true)]
    public byte[] OfficeIMO() => ImageComparisonAdapters.EncodeOfficeImo(_officeImage);

    [Benchmark]
    public byte[] SkiaSharp() => ImageComparisonAdapters.EncodeSkia(_skiaImage);

    [Benchmark]
    public byte[] MagickNET() => ImageComparisonAdapters.EncodeMagick(_magickImage);
}
