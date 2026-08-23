using BenchmarkDotNet.Attributes;
using ImageMagick;
using SkiaSharp;

namespace OfficeIMO.Drawing.Benchmarks.Comparisons;

/// <summary>Resizes one identical opaque RGBA photo using linear/bilinear sampling.</summary>
[MemoryDiagnoser]
public class ImageResizeBenchmarks {
    private const int TargetWidth = 800;
    private const int TargetHeight = 632;
    private OfficeRasterImage _officeImage = null!;
    private SKBitmap _skiaImage = null!;
    private MagickImage _magickImage = null!;

    [GlobalSetup]
    public void Setup() {
        _officeImage = ImageBenchmarkCorpus.Decode(ImageBenchmarkCorpus.Photo.ReadBytes(), "resize source");
        byte[] rgba = _officeImage.GetPixels();
        _skiaImage = ImageComparisonAdapters.CreateSkiaBitmap(rgba, _officeImage.Width, _officeImage.Height);
        _magickImage = ImageComparisonAdapters.CreateMagickImage(rgba, _officeImage.Width, _officeImage.Height);

        byte[] office = OfficeRasterResampler.Resize(_officeImage, TargetWidth, TargetHeight).GetPixels();
        ImageComparisonValidation.AssertSimilarPixels(
            "SkiaSharp resize",
            office,
            ImageComparisonAdapters.ResizeSkiaRgba(_skiaImage, TargetWidth, TargetHeight),
            maximumMeanAbsoluteError: 3D);
        ImageComparisonValidation.AssertSimilarPixels(
            "Magick.NET resize",
            office,
            ImageComparisonAdapters.ResizeMagickRgba(_magickImage, TargetWidth, TargetHeight),
            maximumMeanAbsoluteError: 3D);
    }

    [GlobalCleanup]
    public void Cleanup() {
        _skiaImage.Dispose();
        _magickImage.Dispose();
    }

    [Benchmark(Baseline = true)]
    public int OfficeIMO() => ImageComparisonAdapters.ResizeOfficeImo(_officeImage, TargetWidth, TargetHeight);

    [Benchmark]
    public int SkiaSharp() => ImageComparisonAdapters.ResizeSkia(_skiaImage, TargetWidth, TargetHeight);

    [Benchmark]
    public int MagickNET() => ImageComparisonAdapters.ResizeMagick(_magickImage, TargetWidth, TargetHeight);
}
