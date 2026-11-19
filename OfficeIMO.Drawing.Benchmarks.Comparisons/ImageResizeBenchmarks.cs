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
#if NET10_0_OR_GREATER
    private byte[] _oneImageSource = null!;
    private global::OneImage.PixelResizer _oneImageResizer = null!;
#endif

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
#if NET10_0_OR_GREATER
        _oneImageSource = rgba;
        _oneImageResizer = new global::OneImage.PixelResizer();
        ImageComparisonValidation.ValidateOneImageResize(_officeImage, office, TargetWidth, TargetHeight);
#endif
    }

    [GlobalCleanup]
    public void Cleanup() {
        _skiaImage.Dispose();
        _magickImage.Dispose();
#if NET10_0_OR_GREATER
        _oneImageResizer.Dispose();
#endif
    }

    [Benchmark(Baseline = true)]
    public int OfficeIMO() => ImageComparisonAdapters.ResizeOfficeImo(_officeImage, TargetWidth, TargetHeight);

    [Benchmark]
    public int SkiaSharp() => ImageComparisonAdapters.ResizeSkia(_skiaImage, TargetWidth, TargetHeight);

    [Benchmark]
    public int MagickNET() => ImageComparisonAdapters.ResizeMagick(_magickImage, TargetWidth, TargetHeight);
#if NET10_0_OR_GREATER

    [Benchmark]
    [BenchmarkCategory("library:oneimage")]
    public int OneImage() => OneImageComparisonAdapters.ResizeRgba(
        _oneImageResizer, _oneImageSource, _officeImage.Width, _officeImage.Height, TargetWidth, TargetHeight).Length / 4;
#endif
}
