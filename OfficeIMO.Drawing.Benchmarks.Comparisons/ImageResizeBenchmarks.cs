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

        ImageComparisonValidation.ValidateResize(_officeImage, _skiaImage, _magickImage, TargetWidth, TargetHeight);
#if NET10_0_OR_GREATER
        _oneImageSource = rgba;
        _oneImageResizer = new global::OneImage.PixelResizer();
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
    [BenchmarkCategory("library:OfficeIMO", "case:Default")]
    public int OfficeIMO() => ImageComparisonAdapters.ResizeOfficeImo(_officeImage, TargetWidth, TargetHeight);

    [Benchmark]
    [BenchmarkCategory("library:SkiaSharp", "case:Default")]
    public int SkiaSharp() => ImageComparisonAdapters.ResizeSkia(_skiaImage, TargetWidth, TargetHeight);

    [Benchmark]
    [BenchmarkCategory("library:Magick.NET", "case:Default")]
    public int MagickNET() => ImageComparisonAdapters.ResizeMagick(_magickImage, TargetWidth, TargetHeight);
#if NET10_0_OR_GREATER

    [Benchmark]
    [BenchmarkCategory("library:OneImage", "case:Default")]
    public int OneImage() => OneImageComparisonAdapters.ResizeRgba(
        _oneImageResizer, _oneImageSource, _officeImage.Width, _officeImage.Height, TargetWidth, TargetHeight).Length / 4;
#endif
}
