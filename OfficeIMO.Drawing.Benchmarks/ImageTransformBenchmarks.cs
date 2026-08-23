using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Drawing.Benchmarks;

/// <summary>Measures shared resize and encoded-image optimization workflows.</summary>
[MemoryDiagnoser]
public class ImageTransformBenchmarks {
    private OfficeRasterImage _photo = null!;
    private byte[] _photoBytes = null!;
    private OfficeImageOptimizationRequest _optimization = null!;

    [GlobalSetup]
    public void Setup() {
        _photoBytes = ImageBenchmarkCorpus.Photo.ReadBytes();
        _photo = ImageBenchmarkCorpus.Decode(_photoBytes, "photo setup");
        _optimization = new OfficeImageOptimizationRequest(800, 800) {
            PreserveAspectRatio = true,
            ResamplingMode = OfficeRasterResamplingMode.Bilinear,
            JpegQuality = 85,
            KeepOriginalWhenNotSmaller = false
        };

        OfficeRasterImage resized = ResizeBilinear();
        if (resized.Width != 800 || resized.Height != 632) {
            throw new InvalidOperationException($"Resize produced {resized.Width}x{resized.Height}; expected 800x632.");
        }
        OfficeImageOptimizationResult optimized = OptimizeForPlacement();
        ImageBenchmarkCorpus.AssertIdentified(optimized.Bytes, OfficeImageFormat.Jpeg, 800, 632, nameof(OptimizeForPlacement));
    }

    [Benchmark(Baseline = true)]
    public OfficeRasterImage ResizeBilinear() =>
        OfficeRasterResampler.Resize(_photo, 800, 632, OfficeRasterResamplingMode.Bilinear);

    [Benchmark]
    public OfficeImageOptimizationResult OptimizeForPlacement() =>
        OfficeImageOptimizer.Optimize(_photoBytes, _optimization, "photo.jpg");
}
