using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Drawing.Benchmarks;

/// <summary>Measures shared resize and encoded-image optimization workflows.</summary>
[MemoryDiagnoser]
public class ImageTransformBenchmarks {
    private OfficeRasterImage _photo = null!;
    private byte[] _photoBytes = null!;
    private OfficeImageOptimizationRequest _optimization = null!;
    private int _targetHeight;

    [Params(400, 800, 1600)]
    public int TargetWidth { get; set; } = 800;

    [GlobalSetup]
    public void Setup() {
        _photoBytes = ImageBenchmarkCorpus.Photo.ReadBytes();
        _photo = ImageBenchmarkCorpus.Decode(_photoBytes, "photo setup");
        _targetHeight = Math.Max(1, (int)Math.Round(_photo.Height * (TargetWidth / (double)_photo.Width)));
        _optimization = new OfficeImageOptimizationRequest(TargetWidth, TargetWidth) {
            PreserveAspectRatio = true,
            ResamplingMode = OfficeRasterResamplingMode.Bilinear,
            JpegQuality = 85,
            KeepOriginalWhenNotSmaller = false
        };

        OfficeRasterImage resized = ResizeBilinear();
        if (resized.Width != TargetWidth || resized.Height != _targetHeight) {
            throw new InvalidOperationException(
                $"Resize produced {resized.Width}x{resized.Height}; expected {TargetWidth}x{_targetHeight}.");
        }
        OfficeImageOptimizationResult optimized = OptimizeForPlacement();
        ImageBenchmarkCorpus.AssertIdentified(
            optimized.Bytes,
            OfficeImageFormat.Jpeg,
            TargetWidth,
            _targetHeight,
            nameof(OptimizeForPlacement));
    }

    [Benchmark(Baseline = true)]
    public OfficeRasterImage ResizeBilinear() =>
        OfficeRasterResampler.Resize(_photo, TargetWidth, _targetHeight, OfficeRasterResamplingMode.Bilinear);

    [Benchmark]
    public OfficeImageOptimizationResult OptimizeForPlacement() =>
        OfficeImageOptimizer.Optimize(_photoBytes, _optimization, "photo.jpg");
}
