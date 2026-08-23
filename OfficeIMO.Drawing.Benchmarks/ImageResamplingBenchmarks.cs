using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Drawing.Benchmarks;

/// <summary>Compares alpha-correct resize modes on representative visual content.</summary>
[MemoryDiagnoser]
public class ImageResamplingBenchmarks {
    private OfficeRasterImage _source = null!;
    private int _width;
    private int _height;

    [ParamsSource(nameof(ScenarioIds))]
    public string ScenarioId { get; set; } = ImageBenchmarkScenarios.Photo.Id;

    public IEnumerable<string> ScenarioIds => ImageBenchmarkScenarios.ResamplingIds;

    [GlobalSetup]
    public void Setup() {
        ImageBenchmarkScenario scenario = ImageBenchmarkScenarios.Get(ScenarioId);
        _source = scenario.CreateImage();
        _width = Math.Max(1, _source.Width / 4);
        _height = Math.Max(1, _source.Height / 4);
        Validate(ResizeBilinear(), nameof(ResizeBilinear));
        Validate(ResizeArea(), nameof(ResizeArea));
        Validate(ResizeLanczos3(), nameof(ResizeLanczos3));
    }

    [Benchmark(Baseline = true)]
    public OfficeRasterImage ResizeBilinear() =>
        OfficeRasterResampler.Resize(_source, _width, _height, OfficeRasterResamplingMode.Bilinear);

    [Benchmark]
    public OfficeRasterImage ResizeArea() =>
        OfficeRasterResampler.Resize(_source, _width, _height, OfficeRasterResamplingMode.Area);

    [Benchmark]
    public OfficeRasterImage ResizeLanczos3() =>
        OfficeRasterResampler.Resize(_source, _width, _height, OfficeRasterResamplingMode.Lanczos3);

    private void Validate(OfficeRasterImage image, string method) {
        if (image.Width != _width || image.Height != _height) {
            throw new InvalidOperationException(
                $"{ScenarioId} {method} produced {image.Width}x{image.Height}; expected {_width}x{_height}.");
        }
    }
}
