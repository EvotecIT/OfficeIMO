using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Visio.Benchmarks;

[MemoryDiagnoser]
public class VisioBenchmarks {
    private VisioBenchmarkScale _scale = null!;
    private VisioBenchmarkFixture _fixture = null!;

    [Params("Small", "Normal", "Large")]
    public string Scale { get; set; } = string.Empty;

    [GlobalSetup]
    public void Setup() {
        _scale = VisioBenchmarkCorpus.Scales.Single(scale => scale.Name == Scale);
        _fixture = VisioBenchmarkCorpus.CreateFixture(_scale);
        VisioBenchmarkValidation.ValidatePackage(_fixture);
    }

    [Benchmark(Description = "Create and save VSDX")]
    public byte[] CreateAndSave() => VisioBenchmarkCorpus.CreateAndSave(_scale);

    [Benchmark(Description = "Load and inspect VSDX")]
    public VisioInspectionSnapshot LoadAndInspect() => VisioBenchmarkValidation.LoadAndInspect(_fixture);
}
