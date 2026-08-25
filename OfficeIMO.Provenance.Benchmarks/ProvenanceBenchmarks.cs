using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Provenance.Benchmarks;

[MemoryDiagnoser]
public class ProvenanceBenchmarks {
    private ProvenanceBenchmarkFixture _fixture = null!;

    [ParamsSource(nameof(CorpusFormats))]
    public string Format { get; set; } = string.Empty;

    [ParamsSource(nameof(CorpusScales))]
    public string Scale { get; set; } = string.Empty;

    public IEnumerable<string> CorpusFormats() => ProvenanceBenchmarkCorpus.Formats;
    public IEnumerable<string> CorpusScales() => ProvenanceBenchmarkCorpus.Scales;

    [GlobalSetup]
    public void Setup() {
        _fixture = ProvenanceBenchmarkCorpus.Create(Format, Scale);
        ProvenanceBenchmarkValidation.Validate(_fixture);
    }

    [Benchmark(Description = "Inspect provenance")]
    public OfficeProvenanceReport Inspect() => ProvenanceBenchmarkValidation.Inspect(_fixture);

    [Benchmark(Description = "Remove provenance")]
    public OfficeProvenanceRemovalResult Remove() => ProvenanceBenchmarkValidation.Remove(_fixture);
}
