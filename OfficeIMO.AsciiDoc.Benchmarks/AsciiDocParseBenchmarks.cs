using BenchmarkDotNet.Attributes;

namespace OfficeIMO.AsciiDoc.Benchmarks;

[MemoryDiagnoser]
public class AsciiDocParseBenchmarks {
    private AsciiDocBenchmarkFixture _fixture = null!;

    [ParamsSource(nameof(CorpusScales))]
    public string Scale { get; set; } = string.Empty;

    public IEnumerable<string> CorpusScales() => AsciiDocBenchmarkCorpus.Scales;

    [GlobalSetup]
    public void Setup() {
        _fixture = AsciiDocBenchmarkCorpus.Get(Scale);
        AsciiDocBenchmarkValidation.Validate(_fixture);
    }

    [Benchmark(Baseline = true, Description = "Parse lossless")]
    public AsciiDocParseResult ParseLossless() => AsciiDocDocument.Parse(_fixture.Source);

    [Benchmark(Description = "Parse + preserve write")]
    public string ParseAndPreserveWrite() => AsciiDocDocument.Parse(_fixture.Source).Document.ToAsciiDoc();
}
