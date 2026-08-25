using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Latex.Benchmarks;

/// <summary>Measures validated lossless LaTeX parsing and preserve writing at representative scales.</summary>
[MemoryDiagnoser]
public class LatexParseBenchmarks {
    private LatexBenchmarkFixture _fixture = null!;

    /// <summary>Gets or sets the deterministic document scale.</summary>
    [ParamsSource(nameof(CorpusScales))]
    public string Scale { get; set; } = string.Empty;

    /// <summary>Returns the scales shared by both workflows.</summary>
    public IEnumerable<string> CorpusScales() => LatexBenchmarkCorpus.Scales;

    /// <summary>Builds and validates the selected corpus before measurement begins.</summary>
    [GlobalSetup]
    public void Setup() {
        _fixture = LatexBenchmarkCorpus.Get(Scale);
        LatexBenchmarkValidation.Validate(_fixture);
    }

    /// <summary>Parses the complete source into the lossless syntax and semantic models.</summary>
    [Benchmark(Baseline = true, Description = "Parse lossless")]
    public LatexParseResult ParseLossless() => LatexDocument.Parse(_fixture.Source);

    /// <summary>Parses and writes the complete source through preserve mode.</summary>
    [Benchmark(Description = "Parse + preserve write")]
    public string ParseAndPreserveWrite() => LatexDocument.Parse(_fixture.Source).Document.ToLatex();
}
