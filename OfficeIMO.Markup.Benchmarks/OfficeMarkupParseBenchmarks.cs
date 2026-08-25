using BenchmarkDotNet.Attributes;
using Markdig.Syntax;

namespace OfficeIMO.Markup.Benchmarks;

[MemoryDiagnoser]
public class OfficeMarkupParseBenchmarks {
    private OfficeMarkupBenchmarkFixture _fixture = null!;

    [ParamsSource(nameof(CorpusScales))]
    public string Scale { get; set; } = string.Empty;

    public IEnumerable<string> CorpusScales() => OfficeMarkupBenchmarkCorpus.Scales;

    [GlobalSetup]
    public void Setup() {
        _fixture = OfficeMarkupBenchmarkCorpus.Get(Scale);
        OfficeMarkupBenchmarkValidation.Validate(_fixture);
    }

    [Benchmark(Baseline = true, Description = "OfficeIMO semantic parse")]
    public OfficeMarkupParseResult OfficeIMO() => OfficeMarkupParser.Parse(_fixture.Source, OfficeMarkupBenchmarkValidation.OfficeOptions);

    [Benchmark(Description = "Markdig semantic parse")]
    public MarkdownDocument MarkdigSemanticParse() =>
        global::Markdig.Markdown.Parse(_fixture.Source, OfficeMarkupBenchmarkValidation.MarkdigPipeline);
}
