using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using Markdig;

namespace OfficeIMO.Markdown.Benchmarks;

[MemoryDiagnoser]
[SimpleJob(RuntimeMoniker.Net80)]
public class MarkdownParseBenchmarks {
    private MarkdownReaderOptions _portableOptions = null!;
    private string _markdown = string.Empty;

    [ParamsSource(nameof(CorpusNames))]
    public string CorpusName { get; set; } = string.Empty;

    public IEnumerable<string> CorpusNames() => MarkdownBenchmarkCorpus.Names;

    [GlobalSetup]
    public void Setup() {
        _portableOptions = MarkdownReaderOptions.CreatePortableProfile();
        _markdown = MarkdownBenchmarkCorpus.Get(CorpusName);
    }

    [Benchmark(Baseline = true)]
    public MarkdownDoc OfficeIMO_Parse_Default() => MarkdownReader.Parse(_markdown);

    [Benchmark]
    public MarkdownDoc OfficeIMO_Parse_Portable() => MarkdownReader.Parse(_markdown, _portableOptions);

    [Benchmark]
    public Markdig.Syntax.MarkdownDocument Markdig_Parse() => Markdig.Markdown.Parse(_markdown);
}

[MemoryDiagnoser]
[SimpleJob(RuntimeMoniker.Net80)]
public class MarkdownHtmlBenchmarks {
    private static readonly MarkdownPipeline MarkdigPipeline = new MarkdownPipelineBuilder().Build();

    private MarkdownReaderOptions _portableOptions = null!;
    private string _markdown = string.Empty;

    [ParamsSource(nameof(CorpusNames))]
    public string CorpusName { get; set; } = string.Empty;

    public IEnumerable<string> CorpusNames() => MarkdownBenchmarkCorpus.Names;

    [GlobalSetup]
    public void Setup() {
        _portableOptions = MarkdownReaderOptions.CreatePortableProfile();
        _markdown = MarkdownBenchmarkCorpus.Get(CorpusName);
    }

    [Benchmark(Baseline = true)]
    public string OfficeIMO_ToHtml_Default() => MarkdownReader.Parse(_markdown).ToHtml();

    [Benchmark]
    public string OfficeIMO_ToHtml_Portable() => MarkdownReader.Parse(_markdown, _portableOptions).ToHtml();

    [Benchmark]
    public string Markdig_ToHtml() => Markdig.Markdown.ToHtml(_markdown, MarkdigPipeline);
}
