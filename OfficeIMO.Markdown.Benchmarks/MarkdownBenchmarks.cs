using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using Markdig;

namespace OfficeIMO.Markdown.Benchmarks;

[MemoryDiagnoser]
[SimpleJob(RuntimeMoniker.Net80)]
public class MarkdownParseBenchmarks {
    private MarkdownReaderOptions _commonMarkOptions = null!;
    private MarkdownReaderOptions _portableOptions = null!;
    private string _markdown = string.Empty;

    [ParamsSource(nameof(CorpusNames))]
    public string CorpusName { get; set; } = string.Empty;

    public IEnumerable<string> CorpusNames() => MarkdownBenchmarkCorpus.Names;

    [GlobalSetup]
    public void Setup() {
        _commonMarkOptions = MarkdownReaderOptions.CreateCommonMarkProfile();
        _portableOptions = MarkdownReaderOptions.CreatePortableProfile();
        _markdown = MarkdownBenchmarkCorpus.Get(CorpusName);
    }

    [Benchmark(Baseline = true)]
    public MarkdownDoc OfficeIMO_Parse_CommonMark() => MarkdownReader.Parse(_markdown, _commonMarkOptions);

    [Benchmark]
    public Markdig.Syntax.MarkdownDocument Markdig_Parse_CommonMark() => Markdig.Markdown.Parse(_markdown);

    [Benchmark]
    public MarkdownDoc OfficeIMO_Parse_Default() => MarkdownReader.Parse(_markdown);

    [Benchmark]
    public MarkdownDoc OfficeIMO_Parse_Portable() => MarkdownReader.Parse(_markdown, _portableOptions);

    [Benchmark]
    public MarkdownParseResult OfficeIMO_Parse_WithSyntaxTree_CommonMark() => MarkdownReader.ParseWithSyntaxTree(_markdown, _commonMarkOptions);

    [Benchmark]
    public MarkdownParseResult OfficeIMO_Parse_WithSyntaxTree_Default() => MarkdownReader.ParseWithSyntaxTree(_markdown);

    [Benchmark]
    public MarkdownParseResult OfficeIMO_Parse_WithSyntaxTree_Portable() => MarkdownReader.ParseWithSyntaxTree(_markdown, _portableOptions);
}

[MemoryDiagnoser]
[SimpleJob(RuntimeMoniker.Net80)]
public class MarkdownHtmlBenchmarks {
    private static readonly MarkdownPipeline MarkdigCommonMarkPipeline = new MarkdownPipelineBuilder().Build();

    private MarkdownReaderOptions _commonMarkOptions = null!;
    private MarkdownReaderOptions _portableOptions = null!;
    private string _markdown = string.Empty;

    [ParamsSource(nameof(CorpusNames))]
    public string CorpusName { get; set; } = string.Empty;

    public IEnumerable<string> CorpusNames() => MarkdownBenchmarkCorpus.Names;

    [GlobalSetup]
    public void Setup() {
        _commonMarkOptions = MarkdownReaderOptions.CreateCommonMarkProfile();
        _portableOptions = MarkdownReaderOptions.CreatePortableProfile();
        _markdown = MarkdownBenchmarkCorpus.Get(CorpusName);
    }

    [Benchmark(Baseline = true)]
    public string OfficeIMO_ToHtml_CommonMark() => MarkdownReader.Parse(_markdown, _commonMarkOptions).ToHtml();

    [Benchmark]
    public string Markdig_ToHtml_CommonMark() => Markdig.Markdown.ToHtml(_markdown, MarkdigCommonMarkPipeline);

    [Benchmark]
    public string OfficeIMO_ToHtml_Default() => MarkdownReader.Parse(_markdown).ToHtml();

    [Benchmark]
    public string OfficeIMO_ToHtml_Portable() => MarkdownReader.Parse(_markdown, _portableOptions).ToHtml();
}
