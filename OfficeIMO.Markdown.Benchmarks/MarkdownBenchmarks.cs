using BenchmarkDotNet.Attributes;
using Markdig;

namespace OfficeIMO.Markdown.Benchmarks;

[MemoryDiagnoser]
public class MarkdownParseBenchmarks {
    private MarkdownLoadOptions _commonMarkOptions = null!;
    private MarkdownLoadOptions _portableOptions = null!;
    private string _markdown = string.Empty;

    [ParamsSource(nameof(CorpusNames))]
    public string CorpusName { get; set; } = string.Empty;

    public IEnumerable<string> CorpusNames() => MarkdownBenchmarkCorpus.Names;

    [GlobalSetup]
    public void Setup() {
        _commonMarkOptions = MarkdownLoadOptions.CreateCommonMarkProfile();
        _portableOptions = MarkdownLoadOptions.CreatePortableProfile();
        _markdown = MarkdownBenchmarkCorpus.Get(CorpusName);
        MarkdownBenchmarkValidation.AssertCommonMarkEquivalent(
            CorpusName,
            _markdown,
            _commonMarkOptions,
            MarkdownBenchmarkValidation.CreateOfficeCommonMarkHtmlOptions());
    }

    [Benchmark(Baseline = true, Description = "OfficeIMO-Semantic")]
    public MarkdownDocument OfficeIMO_ParseSemantic_CommonMark() => MarkdownReader.ParseSemantic(_markdown, _commonMarkOptions);

    [Benchmark(Description = "OfficeIMO source-backed")]
    public MarkdownDocument OfficeIMO_Parse_SourceBacked_CommonMark() => MarkdownDocument.Parse(_markdown, _commonMarkOptions);

    [Benchmark(Description = "Markdig")]
    public Markdig.Syntax.MarkdownDocument Markdig_ParseSemanticComparison_CommonMark() => Markdig.Markdown.Parse(_markdown);

    [Benchmark]
    public MarkdownDocument OfficeIMO_Parse_Default() => MarkdownDocument.Parse(_markdown);

    [Benchmark]
    public MarkdownDocument OfficeIMO_Parse_Portable() => MarkdownDocument.Parse(_markdown, _portableOptions);

    [Benchmark]
    public MarkdownParseInfo OfficeIMO_Parse_WithSyntaxTree_CommonMark() => MarkdownReader.ParseWithSyntaxTree(_markdown, _commonMarkOptions);

    [Benchmark]
    public MarkdownParseInfo OfficeIMO_Parse_WithSyntaxTree_Default() => MarkdownReader.ParseWithSyntaxTree(_markdown);

    [Benchmark]
    public MarkdownParseInfo OfficeIMO_Parse_WithSyntaxTree_Portable() => MarkdownReader.ParseWithSyntaxTree(_markdown, _portableOptions);
}

[MemoryDiagnoser]
public class MarkdownHtmlBenchmarks {
    private static readonly MarkdownPipeline MarkdigCommonMarkPipeline = new MarkdownPipelineBuilder().Build();

    private MarkdownLoadOptions _commonMarkOptions = null!;
    private MarkdownLoadOptions _portableOptions = null!;
    private MarkdownHtmlExportOptions _commonMarkHtmlOptions = null!;
    private string _markdown = string.Empty;

    [ParamsSource(nameof(CorpusNames))]
    public string CorpusName { get; set; } = string.Empty;

    public IEnumerable<string> CorpusNames() => MarkdownBenchmarkCorpus.Names;

    [GlobalSetup]
    public void Setup() {
        _commonMarkOptions = MarkdownLoadOptions.CreateCommonMarkProfile();
        _portableOptions = MarkdownLoadOptions.CreatePortableProfile();
        _commonMarkHtmlOptions = MarkdownBenchmarkValidation.CreateOfficeCommonMarkHtmlOptions();
        _markdown = MarkdownBenchmarkCorpus.Get(CorpusName);
        MarkdownBenchmarkValidation.AssertCommonMarkEquivalent(CorpusName, _markdown, _commonMarkOptions, _commonMarkHtmlOptions);
    }

    [Benchmark(Baseline = true, Description = "OfficeIMO")]
    public string OfficeIMO_ToHtml_CommonMark() => MarkdownDocument.Parse(_markdown, _commonMarkOptions).ToHtml(_commonMarkHtmlOptions);

    [Benchmark(Description = "Markdig")]
    public string Markdig_ToHtml_CommonMark() => Markdig.Markdown.ToHtml(_markdown, MarkdigCommonMarkPipeline);

    [Benchmark]
    public string OfficeIMO_ToHtml_Default() => MarkdownDocument.Parse(_markdown).ToHtml();

    [Benchmark]
    public string OfficeIMO_ToHtml_Portable() => MarkdownDocument.Parse(_markdown, _portableOptions).ToHtml();
}

[MemoryDiagnoser]
public class MarkdownTransformBenchmarks {
    private MarkdownLoadOptions _baselineOptions = null!;
    private MarkdownLoadOptions _transformOptions = null!;
    private string _markdown = string.Empty;

    [ParamsSource(nameof(CorpusNames))]
    public string CorpusName { get; set; } = string.Empty;

    public IEnumerable<string> CorpusNames() => MarkdownBenchmarkCorpus.Names;

    [GlobalSetup]
    public void Setup() {
        _baselineOptions = MarkdownLoadOptions.CreateOfficeIMOProfile();
        _transformOptions = CreateTransformOptions();
        _markdown = MarkdownBenchmarkCorpus.Get(CorpusName);
    }

    [Benchmark(Baseline = true)]
    public MarkdownDocument OfficeIMO_Parse_OfficeProfile() => MarkdownDocument.Parse(_markdown, _baselineOptions);

    [Benchmark]
    public MarkdownDocument OfficeIMO_Parse_WithNormalizationTransforms() => MarkdownDocument.Parse(_markdown, _transformOptions);

    [Benchmark]
    public MarkdownParseInfo OfficeIMO_ParseWithSyntaxTreeAndDiagnostics_WithNormalizationTransforms() =>
        MarkdownReader.ParseWithSyntaxTreeAndDiagnostics(_markdown, _transformOptions);

    [Benchmark]
    public string OfficeIMO_ToMarkdown_AfterNormalizationTransforms() =>
        MarkdownDocument.Parse(_markdown, _transformOptions).ToMarkdown();

    private static MarkdownLoadOptions CreateTransformOptions() {
        var options = MarkdownLoadOptions.CreateOfficeIMOProfile();
        options.InputNormalization.NormalizeCompactHeadingBoundaries = true;
        options.InputNormalization.NormalizeHeadingListBoundaries = true;
        options.InputNormalization.NormalizeColonListBoundaries = true;
        options.InputNormalization.NormalizeCompactStrongLabelListBoundaries = true;
        options.InputNormalization.NormalizeStandaloneHashHeadingSeparators = true;
        options.InputNormalization.NormalizeTightArrowStrongBoundaries = true;
        options.InputNormalization.NormalizeBrokenStrongArrowLabels = true;
        options.InputNormalization.NormalizeWrappedSignalFlowStrongRuns = true;
        options.InputNormalization.NormalizeSignalFlowLabelSpacing = true;
        options.InputNormalization.NormalizeOrderedListMarkerSpacing = true;
        options.InputNormalization.NormalizeOrderedListParenMarkers = true;
        options.InputNormalization.NormalizeOrderedListCaretArtifacts = true;
        return options;
    }
}
