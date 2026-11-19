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
        // The measured parse results themselves, rendered with the same CommonMark HTML settings.
        MarkdownBenchmarkValidation.ValidateMeasuredLanes(
            CorpusName,
            "ParseSemantic_CommonMark",
            OfficeIMO_ParseSemantic_CommonMark().ToHtml(MarkdownBenchmarkValidation.CreateOfficeCommonMarkHtmlOptions()),
            Markdig_ParseSemanticComparison_CommonMark().ToHtml(MarkdigCommonMarkPipeline));
    }

    private static readonly MarkdownPipeline MarkdigCommonMarkPipeline = new MarkdownPipelineBuilder().Build();

    [Benchmark(Baseline = true, Description = "OfficeIMO-Semantic")]
    [BenchmarkCategory("library:OfficeIMO", "case:ParseSemantic_CommonMark")]
    public MarkdownDocument OfficeIMO_ParseSemantic_CommonMark() => MarkdownBenchmarkParsing.ParseSemantic(_markdown, _commonMarkOptions);

    [Benchmark(Description = "OfficeIMO source-backed")]
    public MarkdownDocument OfficeIMO_Parse_SourceBacked_CommonMark() => MarkdownDocument.Parse(_markdown, _commonMarkOptions);

    [Benchmark(Description = "Markdig")]
    [BenchmarkCategory("library:Markdig", "case:ParseSemantic_CommonMark")]
    public Markdig.Syntax.MarkdownDocument Markdig_ParseSemanticComparison_CommonMark() => Markdig.Markdown.Parse(_markdown);

    [Benchmark]
    public MarkdownDocument OfficeIMO_Parse_Default() => MarkdownDocument.Parse(_markdown);

    [Benchmark]
    public MarkdownDocument OfficeIMO_Parse_Portable() => MarkdownDocument.Parse(_markdown, _portableOptions);

    [Benchmark]
    public MarkdownParseInfo OfficeIMO_Parse_WithSyntaxTree_CommonMark() => MarkdownBenchmarkParsing.ParseWithSyntaxTree(_markdown, _commonMarkOptions);

    [Benchmark]
    public MarkdownParseInfo OfficeIMO_Parse_WithSyntaxTree_Default() => MarkdownBenchmarkParsing.ParseWithSyntaxTree(_markdown);

    [Benchmark]
    public MarkdownParseInfo OfficeIMO_Parse_WithSyntaxTree_Portable() => MarkdownBenchmarkParsing.ParseWithSyntaxTree(_markdown, _portableOptions);
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
        MarkdownBenchmarkValidation.ValidateMeasuredLanes(CorpusName, "ToHtml_CommonMark", OfficeIMO_ToHtml_CommonMark(), Markdig_ToHtml_CommonMark());
    }

    [Benchmark(Baseline = true, Description = "OfficeIMO source-backed")]
    [BenchmarkCategory("library:OfficeIMO", "case:ToHtml_CommonMark")]
    public string OfficeIMO_ToHtml_CommonMark() => MarkdownDocument.Parse(_markdown, _commonMarkOptions).ToHtml(_commonMarkHtmlOptions);

    [Benchmark(Description = "OfficeIMO-Semantic")]
    public string OfficeIMO_ToHtml_Semantic_CommonMark() =>
        MarkdownBenchmarkParsing.ParseSemantic(_markdown, _commonMarkOptions).ToHtml(_commonMarkHtmlOptions);

    [Benchmark(Description = "Markdig")]
    [BenchmarkCategory("library:Markdig", "case:ToHtml_CommonMark")]
    public string Markdig_ToHtml_CommonMark() => Markdig.Markdown.ToHtml(_markdown, MarkdigCommonMarkPipeline);

    [Benchmark(Description = "OfficeIMO default")]
    public string OfficeIMO_ToHtml_Default() => MarkdownDocument.Parse(_markdown).ToHtml();

    [Benchmark(Description = "OfficeIMO portable")]
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
        MarkdownBenchmarkParsing.ParseWithSyntaxTree(_markdown, _transformOptions);

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
