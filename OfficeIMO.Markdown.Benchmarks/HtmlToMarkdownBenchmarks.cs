using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using OfficeIMO.Html;
using OfficeIMO.Markdown.Html;

namespace OfficeIMO.Markdown.Benchmarks;

[MemoryDiagnoser]
[SimpleJob(RuntimeMoniker.Net80)]
public class HtmlToMarkdownBenchmarks {
    private HtmlToMarkdownOptions _officeOptions = null!;
    private ReverseMarkdown.Converter _reverseDefault = null!;
    private string _html = string.Empty;

    [ParamsSource(nameof(CorpusNames))]
    public string CorpusName { get; set; } = string.Empty;

    public IEnumerable<string> CorpusNames() => HtmlToMarkdownBenchmarkCorpus.ComparisonNames;

    [GlobalSetup]
    public void Setup() {
        _officeOptions = HtmlToMarkdownOptions.CreateOfficeIMOProfile();
        _reverseDefault = new ReverseMarkdown.Converter();
        _html = HtmlToMarkdownBenchmarkCorpus.Get(CorpusName);
        HtmlToMarkdownBenchmarkValidation.AssertDefaultEquivalent(
            CorpusName,
            OfficeIMO_Default_Profile(),
            ReverseMarkdown_Default_Profile());
    }

    [Benchmark(Baseline = true)]
    public string OfficeIMO_Default_Profile() => HtmlConversionDocument.Parse(_html).ToMarkdown(_officeOptions);

    [Benchmark]
    public string ReverseMarkdown_Default_Profile() => _reverseDefault.Convert(_html);
}

[MemoryDiagnoser]
[SimpleJob(RuntimeMoniker.Net80)]
public class HtmlToMarkdownOfficeProfileBenchmarks {
    private HtmlConversionDocument _document = null!;
    private HtmlToMarkdownOptions _officeOptions = null!;
    private HtmlToMarkdownOptions _githubOptions = null!;
    private HtmlToMarkdownOptions _commonMarkOptions = null!;

    [ParamsSource(nameof(CorpusNames))]
    public string CorpusName { get; set; } = string.Empty;

    public IEnumerable<string> CorpusNames() => HtmlToMarkdownBenchmarkCorpus.Names;

    [GlobalSetup]
    public void Setup() {
        _officeOptions = HtmlToMarkdownOptions.CreateOfficeIMOProfile();
        _githubOptions = HtmlToMarkdownOptions.CreateGitHubFlavoredMarkdownProfile();
        _commonMarkOptions = HtmlToMarkdownOptions.CreateCommonMarkProfile();
        _document = HtmlConversionDocument.Parse(HtmlToMarkdownBenchmarkCorpus.Get(CorpusName));
        ValidateConverterOutputs(HtmlToMarkdownBenchmarkCorpus.GetExpectedFragment(CorpusName));
    }

    [Benchmark]
    public string OfficeIMO_GitHub_Profile() => _document.ToMarkdown(_githubOptions);

    [Benchmark]
    public string OfficeIMO_CommonMark_Profile() => _document.ToMarkdown(_commonMarkOptions);

    [Benchmark]
    public string OfficeIMO_Default_Profile() => _document.ToMarkdown(_officeOptions);

    private void ValidateConverterOutputs(string expectedFragment) {
        ValidateOutput(nameof(OfficeIMO_GitHub_Profile), OfficeIMO_GitHub_Profile(), expectedFragment);
        ValidateOutput(nameof(OfficeIMO_CommonMark_Profile), OfficeIMO_CommonMark_Profile(), expectedFragment);
        ValidateOutput(nameof(OfficeIMO_Default_Profile), OfficeIMO_Default_Profile(), expectedFragment);
    }

    private static void ValidateOutput(string laneName, string markdown, string expectedFragment) {
        if (string.IsNullOrWhiteSpace(markdown) ||
            !markdown.Contains(expectedFragment, StringComparison.Ordinal)) {
            throw new InvalidOperationException(
                $"{laneName} did not preserve expected benchmark content '{expectedFragment}'.");
        }
    }
}
