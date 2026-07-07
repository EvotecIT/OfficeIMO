using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using OfficeIMO.Markdown.Html;

namespace OfficeIMO.Markdown.Benchmarks;

[MemoryDiagnoser]
[SimpleJob(RuntimeMoniker.Net80)]
public class HtmlToMarkdownBenchmarks {
    private HtmlToMarkdownConverter _officeConverter = null!;
    private HtmlToMarkdownOptions _officeOptions = null!;
    private HtmlToMarkdownOptions _githubOptions = null!;
    private HtmlToMarkdownOptions _commonMarkOptions = null!;
    private ReverseMarkdown.Converter _reverseDefault = null!;
    private ReverseMarkdown.Converter _reverseGitHub = null!;
    private ReverseMarkdown.Converter _reverseCommonMark = null!;
    private string _html = string.Empty;

    [ParamsSource(nameof(CorpusNames))]
    public string CorpusName { get; set; } = string.Empty;

    public IEnumerable<string> CorpusNames() => HtmlToMarkdownBenchmarkCorpus.Names;

    [GlobalSetup]
    public void Setup() {
        _officeConverter = new HtmlToMarkdownConverter();
        _officeOptions = HtmlToMarkdownOptions.CreateOfficeIMOProfile();
        _githubOptions = HtmlToMarkdownOptions.CreateGitHubFlavoredMarkdownProfile();
        _commonMarkOptions = HtmlToMarkdownOptions.CreateCommonMarkProfile();
        _reverseDefault = new ReverseMarkdown.Converter();
        _reverseGitHub = new ReverseMarkdown.Converter(new ReverseMarkdown.Config {
            Flavor = ReverseMarkdown.Config.MarkdownFlavor.GitHub,
            Links = { SmartHref = true }
        });
        _reverseCommonMark = new ReverseMarkdown.Converter(new ReverseMarkdown.Config {
            Flavor = ReverseMarkdown.Config.MarkdownFlavor.CommonMark
        });
        _html = HtmlToMarkdownBenchmarkCorpus.Get(CorpusName);
        ValidateConverterOutputs(HtmlToMarkdownBenchmarkCorpus.GetExpectedFragment(CorpusName));
    }

    [Benchmark(Baseline = true)]
    public string OfficeIMO_GitHub_Profile() => _officeConverter.Convert(_html, _githubOptions);

    [Benchmark]
    public string OfficeIMO_CommonMark_Profile() => _officeConverter.Convert(_html, _commonMarkOptions);

    [Benchmark]
    public string OfficeIMO_Default_Profile() => _officeConverter.Convert(_html, _officeOptions);

    [Benchmark]
    public string ReverseMarkdown_GitHub_Profile() => _reverseGitHub.Convert(_html);

    [Benchmark]
    public string ReverseMarkdown_CommonMark_Profile() => _reverseCommonMark.Convert(_html);

    [Benchmark]
    public string ReverseMarkdown_Default_Profile() => _reverseDefault.Convert(_html);

    private void ValidateConverterOutputs(string expectedFragment) {
        ValidateOutput(nameof(OfficeIMO_GitHub_Profile), OfficeIMO_GitHub_Profile(), expectedFragment);
        ValidateOutput(nameof(OfficeIMO_CommonMark_Profile), OfficeIMO_CommonMark_Profile(), expectedFragment);
        ValidateOutput(nameof(OfficeIMO_Default_Profile), OfficeIMO_Default_Profile(), expectedFragment);
        ValidateOutput(nameof(ReverseMarkdown_GitHub_Profile), ReverseMarkdown_GitHub_Profile(), expectedFragment);
        ValidateOutput(nameof(ReverseMarkdown_CommonMark_Profile), ReverseMarkdown_CommonMark_Profile(), expectedFragment);
        ValidateOutput(nameof(ReverseMarkdown_Default_Profile), ReverseMarkdown_Default_Profile(), expectedFragment);
    }

    private static void ValidateOutput(string laneName, string markdown, string expectedFragment) {
        if (string.IsNullOrWhiteSpace(markdown) ||
            !markdown.Contains(expectedFragment, StringComparison.Ordinal)) {
            throw new InvalidOperationException(
                $"{laneName} did not preserve expected benchmark content '{expectedFragment}'.");
        }
    }
}
