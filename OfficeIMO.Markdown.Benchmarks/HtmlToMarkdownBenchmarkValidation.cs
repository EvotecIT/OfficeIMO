using Markdig;
using OfficeIMO.Benchmarks;
using OfficeIMO.Html;
using OfficeIMO.Markdown.Html;

namespace OfficeIMO.Markdown.Benchmarks;

internal static class HtmlToMarkdownBenchmarkValidation {
    private static readonly MarkdownPipeline SemanticPipeline = new MarkdownPipelineBuilder()
        .UsePipeTables()
        .Build();

    internal static void AssertDefaultEquivalent(string corpusName, string officeMarkdown, string reverseMarkdown) {
        AssertDefaultEquivalent(corpusName, officeMarkdown, reverseMarkdown, recordValidation: false);
    }

    /// <summary>
    /// Checks both measured Markdown outputs: each keeps the corpus' expected content, and both render (through
    /// Markdig) to the same normalized HTML. Records the result for both libraries when <paramref name="recordValidation"/> is set.
    /// </summary>
    internal static void AssertDefaultEquivalent(string corpusName, string officeMarkdown, string reverseMarkdown, bool recordValidation) {
        const string caseName = "Default_Profile";
        string expectedFragment = HtmlToMarkdownBenchmarkCorpus.GetExpectedFragment(corpusName);
        string? failure = null;
        if (!PreservesFragment(officeMarkdown, expectedFragment)) {
            failure = $"OfficeIMO default did not preserve expected benchmark content '{expectedFragment}'.";
        } else if (!PreservesFragment(reverseMarkdown, expectedFragment)) {
            failure = $"ReverseMarkdown default did not preserve expected benchmark content '{expectedFragment}'.";
        }

        string officeHtml = MarkdownBenchmarkValidation.NormalizeHtml(
            Markdig.Markdown.ToHtml(officeMarkdown, SemanticPipeline));
        string reverseHtml = MarkdownBenchmarkValidation.NormalizeHtml(
            Markdig.Markdown.ToHtml(reverseMarkdown, SemanticPipeline));
        if (failure is null && !string.Equals(officeHtml, reverseHtml, StringComparison.Ordinal)) {
            failure = $"HTML-to-Markdown output differs semantically for corpus '{corpusName}'. " +
                "Keep this corpus out of the competitive benchmark until both measured paths produce equivalent rendered HTML.";
        }

        string detail = $"CorpusName={corpusName}; expected fragment '{expectedFragment}'; rendered HTML characters OfficeIMO={officeHtml.Length}, ReverseMarkdown={reverseHtml.Length}";
        if (failure is null) {
            if (recordValidation) {
                BenchmarkOutputValidationLog.Passed("OfficeIMO", detail + "; identical rendered HTML", caseName);
                BenchmarkOutputValidationLog.Passed("ReverseMarkdown", detail + "; identical rendered HTML", caseName);
            }
            return;
        }

        if (recordValidation) {
            BenchmarkOutputValidationLog.Failed("OfficeIMO", detail + "; " + failure, caseName);
            BenchmarkOutputValidationLog.Failed("ReverseMarkdown", detail + "; " + failure, caseName);
        }
        throw new InvalidOperationException(failure);
    }

    private static bool PreservesFragment(string markdown, string expectedFragment) =>
        !string.IsNullOrWhiteSpace(markdown) && markdown.Contains(expectedFragment, StringComparison.Ordinal);

    internal static void AssertAllDefaultComparisons() {
        var officeOptions = HtmlMarkdownExportOptions.CreateOfficeIMOProfile();
        var reverseConverter = new ReverseMarkdown.Converter();
        foreach (string corpusName in HtmlToMarkdownBenchmarkCorpus.ComparisonNames) {
            string html = HtmlToMarkdownBenchmarkCorpus.Get(corpusName);
            string officeMarkdown = HtmlConversionDocument.Parse(html).ToMarkdown(officeOptions).ToMarkdown();
            string reverseMarkdown = reverseConverter.Convert(html);
            AssertDefaultEquivalent(corpusName, officeMarkdown, reverseMarkdown);
        }
    }
}
