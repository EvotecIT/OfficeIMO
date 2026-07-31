using Markdig;
using OfficeIMO.Html;
using OfficeIMO.Markdown.Html;

namespace OfficeIMO.Markdown.Benchmarks;

internal static class HtmlToMarkdownBenchmarkValidation {
    private static readonly MarkdownPipeline SemanticPipeline = new MarkdownPipelineBuilder()
        .UsePipeTables()
        .Build();

    internal static void AssertDefaultEquivalent(string corpusName, string officeMarkdown, string reverseMarkdown) {
        string expectedFragment = HtmlToMarkdownBenchmarkCorpus.GetExpectedFragment(corpusName);
        ValidateOutput("OfficeIMO default", officeMarkdown, expectedFragment);
        ValidateOutput("ReverseMarkdown default", reverseMarkdown, expectedFragment);

        string officeHtml = MarkdownBenchmarkValidation.NormalizeHtml(
            Markdig.Markdown.ToHtml(officeMarkdown, SemanticPipeline));
        string reverseHtml = MarkdownBenchmarkValidation.NormalizeHtml(
            Markdig.Markdown.ToHtml(reverseMarkdown, SemanticPipeline));
        if (string.Equals(officeHtml, reverseHtml, StringComparison.Ordinal)) {
            return;
        }

        throw new InvalidOperationException(
            $"HTML-to-Markdown output differs semantically for corpus '{corpusName}'. " +
            "Keep this corpus out of the competitive benchmark until both measured paths produce equivalent rendered HTML.");
    }

    internal static void AssertAllDefaultComparisons() {
        var officeOptions = HtmlToMarkdownOptions.CreateOfficeIMOProfile();
        var reverseConverter = new ReverseMarkdown.Converter();
        foreach (string corpusName in HtmlToMarkdownBenchmarkCorpus.ComparisonNames) {
            string html = HtmlToMarkdownBenchmarkCorpus.Get(corpusName);
            string officeMarkdown = HtmlConversionDocument.Parse(html).ToMarkdown(officeOptions);
            string reverseMarkdown = reverseConverter.Convert(html);
            AssertDefaultEquivalent(corpusName, officeMarkdown, reverseMarkdown);
        }
    }

    private static void ValidateOutput(string laneName, string markdown, string expectedFragment) {
        if (string.IsNullOrWhiteSpace(markdown) ||
            !markdown.Contains(expectedFragment, StringComparison.Ordinal)) {
            throw new InvalidOperationException(
                $"{laneName} did not preserve expected benchmark content '{expectedFragment}'.");
        }
    }
}
