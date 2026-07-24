using System.Text;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class MarkdownAllSeverityBatch11SecurityTests {
    [Fact]
    public void MarkdownReader_DeeplyNestedImageAlt_DoesNotRecursivelyMaterializeImages() {
        var markdown = new StringBuilder();
        for (int i = 0; i < 512; i++) markdown.Append("![");
        markdown.Append("safe");
        for (int i = 0; i < 512; i++) markdown.Append("](https://example.com/image.png)");

        MarkdownDoc document = MarkdownReader.Parse(markdown.ToString());

        var image = Assert.IsType<ImageBlock>(Assert.Single(document.Blocks));
        Assert.Contains("safe", image.PlainAlt, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownVisualContract_RejectsOversizedOrMismatchedPayloadsBeforeDecoding() {
        string oversized = new('A', ((MarkdownVisualElementContract.MaxDecodedPayloadBytes + 2) / 3 * 4) + 1);
        Assert.True(MarkdownVisualElementContract.TryParse(CreateVisualAttributes(oversized), out MarkdownVisualElement? element));
        Assert.Null(MarkdownVisualElementContract.TryDecodePayload(element));

        var wrongVersion = CreateVisualAttributes(Convert.ToBase64String(Encoding.UTF8.GetBytes("{}")));
        wrongVersion[MarkdownVisualElementContract.AttributeVisualContract] = "v2";
        Assert.True(MarkdownVisualElementContract.TryParse(wrongVersion, out element));
        Assert.Null(MarkdownVisualElementContract.TryDecodePayload(element));
    }

    [Fact]
    public void MarkdownReader_DeepInlineHtmlWrappers_UseBoundedIndexedMatching() {
        var markdown = new StringBuilder();
        for (int i = 0; i < 512; i++) markdown.Append("<u>");
        markdown.Append("safe");
        for (int i = 0; i < 512; i++) markdown.Append("</u>");

        MarkdownDoc document = MarkdownReader.Parse(markdown.ToString());

        Assert.Single(document.Blocks);
        Assert.Contains("safe", document.ToMarkdown(), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("<u><sup>x</u></sup>")]
    [InlineData("<q><ins>x</q></ins>")]
    public void MarkdownReader_CrossedInlineHtmlWrappers_StayWithinTheCurrentSlice(string markdown) {
        MarkdownDoc document = MarkdownReader.Parse(markdown);

        Assert.Single(document.Blocks);
        Assert.Contains("x", document.ToMarkdown(), StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownReader_ManyUnmatchedInlineHtmlOpeners_AreHandledWithoutRepeatedSuffixScans() {
        string markdown = string.Join(" ", Enumerable.Repeat("<u>", 4_000));

        MarkdownDoc document = MarkdownReader.Parse(markdown);

        Assert.Single(document.Blocks);
    }

    [Fact]
    public void MarkdownReader_BlockImageAltWithManyUnmatchedInlineHtmlOpeners_UsesOneWrapperIndex() {
        string alt = string.Join(" ", Enumerable.Repeat("<u>", 4_000));
        string markdown = "![" + alt + "](https://example.com/image.png)";

        MarkdownDoc document = MarkdownReader.Parse(markdown);

        Assert.IsType<ImageBlock>(Assert.Single(document.Blocks));
    }

    [Fact]
    public void MarkdownReader_LongSetextHeading_IsParsedFromOneUnderlineScan() {
        string markdown = string.Join("\n", Enumerable.Repeat("heading", 2_000)) + "\n====";

        MarkdownDoc document = MarkdownReader.Parse(markdown);

        var heading = Assert.IsType<HeadingBlock>(Assert.Single(document.Blocks));
        Assert.Equal(1, heading.Level);
        Assert.StartsWith("heading", heading.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownReader_OddEmphasisCloserSelection_PreservesFormattingWithManyRuns() {
        string markdown = "***content* " + string.Join(" ", Enumerable.Repeat("***x*", 2_000));

        MarkdownDoc document = MarkdownReader.Parse(markdown);

        Assert.Single(document.Blocks);
        Assert.Contains("content", document.ToMarkdown(), StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownReader_RepeatedShortEmphasisOpeners_UseIndexedClosingRunLookups() {
        string markdown = string.Join(" ", Enumerable.Repeat("**x ", 4_000));

        MarkdownDoc document = MarkdownReader.Parse(markdown);

        Assert.Single(document.Blocks);
    }

    [Fact]
    public void MarkdownListRendering_LargeTaskAndLooseScopes_PreserveExpectedHtml() {
        string markdown = string.Join("\n", Enumerable.Range(0, 4_000).Select(index =>
            index % 2 == 0 ? $"- [ ] item {index}" : $"- item {index}"));
        var options = new MarkdownReaderOptions { TaskLists = true };

        string html = MarkdownReader.Parse(markdown, options).ToHtmlFragment();

        Assert.Contains("contains-task-list", html, StringComparison.Ordinal);
        Assert.Contains("item 3999", html, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_DeepInlineContainers_AvoidRecursiveBlockClassification() {
        var html = new StringBuilder();
        const int depthWithinUntrustedInputBudget = 240;
        for (int i = 0; i < depthWithinUntrustedInputBudget; i++) html.Append("<span>");
        html.Append("<div>safe</div>");
        for (int i = 0; i < depthWithinUntrustedInputBudget; i++) html.Append("</span>");

        string markdown = OfficeIMO.Html.HtmlConversionDocument.Parse(html.ToString()).ToMarkdown();

        Assert.Contains("safe", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_DefinitionListExpansion_EnforcesConfiguredAggregateBudget() {
        const string html = "<dl><dt>a</dt><dt>b</dt><dt>c</dt><dd>one</dd><dd>two</dd></dl>";
        var options = new HtmlToMarkdownOptions { MaxDefinitionListEntryExpansions = 5 };

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToMarkdownDocument(options));

        Assert.Contains("entry expansion limit", exception.Message, StringComparison.Ordinal);
    }

    private static Dictionary<string, string?> CreateVisualAttributes(string payload) => new(StringComparer.Ordinal) {
        [MarkdownVisualElementContract.AttributeVisualContract] = MarkdownVisualElementContract.ContractVersion,
        [MarkdownVisualElementContract.AttributeVisualKind] = "chart",
        [MarkdownVisualElementContract.AttributeFenceLanguage] = "chart",
        [MarkdownVisualElementContract.AttributeVisualHash] = "hash",
        [MarkdownVisualElementContract.AttributeConfigFormat] = MarkdownVisualElementContract.ConfigFormatJson,
        [MarkdownVisualElementContract.AttributeConfigEncoding] = MarkdownVisualElementContract.ConfigEncodingBase64Utf8,
        [MarkdownVisualElementContract.AttributeConfigBase64] = payload
    };
}
