using System.Linq;
using System.Text.Json;
using OfficeIMO.Adf;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Adf.Tests;

public sealed class AdfContractTests {
    [Fact]
    public void ParseWrite_PreservesUnknownNodesMarksAttributesAndProperties() {
        const string json = "{\"version\":1,\"type\":\"doc\",\"vendorRoot\":{\"enabled\":true},\"content\":[{\"type\":\"vendorPanel\",\"attrs\":{\"tone\":\"blue\"},\"vendorNode\":17,\"content\":[{\"type\":\"text\",\"text\":\"Hello\",\"marks\":[{\"type\":\"vendorMark\",\"attrs\":{\"x\":1},\"vendorMarkProperty\":\"kept\"}]}]}]}";

        AdfDocument document = AdfDocument.Parse(json);
        string output = document.ToJson();

        using JsonDocument parsed = JsonDocument.Parse(output);
        JsonElement root = parsed.RootElement;
        Assert.True(root.GetProperty("vendorRoot").GetProperty("enabled").GetBoolean());
        JsonElement node = root.GetProperty("content")[0];
        Assert.Equal("vendorPanel", node.GetProperty("type").GetString());
        Assert.Equal(17, node.GetProperty("vendorNode").GetInt32());
        Assert.Equal("blue", node.GetProperty("attrs").GetProperty("tone").GetString());
        Assert.Equal("kept", node.GetProperty("content")[0].GetProperty("marks")[0].GetProperty("vendorMarkProperty").GetString());
    }

    [Fact]
    public void MarkdownRoundTrip_PreservesCommonDocumentStructure() {
        const string markdown = "# Status\n\nThis is **ready** with [details](https://example.com).\n\n- First\n- Second\n\n```powershell\nGet-Date\n```";

        AdfConversionResult<AdfDocument> adf = AdfConverter.FromMarkdown(markdown);
        AdfConversionResult<string> roundTrip = AdfConverter.ToMarkdown(adf.Value);

        Assert.Contains("# Status", roundTrip.Value);
        Assert.Contains("**ready**", roundTrip.Value);
        Assert.Contains("[details](https://example.com)", roundTrip.Value);
        Assert.Contains("- First", roundTrip.Value);
        Assert.Contains("```powershell", roundTrip.Value);
        Assert.False(adf.Report.HasErrors);
    }

    [Fact]
    public void HtmlConversion_UsesOfficeimoPipelines() {
        AdfConversionResult<AdfDocument> adf = AdfConverter.FromHtml("<h2>Report</h2><p><strong>Ready</strong></p>");
        AdfConversionResult<string> html = AdfConverter.ToHtml(adf.Value);

        Assert.Contains("<h2", html.Value);
        Assert.Contains("Report", html.Value);
        Assert.Contains("<strong>Ready</strong>", html.Value);
        Assert.Contains(adf.Report.Diagnostics, item => item.Code == "ADF_HTML_VIA_MARKDOWN");
        Assert.Contains(html.Report.Diagnostics, item => item.Code == "ADF_TO_HTML_VIA_MARKDOWN");
    }

    [Fact]
    public void Validation_WarnsForUnknownNodesWithoutRejectingRoundTrip() {
        AdfDocument document = AdfDocument.Parse("{\"version\":1,\"type\":\"doc\",\"content\":[{\"type\":\"futureNode\",\"attrs\":{\"a\":1}}]}");

        AdfValidationResult result = document.Validate();

        Assert.True(result.IsValid);
        Assert.Contains(result.Issues, item => item.Code == "ADF_UNKNOWN_NODE");
    }

    [Fact]
    public void MarkdownTaskList_UsesValidListItemFallbackAndReportsFidelity() {
        AdfConversionResult<AdfDocument> result = AdfConverter.FromMarkdown("- [x] Ready\n- [ ] Pending");

        AdfNode list = Assert.Single(result.Value.Content);
        Assert.Equal("bulletList", list.Type);
        Assert.All(list.Content, item => Assert.Equal("listItem", item.Type));
        Assert.True(result.Value.Validate().IsValid);
        Assert.Contains(result.Report.Diagnostics, item => item.Code == "MARKDOWN_TASK_LIST_FALLBACK");
        Assert.Contains("\\[x\\] Ready", AdfConverter.ToMarkdown(result.Value).Value);
    }

    [Fact]
    public void Validation_RejectsTaskItemUnderBulletList() {
        AdfDocument document = AdfDocument.Parse("{\"version\":1,\"type\":\"doc\",\"content\":[{\"type\":\"bulletList\",\"content\":[{\"type\":\"taskItem\",\"attrs\":{\"state\":\"DONE\"},\"content\":[{\"type\":\"paragraph\",\"content\":[{\"type\":\"text\",\"text\":\"Ready\"}]}]}]}]}");

        AdfValidationResult result = document.Validate();

        Assert.False(result.IsValid);
        Assert.Contains(result.Issues, item => item.Code == "ADF_LIST_CHILD");
        Assert.Contains(result.Issues, item => item.Code == "ADF_TASK_ITEM_PARENT");
    }

    [Fact]
    public void LinkProjection_PreservesTitleAndReportsUnsupportedAttributes() {
        var link = new AdfMark("link")
            .SetAttribute("href", "https://example.com/details")
            .SetAttribute("title", "Ready \"now\"")
            .SetAttribute("collection", "other");
        var document = new AdfDocument();
        document.Content.Add(new AdfNode("paragraph") { Content = { AdfNode.TextNode("details", new[] { link }) } });

        AdfConversionResult<string> result = AdfConverter.ToMarkdown(document);

        Assert.Contains("[details](https://example.com/details 'Ready \"now\"')", result.Value);
        Assert.Contains(result.Report.Diagnostics, item => item.Code == "ADF_LINK_ATTRIBUTES_DROPPED");
    }

    [Fact]
    public void CodeBlockProjection_UsesFenceThatCannotCollideWithContent() {
        const string content = "before\n```\n# still code\nafter";
        var document = new AdfDocument();
        document.Content.Add(new AdfNode("codeBlock") { Content = { AdfNode.TextNode(content) } });

        AdfConversionResult<string> markdown = AdfConverter.ToMarkdown(document);
        AdfConversionResult<AdfDocument> roundTrip = AdfConverter.FromMarkdown(markdown.Value);

        Assert.StartsWith(MarkdownFence.BuildSafeFence(content), markdown.Value);
        AdfNode codeBlock = Assert.Single(roundTrip.Value.Content);
        Assert.Equal("codeBlock", codeBlock.Type);
        Assert.Equal(content, Assert.Single(codeBlock.Content).Text);
    }

    [Theory]
    [InlineData("# Heading")]
    [InlineData("> quote")]
    [InlineData("- item")]
    [InlineData("1. item")]
    public void ParagraphProjection_EscapesLineLeadingBlockSyntax(string text) {
        var document = new AdfDocument();
        document.Content.Add(new AdfNode("paragraph") { Content = { AdfNode.TextNode(text) } });

        AdfConversionResult<string> markdown = AdfConverter.ToMarkdown(document);
        AdfConversionResult<AdfDocument> roundTrip = AdfConverter.FromMarkdown(markdown.Value);

        AdfNode paragraph = Assert.Single(roundTrip.Value.Content);
        Assert.Equal("paragraph", paragraph.Type);
        Assert.Equal(text, string.Concat(paragraph.Content.Select(node => node.Text)));
    }

    [Fact]
    public void InlineCodeProjection_PreservesBackticksAndCombinedMarksRegardlessOfOrder() {
        const string content = "value`tick";
        foreach (AdfMark[] marks in new[] {
                     new[] { new AdfMark("strong"), new AdfMark("code") },
                     new[] { new AdfMark("code"), new AdfMark("strong") },
                 }) {
            var document = new AdfDocument();
            document.Content.Add(new AdfNode("paragraph") { Content = { AdfNode.TextNode(content, marks) } });

            AdfConversionResult<string> markdown = AdfConverter.ToMarkdown(document);
            AdfConversionResult<AdfDocument> roundTrip = AdfConverter.FromMarkdown(markdown.Value);

            AdfNode text = Assert.Single(Assert.Single(roundTrip.Value.Content).Content);
            Assert.Equal(content, text.Text);
            Assert.Contains(text.Marks, mark => mark.Type == "strong");
            Assert.Contains(text.Marks, mark => mark.Type == "code");
        }
    }
}
