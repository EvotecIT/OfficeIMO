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

    [Theory]
    [InlineData("text")]
    [InlineData("hardBreak")]
    [InlineData("mention")]
    [InlineData("tableCell")]
    public void Validation_RejectsKnownNonBlockNodesAtDocumentRoot(string nodeType) {
        var document = new AdfDocument();
        document.Content.Add(nodeType == "text" ? AdfNode.TextNode("invalid") : new AdfNode(nodeType));

        AdfValidationResult result = document.Validate();

        Assert.False(result.IsValid);
        AdfValidationIssue issue = Assert.Single(result.Issues, item => item.Code == "ADF_ROOT_CHILD");
        Assert.Equal("$.content[0]", issue.Path);
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
    public void LinkProjection_EscapesAndRoundTripsDestinationDelimiters() {
        const string href = "https://example.test/a(b)/docs\\[one]|two";
        var link = new AdfMark("link").SetAttribute("href", href);
        var document = new AdfDocument();
        document.Content.Add(new AdfNode("paragraph") {
            Content = { AdfNode.TextNode("details", new[] { link }) }
        });

        AdfConversionResult<string> markdown = AdfConverter.ToMarkdown(document);
        AdfConversionResult<AdfDocument> roundTrip = AdfConverter.FromMarkdown(markdown.Value);

        AdfNode text = Assert.Single(Assert.Single(roundTrip.Value.Content).Content);
        AdfMark roundTripLink = Assert.Single(text.Marks, mark => mark.Type == "link");
        Assert.Equal(href, roundTripLink.GetStringAttribute("href"));
    }

    [Fact]
    public void MarkdownDocumentSoftBreak_StaysTextInsteadOfBecomingHardBreak() {
        var markdown = MarkdownDoc.Create()
            .Add(new ParagraphBlock(new InlineSequence().Text("first").SoftBreak().Text("second")));

        AdfConversionResult<AdfDocument> adf = AdfConverter.FromMarkdown(markdown);
        AdfConversionResult<string> roundTrip = AdfConverter.ToMarkdown(adf.Value);

        AdfNode paragraph = Assert.Single(adf.Value.Content);
        Assert.Collection(
            paragraph.Content,
            first => Assert.Equal("first", first.Text),
            newline => Assert.Equal("\n", newline.Text),
            second => Assert.Equal("second", second.Text));
        Assert.True(adf.Value.Validate().IsValid);
        Assert.DoesNotContain(paragraph.Content, node => node.Type == "hardBreak");
        Assert.Equal("first\nsecond", roundTrip.Value.Replace("\r\n", "\n"));
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

    [Fact]
    public void CodeBlockProjection_NormalizesExternalLanguageToOneSafeToken() {
        const string content = "Get-Date";
        var document = new AdfDocument();
        document.Content.Add(new AdfNode("codeBlock") {
            Content = { AdfNode.TextNode(content) }
        }.SetAttribute("language", "powershell\n```\n# Heading"));

        AdfConversionResult<string> markdown = AdfConverter.ToMarkdown(document);
        AdfConversionResult<AdfDocument> roundTrip = AdfConverter.FromMarkdown(markdown.Value);

        Assert.StartsWith("```powershell\n", markdown.Value.Replace("\r\n", "\n"));
        Assert.Contains(markdown.Report.Diagnostics, item => item.Code == "ADF_CODE_LANGUAGE_NORMALIZED");
        AdfNode codeBlock = Assert.Single(roundTrip.Value.Content);
        Assert.Equal("codeBlock", codeBlock.Type);
        Assert.Equal("powershell", codeBlock.GetStringAttribute("language"));
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
    public void ParagraphProjection_PreservesLiteralHtmlLikeAndEntityLikeText() {
        const string content = "before <u>value</u> &copy; after";
        var document = new AdfDocument();
        document.Content.Add(new AdfNode("paragraph") { Content = { AdfNode.TextNode(content) } });

        AdfConversionResult<string> markdown = AdfConverter.ToMarkdown(document);
        AdfConversionResult<AdfDocument> roundTrip = AdfConverter.FromMarkdown(markdown.Value);

        AdfNode paragraph = Assert.Single(roundTrip.Value.Content);
        Assert.Equal("paragraph", paragraph.Type);
        Assert.Equal(content, string.Concat(paragraph.Content.Select(node => node.Text)));
    }

    [Fact]
    public void ParagraphProjection_PreservesTerminalHardBreak() {
        var document = new AdfDocument();
        document.Content.Add(new AdfNode("paragraph") {
            Content = { AdfNode.TextNode("ready"), new AdfNode("hardBreak") }
        });

        AdfConversionResult<string> markdown = AdfConverter.ToMarkdown(document);
        AdfConversionResult<AdfDocument> roundTrip = AdfConverter.FromMarkdown(markdown.Value);

        Assert.EndsWith("<br />", markdown.Value);
        AdfNode paragraph = Assert.Single(roundTrip.Value.Content);
        Assert.Collection(
            paragraph.Content,
            text => Assert.Equal("ready", text.Text),
            hardBreak => Assert.Equal("hardBreak", hardBreak.Type));
    }

    [Fact]
    public void TableProjection_PreservesLiteralPipesAndHardBreaks() {
        var document = new AdfDocument();
        document.Content.Add(new AdfNode("table") {
            Content = {
                new AdfNode("tableRow") {
                    Content = {
                        new AdfNode("tableHeader") {
                            Content = { new AdfNode("paragraph") { Content = { AdfNode.TextNode("Header") } } }
                        },
                        new AdfNode("tableHeader") {
                            Content = { new AdfNode("paragraph") { Content = { AdfNode.TextNode("Other") } } }
                        }
                    }
                },
                new AdfNode("tableRow") {
                    Content = {
                        new AdfNode("tableCell") {
                            Content = {
                                new AdfNode("paragraph") {
                                    Content = {
                                        AdfNode.TextNode("left|right"),
                                        new AdfNode("hardBreak"),
                                        AdfNode.TextNode("next")
                                    }
                                }
                            }
                        },
                        new AdfNode("tableCell") {
                            Content = { new AdfNode("paragraph") { Content = { AdfNode.TextNode("stable") } } }
                        }
                    }
                }
            }
        });

        AdfConversionResult<string> markdown = AdfConverter.ToMarkdown(document);
        AdfConversionResult<AdfDocument> roundTrip = AdfConverter.FromMarkdown(markdown.Value);

        Assert.Contains("left\\|right<br />next", markdown.Value);
        AdfNode table = Assert.Single(roundTrip.Value.Content);
        Assert.Equal(2, table.Content.Count);
        Assert.All(table.Content, row => Assert.Equal(2, row.Content.Count));
        AdfNode paragraph = Assert.Single(table.Content[1].Content[0].Content);
        Assert.Collection(
            paragraph.Content,
            text => Assert.Equal("left", text.Text),
            pipe => Assert.Equal("|", pipe.Text),
            text => Assert.Equal("right", text.Text),
            hardBreak => Assert.Equal("hardBreak", hardBreak.Type),
            text => Assert.Equal("next", text.Text));
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

    [Fact]
    public void InlineCodeProjection_ReportsMarkdownLineBreakNormalization() {
        const string content = "first\nsecond";
        var document = new AdfDocument();
        document.Content.Add(new AdfNode("paragraph") {
            Content = { AdfNode.TextNode(content, new[] { new AdfMark("code") }) }
        });

        AdfConversionResult<string> markdown = AdfConverter.ToMarkdown(document);
        AdfConversionResult<AdfDocument> roundTrip = AdfConverter.FromMarkdown(markdown.Value);

        Assert.Contains(markdown.Report.Diagnostics, item => item.Code == "ADF_CODE_MARK_LINE_BREAK_NORMALIZED");
        AdfNode text = Assert.Single(Assert.Single(roundTrip.Value.Content).Content);
        Assert.Equal("first second", text.Text);
        Assert.Contains(text.Marks, mark => mark.Type == "code");
    }
}
