using System.Linq;
using System.Text.Json;
using OfficeIMO.Adf;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Adf.Tests;

public sealed class AdfContractTests {
    [Fact]
    public void SetAttribute_PreservesCustomJsonValuesOnDynamicRuntimes() {
        var node = new AdfNode("extension").SetAttribute("parameters", new {
            html = "<strong>Ready</strong>",
            enabled = true,
            levels = new[] { 1, 2 }
        });

        JsonElement parameters = node.Attributes["parameters"];
        Assert.Equal("<strong>Ready</strong>", parameters.GetProperty("html").GetString());
        Assert.True(parameters.GetProperty("enabled").GetBoolean());
        Assert.Equal(2, parameters.GetProperty("levels").GetArrayLength());
    }

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
    public void MarkdownProjection_ReportsExplicitEmptyParagraphLoss() {
        var document = new AdfDocument(new[] { new AdfNode("paragraph") });

        AdfConversionResult<string> result = AdfConverter.ToMarkdown(document);

        Assert.Empty(result.Value);
        Assert.False(result.Report.IsLossless);
        AdfConversionDiagnostic diagnostic = Assert.Single(result.Report.Diagnostics, item => item.Code == "ADF_EMPTY_PARAGRAPH_DROPPED");
        Assert.Equal("$.content[0]", diagnostic.Path);
    }

    [Fact]
    public void HeadingProjection_ReportsDroppedProperties() {
        var heading = new AdfNode("heading") { Content = { AdfNode.TextNode("Status") } };
        heading.SetAttribute("level", 2).SetAttribute("localId", "heading-1");
        heading.ExtensionData["vendorHeadingOption"] = JsonSerializer.SerializeToElement(true);
        var document = new AdfDocument(new[] { heading });

        AdfConversionResult<string> result = AdfConverter.ToMarkdown(document);

        Assert.False(result.Report.IsLossless);
        AdfConversionDiagnostic diagnostic = Assert.Single(result.Report.Diagnostics, item => item.Code == "ADF_HEADING_PROPERTIES_DROPPED");
        Assert.Equal("$.content[0]", diagnostic.Path);
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
    [InlineData("{\"type\":\"file\",\"collection\":\"files\"}", "ADF_MEDIA_ID")]
    [InlineData("{\"type\":\"file\",\"id\":17,\"collection\":\"files\"}", "ADF_MEDIA_ID")]
    [InlineData("{\"type\":\"link\",\"id\":\"media-1\"}", "ADF_MEDIA_COLLECTION")]
    [InlineData("{\"type\":\"link\",\"id\":\"media-1\",\"collection\":42}", "ADF_MEDIA_COLLECTION")]
    [InlineData("{\"type\":\"external\"}", "ADF_MEDIA_URL")]
    [InlineData("{\"type\":\"external\",\"url\":42}", "ADF_MEDIA_URL")]
    [InlineData("{\"type\":\"unknown\"}", "ADF_MEDIA_TYPE")]
    [InlineData("{\"type\":42}", "ADF_MEDIA_TYPE")]
    public void Validation_RequiresSchemaDefinedMediaSourceAttributes(string attributesJson, string expectedCode) {
        AdfDocument document = AdfDocument.Parse(
            "{\"version\":1,\"type\":\"doc\",\"content\":[{\"type\":\"mediaSingle\",\"content\":[{\"type\":\"media\",\"attrs\":" + attributesJson + "}]}]}");

        AdfValidationResult result = document.Validate();

        Assert.False(result.IsValid);
        AdfValidationIssue issue = Assert.Single(result.Issues, item => item.Code == expectedCode);
        Assert.StartsWith("$.content[0].content[0].attrs.", issue.Path);
    }

    [Theory]
    [InlineData("{\"type\":\"file\",\"id\":\"media-1\",\"collection\":\"files\"}")]
    [InlineData("{\"type\":\"link\",\"id\":\"media-1\",\"collection\":\"\"}")]
    [InlineData("{\"type\":\"external\",\"url\":\"\"}")]
    public void Validation_AcceptsSchemaDefinedMediaSourceAttributes(string attributesJson) {
        AdfDocument document = AdfDocument.Parse(
            "{\"version\":1,\"type\":\"doc\",\"content\":[{\"type\":\"mediaSingle\",\"content\":[{\"type\":\"media\",\"attrs\":" + attributesJson + "}]}]}");

        Assert.True(document.Validate().IsValid);
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
    public void MarkdownTaskList_UsesNativeAdfTaskNodesAndRoundTrips() {
        AdfConversionResult<AdfDocument> result = AdfConverter.FromMarkdown("- [x] Ready\n- [ ] Pending");

        AdfNode list = Assert.Single(result.Value.Content);
        Assert.Equal("taskList", list.Type);
        Assert.False(string.IsNullOrWhiteSpace(list.GetStringAttribute("localId")));
        Assert.Collection(
            list.Content,
            item => {
                Assert.Equal("taskItem", item.Type);
                Assert.Equal("DONE", item.GetStringAttribute("state"));
                Assert.False(string.IsNullOrWhiteSpace(item.GetStringAttribute("localId")));
                Assert.Equal("Ready", Assert.Single(item.Content).Text);
            },
            item => {
                Assert.Equal("taskItem", item.Type);
                Assert.Equal("TODO", item.GetStringAttribute("state"));
                Assert.False(string.IsNullOrWhiteSpace(item.GetStringAttribute("localId")));
                Assert.Equal("Pending", Assert.Single(item.Content).Text);
            });
        Assert.True(result.Value.Validate().IsValid);
        Assert.DoesNotContain(result.Report.Diagnostics, item => item.Code == "MARKDOWN_TASK_LIST_FALLBACK");

        AdfConversionResult<string> roundTrip = AdfConverter.ToMarkdown(result.Value);
        Assert.Equal("- [x] Ready\n- [ ] Pending", roundTrip.Value.Replace("\r\n", "\n"));
        Assert.DoesNotContain(roundTrip.Report.Diagnostics, item => item.Code == "ADF_UNSUPPORTED_NODE");
    }

    [Fact]
    public void MarkdownComplexTaskList_UsesVisibleMarkerFallback() {
        var task = ListItem.Task("Ready", done: true);
        task.AdditionalParagraphs.Add(new InlineSequence().Text("Details"));
        var list = new UnorderedListBlock();
        list.Items.Add(task);

        AdfConversionResult<AdfDocument> result = AdfConverter.FromMarkdown(MarkdownDoc.Create().Add(list));

        Assert.Equal("bulletList", Assert.Single(result.Value.Content).Type);
        Assert.Contains(result.Report.Diagnostics, item => item.Code == "MARKDOWN_TASK_LIST_FALLBACK");
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
    public void Validation_RejectsKnownNonTextPayloadsAndMarks() {
        var paragraph = new AdfNode("paragraph") { Text = "invalid" };
        paragraph.Marks.Add(new AdfMark("strong"));
        var document = new AdfDocument(new[] { paragraph });

        AdfValidationResult result = document.Validate();

        Assert.False(result.IsValid);
        Assert.Contains(result.Issues, item => item.Code == "ADF_TEXT_NOT_ALLOWED" && item.Path == "$.content[0].text");
        Assert.Contains(result.Issues, item => item.Code == "ADF_MARKS_NOT_ALLOWED" && item.Path == "$.content[0].marks");
    }

    [Fact]
    public void Validation_RejectsEmptyTextNodes() {
        var document = new AdfDocument();
        document.Content.Add(new AdfNode("paragraph") { Content = { AdfNode.TextNode(string.Empty) } });

        AdfValidationResult result = document.Validate();

        Assert.False(result.IsValid);
        Assert.Contains(result.Issues, item => item.Code == "ADF_TEXT_REQUIRED" && item.Path == "$.content[0].content[0].text");
    }

    [Fact]
    public void Validation_ReportsNullNestedNodesWithoutThrowing() {
        var paragraph = new AdfNode("paragraph");
        paragraph.Content.Add(null!);
        var document = new AdfDocument(new[] { paragraph });

        AdfValidationResult result = document.Validate();

        Assert.False(result.IsValid);
        AdfValidationIssue issue = Assert.Single(result.Issues, item => item.Code == "ADF_NULL_NODE");
        Assert.Equal("$.content[0].content[0]", issue.Path);
    }

    [Theory]
    [InlineData("taskList", "{}", "ADF_TASK_LOCAL_ID")]
    [InlineData("taskItem", "{\"state\":\"DONE\"}", "ADF_TASK_LOCAL_ID")]
    [InlineData("taskItem", "{\"localId\":\"item-1\"}", "ADF_TASK_STATE")]
    [InlineData("taskItem", "{\"localId\":\"item-1\",\"state\":\"done\"}", "ADF_TASK_STATE")]
    public void Validation_RequiresTaskIdentityAndState(string nodeType, string attributesJson, string expectedCode) {
        AdfDocument document = AdfDocument.Parse(
            "{\"version\":1,\"type\":\"doc\",\"content\":[{\"type\":\"" + nodeType + "\",\"attrs\":" + attributesJson + "}]}");

        AdfValidationResult result = document.Validate();

        Assert.False(result.IsValid);
        Assert.Contains(result.Issues, item => item.Code == expectedCode);
    }

    [Theory]
    [InlineData("paragraph", "paragraph")]
    [InlineData("table", "paragraph")]
    [InlineData("codeBlock", "hardBreak")]
    [InlineData("mediaGroup", "paragraph")]
    [InlineData("rule", "text")]
    public void Validation_RejectsKnownInvalidParentChildPairs(string parentType, string childType) {
        var parent = new AdfNode(parentType);
        parent.Content.Add(childType == "text" ? AdfNode.TextNode("invalid") : new AdfNode(childType));
        var document = new AdfDocument(new[] { parent });

        AdfValidationResult result = document.Validate();

        Assert.False(result.IsValid);
        Assert.Contains(result.Issues, item => item.Code == "ADF_NODE_CHILD" && item.Path == "$.content[0].content[0]");
    }

    [Theory]
    [InlineData("paragraph", "text")]
    [InlineData("heading", "hardBreak")]
    [InlineData("codeBlock", "text")]
    [InlineData("blockquote", "paragraph")]
    [InlineData("bulletList", "listItem")]
    [InlineData("listItem", "paragraph")]
    [InlineData("taskList", "taskItem")]
    [InlineData("taskItem", "text")]
    [InlineData("table", "tableRow")]
    [InlineData("tableRow", "tableCell")]
    [InlineData("tableCell", "paragraph")]
    [InlineData("mediaSingle", "media")]
    [InlineData("mediaGroup", "media")]
    [InlineData("panel", "paragraph")]
    [InlineData("bodiedExtension", "paragraph")]
    public void Validation_AcceptsKnownParentChildPairs(string parentType, string childType) {
        var parent = new AdfNode(parentType);
        parent.Content.Add(childType == "text" ? AdfNode.TextNode("valid") : new AdfNode(childType));
        var document = new AdfDocument(new[] { parent });

        AdfValidationResult result = document.Validate();

        Assert.DoesNotContain(result.Issues, item => item.Code == "ADF_NODE_CHILD" && item.Path == "$.content[0].content[0]");
        Assert.DoesNotContain(result.Issues, item => item.Code == "ADF_LIST_CHILD" && item.Path == "$.content[0].content[0]");
        Assert.DoesNotContain(result.Issues, item => item.Code == "ADF_TASK_LIST_CHILD" && item.Path == "$.content[0].content[0]");
    }

    [Theory]
    [InlineData(false, null)]
    [InlineData(true, null)]
    [InlineData(false, "")]
    [InlineData(false, " ")]
    public void Validation_RequiresNonEmptyStringHrefOnLinkMarks(bool addNonStringHref, string? href) {
        var link = new AdfMark("link");
        if (addNonStringHref) link.SetAttribute("href", 42);
        else if (href != null) link.SetAttribute("href", href);
        var document = new AdfDocument();
        document.Content.Add(new AdfNode("paragraph") {
            Content = { AdfNode.TextNode("broken", new[] { link }) }
        });

        AdfValidationResult result = document.Validate();

        Assert.False(result.IsValid);
        AdfValidationIssue issue = Assert.Single(result.Issues, item => item.Code == "ADF_LINK_HREF_REQUIRED");
        Assert.Equal("$.content[0].content[0].marks[0].attrs.href", issue.Path);
    }

    [Theory]
    [InlineData("{}", false)]
    [InlineData("{\"level\":0}", false)]
    [InlineData("{\"level\":1}", true)]
    [InlineData("{\"level\":6}", true)]
    [InlineData("{\"level\":7}", false)]
    [InlineData("{\"level\":\"2\"}", false)]
    [InlineData("{\"level\":1.5}", false)]
    public void Validation_RequiresIntegerHeadingLevelFromOneThroughSix(string attributesJson, bool expectedValid) {
        AdfDocument document = AdfDocument.Parse(
            "{\"version\":1,\"type\":\"doc\",\"content\":[{\"type\":\"heading\",\"attrs\":" + attributesJson + ",\"content\":[{\"type\":\"text\",\"text\":\"Title\"}]}]}");

        AdfValidationResult result = document.Validate();

        Assert.Equal(expectedValid, result.IsValid);
        if (expectedValid) {
            Assert.DoesNotContain(result.Issues, item => item.Code == "ADF_HEADING_LEVEL");
        } else {
            AdfValidationIssue issue = Assert.Single(result.Issues, item => item.Code == "ADF_HEADING_LEVEL");
            Assert.Equal("$.content[0].attrs.level", issue.Path);
        }
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
    public void MarkdownEmptyCodeFence_ProducesValidEmptyAdfCodeBlock() {
        AdfConversionResult<AdfDocument> adf = AdfConverter.FromMarkdown("```\n```");

        AdfNode codeBlock = Assert.Single(adf.Value.Content);
        Assert.Equal("codeBlock", codeBlock.Type);
        Assert.Empty(codeBlock.Content);
        Assert.True(adf.Value.Validate().IsValid);

        AdfConversionResult<string> roundTrip = AdfConverter.ToMarkdown(adf.Value);
        Assert.Equal("```\n```", roundTrip.Value.Replace("\r\n", "\n"));
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

    [Fact]
    public void CodeBlockProjection_PreventsLanguageFromExtendingTildeFence() {
        const string content = "````\n~";
        var document = new AdfDocument();
        document.Content.Add(new AdfNode("codeBlock") {
            Content = { AdfNode.TextNode(content) }
        }.SetAttribute("language", "~~~"));
        document.Content.Add(new AdfNode("paragraph") { Content = { AdfNode.TextNode("After") } });

        AdfConversionResult<string> markdown = AdfConverter.ToMarkdown(document);
        AdfConversionResult<AdfDocument> roundTrip = AdfConverter.FromMarkdown(markdown.Value);

        Assert.StartsWith("~~~\n", markdown.Value.Replace("\r\n", "\n"));
        Assert.Contains(markdown.Report.Diagnostics, item => item.Code == "ADF_CODE_LANGUAGE_NORMALIZED");
        Assert.Collection(
            roundTrip.Value.Content,
            codeBlock => Assert.Equal(content, Assert.Single(codeBlock.Content).Text),
            paragraph => Assert.Equal("After", Assert.Single(paragraph.Content).Text));
    }

    [Fact]
    public void CodeBlockProjection_ReportsDroppedAttributesAndExtensionProperties() {
        var codeBlock = new AdfNode("codeBlock") { Content = { AdfNode.TextNode("Get-Date") } };
        codeBlock.SetAttribute("uniqueId", "code-1");
        codeBlock.ExtensionData["vendorCodeOption"] = JsonSerializer.SerializeToElement(true);
        var document = new AdfDocument(new[] { codeBlock });

        AdfConversionResult<string> result = AdfConverter.ToMarkdown(document);

        Assert.False(result.Report.IsLossless);
        AdfConversionDiagnostic diagnostic = Assert.Single(result.Report.Diagnostics, item => item.Code == "ADF_CODE_PROPERTIES_DROPPED");
        Assert.Equal("$.content[0]", diagnostic.Path);
    }

    [Fact]
    public void Validation_RejectsMarksInsideCodeBlockText() {
        var codeBlock = new AdfNode("codeBlock") {
            Content = { AdfNode.TextNode("Get-Date", new[] { new AdfMark("strong") }) }
        };
        var document = new AdfDocument(new[] { codeBlock });

        AdfValidationResult result = document.Validate();

        Assert.False(result.IsValid);
        AdfValidationIssue issue = Assert.Single(result.Issues, item => item.Code == "ADF_CODE_MARKS_NOT_ALLOWED");
        Assert.Equal("$.content[0].content[0].marks", issue.Path);
    }

    [Theory]
    [InlineData("first\r\nsecond")]
    [InlineData("first\rsecond")]
    public void CodeBlockProjection_ReportsCarriageReturnLineEndingNormalization(string content) {
        var document = new AdfDocument();
        document.Content.Add(new AdfNode("codeBlock") { Content = { AdfNode.TextNode(content) } });

        AdfConversionResult<string> markdown = AdfConverter.ToMarkdown(document);
        AdfConversionResult<AdfDocument> roundTrip = AdfConverter.FromMarkdown(markdown.Value);

        Assert.False(markdown.Report.IsLossless);
        AdfConversionDiagnostic diagnostic = Assert.Single(markdown.Report.Diagnostics, item => item.Code == "ADF_CODE_LINE_ENDINGS_NORMALIZED");
        Assert.Equal("$.content[0]", diagnostic.Path);
        Assert.Equal("first\nsecond", Assert.Single(Assert.Single(roundTrip.Value.Content).Content).Text);
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
    public void HeadingProjection_ReportsMultilineTextAsLossy() {
        var document = new AdfDocument();
        document.Content.Add(new AdfNode("heading") {
            Content = { AdfNode.TextNode("first\nsecond") }
        }.SetAttribute("level", 2));

        AdfConversionResult<string> result = AdfConverter.ToMarkdown(document);

        Assert.False(result.Report.IsLossless);
        AdfConversionDiagnostic diagnostic = Assert.Single(result.Report.Diagnostics, item => item.Code == "ADF_HEADING_LINE_BREAK_NORMALIZED");
        Assert.Equal("$.content[0].content[0]", diagnostic.Path);
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
    public void TableProjection_ReportsDroppedCellAttributes() {
        var document = new AdfDocument();
        document.Content.Add(new AdfNode("table") {
            Content = {
                new AdfNode("tableRow") {
                    Content = {
                        new AdfNode("tableHeader") {
                            Content = { new AdfNode("paragraph") { Content = { AdfNode.TextNode("Header") } } }
                        }.SetAttribute("colspan", 2).SetAttribute("colwidth", new[] { 120, 120 })
                    }
                }
            }
        });

        AdfConversionResult<string> result = AdfConverter.ToMarkdown(document);

        Assert.False(result.Report.IsLossless);
        AdfConversionDiagnostic diagnostic = Assert.Single(result.Report.Diagnostics, item => item.Code == "ADF_TABLE_CELL_ATTRIBUTES_DROPPED");
        Assert.Equal("$.content[0].content[0].content[0]", diagnostic.Path);
    }

    [Fact]
    public void TableProjection_ReportsDroppedTableAttributesAndProperties() {
        var table = new AdfNode("table").SetAttribute("layout", "wide");
        table.ExtensionData["vendorTableProperty"] = JsonSerializer.SerializeToElement(true);
        table.Content.Add(new AdfNode("tableRow") { Content = { TableCell("tableHeader", "Header") } });
        var document = new AdfDocument(new[] { table });

        AdfConversionResult<string> result = AdfConverter.ToMarkdown(document);

        Assert.False(result.Report.IsLossless);
        AdfConversionDiagnostic diagnostic = Assert.Single(result.Report.Diagnostics, item => item.Code == "ADF_TABLE_ATTRIBUTES_DROPPED");
        Assert.Equal("$.content[0]", diagnostic.Path);
    }

    [Fact]
    public void MarkdownProjection_ReportsDroppedRootExtensionProperties() {
        var document = new AdfDocument();
        document.ExtensionData["vendorRoot"] = JsonSerializer.SerializeToElement(true);
        document.Content.Add(new AdfNode("paragraph") { Content = { AdfNode.TextNode("Ready") } });

        AdfConversionResult<string> result = AdfConverter.ToMarkdown(document);

        Assert.False(result.Report.IsLossless);
        AdfConversionDiagnostic diagnostic = Assert.Single(result.Report.Diagnostics, item => item.Code == "ADF_ROOT_PROPERTIES_DROPPED");
        Assert.Equal("$", diagnostic.Path);
    }

    [Fact]
    public void TableProjection_ReportsFlattenedCellBlockStructure() {
        var header = TableCell("tableHeader", "First");
        header.Content.Add(new AdfNode("paragraph") { Content = { AdfNode.TextNode("Second") } });
        var document = new AdfDocument();
        document.Content.Add(new AdfNode("table") {
            Content = { new AdfNode("tableRow") { Content = { header } } }
        });

        AdfConversionResult<string> result = AdfConverter.ToMarkdown(document);

        Assert.False(result.Report.IsLossless);
        AdfConversionDiagnostic diagnostic = Assert.Single(result.Report.Diagnostics, item => item.Code == "ADF_TABLE_CELL_BLOCKS_FLATTENED");
        Assert.Equal("$.content[0].content[0].content[0]", diagnostic.Path);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void TableProjection_ReportsHeaderLayoutsMarkdownCannotRepresent(bool mixedFirstRow) {
        var firstRow = new AdfNode("tableRow");
        firstRow.Content.Add(TableCell("tableHeader", "First"));
        firstRow.Content.Add(TableCell(mixedFirstRow ? "tableCell" : "tableHeader", "Second"));
        var secondRow = new AdfNode("tableRow");
        secondRow.Content.Add(TableCell("tableCell", "Third"));
        secondRow.Content.Add(TableCell(mixedFirstRow ? "tableCell" : "tableHeader", "Fourth"));
        var document = new AdfDocument();
        document.Content.Add(new AdfNode("table") { Content = { firstRow, secondRow } });

        AdfConversionResult<string> result = AdfConverter.ToMarkdown(document);

        Assert.False(result.Report.IsLossless);
        AdfConversionDiagnostic diagnostic = Assert.Single(result.Report.Diagnostics, item => item.Code == "ADF_TABLE_HEADER_LAYOUT_NORMALIZED");
        Assert.Equal("$.content[0]", diagnostic.Path);
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

    [Theory]
    [InlineData("strong", "**ready**")]
    [InlineData("em", "*ready*")]
    [InlineData("strike", "~~ready~~")]
    public void DelimiterMarkProjection_MovesBoundaryWhitespaceOutsideMarkup(string markType, string renderedCore) {
        var document = new AdfDocument();
        document.Content.Add(new AdfNode("paragraph") {
            Content = {
                AdfNode.TextNode("before"),
                AdfNode.TextNode(" ready ", new[] { new AdfMark(markType) }),
                AdfNode.TextNode("after"),
            }
        });

        AdfConversionResult<string> markdown = AdfConverter.ToMarkdown(document);
        AdfConversionResult<AdfDocument> roundTrip = AdfConverter.FromMarkdown(markdown.Value);

        Assert.Contains(" " + renderedCore + " ", markdown.Value);
        Assert.False(markdown.Report.IsLossless);
        Assert.Contains(markdown.Report.Diagnostics, item => item.Code == "ADF_MARK_BOUNDARY_WHITESPACE_NORMALIZED");
        AdfNode paragraph = Assert.Single(roundTrip.Value.Content);
        Assert.Equal("before ready after", string.Concat(paragraph.Content.Select(item => item.Text)));
        Assert.Contains(paragraph.Content, item => item.Text == "ready" && item.Marks.Any(mark => mark.Type == markType));
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

    private static AdfNode TableCell(string type, string text) => new AdfNode(type) {
        Content = { new AdfNode("paragraph") { Content = { AdfNode.TextNode(text) } } }
    };
}
