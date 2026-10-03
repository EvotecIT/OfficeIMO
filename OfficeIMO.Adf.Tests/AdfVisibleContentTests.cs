using System;
using System.Linq;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Adf.Tests;

public sealed class AdfVisibleContentTests {
    [Theory]
    [InlineData("> # Heading")]
    [InlineData("> > Nested quote")]
    [InlineData("> ---")]
    public void QuotedBlocksUseSchemaValidVisibleFallbacks(string markdown) {
        var result = AdfConverter.FromMarkdown(markdown);
        Assert.True(result.Value.Validate(new AdfValidationOptions { Profile = AdfValidationProfile.FullSchema }).IsValid);
        Assert.Contains(result.Report.Diagnostics, issue => issue.Code == "MARKDOWN_CONTEXT_FALLBACK");
        Assert.NotEmpty(AdfConverter.ToMarkdown(result.Value).Value);
    }
    [Fact]
    public void NativeInlineMetadataHasVisibleContentAndExplicitLoss() {
        var paragraph = new AdfNode("paragraph");
        paragraph.Content.Add(new AdfNode("mention").SetAttribute("id", "person").SetAttribute("text", "@Jane"));
        paragraph.Content.Add(new AdfNode("emoji").SetAttribute("shortName", ":smile:").SetAttribute("text", "🙂"));
        paragraph.Content.Add(new AdfNode("status").SetAttribute("text", "Approved"));
        paragraph.Content.Add(new AdfNode("date").SetAttribute("timestamp", "0"));
        paragraph.Content.Add(new AdfNode("inlineCard").SetAttribute("url", "https://example.test/item"));
        var document = new AdfDocument(new[] { paragraph });
        string native = document.ToJson();
        var result = AdfConverter.ToMarkdown(document);
        Assert.Contains("@Jane", result.Value);
        Assert.Contains("🙂", result.Value);
        Assert.Contains("Approved", result.Value);
        Assert.Contains("1970-01-01", result.Value);
        Assert.Contains("https://example.test/item", result.Value);
        Assert.True(result.Report.HasLoss);
        Assert.Equal(native, document.ToJson());
    }

    [Fact]
    public void NestedTasksMediaAndRichTableCellsKeepTheirVisibleContent() {
        AdfNode Task(string id, string text) => new AdfNode("taskItem") { Content = { AdfNode.TextNode(text) } }.SetAttribute("localId", id).SetAttribute("state", "DONE");
        var tasks = new AdfNode("taskList").SetAttribute("localId", "list");
        tasks.Content.Add(Task("parent", "Parent"));
        tasks.Content.Add(new AdfNode("taskList") { Content = { Task("child", "Child") } }.SetAttribute("localId", "child-list"));
        var image = new AdfNode("mediaSingle") { Content = { new AdfNode("media").SetAttribute("type", "external").SetAttribute("url", "https://example.test/image.png").SetAttribute("alt", "Image") } };
        var cellParagraph = new AdfNode("paragraph") { Content = { AdfNode.TextNode("Inside table") } };
        var cellList = new AdfNode("bulletList") { Content = { new AdfNode("listItem") { Content = { cellParagraph } } } };
        var richCell = new AdfNode("tableHeader") { Content = { cellList } };
        richCell.SetAttribute("localId", "retained");
        var table = new AdfNode("table") { Content = { new AdfNode("tableRow") { Content = { richCell } }.SetAttribute("localId", "row") } };
        var result = AdfConverter.ToMarkdown(new AdfDocument(new[] { tasks, image, table }));
        Assert.Contains("- [x] Parent", result.Value);
        Assert.Contains("  - [x] Child", result.Value);
        Assert.Contains("![Image](https://example.test/image.png)", result.Value);
        Assert.Contains("Inside table", result.Value);
        Assert.Contains(result.Report.Diagnostics, issue => issue.Code == "ADF_TABLE_CELL_BLOCKS_FLATTENED");
        Assert.Contains(result.Report.Diagnostics, issue => issue.Code == "ADF_NODE_PROPERTIES_DROPPED" && issue.Path == "$.content[2].content[0]");
    }

    [Fact]
    public void TaskIdentityIsDeterministicAndCallerCanSupplyIdentityPolicy() {
        const string markdown = "- [ ] First\n- [x] Second";
        Assert.Equal(AdfConverter.FromMarkdown(markdown).Value.ToJson(), AdfConverter.FromMarkdown(markdown).Value.ToJson());
        var options = new AdfConversionOptions { LocalIdFactory = path => "custom:" + path };
        AdfDocument document = AdfConverter.FromMarkdown(markdown, options).Value;
        Assert.Equal("custom:$.blocks[0]", document.Content.Single().GetStringAttribute("localId"));
        Assert.Equal(3, document.Content.Single().Content.Select(node => node.GetStringAttribute("localId")).Append(document.Content.Single().GetStringAttribute("localId")).Distinct().Count());
        options.LocalIdFactory = _ => "duplicate";
        Assert.Throws<InvalidOperationException>(() => AdfConverter.FromMarkdown(markdown, options));
        Assert.True(document.Validate(new AdfValidationOptions { Profile = AdfValidationProfile.FullSchema }).IsValid);
    }

    [Fact]
    public void CallerResolverProjectsExtensionWithoutChangingNativePayload() {
        var node = new AdfNode("extension").SetAttribute("extensionType", "application.test").SetAttribute("extensionKey", "widget").SetAttribute("parameters", new { value = 42 });
        var document = new AdfDocument(new[] { node });
        string native = document.ToJson();
        var options = new AdfConversionOptions { ExtensionResolver = source => source.GetStringAttribute("extensionKey") == "widget" ? MarkdownReader.Parse("Resolved widget") : null };
        var result = AdfConverter.ToMarkdown(document, options);
        Assert.Contains("Resolved widget", result.Value);
        Assert.Contains(result.Report.Diagnostics, issue => issue.Code == "ADF_EXTENSION_RESOLVED");
        Assert.Equal(native, document.ToJson());
    }

    [Fact]
    public void ExtensionResolverHandlesInlineExtensionsAndOnlyReceivesExtensionNodes() {
        var inline = new AdfNode("inlineExtension").SetAttribute("extensionKey", "inline");
        var paragraph = new AdfNode("paragraph") { Content = { AdfNode.TextNode("Before "), inline, AdfNode.TextNode(" after") } };
        var document = new AdfDocument(new[] { paragraph, new AdfNode("vendorFuture") });
        int calls = 0;
        var options = new AdfConversionOptions { ExtensionResolver = node => { calls++; return MarkdownReader.Parse("*Resolved*"); } };
        var result = AdfConverter.ToMarkdown(document, options);
        Assert.Contains("Before *Resolved* after", result.Value);
        Assert.Equal(1, calls);
        Assert.Contains(result.Report.Diagnostics, issue => issue.Code == "ADF_EXTENSION_RESOLVED" && issue.Path == "$.content[0].content[1]");
        options.ExtensionResolver = _ => MarkdownReader.Parse("# Block");
        Assert.Throws<InvalidOperationException>(() => AdfConverter.ToMarkdown(document, options));
    }
}
