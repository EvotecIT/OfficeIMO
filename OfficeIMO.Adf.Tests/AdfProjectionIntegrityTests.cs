using System;
using System.Linq;
using System.Text.Json;
using OfficeIMO.Adf;
using Xunit;

namespace OfficeIMO.Adf.Tests;

public sealed class AdfProjectionIntegrityTests {
    [Theory]
    [InlineData("strong", "em")]
    [InlineData("em", "strong")]
    [InlineData("strong", "strong")]
    [InlineData("em", "em")]
    [InlineData("strike", "strike")]
    [InlineData("code", "code")]
    public void AdjacentMarkedRunsPreserveVisibleTextAndStyling(string first, string second) {
        var document = Paragraph(AdfNode.TextNode("a", new[] { new AdfMark(first) }), AdfNode.TextNode("b", new[] { new AdfMark(second) }));
        var markdown = AdfConverter.ToMarkdown(document);
        Assert.False(markdown.Report.HasLoss);
        var roundTrip = AdfConverter.FromMarkdown(markdown.Value);
        Assert.False(roundTrip.Report.HasLoss);
        AdfNode[] text = roundTrip.Value.Content.Single().Content.ToArray();
        Assert.Equal("ab", string.Concat(text.Select(node => node.Text)));
        Assert.Contains(text, node => node.Text == "a" && node.Marks.Any(mark => mark.Type == first));
        Assert.Contains(text, node => node.Text == "b" && node.Marks.Any(mark => mark.Type == second));
    }

    [Theory]
    [InlineData("paragraph")]
    [InlineData("bulletList")]
    [InlineData("listItem")]
    [InlineData("rule")]
    public void ProjectedNodeMetadataIsReportedAsOmission(string type) {
        var node = new AdfNode(type).SetAttribute("localId", "source-id");
        if (type == "paragraph") node.Content.Add(AdfNode.TextNode("Visible"));
        var result = AdfConverter.ToMarkdown(new AdfDocument(new[] { node }));
        Assert.True(result.Report.HasLoss);
        Assert.Contains(result.Report.FidelityDiagnostics, diagnostic => diagnostic.Code == "ADF_NODE_PROPERTIES_DROPPED" && diagnostic.LossKind == OfficeConversionLossKind.Omission);
        Assert.Throws<InvalidOperationException>(() => result.Report.RequireNoLoss());
    }

    [Fact]
    public void ParagraphAlignmentIsValidAndItsProjectionReportsTheLostMark() {
        var document = Paragraph(AdfNode.TextNode("Visible"));
        document.Content[0].Marks.Add(new AdfMark("alignment").SetAttribute("align", "center"));
        Assert.True(document.Validate().IsValid);
        var result = AdfConverter.ToMarkdown(document);
        Assert.Equal("Visible", result.Value);
        Assert.Contains(result.Report.FidelityDiagnostics, diagnostic => diagnostic.Code == "ADF_NODE_MARKS_DROPPED" && diagnostic.LossKind == OfficeConversionLossKind.Omission);
    }

    [Theory]
    [InlineData("mention", "text", "@Ada")]
    [InlineData("emoji", "text", "🙂")]
    [InlineData("emoji", "shortName", ":smile:")]
    [InlineData("inlineCard", "url", "https://example.com/card")]
    public void SemanticInlineFallbackRetainsVisibleContent(string type, string attribute, string value) {
        var result = AdfConverter.ToMarkdown(Paragraph(new AdfNode(type).SetAttribute(attribute, value)));
        var roundTrip = AdfConverter.FromMarkdown(result.Value);
        Assert.Equal(value, string.Concat(roundTrip.Value.Content.Single().Content.Select(node => node.Text)));
        Assert.Contains(result.Report.FidelityDiagnostics, diagnostic => diagnostic.Code == "ADF_SEMANTIC_NODE_PROJECTED" && diagnostic.LossKind == OfficeConversionLossKind.Omission);
    }

    [Fact]
    public void InlineWithoutVisibleFallbackReportsActualOmission() {
        var result = AdfConverter.ToMarkdown(Paragraph(new AdfNode("mention").SetAttribute("id", "opaque-id")));
        Assert.Equal(string.Empty, result.Value);
        Assert.Contains(result.Report.FidelityDiagnostics, diagnostic => diagnostic.Code == "ADF_UNSUPPORTED_INLINE_OMITTED" && diagnostic.LossKind == OfficeConversionLossKind.Omission);
    }

    [Fact]
    public void NativeJsonPreservesEmptyPropertiesIncludingRequiredRowContent() {
        const string source = "{\"version\":1,\"type\":\"doc\",\"content\":[{\"type\":\"table\",\"content\":[{\"type\":\"tableRow\",\"content\":[]}]},{\"type\":\"paragraph\",\"attrs\":{},\"marks\":[],\"content\":[]},{\"type\":\"vendor\",\"attrs\":{},\"marks\":[],\"content\":[]}]}";
        using var written = JsonDocument.Parse(AdfDocument.Parse(source).ToJson());
        var content = written.RootElement.GetProperty("content");
        Assert.Equal(0, content[0].GetProperty("content")[0].GetProperty("content").GetArrayLength());
        foreach (int index in new[] { 1, 2 }) {
            Assert.Empty(content[index].GetProperty("attrs").EnumerateObject());
            Assert.Equal(0, content[index].GetProperty("content").GetArrayLength());
            Assert.Equal(0, content[index].GetProperty("marks").GetArrayLength());
        }
        using var created = JsonDocument.Parse(new AdfDocument(new[] { new AdfNode("tableRow") }).ToJson());
        Assert.Equal(0, created.RootElement.GetProperty("content")[0].GetProperty("content").GetArrayLength());
    }

    [Theory]
    [InlineData("bulletList")]
    [InlineData("orderedList")]
    [InlineData("table")]
    [InlineData("tableCell")]
    [InlineData("panel")]
    public void KnownRequiredContentCannotBeEmpty(string type) {
        var result = new AdfDocument(new[] { new AdfNode(type) }).Validate();
        Assert.False(result.IsValid);
        Assert.Contains(result.Issues, issue => issue.Code == "ADF_CONTENT_REQUIRED");
    }

    [Fact]
    public void MediaSingleLinkMarksAndPanelTypesFollowNativeShape() {
        var media = new AdfNode("media").SetAttribute("type", "external").SetAttribute("url", "https://example.com/a.png");
        var container = new AdfNode("mediaSingle");
        container.Content.Add(media);
        container.Marks.Add(new AdfMark("link").SetAttribute("href", "https://example.com"));
        Assert.True(new AdfDocument(new[] { container }).Validate().IsValid);
        var panel = new AdfNode("panel");
        panel.Content.Add(new AdfNode("paragraph"));
        var document = new AdfDocument(new[] { panel });
        Assert.Contains(document.Validate().Issues, issue => issue.Code == "ADF_PANEL_TYPE");
        panel.SetAttribute("panelType", "info");
        Assert.True(document.Validate().IsValid);
    }

    [Fact]
    public void StandaloneMarkdownImageUsesExternalMediaAndRetainsUrlAndAltText() {
        var adf = AdfConverter.FromMarkdown("![Diagram](https://example.com/a.png)");
        Assert.False(adf.Report.HasLoss);
        Assert.Equal("mediaSingle", adf.Value.Content.Single().Type);
        Assert.True(adf.Value.Validate().IsValid);
        var result = AdfConverter.ToMarkdown(adf.Value);
        Assert.Equal("![Diagram](https://example.com/a.png)", result.Value);
    }

    [Fact]
    public void NestedTaskListsRetainHierarchyAndCompletion() {
        var result = AdfConverter.FromMarkdown("- [x] outer\n  - [ ] inner\n  - [x] ready\n- [ ] next");
        Assert.True(result.Value.Validate().IsValid);
        Assert.False(result.Report.HasLoss);
        var list = result.Value.Content.Single();
        Assert.Equal("taskList", list.Type);
        Assert.Equal("DONE", list.Content[0].GetStringAttribute("state"));
        Assert.Equal("taskList", list.Content[1].Type);
        Assert.Equal("DONE", list.Content[1].Content[1].GetStringAttribute("state"));
        string markdown = AdfConverter.ToMarkdown(result.Value).Value;
        Assert.Contains("  - [ ] inner", markdown);
        var back = AdfConverter.FromMarkdown(markdown).Value;
        Assert.Equal("taskList", back.Content.Single().Content[1].Type);
    }

    private static AdfDocument Paragraph(params AdfNode[] nodes) {
        var paragraph = new AdfNode("paragraph");
        paragraph.Content.AddRange(nodes);
        return new AdfDocument(new[] { paragraph });
    }
}
