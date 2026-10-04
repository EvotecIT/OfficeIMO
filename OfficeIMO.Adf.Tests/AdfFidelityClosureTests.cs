using System;
using System.Collections.Generic;
using OfficeIMO.Adf;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Adf.Tests;

public sealed class AdfFidelityClosureTests {
    [Fact]
    public void ReportDiagnosticsCannotDivergeFromItsCapturedFidelityEvidence() {
        var error = new AdfConversionDiagnostic("failure", "$", "Failed conversion", AdfConversionSeverity.Error);
        var source = new[] { error };
        var report = new AdfConversionReport(source);
        source[0] = new AdfConversionDiagnostic("info", "$", "Information", AdfConversionSeverity.Information);

        Assert.Same(error, Assert.Single(report.Diagnostics));
        Assert.Throws<NotSupportedException>(() => ((IList<AdfConversionDiagnostic>)report.Diagnostics)[0] = source[0]);
        Assert.True(report.HasErrors);
        Assert.True(report.HasLoss);
        Assert.Throws<InvalidOperationException>(() => report.RequireNoLoss());
    }

    [Theory]
    [InlineData("bulletList")]
    [InlineData("orderedList")]
    [InlineData("table")]
    [InlineData("blockquote")]
    public void ValidationRejectsEmptyRequiredContainers(string type) {
        var document = new AdfDocument(new[] { new AdfNode(type) });
        Assert.False(document.Validate().IsValid);
        Assert.Contains(document.Validate().Issues, issue => issue.Code == "ADF_CONTENT_REQUIRED");
    }

    [Theory]
    [InlineData("mention", "id")]
    [InlineData("emoji", "shortName")]
    public void RequiredInlineAttributesHaveSchemaStringTypes(string type, string attribute) {
        var inline = new AdfNode(type);
        var document = new AdfDocument(new[] { new AdfNode("paragraph") { Content = { inline } } });
        Assert.False(document.Validate().IsValid);
        inline.SetAttribute(attribute, 42);
        Assert.False(document.Validate().IsValid);
        inline.SetAttribute(attribute, "value");
        Assert.True(document.Validate().IsValid);
        Assert.True(document.Validate(new AdfValidationOptions { Profile = AdfValidationProfile.FullSchema }).IsValid);
    }

    [Fact]
    public void ParagraphAlignmentIsAValidBlockMarkWithExplicitProjectionLoss() {
        var paragraph = new AdfNode("paragraph") { Content = { AdfNode.TextNode("Aligned") }, Marks = { new AdfMark("alignment").SetAttribute("align", "center") } };
        var document = new AdfDocument(new[] { paragraph });
        Assert.True(document.Validate().IsValid);
        Assert.Contains(AdfConverter.ToMarkdown(document).Report.Diagnostics, issue => issue.Code == "ADF_BLOCK_MARKS_DROPPED");
    }

    [Fact]
    public void TaskListInsideQuoteUsesValidVisibleFallback() {
        AdfConversionResult<AdfDocument> result = AdfConverter.FromMarkdown(MarkdownReader.Parse("> - [ ] Pending\n> - [x] Done"));
        Assert.True(result.Value.Validate().IsValid);
        Assert.True(result.Value.Validate(new AdfValidationOptions { Profile = AdfValidationProfile.FullSchema }).IsValid);
        AdfNode quote = Assert.Single(result.Value.Content);
        Assert.Equal("blockquote", quote.Type);
        Assert.Equal("bulletList", Assert.Single(quote.Content).Type);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "MARKDOWN_TASK_LIST_FALLBACK");
        string markdown = AdfConverter.ToMarkdown(result.Value).Value;
        Assert.Contains("Pending", markdown);
        Assert.Contains("Done", markdown);
    }

    [Fact]
    public void ParagraphIdentityAndVendorPropertiesAreReportedAsOmitted() {
        AdfDocument document = AdfDocument.Parse("{\"version\":1,\"type\":\"doc\",\"content\":[{\"type\":\"paragraph\",\"attrs\":{\"localId\":\"p1\"},\"vendor\":42,\"content\":[{\"type\":\"text\",\"text\":\"Visible\"}]}]}");
        var result = AdfConverter.ToMarkdown(document);
        Assert.Contains("Visible", result.Value);
        Assert.False(result.Report.IsLossless);
        Assert.Contains(result.Report.Diagnostics, issue => issue.Code == "ADF_NODE_PROPERTIES_DROPPED" && issue.Path == "$.content[0]");
        Assert.Throws<InvalidOperationException>(() => result.Report.RequireNoLoss());
        Assert.Equal("p1", Assert.Single(AdfDocument.Parse(document.ToJson()).Content).GetStringAttribute("localId"));
    }

    [Theory]
    [InlineData("{\"type\":\"paragraph\",\"marks\":[{\"type\":\"alignment\",\"attrs\":{\"align\":\"sideways\"}}]}")]
    [InlineData("{\"type\":\"orderedList\",\"attrs\":{\"order\":-1},\"content\":[{\"type\":\"listItem\",\"content\":[{\"type\":\"paragraph\"}]}]}")]
    [InlineData("{\"type\":\"paragraph\",\"attrs\":{\"localId\":42}}")]
    [InlineData("{\"type\":\"table\",\"content\":[]}")]
    [InlineData("{\"type\":\"mediaSingle\",\"attrs\":{\"layout\":\"sideways\"},\"content\":[{\"type\":\"media\",\"attrs\":{\"type\":\"external\",\"url\":\"https://example.test/image\"}}]}")]
    public void StrictValidationEnforcesAttributeAndContainerSchema(string nodeJson) {
        AdfDocument document = AdfDocument.Parse("{\"version\":1,\"type\":\"doc\",\"content\":[" + nodeJson + "]}");
        Assert.False(document.Validate(new AdfValidationOptions { Profile = AdfValidationProfile.FullSchema }).IsValid);
    }

    [Fact]
    public void UnknownVendorNodesPreserveSourceAndHaveExplicitStrictValidationFailure() {
        AdfDocument document = AdfDocument.Parse("{\"version\":1,\"type\":\"doc\",\"content\":[{\"type\":\"vendorFuture\",\"attrs\":{\"payload\":42}}]}");
        Assert.True(document.Validate().IsValid);
        Assert.False(document.Validate(new AdfValidationOptions { Profile = AdfValidationProfile.FullSchema }).IsValid);
        Assert.Equal(document.ToJson(), AdfDocument.Parse(document.ToJson()).ToJson());
    }
}
