using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Xunit;

namespace OfficeIMO.Adf.Tests;

public sealed class AdfDestinationPolicyTests {
    [Theory]
    [InlineData(AdfValidationProfile.ForwardCompatible)]
    [InlineData(AdfValidationProfile.FullSchema)]
    public void DestinationRestrictionsReportNestedNodesAndMarksWithoutChangingSource(AdfValidationProfile profile) {
        const string source = "{\"type\":\"doc\",\"version\":1,\"content\":[{\"type\":\"blockquote\",\"content\":[{\"type\":\"paragraph\",\"content\":[{\"type\":\"text\",\"text\":\"Visible\",\"marks\":[{\"type\":\"em\"}]}]}]}]}";
        AdfDocument document = AdfDocument.Parse(source);
        var policy = new AdfDestinationPolicy("Qualified destination 2026-10", new[] { "paragraph", "text" }, new[] { "strong" });
        Assert.True(document.Validate(new AdfValidationOptions { Profile = profile }).IsValid);
        string native = document.ToJson();
        AdfValidationResult result = document.Validate(new AdfValidationOptions { Profile = profile, DestinationPolicy = policy });
        Assert.False(result.IsValid);
        Assert.Collection(result.Issues,
            issue => { Assert.Equal("ADF_DESTINATION_NODE", issue.Code); Assert.Equal("$.content[0].type", issue.Path); },
            issue => { Assert.Equal("ADF_DESTINATION_MARK", issue.Code); Assert.Equal("$.content[0].content[0].content[0].marks[0].type", issue.Path); });
        Assert.All(result.Issues, issue => Assert.Contains(policy.Name, issue.Message));
        Assert.Equal(native, document.ToJson());
        Assert.Throws<NotSupportedException>(() => ((IList<AdfValidationIssue>)result.Issues).Clear());
    }

    [Fact]
    public void AllowListsAreImmutableCaseSensitiveSnapshotsAndDoNotReplaceSchemaChecks() {
        var names = new List<string> { "paragraph", "text", "mention", "mention" };
        var policy = new AdfDestinationPolicy("Target", names, Array.Empty<string>());
        names.Clear();
        Assert.Equal(3, policy.AllowedNodeTypes!.Count);
        Assert.Throws<NotSupportedException>(() => ((IList<string>)policy.AllowedNodeTypes).Clear());
        var document = new AdfDocument(new[] { new AdfNode("paragraph") { Content = { new AdfNode("mention") } } });
        AdfValidationResult invalid = document.Validate(new AdfValidationOptions { Profile = AdfValidationProfile.FullSchema, DestinationPolicy = policy });
        Assert.False(invalid.IsValid);
        Assert.DoesNotContain(invalid.Issues, issue => issue.Code.StartsWith("ADF_DESTINATION", StringComparison.Ordinal));
        document.Content[0].Content[0] = AdfNode.TextNode("Text", new[] { new AdfMark("strong") });
        Assert.Contains(document.Validate(new AdfValidationOptions { DestinationPolicy = policy }).Issues, issue => issue.Code == "ADF_DESTINATION_MARK");
        var uppercase = new AdfDestinationPolicy("Target", new[] { "Paragraph", "text" });
        Assert.Contains(document.Validate(new AdfValidationOptions { DestinationPolicy = uppercase }).Issues, issue => issue.Code == "ADF_DESTINATION_NODE");
    }

    [Fact]
    public void CallerAllowedFutureNodesRemainForwardCompatibleButFailThePinnedSchema() {
        var document = new AdfDocument(new[] { new AdfNode("vendorFuture") });
        var options = new AdfValidationOptions { DestinationPolicy = new AdfDestinationPolicy("Future target", new[] { "vendorFuture" }) };
        AdfValidationResult future = document.Validate(options);
        Assert.True(future.IsValid);
        Assert.Contains(future.Issues, issue => issue.Code == "ADF_UNKNOWN_NODE");
        options.Profile = AdfValidationProfile.FullSchema;
        Assert.False(document.Validate(options).IsValid);
    }

    [Fact]
    public void GeneratedMarkdownAndHtmlReportDestinationFailuresAndRetainInspectableValues() {
        var options = new AdfConversionOptions {
            DestinationValidation = new AdfValidationOptions {
                Profile = AdfValidationProfile.FullSchema,
                DestinationPolicy = new AdfDestinationPolicy("Paragraph-only target", new[] { "paragraph", "text" }, Array.Empty<string>())
            }
        };
        var markdown = AdfConverter.FromMarkdown("# Heading\n\n**Visible**", options);
        Assert.True(markdown.Report.HasErrors);
        Assert.Contains(markdown.Report.Diagnostics, issue => issue.Code == "ADF_DESTINATION_NODE");
        Assert.Contains(markdown.Report.Diagnostics, issue => issue.Code == "ADF_DESTINATION_MARK");
        Assert.Throws<InvalidOperationException>(() => markdown.Report.RequireNoLoss());
        Assert.Equal("heading", markdown.Value.Content[0].Type);
        var html = AdfConverter.FromHtml("<h1>Heading</h1><p>Visible</p>", null, options);
        Assert.True(html.Report.HasErrors);
        Assert.Contains(html.Report.Diagnostics, issue => issue.Code == "ADF_DESTINATION_NODE");
        Assert.Contains("Visible", AdfConverter.ToMarkdown(html.Value).Value);
        Assert.False(AdfConverter.FromMarkdown("Visible", options).Report.HasErrors);
    }

    [Fact]
    public void PolicyValidationPreservesGraphLimitsAndCancellationAndBoundsDiagnostics() {
        var policy = new AdfDestinationPolicy("Empty target", Array.Empty<string>());
        var options = new AdfValidationOptions { DestinationPolicy = policy };
        var node = new AdfNode("paragraph"); node.Content.Add(node);
        var document = new AdfDocument(new[] { node });
        Assert.Contains(document.Validate(options).Issues, issue => issue.Code == "ADF_CYCLIC_CONTENT");
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => document.Validate(options, cancellation.Token));
        document.Content.Clear();
        for (int index = 0; index < 1005; index++) document.Content.Add(new AdfNode("paragraph"));
        AdfValidationResult bounded = document.Validate(options);
        Assert.False(bounded.IsValid);
        Assert.Equal(1000, bounded.Issues.Count(issue => issue.Code == "ADF_DESTINATION_NODE"));
    }
}
