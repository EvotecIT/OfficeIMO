using System;
using System.Linq;
using System.Threading;
using OfficeIMO.Adf;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Adf.Tests;

public sealed class AdfProcessingIntegrationTests {
    [Fact]
    public void SharedProcessingTypeRetainsTheSelectedFullSchemaProfile() {
        var document = new AdfDocument(new[] { new AdfNode("vendorFutureBlock") });
        AdfProcessingOptions options = new AdfValidationOptions { Profile = AdfValidationProfile.FullSchema };
        Assert.False(document.Validate(options).IsValid);
    }

    [Theory]
    [InlineData(AdfValidationProfile.ForwardCompatible)]
    [InlineData(AdfValidationProfile.FullSchema)]
    public void ValidationProfilesUseCallerGraphLimits(AdfValidationProfile profile) {
        var document = new AdfDocument(new[] { new AdfNode("paragraph") { Content = { AdfNode.TextNode("text") } } });
        var options = new AdfValidationOptions { Profile = profile, MaxDepth = 1 };
        Assert.Contains(document.Validate(options).Issues, issue => issue.Code == "ADF_CONTENT_DEPTH_EXCEEDED");
        options.MaxDepth = 64; options.MaxNodes = 1;
        Assert.Contains(document.Validate(options).Issues, issue => issue.Code == "ADF_NODE_LIMIT_EXCEEDED");
        options.MaxNodes = 100; options.MaxTextCharacters = 3;
        Assert.Contains(document.Validate(options).Issues, issue => issue.Code == "ADF_TEXT_LIMIT_EXCEEDED");
    }

    [Theory]
    [InlineData(AdfValidationProfile.ForwardCompatible, true)]
    [InlineData(AdfValidationProfile.ForwardCompatible, false)]
    [InlineData(AdfValidationProfile.FullSchema, true)]
    [InlineData(AdfValidationProfile.FullSchema, false)]
    public void ValidationHonorsBothTokenSourcesWithoutMutatingOptions(AdfValidationProfile profile, bool cancelOptions) {
        using var optionToken = new CancellationTokenSource();
        using var argumentToken = new CancellationTokenSource();
        (cancelOptions ? optionToken : argumentToken).Cancel();
        var options = new AdfValidationOptions { Profile = profile, CancellationToken = optionToken.Token };
        Assert.Throws<OperationCanceledException>(() => new AdfDocument().Validate(options, argumentToken.Token));
        Assert.Equal(optionToken.Token, options.CancellationToken);
    }

    [Fact]
    public void HigherConfiguredDepthWorksForNativeAndFullSchemaValidation() {
        var inner = new AdfNode("paragraph") { Content = { AdfNode.TextNode("Deep") } };
        for (int index = 0; index < 35; index++) {
            var item = new AdfNode("listItem") { Content = { new AdfNode("paragraph") } };
            if (inner.Type == "paragraph") item.Content[0] = inner;
            else item.Content.Add(inner);
            inner = new AdfNode("bulletList") { Content = { item } };
        }
        var document = new AdfDocument(new[] { inner });
        Assert.Contains(document.Validate().Issues, issue => issue.Code == "ADF_CONTENT_DEPTH_EXCEEDED");
        var options = new AdfValidationOptions { Profile = AdfValidationProfile.FullSchema, MaxDepth = 80 };
        Assert.True(document.Validate(options).IsValid);
        Assert.True(AdfDocument.Parse(document.ToJson(options), options).Validate(options).IsValid);
    }

    [Fact]
    public void ExtensionResolverGraphsAreCheckedBeforeRendering() {
        var quote = new QuoteBlock(Array.Empty<string>());
        var projection = MarkdownDoc.Create().Add(quote);
        quote.ChildBlocks.Add(quote);
        var document = new AdfDocument(new[] { new AdfNode("extension") });
        Assert.Throws<InvalidOperationException>(() => AdfConverter.ToMarkdown(document, new AdfConversionOptions {
            ExtensionResolver = _ => projection
        }));
    }

    [Fact]
    public void MultipleNestedTaskBlocksReceiveDistinctStablePaths() {
        var outer = ListItem.Task("Outer");
        outer.NestedBlocks.Add(new UnorderedListBlock { Items = { ListItem.Task("First") } });
        outer.NestedBlocks.Add(new UnorderedListBlock { Items = { ListItem.Task("Second") } });
        var source = MarkdownDoc.Create().Add(new UnorderedListBlock { Items = { outer } });
        var result = AdfConverter.FromMarkdown(source, new AdfConversionOptions { LocalIdFactory = path => path });
        var root = Assert.Single(result.Value.Content);
        var nodes = new[] { root }.Concat(root.Content).Concat(root.Content.Where(node => node.Type == "taskList").SelectMany(node => node.Content));
        var identities = nodes.Select(node => node.GetStringAttribute("localId")).ToArray();
        Assert.Equal(6, identities.Length);
        Assert.Equal(identities.Length, identities.Distinct().Count());
        Assert.True(result.Value.Validate(new AdfValidationOptions { Profile = AdfValidationProfile.FullSchema }).IsValid);
    }
}
