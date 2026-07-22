using OfficeIMO.Adf;
using Xunit;

namespace OfficeIMO.Adf.Tests;

public sealed class AdfSchemaValidationTests {
    [Fact]
    public void Validation_RejectsEmptyListItems() {
        var listItem = new AdfNode("listItem");
        var list = new AdfNode("bulletList") { Content = { listItem } };
        var document = new AdfDocument(new[] { list });

        AdfValidationResult result = document.Validate();

        Assert.False(result.IsValid);
        AdfValidationIssue issue = Assert.Single(result.Issues, item => item.Code == "ADF_LIST_ITEM_CONTENT_REQUIRED");
        Assert.Equal("$.content[0].content[0].content", issue.Path);
    }

    [Fact]
    public void Validation_AllowsSchemaDefinedListItemWithoutLeadingParagraph() {
        var codeBlock = new AdfNode("codeBlock") { Content = { AdfNode.TextNode("Get-Date") } };
        var listItem = new AdfNode("listItem") { Content = { codeBlock } };
        var list = new AdfNode("bulletList") { Content = { listItem } };
        var document = new AdfDocument(new[] { list });

        AdfValidationResult result = document.Validate();

        Assert.True(result.IsValid);
    }

    [Theory]
    [InlineData("{}")]
    [InlineData("{\"url\":42}")]
    [InlineData("{\"url\":\" \"}")]
    [InlineData("{\"data\":null}")]
    [InlineData("{\"url\":\"https://example.test\",\"data\":{}}")]
    public void Validation_RejectsInlineCardsWithoutExactlyOneValidTarget(string attributesJson) {
        AdfDocument document = InlineCardDocument(attributesJson);

        AdfValidationResult result = document.Validate();

        Assert.False(result.IsValid);
        AdfValidationIssue issue = Assert.Single(result.Issues, item => item.Code == "ADF_INLINE_CARD_TARGET");
        Assert.Equal("$.content[0].content[0].attrs", issue.Path);
    }

    [Theory]
    [InlineData("{\"url\":\"https://example.test/card\"}")]
    [InlineData("{\"data\":{\"@type\":\"Link\",\"url\":\"https://example.test/card\"}}")]
    public void Validation_AcceptsInlineCardUrlOrDataTarget(string attributesJson) {
        AdfDocument document = InlineCardDocument(attributesJson);

        AdfValidationResult result = document.Validate();

        Assert.True(result.IsValid);
    }

    private static AdfDocument InlineCardDocument(string attributesJson) => AdfDocument.Parse(
        "{\"version\":1,\"type\":\"doc\",\"content\":[{\"type\":\"paragraph\",\"content\":[{\"type\":\"inlineCard\",\"attrs\":" + attributesJson + "}]}]}");
}
