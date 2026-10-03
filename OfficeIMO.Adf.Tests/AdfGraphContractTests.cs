using System;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading;
using OfficeIMO.Adf;
using Xunit;

namespace OfficeIMO.Adf.Tests;

public sealed class AdfGraphContractTests {
    [Fact]
    public void DeepSupportedListsValidateWithoutCountingSchemaReferencesAsDataDepth() {
        var inner = new AdfNode("paragraph") { Content = { AdfNode.TextNode("Deep") } };
        for (int index = 0; index < 30; index++) {
            var item = new AdfNode("listItem") { Content = { new AdfNode("paragraph") } };
            if (inner.Type == "paragraph") item.Content[0] = inner;
            else item.Content.Add(inner);
            inner = new AdfNode("bulletList") { Content = { item } };
        }
        var document = new AdfDocument(new[] { inner });
        var options = new AdfValidationOptions { Profile = AdfValidationProfile.FullSchema, MaximumSchemaEvaluations = 10000 };
        AdfValidationResult result = document.Validate(options);
        Assert.True(result.IsValid, string.Join("; ", result.Issues.Select(issue => issue.Path + ": " + issue.Message)));
        // An invalid leaf follows the same selected branches and reports its real location.
        var invalid = new AdfNode("vendorFuture");
        AdfNode current = inner;
        while (current.Content[0].Content.Count > 1) current = current.Content[0].Content[1];
        current.Content[0].Content.Add(invalid);
        result = document.Validate(options);
        Assert.False(result.IsValid);
        Assert.DoesNotContain(result.Issues, issue => issue.Code == "ADF_SCHEMA_LIMIT");
        Assert.Contains(result.Issues, issue => issue.Path.EndsWith(".type", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData(AdfValidationProfile.ForwardCompatible)]
    [InlineData(AdfValidationProfile.FullSchema)]
    public void ValidationHonorsCancellationAndSchemaBudgetFailsExplicitly(AdfValidationProfile profile) {
        var document = new AdfDocument(new[] { new AdfNode("paragraph") { Content = { AdfNode.TextNode("Visible") } } });
        var options = new AdfValidationOptions { Profile = profile };
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => document.Validate(options, cancellation.Token));
        if (profile == AdfValidationProfile.FullSchema) {
            options.MaximumSchemaEvaluations = 1;
            AdfValidationResult result = document.Validate(options);
            Assert.False(result.IsValid);
            Assert.Contains(result.Issues, issue => issue.Code == "ADF_SCHEMA_LIMIT");
        }
    }
    [Theory]
    [InlineData(AdfValidationProfile.ForwardCompatible)]
    [InlineData(AdfValidationProfile.FullSchema)]
    public void MutableCyclesAreReportedAndRejectedByAllRecursiveEntryPoints(AdfValidationProfile profile) {
        var node = new AdfNode("vendor-wrapper"); node.Content.Add(node);
        var document = new AdfDocument(new[] { node });
        AdfValidationResult result = document.Validate(new AdfValidationOptions { Profile = profile });
        Assert.False(result.IsValid);
        Assert.Contains(result.Issues, issue => issue.Code == "ADF_CYCLIC_CONTENT" && issue.Path == "$.content[0].content[0]");
        Assert.Throws<InvalidDataException>(() => document.ToJson());
        Assert.Throws<InvalidDataException>(() => AdfConverter.ToMarkdown(document));
    }

    [Fact]
    public void AuthoringDepthIsBoundedAndSharedSiblingsRemainSerializable() {
        var first = new AdfNode("vendor-wrapper");
        var current = first;
        for (int index = 0; index < 64; index++) { var next = new AdfNode("vendor-wrapper"); current.Content.Add(next); current = next; }
        var document = new AdfDocument(new[] { first });
        Assert.Contains(document.Validate().Issues, issue => issue.Code == "ADF_CONTENT_DEPTH_EXCEEDED");
        Assert.Throws<InvalidDataException>(() => document.ToJson());
        var paragraph = new AdfNode("paragraph") { Content = { AdfNode.TextNode("Shared") } };
        var shared = new AdfDocument(new[] { paragraph, paragraph });
        Assert.True(shared.Validate().IsValid);
        Assert.Equal(2, AdfDocument.Parse(shared.ToJson()).Content.Count);
    }

    [Fact]
    public void ExplicitEmptyFieldsAndRequiredEmptyAttributeObjectsSurviveNativeRoundTrips() {
        const string source = "{\"type\":\"doc\",\"version\":1,\"content\":[{\"type\":\"expand\",\"content\":[{\"type\":\"nestedExpand\",\"attrs\":{},\"content\":[{\"type\":\"paragraph\",\"content\":[],\"marks\":[]}]}]}]}";
        AdfDocument document = AdfDocument.Parse(source);
        Assert.True(document.Validate(new AdfValidationOptions { Profile = AdfValidationProfile.FullSchema }).IsValid);
        AdfDocument reopened = AdfDocument.Parse(document.ToJson());
        Assert.True(reopened.Validate(new AdfValidationOptions { Profile = AdfValidationProfile.FullSchema }).IsValid);
        using JsonDocument json = JsonDocument.Parse(reopened.ToJson());
        JsonElement nested = json.RootElement.GetProperty("content")[0].GetProperty("content")[0];
        Assert.Equal(JsonValueKind.Object, nested.GetProperty("attrs").ValueKind);
        Assert.Equal(0, nested.GetProperty("content")[0].GetProperty("content").GetArrayLength());
        Assert.Equal(0, nested.GetProperty("content")[0].GetProperty("marks").GetArrayLength());
    }

    [Fact]
    public void NumericallyEquivalentEnumValuesPassThePinnedJsonSchema() {
        var heading = new AdfNode("heading").SetAttribute("level", 1.0m);
        heading.Content.Add(AdfNode.TextNode("Heading"));
        var document = new AdfDocument(new[] { heading });
        Assert.True(document.Validate().IsValid);
        Assert.Equal(1, heading.GetInt32Attribute("level"));
        Assert.True(document.Validate(new AdfValidationOptions { Profile = AdfValidationProfile.FullSchema }).IsValid);
        Assert.Equal(1, AdfDocument.Parse("{\"type\":\"doc\",\"version\":1.0,\"content\":[]}").Version);
    }
}
