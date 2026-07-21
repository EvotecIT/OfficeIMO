using System.Text.Json;
using OfficeIMO.Adf;
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
    }

    [Fact]
    public void Validation_WarnsForUnknownNodesWithoutRejectingRoundTrip() {
        AdfDocument document = AdfDocument.Parse("{\"version\":1,\"type\":\"doc\",\"content\":[{\"type\":\"futureNode\",\"attrs\":{\"a\":1}}]}");

        AdfValidationResult result = document.Validate();

        Assert.True(result.IsValid);
        Assert.Contains(result.Issues, item => item.Code == "ADF_UNKNOWN_NODE");
    }
}
