using OfficeIMO.Reader;
using OfficeIMO.Reader.Chm;
using OfficeIMO.Reader.Html;
using OfficeIMO.Html;
using System.Security.Cryptography;

namespace OfficeIMO.Chm.Tests;

public sealed class ReaderTests {
    [Theory]
    [InlineData("characters")]
    [InlineData("nodes")]
    [InlineData("depth")]
    public void ReaderPreservesStricterCallerHtmlLimits(string boundary) {
        var html = new HtmlConversionDocumentOptions();
        if (boundary == "characters") html.Limits.MaxInputCharacters = 20;
        if (boundary == "nodes") html.Limits.MaxHtmlNodes = 3;
        if (boundary == "depth") html.Limits.MaxHtmlDepth = 2;
        var options = new ReaderChmOptions { HtmlOptions = new ReaderHtmlOptions { ConversionOptions = html } };
        var reader = new OfficeDocumentReaderBuilder().AddChmHandler(options).Build();
        using var stream = new MemoryStream(ChmFixture.Book());
        Assert.Throws<HtmlDomLimitException>(() => reader.ReadDocument(stream, "manual.chm"));
        if (boundary == "characters") Assert.Equal(20, html.Limits.MaxInputCharacters);
        if (boundary == "nodes") Assert.Equal(3, html.Limits.MaxHtmlNodes);
        if (boundary == "depth") Assert.Equal(2, html.Limits.MaxHtmlDepth);
        Assert.DoesNotContain("chm", html.UrlPolicy.AllowedUrlSchemes);
    }

    [Fact]
    public void ReaderRetainsCallerUrlPolicyAndArchiveRelativeLinks() {
        var html = new HtmlConversionDocumentOptions();
        html.UrlPolicy.AllowedUrlSchemes.Clear();
        var options = new ReaderChmOptions { HtmlOptions = new ReaderHtmlOptions { ConversionOptions = html } };
        var reader = new OfficeDocumentReaderBuilder().AddChmHandler(options).Build();
        using var stream = new MemoryStream(ChmFixture.Archive(new Dictionary<string, byte[]> {
            ["/topic.html"] = ChmFixture.Html("<a href='https://example.invalid/'>External</a><a href='other.html'>Internal</a>"),
            ["/other.html"] = ChmFixture.Html("<p>Other</p>")
        }));
        var result = reader.ReadDocument(stream, "manual.chm");
        Assert.DoesNotContain("https://example.invalid", result.Markdown!);
        Assert.Contains("chm://archive/other.html", result.Markdown!);
        Assert.Empty(html.UrlPolicy.AllowedUrlSchemes);
    }
    [Fact]
    public void ReaderRetainsTopicCitationsRichTablesAndArchiveHash() {
        byte[] bytes = ChmFixture.Book(); using var stream = new MemoryStream(bytes); stream.Position = 11;
        var reader = new OfficeDocumentReaderBuilder().AddChmHandler().Build();
        var result = reader.ReadDocument(stream, "manual.chm", new ReaderOptions { ComputeHashes = true });
        Assert.Equal(ReaderInputKind.Chm, result.Kind); Assert.Equal(11, stream.Position); Assert.True(stream.CanRead);
        using var sha = SHA256.Create(); string hash = BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", "").ToLowerInvariant();
        Assert.Equal(hash, result.Source.SourceHash); Assert.Equal(bytes.Length, result.Source.LengthBytes);
        Assert.All(result.Chunks, chunk => { Assert.Equal(ReaderInputKind.Chm, chunk.Kind); Assert.Equal(hash, chunk.SourceHash); Assert.StartsWith("manual.chm!/", chunk.Location.Path); });
        Assert.Contains(result.Chunks, chunk => chunk.Location.Path == "manual.chm!/guide/details.html" && chunk.Text.Contains("42"));
        Assert.NotEmpty(result.Blocks); Assert.NotEmpty(result.Tables); Assert.NotEmpty(result.Links);
        Assert.Equal(result.Blocks.Count, result.Blocks.Select(block => block.Id).Distinct().Count());
        Assert.Equal(ReaderInputKind.Chm, OfficeDocumentReadResultJson.Deserialize(OfficeDocumentReadResultJson.Serialize(result)).Kind);
        var oldEnvelope = System.Text.Json.Nodes.JsonNode.Parse(OfficeDocumentReadResultJson.Serialize(result))!;
        oldEnvelope["schemaVersion"] = 10;
        string oldBinding = oldEnvelope.ToJsonString();
        Assert.Throws<System.Text.Json.JsonException>(() => OfficeDocumentReadResultJson.Deserialize(oldBinding));
    }

    [Fact]
    public void DetectionFindsItsfSignatureWithUnknownExtensionAndPreservesPosition() {
        var reader = new OfficeDocumentReaderBuilder().AddChmHandler().Build();
        using var stream = new MemoryStream(ChmFixture.Book());
        var detected = reader.Detect(stream, "manual.bin", new ReaderDetectionOptions { Mode = ReaderDetectionMode.PreferContent });
        Assert.Equal(ReaderInputKind.Chm, detected.Kind); Assert.Equal("application/vnd.ms-htmlhelp", detected.MediaType); Assert.Equal(0, stream.Position);
    }

    [Fact]
    public void RegistrationSnapshotsOptionsAndHonorsReaderInputLimit() {
        var options = new ReaderChmOptions { ConversionOptions = new ChmConversionOptions { TopicPaths = new[] { "/welcome.html" } } };
        var reader = new OfficeDocumentReaderBuilder().AddChmHandler(options).Build();
        options.ConversionOptions.TopicPaths = new[] { "/missing.html" };
        using var stream = new MemoryStream(ChmFixture.Book());
        var result = reader.ReadDocument(stream, "manual.chm"); Assert.Contains("Café", result.Markdown!); Assert.DoesNotContain("Second topic", result.Markdown!);
        Assert.ThrowsAny<IOException>(() => reader.ReadDocument(stream, "manual.chm", new ReaderOptions { MaxInputBytes = 20 }));
    }
}
