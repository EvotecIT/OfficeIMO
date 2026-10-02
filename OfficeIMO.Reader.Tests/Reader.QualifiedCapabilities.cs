using OfficeIMO.Reader;
using OfficeIMO.Reader.Word;
using System.Text;
using System.Text.Json;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderQualifiedCapabilityTests {
    [Fact]
    public void FormatProfilesReuseOwnerIdsAndRemainImmutable() {
        var preservation = new[] { "text" };
        var profiles = new[] { new ReaderFormatQualification("test", "custom.test", ReaderFormatSupport.SalvageRead,
            preservation: preservation, limitations: new[] { "No images" }, evidence: new[] { "Fixture set 1" }) };
        var builder = new OfficeDocumentReaderBuilder().AddHandler(new ReaderHandlerRegistration {
            Id = "custom", Extensions = new[] { ".test", ".unqualified" }, FormatQualifications = profiles,
            ReadStream = (_, _, _, _) => Array.Empty<ReaderChunk>()
        }).AddWordHandler();
        preservation[0] = "mutated";
        profiles[0] = new ReaderFormatQualification("test", "changed");
        var reader = builder.Build();
        var custom = reader.GetCapabilities().Single(c => c.Id == "custom");
        Assert.Equal("text", custom.FormatQualifications.Single(p => p.Extension == ".test").Preservation[0]);
        Assert.Equal(ReaderFormatSupport.Unqualified, custom.FormatQualifications.Single(p => p.Extension == ".unqualified").Support);
        var word = reader.GetCapabilities().Single(c => c.Id == OfficeDocumentReaderBuilderWordExtensions.HandlerId);
        Assert.Equal(global::OfficeIMO.Word.WordFormatCatalog.All.Select(f => f.Id).OrderBy(id => id, StringComparer.Ordinal), word.FormatQualifications.Select(f => f.FormatId).OrderBy(id => id, StringComparer.Ordinal));
        using var json = JsonDocument.Parse(reader.GetCapabilityManifestJson());
        Assert.Equal(6, json.RootElement.GetProperty("schemaVersion").GetInt32());
        Assert.Contains(json.RootElement.GetProperty("handlers").EnumerateArray(), h => h.GetProperty("formatQualifications").GetArrayLength() > 0);
    }

    [Fact]
    public void ProfilesCannotClaimAnUnregisteredExtension() {
        Assert.Throws<ArgumentException>(() => new OfficeDocumentReaderBuilder().AddHandler(new ReaderHandlerRegistration {
            Id = "custom", Extensions = new[] { ".test" },
            FormatQualifications = new[] { new ReaderFormatQualification(".other", "other") },
            ReadStream = (_, _, _, _) => Array.Empty<ReaderChunk>()
        }));
    }

    [Theory]
    [InlineData("漢字かな交じり文한국어文本")]
    [InlineData("مرحبا بالعالم שלום עולם")]
    [InlineData("école 👩🏽‍💻🙂 café")]
    public void TokenizerBridgePreservesUnicodeAndHonorsActualCounts(string sample) {
        string text = string.Concat(Enumerable.Repeat(sample, 20));
        int rangeCalls = 0;
        var counter = new ReaderDelegateTokenCounter("application.scalar-v1", value => value.EnumerateRunes().Count(),
            (prefix, source, start, length) => { rangeCalls++; return prefix.EnumerateRunes().Count() + source.Substring(start, length).EnumerateRunes().Count(); });
        var result = ReaderHierarchicalChunker.Chunk(new[] { new ReaderChunk { Id = "source", Text = text } },
            new ReaderHierarchicalChunkingOptions { MaxTokens = 8, OverlapTokens = 0, IncludeContextInText = false, TokenCounter = counter });
        Assert.True(rangeCalls > 0);
        Assert.Equal(text, string.Concat(result.Chunks.Select(c => c.Text)));
        Assert.All(result.Chunks, chunk => {
            Assert.InRange(counter.CountTokens(chunk.Text), 1, 8);
            Assert.Equal(counter.CountTokens(chunk.Text), chunk.TokenEstimate);
            Assert.Equal(chunk.Text, new UTF8Encoding(false, true).GetString(new UTF8Encoding(false, true).GetBytes(chunk.Text)));
        });
        Assert.Equal(counter.Id, result.TokenCounterId);
    }
}
