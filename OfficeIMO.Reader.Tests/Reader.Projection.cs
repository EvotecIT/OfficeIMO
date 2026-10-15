using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using OfficeIMO.Reader;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderProjectionTests {
    private static ReaderOptions Lean() => new() { Projection = ReaderProjection.TextAndMetadata, ComputeHashes = false };

    [Fact]
    public void LeanTextKeepsContentWithoutBuildingMarkdownOrTableProjection() {
        const string content = "| Name | Value |\n| --- | --- |\n| first | second |";
        var reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().Build();
        byte[] input = Encoding.UTF8.GetBytes(content);
        OfficeDocumentReadResult full = reader.ReadDocument(input, "source.txt");
        OfficeDocumentReadResult lean = reader.ReadDocument(input, "source.txt", Lean());

        Assert.Equal(content, Assert.Single(lean.Chunks).Text);
        Assert.Equal(full.Chunks.Select(chunk => chunk.Text), lean.Chunks.Select(chunk => chunk.Text));
        Assert.NotNull(full.Markdown);
        Assert.Null(lean.Markdown);
        Assert.Empty(lean.Tables);
        Assert.Empty(lean.Assets);
        Assert.Empty(lean.Visuals);
        Assert.Empty(lean.OcrCandidates);
        Assert.Contains("officeimo.reader.projection.text-and-metadata", lean.CapabilitiesUsed);
    }

    [Theory]
    [InlineData("document")]
    [InlineData("async")]
    [InlineData("incremental")]
    public async Task UnsupportedProjectionRejectsBeforeInvokingHandler(string route) {
        int calls = 0;
        var reader = new OfficeDocumentReaderBuilder().AddHandler(new ReaderHandlerRegistration {
            Id = "unsupported", Extensions = new[] { ".guard" }, Kind = ReaderInputKind.Text,
            SupportsIncrementalStream = true,
            ReadStream = (_, _, _, _) => { calls++; return new[] { new ReaderChunk { Text = "content" } }; }
        }).Build();
        using var input = new MemoryStream(Encoding.UTF8.GetBytes("content"));
        if (route == "async") {
            await Assert.ThrowsAsync<NotSupportedException>(() => reader.ReadDocumentAsync(input, "source.guard", Lean()));
        } else if (route == "incremental") {
            Assert.Throws<NotSupportedException>(() => reader.EnumerateChunks(input, "source.guard", Lean()).ToArray());
        } else {
            Assert.Throws<NotSupportedException>(() => reader.ReadDocument(input, "source.guard", Lean()));
        }
        Assert.Equal(0, calls);
        Assert.True(input.CanRead);
        Assert.Single(reader.ReadDocument(input, "source.guard").Chunks);
        Assert.Equal(1, calls);
    }

    [Fact]
    public void ProjectionAndHandlerSupportAreSnapshottedBeforeEnumeration() {
        ReaderProjection observed = ReaderProjection.Full;
        long? observedPartLimit = null;
        long? observedPackageLimit = null;
        var registration = new ReaderHandlerRegistration {
            Id = "snapshot", Extensions = new[] { ".snapshot" }, Kind = ReaderInputKind.Text,
            SupportsTextAndMetadata = true, SupportsIncrementalStream = true,
            ReadStream = (_, _, options, _) => {
                observed = options.Projection;
                observedPartLimit = options.OpenXmlMaxPartUncompressedBytes;
                observedPackageLimit = options.OpenXmlMaxTotalUncompressedBytes;
                return new[] { new ReaderChunk { Text = "content" } };
            }
        };
        var reader = new OfficeDocumentReaderBuilder().AddHandler(registration).Build();
        registration.SupportsTextAndMetadata = false;
        var options = Lean();
        options.OpenXmlMaxPartUncompressedBytes = 1024;
        options.OpenXmlMaxTotalUncompressedBytes = 4096;
        using var input = new MemoryStream(Encoding.UTF8.GetBytes("content"));
        var chunks = reader.EnumerateChunks(input, "source.snapshot", options);
        options.Projection = ReaderProjection.Full;
        options.OpenXmlMaxPartUncompressedBytes = null;
        options.OpenXmlMaxTotalUncompressedBytes = null;
        Assert.Single(chunks);
        Assert.Equal(ReaderProjection.TextAndMetadata, observed);
        Assert.Equal(1024L, observedPartLimit);
        Assert.Equal(4096L, observedPackageLimit);
        Assert.True(Assert.Single(reader.GetCapabilities()).SupportsTextAndMetadata);
        using JsonDocument capabilities = JsonDocument.Parse(reader.GetCapabilityManifestJson());
        Assert.True(capabilities.RootElement.GetProperty("handlers")[0].GetProperty("supportsTextAndMetadata").GetBoolean());
    }

    [Fact]
    public void NestedReadCannotReenableAnUnsupportedFullProjection() {
        int childCalls = 0;
        var child = new OfficeDocumentReaderBuilder().AddHandler(new ReaderHandlerRegistration {
            Id = "child", Extensions = new[] { ".child" }, Kind = ReaderInputKind.Text,
            ReadStream = (_, _, _, _) => { childCalls++; return Array.Empty<ReaderChunk>(); }
        }).Build();
        var parent = new OfficeDocumentReaderBuilder().AddHandler(new ReaderHandlerRegistration {
            Id = "parent", Extensions = new[] { ".parent" }, Kind = ReaderInputKind.Text,
            SupportsTextAndMetadata = true,
            ReadDocumentStream = (_, _, _, _) => {
                child.ReadDocument(Encoding.UTF8.GetBytes("child"), "source.child", new ReaderOptions());
                return new OfficeDocumentReadResult();
            }
        }).Build();
        Assert.Throws<NotSupportedException>(() => parent.ReadDocument(Encoding.UTF8.GetBytes("parent"), "source.parent", Lean()));
        Assert.Equal(0, childCalls);
    }

    [Fact]
    public void TextProcessorDoesNotRegenerateUnrequestedMarkdown() {
        var reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers()
            .AddProcessor(new DelegateOfficeDocumentProcessor("edit", (document, _) => {
                document.Chunks[0].Text = "edited";
                return document;
            })).Build();
        OfficeDocumentReadResult result = reader.ReadDocument(Encoding.UTF8.GetBytes("original"), "source.txt", Lean());
        Assert.Equal("edited", Assert.Single(result.Chunks).Text);
        Assert.Equal("edited", Assert.Single(result.Blocks).Text);
        Assert.Null(result.Markdown);
        Assert.Contains("officeimo.reader.projection.text-and-metadata", result.CapabilitiesUsed);
    }

    [Fact]
    public void AHandlerCanQualifyOnlyItsModernFormatForLeanProjection() {
        int calls = 0;
        string[] supported = { ".Modern" };
        var reader = new OfficeDocumentReaderBuilder().AddHandler(new ReaderHandlerRegistration {
            Id = "mixed", Extensions = new[] { ".modern", ".legacy" }, Kind = ReaderInputKind.Text,
            SupportsTextAndMetadata = true, TextAndMetadataExtensions = supported,
            ReadStream = (_, _, _, _) => { calls++; return new[] { new ReaderChunk { Text = "content" } }; }
        }).Build();
        supported[0] = ".legacy";
        byte[] bytes = Encoding.UTF8.GetBytes("content");
        Assert.Throws<NotSupportedException>(() => reader.ReadDocument(bytes, "source.legacy", Lean()));
        Assert.Equal(0, calls);
        Assert.Single(reader.ReadDocument(bytes, "source.modern", Lean()).Chunks);
        Assert.Equal(new[] { ".modern" }, Assert.Single(reader.GetCapabilities()).TextAndMetadataExtensions);
        Assert.Single(reader.ReadDocument(bytes, "source.legacy").Chunks);
        Assert.Equal(2, calls);
    }

    [Fact]
    public void UnknownProjectionIsRejectedBeforeReadingInput() {
        var reader = new OfficeDocumentReaderBuilder().AddPlainTextHandlers().Build();
        Assert.Throws<ArgumentOutOfRangeException>(() => reader.ReadDocument(Array.Empty<byte>(), "source.txt",
            new ReaderOptions { Projection = (ReaderProjection)99 }));
    }
}
