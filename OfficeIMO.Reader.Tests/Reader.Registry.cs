using OfficeIMO.Reader;
using OfficeIMO.Reader.Json;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class ReaderRegistryTests {
    [Fact]
    public void DocumentReader_ExposesOnlyBuiltInCapabilities() {
        IReadOnlyList<ReaderHandlerCapability> capabilities = DocumentReader.GetCapabilities();

        Assert.NotEmpty(capabilities);
        Assert.All(capabilities, capability => {
            Assert.True(capability.IsBuiltIn);
            Assert.Equal(ReaderCapabilitySchema.Id, capability.SchemaId);
            Assert.Equal(ReaderCapabilitySchema.Version, capability.SchemaVersion);
        });
        Assert.Contains(capabilities, capability => capability.Id == "officeimo.reader.word");
        Assert.Contains(capabilities, capability => capability.Id == "officeimo.reader.excel");
        Assert.Contains(capabilities, capability => capability.Id == "officeimo.reader.powerpoint");
        Assert.Contains(capabilities, capability => capability.Id == "officeimo.reader.markdown");
        Assert.Contains(capabilities, capability => capability.Id == "officeimo.reader.pdf");
        Assert.Contains(capabilities, capability => capability.Id == "officeimo.reader.text");
    }

    [Fact]
    public void DocumentReader_CapabilityManifestJson_IsDeterministicAndValid() {
        string first = DocumentReader.GetCapabilityManifestJson();
        string second = DocumentReader.GetCapabilityManifestJson();

        Assert.Equal(first, second);
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(first));
        ReaderChunk[] chunks = DocumentReaderJsonExtensions.ReadJson(
            stream,
            sourceName: "capability-manifest.json",
            jsonOptions: new JsonReadOptions { ChunkRows = 128, IncludeMarkdown = false })
            .ToArray();

        Assert.NotEmpty(chunks);
        Assert.DoesNotContain(chunks, chunk =>
            chunk.Warnings?.Any(warning => warning.Contains("JSON parse error", StringComparison.OrdinalIgnoreCase)) == true);
        Assert.Contains("\"schemaId\":\"officeimo.reader.capability\"", first, StringComparison.Ordinal);
        Assert.Contains("\"supportsDocumentPath\":", first, StringComparison.Ordinal);
        Assert.Contains("\"supportsAsyncStream\":", first, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeDocumentReader_UsesBuilderHandlerWithoutChangingStaticReader() {
        const string extension = ".builderix";
        const string handlerId = "officeimo.tests.builder";
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + extension);

        try {
            File.WriteAllText(path, "input");
            OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
                .AddHandler(new ReaderHandlerRegistration {
                    Id = handlerId,
                    Kind = ReaderInputKind.Text,
                    Extensions = new[] { extension },
                    ReadPath = (sourcePath, options, cancellationToken) => new[] {
                        new ReaderChunk { Id = "builder-1", Kind = ReaderInputKind.Text, Text = "builder-output" }
                    }
                })
                .Build();

            Assert.Equal("builder-output", Assert.Single(reader.Read(path)).Text);
            Assert.Equal(ReaderInputKind.Unknown, DocumentReader.DetectKind(path));
            Assert.Contains(reader.GetCapabilities(), capability => capability.Id == handlerId && !capability.IsBuiltIn);
            Assert.DoesNotContain(DocumentReader.GetCapabilities(), capability => capability.Id == handlerId);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void OfficeDocumentReader_CapabilityManifestIncludesConfiguredHandlers() {
        const string handlerId = "officeimo.tests.manifest";
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddHandler(new ReaderHandlerRegistration {
                Id = handlerId,
                Kind = ReaderInputKind.Text,
                Extensions = new[] { ".manifestix" },
                ReadStream = (stream, sourceName, options, cancellationToken) => Array.Empty<ReaderChunk>()
            })
            .Build();

        ReaderCapabilityManifest manifest = reader.GetCapabilityManifest();

        Assert.Contains(manifest.Handlers, capability => capability.Id == handlerId && capability.SupportsStream);
        Assert.Contains(handlerId, reader.GetCapabilityManifestJson(), StringComparison.Ordinal);
    }
}
