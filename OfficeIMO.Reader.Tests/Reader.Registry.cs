using OfficeIMO.Reader;
using OfficeIMO.Reader.Json;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.Reader.PowerPoint;
using OfficeIMO.Word;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class ReaderRegistryTests {
    [Fact]
    public void DocumentReader_ExposesOnlyOfficeIMOCapabilities() {
        IReadOnlyList<ReaderHandlerCapability> capabilities = OfficeIMO.Reader.Tests.ReaderTestReaders.All.GetCapabilities();

        Assert.NotEmpty(capabilities);
        Assert.All(capabilities, capability => {
            Assert.Equal(ReaderCapabilitySchema.Id, capability.SchemaId);
            Assert.Equal(ReaderCapabilitySchema.Version, capability.SchemaVersion);
            Assert.Equal(ReaderHandlerOrigin.OfficeIMO, capability.Origin);
        });
        Assert.Contains(capabilities, capability => capability.Id == "officeimo.reader.word");
        Assert.Contains(capabilities, capability => capability.Id == "officeimo.reader.excel");
        Assert.Contains(capabilities, capability => capability.Id == "officeimo.reader.powerpoint");
        Assert.Contains(capabilities, capability => capability.Id == "officeimo.reader.markdown");
        Assert.Contains(capabilities, capability => capability.Id == "officeimo.reader.pdf");
        Assert.Contains(capabilities, capability => capability.Id == "officeimo.reader.text");
    }

    [Fact]
    public void DocumentReader_PowerPointCapabilityIncludesBinaryVariants() {
        ReaderHandlerCapability openXml = Assert.Single(
            OfficeIMO.Reader.Tests.ReaderTestReaders.All.GetCapabilities(), item =>
                item.Id == "officeimo.reader.powerpoint");
        ReaderHandlerCapability binary = Assert.Single(
            OfficeIMO.Reader.Tests.ReaderTestReaders.All.GetCapabilities(), item =>
                item.Id == "officeimo.reader.powerpoint.binary");

        Assert.Equal(ReaderInputKind.PowerPoint, openXml.Kind);
        Assert.Equal(
            PowerPointFormatCatalog.All
                .Where(format => format.Generation == OfficeFormatGeneration.Modern)
                .Select(format => format.Extension)
                .OrderBy(extension => extension, StringComparer.Ordinal)
                .ToArray(),
            openXml.Extensions);
        Assert.True(openXml.SupportsPath);
        Assert.True(openXml.SupportsStream);
        Assert.Equal(PowerPointReaderAdapter.DefaultModernMaxInputBytes, openXml.DefaultMaxInputBytes);
        Assert.Equal(ReaderInputKind.PowerPoint, binary.Kind);
        Assert.Equal(
            PowerPointFormatCatalog.All
                .Where(format => format.Generation == OfficeFormatGeneration.Legacy)
                .Select(format => format.Extension)
                .OrderBy(extension => extension, StringComparer.Ordinal)
                .ToArray(),
            binary.Extensions);
        Assert.True(binary.SupportsPath);
        Assert.True(binary.SupportsStream);
        Assert.Equal(LegacyPptImportOptions.DefaultMaxInputBytes,
            binary.DefaultMaxInputBytes);
        Assert.Equal(PowerPointReaderAdapter.DefaultModernMaxInputBytes,
            OfficeIMO.Reader.Tests.ReaderTestReaders.All.GetHandlerDefaultMaxInputBytes("deck.pptx"));
        Assert.Equal(LegacyPptImportOptions.DefaultMaxInputBytes,
            OfficeIMO.Reader.Tests.ReaderTestReaders.All.GetHandlerDefaultMaxInputBytes(
                "deck.ppt"));
    }

    [Fact]
    public void PowerPointReader_UnknownStreamNameKeepsLegacyParserAtLegacyLimit() {
        PowerPointLoadOptions options = PowerPointReaderAdapter.CreateLoadOptions(
            new ReaderOptions(),
            sourceName: null);

        Assert.Equal(PowerPointReaderAdapter.DefaultModernMaxInputBytes, options.PackageSecurity!.MaxPackageBytes);
        Assert.Equal(LegacyPptImportOptions.DefaultMaxInputBytes, options.LegacyPptImportOptions!.MaxInputBytes);
    }

    [Fact]
    public void DocumentReader_WordAndExcelCapabilitiesMatchFormatCatalogs() {
        ReaderHandlerCapability word = Assert.Single(
            OfficeIMO.Reader.Tests.ReaderTestReaders.All.GetCapabilities(), item => item.Id == "officeimo.reader.word");
        ReaderHandlerCapability excel = Assert.Single(
            OfficeIMO.Reader.Tests.ReaderTestReaders.All.GetCapabilities(), item => item.Id == "officeimo.reader.excel");

        Assert.Equal(
            WordFormatCatalog.All.Select(format => format.Extension).OrderBy(extension => extension, StringComparer.Ordinal).ToArray(),
            word.Extensions);
        Assert.Equal(
            ExcelFormatCatalog.All.Select(format => format.Extension).OrderBy(extension => extension, StringComparer.Ordinal).ToArray(),
            excel.Extensions);
    }

    [Fact]
    public void DocumentReader_BoundsUnidentifiedStreamsBeforeDetection() {
        Assert.Equal(64L * 1024L * 1024L,
            DocumentReaderEngine.ResolveStreamMaxInputBytes(null,
                new ReaderOptions(), streamCanSeek: false));
        Assert.Equal(64L * 1024L * 1024L,
            DocumentReaderEngine.ResolveStreamMaxInputBytes("content.bin",
                new ReaderOptions(), streamCanSeek: false));
        Assert.Null(DocumentReaderEngine.ResolveStreamMaxInputBytes(
            "content.bin", new ReaderOptions(), streamCanSeek: true));
        Assert.Equal(64L * 1024L * 1024L, DocumentReaderEngine.ResolveStreamMaxInputBytes(
            "document.docx", new ReaderOptions(), streamCanSeek: false));
    }

    [Fact]
    public void DocumentReader_CapabilityManifestJson_IsDeterministicAndValid() {
        string first = OfficeIMO.Reader.Tests.ReaderTestReaders.All.GetCapabilityManifestJson();
        string second = OfficeIMO.Reader.Tests.ReaderTestReaders.All.GetCapabilityManifestJson();

        Assert.Equal(first, second);
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(first));
        ReaderChunk[] chunks = JsonReaderAdapter.Read(
            stream,
            sourceName: "capability-manifest.json",
            jsonOptions: new JsonReadOptions { ChunkRows = 128, IncludeMarkdown = false })
            .ToArray();

        Assert.NotEmpty(chunks);
        Assert.DoesNotContain(chunks, chunk =>
            chunk.Warnings?.Any(warning => warning.Contains("JSON parse error", StringComparison.OrdinalIgnoreCase)) == true);
        Assert.Contains("\"schemaId\":\"officeimo.reader.capability\"", first, StringComparison.Ordinal);
        Assert.Contains("\"origin\":\"OfficeIMO\"", first, StringComparison.Ordinal);
        Assert.DoesNotContain("isBuiltIn", first, StringComparison.Ordinal);
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
            Assert.Equal(ReaderInputKind.Unknown, OfficeIMO.Reader.Tests.ReaderTestReaders.All.DetectKind(path));
            Assert.Contains(reader.GetCapabilities(), capability => capability.Id == handlerId && capability.Origin == ReaderHandlerOrigin.Custom);
            Assert.DoesNotContain(OfficeIMO.Reader.Tests.ReaderTestReaders.All.GetCapabilities(), capability => capability.Id == handlerId);
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

        Assert.Contains(manifest.Handlers, capability => capability.Id == handlerId && capability.SupportsStream && capability.Origin == ReaderHandlerOrigin.Custom);
        Assert.Contains(handlerId, reader.GetCapabilityManifestJson(), StringComparison.Ordinal);
    }
}
