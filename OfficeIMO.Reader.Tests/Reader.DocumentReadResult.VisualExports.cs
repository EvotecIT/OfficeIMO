using OfficeIMO.Reader;
using System.Text.Json;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderDocumentReadResultVisualExportTests {
    [Fact]
    public void DocumentReader_ReadVisualExports_ReturnsDeterministicSidecarPayloads() {
        byte[] bytes = Encoding.UTF8.GetBytes("# Diagram\n\n```mermaid\ngraph TD\nA-->B\n```\n");
        using var stream = new MemoryStream(bytes, writable: false);

        IReadOnlyList<ReaderVisualExportBundle> exports = OfficeIMO.Reader.Tests.ReaderTestReaders.All.ReadVisualExports(
            stream,
            "diagram.md",
            indentedJson: true);

        ReaderVisualExportBundle export = Assert.Single(exports);
        Assert.Equal("diagram-visual-0000", export.Id);
        Assert.Equal("diagram-visual-0000", export.FileNamePrefix);
        Assert.Equal(".mmd", export.PayloadExtension);
        Assert.Equal("graph TD\nA-->B", export.Payload);
        Assert.Equal("diagram.md", export.Visual.Location?.Path);
        using JsonDocument document = JsonDocument.Parse(export.Json);
        JsonElement root = document.RootElement;
        Assert.Equal("mermaid", root.GetProperty("kind").GetString());
        Assert.Equal("mermaid", root.GetProperty("language").GetString());
        Assert.Equal("graph TD\nA-->B", root.GetProperty("content").GetString());
        Assert.Equal("diagram.md", root.GetProperty("location").GetProperty("path").GetString());
        Assert.Equal("diagram--code-1", root.GetProperty("location").GetProperty("blockAnchor").GetString());
    }

    [Fact]
    public void ReaderVisualExportMaterializer_WriteVisualExportsToDirectory_WritesPayloadAndJsonSidecars() {
        byte[] bytes = Encoding.UTF8.GetBytes("# Diagram\n\n```mermaid\ngraph TD\nA-->B\n```\n");
        using var stream = new MemoryStream(bytes, writable: false);
        IReadOnlyList<ReaderVisualExportBundle> exports = OfficeIMO.Reader.Tests.ReaderTestReaders.All.ReadVisualExports(stream, "diagram.md");
        var directory = Path.Combine(Path.GetTempPath(), "officeimo-reader-visuals-" + Guid.NewGuid().ToString("N"));

        try {
            IReadOnlyList<ReaderVisualMaterializedExport> materialized = exports.WriteVisualExportsToDirectory(directory);

            Assert.Equal(2, materialized.Count);
            Assert.All(materialized, item => Assert.True(item.Written));
            Assert.True(File.Exists(Path.Combine(directory, "diagram-visual-0000.mmd")));
            Assert.True(File.Exists(Path.Combine(directory, "diagram-visual-0000.json")));
            Assert.Equal("graph TD\nA-->B", File.ReadAllText(Path.Combine(directory, "diagram-visual-0000.mmd")));
            using JsonDocument document = JsonDocument.Parse(File.ReadAllText(Path.Combine(directory, "diagram-visual-0000.json")));
            Assert.Equal("mermaid", document.RootElement.GetProperty("kind").GetString());

            IReadOnlyList<ReaderVisualMaterializedExport> skipped = exports.WriteVisualExportsToDirectory(
                directory,
                new ReaderVisualExportMaterializationOptions { Overwrite = false });
            Assert.Equal(2, skipped.Count);
            Assert.All(skipped, item => {
                Assert.False(item.Written);
                Assert.Equal("Destination file already exists.", item.SkippedReason);
            });
        } finally {
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void ReaderVisualExportMaterializer_StreamVisualExports_StreamsSelectedPayloadsThroughCallback() {
        var export = new ReaderVisualExportBundle {
            Id = "manual-visual",
            FileNamePrefix = "manual-visual",
            PayloadExtension = ".mmd",
            Payload = "graph TD\nA-->B",
            Json = "{\"kind\":\"mermaid\"}"
        };
        string captured = string.Empty;

        IReadOnlyList<ReaderVisualMaterializedExport> materialized = new[] { export }.StreamVisualExports(
            (bundle, format, payload) => {
                using var reader = new StreamReader(payload, Encoding.UTF8);
                captured = reader.ReadToEnd();
            },
            new ReaderVisualExportMaterializationOptions {
                IncludePayload = false,
                IncludeJson = true
            });

        ReaderVisualMaterializedExport written = Assert.Single(materialized);
        Assert.True(written.Written);
        Assert.Equal(ReaderVisualExportFormat.Json, written.Format);
        Assert.Equal("manual-visual.json", written.FileName);
        Assert.Equal("{\"kind\":\"mermaid\"}", captured);
    }

    [Fact]
    public void ReaderVisualExportMaterializer_WriteVisualExportsToDirectory_DisambiguatesJsonPayloadAndSidecar() {
        var export = new ReaderVisualExportBundle {
            Id = "chart-visual",
            FileNamePrefix = "chart-visual",
            PayloadExtension = ".json",
            Payload = "{\"payload\":true}",
            Json = "{\"kind\":\"chart\"}"
        };
        var directory = Path.Combine(Path.GetTempPath(), "officeimo-reader-json-visuals-" + Guid.NewGuid().ToString("N"));

        try {
            IReadOnlyList<ReaderVisualMaterializedExport> materialized = new[] { export }.WriteVisualExportsToDirectory(directory);

            Assert.Equal(2, materialized.Count);
            Assert.Contains(materialized, item => item.Format == ReaderVisualExportFormat.Payload && item.FileName == "chart-visual.json");
            Assert.Contains(materialized, item => item.Format == ReaderVisualExportFormat.Json && item.FileName == "chart-visual-metadata.json");
            Assert.Equal("{\"payload\":true}", File.ReadAllText(Path.Combine(directory, "chart-visual.json")));
            Assert.Equal("{\"kind\":\"chart\"}", File.ReadAllText(Path.Combine(directory, "chart-visual-metadata.json")));
        } finally {
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void ReaderVisualExport_ToJson_EmitsStableNormalizedVisualShape() {
        var visual = new ReaderVisual {
            Kind = "chart",
            Language = "ix-chart",
            Content = "{\"type\":\"bar\"}",
            PayloadHash = "payload-1",
            Location = new ReaderLocation {
                Path = "dashboard.md",
                BlockIndex = 2,
                SourceBlockKind = "code",
                BlockAnchor = "dashboard--code-2"
            }
        };

        string json = visual.ToJson();

        using JsonDocument document = JsonDocument.Parse(json);
        JsonElement root = document.RootElement;
        Assert.Equal("chart", root.GetProperty("kind").GetString());
        Assert.Equal("ix-chart", root.GetProperty("language").GetString());
        Assert.Equal("payload-1", root.GetProperty("payloadHash").GetString());
        Assert.Equal("{\"type\":\"bar\"}", root.GetProperty("content").GetString());
        Assert.Equal("dashboard.md", root.GetProperty("location").GetProperty("path").GetString());
        Assert.Equal("code", root.GetProperty("location").GetProperty("sourceBlockKind").GetString());
        Assert.Equal(".json", ReaderVisualExport.GetPayloadExtension(visual));
    }
}
