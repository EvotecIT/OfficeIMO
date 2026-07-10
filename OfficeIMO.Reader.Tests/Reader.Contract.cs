using OfficeIMO.Reader;
using System.Text.Json;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderContractTests {
    [Fact]
    public void OfficeDocumentReadResultSchema_ExposesStableCurrentContract() {
        Assert.Equal(5, OfficeDocumentReadResultSchema.MinimumSupportedVersion);
        Assert.Equal(5, OfficeDocumentReadResultSchema.CurrentVersion);
        Assert.Equal(OfficeDocumentReadResultSchema.CurrentVersion, OfficeDocumentReadResultSchema.Version);
        Assert.True(OfficeDocumentReadResultSchema.IsSupported(
            OfficeDocumentReadResultSchema.Id,
            OfficeDocumentReadResultSchema.CurrentVersion));
        Assert.False(OfficeDocumentReadResultSchema.IsSupported(OfficeDocumentReadResultSchema.Id, 4));
        Assert.False(OfficeDocumentReadResultSchema.IsSupported("other.schema", 5));
    }

    [Fact]
    public void OfficeDocumentReadResultSchema_EmbedsVersionedJsonSchema() {
        string schemaJson = OfficeDocumentReadResultSchema.GetJsonSchema();

        using JsonDocument schema = JsonDocument.Parse(schemaJson);
        JsonElement root = schema.RootElement;
        Assert.Equal(OfficeDocumentReadResultSchema.JsonSchemaId, root.GetProperty("$id").GetString());
        Assert.Equal(OfficeDocumentReadResultSchema.Id,
            root.GetProperty("properties").GetProperty("schemaId").GetProperty("const").GetString());
        Assert.Equal(OfficeDocumentReadResultSchema.CurrentVersion,
            root.GetProperty("properties").GetProperty("schemaVersion").GetProperty("const").GetInt32());

        string[] properties = root.GetProperty("properties")
            .EnumerateObject()
            .Select(property => property.Name)
            .ToArray();
        Assert.Equal(new[] {
            "schemaId",
            "schemaVersion",
            "kind",
            "source",
            "capabilitiesUsed",
            "markdown",
            "html",
            "json",
            "chunks",
            "metadata",
            "pages",
            "blocks",
            "tables",
            "assets",
            "links",
            "forms",
            "ocrCandidates",
            "visuals",
            "diagnostics"
        }, properties);
    }

    [Fact]
    public void OfficeDocumentReadResultJson_RoundTripsCurrentTransportShape() {
        var original = new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Pdf,
            Source = new OfficeDocumentSource {
                Path = "report.pdf",
                SourceId = "source-1",
                LengthBytes = 42
            },
            CapabilitiesUsed = new[] { "officeimo.reader.pdf" },
            Markdown = "# Report",
            Chunks = new[] {
                new ReaderChunk {
                    Id = "chunk-1",
                    Kind = ReaderInputKind.Pdf,
                    Text = "Report body",
                    Location = new ReaderLocation { Path = "report.pdf", Page = 1 }
                }
            },
            Tables = new[] {
                new ReaderTable {
                    Title = "Totals",
                    Columns = new[] { "Name", "Value" },
                    Rows = new[] { (IReadOnlyList<string>)new[] { "Items", "2" } },
                    TotalRowCount = 1
                }
            },
            Diagnostics = new[] {
                new OfficeDocumentDiagnostic {
                    Severity = OfficeDocumentDiagnosticSeverity.Warning,
                    Category = OfficeDocumentDiagnosticCategory.Content,
                    Code = "content-truncated",
                    Message = "Output was bounded.",
                    Source = "officeimo.reader.pdf",
                    IsRecoverable = true,
                    Attributes = new Dictionary<string, string> { ["limit"] = "42" }
                }
            }
        };

        string json = OfficeDocumentReadResultJson.Serialize(original);
        OfficeDocumentReadResult restored = OfficeDocumentReadResultJson.Deserialize(json);

        Assert.Equal(OfficeDocumentReadResultSchema.Id, restored.SchemaId);
        Assert.Equal(OfficeDocumentReadResultSchema.CurrentVersion, restored.SchemaVersion);
        Assert.Equal(ReaderInputKind.Pdf, restored.Kind);
        Assert.Equal("report.pdf", restored.Source.Path);
        Assert.Equal("officeimo.reader.pdf", Assert.Single(restored.CapabilitiesUsed));
        Assert.Equal("Report body", Assert.Single(restored.Chunks).Text);
        Assert.Equal("2", Assert.Single(Assert.Single(restored.Tables).Rows)[1]);
        OfficeDocumentDiagnostic diagnostic = Assert.Single(restored.Diagnostics);
        Assert.Equal(OfficeDocumentDiagnosticCategory.Content, diagnostic.Category);
        Assert.Equal("42", diagnostic.Attributes["limit"]);
    }

    [Theory]
    [InlineData("other.schema", 5)]
    [InlineData("officeimo.document.read-result", 4)]
    [InlineData("officeimo.document.read-result", 6)]
    public void OfficeDocumentReadResultJson_RejectsUnsupportedSchemaHeaders(string schemaId, int schemaVersion) {
        string json = $"{{\"schemaId\":\"{schemaId}\",\"schemaVersion\":{schemaVersion}}}";

        OfficeDocumentReadResultSchemaException exception = Assert.Throws<OfficeDocumentReadResultSchemaException>(
            () => OfficeDocumentReadResultJson.Deserialize(json));

        Assert.Equal(schemaId, exception.SchemaId);
        Assert.Equal(schemaVersion, exception.SchemaVersion);
    }

    [Fact]
    public void OfficeDocumentReadResultJson_RejectsUnsupportedSchemaDuringSerialization() {
        var result = new OfficeDocumentReadResult { SchemaVersion = 4 };

        OfficeDocumentReadResultSchemaException exception = Assert.Throws<OfficeDocumentReadResultSchemaException>(
            () => OfficeDocumentReadResultJson.Serialize(result));

        Assert.Equal(4, exception.SchemaVersion);
    }

    [Fact]
    public void OfficeDocumentReadResultJson_RejectsIncompleteCurrentEnvelope() {
        const string json = "{\"schemaId\":\"officeimo.document.read-result\",\"schemaVersion\":5}";

        JsonException exception = Assert.Throws<JsonException>(() => OfficeDocumentReadResultJson.Deserialize(json));

        Assert.Contains("kind", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeDocumentReadResultJson_RejectsUnknownTransportMembers() {
        string json = OfficeDocumentReadResultJson.Serialize(new OfficeDocumentReadResult());
        string withUnknownMember = json.Insert(json.Length - 1, ",\"futureField\":true");

        JsonException exception = Assert.Throws<JsonException>(
            () => OfficeDocumentReadResultJson.Deserialize(withUnknownMember));

        Assert.Contains("futureField", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeDocumentReadResultJson_RejectsNumericEnums() {
        string json = OfficeDocumentReadResultJson.Serialize(new OfficeDocumentReadResult());
        string withNumericKind = json.Replace("\"kind\":\"Unknown\"", "\"kind\":0", StringComparison.Ordinal);

        Assert.Throws<JsonException>(() => OfficeDocumentReadResultJson.Deserialize(withNumericKind));
    }
}
