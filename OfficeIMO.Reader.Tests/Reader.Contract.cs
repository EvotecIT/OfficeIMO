using OfficeIMO.Reader;
using System.Text.Json;
using System.Text.Json.Nodes;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderContractTests {
    [Fact]
    public void ReaderPackages_ExposeFacadeExtensions_NotFormatAdapterBrains() {
        Type[] publicPackageAnchors = {
            typeof(OfficeIMO.Reader.AsciiDoc.OfficeDocumentReaderBuilderAsciiDocExtensions),
            typeof(OfficeIMO.Reader.Csv.OfficeDocumentReaderBuilderCsvExtensions),
            typeof(OfficeIMO.Reader.Epub.OfficeDocumentReaderBuilderEpubExtensions),
            typeof(OfficeIMO.Reader.Html.OfficeDocumentReaderBuilderHtmlExtensions),
            typeof(OfficeIMO.Reader.Json.OfficeDocumentReaderBuilderJsonExtensions),
            typeof(OfficeIMO.Reader.Latex.OfficeDocumentReaderBuilderLatexExtensions),
            typeof(OfficeIMO.Reader.OpenDocument.OfficeDocumentReaderBuilderOpenDocumentExtensions),
            typeof(OfficeIMO.Reader.OneNote.OfficeDocumentReaderBuilderOneNoteExtensions),
            typeof(OfficeIMO.Reader.Pdf.OfficeDocumentReaderBuilderPdfExtensions),
            typeof(OfficeIMO.Reader.Rtf.OfficeDocumentReaderBuilderRtfExtensions),
            typeof(OfficeIMO.Reader.Visio.OfficeDocumentReaderBuilderVisioExtensions),
            typeof(OfficeIMO.Reader.Xml.OfficeDocumentReaderBuilderXmlExtensions),
            typeof(OfficeIMO.Reader.Yaml.OfficeDocumentReaderBuilderYamlExtensions),
            typeof(OfficeIMO.Reader.Zip.OfficeDocumentReaderBuilderZipExtensions)
        };

        foreach (var assembly in publicPackageAnchors.Select(type => type.Assembly).Distinct()) {
            Assert.DoesNotContain(
                assembly.GetExportedTypes(),
                type => type.Name.EndsWith("ReaderAdapter", StringComparison.Ordinal));
        }
    }

    [Fact]
    public void OfficeDocumentReadResultSchema_ExposesStableCurrentContract() {
        Assert.Equal(5, OfficeDocumentReadResultSchema.MinimumSupportedVersion);
        Assert.Equal(6, OfficeDocumentReadResultSchema.CurrentVersion);
        Assert.True(OfficeDocumentReadResultSchema.IsSupported(
            OfficeDocumentReadResultSchema.Id, 5));
        Assert.True(OfficeDocumentReadResultSchema.IsSupported(
            OfficeDocumentReadResultSchema.Id,
            OfficeDocumentReadResultSchema.CurrentVersion));
        Assert.False(OfficeDocumentReadResultSchema.IsSupported(OfficeDocumentReadResultSchema.Id, 4));
        Assert.False(OfficeDocumentReadResultSchema.IsSupported("other.schema", 5));
    }

    [Fact]
    public void OfficeDocumentReadResultSchema_PreservesTheClosedVersion5KindContract() {
        using JsonDocument schema = JsonDocument.Parse(OfficeDocumentReadResultSchema.GetJsonSchema(5));
        JsonElement root = schema.RootElement;
        string[] kinds = root.GetProperty("properties").GetProperty("kind").GetProperty("enum")
            .EnumerateArray().Select(value => value.GetString()!).ToArray();

        Assert.Equal("urn:officeimo:schema:document-read-result:5", root.GetProperty("$id").GetString());
        Assert.Equal(5, root.GetProperty("properties").GetProperty("schemaVersion")
            .GetProperty("const").GetInt32());
        Assert.DoesNotContain(nameof(ReaderInputKind.Calendar), kinds);
        Assert.DoesNotContain(nameof(ReaderInputKind.VCard), kinds);
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
        Assert.Equal(
            Enum.GetNames(typeof(ReaderInputKind)),
            root.GetProperty("properties").GetProperty("kind").GetProperty("enum")
                .EnumerateArray()
                .Select(value => value.GetString())
                .ToArray());

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
        var chunkLocation = new ReaderLocation {
            Path = "report.pdf",
            Page = 1,
            HeadingPath = "Q1 > Q2"
        };
        ReaderHeadingPath.SetHierarchyPath(chunkLocation, ReaderHeadingPath.Combine(new[] { "Q1 > Q2" }));
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
                    Location = chunkLocation
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
        ReaderChunk restoredChunk = Assert.Single(restored.Chunks);
        Assert.Equal("Report body", restoredChunk.Text);
        Assert.Equal("Q1 > Q2", restoredChunk.Location.HeadingPath);
        Assert.Equal(@"Q1 \> Q2", restoredChunk.Location.HierarchyHeadingPath);
        ReaderChunkHierarchyNode restoredHeading = Assert.Single(
            ReaderHierarchicalChunker.Chunk(restored).Nodes,
            node => node.Kind == ReaderChunkHierarchyNodeKind.Heading);
        Assert.Equal("Q1 > Q2", restoredHeading.Title);
        Assert.Equal("2", Assert.Single(Assert.Single(restored.Tables).Rows)[1]);
        OfficeDocumentDiagnostic diagnostic = Assert.Single(restored.Diagnostics);
        Assert.Equal(OfficeDocumentDiagnosticCategory.Content, diagnostic.Category);
        Assert.Equal("42", diagnostic.Attributes["limit"]);
    }

    [Fact]
    public void OfficeDocumentReadResultJson_SerializesNormalizedSchemaWithoutMutatingInput() {
        var result = new OfficeDocumentReadResult {
            SchemaId = string.Empty,
            SchemaVersion = 0
        };

        string json = OfficeDocumentReadResultJson.Serialize(result);
        OfficeDocumentReadResult restored = OfficeDocumentReadResultJson.Deserialize(json);

        Assert.Equal(OfficeDocumentReadResultSchema.Id, restored.SchemaId);
        Assert.Equal(OfficeDocumentReadResultSchema.CurrentVersion, restored.SchemaVersion);
        Assert.Equal(string.Empty, result.SchemaId);
        Assert.Equal(0, result.SchemaVersion);
    }

    [Fact]
    public void OfficeDocumentReadResultJson_NormalizesRequiredNullMembersWithoutMutatingInput() {
        var result = new OfficeDocumentReadResult {
            Source = null!,
            CapabilitiesUsed = null!,
            Chunks = null!,
            Metadata = null!,
            Pages = null!,
            Blocks = null!,
            Tables = null!,
            Assets = null!,
            Links = null!,
            Forms = null!,
            OcrCandidates = null!,
            Visuals = null!,
            Diagnostics = null!
        };

        string json = OfficeDocumentReadResultJson.Serialize(result);
        OfficeDocumentReadResult restored = OfficeDocumentReadResultJson.Deserialize(json);

        Assert.NotNull(restored.Source);
        Assert.Empty(restored.CapabilitiesUsed);
        Assert.Empty(restored.Chunks);
        Assert.Empty(restored.Metadata);
        Assert.Empty(restored.Pages);
        Assert.Empty(restored.Blocks);
        Assert.Empty(restored.Tables);
        Assert.Empty(restored.Assets);
        Assert.Empty(restored.Links);
        Assert.Empty(restored.Forms);
        Assert.Empty(restored.OcrCandidates);
        Assert.Empty(restored.Visuals);
        Assert.Empty(restored.Diagnostics);
        Assert.Null(result.Source);
        Assert.Null(result.Chunks);
    }

    [Fact]
    public void OfficeDocumentReadResultJson_SortsAttributeDictionariesWithoutMutatingInput() {
        var metadataAttributes = new Dictionary<string, string> {
            ["zeta"] = "last",
            ["alpha"] = "first"
        };
        var diagnosticAttributes = new Dictionary<string, string> {
            ["zeta"] = "last",
            ["alpha"] = "first"
        };
        var result = new OfficeDocumentReadResult {
            Metadata = new[] {
                new OfficeDocumentMetadataEntry {
                    Id = "metadata-1",
                    Category = "core",
                    Name = "fixture",
                    Attributes = metadataAttributes
                }
            },
            Diagnostics = new[] {
                new OfficeDocumentDiagnostic {
                    Code = "fixture",
                    Message = "Fixture",
                    Attributes = diagnosticAttributes
                }
            }
        };

        string json = OfficeDocumentReadResultJson.Serialize(result);
        using JsonDocument parsed = JsonDocument.Parse(json);

        Assert.Equal(new[] { "alpha", "zeta" }, parsed.RootElement.GetProperty("metadata")[0]
            .GetProperty("attributes").EnumerateObject().Select(property => property.Name));
        Assert.Equal(new[] { "alpha", "zeta" }, parsed.RootElement.GetProperty("diagnostics")[0]
            .GetProperty("attributes").EnumerateObject().Select(property => property.Name));
        Assert.Equal(new[] { "zeta", "alpha" }, metadataAttributes.Keys);
        Assert.Equal(new[] { "zeta", "alpha" }, diagnosticAttributes.Keys);
    }

    [Theory]
    [InlineData("other.schema", 5)]
    [InlineData("officeimo.document.read-result", 4)]
    [InlineData("officeimo.document.read-result", 7)]
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
        const string json = "{\"schemaId\":\"officeimo.document.read-result\",\"schemaVersion\":6}";

        JsonException exception = Assert.Throws<JsonException>(() => OfficeDocumentReadResultJson.Deserialize(json));

        Assert.Contains("kind", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeDocumentReadResultJson_RejectsNewKindsWhenExplicitlyWritingVersion5() {
        var result = new OfficeDocumentReadResult {
            SchemaVersion = 5,
            Kind = ReaderInputKind.Calendar
        };

        JsonException exception = Assert.Throws<JsonException>(() =>
            OfficeDocumentReadResultJson.Serialize(result));

        Assert.Contains("schema version 5", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void OfficeDocumentReadResultJson_ReadsVersion5ButRejectsVersion5NewKinds() {
        string legacyJson = OfficeDocumentReadResultJson.Serialize(new OfficeDocumentReadResult {
            SchemaVersion = 5,
            Kind = ReaderInputKind.Pdf
        });
        OfficeDocumentReadResult legacy = OfficeDocumentReadResultJson.Deserialize(legacyJson);
        Assert.Equal(OfficeDocumentReadResultSchema.CurrentVersion, legacy.SchemaVersion);
        Assert.Equal(ReaderInputKind.Pdf, legacy.Kind);

        JsonObject invalid = JsonNode.Parse(OfficeDocumentReadResultJson.Serialize(
            new OfficeDocumentReadResult { Kind = ReaderInputKind.VCard }))!.AsObject();
        invalid["schemaVersion"] = 5;
        Assert.Throws<JsonException>(() => OfficeDocumentReadResultJson.Deserialize(invalid.ToJsonString()));
    }

    [Fact]
    public void OfficeDocumentReadResultJson_RejectsVersion6KindsInVersion5Chunks() {
        var result = new OfficeDocumentReadResult {
            SchemaVersion = 5,
            Kind = ReaderInputKind.Email,
            Chunks = new[] {
                new ReaderChunk { Id = "calendar", Kind = ReaderInputKind.Calendar }
            }
        };

        JsonException writeException = Assert.Throws<JsonException>(() =>
            OfficeDocumentReadResultJson.Serialize(result));
        Assert.Contains("schema version 5", writeException.Message,
            StringComparison.OrdinalIgnoreCase);

        JsonObject invalid = JsonNode.Parse(OfficeDocumentReadResultJson.Serialize(
            new OfficeDocumentReadResult {
                Kind = ReaderInputKind.Email,
                Chunks = new[] {
                    new ReaderChunk { Id = "contact", Kind = ReaderInputKind.VCard }
                }
            }))!.AsObject();
        invalid["schemaVersion"] = 5;

        JsonException readException = Assert.Throws<JsonException>(() =>
            OfficeDocumentReadResultJson.Deserialize(invalid.ToJsonString()));
        Assert.Contains("schema version 5", readException.Message,
            StringComparison.OrdinalIgnoreCase);
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
    public void OfficeDocumentReadResultJson_RejectsUnknownSourceMembers() {
        JsonObject envelope = JsonNode.Parse(
            OfficeDocumentReadResultJson.Serialize(new OfficeDocumentReadResult()))!.AsObject();
        envelope["source"]!["futureField"] = true;

        JsonException exception = Assert.Throws<JsonException>(
            () => OfficeDocumentReadResultJson.Deserialize(envelope.ToJsonString()));

        Assert.Contains("source.futureField", exception.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("source")]
    [InlineData("capabilitiesUsed")]
    [InlineData("chunks")]
    [InlineData("diagnostics")]
    public void OfficeDocumentReadResultJson_RejectsNullRequiredMembers(string propertyName) {
        JsonObject envelope = JsonNode.Parse(
            OfficeDocumentReadResultJson.Serialize(new OfficeDocumentReadResult()))!.AsObject();
        envelope[propertyName] = null;

        JsonException exception = Assert.Throws<JsonException>(
            () => OfficeDocumentReadResultJson.Deserialize(envelope.ToJsonString()));

        Assert.Contains(propertyName, exception.Message, StringComparison.Ordinal);
        Assert.Contains("cannot be null", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeDocumentReadResultJson_AllowsNestedExtensionMembers() {
        var result = new OfficeDocumentReadResult {
            Chunks = new[] {
                new ReaderChunk { Id = "chunk-1", Kind = ReaderInputKind.Text, Text = "Body" }
            }
        };
        JsonObject envelope = JsonNode.Parse(OfficeDocumentReadResultJson.Serialize(result))!.AsObject();
        JsonObject chunk = envelope["chunks"]![0]!.AsObject();
        chunk["futureField"] = true;

        OfficeDocumentReadResult restored = OfficeDocumentReadResultJson.Deserialize(envelope.ToJsonString());

        Assert.Equal("chunk-1", Assert.Single(restored.Chunks).Id);
    }

    [Theory]
    [InlineData("chunks")]
    [InlineData("metadata")]
    [InlineData("pages")]
    [InlineData("blocks")]
    [InlineData("tables")]
    [InlineData("assets")]
    [InlineData("links")]
    [InlineData("forms")]
    [InlineData("ocrCandidates")]
    [InlineData("visuals")]
    [InlineData("diagnostics")]
    public void OfficeDocumentReadResultJson_RejectsNullItemsInRequiredObjectArrays(string propertyName) {
        JsonObject envelope = JsonNode.Parse(
            OfficeDocumentReadResultJson.Serialize(new OfficeDocumentReadResult()))!.AsObject();
        envelope[propertyName] = new JsonArray((JsonNode?)null);

        JsonException exception = Assert.Throws<JsonException>(
            () => OfficeDocumentReadResultJson.Deserialize(envelope.ToJsonString()));

        Assert.Contains(propertyName, exception.Message, StringComparison.Ordinal);
        Assert.Contains("item 0", exception.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("message")]
    [InlineData("attributes")]
    public void OfficeDocumentReadResultJson_RejectsNullRequiredDiagnosticMembers(string propertyName) {
        var result = new OfficeDocumentReadResult {
            Diagnostics = new[] { new OfficeDocumentDiagnostic { Code = "fixture", Message = "Fixture" } }
        };
        JsonObject envelope = JsonNode.Parse(OfficeDocumentReadResultJson.Serialize(result))!.AsObject();
        envelope["diagnostics"]![0]![propertyName] = null;

        JsonException exception = Assert.Throws<JsonException>(
            () => OfficeDocumentReadResultJson.Deserialize(envelope.ToJsonString()));

        Assert.Contains(propertyName, exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void OfficeDocumentReadResultJson_RejectsDiagnosticsWithoutStableCodes() {
        var result = new OfficeDocumentReadResult {
            Diagnostics = new[] { new OfficeDocumentDiagnostic { Message = "Missing code" } }
        };

        JsonException exception = Assert.Throws<JsonException>(() => OfficeDocumentReadResultJson.Serialize(result));

        Assert.Contains("non-empty code", exception.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("message")]
    [InlineData("attributes")]
    public void OfficeDocumentReadResultJson_RejectsNullDiagnosticMembersDuringSerialization(string propertyName) {
        var diagnostic = new OfficeDocumentDiagnostic { Code = "fixture", Message = "Fixture" };
        if (propertyName == "message") diagnostic.Message = null!;
        if (propertyName == "attributes") diagnostic.Attributes = null!;
        var result = new OfficeDocumentReadResult { Diagnostics = new[] { diagnostic } };

        JsonException exception = Assert.Throws<JsonException>(() => OfficeDocumentReadResultJson.Serialize(result));

        Assert.Contains(propertyName, exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void OfficeDocumentReadResultJson_RejectsNumericEnums() {
        string json = OfficeDocumentReadResultJson.Serialize(new OfficeDocumentReadResult());
        string withNumericKind = json.Replace("\"kind\":\"Unknown\"", "\"kind\":0");

        Assert.Throws<JsonException>(() => OfficeDocumentReadResultJson.Deserialize(withNumericKind));
    }

    [Theory]
    [InlineData("markdown")]
    [InlineData("html")]
    [InlineData("json")]
    public void OfficeDocumentReadResultJson_RejectsNullOptionalTextMembers(string propertyName) {
        JsonObject envelope = JsonNode.Parse(
            OfficeDocumentReadResultJson.Serialize(new OfficeDocumentReadResult()))!.AsObject();
        envelope[propertyName] = null;

        JsonException exception = Assert.Throws<JsonException>(
            () => OfficeDocumentReadResultJson.Deserialize(envelope.ToJsonString()));

        Assert.Contains(propertyName, exception.Message, StringComparison.Ordinal);
        Assert.Contains("string", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("kind", "pdf")]
    [InlineData("severity", "warning")]
    [InlineData("category", "ocr")]
    public void OfficeDocumentReadResultJson_RejectsEnumsWithSchemaInvalidCasing(string propertyName, string value) {
        var result = new OfficeDocumentReadResult {
            Diagnostics = new[] { new OfficeDocumentDiagnostic { Code = "fixture", Message = "Fixture" } }
        };
        JsonObject envelope = JsonNode.Parse(OfficeDocumentReadResultJson.Serialize(result))!.AsObject();
        if (propertyName == "kind") {
            envelope[propertyName] = value;
        } else {
            envelope["diagnostics"]![0]![propertyName] = value;
        }

        JsonException exception = Assert.Throws<JsonException>(
            () => OfficeDocumentReadResultJson.Deserialize(envelope.ToJsonString()));

        Assert.Contains(propertyName, exception.Message, StringComparison.Ordinal);
        Assert.Contains("enum", exception.Message, StringComparison.OrdinalIgnoreCase);
    }
}
