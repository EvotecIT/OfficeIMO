using OfficeIMO.Reader;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderStructuredExtractionTests {
    [Fact]
    public void Extract_ProducesDeterministicTypedSectionsAndRecords() {
        OfficeDocumentReadResult document = CreateRepresentativeDocument();

        OfficeDocumentStructuredExtractionResult result = OfficeDocumentStructuredExtractor.Extract(document);

        Assert.Equal(OfficeDocumentStructuredExtractionSchema.Id, result.SchemaId);
        Assert.Equal(OfficeDocumentStructuredExtractionSchema.Version, result.SchemaVersion);
        Assert.Equal(new[] { null, "Overview", "Details" }, result.Sections.Select(section => section.Heading));
        Assert.Equal("Intro", result.Sections[0].Text);
        Assert.Equal("First paragraph\nSecond paragraph", result.Sections[1].Text);
        Assert.Equal(new[] { "named-values", "shape-data" }, result.Tables.Select(table => table.Title));
        Assert.Equal("customer", Assert.Single(result.Forms).Id);

        Assert.Contains(result.Records, record => record.Category == "metadata" && record.Name == "Author" && record.Value == "Ada");
        Assert.Contains(result.Records, record => record.Category == "form" && record.Name == "Customer" && record.Value == "Contoso");
        Assert.Contains(result.Records, record => record.Category == "key-value" && record.Name == "Region" && record.Value == "EU");
        Assert.Contains(result.Records, record => record.Category == "shape-data" && record.Name == "Owner" && record.Value == "Operations");
        Assert.Contains(result.Records, record =>
            record.Category == "chart-summary" &&
            record.Value == "bar" &&
            record.Attributes["datasetCount"] == "1" &&
            record.Attributes["pointCount"] == "2");
        Assert.Contains(result.Records, record => record.Category == "quality-summary" && record.Name == "named-values");
        Assert.Contains(result.Records, record => record.Category == "visual-summary" && record.Name == "Revenue");
        Assert.Contains(result.Records, record => record.Category == "readiness-summary" && record.Value == "review");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "encrypted-source");

        string first = result.ToJson();
        string second = OfficeDocumentStructuredExtractor.Extract(document).ToJson();
        Assert.Equal(first, second);
        using JsonDocument json = JsonDocument.Parse(first);
        Assert.Equal(OfficeDocumentStructuredExtractionSchema.Id, json.RootElement.GetProperty("schemaId").GetString());
        Assert.Equal("Security", json.RootElement.GetProperty("diagnostics")[0].GetProperty("category").GetString());
    }

    [Fact]
    public void Extract_EnforcesIndependentHardLimitsWithDiagnostics() {
        OfficeDocumentStructuredExtractionResult result = OfficeDocumentStructuredExtractor.Extract(
            CreateRepresentativeDocument(),
            new OfficeDocumentStructuredExtractionOptions {
                MaxRecords = 1,
                MaxSections = 1,
                MaxSectionCharacters = 3,
                MaxTables = 1,
                MaxForms = 1,
                MaxDiagnostics = 1
            });

        Assert.Single(result.Records);
        Assert.Single(result.Sections);
        Assert.Equal("Int", result.Sections[0].Text);
        Assert.True(result.Sections[0].Truncated);
        Assert.Single(result.Tables);
        Assert.Single(result.Forms);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "structured-record-limit");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "structured-section-limit");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "structured-section-character-limit");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "structured-table-limit");
    }

    [Fact]
    public void Extract_RejectsInvalidLimitsAndHonorsCancellation() {
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeDocumentStructuredExtractor.Extract(
            new OfficeDocumentReadResult(),
            new OfficeDocumentStructuredExtractionOptions { MaxRecords = 0 }));

        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => OfficeDocumentStructuredExtractor.Extract(
            CreateRepresentativeDocument(),
            cancellationToken: cancellation.Token));
    }

    [Fact]
    public void Extract_IncludesAndDeduplicatesPageFormsBeforeApplyingLimit() {
        var documentForm = new OfficeDocumentFormField { Id = "customer", Name = "Customer" };
        var pageDuplicate = new OfficeDocumentFormField { Id = "customer", Name = "Customer duplicate" };
        var pageOnly = new OfficeDocumentFormField { Id = "approval", Name = "Approval" };
        var document = new OfficeDocumentReadResult {
            Forms = new[] { documentForm },
            Pages = new[] {
                new OfficeDocumentPage { Forms = new[] { pageDuplicate, pageOnly } }
            }
        };

        OfficeDocumentStructuredExtractionResult result = OfficeDocumentStructuredExtractor.Extract(
            document,
            new OfficeDocumentStructuredExtractionOptions { MaxForms = 2 });

        Assert.Equal(new[] { "customer", "approval" }, result.Forms.Select(form => form.Id));
        Assert.Equal(
            new[] { "customer", "approval" },
            result.Records.Where(record => record.Category == "form").Select(record => record.SourceObjectId));
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "structured-form-limit");
    }

    [Fact]
    public void Extract_ChecksCancellationWhileScanningNonReadinessDiagnostics() {
        var document = new OfficeDocumentReadResult {
            Diagnostics = new[] {
                new OfficeDocumentDiagnostic { Code = "general", Category = OfficeDocumentDiagnosticCategory.General }
            }
        };
        var options = new OfficeDocumentStructuredExtractionOptions {
            IncludeMetadata = false,
            IncludeForms = false,
            IncludeKeyValueRows = false,
            IncludeShapeData = false,
            IncludeChartSummaries = false,
            IncludeQualitySummaries = true,
            IncludeSections = false,
            IncludeNamedTables = false,
            IncludeSourceDiagnostics = false
        };
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() => OfficeDocumentStructuredExtractor.Extract(
            document,
            options,
            cancellation.Token));
    }

    [Fact]
    public void StructuredJson_RejectsUnsupportedSchemaVersions() {
        var result = new OfficeDocumentStructuredExtractionResult { SchemaVersion = 2 };

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => result.ToJson());

        Assert.Contains("version 2", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task StaticStructuredReadSupportsSyncAndAsyncByteSurfaces() {
        byte[] markdown = Encoding.UTF8.GetBytes("# Overview\n\nUseful body");

        OfficeDocumentStructuredExtractionResult sync = DocumentReader.ReadStructured(markdown, "note.md");
        OfficeDocumentStructuredExtractionResult asyncResult = await DocumentReader.ReadStructuredAsync(markdown, "note.md");

        Assert.NotEmpty(sync.Sections);
        Assert.Contains("Overview", sync.Sections[0].Heading ?? sync.Sections[0].Text, StringComparison.Ordinal);
        Assert.Equal(sync.ToJson(), asyncResult.ToJson());
    }

    [Fact]
    public async Task InstanceStructuredReadAppliesConfiguredProcessorsFirst() {
        byte[] markdown = Encoding.UTF8.GetBytes("# Overview\n\nUseful body");
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddProcessor(new DelegateOfficeDocumentProcessor("metadata", (document, _) => {
                document.Metadata = new[] {
                    new OfficeDocumentMetadataEntry { Id = "processed", Name = "Processed", Value = "yes" }
                };
                return document;
            }))
            .Build();

        OfficeDocumentStructuredExtractionResult sync = reader.ReadStructured(markdown, "note.md");
        OfficeDocumentStructuredExtractionResult asyncResult = await reader.ReadStructuredAsync(markdown, "note.md");

        Assert.Contains(sync.Records, record => record.Name == "Processed" && record.Value == "yes");
        Assert.Contains(asyncResult.Records, record => record.Name == "Processed" && record.Value == "yes");
    }

    private static OfficeDocumentReadResult CreateRepresentativeDocument() {
        var namedTable = new ReaderTable {
            Title = "named-values",
            Columns = new[] { "Key", "Value" },
            Rows = new[] { (IReadOnlyList<string>)new[] { "Region", "EU" } },
            TotalRowCount = 1,
            Diagnostics = new ReaderTableDiagnostics {
                Confidence = 0.9,
                SchemaConfidence = 0.8,
                CellCompleteness = 1,
                ColumnGeometryConfidence = 0.7,
                SourceRowCount = 1,
                ExpectedCellCount = 2,
                FilledCellCount = 2,
                HasGeometry = true
            }
        };
        var shapeData = new ReaderTable {
            Title = "shape-data",
            Kind = "visio-shape-data",
            Columns = new[] { "OwnerType", "OwnerId", "Name", "Label", "Value", "Type", "Prompt" },
            Rows = new[] {
                (IReadOnlyList<string>)new[] { "Shape", "7", "Owner", "", "Operations", "string", "Responsible team" }
            },
            TotalRowCount = 1
        };
        var chart = new ReaderVisual {
            Kind = "chart",
            Language = "ix-chart",
            SourceName = "Revenue",
            PayloadHash = "chart-hash",
            Content = "{\"type\":\"bar\",\"data\":{\"labels\":[\"Q1\",\"Q2\"],\"datasets\":[{\"data\":[1,2]}]}}",
            Width = 640,
            Height = 480,
            PlacementCount = 1,
            HasGeometry = true,
            IsAxisAligned = true
        };
        return new OfficeDocumentReadResult {
            Source = new OfficeDocumentSource { Path = "representative.docx", Title = "Representative" },
            Blocks = new[] {
                new OfficeDocumentBlock { Id = "preamble", Kind = "paragraph", Text = "Intro" },
                new OfficeDocumentBlock { Id = "overview", Kind = "heading", Text = "Overview", Level = 1 },
                new OfficeDocumentBlock { Id = "p1", Kind = "paragraph", Text = "First paragraph" },
                new OfficeDocumentBlock { Id = "p2", Kind = "paragraph", Text = "Second paragraph" },
                new OfficeDocumentBlock { Id = "details", Kind = "heading", Text = "Details", Level = 2 },
                new OfficeDocumentBlock { Id = "p3", Kind = "paragraph", Text = "Final paragraph" }
            },
            Metadata = new[] {
                new OfficeDocumentMetadataEntry { Id = "author", Category = "core", Name = "Author", Value = "Ada" }
            },
            Forms = new[] {
                new OfficeDocumentFormField { Id = "customer", Name = "Customer", Kind = "text", Value = "Contoso", IsRequired = true }
            },
            Tables = new[] { namedTable, shapeData },
            Visuals = new[] { chart },
            Chunks = new[] {
                new ReaderChunk {
                    Id = "chunk-1",
                    Text = "Useful body",
                    Diagnostics = new ReaderChunkDiagnostics {
                        SourceKind = "word",
                        TableCount = 2,
                        ImageCount = 1,
                        FormFieldCount = 1,
                        HasSecurityState = true,
                        HasEncryption = true
                    }
                }
            },
            Diagnostics = new[] {
                new OfficeDocumentDiagnostic {
                    Category = OfficeDocumentDiagnosticCategory.Security,
                    Code = "encrypted-source",
                    Message = "Source reports encryption.",
                    IsRecoverable = true
                }
            }
        };
    }
}
