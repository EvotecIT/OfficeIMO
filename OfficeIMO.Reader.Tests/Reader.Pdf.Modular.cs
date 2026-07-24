using OfficeIMO.Pdf;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Pdf;
using OfficeIMO.Tests.Pdf;
using System.Text;
using System.Text.Json;
using Xunit;

namespace OfficeIMO.Tests;

[Collection("ReaderRegistryNonParallel")]
public sealed class ReaderPdfModularTests {
    [Fact]
    public void DocumentReaderPdf_ReadPdfStream_EmitsPageAwareChunks() {
        byte[] pdf = BuildTwoPagePdf();
        using var stream = new MemoryStream(pdf, writable: false);

        var chunks = PdfReaderAdapter.Read(
            stream,
            sourceName: " sample.pdf ",
            readerOptions: new ReaderOptions { MaxChars = 8_000, ComputeHashes = true }).ToList();

        Assert.Equal(2, chunks.Count);
        Assert.All(chunks, c => {
            Assert.Equal(ReaderInputKind.Pdf, c.Kind);
            Assert.Equal("sample.pdf", c.Location.Path);
            Assert.False(string.IsNullOrWhiteSpace(c.SourceId));
            Assert.False(string.IsNullOrWhiteSpace(c.SourceHash));
            Assert.False(string.IsNullOrWhiteSpace(c.ChunkHash));
            Assert.True(c.TokenEstimate.HasValue && c.TokenEstimate.Value >= 1);
            Assert.Equal(pdf.Length, c.SourceLengthBytes);
            Assert.Null(c.SourceLastWriteUtc);
            Assert.NotNull(c.Diagnostics);
            Assert.Equal("pdf", c.Diagnostics!.SourceKind);
            Assert.Equal(2, c.Diagnostics.PageCount);
            Assert.Equal(2, c.Diagnostics.SelectedPageCount);
            Assert.Equal(c.Location.Page, c.Diagnostics.PageNumber);
            Assert.False(c.Diagnostics.HasSecurityState);
            Assert.False(c.Diagnostics.HasEncryption);
            Assert.False(c.Diagnostics.HasSignatures);
            Assert.False(c.Diagnostics.HasIncrementalUpdates);
            Assert.False(c.Diagnostics.RequiresAppendOnlyMutation);
            Assert.False(c.Diagnostics.HasOpenAction);
            Assert.False(c.Diagnostics.HasCatalogActions);
            Assert.False(c.Diagnostics.HasPageActions);
            Assert.False(c.Diagnostics.HasAnnotationActions);
            Assert.False(c.Diagnostics.HasActiveContent);
            Assert.Equal(0, c.Diagnostics.CatalogActionCount);
            Assert.Equal(0, c.Diagnostics.PageActionCount);
            Assert.Equal(0, c.Diagnostics.SelectedPageActionCount);
            Assert.Equal(0, c.Diagnostics.AnnotationActionCount);
            Assert.Equal(0, c.Diagnostics.SelectedAnnotationActionCount);
            Assert.Equal(0, c.Diagnostics.TableCount);
            Assert.Equal(0, c.Diagnostics.TableGeometryCount);
            Assert.Equal(0D, c.Diagnostics.TableGeometryCoverage);
            Assert.Null(c.Diagnostics.MinTableConfidence);
            Assert.Null(c.Diagnostics.AverageTableConfidence);
            Assert.Equal(0, c.Diagnostics.ImageCount);
            Assert.Equal(0, c.Diagnostics.ImageGeometryCount);
            Assert.Equal(0D, c.Diagnostics.ImageGeometryCoverage);
            Assert.False(c.Diagnostics.HasXmpMetadata);
            Assert.Equal(0, c.Diagnostics.OutputIntentCount);
            Assert.Equal(0, c.Diagnostics.AttachmentCount);
            Assert.False(c.Diagnostics.HasTaggedContent);
            Assert.Equal(0, c.Diagnostics.TaggedStructureElementCount);
            Assert.Equal(0, c.Diagnostics.TaggedMarkedContentReferenceCount);
            Assert.Equal(0, c.Diagnostics.OptionalContentGroupCount);
            Assert.Equal(0, c.Diagnostics.OptionalContentInitiallyHiddenCount);
            Assert.Equal(0, c.Diagnostics.OptionalContentLockedCount);
            Assert.Equal(0, c.Diagnostics.FormFieldCount);
            Assert.Equal(0, c.Diagnostics.FormWidgetCount);
            Assert.Equal(0, c.Diagnostics.SelectedFormWidgetCount);
            Assert.Equal(0, c.Diagnostics.SelectedFormWidgetAppearanceStateCount);
            Assert.Equal(0D, c.Diagnostics.SelectedFormWidgetAppearanceStateCoverage);
            Assert.Equal(0, c.Diagnostics.SelectedFormWidgetNormalAppearanceStateCount);
        });
        Assert.Contains(chunks, c => c.Location.Page == 1 && (c.Markdown ?? c.Text).Contains("Reader PDF page one", StringComparison.Ordinal));
        Assert.Contains(chunks, c => c.Location.Page == 2 && (c.Markdown ?? c.Text).Contains("Reader PDF page two", StringComparison.Ordinal));
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfBytes_ProvidesChunksDocumentAndJsonOverloads() {
        byte[] pdf = BuildTwoPagePdf();

        var chunks = PdfReaderAdapter.Read(
            pdf,
            sourceName: " bytes.pdf ",
            readerOptions: new ReaderOptions { MaxChars = 8_000, ComputeHashes = true }).ToList();

        Assert.Equal(2, chunks.Count);
        Assert.All(chunks, chunk => {
            Assert.Equal(ReaderInputKind.Pdf, chunk.Kind);
            Assert.Equal("bytes.pdf", chunk.Location.Path);
            Assert.Equal(pdf.Length, chunk.SourceLengthBytes);
            Assert.False(string.IsNullOrWhiteSpace(chunk.SourceHash));
        });

        OfficeDocumentReadResult result = PdfReaderAdapter.ReadDocument(
            pdf,
            sourceName: " bytes-result.pdf ",
            readerOptions: new ReaderOptions { MaxChars = 8_000 });

        Assert.Equal("bytes-result.pdf", result.Source.Path);
        Assert.Equal(pdf.Length, result.Source.LengthBytes);
        Assert.Contains("Reader PDF page one", result.Markdown, StringComparison.Ordinal);
        Assert.Equal(2, result.Pages.Count);
        Assert.Equal(2, result.Chunks.Count);

        string json = PdfReaderAdapter.ReadDocumentJson(
            pdf,
            sourceName: " bytes-json.pdf ",
            readerOptions: new ReaderOptions { MaxChars = 8_000 });

        using JsonDocument document = JsonDocument.Parse(json);
        JsonElement root = document.RootElement;
        Assert.Equal("Pdf", root.GetProperty("kind").GetString());
        Assert.Equal("bytes-json.pdf", root.GetProperty("source").GetProperty("path").GetString());
        Assert.Contains("Reader PDF page two", root.GetProperty("markdown").GetString(), StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfStream_CanSelectPageRanges() {
        byte[] pdf = BuildTwoPagePdf();
        using var stream = new MemoryStream(pdf, writable: false);

        var chunks = PdfReaderAdapter.Read(
            stream,
            sourceName: "ranges.pdf",
            pdfOptions: new ReaderPdfOptions {
                PageRanges = new[] { PdfPageRange.From(2, 2) }
            }).ToList();

        var chunk = Assert.Single(chunks);
        Assert.Equal(2, chunk.Location.Page);
        Assert.Contains("Reader PDF page two", chunk.Markdown ?? chunk.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("Reader PDF page one", chunk.Markdown ?? chunk.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfLogicalDocument_CanSelectPageRanges() {
        PdfLogicalDocument logical = PdfLogicalDocument.Load(BuildTwoPagePdf());

        var chunks = PdfReaderAdapter.Read(
            logical,
            sourceName: "logical-ranges.pdf",
            pdfOptions: new ReaderPdfOptions {
                PageRanges = new[] { PdfPageRange.From(2, 2) }
            }).ToList();

        var chunk = Assert.Single(chunks);
        Assert.Equal(2, chunk.Location.Page);
        Assert.Contains("Reader PDF page two", chunk.Markdown ?? chunk.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("Reader PDF page one", chunk.Markdown ?? chunk.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfLogicalDocument_FiltersOpenActionToSelectedPages() {
        byte[] pdf = PdfDocument.Create()
            .OpenAction(pageNumber: 1, destinationMode: PdfOpenActionDestinationMode.Fit)
            .H1("Reader PDF page one")
            .Paragraph(p => p.Text("First page body."))
            .PageBreak()
            .H1("Reader PDF page two")
            .Paragraph(p => p.Text("Second page body."))
            .ToBytes();
        PdfLogicalDocument logical = PdfLogicalDocument.Load(pdf);

        ReaderChunk chunk = Assert.Single(PdfReaderAdapter.Read(
            logical,
            sourceName: "logical-open-action-ranges.pdf",
            pdfOptions: new ReaderPdfOptions {
                PageRanges = new[] { PdfPageRange.From(2, 2) }
            }).ToList());

        Assert.Equal(2, chunk.Location.Page);
        Assert.NotNull(chunk.Diagnostics);
        Assert.False(chunk.Diagnostics!.HasOpenAction);
        Assert.Null(chunk.Actions);
        Assert.Contains("Reader PDF page two", chunk.Markdown ?? chunk.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("Reader PDF page one", chunk.Markdown ?? chunk.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfDocument_MapsLogicalPagesChunksAndBlocks() {
        byte[] pdf = BuildTwoPagePdf();
        using var stream = new MemoryStream(pdf, writable: false);

        OfficeDocumentReadResult result = PdfReaderAdapter.ReadDocument(
            stream,
            sourceName: " readback.pdf ",
            readerOptions: new ReaderOptions { MaxChars = 8_000, ComputeHashes = true });

        Assert.Equal(OfficeDocumentReadResultSchema.Id, result.SchemaId);
        Assert.Equal(OfficeDocumentReadResultSchema.CurrentVersion, result.SchemaVersion);
        Assert.Equal(ReaderInputKind.Pdf, result.Kind);
        Assert.Equal("readback.pdf", result.Source.Path);
        Assert.Equal(pdf.Length, result.Source.LengthBytes);
        Assert.False(string.IsNullOrWhiteSpace(result.Source.SourceId));
        Assert.False(string.IsNullOrWhiteSpace(result.Source.SourceHash));
        Assert.Contains("officeimo.reader.pdf", result.CapabilitiesUsed);
        Assert.Contains("Reader PDF page one", result.Markdown, StringComparison.Ordinal);
        Assert.Contains("Reader PDF page two", result.Markdown, StringComparison.Ordinal);
        Assert.Equal(2, result.Pages.Count);
        Assert.Equal(2, result.Chunks.Count);
        Assert.Contains(result.Blocks, block => block.Location.Page == 1 && block.Text.Contains("Reader PDF page one", StringComparison.Ordinal));
        Assert.Contains(result.Blocks, block => block.Location.Page == 2 && block.Text.Contains("Reader PDF page two", StringComparison.Ordinal));
        Assert.All(result.Pages, page => Assert.NotEmpty(page.Blocks));
    }

    [Fact]
    public void OfficeDocumentReadResult_ToJson_UsesStableTransportShape() {
        var result = new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Pdf,
            Source = new OfficeDocumentSource {
                Path = "sample.pdf",
                SourceId = "source-sample"
            },
            CapabilitiesUsed = new[] { "officeimo.reader.pdf" },
            Metadata = new[] {
                new OfficeDocumentMetadataEntry {
                    Id = "meta-1",
                    Category = "sample.catalog",
                    Name = "PageMode",
                    Value = "UseOutlines",
                    ValueType = "string",
                    Attributes = new Dictionary<string, string>(StringComparer.Ordinal) {
                        ["zeta"] = "last",
                        ["alpha"] = "first"
                    }
                }
            },
            Blocks = new[] {
                new OfficeDocumentBlock {
                    Id = "block-1",
                    Kind = "heading",
                    Text = "Sample",
                    Location = new ReaderLocation {
                        Path = "sample.pdf",
                        Page = 1,
                        BlockAnchor = "page-1-block-1"
                    }
                }
            },
            Diagnostics = new[] {
                new OfficeDocumentDiagnostic {
                    Severity = OfficeDocumentDiagnosticSeverity.Warning,
                    Code = "sample-warning",
                    Message = "Sample warning"
                }
            }
        };

        string json = result.ToJson();

        Assert.StartsWith("{\"schemaId\":\"officeimo.document.read-result\"", json, StringComparison.Ordinal);
        Assert.DoesNotContain("\"html\"", json, StringComparison.Ordinal);
        using JsonDocument document = JsonDocument.Parse(json);
        JsonElement root = document.RootElement;
        Assert.Equal("officeimo.document.read-result", root.GetProperty("schemaId").GetString());
        Assert.Equal(OfficeDocumentReadResultSchema.CurrentVersion, root.GetProperty("schemaVersion").GetInt32());
        Assert.Equal("Pdf", root.GetProperty("kind").GetString());
        Assert.Equal("sample.pdf", root.GetProperty("source").GetProperty("path").GetString());
        Assert.Equal("officeimo.reader.pdf", root.GetProperty("capabilitiesUsed")[0].GetString());
        JsonElement metadata = root.GetProperty("metadata")[0];
        Assert.Equal("sample.catalog", metadata.GetProperty("category").GetString());
        Assert.Equal("PageMode", metadata.GetProperty("name").GetString());
        Assert.Equal("first", metadata.GetProperty("attributes").GetProperty("alpha").GetString());
        Assert.Equal("last", metadata.GetProperty("attributes").GetProperty("zeta").GetString());
        Assert.Equal("heading", root.GetProperty("blocks")[0].GetProperty("kind").GetString());
        Assert.Equal("Warning", root.GetProperty("diagnostics")[0].GetProperty("severity").GetString());
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfDocumentJson_EmitsPagesChunksAndBlocks() {
        byte[] pdf = BuildTwoPagePdf();
        using var stream = new MemoryStream(pdf, writable: false);

        string json = PdfReaderAdapter.ReadDocumentJson(
            stream,
            sourceName: " json.pdf ",
            readerOptions: new ReaderOptions { MaxChars = 8_000 });

        using JsonDocument document = JsonDocument.Parse(json);
        JsonElement root = document.RootElement;
        Assert.Equal("officeimo.document.read-result", root.GetProperty("schemaId").GetString());
        Assert.Equal("Pdf", root.GetProperty("kind").GetString());
        Assert.Equal("json.pdf", root.GetProperty("source").GetProperty("path").GetString());
        Assert.Contains("Reader PDF page one", root.GetProperty("markdown").GetString(), StringComparison.Ordinal);
        Assert.Equal(2, root.GetProperty("pages").GetArrayLength());
        Assert.Equal(2, root.GetProperty("chunks").GetArrayLength());
        Assert.Contains(root.GetProperty("blocks").EnumerateArray(), block =>
            block.GetProperty("location").GetProperty("page").GetInt32() == 1 &&
            block.GetProperty("text").GetString()!.Contains("Reader PDF page one", StringComparison.Ordinal));
        Assert.Contains(root.GetProperty("blocks").EnumerateArray(), block =>
            block.GetProperty("location").GetProperty("page").GetInt32() == 2 &&
            block.GetProperty("text").GetString()!.Contains("Reader PDF page two", StringComparison.Ordinal));
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfDocumentJson_EmitsStructuredChunkMetadata() {
        byte[] pdf = BuildOpenAndCatalogActionsPdf();
        using var stream = new MemoryStream(pdf, writable: false);

        string json = PdfReaderAdapter.ReadDocumentJson(
            stream,
            sourceName: "open-catalog-actions.pdf",
            readerOptions: new ReaderOptions { MaxChars = 8_000 });

        using JsonDocument document = JsonDocument.Parse(json);
        JsonElement chunk = document.RootElement.GetProperty("chunks")[0];
        Assert.True(chunk.GetProperty("diagnostics").GetProperty("hasCatalogActions").GetBoolean());
        Assert.Equal(1, chunk.GetProperty("diagnostics").GetProperty("catalogActionCount").GetInt32());
        Assert.Contains(chunk.GetProperty("actions").EnumerateArray(), action =>
            action.GetProperty("scope").GetString() == "DocumentOpen" &&
            action.GetProperty("destinationPageNumber").GetInt32() == 1);
        Assert.Contains(chunk.GetProperty("actions").EnumerateArray(), action =>
            action.GetProperty("scope").GetString() == "Catalog" &&
            action.GetProperty("name").GetString() == "Startup");
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfDocument_ExposesOpenAndCatalogActionMetadata() {
        OfficeDocumentReadResult result = PdfReaderAdapter.ReadDocument(
            new MemoryStream(BuildOpenAndCatalogActionsPdf(), writable: false),
            sourceName: "open-catalog-actions.pdf");

        Assert.Equal("2", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-action-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-active-action-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-action-document-open-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-action-catalog-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-action-type-destination-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-action-type-javascript-count").Value);
        Assert.DoesNotContain(result.Metadata, metadata => metadata.Id == "pdf-action-chained-count");

        string json = result.ToJson();
        Assert.DoesNotContain("app.alert", json, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfDocument_ExposesAnnotationActionMetadataWithoutPayloads() {
        OfficeDocumentReadResult result = PdfReaderAdapter.ReadDocument(
            new MemoryStream(BuildAnnotationActionsPdf(), writable: false),
            sourceName: "annotation-actions.pdf");

        Assert.Equal("3", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-action-count").Value);
        Assert.Equal("3", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-active-action-count").Value);
        Assert.Equal("3", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-action-annotation-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-action-chained-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-action-type-javascript-count").Value);
        Assert.Equal("2", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-action-type-launch-count").Value);

        string json = result.ToJson();
        Assert.DoesNotContain("app.alert", json, StringComparison.Ordinal);
        Assert.DoesNotContain("tool.exe", json, StringComparison.Ordinal);
        Assert.DoesNotContain("chain.exe", json, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfDocument_ExposesFreeTextAnnotationAppearanceMetadata() {
        OfficeDocumentReadResult result = PdfReaderAdapter.ReadDocument(
            new MemoryStream(BuildFreeTextAppearanceMetadataPdf(), writable: false),
            sourceName: "freetext-appearance.pdf");

        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-annotation-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-annotation-freetext-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-annotation-visual-style-metadata-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-annotation-freetext-appearance-metadata-count").Value);

        OfficeDocumentMetadataEntry annotation = Assert.Single(result.Metadata, metadata =>
            metadata.Category == "pdf.annotation.freeText" &&
            metadata.Name == "FreeText");
        Assert.Equal("FreeText", annotation.Name);
        Assert.Equal("Reader styled note", annotation.Value);
        Assert.Equal("5", annotation.SourceObjectId);
        Assert.Equal(1, annotation.Location!.Page);
        Assert.Equal("annotation", annotation.Location.SourceBlockKind);
        Assert.Equal("5", annotation.Attributes["objectNumber"]);
        Assert.Equal("FreeText", annotation.Attributes["subtype"]);
        Assert.Equal("false", annotation.Attributes["hasNormalAppearance"]);
        Assert.Equal("font-size: 14pt; color: rgb(51, 102, 153); text-align: center", annotation.Attributes["defaultStyle"]);
        Assert.Equal("Rich reader note", annotation.Attributes["richContentsPlainText"]);
        Assert.Equal("14", annotation.Attributes["effectiveFontSize"]);
        Assert.Equal("0.2,0.4,0.6", annotation.Attributes["effectiveTextColor"]);
        Assert.Equal("Center", annotation.Attributes["effectiveTextAlign"]);
        Assert.Equal("0.2,0.4,0.8", annotation.Attributes["color"]);
        Assert.Equal("0.95,0.98,1", annotation.Attributes["interiorColor"]);
        Assert.Equal("0.5", annotation.Attributes["opacity"]);
        Assert.Equal("1", annotation.Attributes["borderWidth"]);

        string json = result.ToJson();
        Assert.Contains("\"category\":\"pdf.annotation.freeText\"", json, StringComparison.Ordinal);
        Assert.Contains("\"richContentsPlainText\":\"Rich reader note\"", json, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfDocument_ExposesAnnotationPathGeometryMetadata() {
        OfficeDocumentReadResult result = PdfReaderAdapter.ReadDocument(
            new MemoryStream(BuildAnnotationPathGeometryMetadataPdf(), writable: false),
            sourceName: "annotation-paths.pdf");

        Assert.Equal("3", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-annotation-count").Value);
        Assert.Equal("3", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-annotation-path-geometry-metadata-count").Value);

        OfficeDocumentMetadataEntry highlight = Assert.Single(result.Metadata, metadata =>
            metadata.Category == "pdf.annotation" &&
            metadata.Name == "Highlight");
        Assert.Equal("30,100,90,100,30,92,90,92", highlight.Attributes["quadPoints"]);

        OfficeDocumentMetadataEntry line = Assert.Single(result.Metadata, metadata =>
            metadata.Category == "pdf.annotation" &&
            metadata.Name == "Line");
        Assert.Equal("40,100,140,100", line.Attributes["lineCoordinates"]);
        Assert.Equal("OpenArrow", line.Attributes["lineStartEnding"]);
        Assert.Equal("ClosedArrow", line.Attributes["lineEndEnding"]);

        OfficeDocumentMetadataEntry ink = Assert.Single(result.Metadata, metadata =>
            metadata.Category == "pdf.annotation" &&
            metadata.Name == "Ink");
        Assert.Equal("30,30,60,45,90,30;100,30,130,45,160,30", ink.Attributes["inkList"]);

        string json = result.ToJson();
        Assert.Contains("\"id\":\"pdf-annotation-path-geometry-metadata-count\"", json, StringComparison.Ordinal);
        Assert.Contains("\"inkList\":\"30,30,60,45,90,30;100,30,130,45,160,30\"", json, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfDocument_PreservesLogicalPageRangeOrder() {
        PdfLogicalDocument logical = PdfLogicalDocument.Load(BuildTwoPagePdf());

        OfficeDocumentReadResult result = PdfReaderAdapter.ReadDocument(
            logical,
            sourceName: "logical-readback.pdf",
            pdfOptions: new ReaderPdfOptions {
                PageRanges = new[] {
                    PdfPageRange.From(2, 2),
                    PdfPageRange.From(1, 1)
                }
            });

        Assert.Equal(new int?[] { 2, 1 }, result.Pages.Select(page => page.Number).ToArray());
        Assert.Equal(new int?[] { 2, 1 }, result.Chunks.Select(chunk => chunk.Location.Page).ToArray());
        Assert.Contains("Reader PDF page two", result.Markdown, StringComparison.Ordinal);
        Assert.Contains("Reader PDF page one", result.Markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfDocument_ExposesCatalogNavigationAndMetadataEntries() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                CreateOutlineFromHeadings = true,
                CatalogPageMode = PdfCatalogPageMode.UseOutlines,
                CatalogPageLayout = PdfCatalogPageLayout.TwoColumnLeft,
                Language = "en-US"
            })
            .ViewerPreferences(preferences => {
                preferences.DisplayDocTitle = true;
                preferences.HideToolbar = true;
            })
            .OpenAction(pageNumber: 1, destinationMode: PdfOpenActionDestinationMode.Fit)
            .H1("Catalog Root")
            .Bookmark("Details")
            .Paragraph(p => p.Text("Detail body"))
            .H2("Catalog Child")
            .ToBytes();
        using var stream = new MemoryStream(pdf, writable: false);

        OfficeDocumentReadResult result = PdfReaderAdapter.ReadDocument(
            stream,
            sourceName: "catalog-navigation.pdf",
            readerOptions: new ReaderOptions { MaxChars = 8_000 });

        Assert.Contains(result.Metadata, entry =>
            entry.Category == "pdf.catalog" &&
            entry.Name == "PageMode" &&
            entry.Value == "UseOutlines");
        Assert.Contains(result.Metadata, entry =>
            entry.Category == "pdf.catalog" &&
            entry.Name == "PageLayout" &&
            entry.Value == "TwoColumnLeft");
        Assert.Contains(result.Metadata, entry =>
            entry.Category == "pdf.catalog" &&
            entry.Name == "Language" &&
            entry.Value == "en-US");
        Assert.Contains(result.Metadata, entry =>
            entry.Category == "pdf.viewerPreference" &&
            entry.Name == "DisplayDocTitle" &&
            entry.Value == "true");
        Assert.Contains(result.Metadata, entry =>
            entry.Category == "pdf.catalog.openAction" &&
            entry.Attributes["destinationMode"] == "Fit" &&
            entry.Attributes["pageNumber"] == "1");
        Assert.Contains(result.Metadata, entry =>
            entry.Category == "pdf.outline" &&
            entry.Name == "Catalog Root" &&
            entry.Attributes["level"] == "1" &&
            entry.Location?.Page == 1);
        Assert.Contains(result.Metadata, entry =>
            entry.Category == "pdf.destination" &&
            entry.Name == "Details" &&
            entry.Attributes["pageNumber"] == "1" &&
            entry.Location?.Page == 1);

        string json = result.ToJson();
        using JsonDocument document = JsonDocument.Parse(json);
        JsonElement metadata = document.RootElement.GetProperty("metadata");
        Assert.Contains(metadata.EnumerateArray(), entry =>
            entry.GetProperty("category").GetString() == "pdf.destination" &&
            entry.GetProperty("name").GetString() == "Details");
        Assert.Contains(metadata.EnumerateArray(), entry =>
            entry.GetProperty("category").GetString() == "pdf.outline" &&
            entry.GetProperty("name").GetString() == "Catalog Root" &&
            entry.GetProperty("attributes").GetProperty("level").GetString() == "1");
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfDocument_ExposesOutputIntentMetadata() {
        byte[] pdf = PdfDocument.Create(new PdfOptions().SetSrgbOutputIntent())
            .Paragraph(p => p.Text("Output intent metadata proof."))
            .ToBytes();

        OfficeDocumentReadResult result = PdfReaderAdapter.ReadDocument(
            new MemoryStream(pdf, writable: false),
            sourceName: "output-intent.pdf");

        ReaderChunkDiagnostics diagnostics = Assert.Single(result.Chunks).Diagnostics!;
        Assert.False(diagnostics.HasXmpMetadata);
        Assert.Equal(1, diagnostics.OutputIntentCount);
        Assert.Equal(0, diagnostics.AttachmentCount);
        Assert.False(diagnostics.HasTaggedContent);
        Assert.Equal(0, diagnostics.OptionalContentGroupCount);

        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-output-intent-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-output-intent-profile-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-output-intent-icc-signature-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-output-intent-subtype-gts-pdfa1-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-output-intent-profile-color-space-rgb-count").Value);

        OfficeDocumentMetadataEntry outputIntent = Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-output-intent-0000");
        Assert.Equal("pdf.outputIntent", outputIntent.Category);
        Assert.Equal("GTS_PDFA1", outputIntent.Name);
        Assert.Equal(PdfIccProfiles.SrgbIec6196621OutputConditionIdentifier, outputIntent.Value);
        Assert.NotNull(outputIntent.SourceObjectId);
        Assert.Equal("GTS_PDFA1", outputIntent.Attributes["subtype"]);
        Assert.Equal(PdfIccProfiles.SrgbIec6196621OutputConditionIdentifier, outputIntent.Attributes["outputConditionIdentifier"]);
        Assert.Equal("3", outputIntent.Attributes["destinationOutputProfileColorComponents"]);
        Assert.Equal("RGB ", outputIntent.Attributes["destinationOutputProfileColorSpace"]);
        Assert.Equal("true", outputIntent.Attributes["destinationOutputProfileHasIccSignature"]);
        Assert.True(int.Parse(outputIntent.Attributes["destinationOutputProfileSizeBytes"], System.Globalization.CultureInfo.InvariantCulture) > 128);
        Assert.Equal(
            outputIntent.Attributes["destinationOutputProfileSizeBytes"],
            outputIntent.Attributes["destinationOutputProfileDeclaredSizeBytes"]);

        using JsonDocument document = JsonDocument.Parse(result.ToJson());
        JsonElement jsonOutputIntent = document.RootElement.GetProperty("metadata").EnumerateArray()
            .Single(entry => entry.GetProperty("id").GetString() == "pdf-output-intent-0000");
        Assert.Equal("pdf.outputIntent", jsonOutputIntent.GetProperty("category").GetString());
        Assert.Equal("GTS_PDFA1", jsonOutputIntent.GetProperty("attributes").GetProperty("subtype").GetString());
        Assert.Equal("true", jsonOutputIntent.GetProperty("attributes").GetProperty("destinationOutputProfileHasIccSignature").GetString());
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfDocument_ExposesTaggedContentMetadata() {
        byte[] pdf = PdfDocument.Create()
            .TaggedPdfCatalogMarkers()
            .Language("en-US")
            .H1("Reader tagged heading")
            .Paragraph(p => p.Text("Reader tagged paragraph."))
            .ToBytes();

        OfficeDocumentReadResult result = PdfReaderAdapter.ReadDocument(
            new MemoryStream(pdf, writable: false),
            sourceName: "tagged-content.pdf");

        ReaderChunkDiagnostics diagnostics = Assert.Single(result.Chunks).Diagnostics!;
        Assert.True(diagnostics.HasTaggedContent);
        Assert.True(diagnostics.TaggedStructureElementCount >= 3);
        Assert.True(diagnostics.TaggedMarkedContentReferenceCount > 0);

        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-tagged-content-count").Value);
        Assert.True(int.Parse(Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-tagged-content-structure-element-count").Value!, System.Globalization.CultureInfo.InvariantCulture) >= 3);
        Assert.True(int.Parse(Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-tagged-content-parent-tree-entry-count").Value!, System.Globalization.CultureInfo.InvariantCulture) > 0);
        Assert.True(int.Parse(Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-tagged-content-marked-content-reference-count").Value!, System.Globalization.CultureInfo.InvariantCulture) > 0);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-tagged-content-type-document-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-tagged-content-type-h1-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-tagged-content-type-p-count").Value);

        OfficeDocumentMetadataEntry tagged = Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-tagged-content");
        Assert.Equal("pdf.taggedContent", tagged.Category);
        Assert.Equal("TaggedContent", tagged.Name);
        Assert.Equal("true", tagged.Value);
        Assert.Equal("true", tagged.Attributes["marked"]);
        Assert.Contains("Document", tagged.Attributes["structureTypes"], StringComparison.Ordinal);
        Assert.Contains("H1", tagged.Attributes["structureTypes"], StringComparison.Ordinal);
        Assert.Contains("P", tagged.Attributes["structureTypes"], StringComparison.Ordinal);
        Assert.Equal("true", tagged.Attributes["hasDocumentStructureElement"]);
        Assert.Equal("true", tagged.Attributes["hasMarkedContentReferences"]);
        Assert.NotNull(tagged.SourceObjectId);

        Assert.Contains(result.Metadata, metadata =>
            metadata.Category == "pdf.taggedContent.element" &&
            metadata.Name == "Document" &&
            metadata.Attributes["language"] == "en-US");

        Assert.Contains(result.Metadata, metadata =>
            metadata.Category == "pdf.taggedContent.element" &&
            metadata.Name == "P" &&
            int.Parse(metadata.Attributes["markedContentReferenceCount"], System.Globalization.CultureInfo.InvariantCulture) > 0);

        using JsonDocument document = JsonDocument.Parse(result.ToJson());
        JsonElement jsonTagged = document.RootElement.GetProperty("metadata").EnumerateArray()
            .Single(entry => entry.GetProperty("id").GetString() == "pdf-tagged-content");
        Assert.Equal("pdf.taggedContent", jsonTagged.GetProperty("category").GetString());
        Assert.Equal("true", jsonTagged.GetProperty("attributes").GetProperty("marked").GetString());
        Assert.Contains("Document", jsonTagged.GetProperty("attributes").GetProperty("structureTypes").GetString(), StringComparison.Ordinal);

        JsonElement jsonDiagnostics = document.RootElement.GetProperty("chunks")[0].GetProperty("diagnostics");
        Assert.True(jsonDiagnostics.GetProperty("hasTaggedContent").GetBoolean());
        Assert.True(jsonDiagnostics.GetProperty("taggedStructureElementCount").GetInt32() >= 3);
        Assert.True(jsonDiagnostics.GetProperty("taggedMarkedContentReferenceCount").GetInt32() > 0);
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfDocument_ExposesOptionalContentMetadata() {
        OfficeDocumentReadResult result = PdfReaderAdapter.ReadDocument(
            new MemoryStream(BuildOptionalContentMetadataPdf(), writable: false),
            sourceName: "optional-content.pdf");

        ReaderChunkDiagnostics diagnostics = Assert.Single(result.Chunks).Diagnostics!;
        Assert.Equal(2, diagnostics.OptionalContentGroupCount);
        Assert.Equal(1, diagnostics.OptionalContentInitiallyHiddenCount);
        Assert.Equal(1, diagnostics.OptionalContentLockedCount);

        Assert.Equal("2", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-optional-content-group-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-optional-content-initially-visible-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-optional-content-initially-hidden-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-optional-content-locked-count").Value);
        Assert.Equal("2", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-optional-content-default-order-count").Value);

        OfficeDocumentMetadataEntry configuration = Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-optional-content-configuration");
        Assert.Equal("pdf.optionalContent", configuration.Category);
        Assert.Equal("DefaultConfiguration", configuration.Name);
        Assert.Equal("Default layers", configuration.Value);
        Assert.Equal("Default layers", configuration.Attributes["name"]);
        Assert.Equal("OfficeIMO fixture", configuration.Attributes["creator"]);
        Assert.Equal("ON", configuration.Attributes["baseState"]);
        Assert.Equal("5", configuration.Attributes["onGroupObjectNumbers"]);
        Assert.Equal("6", configuration.Attributes["offGroupObjectNumbers"]);
        Assert.Equal("6", configuration.Attributes["lockedGroupObjectNumbers"]);
        Assert.Equal("5,6", configuration.Attributes["orderGroupObjectNumbers"]);

        OfficeDocumentMetadataEntry printLayer = Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-optional-content-group-0000");
        Assert.Equal("pdf.optionalContent.group", printLayer.Category);
        Assert.Equal("Print layer", printLayer.Name);
        Assert.Equal("true", printLayer.Value);
        Assert.Equal("5", printLayer.SourceObjectId);
        Assert.Equal("5", printLayer.Attributes["objectNumber"]);
        Assert.Equal("View,Design", printLayer.Attributes["intents"]);
        Assert.Equal("true", printLayer.Attributes["isInitiallyVisible"]);
        Assert.Equal("false", printLayer.Attributes["isLocked"]);
        Assert.Equal("true", printLayer.Attributes["isInDefaultOrder"]);
        Assert.Equal("ON", printLayer.Attributes["viewState"]);
        Assert.Equal("ON", printLayer.Attributes["printState"]);
        Assert.Equal("OFF", printLayer.Attributes["exportState"]);
        Assert.Equal("OfficeIMO", printLayer.Attributes["usageCreator"]);
        Assert.Equal("Artwork", printLayer.Attributes["usageSubtype"]);

        OfficeDocumentMetadataEntry hiddenLayer = Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-optional-content-group-0001");
        Assert.Equal("Hidden layer", hiddenLayer.Name);
        Assert.Equal("false", hiddenLayer.Value);
        Assert.Equal("6", hiddenLayer.SourceObjectId);
        Assert.Equal("false", hiddenLayer.Attributes["isInitiallyVisible"]);
        Assert.Equal("true", hiddenLayer.Attributes["isLocked"]);
        Assert.Equal("ON", hiddenLayer.Attributes["exportState"]);

        using JsonDocument document = JsonDocument.Parse(result.ToJson());
        JsonElement jsonLayer = document.RootElement.GetProperty("metadata").EnumerateArray()
            .Single(entry => entry.GetProperty("id").GetString() == "pdf-optional-content-group-0001");
        Assert.Equal("pdf.optionalContent.group", jsonLayer.GetProperty("category").GetString());
        Assert.Equal("Hidden layer", jsonLayer.GetProperty("name").GetString());
        Assert.Equal("true", jsonLayer.GetProperty("attributes").GetProperty("isLocked").GetString());
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfDocument_ExposesXmpMetadata() {
        byte[] pdf = PdfDocument.Create(new PdfOptions()
                .SetPdfAIdentification(3, "B")
                .SetPdfUaIdentification()
                .SetElectronicInvoiceMetadata("EN 16931"))
            .Meta(title: "Reader XMP readback", author: "OfficeIMO", subject: "Reader metadata", keywords: "delta, epsilon")
            .Paragraph(p => p.Text("Reader generated XMP readback."))
            .ToBytes();

        OfficeDocumentReadResult result = PdfReaderAdapter.ReadDocument(
            new MemoryStream(pdf, writable: false),
            sourceName: "xmp-metadata.pdf");

        ReaderChunkDiagnostics diagnostics = Assert.Single(result.Chunks).Diagnostics!;
        Assert.True(diagnostics.HasXmpMetadata);
        Assert.Equal(0, diagnostics.OutputIntentCount);
        Assert.False(diagnostics.HasTaggedContent);

        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-xmp-metadata-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-xmp-pdfa-identification-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-xmp-pdfua-identification-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-xmp-electronic-invoice-metadata-count").Value);
        Assert.DoesNotContain(result.Metadata, metadata => metadata.Id == "pdf-xmp-unsupported-filter-count");

        OfficeDocumentMetadataEntry xmp = Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-xmp-metadata");
        Assert.Equal("pdf.xmp", xmp.Category);
        Assert.Equal("XmpMetadata", xmp.Name);
        Assert.Equal("Reader XMP readback", xmp.Value);
        Assert.NotNull(xmp.SourceObjectId);
        Assert.Equal("Reader XMP readback", xmp.Attributes["title"]);
        Assert.Equal("OfficeIMO", xmp.Attributes["creator"]);
        Assert.Equal("Reader metadata", xmp.Attributes["description"]);
        Assert.Equal("delta,epsilon", xmp.Attributes["subjects"]);
        Assert.Equal("OfficeIMO.Pdf", xmp.Attributes["producer"]);
        Assert.Equal("3", xmp.Attributes["pdfAPart"]);
        Assert.Equal("B", xmp.Attributes["pdfAConformance"]);
        Assert.Equal("1", xmp.Attributes["pdfUaPart"]);
        Assert.Equal("INVOICE", xmp.Attributes["electronicInvoiceDocumentType"]);
        Assert.Equal("factur-x.xml", xmp.Attributes["electronicInvoiceDocumentFileName"]);
        Assert.Equal("1.0", xmp.Attributes["electronicInvoiceVersion"]);
        Assert.Equal("EN 16931", xmp.Attributes["electronicInvoiceConformanceLevel"]);
        Assert.Equal("true", xmp.Attributes["isWellFormedXml"]);
        Assert.True(int.Parse(xmp.Attributes["streamSizeBytes"], System.Globalization.CultureInfo.InvariantCulture) > 0);
        Assert.True(int.Parse(xmp.Attributes["decodedSizeBytes"], System.Globalization.CultureInfo.InvariantCulture) > 0);

        using JsonDocument document = JsonDocument.Parse(result.ToJson());
        JsonElement jsonXmp = document.RootElement.GetProperty("metadata").EnumerateArray()
            .Single(entry => entry.GetProperty("id").GetString() == "pdf-xmp-metadata");
        Assert.Equal("pdf.xmp", jsonXmp.GetProperty("category").GetString());
        Assert.Equal("Reader XMP readback", jsonXmp.GetProperty("attributes").GetProperty("title").GetString());
        Assert.Equal("EN 16931", jsonXmp.GetProperty("attributes").GetProperty("electronicInvoiceConformanceLevel").GetString());
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfDocument_ExposesSecuritySignatureAndDssMetadata() {
        OfficeDocumentReadResult result = PdfReaderAdapter.ReadDocument(
            new MemoryStream(BuildSignedSecurityMetadataPdf(), writable: false),
            sourceName: "signed-security.pdf");

        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-security-state-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-security-signature-field-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-security-signature-count").Value);
        Assert.Equal("2", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-security-byte-range-segment-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-security-dss-vri-count").Value);
        Assert.Equal("8", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-security-dss-evidence-count").Value);
        Assert.Equal("2", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-security-startxref-count").Value);
        Assert.Equal("false", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-preflight-can-append-metadata-revision").Value);
        Assert.Equal("false", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-preflight-can-append-form-field-revision").Value);
        Assert.Equal("false", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-preflight-can-prepare-external-signature-revision").Value);

        OfficeDocumentMetadataEntry security = Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-security-state");
        Assert.Equal("pdf.security", security.Category);
        Assert.Equal("AppendOnlyRequired", security.Value);
        Assert.Equal("true", security.Attributes["hasSignatures"]);
        Assert.Equal("true", security.Attributes["acroFormAppendOnly"]);
        Assert.Equal("true", security.Attributes["hasDocMDPPermissions"]);
        Assert.Equal("2", security.Attributes["docMDPPermissionLevel"]);
        Assert.Equal("true", security.Attributes["hasUsageRights"]);
        Assert.Equal("6", security.Attributes["usageRightsObjectNumbers"]);
        Assert.Equal("true", security.Attributes["hasDocumentSecurityStore"]);
        Assert.Equal("true", security.Attributes["hasLongTermValidationEvidence"]);
        Assert.Equal("true", security.Attributes["requiresAppendOnlyMutation"]);
        Assert.Equal("true", security.Attributes["hasIncrementalUpdates"]);
        Assert.Equal("100,200", security.Attributes["startXrefOffsets"]);

        OfficeDocumentMetadataEntry signature = Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-security-signature-0000");
        Assert.Equal("pdf.security.signature", signature.Category);
        Assert.Equal("Approval", signature.Name);
        Assert.Equal("Alice", signature.Value);
        Assert.Equal("6", signature.SourceObjectId);
        Assert.Equal("5", signature.Attributes["fieldObjectNumber"]);
        Assert.Equal("Approval", signature.Attributes["fieldName"]);
        Assert.Equal("Adobe.PPKLite", signature.Attributes["filter"]);
        Assert.Equal("adbe.pkcs7.detached", signature.Attributes["subFilter"]);
        Assert.Equal("Alice", signature.Attributes["signerName"]);
        Assert.Equal("Warsaw", signature.Attributes["location"]);
        Assert.Equal("Approval", signature.Attributes["reason"]);
        Assert.Equal("0,10,20,30", signature.Attributes["byteRangeValues"]);
        Assert.Equal("Include", signature.Attributes["fieldLockAction"]);
        Assert.Equal("Total,Approver", signature.Attributes["fieldLockFields"]);
        Assert.Equal("Adobe.PPKLite", signature.Attributes["seedValueFilter"]);
        Assert.Equal("adbe.pkcs7.detached", signature.Attributes["seedValueSubFilters"]);
        Assert.Equal("SHA256,SHA512", signature.Attributes["seedValueDigestMethods"]);
        Assert.Equal("Approval,Final", signature.Attributes["seedValueReasons"]);
        Assert.Equal("true", signature.Attributes["seedValueAddRevInfo"]);
        Assert.Equal("2", signature.Attributes["seedValueMdpPermissionLevel"]);

        OfficeDocumentMetadataEntry dss = Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-security-dss");
        Assert.Equal("pdf.security.dss", dss.Category);
        Assert.Equal("DocumentSecurityStore", dss.Name);
        Assert.Equal("ABCDEF", dss.Value);
        Assert.Equal("9", dss.SourceObjectId);
        Assert.Equal("ABCDEF", dss.Attributes["vriKeys"]);
        Assert.Equal("10,11", dss.Attributes["certificateObjectNumbers"]);
        Assert.Equal("12", dss.Attributes["ocspObjectNumbers"]);
        Assert.Equal("13", dss.Attributes["crlObjectNumbers"]);
        Assert.Equal("10", dss.Attributes["vriCertificateObjectNumbers"]);
        Assert.Equal("12", dss.Attributes["vriOcspObjectNumbers"]);
        Assert.Equal("13", dss.Attributes["vriCrlObjectNumbers"]);
        Assert.Equal("14", dss.Attributes["timestampObjectNumbers"]);

        using JsonDocument document = JsonDocument.Parse(result.ToJson());
        JsonElement jsonSignature = document.RootElement.GetProperty("metadata").EnumerateArray()
            .Single(entry => entry.GetProperty("id").GetString() == "pdf-security-signature-0000");
        Assert.Equal("pdf.security.signature", jsonSignature.GetProperty("category").GetString());
        Assert.Equal("Alice", jsonSignature.GetProperty("attributes").GetProperty("signerName").GetString());
        Assert.Equal("SHA256,SHA512", jsonSignature.GetProperty("attributes").GetProperty("seedValueDigestMethods").GetString());

        JsonElement jsonAppendPolicy = document.RootElement.GetProperty("metadata").EnumerateArray()
            .Single(entry => entry.GetProperty("id").GetString() == "pdf-preflight-can-prepare-external-signature-revision");
        Assert.Equal("pdf.preflight.capability", jsonAppendPolicy.GetProperty("category").GetString());
        Assert.Equal("false", jsonAppendPolicy.GetProperty("value").GetString());
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfDocument_FiltersLogicalMetadataToSelectedPages() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                CreateOutlineFromHeadings = true
            })
            .H1("First")
            .Bookmark("FirstDest")
            .Paragraph(p => p.Text("First page body."))
            .PageBreak()
            .H1("Second")
            .Bookmark("SecondDest")
            .Paragraph(p => p.Text("Second page body."))
            .ToBytes();
        PdfLogicalDocument logical = PdfLogicalDocument.Load(pdf);

        OfficeDocumentReadResult result = PdfReaderAdapter.ReadDocument(
            logical,
            sourceName: "selected-metadata.pdf",
            pdfOptions: new ReaderPdfOptions {
                PageRanges = new[] { PdfPageRange.From(1, 1) }
            });

        Assert.Single(result.Pages);
        Assert.Contains(result.Metadata, entry => entry.Category == "pdf.destination" && entry.Name == "FirstDest");
        Assert.DoesNotContain(result.Metadata, entry => entry.Category == "pdf.destination" && entry.Name == "SecondDest");
        Assert.Contains(result.Metadata, entry => entry.Category == "pdf.outline" && entry.Name == "First");
        Assert.DoesNotContain(result.Metadata, entry => entry.Category == "pdf.outline" && entry.Name == "Second");
        Assert.Contains(result.Metadata, entry =>
            entry.Id == "pdf-named-destination-count" &&
            entry.Value == "1");
        Assert.Contains(result.Metadata, entry =>
            entry.Id == "pdf-outline-count" &&
            entry.Value == "1");
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfDocument_FiltersSelectedPageFormMetadata() {
        byte[] pdf = PdfDocument.Create()
            .H1("First")
            .Paragraph(p => p.Text("First page body."))
            .PageBreak()
            .H1("Second")
            .TextField("Second.Only", value: "hidden")
            .ToBytes();
        PdfLogicalDocument logical = PdfLogicalDocument.Load(pdf);

        OfficeDocumentReadResult result = PdfReaderAdapter.ReadDocument(
            logical,
            sourceName: "selected-form.pdf",
            pdfOptions: new ReaderPdfOptions {
                PageRanges = new[] { PdfPageRange.From(1, 1) }
            });

        Assert.Empty(result.Forms);
        Assert.DoesNotContain(result.Metadata, entry => entry.Id == "pdf-form-field-count");
        Assert.DoesNotContain(result.Metadata, entry => entry.Id == "pdf-form-widget-count");
        Assert.DoesNotContain(result.Metadata, entry => entry.Id == "pdf-form-widget-geometry-count");
        Assert.DoesNotContain(result.Metadata, entry => entry.Id == "pdf-form-widget-geometry-coverage");
        Assert.DoesNotContain(result.Metadata, entry => entry.Id == "pdf-form-text-count");
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfDocument_ExposesFormWidgetMetadata() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 240,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Paragraph(p => p.Text("Reader PDF form metadata."))
            .TextField("Contact.Email", value: "info@example.com", width: 180, height: 24)
            .ChoiceField("Contact.Country", new[] { "PL", "DE" }, value: "PL", width: 180, height: 24)
            .ToBytes();

        OfficeDocumentReadResult result = PdfReaderAdapter.ReadDocument(
            new MemoryStream(pdf, writable: false),
            sourceName: "form-metadata.pdf");

        Assert.Equal(2, result.Forms.Count);
        Assert.Equal("2", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-form-field-count").Value);
        Assert.Equal("2", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-form-widget-count").Value);
        Assert.Equal("2", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-form-widget-geometry-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-form-widget-geometry-coverage").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-form-text-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-form-choice-count").Value);
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfDocument_PreservesRemoteLinkDestinations() {
        OfficeDocumentReadResult result = PdfReaderAdapter.ReadDocument(
            new MemoryStream(BuildRemoteGoToLinkPdf(), writable: false),
            sourceName: "remote-link.pdf");

        OfficeDocumentLink link = Assert.Single(result.Links);
        Assert.Equal("remote", link.Kind);
        Assert.Equal("remote-report.pdf", link.RemoteFile);
        Assert.Null(link.RemoteDestinationName);
        Assert.Equal(2, link.RemoteDestinationPageNumber);
        Assert.Equal(nameof(PdfOpenActionDestinationMode.FitHorizontal), link.RemoteDestinationMode);
        Assert.Equal(144D, link.RemoteDestinationTop);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-link-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-link-geometry-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-link-geometry-coverage").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-link-remote-count").Value);

        using JsonDocument document = JsonDocument.Parse(result.ToJson());
        JsonElement jsonLink = document.RootElement.GetProperty("links")[0];
        Assert.Equal("remote-report.pdf", jsonLink.GetProperty("remoteFile").GetString());
        Assert.Equal(2, jsonLink.GetProperty("remoteDestinationPageNumber").GetInt32());
        Assert.Equal(nameof(PdfOpenActionDestinationMode.FitHorizontal), jsonLink.GetProperty("remoteDestinationMode").GetString());
        Assert.Equal(144D, jsonLink.GetProperty("remoteDestinationTop").GetDouble());
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfDocument_PreservesInternalLinkDestinationCoordinates() {
        OfficeDocumentReadResult result = PdfReaderAdapter.ReadDocument(
            new MemoryStream(BuildInternalDestinationLinkPdf(), writable: false),
            sourceName: "internal-link.pdf");

        OfficeDocumentLink link = Assert.Single(result.Links);
        Assert.Equal("destination", link.Kind);
        Assert.Equal(2, link.DestinationPageNumber);
        Assert.Equal(nameof(PdfOpenActionDestinationMode.Xyz), link.DestinationMode);
        Assert.Equal(24D, link.DestinationLeft);
        Assert.Equal(144D, link.DestinationTop);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-link-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-link-geometry-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-link-geometry-coverage").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-link-destination-count").Value);

        using JsonDocument document = JsonDocument.Parse(result.ToJson());
        JsonElement jsonLink = document.RootElement.GetProperty("links")[0];
        Assert.Equal(2, jsonLink.GetProperty("destinationPageNumber").GetInt32());
        Assert.Equal(nameof(PdfOpenActionDestinationMode.Xyz), jsonLink.GetProperty("destinationMode").GetString());
        Assert.Equal(24D, jsonLink.GetProperty("destinationLeft").GetDouble());
        Assert.Equal(144D, jsonLink.GetProperty("destinationTop").GetDouble());
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfDocument_ExposesAttachmentMetadata() {
        byte[] invoiceXml = Encoding.UTF8.GetBytes("<invoice>42</invoice>");
        byte[] sourceBytes = Encoding.UTF8.GetBytes("Source payload");
        byte[] pdf = PdfDocument.Create()
            .AttachFile("invoice.xml", invoiceXml, "application/xml", PdfAssociatedFileRelationship.Data, "Structured invoice XML")
            .AttachFile("source.txt", sourceBytes, "text/plain", PdfAssociatedFileRelationship.Source)
            .Paragraph(p => p.Text("Attachment metadata proof."))
            .ToBytes();

        OfficeDocumentReadResult result = PdfReaderAdapter.ReadDocument(
            new MemoryStream(pdf, writable: false),
            sourceName: "attachments.pdf");

        ReaderChunkDiagnostics diagnostics = Assert.Single(result.Chunks).Diagnostics!;
        Assert.Equal(2, diagnostics.AttachmentCount);
        Assert.False(diagnostics.HasXmpMetadata);
        Assert.Equal(0, diagnostics.OutputIntentCount);

        Assert.Equal("2", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-attachment-count").Value);
        Assert.DoesNotContain(result.Metadata, metadata => metadata.Id == "pdf-attachment-associated-count");
        Assert.Equal(
            (invoiceXml.Length + sourceBytes.Length).ToString(System.Globalization.CultureInfo.InvariantCulture),
            Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-attachment-total-size-bytes").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-attachment-relationship-data-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-attachment-relationship-source-count").Value);
        Assert.Equal("2", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-attachment-source-names-embeddedfiles-count").Value);

        OfficeDocumentMetadataEntry invoice = Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-attachment-0000");
        Assert.Equal("pdf.attachment", invoice.Category);
        Assert.Equal("invoice.xml", invoice.Name);
        Assert.Equal("invoice.xml", invoice.Value);
        Assert.NotNull(invoice.SourceObjectId);
        Assert.Equal("invoice.xml", invoice.Attributes["name"]);
        Assert.Equal("invoice.xml", invoice.Attributes["fileName"]);
        Assert.Equal("invoice.xml", invoice.Attributes["unicodeFileName"]);
        Assert.Equal("Structured invoice XML", invoice.Attributes["description"]);
        Assert.Equal("application/xml", invoice.Attributes["mimeType"]);
        Assert.Equal(nameof(PdfAssociatedFileRelationship.Data), invoice.Attributes["relationship"]);
        Assert.Equal("Names/EmbeddedFiles", invoice.Attributes["source"]);
        Assert.Equal("false", invoice.Attributes["isAssociatedFile"]);
        Assert.Equal(invoiceXml.Length.ToString(System.Globalization.CultureInfo.InvariantCulture), invoice.Attributes["sizeBytes"]);
        Assert.True(int.Parse(invoice.Attributes["fileSpecObjectNumber"], System.Globalization.CultureInfo.InvariantCulture) > 0);
        Assert.True(int.Parse(invoice.Attributes["embeddedFileObjectNumber"], System.Globalization.CultureInfo.InvariantCulture) > 0);

        OfficeDocumentMetadataEntry source = Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-attachment-0001");
        Assert.Equal("source.txt", source.Name);
        Assert.Equal("text/plain", source.Attributes["mimeType"]);
        Assert.Equal(nameof(PdfAssociatedFileRelationship.Source), source.Attributes["relationship"]);
        Assert.Equal(sourceBytes.Length.ToString(System.Globalization.CultureInfo.InvariantCulture), source.Attributes["sizeBytes"]);

        using JsonDocument document = JsonDocument.Parse(result.ToJson());
        JsonElement jsonAttachment = document.RootElement.GetProperty("metadata").EnumerateArray()
            .Single(entry => entry.GetProperty("id").GetString() == "pdf-attachment-0000");
        Assert.Equal("pdf.attachment", jsonAttachment.GetProperty("category").GetString());
        Assert.Equal("application/xml", jsonAttachment.GetProperty("attributes").GetProperty("mimeType").GetString());
        Assert.Equal(nameof(PdfAssociatedFileRelationship.Data), jsonAttachment.GetProperty("attributes").GetProperty("relationship").GetString());
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfStream_ReadsCompleteSeekableInputAndRestoresPosition() {
        byte[] pdf = BuildTwoPagePdf();
        using var stream = new MemoryStream(pdf, writable: false);
        stream.Position = 10;

        var chunks = PdfReaderAdapter.Read(stream, sourceName: "embedded.pdf").ToList();

        Assert.NotEmpty(chunks);
        Assert.Equal(10, stream.Position);
        Assert.Contains(chunks, c => (c.Markdown ?? c.Text).Contains("Reader PDF page one", StringComparison.Ordinal));
        Assert.Contains(chunks, c => (c.Markdown ?? c.Text).Contains("Reader PDF page two", StringComparison.Ordinal));
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfStream_MaxInputBytesUsesCompleteSeekableInput() {
        byte[] pdf = BuildTwoPagePdf();
        byte[] prefix = Encoding.ASCII.GetBytes("not-a-pdf-prefix-that-is-longer-than-the-limit");
        using var stream = new MemoryStream(prefix.Concat(pdf).ToArray(), writable: false);
        stream.Position = prefix.Length;

        Assert.Throws<IOException>(() => PdfReaderAdapter.Read(
                stream,
                sourceName: "embedded-limited.pdf",
                readerOptions: new ReaderOptions { MaxInputBytes = pdf.Length })
            .ToList());
        Assert.Equal(prefix.Length, stream.Position);
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfStream_DuplicatePageRangeSelectionsEmitUniqueChunkIds() {
        byte[] pdf = BuildTwoPagePdf();
        using var stream = new MemoryStream(pdf, writable: false);

        var chunks = PdfReaderAdapter.Read(
            stream,
            sourceName: "duplicate-ranges.pdf",
            pdfOptions: new ReaderPdfOptions {
                PageRanges = new[] {
                    PdfPageRange.From(1, 1),
                    PdfPageRange.From(1, 1)
                }
            }).ToList();

        Assert.Equal(2, chunks.Count);
        Assert.All(chunks, chunk => Assert.Equal(1, chunk.Location.Page));
        Assert.Equal(chunks.Count, chunks.Select(chunk => chunk.Id).Distinct(StringComparer.Ordinal).Count());
        Assert.Equal(chunks.Count, chunks.Select(chunk => chunk.Location.BlockAnchor).Distinct(StringComparer.Ordinal).Count());
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfStream_RendersUnsafeUriLinksAsInertMarkdown() {
        byte[] pdf = CreateLinkAnnotationPdf("javascript:alert(1)");
        using var stream = new MemoryStream(pdf, writable: false);

        var chunk = Assert.Single(PdfReaderAdapter.Read(
            stream,
            sourceName: "unsafe-link.pdf",
            pdfOptions: ReaderPdfOptions.CreateOfficeIMOProfile()).ToList());

        string markdown = chunk.Markdown ?? chunk.Text;
        Assert.DoesNotContain("](javascript:", markdown, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Link: Unsafe link -> javascript:alert(1)", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfDocument_ExposesPdfPreflightRewriteBlockers() {
        byte[] pdf = BuildActiveContentPdf();
        using var stream = new MemoryStream(pdf, writable: false);

        OfficeDocumentReadResult result = PdfReaderAdapter.ReadDocument(
            stream,
            sourceName: "active-content.pdf",
            readerOptions: new ReaderOptions { MaxChars = 8_000 });

        Assert.NotEmpty(result.Pages);
        Assert.Contains(result.Metadata, entry =>
            entry.Category == "pdf.preflight.capability" &&
            entry.Name == "CanRead" &&
            entry.Value == "true" &&
            entry.ValueType == "boolean");
        Assert.Contains(result.Metadata, entry =>
            entry.Category == "pdf.preflight.capability" &&
            entry.Name == "CanRewrite" &&
            entry.Value == "false" &&
            entry.ValueType == "boolean");
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Severity == OfficeDocumentDiagnosticSeverity.Warning &&
            diagnostic.Code == "pdf-rewrite-blocker" &&
            diagnostic.Message == "PDF active content is not supported for rewriting by OfficeIMO.Pdf yet.");

        string json = result.ToJson();
        using JsonDocument document = JsonDocument.Parse(json);
        Assert.Contains(document.RootElement.GetProperty("metadata").EnumerateArray(), entry =>
            entry.GetProperty("category").GetString() == "pdf.preflight.capability" &&
            entry.GetProperty("name").GetString() == "CanRewrite" &&
            entry.GetProperty("value").GetString() == "false" &&
            entry.GetProperty("valueType").GetString() == "boolean");
        Assert.Contains(document.RootElement.GetProperty("diagnostics").EnumerateArray(), diagnostic =>
            diagnostic.GetProperty("severity").GetString() == "Warning" &&
            diagnostic.GetProperty("code").GetString() == "pdf-rewrite-blocker" &&
            diagnostic.GetProperty("message").GetString() == "PDF active content is not supported for rewriting by OfficeIMO.Pdf yet.");
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfLogicalDocument_RangedReadRetainsCatalogActions() {
        PdfLogicalDocument logical = PdfLogicalDocument.Load(BuildTwoPageCatalogActionsPdf());

        ReaderChunk chunk = Assert.Single(PdfReaderAdapter.Read(
            logical,
            sourceName: "partial-catalog-actions.pdf",
            pdfOptions: new ReaderPdfOptions {
                PageRanges = new[] {
                    PdfPageRange.From(2, 2)
                }
            },
            readerOptions: new ReaderOptions { MaxChars = 8_000 }).ToList());

        Assert.Equal(2, chunk.Location.Page);
        Assert.NotNull(chunk.Diagnostics);
        Assert.True(chunk.Diagnostics!.HasCatalogActions);
        Assert.True(chunk.Diagnostics.HasActiveContent);
        Assert.Equal(1, chunk.Diagnostics.CatalogActionCount);
        Assert.NotNull(chunk.Actions);
        Assert.Contains(chunk.Actions!, action => action.Scope == ReaderActionScope.Catalog && action.IsPotentiallyUnsafe);
        Assert.Contains("Second catalog-safe page", chunk.Markdown ?? chunk.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("First catalog page", chunk.Markdown ?? chunk.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfLogicalDocument_DuplicateRangesRetainCatalogActionsOnlyOnce() {
        PdfLogicalDocument logical = PdfLogicalDocument.Load(BuildTwoPageCatalogActionsPdf());

        List<ReaderChunk> chunks = PdfReaderAdapter.Read(
            logical,
            sourceName: "catalog-actions-duplicate-ranges.pdf",
            pdfOptions: new ReaderPdfOptions {
                PageRanges = new[] {
                    PdfPageRange.From(2, 2),
                    PdfPageRange.From(2, 2)
                }
            }).ToList();

        Assert.Equal(2, chunks.Count);
        Assert.Contains(chunks[0].Actions!, action => action.Scope == ReaderActionScope.Catalog);
        Assert.DoesNotContain(chunks[1].Actions ?? Array.Empty<ReaderActionSummary>(), action => action.Scope == ReaderActionScope.Catalog);
        Assert.All(chunks, chunk => {
            Assert.NotNull(chunk.Diagnostics);
            Assert.Equal(1, chunk.Diagnostics!.CatalogActionCount);
            Assert.Equal(1, chunk.Diagnostics.PotentiallyUnsafeActionCount);
            Assert.Equal(1, chunk.Diagnostics.JavaScriptActionCount);
        });
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfStream_ExposesDetectedTables() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Paragraph(p => p.Text("Reader PDF table marker."))
            .Table(new[] {
                new[] { "Code", "Name", "Qty" },
                new[] { "A-100", "Alpha", "2" },
                new[] { "B-200", "Beta", "14" }
            }, style: new PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 70, 170, 60 },
                HeaderRowCount = 1,
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .ToBytes();
        using var stream = new MemoryStream(pdf, writable: false);

        var chunk = Assert.Single(PdfReaderAdapter.Read(
            stream,
            sourceName: "table.pdf",
            pdfOptions: new ReaderPdfOptions {
                LayoutOptions = new PdfTextLayoutOptions {
                    ForceSingleColumn = true
                }
            },
            readerOptions: new ReaderOptions { MaxChars = 8_000 }).ToList());

        Assert.NotNull(chunk.Diagnostics);
        Assert.Equal(1, chunk.Diagnostics!.TableCount);
        Assert.Equal(1, chunk.Diagnostics.TableGeometryCount);
        Assert.Equal(1D, chunk.Diagnostics.TableGeometryCoverage, 3);
        Assert.True(chunk.Diagnostics.MinTableConfidence >= 0.95D);
        Assert.True(chunk.Diagnostics.AverageTableConfidence >= 0.95D);
        Assert.Equal(0, chunk.Diagnostics.ImageCount);
        Assert.Equal(0D, chunk.Diagnostics.ImageGeometryCoverage);
        Assert.NotNull(chunk.Tables);
        var table = Assert.Single(chunk.Tables!);
        Assert.NotNull(table.Location);
        Assert.Equal(1, table.Location!.Page);
        Assert.Equal(0, table.Location.TableIndex);
        Assert.Equal("table", table.Location.SourceBlockKind);
        Assert.Equal("page-1-selection-0000-table-0", table.Location.BlockAnchor);
        Assert.Equal(new[] { "Code", "Name", "Qty" }, table.Columns);
        Assert.Equal(3, table.ColumnProfiles.Count);
        Assert.Equal(ReaderTableColumnKind.Text, table.ColumnProfiles[0].Kind);
        Assert.Equal(ReaderTableColumnKind.Text, table.ColumnProfiles[1].Kind);
        Assert.Equal(ReaderTableColumnKind.Numeric, table.ColumnProfiles[2].Kind);
        Assert.Equal("Qty", table.ColumnProfiles[2].Name);
        Assert.True(table.ColumnProfiles[2].IsNumeric);
        Assert.Equal(2, table.ColumnProfiles[2].NonEmptyCellCount);
        Assert.Equal(2, table.ColumnProfiles[2].NumericCellCount);
        Assert.Equal(1d, table.ColumnProfiles[2].Confidence);
        Assert.NotNull(table.Diagnostics);
        Assert.True(table.Diagnostics!.HasGeometry);
        Assert.True(table.Diagnostics.Width > 0);
        Assert.True(table.Diagnostics.Height > 0);
        Assert.Equal(1D, table.Diagnostics.SchemaConfidence, 3);
        Assert.Equal(1D, table.Diagnostics.CellCompleteness, 3);
        Assert.Equal(1D, table.Diagnostics.ColumnGeometryConfidence, 3);
        Assert.True(table.Diagnostics.Confidence >= 0.95D);
        Assert.Equal(2, table.TotalRowCount);
        Assert.False(table.Truncated);
        Assert.Equal(2, table.Rows.Count);
        Assert.Equal(new[] { "A-100", "Alpha", "2" }, table.Rows[0]);
        Assert.Equal(new[] { "B-200", "Beta", "14" }, table.Rows[1]);
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfDocument_ExposesTablesOnResultAndPage() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Paragraph(p => p.Text("Reader PDF table marker."))
            .Table(new[] {
                new[] { "Code", "Name", "Qty" },
                new[] { "A-100", "Alpha", "2" },
                new[] { "B-200", "Beta", "14" }
            }, style: new PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 70, 170, 60 },
                HeaderRowCount = 1,
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .ToBytes();
        using var stream = new MemoryStream(pdf, writable: false);

        OfficeDocumentReadResult result = PdfReaderAdapter.ReadDocument(
            stream,
            sourceName: "table-readback.pdf",
            pdfOptions: new ReaderPdfOptions {
                LayoutOptions = new PdfTextLayoutOptions {
                    ForceSingleColumn = true
                }
            },
            readerOptions: new ReaderOptions { MaxChars = 8_000 });

        ReaderTable table = Assert.Single(result.Tables);
        Assert.Equal(new[] { "Code", "Name", "Qty" }, table.Columns);
        Assert.Equal(2, table.TotalRowCount);
        Assert.Equal("table-readback.pdf", table.Location?.Path);
        OfficeDocumentPage page = Assert.Single(result.Pages);
        Assert.Same(table, Assert.Single(page.Tables));
        Assert.Equal("table-readback.pdf", page.Tables[0].Location?.Path);
        Assert.Contains(result.Blocks, block => block.Kind == "table" && block.Location.Page == 1);
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfTables_ReturnsLogicalTablesOnly() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Table(new[] {
                new[] { "Code", "Name", "Qty" },
                new[] { "A-100", "Alpha", "2" },
                new[] { "B-200", "Beta", "14" }
            }, style: new PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 70, 170, 60 },
                HeaderRowCount = 1,
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .ToBytes();
        using var stream = new MemoryStream(pdf, writable: false);

        IReadOnlyList<ReaderTable> tables = PdfReaderAdapter.ReadTables(
            stream,
            sourceName: "tables-only.pdf",
            pdfOptions: new ReaderPdfOptions {
                LayoutOptions = new PdfTextLayoutOptions {
                    ForceSingleColumn = true
                }
            });

        ReaderTable table = Assert.Single(tables);
        Assert.Equal("tables-only.pdf", table.Location?.Path);
        Assert.Equal(1, table.Location?.Page);
        Assert.Equal(new[] { "Code", "Name", "Qty" }, table.Columns);
        Assert.NotNull(table.Diagnostics);
        Assert.True(table.Diagnostics!.HasGeometry);
        Assert.True(table.Diagnostics.Confidence >= 0.95D);
        Assert.Equal(2, table.TotalRowCount);
        Assert.Equal(new[] { "B-200", "Beta", "14" }, table.Rows[1]);

        using var exportStream = new MemoryStream(pdf, writable: false);
        ReaderTableExportBundle export = Assert.Single(PdfReaderAdapter.ReadTableExports(
            exportStream,
            sourceName: "tables-only.pdf",
            pdfOptions: new ReaderPdfOptions {
                LayoutOptions = new PdfTextLayoutOptions {
                    ForceSingleColumn = true
                }
            }));
        Assert.Equal("tables-only-page-0001-table-0000", export.Id);
        Assert.Contains("A-100,Alpha,2", export.Csv, StringComparison.Ordinal);
        using JsonDocument exportJson = JsonDocument.Parse(export.Json);
        Assert.True(exportJson.RootElement.GetProperty("diagnostics").GetProperty("hasGeometry").GetBoolean());

        IReadOnlyList<ReaderTable> byteTables = PdfReaderAdapter.ReadTables(
            pdf,
            sourceName: "tables-bytes.pdf",
            pdfOptions: new ReaderPdfOptions {
                LayoutOptions = new PdfTextLayoutOptions {
                    ForceSingleColumn = true
                }
            });
        Assert.Equal("tables-bytes.pdf", Assert.Single(byteTables).Location?.Path);

        ReaderTableExportBundle byteExport = Assert.Single(PdfReaderAdapter.ReadTableExports(
            pdf,
            sourceName: "tables-bytes.pdf",
            pdfOptions: new ReaderPdfOptions {
                LayoutOptions = new PdfTextLayoutOptions {
                    ForceSingleColumn = true
                }
            }));
        Assert.Equal("tables-bytes-page-0001-table-0000", byteExport.Id);
        Assert.Contains("B-200,Beta,14", byteExport.Csv, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfStream_ExposesFormAndSourceDiagnostics() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 240,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Paragraph(p => p.Text("Reader PDF form marker."))
            .TextField("Contact.Email", value: "info@example.com", width: 180, height: 24)
            .ChoiceField("Contact.Country", new[] { "PL", "DE" }, value: "PL", width: 180, height: 24)
            .ToBytes();
        using var stream = new MemoryStream(pdf, writable: false);

        ReaderChunk chunk = Assert.Single(PdfReaderAdapter.Read(
            stream,
            sourceName: "form.pdf",
            pdfOptions: new ReaderPdfOptions {
                LayoutOptions = new PdfTextLayoutOptions {
                    ForceSingleColumn = true
                }
            },
            readerOptions: new ReaderOptions { MaxChars = 8_000 }).ToList());

        Assert.NotNull(chunk.Diagnostics);
        Assert.Equal(1, chunk.Diagnostics!.PageCount);
        Assert.Equal(1, chunk.Diagnostics.SelectedPageCount);
        Assert.Equal(1, chunk.Diagnostics.PageNumber);
        Assert.Equal(2, chunk.Diagnostics.FormFieldCount);
        Assert.Equal(2, chunk.Diagnostics.FormWidgetCount);
        Assert.Equal(2, chunk.Diagnostics.SelectedFormWidgetCount);
        Assert.False(chunk.Diagnostics.HasSecurityState);
        Assert.NotNull(chunk.FormFields);
        Assert.Equal(2, chunk.FormFields!.Count);

        ReaderFormField email = Assert.Single(chunk.FormFields, field => field.Name == "Contact.Email");
        Assert.Equal(ReaderFormFieldKind.Text, email.Kind);
        Assert.Equal("Tx", email.FieldType);
        Assert.Equal("info@example.com", email.Value);
        Assert.Equal(new[] { "info@example.com" }, email.Values);
        Assert.Equal(1, email.WidgetCount);
        Assert.Equal(new[] { 1 }, email.PageNumbers);
        ReaderFormWidget emailWidget = Assert.Single(email.Widgets);
        Assert.Equal(1, emailWidget.PageNumber);
        Assert.True(emailWidget.Width > 0);
        Assert.True(emailWidget.Height > 0);

        ReaderFormField country = Assert.Single(chunk.FormFields, field => field.Name == "Contact.Country");
        Assert.Equal(ReaderFormFieldKind.Choice, country.Kind);
        Assert.Equal("Ch", country.FieldType);
        Assert.Equal("PL", country.Value);
        Assert.Equal(new[] { "PL" }, country.Values);
        Assert.Equal(2, country.OptionCount);
        Assert.Equal(1, country.SelectedOptionCount);
        Assert.Equal(1, country.WidgetCount);
        Assert.Contains("Reader PDF form marker", chunk.Markdown ?? chunk.Text, StringComparison.Ordinal);

        using JsonDocument document = JsonDocument.Parse(new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Pdf,
            Chunks = new[] { chunk }
        }.ToJson());
        JsonElement jsonChunk = document.RootElement.GetProperty("chunks")[0];
        Assert.Equal(2, jsonChunk.GetProperty("formFields").GetArrayLength());
        Assert.Equal("Contact.Email", jsonChunk.GetProperty("formFields")[0].GetProperty("name").GetString());
        Assert.Equal(2, jsonChunk.GetProperty("diagnostics").GetProperty("formFieldCount").GetInt32());
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfStream_ExposesFormWidgetAppearanceStates() {
        byte[] pdf = BuildWidgetAppearanceFormPdf();
        using var stream = new MemoryStream(pdf, writable: false);

        ReaderChunk chunk = Assert.Single(PdfReaderAdapter.Read(
            stream,
            sourceName: "widget-appearance.pdf",
            readerOptions: new ReaderOptions { MaxChars = 8_000 }).ToList());

        Assert.NotNull(chunk.Diagnostics);
        Assert.Equal(1, chunk.Diagnostics!.FormFieldCount);
        Assert.Equal(1, chunk.Diagnostics.FormWidgetCount);
        Assert.Equal(1, chunk.Diagnostics.SelectedFormWidgetCount);
        Assert.Equal(1, chunk.Diagnostics.SelectedFormWidgetAppearanceStateCount);
        Assert.Equal(1D, chunk.Diagnostics.SelectedFormWidgetAppearanceStateCoverage, 3);
        Assert.Equal(2, chunk.Diagnostics.SelectedFormWidgetNormalAppearanceStateCount);

        Assert.NotNull(chunk.FormFields);
        ReaderFormField field = Assert.Single(chunk.FormFields!);
        Assert.Equal("AcceptTerms", field.Name);
        Assert.Equal(ReaderFormFieldKind.Button, field.Kind);
        Assert.Equal("Yes", field.Value);
        ReaderFormWidget widget = Assert.Single(field.Widgets);
        Assert.Equal("Yes", widget.AppearanceState);
        Assert.Equal(2, widget.NormalAppearanceStateCount);
        Assert.Equal(new[] { "Off", "Yes" }, widget.NormalAppearanceStates);
        Assert.True(widget.IsPrint);
        Assert.False(widget.IsHidden);
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfDocument_ExposesAcroFormXfaMetadataWithoutRenderingXfa() {
        byte[] pdf = BuildAcroFormXfaPdf();

        PdfLogicalDocument logical = PdfLogicalDocument.Load(pdf);
        Assert.NotNull(logical.AcroFormXfa);
        Assert.True(logical.HasAcroFormXfa);
        Assert.Equal("array", logical.AcroFormXfa!.ObjectKind);
        Assert.Equal(2, logical.AcroFormXfa.PacketCount);
        Assert.Equal(new[] { "template", "datasets" }, logical.AcroFormXfa.PacketNames);
        Assert.Equal(2, logical.AcroFormXfa.StreamCount);
        Assert.True(logical.AcroFormXfa.TotalPayloadBytes > 0);
        Assert.True(logical.AcroFormXfa.HasTemplatePacket);
        Assert.True(logical.AcroFormXfa.HasDatasetsPacket);

        PdfDocumentInfo info = PdfInspector.Inspect(pdf);
        Assert.True(info.HasForms);
        Assert.True(info.HasAcroFormXfa);
        Assert.Equal(2, info.AcroFormXfa!.PacketCount);

        OfficeDocumentReadResult result = PdfReaderAdapter.ReadDocument(
            new MemoryStream(pdf, writable: false),
            sourceName: "xfa-form.pdf",
            readerOptions: new ReaderOptions { MaxChars = 8_000 });

        Assert.Equal("true", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-acroform-xfa-present").Value);
        Assert.Equal("2", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-acroform-xfa-packet-count").Value);
        Assert.Equal("2", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-acroform-xfa-stream-count").Value);
        OfficeDocumentMetadataEntry xfa = Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-acroform-xfa");
        Assert.Equal("pdf.form.xfa", xfa.Category);
        Assert.Equal("array", xfa.Value);
        Assert.Equal("template,datasets", xfa.Attributes["packetNames"]);
        Assert.Equal("true", xfa.Attributes["hasTemplatePacket"]);
        Assert.Equal("true", xfa.Attributes["hasDatasetsPacket"]);

        string json = result.ToJson();
        using JsonDocument document = JsonDocument.Parse(json);
        JsonElement jsonXfa = document.RootElement.GetProperty("metadata").EnumerateArray()
            .Single(entry => entry.GetProperty("id").GetString() == "pdf-acroform-xfa");
        Assert.Equal("template,datasets", jsonXfa.GetProperty("attributes").GetProperty("packetNames").GetString());
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfStream_ChunkHashIncludesActionMetadata() {
        ReaderOptions readerOptions = new() {
            MaxChars = 8_000,
            ComputeHashes = true
        };

        ReaderChunk first = Assert.Single(PdfReaderAdapter.Read(
            new MemoryStream(BuildOpenActionHashPdf(120), writable: false),
            sourceName: "action-hash.pdf",
            readerOptions: readerOptions).ToList());
        ReaderChunk second = Assert.Single(PdfReaderAdapter.Read(
            new MemoryStream(BuildOpenActionHashPdf(160), writable: false),
            sourceName: "action-hash.pdf",
            readerOptions: readerOptions).ToList());

        Assert.Equal(first.Text, second.Text);
        Assert.NotNull(first.Actions);
        Assert.NotNull(second.Actions);
        Assert.Equal(120D, Assert.Single(first.Actions!).DestinationTop);
        Assert.Equal(160D, Assert.Single(second.Actions!).DestinationTop);
        Assert.False(string.IsNullOrWhiteSpace(first.ChunkHash));
        Assert.False(string.IsNullOrWhiteSpace(second.ChunkHash));
        Assert.NotEqual(first.ChunkHash, second.ChunkHash);
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfStream_ExposesImageVisualGeometry() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 300,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Paragraph(p => p.Text("Reader PDF visual marker."))
            .Image(PdfPngTestImages.CreateRgbPng(3, 2), 48, 32, alternativeText: "Reader visual badge")
            .ToBytes();
        using var stream = new MemoryStream(pdf, writable: false);

        ReaderChunk chunk = Assert.Single(PdfReaderAdapter.Read(
            stream,
            sourceName: "image-visual.pdf",
            readerOptions: new ReaderOptions { MaxChars = 8_000 }).ToList());

        Assert.NotNull(chunk.Diagnostics);
        Assert.Equal(1, chunk.Diagnostics!.ImageCount);
        Assert.Equal(1, chunk.Diagnostics.ImageGeometryCount);
        Assert.Equal(1D, chunk.Diagnostics.ImageGeometryCoverage, 3);
        Assert.Equal(0, chunk.Diagnostics.ImageNonAxisAlignedCount);
        Assert.Equal(0D, chunk.Diagnostics.ImageNonAxisAlignedCoverage);
        Assert.Equal(0, chunk.Diagnostics.TableCount);
        Assert.Equal(0, chunk.Diagnostics.TableGeometryCount);
        Assert.Equal(0D, chunk.Diagnostics.TableGeometryCoverage);
        Assert.Null(chunk.Diagnostics.MinTableConfidence);
        Assert.Null(chunk.Diagnostics.AverageTableConfidence);
        Assert.NotNull(chunk.Visuals);
        ReaderVisual visual = Assert.Single(chunk.Visuals!);
        Assert.Equal("image", visual.Kind);
        Assert.Equal("pdf-image", visual.Language);
        Assert.Equal(1, visual.Location!.Page);
        Assert.Equal("image", visual.Location.SourceBlockKind);
        Assert.Equal("page-1-selection-0000-image-0", visual.Location.BlockAnchor);
        Assert.False(string.IsNullOrWhiteSpace(visual.SourceName));
        Assert.Equal("image/png", visual.MimeType);
        Assert.Equal(3D, visual.Width);
        Assert.Equal(2D, visual.Height);
        Assert.Equal(48D, visual.PlacedWidth!.Value, 3);
        Assert.Equal(32D, visual.PlacedHeight!.Value, 3);
        Assert.Equal(1, visual.PlacementCount);
        Assert.True(visual.HasGeometry);
        Assert.True(visual.IsAxisAligned);
        Assert.False(string.IsNullOrWhiteSpace(visual.PayloadHash));
        Assert.Contains("Reader PDF visual marker", chunk.Markdown ?? chunk.Text, StringComparison.Ordinal);

        ReaderVisual extracted = Assert.Single(OfficeIMO.Reader.Tests.ReaderTestReaders.All.ExtractVisuals(new[] { chunk }));
        Assert.Equal("image-visual.pdf", extracted.Location?.Path);
        Assert.Equal("image/png", extracted.MimeType);
        Assert.Equal(48D, extracted.PlacedWidth!.Value, 3);
        Assert.Equal(32D, extracted.PlacedHeight!.Value, 3);
        Assert.Equal(1, extracted.PlacementCount);
        Assert.True(extracted.HasGeometry);
        Assert.True(extracted.IsAxisAligned);
        using JsonDocument visualJson = JsonDocument.Parse(extracted.ToJson());
        Assert.Equal("image/png", visualJson.RootElement.GetProperty("mimeType").GetString());
        Assert.True(visualJson.RootElement.GetProperty("hasGeometry").GetBoolean());
        Assert.True(visualJson.RootElement.GetProperty("isAxisAligned").GetBoolean());
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfStream_ExposesNonAxisAlignedImageDiagnostics() {
        byte[] pdf = BuildSkewedImagePdf();
        using var stream = new MemoryStream(pdf, writable: false);

        ReaderChunk chunk = Assert.Single(PdfReaderAdapter.Read(
            stream,
            sourceName: "skewed-image.pdf",
            readerOptions: new ReaderOptions { MaxChars = 8_000 }).ToList());

        Assert.NotNull(chunk.Diagnostics);
        Assert.Equal(1, chunk.Diagnostics!.ImageCount);
        Assert.Equal(1, chunk.Diagnostics.ImageGeometryCount);
        Assert.Equal(1D, chunk.Diagnostics.ImageGeometryCoverage, 3);
        Assert.Equal(1, chunk.Diagnostics.ImageNonAxisAlignedCount);
        Assert.Equal(1D, chunk.Diagnostics.ImageNonAxisAlignedCoverage, 3);

        ReaderVisual visual = Assert.Single(chunk.Visuals!);
        Assert.True(visual.HasGeometry);
        Assert.False(visual.IsAxisAligned);
        Assert.True(visual.PlacedWidth > 0D);
        Assert.True(visual.PlacedHeight > 0D);

        OfficeDocumentReadResult result = PdfReaderAdapter.ReadDocument(
            new MemoryStream(pdf, writable: false),
            sourceName: "skewed-image.pdf");

        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-image-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-image-geometry-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-image-non-axis-aligned-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-image-non-axis-aligned-coverage").Value);

        using JsonDocument resultJson = JsonDocument.Parse(result.ToJson());
        JsonElement diagnostics = resultJson.RootElement.GetProperty("chunks")[0].GetProperty("diagnostics");
        Assert.Equal(1, diagnostics.GetProperty("imageNonAxisAlignedCount").GetInt32());
        Assert.Equal(1D, diagnostics.GetProperty("imageNonAxisAlignedCoverage").GetDouble());
        Assert.False(resultJson.RootElement.GetProperty("visuals")[0].GetProperty("isAxisAligned").GetBoolean());
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfStream_ExposesActiveContentDiagnostics() {
        byte[] pdf = BuildPageAdditionalActionsPdf();
        using var stream = new MemoryStream(pdf, writable: false);

        ReaderChunk chunk = Assert.Single(PdfReaderAdapter.Read(
            stream,
            sourceName: "active.pdf",
            readerOptions: new ReaderOptions { MaxChars = 8_000 }).ToList());

        Assert.NotNull(chunk.Diagnostics);
        Assert.Equal(1, chunk.Diagnostics!.PageCount);
        Assert.Equal(1, chunk.Diagnostics.PageNumber);
        Assert.False(chunk.Diagnostics.HasOpenAction);
        Assert.False(chunk.Diagnostics.HasCatalogActions);
        Assert.True(chunk.Diagnostics.HasPageActions);
        Assert.False(chunk.Diagnostics.HasAnnotationActions);
        Assert.True(chunk.Diagnostics.HasActiveContent);
        Assert.Equal(0, chunk.Diagnostics.CatalogActionCount);
        Assert.Equal(2, chunk.Diagnostics.PageActionCount);
        Assert.Equal(2, chunk.Diagnostics.SelectedPageActionCount);
        Assert.Equal(0, chunk.Diagnostics.AnnotationActionCount);
        Assert.Equal(0, chunk.Diagnostics.SelectedAnnotationActionCount);
        Assert.NotNull(chunk.Actions);
        Assert.Equal(2, chunk.Actions!.Count);
        ReaderActionSummary openAction = Assert.Single(chunk.Actions, action => action.ActionPath == "O");
        Assert.Equal(ReaderActionScope.Page, openAction.Scope);
        Assert.Equal("JavaScript", openAction.ActionType);
        Assert.Equal("Page/AA", openAction.Source);
        Assert.Equal("O", openAction.TriggerName);
        Assert.Equal(1, openAction.PageNumber);
        Assert.False(openAction.IsChainedAction);
        ReaderActionSummary closeAction = Assert.Single(chunk.Actions, action => action.ActionPath == "C");
        Assert.Equal("Launch", closeAction.ActionType);
        Assert.Equal("C", closeAction.TriggerName);
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfStream_ExposesOpenAndCatalogActionSummaries() {
        byte[] pdf = BuildOpenAndCatalogActionsPdf();
        using var stream = new MemoryStream(pdf, writable: false);

        ReaderChunk chunk = Assert.Single(PdfReaderAdapter.Read(
            stream,
            sourceName: "open-catalog-actions.pdf",
            readerOptions: new ReaderOptions { MaxChars = 8_000 }).ToList());

        Assert.NotNull(chunk.Diagnostics);
        Assert.True(chunk.Diagnostics!.HasOpenAction);
        Assert.True(chunk.Diagnostics.HasCatalogActions);
        Assert.False(chunk.Diagnostics.HasPageActions);
        Assert.False(chunk.Diagnostics.HasAnnotationActions);
        Assert.True(chunk.Diagnostics.HasActiveContent);
        Assert.Equal(1, chunk.Diagnostics.CatalogActionCount);
        Assert.Equal(0, chunk.Diagnostics.PageActionCount);
        Assert.Equal(0, chunk.Diagnostics.SelectedPageActionCount);
        Assert.Equal(0, chunk.Diagnostics.AnnotationActionCount);
        Assert.Equal(0, chunk.Diagnostics.SelectedAnnotationActionCount);
        Assert.NotNull(chunk.Actions);
        Assert.Equal(2, chunk.Actions!.Count);

        ReaderActionSummary openAction = Assert.Single(chunk.Actions, action => action.Scope == ReaderActionScope.DocumentOpen);
        Assert.Equal("Destination", openAction.ActionType);
        Assert.Equal("OpenAction", openAction.Source);
        Assert.Equal(1, openAction.DestinationPageNumber);
        Assert.Equal("FitHorizontal", openAction.DestinationMode);
        Assert.Equal(180D, openAction.DestinationTop);
        Assert.Null(openAction.Name);
        Assert.Null(openAction.TriggerName);
        Assert.Null(openAction.ActionPath);

        ReaderActionSummary catalogAction = Assert.Single(chunk.Actions, action => action.Scope == ReaderActionScope.Catalog);
        Assert.Equal("JavaScript", catalogAction.ActionType);
        Assert.Equal("Names/JavaScript", catalogAction.Source);
        Assert.Equal("Startup", catalogAction.Name);
        Assert.Null(catalogAction.TriggerName);
        Assert.Null(catalogAction.ActionPath);
        Assert.Null(catalogAction.DestinationPageNumber);
        Assert.DoesNotContain("app.alert", catalogAction.Name ?? string.Empty, StringComparison.Ordinal);
        Assert.DoesNotContain("app.alert", catalogAction.ActionType, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfStream_ExposesAnnotationActionSummariesWithoutPayloads() {
        byte[] pdf = BuildAnnotationActionsPdf();
        using var stream = new MemoryStream(pdf, writable: false);

        ReaderChunk chunk = Assert.Single(PdfReaderAdapter.Read(
            stream,
            sourceName: "annotation-actions.pdf",
            readerOptions: new ReaderOptions { MaxChars = 8_000 }).ToList());

        Assert.NotNull(chunk.Diagnostics);
        Assert.False(chunk.Diagnostics!.HasCatalogActions);
        Assert.False(chunk.Diagnostics.HasPageActions);
        Assert.True(chunk.Diagnostics.HasAnnotationActions);
        Assert.True(chunk.Diagnostics.HasActiveContent);
        Assert.Equal(3, chunk.Diagnostics.AnnotationActionCount);
        Assert.Equal(3, chunk.Diagnostics.SelectedAnnotationActionCount);
        Assert.NotNull(chunk.Actions);
        Assert.Equal(3, chunk.Actions!.Count);

        ReaderActionSummary primaryAction = Assert.Single(chunk.Actions, action => action.Source == "Annotation/A");
        Assert.Equal(ReaderActionScope.Annotation, primaryAction.Scope);
        Assert.Equal("JavaScript", primaryAction.ActionType);
        Assert.Equal("Link", primaryAction.Name);
        Assert.Equal("A", primaryAction.ActionPath);
        Assert.Equal(1, primaryAction.PageNumber);
        Assert.False(primaryAction.IsChainedAction);

        ReaderActionSummary additionalAction = Assert.Single(chunk.Actions, action => action.Source == "Annotation/AA");
        Assert.Equal("Launch", additionalAction.ActionType);
        Assert.Equal("Text", additionalAction.Name);
        Assert.Equal("E", additionalAction.TriggerName);
        Assert.Equal("AA.E", additionalAction.ActionPath);
        Assert.Equal(1, additionalAction.PageNumber);
        Assert.False(additionalAction.IsChainedAction);

        ReaderActionSummary chainedAction = Assert.Single(chunk.Actions, action => action.Source == "Annotation/Next");
        Assert.Equal("Launch", chainedAction.ActionType);
        Assert.Equal("A", chainedAction.TriggerName);
        Assert.Equal("A.Next", chainedAction.ActionPath);
        Assert.True(chainedAction.IsChainedAction);

        string markdown = chunk.Markdown ?? chunk.Text;
        Assert.DoesNotContain("app.alert", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("tool.exe", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("chain.exe", markdown, StringComparison.Ordinal);
        Assert.All(chunk.Actions, action => {
            Assert.DoesNotContain("app.alert", action.ActionType, StringComparison.Ordinal);
            Assert.DoesNotContain("tool.exe", action.ActionType, StringComparison.Ordinal);
            Assert.DoesNotContain("chain.exe", action.ActionType, StringComparison.Ordinal);
        });
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfStream_ClassifiesRichMediaAnnotationActionsAsUnsafe() {
        byte[] pdf = BuildRichMediaAnnotationActionPdf();
        using var stream = new MemoryStream(pdf, writable: false);

        ReaderChunk chunk = Assert.Single(PdfReaderAdapter.Read(
            stream,
            sourceName: "rich-media-action.pdf",
            readerOptions: new ReaderOptions { MaxChars = 8_000 }).ToList());

        ReaderActionSummary action = Assert.Single(chunk.Actions!);
        Assert.Equal("RichMedia", action.ActionType);
        Assert.True(action.IsPotentiallyUnsafe);
        Assert.True(chunk.Diagnostics!.HasActiveContent);
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfDocument_FlagsImageOnlyPagesAsOcrCandidates() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 300,
                PageHeight = 220,
                MarginLeft = 24,
                MarginRight = 24,
                MarginTop = 24,
                MarginBottom = 24
            })
            .Image(CreateMinimalRgbPng(), 180, 120)
            .ToBytes();
        using var stream = new MemoryStream(pdf, writable: false);

        OfficeDocumentReadResult result = PdfReaderAdapter.ReadDocument(
            stream,
            sourceName: "scan-candidate.pdf",
            readerOptions: new ReaderOptions { MaxChars = 8_000 });

        OfficeDocumentOcrCandidate candidate = Assert.Single(result.OcrCandidates);
        OfficeDocumentAsset asset = Assert.Single(result.Assets);
        Assert.Equal(OfficeDocumentAssetNaming.BuildFileName(asset.Id, asset.Extension), asset.FileName);
        Assert.NotNull(asset.PayloadBytes);
        Assert.NotNull(asset.Region);
        Assert.True(asset.Region!.Width > 0);
        Assert.True(asset.Region.Height > 0);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-image-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-image-geometry-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-image-geometry-coverage").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-ocr-candidate-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-ocr-image-candidate-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-ocr-asset-linked-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-ocr-candidate-geometry-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-ocr-candidate-geometry-coverage").Value);
        Assert.Equal("image", candidate.Kind);
        Assert.Equal(1, candidate.Location.Page);
        Assert.Equal(1, candidate.ImageCount);
        Assert.Equal(0, candidate.TextBlockCount);
        Assert.Equal(asset.Id, candidate.AssetId);
        Assert.NotNull(candidate.Region);
        Assert.True(candidate.Region!.Width > 0);
        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "ocr-needed" &&
            diagnostic.Location?.Page == 1);
        OfficeDocumentPage page = Assert.Single(result.Pages);
        Assert.Same(candidate, Assert.Single(page.OcrCandidates));

        string json = result.ToJson();
        using JsonDocument document = JsonDocument.Parse(json);
        JsonElement root = document.RootElement;
        Assert.Equal(OfficeDocumentReadResultSchema.Id, root.GetProperty("schemaId").GetString());
        Assert.Equal(OfficeDocumentReadResultSchema.CurrentVersion, root.GetProperty("schemaVersion").GetInt32());
        Assert.Equal("ocr-needed", root.GetProperty("diagnostics")[0].GetProperty("code").GetString());
        Assert.Equal(asset.FileName, root.GetProperty("assets")[0].GetProperty("fileName").GetString());
        Assert.True(root.GetProperty("assets")[0].GetProperty("region").GetProperty("width").GetDouble() > 0D);
        Assert.Equal("image", root.GetProperty("ocrCandidates")[0].GetProperty("kind").GetString());
        Assert.Equal("image", root.GetProperty("pages")[0].GetProperty("ocrCandidates")[0].GetProperty("kind").GetString());
    }

    [Fact]
    public void DocumentReaderPdf_ApplyOcrResults_EnrichesImageOnlyPagesWithoutRunningOcrInCore() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 300,
                PageHeight = 220,
                MarginLeft = 24,
                MarginRight = 24,
                MarginTop = 24,
                MarginBottom = 24
            })
            .Image(CreateMinimalRgbPng(), 180, 120)
            .ToBytes();
        using var stream = new MemoryStream(pdf, writable: false);
        OfficeDocumentReadResult result = PdfReaderAdapter.ReadDocument(
            stream,
            sourceName: "scan-candidate.pdf",
            readerOptions: new ReaderOptions { MaxChars = 8_000 });
        OfficeDocumentOcrCandidate candidate = Assert.Single(result.OcrCandidates);

        OfficeDocumentOcrEnrichmentResult enrichment = result.ApplyOcrResults(new[] {
            new OfficeDocumentOcrTextResult {
                CandidateId = candidate.Id,
                Text = "Invoice 1042\nTotal 123.45 EUR",
                Confidence = 0.97D,
                Language = "en",
                Provider = "external-ocr-contract",
                Model = "fixture"
            }
        });

        OfficeDocumentReadResult enriched = enrichment.Document;
        Assert.Equal(1, enrichment.Report.CandidateCount);
        Assert.Equal(1, enrichment.Report.ResultCount);
        Assert.Equal(1, enrichment.Report.AppliedResultCount);
        Assert.Equal(0, enrichment.Report.UnresolvedCandidateCount);
        Assert.Equal(0, enrichment.Report.UnmatchedResultCount);
        Assert.Equal(1, enrichment.Report.EnrichedBlockCount);
        Assert.Equal(1, enrichment.Report.EnrichedChunkCount);
        Assert.Equal(candidate.Id, Assert.Single(enrichment.Report.AppliedCandidateIds));
        Assert.Empty(enriched.OcrCandidates);
        Assert.DoesNotContain(enriched.Diagnostics, diagnostic => diagnostic.Code == "ocr-needed");
        Assert.Contains("officeimo.reader.ocr-enrichment", enriched.CapabilitiesUsed);
        Assert.Contains("Invoice 1042", enriched.Markdown, StringComparison.Ordinal);

        OfficeDocumentBlock block = Assert.Single(enriched.Blocks, item => item.Kind == "ocr-text");
        Assert.Equal("Invoice 1042\nTotal 123.45 EUR", block.Text);
        Assert.Equal(candidate.Location.Page, block.Location.Page);
        Assert.Equal(candidate.Region!.Width, block.Region!.Width);
        ReaderChunk chunk = Assert.Single(enriched.Chunks, item => item.Id == candidate.Id + "-chunk");
        Assert.Equal(block.Text, chunk.Text);
        Assert.Equal(ReaderInputKind.Pdf, chunk.Kind);
        OfficeDocumentPage page = Assert.Single(enriched.Pages);
        Assert.Contains(page.Blocks, item => item.Id == block.Id);
        Assert.Empty(page.OcrCandidates);

        Assert.Equal("1", Assert.Single(enriched.Metadata, metadata => metadata.Id == "reader-ocr-applied-count").Value);
        Assert.Equal("0", Assert.Single(enriched.Metadata, metadata => metadata.Id == "reader-ocr-unresolved-candidate-count").Value);
        OfficeDocumentMetadataEntry applied = Assert.Single(enriched.Metadata, metadata => metadata.Id == "reader-ocr-applied-0001");
        Assert.Equal(candidate.Id, applied.Value);
        Assert.Equal("external-ocr-contract", applied.Attributes["provider"]);
        Assert.Equal("fixture", applied.Attributes["model"]);
        Assert.Equal("en", applied.Attributes["language"]);
        Assert.Equal("0.97", applied.Attributes["confidence"]);

        using JsonDocument document = JsonDocument.Parse(enriched.ToJson());
        JsonElement root = document.RootElement;
        Assert.Equal("Invoice 1042\nTotal 123.45 EUR", root.GetProperty("blocks").EnumerateArray().Single(item => item.GetProperty("kind").GetString() == "ocr-text").GetProperty("text").GetString());
        Assert.Empty(root.GetProperty("ocrCandidates").EnumerateArray());
    }

    [Fact]
    public void DocumentReaderPdf_ApplyOcrResults_KeepsDiagnosticsForUnresolvedCandidatesInSameContainer() {
        var location = new ReaderLocation { Page = 1, Path = "scan.pdf" };
        var result = new OfficeDocumentReadResult {
            Source = new OfficeDocumentSource { Path = "scan.pdf" },
            OcrCandidates = new[] {
                new OfficeDocumentOcrCandidate { Id = "ocr-1", Kind = "image", Reason = "First image needs OCR.", Location = location },
                new OfficeDocumentOcrCandidate { Id = "ocr-2", Kind = "image", Reason = "Second image needs OCR.", Location = location }
            },
            Diagnostics = new[] {
                new OfficeDocumentDiagnostic { Code = "ocr-needed", Message = "First image needs OCR.", Location = location },
                new OfficeDocumentDiagnostic { Code = "ocr-needed", Message = "Second image needs OCR.", Location = location }
            }
        };

        OfficeDocumentOcrEnrichmentResult enrichment = result.ApplyOcrResults(new[] {
            new OfficeDocumentOcrTextResult { CandidateId = "ocr-1", Text = "Resolved text" }
        });

        OfficeDocumentReadResult enriched = enrichment.Document;
        OfficeDocumentOcrCandidate unresolved = Assert.Single(enriched.OcrCandidates);
        Assert.Equal("ocr-2", unresolved.Id);
        OfficeDocumentDiagnostic diagnostic = Assert.Single(enriched.Diagnostics, item => item.Code == "ocr-needed");
        Assert.Equal("Second image needs OCR.", diagnostic.Message);
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfStream_DuplicatePageRangeTableSelectionsEmitUniqueAnchors() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Paragraph(p => p.Text("Reader PDF table marker."))
            .Table(new[] {
                new[] { "Code", "Name", "Qty" },
                new[] { "A-100", "Alpha", "2" },
                new[] { "B-200", "Beta", "14" }
            }, style: new PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 70, 170, 60 },
                HeaderRowCount = 1,
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .ToBytes();
        using var stream = new MemoryStream(pdf, writable: false);

        var chunks = PdfReaderAdapter.Read(
            stream,
            sourceName: "duplicate-table-ranges.pdf",
            pdfOptions: new ReaderPdfOptions {
                LayoutOptions = new PdfTextLayoutOptions {
                    ForceSingleColumn = true
                },
                PageRanges = new[] {
                    PdfPageRange.From(1, 1),
                    PdfPageRange.From(1, 1)
                }
            },
            readerOptions: new ReaderOptions { MaxChars = 8_000 }).ToList();

        var tableAnchors = chunks
            .Where(chunk => chunk.Tables?.Count > 0)
            .Select(chunk => Assert.Single(chunk.Tables!).Location!.BlockAnchor)
            .ToArray();

        Assert.Equal(new[] {
            "page-1-selection-0000-table-0",
            "page-1-selection-0001-table-0"
        }, tableAnchors);
        Assert.Equal(tableAnchors.Length, tableAnchors.Distinct(StringComparer.Ordinal).Count());
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfTableExports_DuplicatePageRangeTableSelectionsEmitUniqueIds() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Paragraph(p => p.Text("Reader PDF table marker."))
            .Table(new[] {
                new[] { "Code", "Name", "Qty" },
                new[] { "A-100", "Alpha", "2" },
                new[] { "B-200", "Beta", "14" }
            }, style: new PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 70, 170, 60 },
                HeaderRowCount = 1,
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .ToBytes();
        using var stream = new MemoryStream(pdf, writable: false);

        IReadOnlyList<ReaderTableExportBundle> exports = PdfReaderAdapter.ReadTableExports(
            stream,
            sourceName: "duplicate-table-ranges.pdf",
            pdfOptions: new ReaderPdfOptions {
                LayoutOptions = new PdfTextLayoutOptions {
                    ForceSingleColumn = true
                },
                PageRanges = new[] {
                    PdfPageRange.From(1, 1),
                    PdfPageRange.From(1, 1)
                }
            },
            readerOptions: new ReaderOptions { MaxChars = 8_000 });

        Assert.Equal(2, exports.Count);
        Assert.Equal(exports.Count, exports.Select(export => export.Id).Distinct(StringComparer.Ordinal).Count());
        Assert.Equal(new[] {
            "duplicate-table-ranges-page-0001-table-0000",
            "duplicate-table-ranges-page-0001-selection-0001-table-0000"
        }, exports.Select(export => export.Id).ToArray());
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfStream_TableRowCapsApplyAfterDetectedHeader() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Table(new[] {
                new[] { "Code", "Name", "Qty" },
                new[] { "A-100", "Alpha", "2" },
                new[] { "B-200", "Beta", "14" }
            }, style: new PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 70, 170, 60 },
                HeaderRowCount = 1,
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .ToBytes();
        using var stream = new MemoryStream(pdf, writable: false);

        var chunk = Assert.Single(PdfReaderAdapter.Read(
            stream,
            sourceName: "table-row-cap.pdf",
            pdfOptions: new ReaderPdfOptions {
                LayoutOptions = new PdfTextLayoutOptions {
                    ForceSingleColumn = true
                }
            },
            readerOptions: new ReaderOptions {
                MaxChars = 8_000,
                MaxTableRows = 1
            }).ToList());

        Assert.NotNull(chunk.Tables);
        var table = Assert.Single(chunk.Tables!);
        Assert.Equal(new[] { "Code", "Name", "Qty" }, table.Columns);
        Assert.Equal(2, table.TotalRowCount);
        Assert.True(table.Truncated);
        Assert.Single(table.Rows);
        Assert.Equal(new[] { "A-100", "Alpha", "2" }, table.Rows[0]);
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfStream_ExposesHeaderlessKeyValueTables() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .KeyValueTable(new[] {
                PdfKeyValueRow.Text("InvoiceId", "INV-001"),
                PdfKeyValueRow.Text("Customer", "Evotec"),
                PdfKeyValueRow.Text("Due", "2026-06-30")
            }, style: new PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 120, 170 },
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .ToBytes();
        using var stream = new MemoryStream(pdf, writable: false);

        var chunk = Assert.Single(PdfReaderAdapter.Read(
            stream,
            sourceName: "invoice-facts.pdf",
            pdfOptions: new ReaderPdfOptions {
                LayoutOptions = new PdfTextLayoutOptions {
                    ForceSingleColumn = true
                }
            },
            readerOptions: new ReaderOptions { MaxChars = 8_000 }).ToList());

        Assert.NotNull(chunk.Tables);
        var table = Assert.Single(chunk.Tables!);
        Assert.NotNull(table.Location);
        Assert.Equal(1, table.Location!.Page);
        Assert.Equal(0, table.Location.TableIndex);
        Assert.Equal(new[] { "Key", "Value" }, table.Columns);
        Assert.Equal(3, table.TotalRowCount);
        Assert.False(table.Truncated);
        Assert.Equal(new[] { "InvoiceId", "INV-001" }, table.Rows[0]);
        Assert.Equal(new[] { "Customer", "Evotec" }, table.Rows[1]);
        Assert.Equal(new[] { "Due", "2026-06-30" }, table.Rows[2]);
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfStream_SplitsByMaxChars() {
        string longText = string.Join(" ", Enumerable.Repeat("Reader PDF split marker", 40));
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 620,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Paragraph(p => p.Text(longText))
            .ToBytes();
        using var stream = new MemoryStream(pdf, writable: false);

        var chunks = PdfReaderAdapter.Read(
            stream,
            sourceName: "split.pdf",
            readerOptions: new ReaderOptions { MaxChars = 96 }).ToList();

        Assert.True(chunks.Count > 1);
        Assert.All(chunks, c => Assert.Equal(ReaderInputKind.Pdf, c.Kind));
        Assert.Contains(chunks, c =>
            c.Warnings != null &&
            c.Warnings.Any(w => w.Contains("split due to MaxChars", StringComparison.OrdinalIgnoreCase)));
    }

    [Fact]
    public void DocumentReaderPdf_BuilderHandler_DispatchesPdfStream() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddPdfHandler().Build();

        byte[] pdf = BuildTwoPagePdf();
        using var stream = new MemoryStream(pdf, writable: false);
        var chunks = reader.Read(stream, "registry.pdf").ToList();

        Assert.NotEmpty(chunks);
        Assert.Contains(chunks, c =>
            c.Kind == ReaderInputKind.Pdf &&
            string.Equals(c.Location.Path, "registry.pdf", StringComparison.OrdinalIgnoreCase) &&
            (c.Markdown ?? c.Text).Contains("Reader PDF page one", StringComparison.Ordinal));

        stream.Position = 0;
        OfficeDocumentReadResult result = reader.ReadDocument(stream, "registry.pdf");
        Assert.Contains("officeimo.pdf.logical-document", result.CapabilitiesUsed);
        Assert.Equal(2, result.Pages.Count);
        Assert.NotEmpty(result.Blocks);

        ReaderHandlerCapability capability = Assert.Single(
            reader.GetCapabilities(),
            item => item.Id == OfficeDocumentReaderBuilderPdfExtensions.HandlerId);
        Assert.True(capability.SupportsDocumentPath);
        Assert.True(capability.SupportsDocumentStream);
    }

    [Fact]
    public void DocumentReaderPdf_BuilderManifest_DescribesConfiguredHandler() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddPdfHandler().Build();
        ReaderCapabilityManifest manifest = reader.GetCapabilityManifest();

        Assert.Contains(manifest.Handlers, handler =>
                handler.Id == OfficeDocumentReaderBuilderPdfExtensions.HandlerId &&
                handler.Kind == ReaderInputKind.Pdf &&
                handler.Extensions.Contains(".pdf"));
        Assert.Equal(1, manifest.Handlers.Count(handler =>
            string.Equals(handler.Id, OfficeDocumentReaderBuilderPdfExtensions.HandlerId, StringComparison.Ordinal)));
        Assert.Contains(manifest.Handlers, handler =>
            handler.Origin == ReaderHandlerOrigin.OfficeIMO &&
            handler.Extensions.Contains(".pdf", StringComparer.OrdinalIgnoreCase));
        Assert.Contains(
            OfficeDocumentReaderBuilderPdfExtensions.HandlerId,
            reader.GetCapabilityManifestJson(),
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfStream_NonSeekable_EnforcesMaxInputBytes() {
        byte[] pdf = BuildTwoPagePdf();
        using var stream = new NonSeekableReadStream(pdf);

        var ex = Assert.Throws<IOException>(() => PdfReaderAdapter.Read(
            stream,
            sourceName: "nonseekable.pdf",
            readerOptions: new ReaderOptions { MaxInputBytes = 16 }).ToList());

        Assert.Contains("Input exceeds MaxInputBytes", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ReaderPdfOptions_CloneCopiesNestedOptionsIndependently() {
        var options = ReaderPdfOptions.CreateOfficeIMOProfile();
        options.LayoutOptions = new PdfTextLayoutOptions {
            ForceSingleColumn = true,
            MarginLeft = 12
        };
        options.PageRanges = new[] { PdfPageRange.From(1, 1) };
        options.MarkdownOptions!.AlignNumericTableColumns = false;

        var clone = options.Clone();

        Assert.NotSame(options, clone);
        Assert.NotSame(options.LayoutOptions, clone.LayoutOptions);
        Assert.NotSame(options.MarkdownOptions, clone.MarkdownOptions);
        Assert.Equal(options.LayoutOptions.MarginLeft, clone.LayoutOptions!.MarginLeft);
        Assert.Equal(options.MarkdownOptions!.IncludeLinkAnnotations, clone.MarkdownOptions!.IncludeLinkAnnotations);
        Assert.False(clone.MarkdownOptions.AlignNumericTableColumns);
        Assert.Equal(options.PageRanges.Single(), clone.PageRanges!.Single());

        clone.LayoutOptions.MarginLeft = 48;
        clone.MarkdownOptions.IncludeLinkAnnotations = false;
        clone.MarkdownOptions.AlignNumericTableColumns = true;

        Assert.Equal(12, options.LayoutOptions.MarginLeft);
        Assert.True(options.MarkdownOptions.IncludeLinkAnnotations);
        Assert.False(options.MarkdownOptions.AlignNumericTableColumns);
    }

    [Fact]
    public void ReaderPdf_ProfileContract_DescribesConfiguredHandlerAndChunkShape() {
        ReaderPdfProfileContract contract = ReaderPdfProfileContracts.OfficeIMO;

        Assert.Equal(OfficeDocumentReaderBuilderPdfExtensions.HandlerId, contract.Id);
        Assert.Contains("OfficeIMO.Pdf logical model", contract.Pipeline, StringComparison.Ordinal);
        Assert.Contains("page-aware locations", contract.OutputContract, StringComparison.Ordinal);
        Assert.Contains("MaxChars", contract.OutputContract, StringComparison.Ordinal);
        Assert.Contains("Reader input limits", contract.SafetyContract, StringComparison.Ordinal);
        Assert.Contains("unsafe links", contract.SafetyContract, StringComparison.Ordinal);
        Assert.Contains("scanned PDFs require OCR", contract.UnsupportedScope, StringComparison.Ordinal);
    }

    private static byte[] BuildTwoPagePdf() {
        return PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .H1("Reader PDF page one")
            .Paragraph(p => p.Text("First page body."))
            .PageBreak()
            .H1("Reader PDF page two")
            .Paragraph(p => p.Text("Second page body."))
            .ToBytes();
    }

    private static byte[] CreateMinimalRgbPng() => PdfPngTestImages.CreateRgbPng(1, 1);

    private static byte[] BuildSkewedImagePdf() {
        string content = "q\n48 18 16 32 80 120 cm\n/Im1 Do\nQ\n";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 260 220] /Resources << /XObject << /Im1 5 0 R >> >> /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length " + Encoding.ASCII.GetByteCount(content).ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>",
            "stream",
            content.TrimEnd('\n'),
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length 3 >>",
            "stream",
            "abc",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] CreateLinkAnnotationPdf(string uri) {
        string escapedUri = uri.Replace("\\", "\\\\").Replace("(", "\\(").Replace(")", "\\)");
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 320 220] /Annots [4 0 R] >>",
            "endobj",
            "4 0 obj",
            $"<< /Type /Annot /Subtype /Link /Rect [40 160 180 182] /Contents (Unsafe link) /A << /S /URI /URI ({escapedUri}) >> >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildRemoteGoToLinkPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R /Annots [5 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Annot /Subtype /Link /Rect [10 20 90 42] /Contents (Remote report link) /A << /S /GoToR /F << /F (fallback.pdf) /UF (remote-report.pdf) >> /D [1 /FitH 144] >> >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildInternalDestinationLinkPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 2 /Kids [3 0 R 4 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 5 0 R /Annots [6 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 7 0 R >>",
            "endobj",
            "5 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "6 0 obj",
            "<< /Type /Annot /Subtype /Link /Rect [10 20 90 42] /Contents (Internal report link) /Dest [4 0 R /XYZ 24 144 1] >>",
            "endobj",
            "7 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 8 >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildActiveContentPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Names << /JavaScript << /Names [(Open) 5 0 R] >> >> >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /S /JavaScript /JS (app.alert('OfficeIMO')) >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildTwoPageCatalogActionsPdf() {
        string first = string.Join("\n", new[] {
            "BT",
            "/F1 12 Tf",
            "50 180 Td",
            "(First catalog page) Tj",
            "ET"
        });
        string second = string.Join("\n", new[] {
            "BT",
            "/F1 12 Tf",
            "50 180 Td",
            "(Second catalog-safe page) Tj",
            "ET"
        });
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Names << /JavaScript << /Names [(Startup) 7 0 R] >> >> >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 2 /Kids [3 0 R 4 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 320 220] /Resources << /Font << /F1 8 0 R >> >> /Contents 5 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 320 220] /Resources << /Font << /F1 8 0 R >> >> /Contents 6 0 R >>",
            "endobj",
            "5 0 obj",
            "<< /Length " + Encoding.ASCII.GetByteCount(first).ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>",
            "stream",
            first,
            "endstream",
            "endobj",
            "6 0 obj",
            "<< /Length " + Encoding.ASCII.GetByteCount(second).ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>",
            "stream",
            second,
            "endstream",
            "endobj",
            "7 0 obj",
            "<< /S /JavaScript /JS (app.alert('OfficeIMO')) >>",
            "endobj",
            "8 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 9 >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildPageAdditionalActionsPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R /AA << /O << /S /JavaScript /JS (app.alert('Page open')) >> /C << /S /Launch /F (tool.exe) >> >> >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 5 >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildWidgetAppearanceFormPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R /Annots [8 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Fields [7 0 R] >>",
            "endobj",
            "7 0 obj",
            "<< /FT /Btn /T (AcceptTerms) /V /Yes /Kids [8 0 R] >>",
            "endobj",
            "8 0 obj",
            "<< /Type /Annot /Subtype /Widget /Parent 7 0 R /Rect [20 100 36 116] /F 4 /AS /Yes /AP << /N << /Off 9 0 R /Yes 10 0 R >> >> >>",
            "endobj",
            "9 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "10 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 11 >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildAcroFormXfaPdf() {
        string template = "<template/>";
        string datasets = "<datasets/>";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Fields [] /XFA [(template) 6 0 R (datasets) 7 0 R] >>",
            "endobj",
            "6 0 obj",
            "<< /Length " + template.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>",
            "stream",
            template,
            "endstream",
            "endobj",
            "7 0 obj",
            "<< /Length " + datasets.Length.ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>",
            "stream",
            datasets,
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 8 >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildOpenActionHashPdf(double destinationTop) {
        string destinationTopText = destinationTop.ToString(System.Globalization.CultureInfo.InvariantCulture);
        string content = string.Join("\n", new[] {
            "BT",
            "/F1 12 Tf",
            "50 180 Td",
            "(Action hash marker) Tj",
            "ET"
        });

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /OpenAction [3 0 R /FitH " + destinationTopText + "] >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 320 220] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length " + Encoding.ASCII.GetByteCount(content).ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>",
            "stream",
            content,
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildOpenAndCatalogActionsPdf() {
        string content = string.Join("\n", new[] {
            "BT",
            "/F1 12 Tf",
            "50 180 Td",
            "(Open and catalog action marker) Tj",
            "ET"
        });

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /OpenAction [3 0 R /FitH 180] /Names << /JavaScript << /Names [(Startup) 5 0 R] >> >> >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 320 220] /Resources << /Font << /F1 6 0 R >> >> /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length " + Encoding.ASCII.GetByteCount(content).ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>",
            "stream",
            content,
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /S /JavaScript /JS (app.alert('OfficeIMO')) >>",
            "endobj",
            "6 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildAnnotationActionsPdf() {
        string content = string.Join("\n", new[] {
            "BT",
            "/F1 12 Tf",
            "50 180 Td",
            "(Annotation action marker) Tj",
            "ET"
        });

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 320 220] /Resources << /Font << /F1 8 0 R >> >> /Contents 4 0 R /Annots [5 0 R 7 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length " + Encoding.ASCII.GetByteCount(content).ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>",
            "stream",
            content,
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Annot /Subtype /Link /Rect [40 150 180 172] /A << /S /JavaScript /JS (app.alert('Annotation')) /Next 6 0 R >> >>",
            "endobj",
            "6 0 obj",
            "<< /S /Launch /F (chain.exe) >>",
            "endobj",
            "7 0 obj",
            "<< /Type /Annot /Subtype /Text /Rect [40 120 80 150] /Contents (Review note) /AA << /E << /S /Launch /F (tool.exe) >> >> >>",
            "endobj",
            "8 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 9 >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildRichMediaAnnotationActionPdf() {
        string content = string.Join("\n", new[] {
            "BT",
            "/F1 12 Tf",
            "50 180 Td",
            "(Rich media action marker) Tj",
            "ET"
        });

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 320 220] /Resources << /Font << /F1 6 0 R >> >> /Contents 4 0 R /Annots [5 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length " + Encoding.ASCII.GetByteCount(content).ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>",
            "stream",
            content,
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Annot /Subtype /Screen /Rect [40 120 180 180] /A << /S /RichMedia >> >>",
            "endobj",
            "6 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildFreeTextAppearanceMetadataPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 4 0 R /Annots [5 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Annot /Subtype /FreeText /Rect [10 20 190 100] /Contents (Reader styled note) /DS (font-size: 14pt; color: rgb(51, 102, 153); text-align: center) /RC (<body><p>Rich <b>reader</b> note</p></body>) /Border [0 0 1] /C [0.2 0.4 0.8] /IC [0.95 0.98 1] /CA 0.5 >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildAnnotationPathGeometryMetadataPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 300] /Contents 4 0 R /Annots [5 0 R 6 0 R 7 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Annot /Subtype /Highlight /Rect [20 80 160 110] /Contents (Reader highlight) /C [1 0.8 0.1] /QuadPoints [30 100 90 100 30 92 90 92] >>",
            "endobj",
            "6 0 obj",
            "<< /Type /Annot /Subtype /Line /Rect [20 80 160 120] /Contents (Reader line) /L [40 100 140 100] /C [0.1 0.2 0.7] /Border [0 0 2] /LE [/OpenArrow /ClosedArrow] >>",
            "endobj",
            "7 0 obj",
            "<< /Type /Annot /Subtype /Ink /Rect [20 20 180 60] /Contents (Reader ink) /InkList [[30 30 60 45 90 30] [100 30 130 45 160 30]] /C [0.1 0.2 0.7] /Border [0 0 2] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 8 >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildOptionalContentMetadataPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.5",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /OCProperties << /OCGs [5 0 R 6 0 R] /D << /Name (Default layers) /Creator (OfficeIMO fixture) /BaseState /ON /ON [5 0 R] /OFF [6 0 R] /Locked [6 0 R] /Order [(Layers) [5 0 R 6 0 R]] >> >> >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /OCG /Name (Print layer) /Intent [/View /Design] /Usage << /CreatorInfo << /Creator (OfficeIMO) /Subtype /Artwork >> /View << /ViewState /ON >> /Print << /PrintState /ON >> /Export << /ExportState /OFF >> >> >>",
            "endobj",
            "6 0 obj",
            "<< /Type /OCG /Name (Hidden layer) /Intent /View /Usage << /View << /ViewState /OFF >> /Print << /PrintState /OFF >> /Export << /ExportState /ON >> >> >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildSignedSecurityMetadataPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 7 0 R /Perms << /DocMDP 6 0 R /UR3 6 0 R >> /DSS 9 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R /Annots [5 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /FT /Sig /T (Approval) /V 6 0 R /Subtype /Widget /Rect [10 10 120 40] /Lock << /Type /SigFieldLock /Action /Include /Fields [(Total) (Approver)] >> /SV << /Filter /Adobe.PPKLite /SubFilter [/adbe.pkcs7.detached] /DigestMethod [/SHA256 /SHA512] /Reasons [(Approval) (Final)] /Ff 3 /AddRevInfo true /MDP << /P 2 >> >> >>",
            "endobj",
            "6 0 obj",
            "<< /Type /Sig /Filter /Adobe.PPKLite /SubFilter /adbe.pkcs7.detached /Name (Alice) /Location (Warsaw) /Reason (Approval) /ContactInfo (alice@example.test) /M (D:20260607120000+02'00') /ByteRange [0 10 20 30] /Contents <001122> /Reference [<< /TransformMethod /DocMDP /TransformParams << /Type /TransformParams /V /1.2 /P 2 >> >>] >>",
            "endobj",
            "7 0 obj",
            "<< /Fields [5 0 R] /SigFlags 3 >>",
            "endobj",
            "8 0 obj",
            "<< /Producer (OfficeIMO signed fixture) >>",
            "endobj",
            "9 0 obj",
            "<< /Certs [10 0 R 11 0 R] /OCSPs [12 0 R] /CRLs [13 0 R] /VRI << /ABCDEF << /Cert [10 0 R] /OCSP [12 0 R] /CRL [13 0 R] /TS 14 0 R >> >> >>",
            "endobj",
            "10 0 obj",
            "<< /Type /EmbeddedFile /Length 0 >>",
            "endobj",
            "11 0 obj",
            "<< /Type /EmbeddedFile /Length 0 >>",
            "endobj",
            "12 0 obj",
            "<< /Type /EmbeddedFile /Length 0 >>",
            "endobj",
            "13 0 obj",
            "<< /Type /EmbeddedFile /Length 0 >>",
            "endobj",
            "14 0 obj",
            "<< /Type /TimestampEvidence /Length 0 >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Info 8 0 R /ID [(abc) (def)] /Size 15 /Prev 100 >>",
            "startxref",
            "100",
            "%%EOF",
            "startxref",
            "200",
            "%%EOF"
        });

        return Encoding.ASCII.GetBytes(pdf);
    }
}
