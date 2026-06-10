using OfficeIMO.Pdf;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Pdf;
using System.Text;
using System.Text.Json;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderPdfModularTests {
    [Fact]
    public void DocumentReaderPdf_ReadPdfStream_EmitsPageAwareChunks() {
        byte[] pdf = BuildTwoPagePdf();
        using var stream = new MemoryStream(pdf, writable: false);

        var chunks = DocumentReaderPdfExtensions.ReadPdf(
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
        });
        Assert.Contains(chunks, c => c.Location.Page == 1 && (c.Markdown ?? c.Text).Contains("Reader PDF page one", StringComparison.Ordinal));
        Assert.Contains(chunks, c => c.Location.Page == 2 && (c.Markdown ?? c.Text).Contains("Reader PDF page two", StringComparison.Ordinal));
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfStream_CanSelectPageRanges() {
        byte[] pdf = BuildTwoPagePdf();
        using var stream = new MemoryStream(pdf, writable: false);

        var chunks = DocumentReaderPdfExtensions.ReadPdf(
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

        var chunks = DocumentReaderPdfExtensions.ReadPdf(
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
    public void DocumentReaderPdf_ReadPdfDocument_MapsLogicalPagesChunksAndBlocks() {
        byte[] pdf = BuildTwoPagePdf();
        using var stream = new MemoryStream(pdf, writable: false);

        OfficeDocumentReadResult result = DocumentReaderPdfExtensions.ReadPdfDocument(
            stream,
            sourceName: " readback.pdf ",
            readerOptions: new ReaderOptions { MaxChars = 8_000, ComputeHashes = true });

        Assert.Equal(OfficeDocumentReadResultSchema.Id, result.SchemaId);
        Assert.Equal(OfficeDocumentReadResultSchema.Version, result.SchemaVersion);
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
        Assert.Equal(1, root.GetProperty("schemaVersion").GetInt32());
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

        string json = DocumentReaderPdfExtensions.ReadPdfDocumentJson(
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
    public void DocumentReaderPdf_ReadPdfDocument_PreservesLogicalPageRangeOrder() {
        PdfLogicalDocument logical = PdfLogicalDocument.Load(BuildTwoPagePdf());

        OfficeDocumentReadResult result = DocumentReaderPdfExtensions.ReadPdfDocument(
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

        OfficeDocumentReadResult result = DocumentReaderPdfExtensions.ReadPdfDocument(
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

        OfficeDocumentReadResult result = DocumentReaderPdfExtensions.ReadPdfDocument(
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
    public void DocumentReaderPdf_ReadPdfDocument_PreservesRemoteLinkDestinations() {
        OfficeDocumentReadResult result = DocumentReaderPdfExtensions.ReadPdfDocument(
            new MemoryStream(BuildRemoteGoToLinkPdf(), writable: false),
            sourceName: "remote-link.pdf");

        OfficeDocumentLink link = Assert.Single(result.Links);
        Assert.Equal("remote", link.Kind);
        Assert.Equal("remote-report.pdf", link.RemoteFile);
        Assert.Null(link.RemoteDestinationName);
        Assert.Equal(2, link.RemoteDestinationPageNumber);
        Assert.Equal(nameof(PdfOpenActionDestinationMode.FitHorizontal), link.RemoteDestinationMode);
        Assert.Equal(144D, link.RemoteDestinationTop);

        using JsonDocument document = JsonDocument.Parse(result.ToJson());
        JsonElement jsonLink = document.RootElement.GetProperty("links")[0];
        Assert.Equal("remote-report.pdf", jsonLink.GetProperty("remoteFile").GetString());
        Assert.Equal(2, jsonLink.GetProperty("remoteDestinationPageNumber").GetInt32());
        Assert.Equal(nameof(PdfOpenActionDestinationMode.FitHorizontal), jsonLink.GetProperty("remoteDestinationMode").GetString());
        Assert.Equal(144D, jsonLink.GetProperty("remoteDestinationTop").GetDouble());
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfStream_UsesCurrentSeekableStreamPosition() {
        byte[] pdf = BuildTwoPagePdf();
        byte[] prefix = Encoding.ASCII.GetBytes("not-a-pdf-prefix");
        using var stream = new MemoryStream(prefix.Concat(pdf).ToArray(), writable: false);
        stream.Position = prefix.Length;

        var chunks = DocumentReaderPdfExtensions.ReadPdf(stream, sourceName: "embedded.pdf").ToList();

        Assert.NotEmpty(chunks);
        Assert.Contains(chunks, c => (c.Markdown ?? c.Text).Contains("Reader PDF page one", StringComparison.Ordinal));
        Assert.Contains(chunks, c => (c.Markdown ?? c.Text).Contains("Reader PDF page two", StringComparison.Ordinal));
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfStream_MaxInputBytesUsesCurrentSeekableStreamPosition() {
        byte[] pdf = BuildTwoPagePdf();
        byte[] prefix = Encoding.ASCII.GetBytes("not-a-pdf-prefix-that-is-longer-than-the-limit");
        using var stream = new MemoryStream(prefix.Concat(pdf).ToArray(), writable: false);
        stream.Position = prefix.Length;

        var chunks = DocumentReaderPdfExtensions.ReadPdf(
            stream,
            sourceName: "embedded-limited.pdf",
            readerOptions: new ReaderOptions { MaxInputBytes = pdf.Length }).ToList();

        Assert.NotEmpty(chunks);
        Assert.All(chunks, c => Assert.Equal(pdf.Length, c.SourceLengthBytes));
        Assert.Contains(chunks, c => (c.Markdown ?? c.Text).Contains("Reader PDF page one", StringComparison.Ordinal));
        Assert.Contains(chunks, c => (c.Markdown ?? c.Text).Contains("Reader PDF page two", StringComparison.Ordinal));
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfStream_DuplicatePageRangeSelectionsEmitUniqueChunkIds() {
        byte[] pdf = BuildTwoPagePdf();
        using var stream = new MemoryStream(pdf, writable: false);

        var chunks = DocumentReaderPdfExtensions.ReadPdf(
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

        var chunk = Assert.Single(DocumentReaderPdfExtensions.ReadPdf(
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

        OfficeDocumentReadResult result = DocumentReaderPdfExtensions.ReadPdfDocument(
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

        var chunk = Assert.Single(DocumentReaderPdfExtensions.ReadPdf(
            stream,
            sourceName: "table.pdf",
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

        OfficeDocumentReadResult result = DocumentReaderPdfExtensions.ReadPdfDocument(
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

        IReadOnlyList<ReaderTable> tables = DocumentReaderPdfExtensions.ReadPdfTables(
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
        Assert.Equal(2, table.TotalRowCount);
        Assert.Equal(new[] { "B-200", "Beta", "14" }, table.Rows[1]);

        using var exportStream = new MemoryStream(pdf, writable: false);
        ReaderTableExportBundle export = Assert.Single(DocumentReaderPdfExtensions.ReadPdfTableExports(
            exportStream,
            sourceName: "tables-only.pdf",
            pdfOptions: new ReaderPdfOptions {
                LayoutOptions = new PdfTextLayoutOptions {
                    ForceSingleColumn = true
                }
            }));
        Assert.Equal("tables-only-page-0001-table-0000", export.Id);
        Assert.Contains("A-100,Alpha,2", export.Csv, StringComparison.Ordinal);
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

        OfficeDocumentReadResult result = DocumentReaderPdfExtensions.ReadPdfDocument(
            stream,
            sourceName: "scan-candidate.pdf",
            readerOptions: new ReaderOptions { MaxChars = 8_000 });

        OfficeDocumentOcrCandidate candidate = Assert.Single(result.OcrCandidates);
        OfficeDocumentAsset asset = Assert.Single(result.Assets);
        Assert.Equal(OfficeDocumentAssetNaming.BuildFileName(asset.Id, asset.Extension), asset.FileName);
        Assert.NotNull(asset.PayloadBytes);
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
        Assert.Equal(OfficeDocumentReadResultSchema.Version, root.GetProperty("schemaVersion").GetInt32());
        Assert.Equal("ocr-needed", root.GetProperty("diagnostics")[0].GetProperty("code").GetString());
        Assert.Equal(asset.FileName, root.GetProperty("assets")[0].GetProperty("fileName").GetString());
        Assert.Equal("image", root.GetProperty("ocrCandidates")[0].GetProperty("kind").GetString());
        Assert.Equal("image", root.GetProperty("pages")[0].GetProperty("ocrCandidates")[0].GetProperty("kind").GetString());
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

        var chunks = DocumentReaderPdfExtensions.ReadPdf(
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

        IReadOnlyList<ReaderTableExportBundle> exports = DocumentReaderPdfExtensions.ReadPdfTableExports(
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

        var chunk = Assert.Single(DocumentReaderPdfExtensions.ReadPdf(
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

        var chunk = Assert.Single(DocumentReaderPdfExtensions.ReadPdf(
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

        var chunks = DocumentReaderPdfExtensions.ReadPdf(
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
    public void DocumentReaderPdf_Registration_DispatchesPdfStream() {
        try {
            DocumentReaderPdfRegistrationExtensions.RegisterPdfHandler();

            byte[] pdf = BuildTwoPagePdf();
            using var stream = new MemoryStream(pdf, writable: false);
            var chunks = DocumentReader.Read(stream, "registry.pdf").ToList();

            Assert.NotEmpty(chunks);
            Assert.Contains(chunks, c =>
                c.Kind == ReaderInputKind.Pdf &&
                string.Equals(c.Location.Path, "registry.pdf", StringComparison.OrdinalIgnoreCase) &&
                (c.Markdown ?? c.Text).Contains("Reader PDF page one", StringComparison.Ordinal));
        } finally {
            DocumentReaderPdfRegistrationExtensions.UnregisterPdfHandler();
        }
    }

    [Fact]
    public void DocumentReaderPdf_BootstrapFromAssembly_RegistersPdfHandlerAndManifest() {
        try {
            DocumentReaderPdfRegistrationExtensions.UnregisterPdfHandler();

            var result = DocumentReader.BootstrapHostFromAssemblies(
                new[] { typeof(DocumentReaderPdfRegistrationExtensions).Assembly },
                new ReaderHostBootstrapOptions {
                    ReplaceExistingHandlers = true,
                    IncludeBuiltInCapabilities = true,
                    IncludeCustomCapabilities = true,
                    IndentedManifestJson = false
                });

            Assert.NotNull(result);
            Assert.Contains(result.RegisteredHandlers, handler => handler.HandlerId == DocumentReaderPdfRegistrationExtensions.HandlerId);
            Assert.Contains(result.Manifest.Handlers, handler =>
                handler.Id == DocumentReaderPdfRegistrationExtensions.HandlerId &&
                handler.Kind == ReaderInputKind.Pdf &&
                handler.Extensions.Contains(".pdf"));
            Assert.Equal(1, result.Manifest.Handlers.Count(handler =>
                string.Equals(handler.Id, DocumentReaderPdfRegistrationExtensions.HandlerId, StringComparison.Ordinal)));
            Assert.DoesNotContain(result.Manifest.Handlers, handler =>
                handler.IsBuiltIn &&
                handler.Extensions.Contains(".pdf", StringComparer.OrdinalIgnoreCase));
            Assert.Contains(DocumentReaderPdfRegistrationExtensions.HandlerId, result.ManifestJson, StringComparison.OrdinalIgnoreCase);
        } finally {
            DocumentReaderPdfRegistrationExtensions.UnregisterPdfHandler();
        }
    }

    [Fact]
    public void DocumentReaderPdf_ReadPdfStream_NonSeekable_EnforcesMaxInputBytes() {
        byte[] pdf = BuildTwoPagePdf();
        using var stream = new NonSeekableReadStream(pdf);

        var ex = Assert.Throws<IOException>(() => DocumentReaderPdfExtensions.ReadPdf(
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
    public void ReaderPdf_ProfileContract_DescribesRegisteredHandlerAndChunkShape() {
        ReaderPdfProfileContract contract = ReaderPdfProfileContracts.OfficeIMO;

        Assert.Equal(DocumentReaderPdfRegistrationExtensions.HandlerId, contract.Id);
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

    private static byte[] CreateMinimalRgbPng() {
        return new byte[] {
            137, 80, 78, 71, 13, 10, 26, 10,
            0, 0, 0, 13,
            73, 72, 68, 82,
            0, 0, 0, 1,
            0, 0, 0, 1,
            8, 2, 0, 0, 0,
            0, 0, 0, 0,
            0, 0, 0, 12,
            73, 68, 65, 84,
            0x78, 0x9C, 0x63, 0xF8, 0xCF, 0xC0, 0x00, 0x00, 0x03, 0x01, 0x01, 0x00,
            0, 0, 0, 0,
            0, 0, 0, 0,
            73, 69, 78, 68,
            0, 0, 0, 0
        };
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
}
