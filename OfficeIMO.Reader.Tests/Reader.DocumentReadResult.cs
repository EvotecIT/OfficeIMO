using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Reader;
using OfficeIMO.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.Json;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderDocumentReadResultTests {
    [Fact]
    public void DocumentReader_ReadDocument_EmitsSharedEnvelopeForMarkdown() {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".md");
        try {
            File.WriteAllText(path, "# Top\n\nPara 1.\n\n## Child\n\nPara 2.\n");

            OfficeDocumentReadResult result = OfficeIMO.Reader.Tests.ReaderTestReaders.All.ReadDocument(path);

            Assert.Equal(OfficeDocumentReadResultSchema.Id, result.SchemaId);
            Assert.Equal(OfficeDocumentReadResultSchema.CurrentVersion, result.SchemaVersion);
            Assert.Equal(ReaderInputKind.Markdown, result.Kind);
            Assert.Equal(path, result.Source.Path);
            Assert.False(string.IsNullOrWhiteSpace(result.Source.SourceId));
            Assert.Contains("officeimo.reader", result.CapabilitiesUsed);
            Assert.Contains("officeimo.reader.markdown", result.CapabilitiesUsed);
            Assert.Contains("# Top", result.Markdown, StringComparison.Ordinal);
            Assert.Contains("## Child", result.Markdown, StringComparison.Ordinal);
            Assert.Equal(2, result.Chunks.Count);
            Assert.Equal(2, result.Blocks.Count);
            Assert.Contains(result.Metadata, entry =>
                entry.Category == "reader.summary" &&
                entry.Name == "ChunkCount" &&
                entry.Value == "2" &&
                entry.ValueType == "count");
            Assert.Contains(result.Metadata, entry =>
                entry.Category == "reader.summary" &&
                entry.Name == "BlockCount" &&
                entry.Value == "2" &&
                entry.ValueType == "count");
            Assert.Empty(result.Pages);
            Assert.Empty(result.Diagnostics);
            Assert.Contains(result.Blocks, block => block.Location.HeadingPath == "Top");
            Assert.Contains(result.Blocks, block => block.Location.HeadingPath == "Top > Child");
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReader_ReadDocument_EmitsMarkdownTableSourceLocations() {
        byte[] bytes = Encoding.UTF8.GetBytes("# Inventory\n\n| Name | Qty |\n| --- | ---: |\n| Alpha | 2 |\n");
        using var stream = new MemoryStream(bytes, writable: false);

        OfficeDocumentReadResult result = OfficeIMO.Reader.Tests.ReaderTestReaders.All.ReadDocument(stream, "inventory.md");

        ReaderTable table = Assert.Single(result.Tables);
        Assert.Equal("inventory.md", table.Location?.Path);
        Assert.Equal(0, table.Location?.BlockIndex);
        Assert.Equal(1, table.Location?.SourceBlockIndex);
        Assert.Equal(3, table.Location?.StartLine);
        Assert.Equal(5, table.Location?.EndLine);
        Assert.Equal(3, table.Location?.NormalizedStartLine);
        Assert.Equal(5, table.Location?.NormalizedEndLine);
        Assert.Equal("Inventory", table.Location?.HeadingPath);
        Assert.Equal("inventory", table.Location?.HeadingSlug);
        Assert.Equal("table", table.Location?.SourceBlockKind);
        Assert.Equal("inventory--table-1", table.Location?.BlockAnchor);
        Assert.Equal(0, table.Location?.TableIndex);
        Assert.Contains(result.Metadata, entry =>
            entry.Category == "reader.summary" &&
            entry.Name == "TableCount" &&
            entry.Value == "1" &&
            entry.ValueType == "count");

        using JsonDocument jsonDocument = JsonDocument.Parse(result.ToJson());
        JsonElement location = jsonDocument.RootElement.GetProperty("tables")[0].GetProperty("location");
        Assert.Equal("inventory.md", location.GetProperty("path").GetString());
        Assert.Equal("table", location.GetProperty("sourceBlockKind").GetString());
        Assert.Equal(0, location.GetProperty("tableIndex").GetInt32());
    }

    [Fact]
    public void DocumentReader_ReadDocument_EmitsMarkdownVisualSourceLocations() {
        byte[] bytes = Encoding.UTF8.GetBytes("# Diagram\n\n```mermaid\ngraph TD\nA-->B\n```\n");
        using var stream = new MemoryStream(bytes, writable: false);

        OfficeDocumentReadResult result = OfficeIMO.Reader.Tests.ReaderTestReaders.All.ReadDocument(stream, "diagram.md");

        ReaderVisual visual = Assert.Single(result.Visuals);
        Assert.Equal("mermaid", visual.Kind);
        Assert.Equal("mermaid", visual.Language);
        Assert.Equal("diagram.md", visual.Location?.Path);
        Assert.Equal(0, visual.Location?.BlockIndex);
        Assert.Equal(1, visual.Location?.SourceBlockIndex);
        Assert.Equal(3, visual.Location?.StartLine);
        Assert.Equal(6, visual.Location?.EndLine);
        Assert.Equal(3, visual.Location?.NormalizedStartLine);
        Assert.Equal(6, visual.Location?.NormalizedEndLine);
        Assert.Equal("Diagram", visual.Location?.HeadingPath);
        Assert.Equal("diagram", visual.Location?.HeadingSlug);
        Assert.Equal("code", visual.Location?.SourceBlockKind);
        Assert.Equal("diagram--code-1", visual.Location?.BlockAnchor);
        Assert.Contains(result.Metadata, entry =>
            entry.Category == "reader.summary" &&
            entry.Name == "VisualCount" &&
            entry.Value == "1" &&
            entry.ValueType == "count");

        using JsonDocument jsonDocument = JsonDocument.Parse(result.ToJson());
        JsonElement location = jsonDocument.RootElement.GetProperty("visuals")[0].GetProperty("location");
        Assert.Equal("diagram.md", location.GetProperty("path").GetString());
        Assert.Equal("code", location.GetProperty("sourceBlockKind").GetString());
        Assert.Equal("diagram--code-1", location.GetProperty("blockAnchor").GetString());
    }

    [Fact]
    public void DocumentReader_ReadVisuals_ReturnsMarkdownVisualsFromStream() {
        byte[] bytes = Encoding.UTF8.GetBytes("# Diagram\n\n```mermaid\ngraph TD\nA-->B\n```\n");
        using var stream = new MemoryStream(bytes, writable: false);

        IReadOnlyList<ReaderVisual> visuals = OfficeIMO.Reader.Tests.ReaderTestReaders.All.ReadVisuals(stream, "diagram.md");

        ReaderVisual visual = Assert.Single(visuals);
        Assert.Equal("mermaid", visual.Kind);
        Assert.Equal("mermaid", visual.Language);
        Assert.Equal("graph TD\nA-->B", visual.Content);
        Assert.Equal("diagram.md", visual.Location?.Path);
        Assert.Equal(0, visual.Location?.BlockIndex);
        Assert.Equal(1, visual.Location?.SourceBlockIndex);
        Assert.Equal(3, visual.Location?.StartLine);
        Assert.Equal(6, visual.Location?.EndLine);
        Assert.Equal("Diagram", visual.Location?.HeadingPath);
        Assert.Equal("code", visual.Location?.SourceBlockKind);
        Assert.Equal("diagram--code-1", visual.Location?.BlockAnchor);
    }

    [Fact]
    public void DocumentReader_ExtractVisuals_AddsChunkLocationFallback() {
        var chunks = new[] {
            new ReaderChunk {
                Kind = ReaderInputKind.Markdown,
                Location = new ReaderLocation {
                    Path = "diagram.md",
                    BlockIndex = 4,
                    SourceBlockIndex = 7,
                    HeadingPath = "Architecture > Flow",
                    HeadingSlug = "architecture-flow",
                    SourceBlockKind = "code",
                    BlockAnchor = "architecture-flow--code-7"
                },
                Visuals = new[] {
                    new ReaderVisual {
                        Kind = "chart",
                        Language = "ix-chart",
                        Content = "{\"type\":\"bar\"}",
                        PayloadHash = "payload-1"
                    }
                }
            }
        };

        ReaderVisual visual = Assert.Single(OfficeIMO.Reader.Tests.ReaderTestReaders.All.ExtractVisuals(chunks));

        Assert.Equal("chart", visual.Kind);
        Assert.Equal("ix-chart", visual.Language);
        Assert.Equal("{\"type\":\"bar\"}", visual.Content);
        Assert.Equal("payload-1", visual.PayloadHash);
        Assert.Equal("diagram.md", visual.Location?.Path);
        Assert.Equal(4, visual.Location?.BlockIndex);
        Assert.Equal(7, visual.Location?.SourceBlockIndex);
        Assert.Equal("Architecture > Flow", visual.Location?.HeadingPath);
        Assert.Equal("architecture-flow", visual.Location?.HeadingSlug);
        Assert.Equal("code", visual.Location?.SourceBlockKind);
        Assert.Equal("architecture-flow--code-7", visual.Location?.BlockAnchor);
    }

    [Fact]
    public void DocumentReader_ReadDocument_EmitsWorkbookMetadataAndTableLocationsForExcel() {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
        try {
            using (var workbook = ExcelDocument.Create(path)) {
                var sheet = workbook.AddWorksheet("Data");
                sheet.Cell(1, 1, "Name");
                sheet.Cell(1, 2, "Value");
                sheet.Cell(2, 1, "Alpha");
                sheet.Cell(2, 2, 1);
                sheet.Cell(3, 1, "Beta");
                sheet.Cell(3, 2, 2);
                workbook.Save();
            }

            OfficeDocumentReadResult result = OfficeIMO.Reader.Tests.ReaderTestReaders.Excel("Data", "A1:B3").ReadDocument(path);

            Assert.Equal(ReaderInputKind.Excel, result.Kind);
            Assert.Contains("officeimo.reader.excel", result.CapabilitiesUsed);
            Assert.Single(result.Chunks);
            Assert.Single(result.Blocks);

            OfficeDocumentPage page = Assert.Single(result.Pages);
            Assert.Equal("Data", page.Name);
            Assert.Equal("sheet", page.Location.SourceBlockKind);
            Assert.Single(page.Tables);

            ReaderTable table = Assert.Single(result.Tables);
            Assert.Equal("Data", table.Location?.Sheet);
            Assert.Equal("A1:B3", table.Location?.A1Range);
            Assert.Equal("table", table.Location?.SourceBlockKind);
            Assert.Equal(0, table.Location?.SourceBlockIndex);
            Assert.Equal(0, table.Location?.TableIndex);
            Assert.Equal(new[] { "Name", "Value" }, table.Columns);
            Assert.Equal(2, table.Rows.Count);
            Assert.Equal("Beta", table.Rows[1][0]);
            Assert.Equal("2", table.Rows[1][1]);

            Assert.Contains(result.Metadata, entry =>
                entry.Category == "reader.summary" &&
                entry.Name == "ChunkCount" &&
                entry.Value == "1" &&
                entry.ValueType == "count");
            Assert.Contains(result.Metadata, entry =>
                entry.Category == "reader.summary" &&
                entry.Name == "TableCount" &&
                entry.Value == "1" &&
                entry.ValueType == "count");
            Assert.Contains(result.Metadata, entry =>
                entry.Category == "reader.container" &&
                entry.Name == "SheetCount" &&
                entry.Value == "1" &&
                entry.ValueType == "count");

            using JsonDocument jsonDocument = JsonDocument.Parse(result.ToJson());
            JsonElement root = jsonDocument.RootElement;
            Assert.Equal("Excel", root.GetProperty("kind").GetString());
            Assert.Equal("Data", root.GetProperty("pages")[0].GetProperty("name").GetString());
            Assert.Equal("Data", root.GetProperty("tables")[0].GetProperty("location").GetProperty("sheet").GetString());
            Assert.Contains(root.GetProperty("metadata").EnumerateArray(), entry =>
                entry.GetProperty("category").GetString() == "reader.container" &&
                entry.GetProperty("name").GetString() == "SheetCount" &&
                entry.GetProperty("value").GetString() == "1");
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReader_ReadTableExports_UsesDocumentWideExcelTableIndexesAcrossChunks() {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
        try {
            using (var workbook = ExcelDocument.Create(path)) {
                var sheet = workbook.AddWorksheet("Data");
                sheet.Cell(1, 1, "Name");
                sheet.Cell(1, 2, "Value");
                sheet.Cell(2, 1, "Alpha");
                sheet.Cell(2, 2, 1);
                sheet.Cell(3, 1, "Beta");
                sheet.Cell(3, 2, 2);
                sheet.Cell(4, 1, "Gamma");
                sheet.Cell(4, 2, 3);
                workbook.Save();
            }

            IReadOnlyList<ReaderTableExportBundle> exports = OfficeIMO.Reader.Tests.ReaderTestReaders.Excel("Data", "A1:B4", chunkRows: 1).ReadTableExports(path);

            Assert.True(exports.Count > 1);
            int[] indexes = Enumerable.Range(0, exports.Count).ToArray();
            Assert.Equal(indexes, exports.Select(export => export.Table.Location?.TableIndex ?? -1).ToArray());
            Assert.Equal(exports.Count, exports.Select(export => export.Id).Distinct(StringComparer.Ordinal).Count());
            Assert.All(indexes, index => Assert.EndsWith("-sheet-Data-table-" + index.ToString("D4"), exports[index].Id));
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReader_ReadDocument_PreservesWorkbookSheetOrderForPages() {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
        try {
            using (var workbook = ExcelDocument.Create(path)) {
                var zSheet = workbook.AddWorksheet("Z");
                zSheet.Cell(1, 1, "Name");
                zSheet.Cell(2, 1, "Zulu");
                var aSheet = workbook.AddWorksheet("A");
                aSheet.Cell(1, 1, "Name");
                aSheet.Cell(2, 1, "Alpha");
                workbook.Save();
            }

            OfficeDocumentReadResult result = OfficeIMO.Reader.Tests.ReaderTestReaders.All.ReadDocument(path);

            Assert.Equal(new[] { "Z", "A" }, result.Pages.Select(page => page.Name).ToArray());
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReader_ReadDocument_NonSeekableStream_EnforcesMaxInputBytesBeforeSnapshot() {
        using var package = new MemoryStream();
        using (WordDocument document = WordDocument.Create(package)) {
            document.AddParagraph("Large document");
            document.Save();
        }

        using var stream = new NonSeekableReadStream(package.ToArray());

        IOException ex = Assert.Throws<IOException>(() => OfficeIMO.Reader.Tests.ReaderTestReaders.All.ReadDocument(
            stream,
            "large.docx",
            new ReaderOptions { MaxInputBytes = 16 }));
        Assert.Contains("Input exceeds MaxInputBytes", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReader_Read_NonSeekableStream_EnforcesMaxInputBytesBeforeSnapshot() {
        using var package = new MemoryStream();
        using (WordDocument document = WordDocument.Create(package)) {
            document.AddParagraph("Large document");
            document.Save();
        }

        using var stream = new NonSeekableReadStream(package.ToArray());

        IOException ex = Assert.Throws<IOException>(() => {
            foreach (ReaderChunk _ in OfficeIMO.Reader.Tests.ReaderTestReaders.All.Read(
                stream,
                "large.docx",
                new ReaderOptions { MaxInputBytes = 16 })) {
            }
        });
        Assert.Contains("Input exceeds MaxInputBytes", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void DocumentReader_ReadDocument_EmitsChunkMetadataForWordDocument() {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        try {
            using (var document = WordDocument.Create(path)) {
                document.AddParagraph("Policy").Style = WordParagraphStyles.Heading1;
                document.AddParagraph("This is the body.");
                document.AddParagraph("Details").Style = WordParagraphStyles.Heading2;
                document.AddParagraph("Nested body.");
                document.Save();
            }

            OfficeDocumentReadResult result = OfficeIMO.Reader.Tests.ReaderTestReaders.All.ReadDocument(path);

            Assert.Equal(ReaderInputKind.Word, result.Kind);
            Assert.Contains("officeimo.reader.word", result.CapabilitiesUsed);
            Assert.NotEmpty(result.Chunks);
            Assert.NotEmpty(result.Blocks);
            Assert.Empty(result.Pages);
            Assert.Contains("# Policy", result.Markdown, StringComparison.Ordinal);
            Assert.Contains("## Details", result.Markdown, StringComparison.Ordinal);
            Assert.Contains(result.Blocks, block => block.Location.HeadingPath == "Policy");
            Assert.Contains(result.Blocks, block => block.Text.Contains("Nested body.", StringComparison.Ordinal));
            Assert.Contains(result.Metadata, entry =>
                entry.Category == "reader.summary" &&
                entry.Name == "ChunkCount" &&
                entry.Value == result.Chunks.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) &&
                entry.ValueType == "count");
            Assert.Contains(result.Metadata, entry =>
                entry.Category == "reader.summary" &&
                entry.Name == "BlockCount" &&
                entry.Value == result.Blocks.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) &&
                entry.ValueType == "count");
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReader_ReadDocument_EmitsSlideContainersForPowerPointPresentation() {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
        try {
            using (var presentation = PowerPointPresentation.Create(path)) {
                presentation.AddSlide().AddTextBox("Intro slide");
                presentation.AddSlide().AddTextBox("Details slide");
                presentation.Save();
            }

            OfficeDocumentReadResult result = OfficeIMO.Reader.Tests.ReaderTestReaders.All.ReadDocument(path);

            Assert.Equal(ReaderInputKind.PowerPoint, result.Kind);
            Assert.Contains("officeimo.reader.powerpoint", result.CapabilitiesUsed);
            Assert.Equal(2, result.Pages.Count);
            Assert.Equal(new[] { 1, 2 }, result.Pages.Select(page => page.Number.GetValueOrDefault()).ToArray());
            Assert.All(result.Pages, page => Assert.Equal("slide", page.Location.SourceBlockKind));
            Assert.Contains(result.Blocks, block => block.Location.Slide == 1 && block.Text.Contains("Intro slide", StringComparison.Ordinal));
            Assert.Contains(result.Blocks, block => block.Location.Slide == 2 && block.Text.Contains("Details slide", StringComparison.Ordinal));
            Assert.Contains(result.Metadata, entry =>
                entry.Category == "reader.container" &&
                entry.Name == "SlideCount" &&
                entry.Value == "2" &&
                entry.ValueType == "count");
            using JsonDocument jsonDocument = JsonDocument.Parse(result.ToJson());
            JsonElement root = jsonDocument.RootElement;
            Assert.Equal("PowerPoint", root.GetProperty("kind").GetString());
            Assert.Equal(2, root.GetProperty("pages").GetArrayLength());
            Assert.Equal("slide", root.GetProperty("pages")[0].GetProperty("location").GetProperty("sourceBlockKind").GetString());
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void DocumentReader_ReadDocumentJson_EmitsStableJsonForMarkdownStream() {
        byte[] bytes = Encoding.UTF8.GetBytes("# Top\n\nPara 1.\n\n## Child\n\nPara 2.\n");
        using var stream = new MemoryStream(bytes, writable: false);

        string json = OfficeIMO.Reader.Tests.ReaderTestReaders.All.ReadDocumentJson(stream, " notes.md ");

        using JsonDocument document = JsonDocument.Parse(json);
        JsonElement root = document.RootElement;
        Assert.Equal(OfficeDocumentReadResultSchema.Id, root.GetProperty("schemaId").GetString());
        Assert.Equal(OfficeDocumentReadResultSchema.CurrentVersion, root.GetProperty("schemaVersion").GetInt32());
        Assert.Equal("Markdown", root.GetProperty("kind").GetString());
        Assert.Equal("notes.md", root.GetProperty("source").GetProperty("path").GetString());
        Assert.Equal("officeimo.reader.markdown", root.GetProperty("capabilitiesUsed")[1].GetString());
        Assert.Equal(2, root.GetProperty("chunks").GetArrayLength());
        Assert.Contains(root.GetProperty("metadata").EnumerateArray(), entry =>
            entry.GetProperty("category").GetString() == "reader.summary" &&
            entry.GetProperty("name").GetString() == "ChunkCount" &&
            entry.GetProperty("value").GetString() == "2" &&
            entry.GetProperty("valueType").GetString() == "count");
        Assert.Equal(2, root.GetProperty("blocks").GetArrayLength());
        Assert.Contains("# Top", root.GetProperty("markdown").GetString(), StringComparison.Ordinal);
        Assert.False(root.TryGetProperty("html", out _));
    }

    [Fact]
    public void OfficeDocumentReadResult_ToJson_KeepsStableTopLevelSchemaOrder() {
        var result = new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Markdown,
            Source = new OfficeDocumentSource { Path = "sample.md" },
            CapabilitiesUsed = new[] { "officeimo.reader.markdown" },
            Markdown = "# Sample",
            Html = "<h1>Sample</h1>",
            Json = "{\"sample\":true}"
        };

        string json = result.ToJson();

        using JsonDocument document = JsonDocument.Parse(json);
        string[] names = document.RootElement.EnumerateObject().Select(property => property.Name).ToArray();
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
        }, names);
    }

    [Fact]
    public void DocumentReader_ReadTables_ReturnsMarkdownTablesFromStream() {
        byte[] bytes = Encoding.UTF8.GetBytes("| Name | Qty |\n| --- | ---: |\n| Alpha | 2 |\n| Beta | 14 |\n");
        using var stream = new MemoryStream(bytes, writable: false);

        IReadOnlyList<ReaderTable> tables = OfficeIMO.Reader.Tests.ReaderTestReaders.All.ReadTables(
            stream,
            "inventory.md",
            new ReaderOptions { MaxTableRows = 1 });

        ReaderTable table = Assert.Single(tables);
        Assert.Equal("inventory.md", table.Location?.Path);
        Assert.Equal("table", table.Location?.SourceBlockKind);
        Assert.Equal(0, table.Location?.SourceBlockIndex);
        Assert.Equal(0, table.Location?.TableIndex);
        Assert.Equal(1, table.Location?.StartLine);
        Assert.Equal(new[] { "Name", "Qty" }, table.Columns);
        Assert.Equal(2, table.TotalRowCount);
        Assert.True(table.Truncated);
        Assert.Single(table.Rows);
        Assert.Equal(new[] { "Alpha", "2" }, table.Rows[0]);
        Assert.Equal(ReaderTableColumnKind.Numeric, table.ColumnProfiles[1].Kind);
    }

    [Fact]
    public void DocumentReader_ReadTableExports_ReturnsDeterministicSidecarPayloads() {
        byte[] bytes = Encoding.UTF8.GetBytes("| Name | Qty |\n| --- | ---: |\n| Alpha | 2 |\n");
        using var stream = new MemoryStream(bytes, writable: false);

        IReadOnlyList<ReaderTableExportBundle> exports = OfficeIMO.Reader.Tests.ReaderTestReaders.All.ReadTableExports(
            stream,
            "inventory.md",
            indentedJson: true);

        ReaderTableExportBundle export = Assert.Single(exports);
        Assert.Equal("inventory-table-0000", export.Id);
        Assert.Equal("inventory-table-0000", export.FileNamePrefix);
        Assert.Equal("inventory.md", export.Table.Location?.Path);
        Assert.Equal("Name,Qty\r\nAlpha,2", export.Csv);
        Assert.Contains("| Name | Qty |", export.Markdown, StringComparison.Ordinal);
        using JsonDocument document = JsonDocument.Parse(export.Json);
        JsonElement root = document.RootElement;
        Assert.Equal("inventory.md", root.GetProperty("location").GetProperty("path").GetString());
        Assert.Equal("Alpha", root.GetProperty("rows")[0][0].GetString());
    }

    [Fact]
    public void ReaderTableExportMaterializer_WriteTableExportsToDirectory_WritesCsvMarkdownAndJsonSidecars() {
        byte[] bytes = Encoding.UTF8.GetBytes("| Name | Qty |\n| --- | ---: |\n| Alpha | 2 |\n");
        using var stream = new MemoryStream(bytes, writable: false);
        IReadOnlyList<ReaderTableExportBundle> exports = OfficeIMO.Reader.Tests.ReaderTestReaders.All.ReadTableExports(stream, "inventory.md");
        var directory = Path.Combine(Path.GetTempPath(), "officeimo-reader-tables-" + Guid.NewGuid().ToString("N"));

        try {
            IReadOnlyList<ReaderTableMaterializedExport> materialized = exports.WriteTableExportsToDirectory(directory);

            Assert.Equal(3, materialized.Count);
            Assert.All(materialized, item => Assert.True(item.Written));
            Assert.True(File.Exists(Path.Combine(directory, "inventory-table-0000.csv")));
            Assert.True(File.Exists(Path.Combine(directory, "inventory-table-0000.md")));
            Assert.True(File.Exists(Path.Combine(directory, "inventory-table-0000.json")));
            Assert.Equal("Name,Qty\r\nAlpha,2", File.ReadAllText(Path.Combine(directory, "inventory-table-0000.csv")));
            Assert.Contains("| Name | Qty |", File.ReadAllText(Path.Combine(directory, "inventory-table-0000.md")), StringComparison.Ordinal);
            using JsonDocument document = JsonDocument.Parse(File.ReadAllText(Path.Combine(directory, "inventory-table-0000.json")));
            Assert.Equal("inventory.md", document.RootElement.GetProperty("location").GetProperty("path").GetString());

            IReadOnlyList<ReaderTableMaterializedExport> skipped = exports.WriteTableExportsToDirectory(
                directory,
                new ReaderTableExportMaterializationOptions { Overwrite = false });
            Assert.Equal(3, skipped.Count);
            Assert.All(skipped, item => {
                Assert.False(item.Written);
                Assert.Equal("Destination file already exists.", item.SkippedReason);
            });
        } finally {
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void ReaderTableExportMaterializer_StreamTableExports_StreamsSelectedPayloadsThroughCallback() {
        var export = new ReaderTableExportBundle {
            Id = "manual-table",
            FileNamePrefix = "manual-table",
            Csv = "Name,Qty",
            Markdown = "| Name | Qty |",
            Json = "{\"columns\":[\"Name\",\"Qty\"]}"
        };
        string captured = string.Empty;

        IReadOnlyList<ReaderTableMaterializedExport> materialized = new[] { export }.StreamTableExports(
            (bundle, format, payload) => {
                using var reader = new StreamReader(payload, Encoding.UTF8);
                captured = reader.ReadToEnd();
            },
            new ReaderTableExportMaterializationOptions {
                IncludeCsv = false,
                IncludeMarkdown = false,
                IncludeJson = true
            });

        ReaderTableMaterializedExport written = Assert.Single(materialized);
        Assert.True(written.Written);
        Assert.Equal(ReaderTableExportFormat.Json, written.Format);
        Assert.Equal("manual-table.json", written.FileName);
        Assert.Equal("{\"columns\":[\"Name\",\"Qty\"]}", captured);
    }

    [Fact]
    public void OfficeDocumentAssetNaming_BuildFileName_SanitizesIdsAndExtensions() {
        string fileName = OfficeDocumentAssetNaming.BuildFileName("Page 1/Image:Main", ".PNG");

        Assert.Equal("page-1-image-main.png", fileName);
    }

    [Fact]
    public void OfficeDocumentAssetMaterializer_WriteAssetsToDirectory_WritesPayloadsAndSkipsManifestOnlyAssets() {
        byte[] payload = Encoding.ASCII.GetBytes("asset payload");
        var directory = Path.Combine(Path.GetTempPath(), "officeimo-reader-assets-" + Guid.NewGuid().ToString("N"));
        var result = new OfficeDocumentReadResult {
            Assets = new[] {
                new OfficeDocumentAsset {
                    Id = "asset:one",
                    Kind = "image",
                    Extension = ".bin",
                    FileName = "nested/asset-one.bin",
                    PayloadBytes = payload
                },
                new OfficeDocumentAsset {
                    Id = "asset:two",
                    Kind = "preview",
                    Extension = ".svg",
                    FileName = "asset-two.svg"
                }
            }
        };

        try {
            IReadOnlyList<OfficeDocumentMaterializedAsset> materialized = result.WriteAssetsToDirectory(directory);

            Assert.Equal(2, materialized.Count);
            OfficeDocumentMaterializedAsset written = Assert.Single(materialized, item => item.Written);
            Assert.Equal("asset-one.bin", written.FileName);
            Assert.True(File.Exists(written.Path));
            Assert.Equal(payload, File.ReadAllBytes(written.Path!));
            OfficeDocumentMaterializedAsset skipped = Assert.Single(materialized, item => !item.Written);
            Assert.Equal("Asset has no in-memory payload.", skipped.SkippedReason);

            string json = result.ToJson();
            using JsonDocument document = JsonDocument.Parse(json);
            Assert.False(document.RootElement.GetProperty("assets")[0].TryGetProperty("payloadBytes", out _));
        } finally {
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void OfficeDocumentAssetMaterializer_StreamAssets_StreamsPayloadsThroughCallback() {
        byte[] payload = Encoding.ASCII.GetBytes("streamed payload");
        byte[] captured = Array.Empty<byte>();
        var result = new OfficeDocumentReadResult {
            Assets = new[] {
                new OfficeDocumentAsset {
                    Id = "asset-stream",
                    Kind = "image",
                    Extension = ".bin",
                    PayloadBytes = payload
                }
            }
        };

        IReadOnlyList<OfficeDocumentMaterializedAsset> materialized = result.StreamAssets((asset, stream) => {
            using var memory = new MemoryStream();
            stream.CopyTo(memory);
            captured = memory.ToArray();
        });

        OfficeDocumentMaterializedAsset written = Assert.Single(materialized);
        Assert.True(written.Written);
        Assert.Equal("asset-stream.bin", written.FileName);
        Assert.Equal(payload, captured);
    }

    [Fact]
    public void OfficeDocumentAssetMaterializer_WhenHashValidationEnabled_SkipsMismatchedPayloads() {
        byte[] payload = Encoding.ASCII.GetBytes("payload");
        var directory = Path.Combine(Path.GetTempPath(), "officeimo-reader-assets-" + Guid.NewGuid().ToString("N"));
        var result = new OfficeDocumentReadResult {
            Assets = new[] {
                new OfficeDocumentAsset {
                    Id = "bad-hash",
                    Kind = "image",
                    Extension = ".bin",
                    PayloadBytes = payload,
                    PayloadHash = new string('0', 64)
                }
            }
        };

        try {
            IReadOnlyList<OfficeDocumentMaterializedAsset> materialized = result.WriteAssetsToDirectory(
                directory,
                new OfficeDocumentAssetMaterializationOptions { ValidatePayloadHash = true });

            OfficeDocumentMaterializedAsset skipped = Assert.Single(materialized);
            Assert.False(skipped.Written);
            Assert.Equal("Asset payload hash does not match PayloadHash.", skipped.SkippedReason);
            Assert.False(File.Exists(skipped.Path));
        } finally {
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void OfficeDocumentAssetHash_PayloadHashMatches_ComputesSha256PayloadHash() {
        byte[] payload = Encoding.ASCII.GetBytes("payload");
        string hash = OfficeDocumentAssetHash.ComputeSha256Hex(payload);
        var asset = new OfficeDocumentAsset {
            Id = "hash",
            Kind = "image",
            PayloadBytes = payload,
            PayloadHash = hash.ToUpperInvariant()
        };

        Assert.True(asset.PayloadHashMatches(out string? actualHash));
        Assert.Equal(hash, actualHash);
    }

    [Fact]
    public void OfficeDocumentAssetDataUri_TryBuildDataUri_EmbedsSmallPayloadsOnlyWhenRequested() {
        var asset = new OfficeDocumentAsset {
            Id = "asset-inline",
            Kind = "preview",
            MediaType = "image/svg+xml",
            Extension = ".svg",
            PayloadBytes = Encoding.UTF8.GetBytes("<svg/>")
        };

        Assert.True(asset.TryBuildDataUri(out string? dataUri));
        Assert.Equal("data:image/svg+xml;base64,PHN2Zy8+", dataUri);

        Assert.False(asset.TryBuildDataUri(out string? cappedDataUri, maxInlineBytes: 2));
        Assert.Null(cappedDataUri);
    }

    [Fact]
    public void OfficeDocumentAssetDataUri_BuildAssetDataUriMap_FiltersManifestOnlyAndOversizedAssets() {
        var result = new OfficeDocumentReadResult {
            Assets = new[] {
                new OfficeDocumentAsset {
                    Id = "small",
                    Kind = "image",
                    MediaType = "text/plain",
                    PayloadBytes = Encoding.ASCII.GetBytes("ok")
                },
                new OfficeDocumentAsset {
                    Id = "large",
                    Kind = "image",
                    MediaType = "text/plain",
                    PayloadBytes = Encoding.ASCII.GetBytes("too-large")
                },
                new OfficeDocumentAsset {
                    Id = "manifest-only",
                    Kind = "preview",
                    MediaType = "image/svg+xml"
                }
            }
        };

        IReadOnlyDictionary<string, string> dataUris = result.BuildAssetDataUriMap(new OfficeDocumentAssetDataUriOptions {
            MaxInlineBytes = 2,
            Predicate = asset => asset.Kind == "image"
        });

        string dataUri = Assert.Single(dataUris).Value;
        Assert.Equal("data:text/plain;base64,b2s=", dataUri);
        Assert.True(dataUris.ContainsKey("small"));
    }

    [Fact]
    public void ReaderTableExport_ToCsv_QuotesSpecialCellsAndNormalizesRowWidth() {
        var table = new ReaderTable {
            Columns = new[] { "Name", "Qty" },
            Rows = new[] {
                new[] { "Alpha, Inc", "2" },
                new[] { "Pipe | row", "multi\nline" },
                new[] { "Quote \"x\"", "3", "extra" }
            }
        };

        string csv = table.ToCsv();

        Assert.Equal("Name,Qty,Column3\r\n\"Alpha, Inc\",2,\r\nPipe | row,\"multi\nline\",\r\n\"Quote \"\"x\"\"\",3,extra", csv);
    }

    [Fact]
    public void ReaderTableExport_ToMarkdownTable_EscapesPipesAndLineBreaks() {
        var table = new ReaderTable {
            Columns = new[] { "Name", "Qty" },
            Rows = new[] {
                new[] { "Pipe | row", "multi\r\nline" },
                new[] { "Plain", "3" }
            }
        };

        string markdown = table.ToMarkdownTable();

        Assert.Equal("| Name | Qty |\n| --- | --- |\n| Pipe \\| row | multi<br>line |\n| Plain | 3 |", markdown.Replace("\r\n", "\n"));
    }

    [Fact]
    public void ReaderTableExport_ToJson_EmitsStableNormalizedTableShape() {
        var table = new ReaderTable {
            Title = "Inventory",
            Kind = "sample",
            Location = new ReaderLocation {
                Path = "inventory.md",
                Page = 2,
                TableIndex = 1
            },
            Columns = new[] { "Name" },
            Rows = new[] {
                new[] { "Alpha", "2" }
            },
            TotalRowCount = 5,
            Truncated = true,
            ColumnProfiles = new[] {
                new ReaderTableColumnProfile {
                    Index = 1,
                    Name = "Column2",
                    Kind = ReaderTableColumnKind.Numeric,
                    NonEmptyCellCount = 1,
                    NumericCellCount = 1,
                    Confidence = 1D
                }
            }
        };

        string json = table.ToJson();

        using JsonDocument document = JsonDocument.Parse(json);
        JsonElement root = document.RootElement;
        Assert.Equal("Inventory", root.GetProperty("title").GetString());
        Assert.Equal("sample", root.GetProperty("kind").GetString());
        Assert.Equal("inventory.md", root.GetProperty("location").GetProperty("path").GetString());
        Assert.Equal(2, root.GetProperty("location").GetProperty("page").GetInt32());
        Assert.Equal(1, root.GetProperty("location").GetProperty("tableIndex").GetInt32());
        Assert.True(root.GetProperty("truncated").GetBoolean());
        Assert.Equal(5, root.GetProperty("totalRowCount").GetInt32());
        Assert.Equal("Name", root.GetProperty("columns")[0].GetString());
        Assert.Equal("Column2", root.GetProperty("columns")[1].GetString());
        Assert.Equal("2", root.GetProperty("rows")[0][1].GetString());
        Assert.Equal("Numeric", root.GetProperty("columnProfiles")[0].GetProperty("kind").GetString());
    }
}
