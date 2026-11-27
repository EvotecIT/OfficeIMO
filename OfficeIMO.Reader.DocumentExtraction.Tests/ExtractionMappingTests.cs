using Microsoft.Extensions.AI;
using Microsoft.Extensions.Documents;
using Xunit;

namespace OfficeIMO.Reader.DocumentExtraction.Tests;

public sealed class ExtractionMappingTests {
    [Fact]
    public void HeadingsOpenNestedSectionsAndEarlierContentGetsARootSection() {
        OfficeDocumentReadResult result = Result(ReaderInputKind.Word,
            Block("paragraph", "Preamble"),
            Block("heading", "Chapter", level: 1),
            Block("paragraph", "Chapter body"),
            Block("heading", "Detail", level: 2),
            Block("paragraph", "Detail body"),
            Block("heading", "Next chapter", level: 1));

        ExtractionDocument document = result.ToExtractionDocument("doc");

        Assert.Equal(3, document.Sections.Count);
        Assert.Equal("Preamble", Text(Assert.Single(document.Sections[0].Elements)));

        ExtractionDocumentSection chapter = document.Sections[1];
        var chapterHeading = Assert.IsType<ExtractionDocumentTextElement>(chapter.Elements[0]);
        Assert.Equal("heading", chapterHeading.Role);
        Assert.Equal(1, chapterHeading.HeadingLevel);
        Assert.Equal("Chapter body", Text(chapter.Elements[1]));
        var detail = Assert.IsType<ExtractionDocumentSection>(chapter.Elements[2]);
        Assert.Equal(2, Assert.IsType<ExtractionDocumentTextElement>(detail.Elements[0]).HeadingLevel);
        Assert.Equal("Detail body", Text(detail.Elements[1]));

        Assert.Equal("Next chapter", Text(Assert.Single(document.Sections[2].Elements)));
        Assert.Null(chapter.Content);
    }

    [Fact]
    public void ListItemsKeepTheirRoleLevelMarkerAndIndex() {
        OfficeDocumentBlock item = Block("list-item", "Second step", level: 1);
        item.Marker = "2.";
        item.ListIndex = 2;

        ExtractionDocument document = Result(ReaderInputKind.Word, item).ToExtractionDocument("doc");

        var element = Assert.IsType<ExtractionDocumentTextElement>(Assert.Single(document.EnumerateContent()));
        Assert.Equal("listItem", element.Role);
        Assert.Null(element.HeadingLevel);
        AdditionalPropertiesDictionary properties = element.Content!.AdditionalProperties!;
        Assert.Equal(1, properties[OfficeIMOExtractionPropertyNames.ListLevel]);
        Assert.Equal("2.", properties[OfficeIMOExtractionPropertyNames.ListMarker]);
        Assert.Equal(2, properties[OfficeIMOExtractionPropertyNames.ListIndex]);
    }

    [Theory]
    [InlineData("paragraph", null)]
    [InlineData("header", "pageHeader")]
    [InlineData("footer", "pageFooter")]
    [InlineData("footnote", "footnote")]
    [InlineData("speaker-notes", "speaker-notes")]
    public void BlockKindsMapToSharedOrPreservedRoles(string kind, string? expectedRole) {
        ExtractionDocument document = Result(ReaderInputKind.Word, Block(kind, "Text")).ToExtractionDocument("doc");

        var element = Assert.IsType<ExtractionDocumentTextElement>(Assert.Single(document.EnumerateContent()));
        Assert.Equal(expectedRole, element.Role);
        Assert.Equal(kind, element.Content!.AdditionalProperties![OfficeIMOExtractionPropertyNames.BlockKind]);
    }

    [Fact]
    public void TablesKeepAHeaderRowAndAddressableCells() {
        var table = new ReaderTable {
            Title = "Totals",
            Columns = new[] { "Region", "Total" },
            Rows = new IReadOnlyList<string>[] { new[] { "North", "10" }, new[] { "South", "" } },
            TotalRowCount = 5,
            Truncated = true,
            Location = new ReaderLocation { BlockIndex = 0 }
        };
        OfficeDocumentReadResult result = Result(ReaderInputKind.Word);
        result.Tables = new[] { table };

        ExtractionDocument document = result.ToExtractionDocument("doc");

        var element = Assert.IsType<ExtractionDocumentTableElement>(Assert.Single(document.EnumerateContent()));
        Assert.Equal(3, element.Cells.RowCount);
        Assert.Equal(2, element.Cells.ColumnCount);
        Assert.Equal("columnHeader", element.Cells[0, 0]!.Role);
        Assert.Equal("Region", Text(element.Cells[0, 0]!.Content!));
        Assert.Null(element.Cells[1, 0]!.Role);
        Assert.Equal("10", Text(element.Cells[1, 1]!.Content!));
        Assert.Null(element.Cells[2, 1]!.Content);
        Assert.Equal(1, element.Cells[2, 1]!.RowSpan);
        Assert.Same(table, element.Content!.RawRepresentation);
        AdditionalPropertiesDictionary properties = element.Content.AdditionalProperties!;
        Assert.Equal("Totals", properties[OfficeIMOExtractionPropertyNames.TableTitle]);
        Assert.Equal(true, properties[OfficeIMOExtractionPropertyNames.TableTruncated]);
        Assert.Equal(5, properties[OfficeIMOExtractionPropertyNames.TableTotalRowCount]);
    }

    [Fact]
    public void IdenticalTablesInDifferentContainersAreBothKept() {
        static ReaderTable Copy(ReaderLocation location) => new ReaderTable {
            Columns = new[] { "Name" }, Rows = new IReadOnlyList<string>[] { new[] { "Bolt" } }, Location = location
        };
        OfficeDocumentReadResult result = Result(ReaderInputKind.Pdf);
        result.Tables = new[] {
            Copy(new ReaderLocation { Page = 1, BlockIndex = 0, BlockAnchor = "page-1-table-0" }),
            Copy(new ReaderLocation { Page = 2, BlockIndex = 0 })
        };

        ExtractionDocument document = result.ToExtractionDocument("doc");

        ExtractionDocumentTableElement[] tables = document.EnumerateContent().OfType<ExtractionDocumentTableElement>().ToArray();
        Assert.Equal(new int?[] { 1, 2 }, tables.Select(table => Assert.Single(table.Sources).Page.PageNumber).OrderBy(page => page).ToArray());
    }

    [Fact]
    public void TableBlocksWithoutAStructuredTableKeepTheirText() {
        ExtractionDocument document = Result(ReaderInputKind.Word, Block("table", "Name | Qty")).ToExtractionDocument("doc");

        var element = Assert.IsType<ExtractionDocumentTextElement>(Assert.Single(document.EnumerateContent()));
        Assert.Equal("table", element.Role);
        Assert.Equal("Name | Qty", Text(element));
    }

    [Fact]
    public void TablesAboveTheCellLimitBecomeTabSeparatedText() {
        OfficeDocumentReadResult result = Result(ReaderInputKind.Word);
        result.Tables = new[] {
            new ReaderTable {
                Columns = new[] { "A", "B" },
                Rows = new IReadOnlyList<string>[] { new[] { "1", "2" }, new[] { "3", "4", "5" } },
                Kind = "data", Truncated = true, TotalRowCount = 10,
                Location = new ReaderLocation { BlockIndex = 0 }
            },
            new ReaderTable {
                Columns = new[] { "Column 1" },
                Rows = new IReadOnlyList<string>[] { new[] { "" }, new[] { "value" } },
                Location = new ReaderLocation { BlockIndex = 1 }
            }
        };

        ExtractionDocument document = result.ToExtractionDocument("doc", new OfficeIMODocumentExtractionOptions { MaxTableCells = 1 });

        ExtractionDocumentTextElement[] elements = document.EnumerateContent().Cast<ExtractionDocumentTextElement>().ToArray();
        ExtractionDocumentTextElement element = elements[0];
        Assert.Equal("table", element.Role);
        Assert.Equal("A\tB\n1\t2\n3\t4\t5", Text(element));
        AdditionalPropertiesDictionary properties = element.Content!.AdditionalProperties!;
        Assert.Equal(true, properties[OfficeIMOExtractionPropertyNames.TableGridLimitExceeded]);
        Assert.Equal(3, properties[OfficeIMOExtractionPropertyNames.TableRowCount]);
        Assert.Equal(3, properties[OfficeIMOExtractionPropertyNames.TableColumnCount]);
        Assert.Equal("data", properties[OfficeIMOExtractionPropertyNames.TableKind]);
        Assert.Equal(true, properties[OfficeIMOExtractionPropertyNames.TableTruncated]);
        Assert.Equal(10, properties[OfficeIMOExtractionPropertyNames.TableTotalRowCount]);
        Assert.Equal("\nvalue", Text(Assert.Single(elements.Skip(1))));
    }

    [Fact]
    public void ImagesCarryPayloadAltTextAndOcrTextWithoutRepeatingTheOcrBlock() {
        byte[] png = { 137, 80, 78, 71, 13, 10, 26, 10 };
        OfficeDocumentBlock caption = Block("paragraph", "Figure follows");
        caption.Location.BlockAnchor = "image-anchor";
        OfficeDocumentBlock ocr = Block("ocr-text", "Scanned words");
        ocr.Id = "candidate-1-block";
        ocr.Recognition = new OfficeDocumentRecognitionEvidence(provider: "tesseract", confidence: 0.9);
        OfficeDocumentReadResult result = Result(ReaderInputKind.Word, caption, ocr);
        result.Assets = new[] {
            new OfficeDocumentAsset {
                Id = "asset-1", Kind = "image", MediaType = "image/png", FileName = "scan.png",
                AltText = "Scanned page", PayloadBytes = png, Width = 10, Height = 20,
                Location = new ReaderLocation { BlockAnchor = "image-anchor" }
            }
        };
        result.OcrCandidates = new[] { new OfficeDocumentOcrCandidate { Id = "candidate-1", Kind = "image", AssetId = "asset-1" } };

        ExtractionDocument document = result.ToExtractionDocument("doc");

        ExtractionDocumentElement[] content = document.EnumerateContent().ToArray();
        Assert.Equal(2, content.Length);
        Assert.Equal("Figure follows", Text(content[0]));
        var image = Assert.IsType<ExtractionDocumentImageElement>(content[1]);
        Assert.Equal("Scanned page", image.AlternativeText);
        Assert.Equal("Scanned words", image.Text);
        var data = Assert.IsType<DataContent>(image.Content);
        Assert.Equal("image/png", data.MediaType);
        Assert.Equal(png, data.Data.ToArray());
        Assert.Equal("tesseract", data.AdditionalProperties![OfficeIMOExtractionPropertyNames.OcrProvider]);
        Assert.Equal(10, data.AdditionalProperties[OfficeIMOExtractionPropertyNames.ImageWidth]);
    }

    [Fact]
    public void ImageBlocksPairWithAssetsByAnchorBeforeOrder() {
        OfficeDocumentBlock first = Block("image", "First");
        first.Location.BlockAnchor = "img-a";
        OfficeDocumentBlock second = Block("image", "Second");
        second.Location.BlockAnchor = "img-b";
        OfficeDocumentReadResult result = Result(ReaderInputKind.Html, first, second);
        result.Assets = new[] {
            new OfficeDocumentAsset { Id = "b", Kind = "image", MediaType = "image/png", Location = new ReaderLocation { BlockAnchor = "img-b" } },
            new OfficeDocumentAsset { Id = "a", Kind = "image", MediaType = "image/png", Location = new ReaderLocation { BlockAnchor = "img-a" } }
        };

        ExtractionDocument document = result.ToExtractionDocument("doc");

        string?[] assetIds = document.EnumerateContent().Cast<ExtractionDocumentImageElement>()
            .Select(image => image.Content!.AdditionalProperties![OfficeIMOExtractionPropertyNames.AssetId] as string).ToArray();
        Assert.Equal(new[] { "a", "b" }, assetIds);
    }

    [Fact]
    public void ImagePayloadsCanBeLeftOut() {
        OfficeDocumentReadResult result = Result(ReaderInputKind.Word);
        result.Assets = new[] { new OfficeDocumentAsset { Id = "a", Kind = "image", MediaType = "image/png", PayloadBytes = new byte[] { 1 } } };

        ExtractionDocument document = result.ToExtractionDocument("doc", new OfficeIMODocumentExtractionOptions { IncludeImagePayloads = false });

        var image = Assert.IsType<ExtractionDocumentImageElement>(Assert.Single(document.EnumerateContent()));
        Assert.IsNotType<DataContent>(image.Content);
        Assert.Equal("a", image.Content!.AdditionalProperties![OfficeIMOExtractionPropertyNames.AssetId]);
    }

    [Fact]
    public void PdfPagesBecomeFramesAndRegionsBecomeBottomLeftPolygons() {
        OfficeDocumentBlock heading = Block("heading", "Title", level: 1);
        heading.Location.Page = 1;
        heading.Region = new OfficeDocumentRegion { X = 10, Y = 700, Width = 200, Height = 0 };
        OfficeDocumentBlock figure = Block("paragraph", "Boxed");
        figure.Location.Page = 1;
        figure.Region = new OfficeDocumentRegion { X = 10, Y = 600, Width = 100, Height = 50 };
        OfficeDocumentReadResult result = Result(ReaderInputKind.Pdf, heading, figure);
        result.Pages = new[] {
            new OfficeDocumentPage { Number = 1, Width = 612, Height = 792 },
            new OfficeDocumentPage { Number = 2, Width = 612, Height = 792 }
        };

        ExtractionDocument document = result.ToExtractionDocument("doc");

        Assert.Equal(new[] { "page-1", "page-2" }, document.SourcePages.Select(page => page.Identifier).ToArray());
        ExtractionDocumentSourcePage first = document.SourcePages[0];
        Assert.Equal(1, first.PageNumber);
        Assert.Equal(new ExtractionDocumentPageDimensions(612, 792), first.Dimensions);
        Assert.Equal(ExtractionDocumentCoordinateUnit.Point, first.CoordinateUnit);
        Assert.Equal(ExtractionDocumentCoordinateOrigin.BottomLeft, first.CoordinateOrigin);

        ExtractionDocumentElement[] content = document.EnumerateContent().ToArray();
        ExtractionDocumentSourceReference baseline = Assert.Single(content[0].Sources);
        Assert.Same(first, baseline.Page);
        Assert.Empty(baseline.Polygon);
        Assert.Equal(
            new[] { new ExtractionDocumentPoint(10, 650), new ExtractionDocumentPoint(110, 650), new ExtractionDocumentPoint(110, 600), new ExtractionDocumentPoint(10, 600) },
            Assert.Single(content[1].Sources).Polygon.ToArray());
    }

    [Fact]
    public void PdfFramesReportTheirCoordinateSystemWithoutKnownDimensions() {
        OfficeDocumentBlock block = Block("paragraph", "Boxed");
        block.Location.Page = 3;
        block.Region = new OfficeDocumentRegion { X = 0, Y = 0, Width = 10, Height = 10 };

        ExtractionDocument document = Result(ReaderInputKind.Pdf, block).ToExtractionDocument("doc");

        ExtractionDocumentSourcePage page = Assert.Single(document.SourcePages);
        Assert.Null(page.Dimensions);
        Assert.Equal(ExtractionDocumentCoordinateUnit.Point, page.CoordinateUnit);
        Assert.Equal(ExtractionDocumentCoordinateOrigin.BottomLeft, page.CoordinateOrigin);
        Assert.Equal(4, Assert.Single(Assert.Single(document.EnumerateContent()).Sources).Polygon.Count);
    }

    [Fact]
    public void SheetAndSlideLocationsBecomeFramesAndProperties() {
        OfficeDocumentBlock cell = Block("paragraph", "Cell text");
        cell.Location.Sheet = "Budget";
        cell.Location.A1Range = "B2:C4";
        OfficeDocumentBlock slide = Block("paragraph", "Slide text");
        slide.Location.Slide = 3;

        ExtractionDocument document = Result(ReaderInputKind.Unknown, cell, slide).ToExtractionDocument("doc");

        // Content in different containers has no shared order, so look elements up by text.
        ExtractionDocumentElement cellElement = document.EnumerateContent().Single(element => Text(element) == "Cell text");
        ExtractionDocumentElement slideElement = document.EnumerateContent().Single(element => Text(element) == "Slide text");
        ExtractionDocumentSourcePage sheet = Assert.Single(cellElement.Sources).Page;
        Assert.Equal("sheet-Budget", sheet.Identifier);
        Assert.Null(sheet.PageNumber);
        Assert.Null(sheet.CoordinateUnit);
        Assert.Equal("B2:C4", cellElement.Content!.AdditionalProperties![OfficeIMOExtractionPropertyNames.CellRange]);
        ExtractionDocumentSourcePage slideFrame = Assert.Single(slideElement.Sources).Page;
        Assert.Equal("slide-3", slideFrame.Identifier);
        Assert.Equal(3, slideFrame.PageNumber);
        Assert.Equal(2, document.SourcePages.Count);
    }

    [Fact]
    public void DocumentMetadataLinksFormsAndRawResultAreKept() {
        OfficeDocumentReadResult result = Result(ReaderInputKind.Pdf, Block("paragraph", "Body"));
        result.Source.Title = "Annual report";
        result.Source.Author = "Finance";
        result.Metadata = new[] { new OfficeDocumentMetadataEntry { Category = "custom", Name = "Department", Value = "Sales" } };
        result.Links = new[] { new OfficeDocumentLink { Kind = "uri", Text = "site", Uri = "https://example.test/" } };
        result.Forms = new[] { new OfficeDocumentFormField { Kind = "text", Name = "Name", Value = "Ada", IsRequired = true } };

        ExtractionDocument document = result.ToExtractionDocument("report.pdf");

        Assert.Equal("report.pdf", document.Identifier);
        Assert.Same(result, document.RawRepresentation);
        Assert.Equal("Annual report", document.Metadata["title"]);
        Assert.Equal("Finance", document.Metadata["author"]);
        Assert.Equal("Sales", document.Metadata["custom.Department"]);
        Assert.Equal("Pdf", document.Metadata[OfficeIMOExtractionPropertyNames.InputKind]);
        var link = Assert.Single(Assert.IsType<List<Dictionary<string, object?>>>(document.Metadata[OfficeIMOExtractionPropertyNames.Links]));
        Assert.Equal("https://example.test/", link["uri"]);
        var field = Assert.Single(Assert.IsType<List<Dictionary<string, object?>>>(document.Metadata[OfficeIMOExtractionPropertyNames.FormFields]));
        Assert.Equal("Ada", field["value"]);
        Assert.Equal(true, field["required"]);
    }

    internal static string Text(ExtractionDocumentElement element) =>
        Assert.IsType<TextContent>(Assert.IsType<ExtractionDocumentTextElement>(element).Content).Text;

    private static OfficeDocumentReadResult Result(ReaderInputKind kind, params OfficeDocumentBlock[] blocks) {
        for (int i = 0; i < blocks.Length; i++) {
            if (string.IsNullOrEmpty(blocks[i].Id)) blocks[i].Id = "block-" + i;
            blocks[i].Location.BlockIndex = i;
        }
        return new OfficeDocumentReadResult { Kind = kind, Blocks = blocks };
    }

    private static OfficeDocumentBlock Block(string kind, string text, int? level = null) =>
        new OfficeDocumentBlock { Kind = kind, Text = text, Level = level };
}
