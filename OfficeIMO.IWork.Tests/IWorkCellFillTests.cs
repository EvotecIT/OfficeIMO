using System.Text.Json;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Independent_selected_Numbers_fills_survive_saved_XLSX_including_styled_empty_cells() {
        using var manifest = JsonDocument.Parse(File.ReadAllText(CorpusFixture("numbers-cell-fills.json")));
        var expectedSource = manifest.RootElement.GetProperty("source");
        string path = CorpusFixture(expectedSource.GetProperty("path").GetString()!);
        Assert.Equal(expectedSource.GetProperty("sha256").GetString(), Convert.ToHexString(
            System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant());
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(path,
            conversionOptions: new IWorkConversionOptions { AllowPartialEditableReconstruction = true, NormalizeWorksheetNames = true });
        Assert.False(result.IsVisualFallback);
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = OfficeIMO.Excel.ExcelDocument.Load(saved);
        foreach (var expectedTable in manifest.RootElement.GetProperty("tables").EnumerateArray()) {
            var sourceSheet = Assert.Single(result.Projection.Sheets, sheet => sheet.Tables.Any(table =>
                table.ModelRecord!.Identifier == expectedTable.GetProperty("modelIdentifier").GetUInt64()));
            var table = Assert.Single(sourceSheet.Tables, table =>
                table.ModelRecord!.Identifier == expectedTable.GetProperty("modelIdentifier").GetUInt64());
            string destinationName = Assert.Single(result.WorksheetMappings, mapping =>
                mapping.SourceSheetIndex == result.Projection.Sheets.ToList().IndexOf(sourceSheet) + 1
                && mapping.SourceTableIndex == sourceSheet.Tables.ToList().IndexOf(table) + 1).DestinationName;
            var destination = Assert.Single(reopened.Sheets, sheet => sheet.Name == destinationName);
            foreach (var expected in expectedTable.GetProperty("cells").EnumerateArray()) {
                int row = expected.GetProperty("row").GetInt32(), column = expected.GetProperty("column").GetInt32();
                var source = Assert.IsType<IWorkTableCell>(table.GetCell(row, column));
                Assert.Equal(expected.GetProperty("empty").GetBoolean(), source.Kind == IWorkCellKind.Empty);
                var fill = expected.GetProperty("fill");
                string? rgb = fill.GetProperty("kind").GetString() == "none" ? null : fill.GetProperty("rgbHex").GetString();
                Assert.IsType<IWorkCellFill>(source.Fill);
                Assert.Equal(rgb, source.Fill!.Color?.RgbHex);
                Assert.Equal(rgb == null ? null : "FF" + rgb, destination.GetCellStyle(row, column).FillColorArgb);
            }
        }
        Assert.DoesNotContain(result.Report.Diagnostics, d => d.Code == "IWORK_NUMBERS_CELL_FILL_OMITTED");
    }

    [Fact]
    public void Independent_selected_Pages_cell_fills_and_layout_survive_saved_DOCX_including_empty_cells() {
        using var manifest = JsonDocument.Parse(File.ReadAllText(CorpusFixture("pages-cell-fills.json")));
        using var result = WordIWorkConverter.ConvertPagesToWordResult(CorpusFixture("picodocs/sample-v14.4.pages"),
            conversionOptions: new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
        Assert.False(result.IsVisualFallback);
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = OfficeIMO.Word.WordDocument.Load(saved);
        foreach (JsonElement native in manifest.RootElement.GetProperty("tables").EnumerateArray()) {
            IWorkTable table = Assert.Single(result.Projection.Tables,
                table => table.Name == native.GetProperty("name").GetString());
            int position = result.Projection.Tables.ToList().IndexOf(table);
            foreach (JsonElement cell in native.GetProperty("cells").EnumerateArray()) {
                int row = cell.GetProperty("row").GetInt32(), column = cell.GetProperty("column").GetInt32();
                IWorkTableCell source = Assert.IsType<IWorkTableCell>(table.GetCell(row, column));
                Assert.Equal(cell.GetProperty("empty").GetBoolean(), source.Kind == IWorkCellKind.Empty);
                IWorkCellFill fill = Assert.IsType<IWorkCellFill>(source.Fill);
                var target = reopened.Tables[position].Rows[row - 1].Cells[column - 1];
                JsonElement inset = cell.GetProperty("paddingPoints");
                IWorkCellPadding padding = Assert.IsType<IWorkCellPadding>(source.Padding);
                Assert.Equal(inset.GetProperty("left").GetDouble(), padding.LeftPoints);
                Assert.Equal(inset.GetProperty("top").GetDouble(), padding.TopPoints);
                Assert.Equal(inset.GetProperty("right").GetDouble(), padding.RightPoints);
                Assert.Equal(inset.GetProperty("bottom").GetDouble(), padding.BottomPoints);
                Assert.Equal((short)(padding.LeftPoints * 20), target.MarginLeftWidth);
                Assert.Equal((short)(padding.TopPoints * 20), target.MarginTopWidth);
                Assert.Equal((short)(padding.RightPoints * 20), target.MarginRightWidth);
                Assert.Equal((short)(padding.BottomPoints * 20), target.MarginBottomWidth);
                Assert.Equal("middle", cell.GetProperty("verticalAlignment").GetString());
                Assert.Equal(IWorkCellVerticalAlignment.Middle, source.VerticalAlignment);
                Assert.Equal(OfficeIMO.Word.WordTableVerticalAlignment.Center, target.VerticalAlignment);
                JsonElement expected = cell.GetProperty("fill");
                if (expected.GetProperty("kind").GetString() == "none") {
                    Assert.True(fill.IsNone);
                    Assert.Equal(OfficeIMO.Word.WordShadingPattern.Nil, target.ShadingPattern);
                } else {
                    string hex = expected.GetProperty("rgbHex").GetString()!;
                    Assert.Equal(hex, fill.Color!.RgbHex);
                    Assert.Equal(hex, target.ShadingFillColorHex);
                }
            }
        }
        Assert.DoesNotContain(result.Report.Diagnostics, diagnostic => diagnostic.Code == "IWORK_TABLE_CELL_FILL_UNSUPPORTED");
        saved.Position = 0;
        using var xml = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(saved, false);
        var validator = new DocumentFormat.OpenXml.Validation.OpenXmlValidator();
        Assert.All(xml.MainDocumentPart!.Document!.Descendants<DocumentFormat.OpenXml.Wordprocessing.Shading>(),
            shading => Assert.Empty(validator.Validate(shading)));
        Assert.All(xml.MainDocumentPart.Document.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableCellMargin>(),
            margins => Assert.Empty(validator.Validate(margins)));
        Assert.All(xml.MainDocumentPart.Document.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableCellVerticalAlignment>(),
            alignment => Assert.Empty(validator.Validate(alignment)));
    }

    [Theory]
    [InlineData("inherit", "FF0000")]
    [InlineData("solid", "0000FF")]
    [InlineData("none", null)]
    public void Selected_cell_fill_inheritance_and_explicit_clear_survive_sparse_Pages_conversion(string mode, string? expected) {
        byte[]? declaration = mode == "inherit" ? null : mode == "none" ? Array.Empty<byte>() : FillColor(0, 0, 1);
        using var package = CellFillPackage(new[] { FillStyle(30, declaration, 31), FillStyle(31, FillColor(1, 0, 0)) });
        IWorkSourceDocument source = IWorkSourceDocument.Open(package);
        using var result = source.ToWordDocumentResult();
        Assert.False(result.IsVisualFallback);
        IWorkTable table = Assert.Single(result.Projection.Tables);
        Assert.Equal(2, table.Cells.Count);
        Assert.Equal(IWorkCellKind.Empty, table.GetCell(1, 2)!.Kind);
        Assert.All(table.Cells, cell => Assert.Equal(expected, cell.Fill!.Color?.RgbHex));
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = OfficeIMO.Word.WordDocument.Load(saved);
        Assert.All(reopened.Tables[0].Rows[0].Cells, cell => {
            if (expected == null) Assert.Equal(OfficeIMO.Word.WordShadingPattern.Nil, cell.ShadingPattern);
            else Assert.Equal(expected, cell.ShadingFillColorHex);
        });
    }

    [Theory]
    [InlineData("missing-style")]
    [InlineData("wrong-style-type")]
    [InlineData("missing-parent")]
    [InlineData("cycle")]
    [InlineData("depth")]
    [InlineData("duplicate-super")]
    [InlineData("duplicate-fill")]
    [InlineData("bad-color")]
    [InlineData("transparent")]
    [InlineData("p3")]
    [InlineData("gradient")]
    [InlineData("ambiguous-catalog")]
    [InlineData("wrong-catalog-type")]
    public void Unsupported_selected_cell_fills_preserve_values_and_prevent_complete_editable_output(string defect) {
        byte[] style = defect switch {
            "wrong-style-type" => ArchiveRecord(30, 2022, Message()),
            "missing-parent" => FillStyle(30, FillColor(1, 0, 0), 99),
            "cycle" => FillStyle(30, FillColor(1, 0, 0), 30),
            "depth" => FillStyle(30, null, 31),
            "duplicate-super" => ArchiveRecord(30, 6004, Message(BytesField(1, ReferenceField(3, 31)),
                BytesField(1, ReferenceField(3, 31)), BytesField(11, BytesField(1, FillColor(1, 0, 0))))),
            "duplicate-fill" => ArchiveRecord(30, 6004, BytesField(11, Message(
                BytesField(1, FillColor(1, 0, 0)), BytesField(1, Message())))),
            "bad-color" => FillStyle(30, FillColor(float.NaN, 0, 0)),
            "transparent" => FillStyle(30, FillColor(1, 0, 0, .5f)),
            "p3" => FillStyle(30, FillColor(1, 0, 0, space: 2)),
            "gradient" => FillStyle(30, BytesField(2, Message())),
            _ => FillStyle(30, FillColor(1, 0, 0))
        };
        using var package = CellFillPackage(defect == "missing-style" ? Array.Empty<byte[]>()
                : new[] { style, FillStyle(31, FillColor(0, 0, 1)) },
            ambiguous: defect == "ambiguous-catalog", listType: defect == "wrong-catalog-type" ? 2 : 4);
        var source = IWorkSourceDocument.Open(package, new IWorkReadOptions {
            MaximumTextStyleInheritanceDepth = defect == "depth" ? 1 : 128
        });
        using var result = source.ToWordDocumentResult(options:new IWorkConversionOptions { RequireCompleteVisualCoverage = false });
        Assert.True(result.IsVisualFallback);
        IWorkTable table = Assert.Single(result.Projection.Tables);
        Assert.Equal(42d, table.GetCell(1, 1)!.Value);
        Assert.Null(table.GetCell(1, 1)!.Fill);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "IWORK_TABLE_CELL_FILL_UNSUPPORTED"
            && diagnostic.LossKind == global::OfficeIMO.OfficeConversionLossKind.Unassessed);
        Assert.NotEmpty(result.Report.SourceDeclarationIssues.Concat<object>(result.Report.SourceReferenceIssues));
    }

    [Fact]
    public void Selected_cell_fill_catalog_and_styled_empty_cells_obey_existing_projection_limits() {
        byte[][] styles = { FillStyle(30, FillColor(1, 0, 0)) };
        using var package = CellFillPackage(styles);
        Assert.Throws<InvalidDataException>(() => IWorkSourceDocument.Open(package,
            new IWorkReadOptions { MaximumMaterializedCells = 1 }).ReadPages());
        package.Position = 0;
        using var moreEntries = CellFillPackage(styles, additionalEntry: true);
        Assert.Throws<InvalidDataException>(() => IWorkSourceDocument.Open(moreEntries,
            new IWorkReadOptions { MaximumTableCatalogEntries = 1 }).ReadPages());
    }

    [Fact]
    public void Unselected_cell_style_values_are_not_traversed() {
        using var package = CellFillPackage(new[] { FillStyle(30, FillColor(1, 0, 0)) }, additionalEntry: true);
        IWorkPagesProjection projection = IWorkSourceDocument.Open(package).ReadPages();
        Assert.True(projection.HasEditableContent);
        Assert.DoesNotContain(projection.SourceReferenceIssues, issue => issue.TargetIdentifier == 99);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Numbers, "inherit", "FF0000")]
    [InlineData(IWorkDocumentKind.Numbers, "solid", "0000FF")]
    [InlineData(IWorkDocumentKind.Numbers, "none", null)]
    [InlineData(IWorkDocumentKind.Keynote, "inherit", "FF0000")]
    [InlineData(IWorkDocumentKind.Keynote, "solid", "0000FF")]
    [InlineData(IWorkDocumentKind.Keynote, "none", null)]
    public void Selected_cell_fills_survive_saved_destinations_and_Reader_empty_cells_remain_empty(IWorkDocumentKind kind, string mode, string? expected) {
        byte[]? declaration = mode == "inherit" ? null : mode == "none" ? Array.Empty<byte>() : FillColor(0, 0, 1);
        using var package = CellFillPackage(new[] { FillStyle(30, declaration, 31), FillStyle(31, FillColor(1, 0, 0)) }, kind: kind);
        var source = IWorkSourceDocument.Open(package);
        var policy = new IWorkConversionOptions { AllowPartialEditableReconstruction = true };
        using var saved = new MemoryStream();
        if (kind == IWorkDocumentKind.Numbers) {
            using var result = source.ToExcelDocumentResult(policy);
            Assert.False(result.IsVisualFallback);
            Assert.DoesNotContain(result.Report.Diagnostics, diagnostic => diagnostic.Code == "IWORK_NUMBERS_CELL_FILL_OMITTED");
            result.Value.Save(saved); saved.Position = 0;
            using var reopened = OfficeIMO.Excel.ExcelDocument.Load(saved);
            Assert.All(new[] { 1, 2 }, column => Assert.Equal(expected == null ? null : "FF" + expected,
                reopened.Sheets[0].GetCellStyle(1, column).FillColorArgb));
            Assert.True(reopened.Sheets[0].TryGetCellValueSnapshot(1, 1, out var value));
            Assert.Equal(OfficeIMO.Excel.ExcelCellValueKind.Number, value!.Kind); Assert.Equal("42", value.RawValue);
        } else {
            using var result = source.ToPowerPointPresentationResult(policy);
            Assert.False(result.IsVisualFallback);
            Assert.DoesNotContain(result.Report.Diagnostics, diagnostic => diagnostic.Code == "IWORK_KEYNOTE_CELL_FILL_OMITTED");
            result.Value.Save(saved); saved.Position = 0;
            using var reopened = OfficeIMO.PowerPoint.PowerPointPresentation.Load(saved);
            var table = Assert.Single(reopened.Slides[0].Tables);
            Assert.Equal("42", table.GetCell(0, 0).Text);
            Assert.Equal(string.Empty, table.GetCell(0, 1).Text);
            Assert.All(new[] { 0, 1 }, column => {
                Assert.Equal(expected, table.GetCell(0, column).FillColor);
                Assert.Equal(expected == null, table.GetCell(0, column).NoFill);
            });
            Assert.Empty(reopened.ValidateDocument());
        }
        package.Position = 0;
        var reader = new OfficeIMO.Reader.OfficeDocumentReaderBuilder();
        OfficeIMO.Reader.IWork.OfficeDocumentReaderBuilderIWorkExtensions.AddIWorkHandler(reader);
        var read = reader.Build().ReadDocument(package, kind == IWorkDocumentKind.Numbers ? "fill.numbers" : "fill.key");
        Assert.Equal(new[] { "42", string.Empty }, Assert.Single(Assert.Single(read.Tables).Rows));
        Assert.Contains(read.Diagnostics, diagnostic => diagnostic.Code == "IWORK_READER_TABLE_STYLE_PARTIAL");
        if (kind == IWorkDocumentKind.Keynote) {
            var counts = Assert.Single(read.Tables).Diagnostics!;
            Assert.Equal(1, counts.FilledCellCount);
            Assert.Equal(1, counts.MissingCellCount);
        }
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages, false)]
    [InlineData(IWorkDocumentKind.Numbers, false)]
    [InlineData(IWorkDocumentKind.Keynote, false)]
    [InlineData(IWorkDocumentKind.Pages, true)]
    [InlineData(IWorkDocumentKind.Numbers, true)]
    [InlineData(IWorkDocumentKind.Keynote, true)]
    public void Styled_covered_merge_cells_block_conversion_only_when_they_have_content(IWorkDocumentKind kind, bool coveredContent) {
        using var package = CellFillPackage(new[] { FillStyle(30, Array.Empty<byte>()) }, kind: kind,
            merge: true, coveredContent: coveredContent);
        var source = IWorkSourceDocument.Open(package);
        var policy = new IWorkConversionOptions { AllowPartialEditableReconstruction = true };
        if (kind == IWorkDocumentKind.Pages) {
            using var result = source.ToWordDocumentResult(IWorkTestPolicy.ForIncompletePreview(policy));
            Assert.Equal(coveredContent, result.IsVisualFallback);
            if (!coveredContent) {
                Assert.Equal(2, result.Value.Tables[0].Rows[0].Cells[0].ColumnSpan);
                Assert.Equal("42", result.Value.Tables[0].Rows[0].Cells[0].Paragraphs[0].Text);
            }
        } else if (kind == IWorkDocumentKind.Numbers) {
            using var result = source.ToExcelDocumentResult(IWorkTestPolicy.ForIncompletePreview(policy));
            Assert.Equal(coveredContent, result.IsVisualFallback);
            if (!coveredContent) Assert.Equal("A1:B1", Assert.Single(result.Value.Sheets[0].GetMergedRanges()).A1Range);
        } else {
            using var result = source.ToPowerPointPresentationResult(IWorkTestPolicy.ForIncompletePreview(policy));
            Assert.Equal(coveredContent, result.IsVisualFallback);
            if (!coveredContent) Assert.Equal((1, 2), Assert.Single(result.Value.Slides[0].Tables).GetCell(0, 0).Merge);
        }
    }

    private static byte[] FillColor(float red, float green, float blue, float alpha = 1, ulong space = 1) =>
        BytesField(1, Message(VarintField(1, 1), FloatField(3, red), FloatField(4, green),
            FloatField(5, blue), FloatField(6, alpha), VarintField(12, space)));

    private static byte[] FillStyle(ulong id, byte[]? fill, ulong? parent = null) =>
        ArchiveRecord(id, 6004, Message(parent.HasValue ? BytesField(1, ReferenceField(3, parent.Value)) : Array.Empty<byte>(),
            fill == null ? Array.Empty<byte>() : BytesField(11, BytesField(1, fill))));

    private static MemoryStream CellFillPackage(byte[][] styles, bool ambiguous = false,
        int listType = 4, bool additionalEntry = false, IWorkDocumentKind kind = IWorkDocumentKind.Pages,
        bool merge = false, bool coveredContent = false) {
        byte[] first = new byte[24]; first[0] = 5; first[1] = 2; first[8] = 0x22;
        Buffer.BlockCopy(BitConverter.GetBytes(42d), 0, first, 12, 8); first[20] = 1;
        byte[] second = new byte[16]; second[0] = 5; second[8] = 0x20; second[12] = 1;
        if (coveredContent) {
            second = new byte[24]; second[0] = 5; second[1] = 2; second[8] = 0x22;
            Buffer.BlockCopy(BitConverter.GetBytes(7d), 0, second, 12, 8); second[20] = 1;
        }
        byte[] tile = Message(VarintField(1, 1), VarintField(2, 0), VarintField(3, 2), VarintField(4, 1),
            BytesField(5, Message(VarintField(1, 0), VarintField(2, 2),
                BytesField(6, Message(first, second)), BytesField(7, new byte[] { 0, 0, 24, 0 }))),
            VarintField(6, 5), VarintField(7, 1));
        var records = new List<byte[]>();
        if (kind == IWorkDocumentKind.Pages) {
            records.Add(ArchiveRecord(1, 10000, ReferenceField(4, 2), new ulong[] { 2, 10 }));
            records.Add(ArchiveRecord(2, 2001, StringField(3, "Body")));
        } else if (kind == IWorkDocumentKind.Numbers) {
            records.Add(ArchiveRecord(1, 1, ReferenceField(1, 2)));
            records.Add(ArchiveRecord(2, 2, Message(StringField(1, "Sheet"), ReferenceField(2, 10))));
        } else {
            records.Add(ArchiveRecord(1, 1, ReferenceField(2, 2)));
            records.Add(ArchiveRecord(2, 2, KeynoteShow(ReferenceField(2, 3))));
            records.Add(ArchiveRecord(3, 4, ReferenceField(2, 4)));
            records.Add(ArchiveRecord(4, 5, ReferenceField(6, 10)));
        }
        records.Add(ArchiveRecord(10, 6000, Message(ReferenceField(2, 11),
            kind == IWorkDocumentKind.Keynote ? BytesField(1, GeometryDrawable(0, 0, 144, 24)) : Array.Empty<byte>())));
        byte[] tileList = BytesField(3, BytesField(1, Message(VarintField(1, 0), ReferenceField(2, 12))));
        byte[] store = Message(ReferenceField(5, 13), tileList);
        records.Add(ArchiveRecord(11, 6001, Message(VarintField(6, 1), VarintField(7, 2), BytesField(4, store),
            merge ? MergeTestOwner(MergeTestPair(0, 0, 0, 1)) : Array.Empty<byte>())));
        records.Add(ArchiveRecord(12, 6002, tile));
        byte[] entry = Message(VarintField(1, 1), ReferenceField(4, 30));
        records.Add(ArchiveRecord(13, 6005, Message(VarintField(1, (ulong)listType), BytesField(3, entry),
            ambiguous ? BytesField(3, entry) : Array.Empty<byte>(),
            additionalEntry ? BytesField(3, Message(VarintField(1, 2), ReferenceField(4, 99))) : Array.Empty<byte>())));
        records.AddRange(styles);
        return CreatePackage(("Index/Document.iwa", FrameIwa(Message(records.ToArray()))), ("preview.png", ValidPreviewPng()));
    }
}
