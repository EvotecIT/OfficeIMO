using System.Threading;
using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.PowerPoint;
using OfficeIMO.Reader;
using OfficeIMO.Reader.IWork;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(IWorkDocumentKind.Pages, "sample.pages")]
    [InlineData(IWorkDocumentKind.Numbers, "sample.numbers")]
    [InlineData(IWorkDocumentKind.Keynote, "sample.key")]
    public void Reader_reports_unrepresented_grid_borders_even_without_materialized_value_cells(IWorkDocumentKind kind, string name) {
        using var package = BorderPackage(kind, sparse: false);
        var result = IWorkReaderAdapter.ReadDocument(package, name, new ReaderOptions(), new ReaderIWorkOptions(), CancellationToken.None);
        Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "IWORK_READER_TABLE_STYLE_PARTIAL");
        Assert.Single(result.Tables);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Header_and_footer_vertical_regions_cover_the_header_column_intersection(IWorkDocumentKind kind) {
        using var package = DimensionPackage(kind, tableStyleReference: ReferenceField(3, 40),
            modelDeclarations: Message(VarintField(9, 1), VarintField(10, 1), VarintField(11, 1)),
            styleRecords: new[] { ArchiveRecord(40, 6003, BytesField(11, Message(
                BytesField(49, BorderStroke(3, 1, 0, 0)), BytesField(51, BorderStroke(1)), BytesField(57, BorderStroke(2, 0, 0, 1))))) });
        var source = IWorkSourceDocument.Open(package);
        var table = kind switch {
            IWorkDocumentKind.Pages => source.ReadPages().Tables[0],
            IWorkDocumentKind.Numbers => source.ReadNumbers().Sheets[0].Tables[0],
            _ => source.ReadKeynote().Slides[0].Tables[0]
        };
        AssertBorder(table.GetBorders(1, 1).Right, "FF0000", 3); AssertBorder(table.GetBorders(1, 2).Left, "FF0000", 3);
        AssertBorder(table.GetBorders(2, 1).Right, "000000", 1);
        AssertBorder(table.GetBorders(3, 1).Right, "0000FF", 2); AssertBorder(table.GetBorders(3, 2).Left, "0000FF", 2);
    }

    [Fact]
    public void Header_column_strokes_are_inactive_in_an_all_header_row_table() {
        using var package = DimensionPackage(IWorkDocumentKind.Keynote, rowCount: 1,
            firstBucketPayload: Message(VarintField(1, 0), BytesField(2, DimensionHeader(0, 22))), secondBucketPayload: Message(VarintField(1, 1)),
            tableStyleReference: ReferenceField(3, 40), modelDeclarations: Message(VarintField(9, 1), VarintField(10, 1)),
            styleRecords: new[] { ArchiveRecord(40, 6003, BytesField(11, Message(
                BytesField(47, BorderStroke(2)), BytesField(49, BorderStroke(1)), BytesField(51, BorderStroke(1, defect: "pattern"))))) });
        var projection = IWorkSourceDocument.Open(package).ReadKeynote();
        Assert.Empty(projection.SourceDeclarationIssues);
        Assert.DoesNotContain(projection.Diagnostics, diagnostic => diagnostic.Code == "IWORK_TABLE_BORDERS_UNSUPPORTED");
        AssertBorder(projection.Slides[0].Tables[0].GetBorders(1, 1).Right, "000000", 1);
        AssertBorder(projection.Slides[0].Tables[0].GetBorders(1, 1).Bottom, "000000", 2);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Inherited_borders_and_later_sparse_shared_edge_edits_survive_each_saved_destination(IWorkDocumentKind kind) {
        using var package = BorderPackage(kind);
        var source = IWorkSourceDocument.Open(package);
        var policy = new IWorkConversionOptions { AllowPartialEditableReconstruction = true };
        using var saved = new MemoryStream();
        IWorkTable table;
        if (kind == IWorkDocumentKind.Pages) {
            using var result = source.ToWordDocumentResult(policy); Assert.False(result.IsVisualFallback);
            table = Assert.Single(result.Projection.Tables); result.Value.Save(saved); saved.Position = 0;
            using var reopened = WordDocument.Load(saved);
            var output = Assert.Single(reopened.Tables);
            Assert.Equal("FF0000", output.Rows[0].Cells[0].Borders.BottomColorHex);
            Assert.Equal(24U, output.Rows[0].Cells[0].Borders.BottomSize);
            Assert.Equal(WordBorderStyle.Nil, output.Rows[0].Cells[1].Borders.BottomStyle);
            Assert.Equal(WordBorderStyle.Nil, output.Rows[1].Cells[1].Borders.TopStyle);
            Assert.Equal("000000", output.Rows[2].Cells[1].Borders.BottomColorHex);
            Assert.Equal(8U, output.Rows[2].Cells[1].Borders.BottomSize);
        } else if (kind == IWorkDocumentKind.Numbers) {
            using var result = source.ToExcelDocumentResult(policy); Assert.False(result.IsVisualFallback);
            table = Assert.Single(Assert.Single(result.Projection.Sheets).Tables); result.Value.Save(saved); saved.Position = 0;
            using var reopened = ExcelDocument.Load(saved); var output = Assert.Single(reopened.Sheets);
            Assert.Equal("thick", output.GetCellStyle(1, 1).Border!.Bottom!.Style);
            Assert.Equal("FFFF0000", output.GetCellStyle(1, 1).Border!.Bottom!.ColorArgb);
            Assert.Equal("none", output.GetCellStyle(1, 2).Border!.Bottom!.Style);
            Assert.Equal("none", output.GetCellStyle(2, 2).Border!.Top!.Style);
            Assert.Equal("thin", output.GetCellStyle(3, 2).Border!.Bottom!.Style);
            Assert.Equal("FF000000", output.GetCellStyle(3, 2).Border!.Bottom!.ColorArgb);
        } else {
            using var result = source.ToPowerPointPresentationResult(policy); Assert.False(result.IsVisualFallback);
            table = Assert.Single(Assert.Single(result.Projection.Slides).Tables); result.Value.Save(saved); saved.Position = 0;
            using var reopened = PowerPointPresentation.Load(saved); Assert.Empty(reopened.ValidateDocument());
            saved.Position = 0; using var zip = new System.IO.Compression.ZipArchive(saved, System.IO.Compression.ZipArchiveMode.Read, true);
            var entry = zip.Entries.Single(item => item.FullName.StartsWith("ppt/slides/", StringComparison.Ordinal) && item.FullName.EndsWith(".xml", StringComparison.Ordinal));
            using var xml = entry.Open(); var document = System.Xml.Linq.XDocument.Load(xml);
            System.Xml.Linq.XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            var rows = document.Descendants(a + "tbl").Single().Elements(a + "tr").ToArray();
            var first = rows[0].Elements(a + "tc").ToArray();
            Assert.Equal("38100", (string?)first[0].Element(a + "tcPr")!.Element(a + "lnB")!.Attribute("w"));
            Assert.Equal("FF0000", (string?)first[0].Element(a + "tcPr")!.Element(a + "lnB")!.Descendants(a + "srgbClr").Single().Attribute("val"));
            Assert.NotNull(first[1].Element(a + "tcPr")!.Element(a + "lnB")!.Element(a + "noFill"));
            Assert.NotNull(rows[1].Elements(a + "tc").Last().Element(a + "tcPr")!.Element(a + "lnT")!.Element(a + "noFill"));
        }
        Assert.Empty(table.Cells); AssertBorder(table.GetBorders(1, 1).Bottom, "FF0000", 3);
        AssertBorder(table.GetBorders(2, 1).Top, "FF0000", 3);
        AssertBorder(table.GetBorders(1, 2).Bottom, "000000", 0);
        AssertBorder(table.GetBorders(2, 2).Top, "000000", 0);
        AssertBorder(table.GetBorders(3, 2).Bottom, "000000", 1);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Selected_inherited_cell_edges_retain_styled_empty_cells_and_shared_explicit_none(IWorkDocumentKind kind) {
        byte[] properties = Message(BytesField(10, BorderStroke(3, 1, 0, 0)), BytesField(12, BorderStroke(1, 0, 0, 1)),
            BytesField(11, BorderStroke(0)), BytesField(13, BorderStroke(0)));
        var styles = new[] { ArchiveRecord(30, 6004, BytesField(1, ReferenceField(3, 31))),
            ArchiveRecord(31, 6004, BytesField(11, properties)) };
        using var package = CellFillPackage(styles, kind: kind); var source = IWorkSourceDocument.Open(package);
        IWorkTable table = kind switch {
            IWorkDocumentKind.Pages => source.ReadPages().Tables[0],
            IWorkDocumentKind.Numbers => source.ReadNumbers().Sheets[0].Tables[0],
            _ => source.ReadKeynote().Slides[0].Tables[0]
        };
        Assert.True(table.HasBorderDeclarations); Assert.Equal(2, table.Cells.Count);
        Assert.Equal(42d, table.GetCell(1, 1)!.Value); Assert.Equal(IWorkCellKind.Empty, table.GetCell(1, 2)!.Kind);
        AssertBorder(table.GetBorders(1, 1).Top, "FF0000", 3); AssertBorder(table.GetBorders(1, 2).Bottom, "0000FF", 1);
        AssertBorder(table.GetBorders(1, 1).Right, "000000", 0); AssertBorder(table.GetBorders(1, 2).Left, "000000", 0);
        AssertBorder(table.GetCell(1, 2)!.Borders!.Value.Top, "FF0000", 3);
    }

    [Theory]
    [InlineData("pattern")]
    [InlineData("transparent")]
    [InlineData("duplicate-width")]
    [InlineData("excessive-pattern-padding")]
    [InlineData("visibility")]
    [InlineData("bad-run-range")]
    [InlineData("equal-order-conflict")]
    public void Unsupported_or_ambiguous_border_declarations_keep_source_dimensions_and_report_loss(string defect) {
        using var package = BorderPackage(IWorkDocumentKind.Numbers, defect: defect);
        var projection = IWorkSourceDocument.Open(package).ReadNumbers(); var table = Assert.Single(Assert.Single(projection.Sheets).Tables);
        Assert.False(projection.HasEditableContent); Assert.Equal(3, table.RowCount); Assert.Equal(2, table.ColumnCount);
        Assert.Contains(projection.Diagnostics, diagnostic => diagnostic.Code == "IWORK_TABLE_BORDERS_UNSUPPORTED"
            && diagnostic.LossKind == OfficeConversionLossKind.Unassessed);
        Assert.NotEmpty(projection.SourceDeclarationIssues);
        if (defect is "bad-run-range" or "equal-order-conflict") Assert.Null(table.GetBorders(1, 2).Bottom);
        else Assert.Null(table.GetBorders(3, defect == "visibility" ? 1 : 2).Right);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages, 13f)]
    [InlineData(IWorkDocumentKind.Keynote, 1584.25f)]
    public void Destination_border_width_bounds_reject_before_constructing_invalid_office_output(IWorkDocumentKind kind, float width) {
        using var package = BorderPackage(kind, width: width, sparse: false); var source = IWorkSourceDocument.Open(package);
        var policy = IWorkTestPolicy.ForIncompletePreview(new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
        if (kind == IWorkDocumentKind.Pages) {
            using var result = source.ToWordDocumentResult(policy); Assert.True(result.IsVisualFallback);
            Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Message.Contains("border widths", StringComparison.Ordinal));
        } else {
            using var result = source.ToPowerPointPresentationResult(policy); Assert.True(result.IsVisualFallback);
            Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Message.Contains("border strokes", StringComparison.Ordinal));
        }
    }

    [Fact]
    public void Competing_selected_cell_styles_do_not_choose_an_arbitrary_shared_edge() {
        using var package = CellFillPackage(new[] { ArchiveRecord(30, 6004, BytesField(11, Message(
            BytesField(11, BorderStroke(3, 1, 0, 0)), BytesField(13, BorderStroke(1, 0, 0, 1))))) });
        var projection = IWorkSourceDocument.Open(package).ReadPages(); var table = Assert.Single(projection.Tables);
        Assert.Equal(42d, table.GetCell(1, 1)!.Value); Assert.False(projection.HasEditableContent);
        Assert.Null(table.GetBorders(1, 1).Right); Assert.Null(table.GetBorders(1, 2).Left);
        Assert.Contains(projection.Diagnostics, diagnostic => diagnostic.Code == "IWORK_TABLE_BORDERS_UNSUPPORTED");
        Assert.Contains(projection.SourceDeclarationIssues, issue => issue.FieldPath == "$cell-border-edges");
    }

    [Fact]
    public void Region_border_defaults_resolve_the_other_neighbor_when_the_first_has_no_declaration() {
        using var package = DimensionPackage(IWorkDocumentKind.Numbers,
            tableStyleReference: ReferenceField(3, 40), modelDeclarations: ReferenceField(18, 50),
            styleRecords: new[] { ArchiveRecord(40, 6003, BytesField(11, BytesField(60, BorderStroke(1)))),
                ArchiveRecord(50, 6004, BytesField(11, BytesField(10, BorderStroke(3, 1, 0, 0)))) });
        var projection = IWorkSourceDocument.Open(package).ReadNumbers(); var table = Assert.Single(Assert.Single(projection.Sheets).Tables);
        Assert.True(projection.HasEditableContent); AssertBorder(table.GetBorders(1, 1).Bottom, "FF0000", 3);
        AssertBorder(table.GetBorders(2, 1).Top, "FF0000", 3);
    }

    [Fact]
    public void Enabled_or_unresolved_grouping_cannot_claim_complete_rectangular_layout_but_disabled_metadata_can() {
        foreach (string state in new[] { "disabled", "enabled", "missing-group" }) {
            using var package = DimensionPackage(IWorkDocumentKind.Numbers, modelDeclarations: ReferenceField(86, 70),
                tableStyleReference: ReferenceField(3, 40), styleRecords: new[] {
                    ArchiveRecord(40, 6003, BytesField(11, BytesField(60, BorderStroke(1)))),
                    ArchiveRecord(70, 6372, ReferenceField(1, state == "missing-group" ? 99UL : 71UL)),
                    ArchiveRecord(71, 6373, VarintField(6, state == "enabled" ? 1UL : 0UL)) });
            var projection = IWorkSourceDocument.Open(package).ReadNumbers(); var table = Assert.Single(Assert.Single(projection.Sheets).Tables);
            Assert.Equal(state == "disabled", projection.HasEditableContent);
            if (state == "disabled") { AssertBorder(table.GetBorders(1, 1).Bottom, "000000", 1); Assert.Empty(projection.SourceDeclarationIssues); }
            else {
                Assert.Null(table.GetBorders(1, 1).Bottom);
                Assert.Contains(projection.Diagnostics, diagnostic => diagnostic.Code == "IWORK_TABLE_GROUPING_UNSUPPORTED"
                    && diagnostic.LossKind == OfficeConversionLossKind.Unassessed);
            }
        }
    }

    [Fact]
    public void Grouping_reference_expansion_retains_the_source_wide_dimension_limit_as_fatal() {
        using var package = DimensionPackage(IWorkDocumentKind.Numbers, modelDeclarations: ReferenceField(86, 70),
            styleRecords: new[] { ArchiveRecord(70, 6372, ReferenceField(1, 71)), ArchiveRecord(71, 6373, VarintField(6, 0)) });
        var error = Assert.Throws<InvalidDataException>(() => IWorkSourceDocument.Open(package,
            new IWorkReadOptions { MaximumTableDimensionEntries = 1 }).ReadNumbers());
        Assert.Contains("dimension entries", error.Message);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Legacy_category_metadata_and_pivot_owners_retain_layout_qualification_limits(bool pivot) {
        byte[] grouping = pivot ? ReferenceField(85, 70) : BytesField(81, BytesField(2, VarintField(6, 1)));
        using var package = DimensionPackage(IWorkDocumentKind.Numbers, modelDeclarations: grouping,
            styleRecords: new[] { ArchiveRecord(70, 6370, Message()) });
        var projection = IWorkSourceDocument.Open(package).ReadNumbers(); Assert.False(projection.HasEditableContent);
        Assert.Contains(projection.Diagnostics, diagnostic => diagnostic.Code == "IWORK_TABLE_GROUPING_UNSUPPORTED");
        Assert.Contains(projection.SourceDeclarationIssues, issue => issue.FieldPath == (pivot ? "85" : "81/2[1]/6"));
    }

    [Fact]
    public void Hidden_inherited_patterned_grid_strokes_do_not_reject_visible_solid_borders() {
        using var package = DimensionPackage(IWorkDocumentKind.Numbers, tableStyleReference: ReferenceField(3, 40),
            styleRecords: new[] { ArchiveRecord(40, 6003, Message(BytesField(1, ReferenceField(3, 41)), BytesField(11, VarintField(33, 0)))),
                ArchiveRecord(41, 6003, BytesField(11, Message(BytesField(59, BorderStroke(1)), BytesField(61, BorderStroke(1, defect: "pattern"))))) });
        var projection = IWorkSourceDocument.Open(package).ReadNumbers(); var table = Assert.Single(Assert.Single(projection.Sheets).Tables);
        Assert.True(projection.HasEditableContent); Assert.Empty(projection.SourceDeclarationIssues);
        AssertBorder(table.GetBorders(3, 1).Right, "000000", 0); AssertBorder(table.GetBorders(3, 2).Right, "000000", 1);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Destination_border_precision_requires_partial_policy_and_records_rounding(IWorkDocumentKind kind) {
        using var package = BorderPackage(kind, width: 0.3f, sparse: false); var source = IWorkSourceDocument.Open(package);
        var strict = IWorkTestPolicy.ForIncompletePreview(new IWorkConversionOptions());
        var partial = new IWorkConversionOptions { AllowPartialEditableReconstruction = true };
        if (kind == IWorkDocumentKind.Pages) {
            using var rejected = source.ToWordDocumentResult(strict); Assert.True(rejected.IsVisualFallback);
            using var converted = source.ToWordDocumentResult(partial); Assert.False(converted.IsVisualFallback);
            Assert.Contains(converted.Report.Diagnostics, diagnostic => diagnostic.Code == "IWORK_PAGES_DOCX_PRECISION");
            using var saved = new MemoryStream(); converted.Value.Save(saved); saved.Position = 0;
            using var reopened = WordDocument.Load(saved); Assert.Equal(2U, reopened.Tables[0].Rows[0].Cells[0].Borders.TopSize);
        } else {
            using var rejected = source.ToPowerPointPresentationResult(strict); Assert.True(rejected.IsVisualFallback);
            using var converted = source.ToPowerPointPresentationResult(partial); Assert.False(converted.IsVisualFallback);
            Assert.Contains(converted.Report.Diagnostics, diagnostic => diagnostic.Code == "IWORK_KEYNOTE_PPTX_PRECISION");
            using var saved = new MemoryStream(); converted.Value.Save(saved); saved.Position = 0;
            using var reopened = PowerPointPresentation.Load(saved); Assert.Empty(reopened.ValidateDocument());
        }
    }

    [Theory]
    [InlineData(1.75f, "medium")]
    [InlineData(0.6f, "hair")]
    public void Numbers_quantized_border_widths_report_approximation_and_preserve_saved_color(float width, string expected) {
        using var package = BorderPackage(IWorkDocumentKind.Numbers, width: width, sparse: false);
        using var result = IWorkSourceDocument.Open(package).ToExcelDocumentResult(new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
        Assert.False(result.IsVisualFallback);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "IWORK_NUMBERS_BORDER_WIDTH_APPROXIMATED"
            && diagnostic.LossKind == OfficeConversionLossKind.Approximation);
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved); var side = reopened.Sheets[0].GetCellStyle(3, 2).Border!.Bottom!;
        Assert.Equal(expected, side.Style); Assert.Equal("FF000000", side.ColorArgb);
    }

    private static byte[] BorderStroke(float width, float red = 0, float green = 0, float blue = 0, string? defect = null) =>
        Message(FillColor(red, green, blue, defect == "transparent" ? 0.5f : 1), FloatField(2, width),
            defect == "duplicate-width" ? FloatField(2, width) : Array.Empty<byte>(),
            BytesField(6, Message(VarintField(1, defect == "pattern" ? 0UL : width == 0 ? 2UL : 1UL),
                defect == "excessive-pattern-padding" ? Message(Enumerable.Range(0, 7).Select(_ => FloatField(4, 0)).ToArray()) : Array.Empty<byte>())));

    private static MemoryStream BorderPackage(IWorkDocumentKind kind, float width = 1, string? defect = null, bool sparse = true) {
        byte[] parent = Message(new[] { 58, 59, 60, 61 }.Select(field => BytesField(field, BorderStroke(width, defect: defect))).ToArray());
        var records = new List<byte[]> {
            ArchiveRecord(40, 6003, Message(BytesField(1, ReferenceField(3, 41)),
                defect == "visibility" ? BytesField(11, VarintField(33, 2)) : Array.Empty<byte>())),
            ArchiveRecord(41, 6003, BytesField(11, parent))
        };
        if (sparse) {
            records.Add(ArchiveRecord(60, 6305, Message(VarintField(1, 2), VarintField(2, 2), VarintField(3, 3),
                ReferenceField(6, 61), ReferenceField(7, 62))));
            // Reversed archive order deliberately differs from edit order on the shared edge.
            records.Add(ArchiveRecord(61, 6306, Message(VarintField(1, 1), BytesField(2, Message(
                VarintField(1, defect == "bad-run-range" ? 2UL : 1UL), VarintField(2, 1), BytesField(3, BorderStroke(0)),
                VarintField(4, defect == "equal-order-conflict" ? 1UL : 2UL))))));
            records.Add(ArchiveRecord(62, 6306, Message(VarintField(1, 0), BytesField(2, Message(
                VarintField(1, 0), VarintField(2, 2), BytesField(3, BorderStroke(3, 1, 0, 0)), VarintField(4, 1))))));
        }
        return DimensionPackage(kind, tableStyleReference: ReferenceField(3, 40), styleRecords: records.ToArray(),
            modelDeclarations: sparse ? ReferenceField(49, 60) : null);
    }
}
