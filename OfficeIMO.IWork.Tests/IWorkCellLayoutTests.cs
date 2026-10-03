using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData("inherit", 4, 4, 4, 4, IWorkCellVerticalAlignment.Middle)]
    [InlineData("clear", 0, 0, 0, 0, IWorkCellVerticalAlignment.Top)]
    [InlineData("partial", 2, 0, 0, 0, IWorkCellVerticalAlignment.Bottom)]
    public void Selected_cell_layout_inheritance_and_whole_padding_override_retain_empty_cells(
        string mode, double left, double top, double right, double bottom, IWorkCellVerticalAlignment vertical) {
        byte[] child = mode switch {
            "clear" => Message(VarintField(8, 0), BytesField(9, Message())),
            "partial" => Message(VarintField(8, 2), BytesField(9, FloatField(1, 2))),
            _ => Message()
        };
        using var package = CellFillPackage(new[] { CellLayoutStyle(30, child, 31),
            CellLayoutStyle(31, Message(VarintField(8, 1), BytesField(9, CellPadding(4)))) });
        var projection = IWorkSourceDocument.Open(package).ReadPages();
        IWorkTable table = Assert.Single(projection.Tables);
        Assert.Equal(2, table.Cells.Count);
        Assert.All(table.Cells, cell => {
            Assert.Null(cell.Fill);
            Assert.Equal(vertical, cell.VerticalAlignment);
            Assert.Equal((left, top, right, bottom), (cell.Padding!.LeftPoints, cell.Padding.TopPoints,
                cell.Padding.RightPoints, cell.Padding.BottomPoints));
        });
        Assert.Equal(IWorkCellKind.Empty, table.GetCell(1, 2)!.Kind);
        Assert.DoesNotContain(projection.Diagnostics, diagnostic => diagnostic.Code == "IWORK_TABLE_CELL_LAYOUT_UNSUPPORTED");
        package.Position = 0;
        Assert.Throws<InvalidDataException>(() => IWorkSourceDocument.Open(package,
            new IWorkReadOptions { MaximumMaterializedCells = 1 }).ReadPages());
    }

    [Theory]
    [InlineData("duplicate-alignment", "11/8")]
    [InlineData("wrong-alignment-wire", "11/8")]
    [InlineData("unsupported-alignment", "11/8")]
    [InlineData("duplicate-padding", "11/9")]
    [InlineData("wrong-padding-wire", "11/9")]
    [InlineData("duplicate-side", "11/9")]
    [InlineData("wrong-side-wire", "11/9")]
    [InlineData("unknown-side", "11/9")]
    [InlineData("negative", "11/9")]
    [InlineData("nonfinite", "11/9")]
    public void Invalid_selected_cell_layout_keeps_values_and_valid_fill_with_source_evidence(string defect, string path) {
        byte[] properties = defect switch {
            "duplicate-alignment" => Message(VarintField(8, 1), VarintField(8, 2)),
            "wrong-alignment-wire" => FloatField(8, 1),
            "unsupported-alignment" => VarintField(8, 3),
            "duplicate-padding" => Message(BytesField(9, CellPadding(4)), BytesField(9, CellPadding(0))),
            "wrong-padding-wire" => VarintField(9, 1),
            "duplicate-side" => BytesField(9, Message(FloatField(1, 4), FloatField(1, 2))),
            "wrong-side-wire" => BytesField(9, VarintField(1, 4)),
            "unknown-side" => BytesField(9, FloatField(5, 4)),
            "negative" => BytesField(9, CellPadding(-1)),
            _ => BytesField(9, CellPadding(float.PositiveInfinity))
        };
        using var package = CellFillPackage(new[] { CellLayoutStyle(30, Message(BytesField(1, FillColor(1, 0, 0)), properties)) });
        using var result = IWorkSourceDocument.Open(package).ToWordDocumentResult();
        Assert.True(result.IsVisualFallback);
        var cell = Assert.Single(result.Projection.Tables).GetCell(1, 1)!;
        Assert.Equal(42d, cell.Value);
        Assert.Equal("FF0000", cell.Fill!.Color!.RgbHex);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "IWORK_TABLE_CELL_LAYOUT_UNSUPPORTED"
            && diagnostic.LossKind == global::OfficeIMO.OfficeConversionLossKind.Unassessed);
        Assert.DoesNotContain(result.Report.Diagnostics, diagnostic => diagnostic.Code == "IWORK_TABLE_CELL_FILL_UNSUPPORTED");
        Assert.Contains(result.Report.SourceDeclarationIssues, issue => issue.Owner.RecordIdentifier == 30 && issue.FieldPath == path);
    }

    [Fact]
    public void Unsupported_fill_does_not_discard_qualified_cell_layout() {
        using var package = CellFillPackage(new[] { CellLayoutStyle(30, Message(BytesField(1, FillColor(float.NaN, 0, 0)),
            VarintField(8, 2), BytesField(9, CellPadding(4)))) });
        var projection = IWorkSourceDocument.Open(package).ReadPages();
        var cell = Assert.Single(projection.Tables).GetCell(1, 1)!;
        Assert.Null(cell.Fill);
        Assert.Equal(4, cell.Padding!.LeftPoints);
        Assert.Equal(IWorkCellVerticalAlignment.Bottom, cell.VerticalAlignment);
        Assert.Contains(projection.Diagnostics, diagnostic => diagnostic.Code == "IWORK_TABLE_CELL_FILL_UNSUPPORTED");
        Assert.DoesNotContain(projection.Diagnostics, diagnostic => diagnostic.Code == "IWORK_TABLE_CELL_LAYOUT_UNSUPPORTED");
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Cell_layout_survives_supported_saved_destinations_and_reports_XLSX_padding_loss(IWorkDocumentKind kind) {
        using var package = CellFillPackage(new[] { CellLayoutStyle(30, Message(VarintField(8, 1), BytesField(9, CellPadding(4)))) }, kind: kind);
        var source = IWorkSourceDocument.Open(package);
        var policy = new IWorkConversionOptions { AllowPartialEditableReconstruction = true };
        using var saved = new MemoryStream();
        if (kind == IWorkDocumentKind.Pages) {
            using var result = source.ToWordDocumentResult(policy);
            Assert.False(result.IsVisualFallback);
            result.Value.Save(saved); saved.Position = 0;
            using var reopened = OfficeIMO.Word.WordDocument.Load(saved);
            Assert.All(reopened.Tables[0].Rows[0].Cells, cell => {
                Assert.Equal((short)80, cell.MarginLeftWidth);
                Assert.Equal((short)80, cell.MarginTopWidth);
                Assert.Equal((short)80, cell.MarginRightWidth);
                Assert.Equal((short)80, cell.MarginBottomWidth);
                Assert.Equal(OfficeIMO.Word.WordTableVerticalAlignment.Center, cell.VerticalAlignment);
            });
        } else if (kind == IWorkDocumentKind.Numbers) {
            using var result = source.ToExcelDocumentResult(policy);
            Assert.False(result.IsVisualFallback);
            result.Value.Save(saved); saved.Position = 0;
            using var reopened = OfficeIMO.Excel.ExcelDocument.Load(saved);
            Assert.Equal("center", reopened.Sheets[0].GetCellStyle(1, 1).VerticalAlignment);
            Assert.Equal("center", reopened.Sheets[0].GetCellStyle(1, 2).VerticalAlignment);
            Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "IWORK_NUMBERS_CELL_PADDING_OMITTED"
                && diagnostic.LossKind == global::OfficeIMO.OfficeConversionLossKind.Omission);
            Assert.True(result.Report.IsPartialEditableReconstruction);
            Assert.Throws<InvalidOperationException>(() => result.Report.RequireCompleteEditableReconstruction());
        } else {
            using var result = source.ToPowerPointPresentationResult(policy);
            Assert.False(result.IsVisualFallback);
            result.Value.Save(saved); saved.Position = 0;
            using var reopened = OfficeIMO.PowerPoint.PowerPointPresentation.Load(saved);
            var table = Assert.Single(reopened.Slides[0].Tables);
            foreach (int column in new[] { 0, 1 }) {
                var cell = table.GetCell(0, column);
                Assert.Equal(4, cell.PaddingLeftPoints);
                Assert.Equal(4, cell.PaddingTopPoints);
                Assert.Equal(4, cell.PaddingRightPoints);
                Assert.Equal(4, cell.PaddingBottomPoints);
                Assert.Equal(OfficeIMO.PowerPoint.PowerPointTextVerticalAlignment.Center, cell.VerticalAlignment);
            }
        }
        package.Position = 0;
        var builder = new OfficeIMO.Reader.OfficeDocumentReaderBuilder();
        OfficeIMO.Reader.IWork.OfficeDocumentReaderBuilderIWorkExtensions.AddIWorkHandler(builder);
        var read = builder.Build().ReadDocument(package, "layout." + (kind == IWorkDocumentKind.Pages ? "pages" : kind == IWorkDocumentKind.Numbers ? "numbers" : "key"));
        Assert.Equal(new[] { "42", string.Empty }, Assert.Single(Assert.Single(read.Tables).Rows));
        Assert.Contains(read.Diagnostics, diagnostic => diagnostic.Code == "IWORK_READER_TABLE_STYLE_PARTIAL");
    }

    [Theory]
    [InlineData(IWorkConversionMode.Auto)]
    [InlineData(IWorkConversionMode.EditableOnly)]
    public void Numbers_destination_padding_omission_requires_explicit_partial_policy(IWorkConversionMode mode) {
        using var package = CellFillPackage(new[] { CellLayoutStyle(30, BytesField(9, CellPadding(4))) },
            kind: IWorkDocumentKind.Numbers);
        var source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers);
        Assert.True(source.ReadNumbers().HasEditableContent);
        var policy = new IWorkConversionOptions { Mode = mode };
        if (mode == IWorkConversionMode.EditableOnly) {
            Assert.Contains("padding", Assert.Throws<InvalidDataException>(() => source.ToExcelDocumentResult(policy)).Message,
                StringComparison.OrdinalIgnoreCase);
        } else {
            using var result = source.ToExcelDocumentResult(policy);
            Assert.True(result.IsVisualFallback);
            Assert.Throws<InvalidOperationException>(() => result.Report.RequireCompleteEditableReconstruction());
        }
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages, 2000)]
    [InlineData(IWorkDocumentKind.Keynote, 200000)]
    public void Destination_cell_padding_overflow_uses_fallback_without_losing_source_points(IWorkDocumentKind kind, float padding) {
        using var package = CellFillPackage(new[] { CellLayoutStyle(30, BytesField(9, CellPadding(padding))) }, kind: kind);
        var source = IWorkSourceDocument.Open(package);
        var policy = new IWorkConversionOptions { AllowPartialEditableReconstruction = true };
        if (kind == IWorkDocumentKind.Pages) {
            using var result = source.ToWordDocumentResult(policy);
            Assert.True(result.IsVisualFallback);
            Assert.Equal(padding, result.Projection.Tables[0].GetCell(1, 1)!.Padding!.LeftPoints);
        } else {
            using var result = source.ToPowerPointPresentationResult(policy);
            Assert.True(result.IsVisualFallback);
            Assert.Equal(padding, result.Projection.Slides[0].Tables[0].GetCell(1, 1)!.Padding!.LeftPoints);
        }
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages, false)]
    [InlineData(IWorkDocumentKind.Pages, true)]
    [InlineData(IWorkDocumentKind.Keynote, false)]
    [InlineData(IWorkDocumentKind.Keynote, true)]
    public void Cell_padding_rounding_requires_partial_policy_and_keeps_exact_source_points(IWorkDocumentKind kind, bool partial) {
        const float native = .125f;
        using var package = CellFillPackage(new[] { CellLayoutStyle(30, BytesField(9, CellPadding(native))) }, kind: kind);
        var source = IWorkSourceDocument.Open(package);
        var policy = new IWorkConversionOptions { AllowPartialEditableReconstruction = partial };
        if (kind == IWorkDocumentKind.Pages) {
            using var result = source.ToWordDocumentResult(policy);
            Assert.Equal(!partial, result.IsVisualFallback);
            Assert.Equal(native, result.Projection.Tables[0].GetCell(1, 1)!.Padding!.LeftPoints);
            if (partial) {
                Assert.Equal((short)3, result.Value.Tables[0].Rows[0].Cells[0].MarginLeftWidth);
                Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "IWORK_PAGES_DOCX_PRECISION");
            }
        } else {
            using var result = source.ToPowerPointPresentationResult(policy);
            Assert.Equal(!partial, result.IsVisualFallback);
            Assert.Equal(native, result.Projection.Slides[0].Tables[0].GetCell(1, 1)!.Padding!.LeftPoints);
            if (partial) {
                Assert.Equal(1588d / 12700d, Assert.Single(result.Value.Slides[0].Tables).GetCell(0, 0).PaddingLeftPoints);
                Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "IWORK_KEYNOTE_PPTX_PRECISION");
            }
        }
    }

    private static byte[] CellPadding(float points) => Message(FloatField(1, points), FloatField(2, points), FloatField(3, points), FloatField(4, points));
    private static byte[] CellLayoutStyle(ulong id, byte[] properties, ulong? parent = null) =>
        ArchiveRecord(id, 6004, Message(parent.HasValue ? BytesField(1, ReferenceField(3, parent.Value)) : Array.Empty<byte>(), BytesField(11, properties)));
}
