using OfficeIMO.Excel;
using OfficeIMO.Excel.IWork;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Partial_rich_formula_cache_remains_editable_with_a_cached_value() {
        using MemoryStream package = CreateNumbersWithPartialRichCell(hasFormula: true);
        Assert.Throws<InvalidDataException>(() => ExcelIWorkConverter.ConvertNumbersToExcelResult(package,
            conversionOptions: new IWorkConversionOptions { Mode = IWorkConversionMode.EditableOnly }));
        package.Position = 0;

        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package,
            conversionOptions: new IWorkConversionOptions {
                Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = true
            });
        IWorkTableCell cell = Assert.Single(Assert.Single(Assert.Single(
            result.Projection.Sheets).Tables).Cells);

        Assert.False(result.IsVisualFallback);
        Assert.True(result.Report.IsPartialEditableReconstruction);
        Assert.Throws<InvalidOperationException>(() => result.Report.RequireCompleteEditableReconstruction());
        IWorkFormulaCellStatus assessment = Assert.Single(result.Report.FormulaCells);
        Assert.True(assessment.ExpressionIsComplete);
        Assert.Equal(IWorkFormulaCacheStatus.Partial, assessment.CacheStatus);
        Assert.Equal(1, result.Report.FormulaSummary.PartialCacheCount);
        Assert.Equal(IWorkCellKind.Formula, cell.Kind);
        Assert.Equal("Beforeafter", cell.Value);
        Assert.True(cell.FormulaIsComplete);
        Assert.False(cell.CachedValueIsComplete);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_TABLE_FORMULA_CACHE_PARTIAL");
        Assert.DoesNotContain(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_TABLE_FORMULA_PARTIAL");
        Assert.DoesNotContain(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_TABLE_RICH_TEXT_STORAGE_UNSUPPORTED");

        ExcelCellData exported = result.Value.Sheets[0].CellAt(1, 1).GetValue();
        Assert.Equal(ExcelCellDataKind.Formula, exported.Kind);
        Assert.Equal("1", result.Value.Sheets[0].GetFormulaText(1, 1));
        Assert.Null(exported.Value);
        using var saved = new MemoryStream();
        result.Value.Save(saved);
        saved.Position = 0;
        using ExcelDocument reopened = ExcelDocument.Load(saved);
        ExcelCellData persisted = reopened.Sheets[0].CellAt(1, 1).GetValue();
        Assert.Equal(ExcelCellDataKind.Formula, persisted.Kind);
        Assert.Equal("1", reopened.Sheets[0].GetFormulaText(1, 1));
        Assert.Null(persisted.Value);
    }

    [Fact]
    public void Partial_rich_formula_cache_without_an_expression_uses_visual_fallback() {
        using MemoryStream package = CreateNumbersWithPartialRichCell(hasFormula: true,
            includeFormulaRecord: false);

        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);
        IWorkTableCell cell = Assert.Single(Assert.Single(Assert.Single(
            result.Projection.Sheets).Tables).Cells);

        Assert.True(result.IsVisualFallback);
        IWorkFormulaCellStatus assessment = Assert.Single(result.Report.FormulaCells);
        Assert.False(assessment.ExpressionIsComplete);
        Assert.Equal(IWorkFormulaCacheStatus.Partial, assessment.CacheStatus);
        Assert.False(cell.FormulaIsComplete);
        Assert.False(cell.CachedValueIsComplete);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_TABLE_FORMULA_CACHE_PARTIAL");
    }

    [Fact]
    public void Partial_rich_nonformula_cell_still_requires_visual_fallback() {
        using MemoryStream package = CreateNumbersWithPartialRichCell(hasFormula: false);

        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);
        IWorkTableCell cell = Assert.Single(Assert.Single(Assert.Single(
            result.Projection.Sheets).Tables).Cells);

        Assert.True(result.IsVisualFallback);
        Assert.Equal("Beforeafter", cell.Value);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_TABLE_RICH_TEXT_STORAGE_UNSUPPORTED");
    }

    [Fact]
    public void Unused_partial_rich_catalog_entry_does_not_block_editable_cells() {
        using MemoryStream package = CreateNumbersWithPartialRichCell(hasFormula: true,
            includeUnusedRichEntry: true);

        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package,
            conversionOptions: new IWorkConversionOptions {
                Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = true
            });

        Assert.False(result.IsVisualFallback);
        Assert.True(result.Projection.HasEditableContent);
        Assert.DoesNotContain(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_TABLE_RICH_TEXT_STORAGE_UNSUPPORTED");
    }

    [Fact]
    public void Visual_only_numbers_conversion_does_not_report_excel_rich_text_loss() {
        using MemoryStream package = CreateNumbersWithPartialRichCell(hasFormula: true);

        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package,
            conversionOptions: new IWorkConversionOptions { Mode = IWorkConversionMode.VisualOnly });

        Assert.True(result.IsVisualFallback);
        Assert.DoesNotContain(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_NUMBERS_EXCEL_RICH_TEXT_PARTIAL");
    }

    [Fact]
    public void Automatic_visual_fallback_does_not_report_unexported_rich_text_loss() {
        using MemoryStream package = CreateNumbersWithPartialRichCell(hasFormula: false);

        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);

        Assert.True(result.IsVisualFallback);
        Assert.DoesNotContain(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_NUMBERS_EXCEL_RICH_TEXT_PARTIAL");
    }

    [Theory]
    [InlineData(IWorkTextAlignment.Right)]
    [InlineData(IWorkTextAlignment.Natural)]
    public void Excel_rich_text_loss_detection_includes_paragraph_alignment(
        IWorkTextAlignment alignment) {
        var textStyle = new IWorkTextStyle(null, null, null, null, null, null,
            null, null, null);
        var paragraphStyle = new IWorkParagraphStyle(null, alignment, null, null,
            null, null, null, null, null, null, textStyle);
        var content = new IWorkTextContent(new[] {
            new IWorkTextParagraph(new[] { new IWorkTextRun("Aligned", textStyle, null) },
                paragraphStyle, null, -1, null, IWorkParagraphBreakKind.None)
        }, isComplete: true, isTextComplete: true);

        Assert.True(ExcelIWorkConverter.HasUnsupportedRichText(content, formula: false));

    }

    [Fact]
    public void Unused_invalid_rich_catalog_entry_does_not_block_editable_cells() {
        using MemoryStream package = CreateNumbersWithPartialRichCell(hasFormula: true,
            includeUnusedRichEntry: true, invalidUnusedWrapper: true);

        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package,
            conversionOptions: new IWorkConversionOptions {
                Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = true
            });

        Assert.False(result.IsVisualFallback);
        Assert.True(result.Projection.HasEditableContent);
        Assert.DoesNotContain(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_TABLE_RICH_TEXT_STORAGE_UNSUPPORTED");
    }

    private static MemoryStream CreateNumbersWithPartialRichCell(bool hasFormula,
        bool includeFormulaRecord = true, bool includeUnusedRichEntry = false,
        bool invalidUnusedWrapper = false) {
        byte[] cell = new byte[hasFormula ? 20 : 16];
        cell[0] = 5;
        cell[1] = 9;
        WriteUInt32(cell, 8, (1u << 4) | (hasFormula ? 1u << 9 : 0u));
        WriteUInt32(cell, 12, 1);
        byte[] row = Message(VarintField(1, 0), BytesField(6, cell),
            BytesField(7, new byte[] { 0, 0 }));
        byte[] tileStorage = Message(BytesField(1,
            Message(VarintField(1, 0), ReferenceField(2, 12))));
        byte[] store = Message(BytesField(3, tileStorage), ReferenceField(17, 13),
            hasFormula ? ReferenceField(6, 16) : Array.Empty<byte>());
        byte[] records = Message(
            ArchiveRecord(1, 1, Message(ReferenceField(1, 2)), new ulong[] { 2 }),
            ArchiveRecord(2, 2, Message(StringField(1, "Sheet"), ReferenceField(2, 10)),
                new ulong[] { 10 }),
            ArchiveRecord(10, 6000, Message(ReferenceField(2, 11)), new ulong[] { 11 }),
            ArchiveRecord(11, 6001, Message(BytesField(4, store), VarintField(6, 1),
                VarintField(7, 1), StringField(8, "Cached")),
                hasFormula ? new ulong[] { 12, 13, 16 } : new ulong[] { 12, 13 }),
            ArchiveRecord(12, 6002, Message(BytesField(5, row))),
            ArchiveRecord(13, 6005, Message(VarintField(1, 8), BytesField(3,
                Message(VarintField(1, 1), ReferenceField(9, 14))),
                includeUnusedRichEntry ? BytesField(3,
                    Message(VarintField(1, 2), ReferenceField(9, 17))) : Array.Empty<byte>()),
                includeUnusedRichEntry ? new ulong[] { 14, 17 } : new ulong[] { 14 }),
            ArchiveRecord(14, 6218, Message(ReferenceField(1, 15)), new ulong[] { 15 }),
            ArchiveRecord(15, 2001, Message(StringField(3, "Before\uFFFCafter"))),
            hasFormula && includeFormulaRecord ? ArchiveRecord(16, 6201, Message(VarintField(1, 3), BytesField(3,
                Message(VarintField(1, 0), BytesField(5, FormulaConstant(1d))))))
                : Array.Empty<byte>(),
            includeUnusedRichEntry ? invalidUnusedWrapper
                ? ArchiveRecord(17, 2001, Message(StringField(3, "Not a wrapper")))
                : ArchiveRecord(17, 6218, Message(ReferenceField(1, 18)), new ulong[] { 18 })
                : Array.Empty<byte>(),
            includeUnusedRichEntry && !invalidUnusedWrapper ? ArchiveRecord(18, 2001,
                Message(StringField(3, "Unused\uFFFCtext"))) : Array.Empty<byte>());
        return CreatePackage(("Index/Document.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));
    }
}
