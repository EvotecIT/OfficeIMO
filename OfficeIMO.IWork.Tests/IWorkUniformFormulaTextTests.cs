using OfficeIMO.Excel;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Uniform_formula_text_styles_preserve_formula_cache_and_font_in_saved_XLSX(bool splitRuns) {
        byte[] attributes = splitRuns
            ? AttributeTable(8, AttributeEntry(0, ReferenceField(2, 30)), AttributeEntry(3, ReferenceField(2, 31)))
            : AttributeTable(8, AttributeEntry(0, ReferenceField(2, 30)));
        byte[] character = Message(VarintField(1, 1), VarintField(2, 1), FloatField(3, 18),
            StringField(5, "Courier New"), VarintField(11, 1), VarintField(12, 1));
        using var package = CreateNumbersWithPartialRichCell(true, richText: "Styled",
            richAttributes: attributes, additionalRecords: Message(
                ArchiveRecord(30, 2021, BytesField(11, character)),
                ArchiveRecord(31, 2021, BytesField(11, character))));
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package,
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly });
        result.Report.RequireCompleteEditableReconstruction();
        Assert.DoesNotContain(result.Report.Diagnostics, d => d.Code == "IWORK_NUMBERS_EXCEL_RICH_TEXT_PARTIAL");
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        ExcelSheet sheet = reopened.Sheets[0];
        Assert.Equal("1", sheet.GetFormulaText(1, 1));
        Assert.Equal("Styled", sheet.CellAt(1, 1).GetValue<string>());
        var style = sheet.GetCellStyle(1, 1);
        Assert.True(style.Bold); Assert.True(style.Italic); Assert.True(style.Underline); Assert.True(style.Strikethrough);
        Assert.Equal(18, style.FontSize); Assert.Equal("Courier New", style.FontName);
    }

    [Fact]
    public void Mixed_formula_run_styles_require_explicit_partial_policy_without_dropping_the_formula_cache() {
        using var package = CreateNumbersWithPartialRichCell(true, richText: "Styled",
            richAttributes: AttributeTable(8, AttributeEntry(0, ReferenceField(2, 30)),
                AttributeEntry(3, ReferenceField(2, 31))), additionalRecords: Message(
                ArchiveRecord(30, 2021, BytesField(11, VarintField(1, 1))),
                ArchiveRecord(31, 2021, BytesField(11, VarintField(1, 0)))));
        Assert.Throws<InvalidDataException>(() => ExcelIWorkConverter.ConvertNumbersToExcelResult(package,
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly }));
        package.Position = 0;
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package,
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = true });
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "IWORK_NUMBERS_EXCEL_RICH_TEXT_PARTIAL"
            && d.LossKind == OfficeConversionLossKind.Omission);
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        Assert.Equal("1", reopened.Sheets[0].GetFormulaText(1, 1));
        Assert.Equal("Styled", reopened.Sheets[0].CellAt(1, 1).GetValue<string>());
    }

    [Fact]
    public void Uniform_formula_font_outside_XLSX_bounds_is_rejected_before_destination_mutation() {
        using var package = CreateNumbersWithPartialRichCell(true, richText: "Styled",
            richAttributes: AttributeTable(8, AttributeEntry(0, ReferenceField(2, 30))),
            additionalRecords: ArchiveRecord(30, 2021, BytesField(11, FloatField(3, 410))));
        Assert.Throws<InvalidDataException>(() => ExcelIWorkConverter.ConvertNumbersToExcelResult(package,
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = true }));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Formula_result_links_preserve_expression_and_cache_and_report_mixed_targets(bool mixedTargets) {
        const string first = "https://example.test/first";
        using var package = CreateNumbersWithPartialRichCell(true, richText: "Styled",
            richAttributes: AttributeTable(11, AttributeEntry(0, ReferenceField(2, 30)),
                AttributeEntry(3, ReferenceField(2, mixedTargets ? 31ul : 30ul))), additionalRecords: Message(
                ArchiveRecord(30, 2032, StringField(2, first)),
                ArchiveRecord(31, 2032, StringField(2, "https://example.test/second"))));
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package,
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = mixedTargets });
        Assert.Equal(mixedTargets, result.Report.Diagnostics.Any(d => d.Code == "IWORK_NUMBERS_EXCEL_RICH_TEXT_PARTIAL"));
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        ExcelSheet sheet = reopened.Sheets[0];
        Assert.Equal("1", sheet.GetFormulaText(1, 1));
        Assert.Equal("Styled", sheet.CellAt(1, 1).GetValue<string>());
        if (mixedTargets) Assert.Empty(sheet.GetHyperlinks());
        else Assert.Equal(first, Assert.Single(sheet.GetHyperlinks()).Value.Target);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Incomplete_formula_cache_text_requires_partial_acceptance_even_with_complete_fonts(bool unexpectedWire) {
        using var package = CreateNumbersWithPartialRichCell(true, richStorage: Message(
            StringField(3, "Readable"), unexpectedWire ? VarintField(3, 1) : BytesField(3, new byte[] { 0xff })));
        IWorkTableCell sourceCell = Assert.Single(Assert.Single(Assert.Single(
            IWorkSourceDocument.Open(package).ReadNumbers().Sheets).Tables).Cells);
        Assert.False(sourceCell.RichText!.IsTextComplete);
        Assert.True(sourceCell.RichText.IsFormattingComplete);
        package.Position = 0;
        Assert.Throws<InvalidDataException>(() => ExcelIWorkConverter.ConvertNumbersToExcelResult(package,
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly }));
        package.Position = 0;
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package,
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = true });
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "IWORK_TABLE_FORMULA_CACHE_PARTIAL");
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "IWORK_NUMBERS_EXCEL_RICH_TEXT_PARTIAL");
        Assert.Equal("1", result.Value.Sheets[0].GetFormulaText(1, 1));
        Assert.Null(result.Value.Sheets[0].CellAt(1, 1).GetValue().Value);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Formula_natural_alignment_is_uniform_only_when_paragraph_directions_agree(bool mixedDirections) {
        string text = mixedDirections ? "Latin\nשלום" : "Latin\nEnglish";
        using var package = CreateNumbersWithPartialRichCell(true, richText: text,
            richAttributes: AttributeTable(5, AttributeEntry(0, ReferenceField(2, 19))),
            additionalRecords: ArchiveRecord(19, 2022, BytesField(12, VarintField(1, 4))));
        if (mixedDirections) {
            Assert.Throws<InvalidDataException>(() => ExcelIWorkConverter.ConvertNumbersToExcelResult(package,
                conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly }));
            package.Position = 0;
        }
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package,
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = mixedDirections });
        Assert.Equal(mixedDirections, result.Report.Diagnostics.Any(d => d.Code == "IWORK_NUMBERS_EXCEL_RICH_TEXT_PARTIAL"));
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        Assert.Equal("1", reopened.Sheets[0].GetFormulaText(1, 1));
        Assert.Equal(text, reopened.Sheets[0].CellAt(1, 1).GetValue<string>());
        if (!mixedDirections) Assert.Equal("general", reopened.Sheets[0].GetCellStyle(1, 1).HorizontalAlignment);
    }
}
