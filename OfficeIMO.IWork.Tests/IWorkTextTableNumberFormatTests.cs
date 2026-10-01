using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    public static IEnumerable<object[]> TextTableNumericCases() {
        foreach (IWorkDocumentKind kind in new[] { IWorkDocumentKind.Pages, IWorkDocumentKind.Keynote }) {
            yield return new object[] { kind, "number", -1234.5d, "(1,234.50)", "FF0000" };
            yield return new object[] { kind, "percentage", 0.125d, "12.50%", "" };
            yield return new object[] { kind, "automatic", 0.125d, "12.5%", "" };
            yield return new object[] { kind, "scientific", -0.0625d, "-6.250E-02", "" };
            yield return new object[] { kind, "fraction", -1.375d, "-1 3/8", "" };
            yield return new object[] { kind, "currency", -1234.5d, "PLN (1,234.50)", "" };
            yield return new object[] { kind, "zero", 0d, "0.00", "" };
        }
    }

    [Theory]
    [MemberData(nameof(TextTableNumericCases))]
    public void Numeric_format_text_and_color_survive_saved_Word_and_PowerPoint(
        IWorkDocumentKind kind, string formatKind, double value, string expectedText, string expectedColor) {
        byte[] format = formatKind switch {
            "number" => NumericFormat(256, 2, 3, 1),
            "percentage" => NumericFormat(258, 2),
            "automatic" => NumericFormat(258, 253),
            "scientific" => NumericFormat(259, 3),
            "fraction" => FractionFormat(8),
            "currency" => CurrencyFormat("PLN", 2, grouping: 1, accounting: 1),
            _ => NumericFormat(256, 2)
        };
        using MemoryStream package = NumberFormatPackage(kind, format, value: value, currency: formatKind == "currency");
        AssertSavedTextTableNumber(package, kind, value, expectedText, expectedColor);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Numeric_formula_cache_is_formatted_without_replacing_source_expression(IWorkDocumentKind kind) {
        byte[] cell = new byte[28]; cell[0] = 5; cell[1] = 10;
        WriteUInt32(cell, 8, (1u << 1) | (1u << 9) | (1u << 13));
        Buffer.BlockCopy(BitConverter.GetBytes(0.125d), 0, cell, 12, 8);
        WriteUInt32(cell, 20, 0); WriteUInt32(cell, 24, 1);
        using MemoryStream package = TableDependencyPackage(kind,
            Message(ReferenceField(22, 13), ReferenceField(6, 14)), cellPayload: cell,
            additionalRecords: Message(
                ArchiveRecord(13, 6005, Message(VarintField(1, 2), BytesField(3, FormatEntry(NumericFormat(258, 2))))),
                ArchiveRecord(14, 6201, Message(BytesField(3, Message(VarintField(1, 0), BytesField(5, FormulaConstant(0.125d))))))));
        IWorkTableCell source = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package, kind), kind).Item1.Cells);
        Assert.Equal(IWorkCellKind.Formula, source.Kind);
        Assert.Equal("=0.125", source.Formula);
        Assert.Equal("0.125", source.CachedDisplayText);
        package.Position = 0;
        AssertSavedTextTableNumber(package, kind, 0.125d, "12.50%", "");
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Percentage_format_overflow_preserves_raw_text_and_reports_omission(IWorkDocumentKind kind) {
        using MemoryStream package = NumberFormatPackage(kind, NumericFormat(258, 2), value: 1e308);
        AssertSavedTextTableNumber(package, kind, 1e308, "1E+308", "", omitted: true);
    }

    private static void AssertSavedTextTableNumber(MemoryStream package, IWorkDocumentKind kind,
        double value, string expectedText, string expectedColor, bool omitted = false) {
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, kind);
        using var saved = new MemoryStream();
        IWorkTableCell cell;
        IWorkConversionReport report;
        if (kind == IWorkDocumentKind.Pages) {
            using var result = source.ToWordDocumentResult(new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
            Assert.False(result.IsVisualFallback);
            cell = Assert.Single(Assert.Single(result.Projection.Tables).Cells); report = result.Report;
            result.Value.Save(saved); saved.Position = 0;
            using var reopened = OfficeIMO.Word.WordDocument.Load(saved);
            var paragraph = Assert.Single(reopened.Tables).Rows[0].Cells[0].Paragraphs[0];
            Assert.Equal(expectedText, paragraph.Text);
            if (expectedColor.Length > 0) Assert.Equal(expectedColor, paragraph.ColorHex);
        } else {
            using var result = source.ToPowerPointPresentationResult(new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
            Assert.False(result.IsVisualFallback);
            cell = Assert.Single(Assert.Single(Assert.Single(result.Projection.Slides).Tables).Cells); report = result.Report;
            result.Value.Save(saved); saved.Position = 0;
            using var reopened = OfficeIMO.PowerPoint.PowerPointPresentation.Load(saved);
            var target = Assert.Single(reopened.Slides[0].Tables).GetCell(0, 0);
            Assert.Equal(expectedText, target.Text);
            if (expectedColor.Length > 0) Assert.Equal(expectedColor, target.Paragraphs[0].Runs[0].Color);
            Assert.Empty(reopened.ValidateDocument());
        }
        Assert.Equal(value, cell.Value);
        Assert.Equal(value.ToString(System.Globalization.CultureInfo.InvariantCulture), cell.CachedDisplayText);
        string prefix = kind == IWorkDocumentKind.Pages ? "IWORK_PAGES" : "IWORK_KEYNOTE";
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code == prefix +
            (omitted ? "_NUMBER_FORMAT_OMITTED" : "_NUMBER_FORMAT_APPROXIMATED")
            && diagnostic.LossKind == (omitted ? global::OfficeIMO.OfficeConversionLossKind.Omission
                : global::OfficeIMO.OfficeConversionLossKind.Approximation));
        Assert.DoesNotContain(report.Diagnostics, diagnostic => diagnostic.Code == prefix +
            (omitted ? "_NUMBER_FORMAT_APPROXIMATED" : "_NUMBER_FORMAT_OMITTED"));
    }
}
