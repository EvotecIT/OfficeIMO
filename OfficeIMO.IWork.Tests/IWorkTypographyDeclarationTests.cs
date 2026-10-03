using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Selected_unmapped_typography_is_not_reported_as_complete_formatting(IWorkDocumentKind kind) {
        using var package = SelectedRichTextPackage(kind,
            AttributeTable(8, AttributeEntry(0, ReferenceField(2, 19))),
            additionalRecords: ArchiveRecord(19, 2021, BytesField(11, Message(
                VarintField(10, 2), VarintField(13, 2), FloatField(14, 3), FloatField(15, -1), FloatField(27, 0.5f)))));
        var (table, references) = ReadSelectedRichTable(IWorkSourceDocument.Open(package, kind), kind);
        IWorkTextContent text = Assert.Single(table.Cells).RichText!;
        Assert.Equal("Value", text.PlainText);
        Assert.False(text.IsFormattingComplete);
        Assert.Empty(references);
        package.Position = 0;
        if (kind == IWorkDocumentKind.Pages) {
            using var fallback = WordIWorkConverter.ConvertPagesToWordResult(package, conversionOptions: new IWorkConversionOptions { RequireCompleteVisualCoverage = false });
            Assert.True(fallback.IsVisualFallback);
        } else if (kind == IWorkDocumentKind.Numbers) {
            using var fallback = ExcelIWorkConverter.ConvertNumbersToExcelResult(package, conversionOptions: new IWorkConversionOptions { RequireCompleteVisualCoverage = false });
            Assert.True(fallback.IsVisualFallback);
        } else {
            using var fallback = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package, conversionOptions: new IWorkConversionOptions { RequireCompleteVisualCoverage = false });
            Assert.True(fallback.IsVisualFallback);
        }
        package.Position = 0;
        Assert.True(ConvertUnitReport(package, kind, visual: false).IsPartialEditableReconstruction);
        package.Position = 0;
        IWorkConversionReport report = ConvertUnitReport(package, kind, visual: true,
            readOptions: new IWorkReadOptions { PreserveSourceRecords = false });
        Assert.Equal(new[] { "11/10", "11/13", "11/14", "11/15", "11/27" },
            report.SourceDeclarationIssues.Select(issue => issue.FieldPath).OrderBy(path => path));
        Assert.All(report.SourceDeclarationIssues, issue => {
            Assert.Equal(19ul, issue.Owner.RecordIdentifier);
            Assert.Equal(1, issue.DeclaredValueCount);
            Assert.Equal(IWorkSourceDeclarationIssueKind.UnsupportedField, issue.Kind);
        });
        Assert.Empty(report.PreservedRecords);
        Assert.Contains(report.FidelityDiagnostics, diagnostic => diagnostic.Code == "IWORK_SOURCE_DECLARATIONS_UNASSESSED"
            && diagnostic.LossKind == OfficeConversionLossKind.Unassessed);
    }

    [Fact]
    public void Neutral_typography_values_remain_complete_and_unselected_typography_is_not_traversed() {
        using var package = SelectedRichTextPackage(IWorkDocumentKind.Numbers,
            AttributeTable(8, AttributeEntry(0, ReferenceField(2, 19))),
            additionalRecords: Message(
                ArchiveRecord(19, 2021, BytesField(11, Message(VarintField(10, 0), VarintField(13, 0),
                    FloatField(14, 0), FloatField(15, 0), FloatField(27, 0)))),
                ArchiveRecord(20, 2021, BytesField(11, VarintField(10, 1)))));
        var (table, _) = ReadSelectedRichTable(IWorkSourceDocument.Open(package), IWorkDocumentKind.Numbers);
        Assert.True(Assert.Single(table.Cells).RichText!.IsFormattingComplete);
        package.Position = 0;
        Assert.Empty(ConvertUnitReport(package, IWorkDocumentKind.Numbers).SourceDeclarationIssues);
    }

    [Fact]
    public void Invalid_typography_values_are_distinguished_from_valid_unmapped_properties() {
        using var package = SelectedRichTextPackage(IWorkDocumentKind.Numbers,
            AttributeTable(8, AttributeEntry(0, ReferenceField(2, 19))),
            additionalRecords: ArchiveRecord(19, 2021, BytesField(11, Message(
                VarintField(10, 3), VarintField(13, 0), VarintField(13, 1),
                FloatField(14, float.NaN), VarintField(15, 0), FloatField(27, float.PositiveInfinity)))));
        IWorkConversionReport report = ConvertUnitReport(package, IWorkDocumentKind.Numbers);
        Assert.Equal(5, report.SourceDeclarationIssues.Count);
        Assert.All(report.SourceDeclarationIssues, issue => Assert.Equal(IWorkSourceDeclarationIssueKind.InvalidValue, issue.Kind));
        Assert.Equal(2, Assert.Single(report.SourceDeclarationIssues, issue => issue.FieldPath == "11/13").DeclaredValueCount);
        package.Position = 0;
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, new IWorkReadOptions { MaximumSourceDeclarationIssues = 1 });
        Assert.Contains("declaration issues", Assert.Throws<InvalidDataException>(() => source.ReadNumbers()).Message);
    }
}
