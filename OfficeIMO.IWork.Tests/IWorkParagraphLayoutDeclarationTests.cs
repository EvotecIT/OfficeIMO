using System.Text.Json;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Selected_line_spacing_and_tabs_require_partial_formatting_and_retain_physical_paths(IWorkDocumentKind kind) {
        using var package = ParagraphLayoutPackage(kind, Message(
            BytesField(13, Message(VarintField(1, 2), FloatField(2, 24))),
            BytesField(25, BytesField(1, FloatField(1, 18)))));
        var text = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package, kind), kind).Item1.Cells).RichText!;
        Assert.Equal("Value", text.PlainText);
        Assert.False(text.IsFormattingComplete);
        package.Position = 0;
        if (kind == IWorkDocumentKind.Pages) {
            using var fallback = WordIWorkConverter.ConvertPagesToWordResult(package);
            Assert.True(fallback.IsVisualFallback);
        } else if (kind == IWorkDocumentKind.Numbers) {
            using var fallback = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);
            Assert.True(fallback.IsVisualFallback);
        } else {
            using var fallback = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package);
            Assert.True(fallback.IsVisualFallback);
        }
        package.Position = 0;
        var report = ConvertUnitReport(package, kind, visual: true,
            readOptions: new IWorkReadOptions { PreserveSourceRecords = false });
        Assert.Equal(new[] { "12/13", "12/25" }, report.SourceDeclarationIssues.Select(issue => issue.FieldPath).OrderBy(path => path));
        Assert.All(report.SourceDeclarationIssues, issue => {
            Assert.Equal(19ul, issue.Owner.RecordIdentifier);
            Assert.Equal(1, issue.DeclaredValueCount);
            Assert.Equal(IWorkSourceDeclarationIssueKind.UnsupportedField, issue.Kind);
        });
        Assert.Empty(report.PreservedRecords);
        Assert.Empty(report.SourceReferenceIssues);
        package.Position = 0;
        Assert.True(ConvertUnitReport(package, kind, visual: false).IsPartialEditableReconstruction);
    }

    [Fact]
    public void Cleared_paragraph_layout_and_unselected_styles_do_not_invent_active_layout_issues() {
        using var package = SelectedRichTextPackage(IWorkDocumentKind.Numbers,
            AttributeTable(5, AttributeEntry(0, ReferenceField(2, 19))),
            additionalRecords: Message(
                ArchiveRecord(19, 2022, BytesField(12, Message(VarintField(12, 1), VarintField(24, 1)))),
                ArchiveRecord(20, 2022, BytesField(12, BytesField(13, FloatField(2, 24))))));
        var text = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package), IWorkDocumentKind.Numbers).Item1.Cells).RichText!;
        Assert.True(text.IsFormattingComplete);
        package.Position = 0;
        Assert.Empty(ConvertUnitReport(package, IWorkDocumentKind.Numbers).SourceDeclarationIssues);
    }

    [Theory]
    [InlineData(0, "12/13", IWorkSourceDeclarationIssueKind.MalformedMessage)]
    [InlineData(1, "12/25", IWorkSourceDeclarationIssueKind.InvalidValue)]
    [InlineData(2, "12/13", IWorkSourceDeclarationIssueKind.InvalidValue)]
    [InlineData(3, "12/24", IWorkSourceDeclarationIssueKind.InvalidValue)]
    public void Malformed_or_conflicting_paragraph_layout_preserves_bounded_diagnostics(int defect, string path, IWorkSourceDeclarationIssueKind kind) {
        byte[] fields = defect switch {
            0 => BytesField(13, new byte[] { 0x80 }),
            1 => Message(BytesField(25, Message()), BytesField(25, Message())),
            2 => Message(VarintField(12, 1), BytesField(13, FloatField(2, 24))),
            _ => VarintField(24, 2)
        };
        using var package = ParagraphLayoutPackage(IWorkDocumentKind.Numbers, fields);
        var issue = Assert.Single(ConvertUnitReport(package, IWorkDocumentKind.Numbers).SourceDeclarationIssues);
        Assert.Equal(path, issue.FieldPath); Assert.Equal(kind, issue.Kind);
        Assert.Equal(defect == 1 ? 2 : 1, issue.DeclaredValueCount);
    }

    [Fact]
    public void Paragraph_layout_declarations_share_the_existing_issue_budget() {
        using var package = ParagraphLayoutPackage(IWorkDocumentKind.Numbers,
            Message(BytesField(13, Message()), BytesField(25, Message())));
        var source = IWorkSourceDocument.Open(package, new IWorkReadOptions { MaximumSourceDeclarationIssues = 1 });
        Assert.Contains("declaration issues", Assert.Throws<InvalidDataException>(() => source.ReadNumbers()).Message);
    }

    [Fact]
    public void Native_Pages_selected_layout_declarations_match_independent_style_evidence() {
        using var manifest = JsonDocument.Parse(File.ReadAllText(CorpusFixture("pages-paragraph-layout.json")));
        string path = CorpusFixture(manifest.RootElement.GetProperty("source").GetString()!);
        Assert.Equal(manifest.RootElement.GetProperty("sourceSha256").GetString(), HashFile(path));
        using var result = WordIWorkConverter.ConvertPagesToWordResult(path,
            conversionOptions: new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
        foreach (var style in manifest.RootElement.GetProperty("styles").EnumerateArray()) {
            ulong identifier = style.GetProperty("recordIdentifier").GetUInt64();
            foreach (var declaration in style.GetProperty("declarations").EnumerateArray()) {
                if (declaration.TryGetProperty("relativeMultiplier", out _)) {
                    Assert.DoesNotContain(result.Report.SourceDeclarationIssues, item => item.Owner.RecordIdentifier == identifier
                        && item.FieldPath == declaration.GetProperty("fieldPath").GetString());
                    continue;
                }
                var issue = Assert.Single(result.Report.SourceDeclarationIssues, item => item.Owner.RecordIdentifier == identifier
                    && item.FieldPath == declaration.GetProperty("fieldPath").GetString());
                Assert.Equal(IWorkSourceDeclarationIssueKind.UnsupportedField, issue.Kind);
            }
        }
        double expected = manifest.RootElement.GetProperty("styles").EnumerateArray()
            .SelectMany(style => style.GetProperty("declarations").EnumerateArray())
            .First(declaration => declaration.TryGetProperty("relativeMultiplier", out _))
            .GetProperty("relativeMultiplier").GetDouble();
        Assert.Contains(result.Projection.Body.Paragraphs, paragraph => paragraph.Style.LineSpacingMultiplier == expected);
        Assert.Contains(result.Value.Paragraphs, paragraph => paragraph.LineSpacingRule == OfficeIMO.Word.WordLineSpacingRule.Auto
            && paragraph.LineSpacing == 276);
        Assert.False(result.IsVisualFallback);
        Assert.True(result.Report.IsPartialEditableReconstruction);
    }

    private static MemoryStream ParagraphLayoutPackage(IWorkDocumentKind kind, byte[] properties) =>
        SelectedRichTextPackage(kind, AttributeTable(5, AttributeEntry(0, ReferenceField(2, 19))),
            additionalRecords: ArchiveRecord(19, 2022, BytesField(12, properties)));
}
