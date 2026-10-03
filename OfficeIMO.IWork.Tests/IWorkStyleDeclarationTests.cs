using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Ambiguous_selected_style_flags_preserve_inheritance_and_identify_the_physical_property(IWorkDocumentKind kind) {
        using MemoryStream package = AmbiguousStylePackage(kind);
        IWorkTextContent text = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package, kind), kind).Item1.Cells).RichText!;
        Assert.False(text.IsFormattingComplete);
        Assert.False(Assert.Single(Assert.Single(text.Paragraphs).Runs).Style.Bold);
        package.Position = 0;
        IWorkConversionReport report = ConvertUnitReport(package, kind);
        IWorkSourceDeclarationIssue issue = Assert.Single(report.SourceDeclarationIssues);
        Assert.Equal(19ul, issue.Owner.RecordIdentifier);
        Assert.Equal("11/1", issue.FieldPath);
        Assert.Equal(2, issue.DeclaredValueCount);
        Assert.Equal(IWorkSourceDeclarationIssueKind.InvalidValue, issue.Kind);
        if (kind == IWorkDocumentKind.Numbers) return;
        package.Position = 0;
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, kind);
        var policy = new IWorkConversionOptions { AllowPartialEditableReconstruction = true };
        using var saved = new MemoryStream();
        if (kind == IWorkDocumentKind.Pages) {
            using var result = source.ToWordDocumentResult(policy);
            Assert.False(result.IsVisualFallback); Assert.True(result.Report.IsPartialEditableReconstruction);
            result.Value.Save(saved); saved.Position = 0;
            using var reopened = OfficeIMO.Word.WordDocument.Load(saved);
            var run = Assert.Single(reopened.Tables[0].Rows[0].Cells[0].Paragraphs, paragraph => paragraph.Text == "Value");
            Assert.False(run.Bold);
        } else {
            using var result = source.ToPowerPointPresentationResult(policy);
            Assert.False(result.IsVisualFallback); Assert.True(result.Report.IsPartialEditableReconstruction);
            result.Value.Save(saved); saved.Position = 0;
            using var reopened = OfficeIMO.PowerPoint.PowerPointPresentation.Load(saved);
            var cell = Assert.Single(reopened.Slides[0].Tables).GetCell(0, 0);
            Assert.Equal("Value", cell.Text); Assert.False(cell.Paragraphs[0].Runs[0].Bold);
            Assert.Empty(reopened.ValidateDocument());
        }
    }

    private static MemoryStream AmbiguousStylePackage(IWorkDocumentKind kind) =>
        SelectedRichTextPackage(kind, AttributeTable(8, AttributeEntry(0, ReferenceField(2, 19))),
            additionalRecords: Message(
                ArchiveRecord(19, 2021, Message(BytesField(1, ReferenceField(3, 20)),
                    BytesField(11, Message(VarintField(1, 0), VarintField(1, 1))))),
                ArchiveRecord(20, 2021, Message(BytesField(11, Message(VarintField(1, 0), FloatField(3, 14), StringField(5, "Arial")))))));

    public static IEnumerable<object[]> SelectedStyleValueCases() {
        foreach (IWorkDocumentKind kind in new[] { IWorkDocumentKind.Pages, IWorkDocumentKind.Numbers, IWorkDocumentKind.Keynote })
            foreach (string defect in new[] { "font-size", "font-name", "font-clear-conflict", "foreground-color", "background-color", "alignment", "paragraph-spacing", "style-name", "hyperlink-text", "hyperlink-missing" })
                yield return new object[] { kind, defect };
    }

    [Theory]
    [MemberData(nameof(SelectedStyleValueCases))]
    public void Selected_style_value_failures_retain_owner_paths_without_inventing_reference_or_content_counts(
        IWorkDocumentKind kind, string defect) {
        var (type, attribute, payload, path) = defect switch {
            "font-size" => (2021u, 8, BytesField(11, FloatField(3, float.NaN)), "11/3"),
            "font-name" => (2021u, 8, BytesField(11, BytesField(5, new byte[] { 0xff })), "11/5"),
            "font-clear-conflict" => (2021u, 8, BytesField(11, Message(VarintField(4, 1), StringField(5, "Arial"))), "11/5"),
            "foreground-color" => (2021u, 8, BytesField(11, BytesField(7, FloatField(3, 1))), "11/7"),
            "background-color" => (2021u, 8, BytesField(11, BytesField(26, new byte[] { 0x80 })), "11/26"),
            "alignment" => (2022u, 5, BytesField(12, VarintField(1, 9)), "12/1"),
            "paragraph-spacing" => (2022u, 5, BytesField(12, FloatField(21, float.PositiveInfinity)), "12/21"),
            "style-name" => (2021u, 8, BytesField(1, BytesField(1, new byte[] { 0xff })), "1/1"),
            "hyperlink-text" => (2032u, 11, BytesField(2, new byte[] { 0xff }), "2"),
            _ => (2032u, 11, Message(), "2")
        };
        using MemoryStream package = SelectedRichTextPackage(kind,
            AttributeTable(attribute, AttributeEntry(0, ReferenceField(2, 19))),
            additionalRecords: ArchiveRecord(19, type, payload));
        IWorkTextContent text = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package, kind), kind).Item1.Cells).RichText!;
        Assert.Equal("Value", text.PlainText);
        Assert.False(text.IsFormattingComplete);
        package.Position = 0;
        IWorkConversionReport report = ConvertUnitReport(package, kind, visual: true,
            readOptions: new IWorkReadOptions { PreserveSourceRecords = false });
        IWorkSourceDeclarationIssue issue = Assert.Single(report.SourceDeclarationIssues);
        Assert.Equal(19ul, issue.Owner.RecordIdentifier);
        Assert.Equal(path, issue.FieldPath);
        Assert.Equal(defect == "hyperlink-missing" ? 0 : 1, issue.DeclaredValueCount);
        Assert.Equal(IWorkSourceDeclarationIssueKind.InvalidValue, issue.Kind);
        Assert.Empty(report.SourceReferenceIssues);
        Assert.Empty(report.PreservedRecords);
        Assert.Contains(report.FidelityDiagnostics, diagnostic => diagnostic.Code == "IWORK_SOURCE_DECLARATIONS_UNASSESSED"
            && diagnostic.LossKind == OfficeConversionLossKind.Unassessed);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Ambiguous_clear_selection_preserves_inherited_font_and_shared_evidence_is_bounded(IWorkDocumentKind kind) {
        using MemoryStream package = SelectedRichTextPackage(kind,
            AttributeTable(8, AttributeEntry(0, ReferenceField(2, 19))), aliases: true,
            additionalRecords: Message(
                ArchiveRecord(19, 2021, Message(BytesField(1, ReferenceField(3, 20)),
                    BytesField(11, Message(VarintField(4, 0), VarintField(4, 1), StringField(5, "Other"))))),
                ArchiveRecord(20, 2021, Message(BytesField(11, StringField(5, "Arial"))))));
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, kind,
            new IWorkReadOptions { MaximumSourceDeclarationIssues = 1 });
        IWorkTable table = ReadSelectedRichTable(source, kind).Item1;
        Assert.Equal(2, table.Cells.Count);
        foreach (IWorkTableCell cell in table.Cells) {
            Assert.False(cell.RichText!.IsFormattingComplete);
            Assert.Equal("Arial", Assert.Single(Assert.Single(cell.RichText.Paragraphs).Runs).Style.FontName);
        }
        package.Position = 0;
        IWorkSourceDeclarationIssue issue = Assert.Single(ConvertUnitReport(package, kind,
            readOptions: new IWorkReadOptions { MaximumSourceDeclarationIssues = 1 }).SourceDeclarationIssues);
        Assert.Equal("11/4", issue.FieldPath); Assert.Equal(2, issue.DeclaredValueCount);
    }

    [Fact]
    public void Selected_style_evidence_limit_remains_fatal_before_partial_conversion() {
        using MemoryStream package = SelectedRichTextPackage(IWorkDocumentKind.Numbers,
            AttributeTable(8, AttributeEntry(0, ReferenceField(2, 19))),
            additionalRecords: ArchiveRecord(19, 2021, BytesField(11,
                Message(VarintField(1, 2), FloatField(3, float.NaN)))));
        IWorkSourceDocument source = IWorkSourceDocument.Open(package,
            new IWorkReadOptions { MaximumSourceDeclarationIssues = 1 });
        Assert.Contains("declaration issues", Assert.Throws<InvalidDataException>(() => source.ReadNumbers()).Message);
    }

    [Fact]
    public void Selected_style_nested_color_field_limit_is_not_downgraded_to_invalid_value_evidence() {
        using MemoryStream package = SelectedRichTextPackage(IWorkDocumentKind.Numbers,
            AttributeTable(8, AttributeEntry(0, ReferenceField(2, 19))),
            additionalRecords: ArchiveRecord(19, 2021, BytesField(11, BytesField(7,
                Message(Enumerable.Range(3, 9).Select(field => FloatField(field, 0)).ToArray())))));
        IWorkSourceDocument source = IWorkSourceDocument.Open(package,
            new IWorkReadOptions { MaximumProtobufFieldCount = 8 });
        Assert.Contains("field limit", Assert.Throws<InvalidDataException>(() => source.ReadNumbers()).Message);
    }

}
