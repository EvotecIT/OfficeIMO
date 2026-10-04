using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(IWorkDocumentKind.Pages, false)]
    [InlineData(IWorkDocumentKind.Pages, true)]
    [InlineData(IWorkDocumentKind.Numbers, false)]
    [InlineData(IWorkDocumentKind.Numbers, true)]
    [InlineData(IWorkDocumentKind.Keynote, false)]
    [InlineData(IWorkDocumentKind.Keynote, true)]
    public void Modern_solid_text_fill_overrides_the_legacy_color_for_selected_character_and_paragraph_styles(
        IWorkDocumentKind kind, bool paragraphStyle) {
        byte[] properties = Message(BytesField(7, TextRgbColor(1, 0, 0)),
            BytesField(46, BytesField(1, TextRgbColor(0, .2f, .4f))));
        using var package = SelectedRichTextPackage(kind,
            AttributeTable(paragraphStyle ? 5 : 8, AttributeEntry(0, ReferenceField(2, 19))),
            additionalRecords: Message(ArchiveRecord(19, paragraphStyle ? 2022u : 2021u, BytesField(11, properties)),
                ArchiveRecord(20, 2021, BytesField(11, VarintField(46, 1)))));
        var (table, _) = ReadSelectedRichTable(IWorkSourceDocument.Open(package, kind), kind);
        IWorkTextContent text = Assert.Single(table.Cells).RichText!;
        Assert.True(text.IsFormattingComplete);
        Assert.Equal("003366", Assert.Single(Assert.Single(text.Paragraphs).Runs).Style!.Color!.RgbHex);
        package.Position = 0;
        Assert.Empty(ConvertUnitReport(package, kind, visual: false).SourceDeclarationIssues);
    }

    [Theory]
    [InlineData("gradient", "11/46", IWorkSourceDeclarationIssueKind.InvalidValue)]
    [InlineData("alpha", "11/46", IWorkSourceDeclarationIssueKind.InvalidValue)]
    [InlineData("p3", "11/46", IWorkSourceDeclarationIssueKind.InvalidValue)]
    [InlineData("wire", "11/46", IWorkSourceDeclarationIssueKind.InvalidValue)]
    [InlineData("duplicate", "11/46", IWorkSourceDeclarationIssueKind.InvalidValue)]
    [InlineData("clear-conflict", "11/46", IWorkSourceDeclarationIssueKind.InvalidValue)]
    [InlineData("empty", "11/46", IWorkSourceDeclarationIssueKind.UnsupportedField)]
    [InlineData("clear", "11/45", IWorkSourceDeclarationIssueKind.UnsupportedField)]
    [InlineData("container", "11/47", IWorkSourceDeclarationIssueKind.UnsupportedField)]
    public void Selected_text_fill_outside_the_opaque_solid_profile_retains_physical_evidence(
        string defect, string path, IWorkSourceDeclarationIssueKind expectedKind) {
        byte[] solid = BytesField(46, BytesField(1, TextRgbColor(0, .2f, .4f)));
        byte[] properties = defect switch {
            "gradient" => BytesField(46, BytesField(2, Message())),
            "alpha" => BytesField(46, BytesField(1, TextRgbColor(0, .2f, .4f, .5f))),
            "p3" => BytesField(46, BytesField(1, TextRgbColor(0, .2f, .4f, colorSpace: 2))),
            "wire" => VarintField(46, 1),
            "duplicate" => Message(solid, solid),
            "clear-conflict" => Message(VarintField(45, 1), solid),
            "empty" => BytesField(46, Message()),
            "clear" => VarintField(45, 1),
            _ => Message(solid, VarintField(47, 1))
        };
        using var package = SelectedRichTextPackage(IWorkDocumentKind.Numbers,
            AttributeTable(8, AttributeEntry(0, ReferenceField(2, 19))),
            additionalRecords: ArchiveRecord(19, 2021, BytesField(11, properties)));
        var (table, _) = ReadSelectedRichTable(IWorkSourceDocument.Open(package), IWorkDocumentKind.Numbers);
        Assert.False(Assert.Single(table.Cells).RichText!.IsFormattingComplete);
        package.Position = 0;
        IWorkConversionReport report = ConvertUnitReport(package, IWorkDocumentKind.Numbers, visual: false);
        Assert.True(report.IsPartialEditableReconstruction);
        IWorkSourceDeclarationIssue issue = Assert.Single(report.SourceDeclarationIssues);
        Assert.Equal(19ul, issue.Owner.RecordIdentifier);
        Assert.Equal(path, issue.FieldPath);
        Assert.Equal(expectedKind, issue.Kind);
        Assert.Equal(defect == "duplicate" ? 2 : 1, issue.DeclaredValueCount);
    }

    private static byte[] TextRgbColor(float red, float green, float blue, float alpha = 1, ulong colorSpace = 1) =>
        Message(VarintField(1, 1), FloatField(3, red), FloatField(4, green), FloatField(5, blue),
            FloatField(6, alpha), VarintField(12, colorSpace));
}
