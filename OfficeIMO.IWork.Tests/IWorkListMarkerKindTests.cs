using OfficeIMO.IWork;
using OfficeIMO.PowerPoint;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Image_markers_do_not_reuse_dormant_text_or_claim_complete_formatting(IWorkDocumentKind kind) {
        using MemoryStream package = ListDeclarationPackage(kind,
            Message(VarintField(11, 1), StringField(16, "1.")),
            inherited: false, indent: 0, aliases: true);
        var source = IWorkSourceDocument.Open(package, kind);
        foreach (var cell in ReadSelectedRichTable(source, kind).Item1.Cells) {
            var paragraph = Assert.Single(cell.RichText!.Paragraphs);
            Assert.Equal(IWorkListMarkerKind.Image, paragraph.ListMarkerKind);
            Assert.Null(paragraph.ListLabel);
            Assert.False(cell.RichText.IsFormattingComplete);
        }
        switch (kind) {
            case IWorkDocumentKind.Pages:
                using (var result = source.ToWordDocumentResult()) Assert.True(result.IsVisualFallback);
                break;
            case IWorkDocumentKind.Numbers:
                using (var result = source.ToExcelDocumentResult()) Assert.True(result.IsVisualFallback);
                break;
            default:
                using (var result = source.ToPowerPointPresentationResult()) Assert.True(result.IsVisualFallback);
                break;
        }
        package.Position = 0;
        var report = ConvertUnitReport(package, kind, visual: true,
            readOptions: new IWorkReadOptions { PreserveSourceRecords = false, MaximumSourceDeclarationIssues = 1 });
        var issue = Assert.Single(report.SourceDeclarationIssues);
        Assert.Equal(19ul, issue.Owner.RecordIdentifier);
        Assert.Equal("11", issue.FieldPath);
        Assert.Equal(IWorkSourceDeclarationIssueKind.UnsupportedField, issue.Kind);
        Assert.Empty(report.PreservedRecords);
    }

    [Theory]
    [InlineData("1.")]
    [InlineData("iv.")]
    [InlineData("(a)")]
    [InlineData("Ab.")]
    public void Literal_text_markers_survive_docx_without_becoming_counters(string label) {
        using MemoryStream package = CreatePagesPackageWithListLabel(label);
        using var result = WordIWorkConverter.ConvertPagesToWordResult(package);
        Assert.False(result.IsVisualFallback);
        Assert.Equal(IWorkListMarkerKind.Text, Assert.Single(result.Projection.Body.Paragraphs).ListMarkerKind);
        using var saved = new MemoryStream();
        result.Value.Save(saved); saved.Position = 0;
        using var document = WordprocessingDocument.Open(saved, false);
        Level level = Assert.Single(document.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!
            .Elements<AbstractNum>().SelectMany(definition => definition.Elements<Level>()));
        Assert.Equal(NumberFormatValues.Bullet, level.NumberingFormat!.Val!.Value);
        Assert.Equal(label, level.LevelText!.Val!.Value);
        Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(document));
    }

    [Theory]
    [InlineData(0, "decimal", "%1.", PowerPointNumberingScheme.ArabicPeriod)]
    [InlineData(1, "decimal", "(%1)", PowerPointNumberingScheme.ArabicParenBoth)]
    [InlineData(2, "decimal", "%1)", PowerPointNumberingScheme.ArabicParenR)]
    [InlineData(3, "upperRoman", "%1.", PowerPointNumberingScheme.RomanUpperCharacterPeriod)]
    [InlineData(4, "upperRoman", "(%1)", PowerPointNumberingScheme.RomanUpperCharacterParenBoth)]
    [InlineData(5, "upperRoman", "%1)", PowerPointNumberingScheme.RomanUpperCharacterParenR)]
    [InlineData(6, "lowerRoman", "%1.", PowerPointNumberingScheme.RomanLowerCharacterPeriod)]
    [InlineData(7, "lowerRoman", "(%1)", PowerPointNumberingScheme.RomanLowerCharacterParenBoth)]
    [InlineData(8, "lowerRoman", "%1)", PowerPointNumberingScheme.RomanLowerCharacterParenR)]
    [InlineData(9, "upperLetter", "%1.", PowerPointNumberingScheme.AlphaUpperCharacterPeriod)]
    [InlineData(10, "upperLetter", "(%1)", PowerPointNumberingScheme.AlphaUpperCharacterParenBoth)]
    [InlineData(11, "upperLetter", "%1)", PowerPointNumberingScheme.AlphaUpperCharacterParenR)]
    [InlineData(12, "lowerLetter", "%1.", PowerPointNumberingScheme.AlphaLowerCharacterPeriod)]
    [InlineData(13, "lowerLetter", "(%1)", PowerPointNumberingScheme.AlphaLowerCharacterParenBoth)]
    [InlineData(14, "lowerLetter", "%1)", PowerPointNumberingScheme.AlphaLowerCharacterParenR)]
    public void Native_number_kinds_survive_saved_word_and_powerpoint(int kind, string wordFormat,
        string wordMarker, PowerPointNumberingScheme powerPointScheme) {
        byte[] fields = Message(VarintField(11, 3), VarintField(15, (ulong)kind));
        var policy = new IWorkConversionOptions { AllowPartialEditableReconstruction = true };
        using (MemoryStream package = ListDeclarationPackage(IWorkDocumentKind.Pages, fields, inherited: false, indent: 0)) {
            using var result = IWorkSourceDocument.Open(package).ToWordDocumentResult(policy);
            Assert.True(result.Report.IsPartialEditableReconstruction);
            using var saved = new MemoryStream();
            result.Value.Save(saved); saved.Position = 0;
            using var document = WordprocessingDocument.Open(saved, false);
            var level = Assert.Single(document.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!
                .Elements<AbstractNum>().SelectMany(definition => definition.Elements<Level>()));
            Assert.Equal(wordFormat, level.NumberingFormat!.Val!.InnerText);
            Assert.Equal(wordMarker, level.LevelText!.Val!.Value);
            Assert.Equal(1, level.StartNumberingValue!.Val!.Value);
            Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(document));
        }
        using (MemoryStream package = ListDeclarationPackage(IWorkDocumentKind.Keynote, fields, inherited: false, indent: 0)) {
            using var result = IWorkSourceDocument.Open(package).ToPowerPointPresentationResult(policy);
            Assert.True(result.Report.IsPartialEditableReconstruction);
            using var saved = new MemoryStream();
            result.Value.Save(saved); saved.Position = 0;
            using var reopened = PowerPointPresentation.Load(saved);
            var paragraph = Assert.Single(Assert.Single(reopened.Slides).Tables).GetCell(0, 0).Paragraphs[0];
            Assert.Equal(powerPointScheme, paragraph.NumberingScheme);
            Assert.Equal(1, paragraph.NumberingStartAt);
            Assert.Empty(reopened.ValidateDocument());
        }
    }

    [Fact]
    public void Image_marker_at_an_unselected_level_does_not_make_text_bullets_incomplete() {
        using MemoryStream package = ListDeclarationPackage(IWorkDocumentKind.Numbers,
            Message(VarintField(11, 1), VarintField(11, 2), FloatField(13, 0), FloatField(13, 18),
                StringField(16, "Dormant"), StringField(16, "•")), inherited: false, indent: 18);
        var projection = IWorkSourceDocument.Open(package).ReadNumbers();
        var text = Assert.Single(Assert.Single(projection.Sheets).Tables).Cells[0].RichText!;
        Assert.True(text.IsFormattingComplete);
        Assert.Equal(IWorkListMarkerKind.Text, Assert.Single(text.Paragraphs).ListMarkerKind);
        Assert.Empty(projection.SourceDeclarationIssues);
    }

    [Fact]
    public void Missing_number_kind_does_not_reuse_dormant_string_labels() {
        using MemoryStream package = ListDeclarationPackage(IWorkDocumentKind.Numbers,
            Message(VarintField(11, 3), StringField(16, "4.")), inherited: false, indent: 0);
        var text = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package),
            IWorkDocumentKind.Numbers).Item1.Cells).RichText!;
        Assert.Null(Assert.Single(text.Paragraphs).ListLabel);
        Assert.Equal(IWorkListMarkerKind.Number, Assert.Single(text.Paragraphs).ListMarkerKind);
        Assert.False(text.IsFormattingComplete);
    }

    [Theory]
    [InlineData("1.")]
    [InlineData("a")]
    public void PowerPoint_does_not_infer_numbering_from_literal_text(string label) {
        using MemoryStream package = CreateKeynotePackageWithRepeatedSlides(1, text: "Item", listLabel: label);
        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package);
        if (label.Length > 1) {
            Assert.True(result.IsVisualFallback);
        } else {
            Assert.False(result.IsVisualFallback);
            using var saved = new MemoryStream();
            result.Value.Save(saved); saved.Position = 0;
            using var reopened = OfficeIMO.PowerPoint.PowerPointPresentation.Load(saved);
            var paragraph = Assert.Single(Assert.Single(Assert.Single(reopened.Slides).TextBoxes).Paragraphs);
            Assert.False(paragraph.IsNumbered);
            Assert.Equal("a", paragraph.BulletCharacter);
            Assert.Empty(reopened.ValidateDocument());
        }
    }
}
