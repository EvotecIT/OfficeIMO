using OfficeIMO.IWork;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(1.235f, false, 0f)]
    [InlineData(1.23f, true, 0f)]
    [InlineData(1.23f, false, 18f)]
    public void Unqualified_marker_measurements_do_not_claim_editable_PPTX(
        float scale, bool mixedFontSizes, float additionalIndent) {
        using var package = CharacterListLayoutPackage(IWorkDocumentKind.Keynote, scale,
            mixedFontSizes: mixedFontSizes, additionalIndent: additionalIndent);
        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package,
            conversionOptions: new IWorkConversionOptions {
                AllowPartialEditableReconstruction = true, RequireCompleteVisualCoverage = false
            });
        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_KEYNOTE_POWERPOINT_DESTINATION_UNSUPPORTED");
        Assert.NotNull(result.Projection.Slides[0].TitleBox!.Content.Paragraphs[0].ListLayout);
    }

    [Fact]
    public void Absolute_marker_size_is_retained_as_an_unqualified_source_declaration() {
        using var package = CharacterListLayoutPackage(IWorkDocumentKind.Keynote, 1.23f,
            scaleWithText: false);
        var projection = IWorkSourceDocument.Open(package).ReadKeynote();
        Assert.False(projection.HasEditableContent);
        Assert.Contains(projection.SourceDeclarationIssues, issue =>
            issue.Owner.RecordIdentifier == 7 && issue.FieldPath == "14");
    }

    [Fact]
    public void DOCX_character_marker_geometry_requires_partial_policy_and_reports_approximation() {
        using var package = CharacterListLayoutPackage(IWorkDocumentKind.Pages, 1.23f);
        using var strict = WordIWorkConverter.ConvertPagesToWordResult(package,
            conversionOptions: new IWorkConversionOptions { RequireCompleteVisualCoverage = false });
        Assert.True(strict.IsVisualFallback);
        package.Position = 0;
        using var partial = WordIWorkConverter.ConvertPagesToWordResult(package,
            conversionOptions: new IWorkConversionOptions { Mode = IWorkConversionMode.EditableOnly,
                AllowPartialEditableReconstruction = true });
        Assert.False(partial.IsVisualFallback);
        Assert.Contains(partial.Report.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_PAGES_LIST_LAYOUT_APPROXIMATED"
            && diagnostic.LossKind == OfficeConversionLossKind.Approximation);
        Assert.Throws<InvalidOperationException>(() => partial.Report.RequireCompleteEditableReconstruction());
        Assert.NotNull(partial.Projection.Body.Paragraphs[0].ListLayout);
        Assert.Empty(partial.Value.ValidateDocument());
    }

    private static MemoryStream CharacterListLayoutPackage(IWorkDocumentKind kind, float scale,
        bool mixedFontSizes = false, float additionalIndent = 0f, bool scaleWithText = true) {
        byte[] roots = kind == IWorkDocumentKind.Pages
            ? ArchiveRecord(1, 10000, Message(ReferenceField(4, 6)))
            : Message(ArchiveRecord(1, 1, Message(ReferenceField(2, 2))),
                ArchiveRecord(2, 2, KeynoteShow(Message(ReferenceField(2, 3)))),
                ArchiveRecord(3, 4, Message(ReferenceField(2, 4))),
                ArchiveRecord(4, 5, Message(ReferenceField(5, 5))),
                ArchiveRecord(5, 2011, Message(ReferenceField(2, 6))));
        byte[] records = Message(roots,
            ArchiveRecord(6, 2001, Message(StringField(3, "Bullet"),
                AttributeTable(7, AttributeEntry(0, ReferenceField(2, 7))),
                AttributeTable(5, AttributeEntry(0, ReferenceField(2, 8))),
                mixedFontSizes ? AttributeTable(8, AttributeEntry(2, ReferenceField(2, 9))) : Message())),
            ArchiveRecord(7, 2023, Message(VarintField(11, 2), FloatField(12, 1), FloatField(13, 0),
                BytesField(14, Message(FloatField(1, scale), FloatField(2, 0),
                    VarintField(3, scaleWithText ? 1ul : 0ul))), StringField(16, "•"))),
            ArchiveRecord(8, 2022, Message(BytesField(11, FloatField(3, 38)),
                BytesField(12, FloatField(11, additionalIndent)))),
            ArchiveRecord(9, 2021, BytesField(11, FloatField(3, 30))));
        return CreatePackage(("Index/Document.iwa", FrameIwa(records)), ("preview.png", ValidPreviewPng()));
    }
}
