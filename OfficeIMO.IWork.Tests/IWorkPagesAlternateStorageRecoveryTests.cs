using OfficeIMO.IWork;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(2)]
    [InlineData(4)]
    public void Pages_recovers_one_valid_storage_when_the_alternate_target_has_the_wrong_type(int validField) {
        int invalidField = validField == 2 ? 4 : 2;
        using var package = TextReferencePackage(IWorkDocumentKind.Pages,
            Message(BytesField(1, BytesField(1, GeometryDrawable(36f, 72f, 216f, 108f))),
                ReferenceField(validField, 11), ReferenceField(invalidField, 12)),
            Message(ArchiveRecord(11, 2001, StringField(3, "Recoverable text")),
                ArchiveRecord(12, 15, Message())));
        var projection = IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages).ReadPages();
        Assert.Equal("Recoverable text", Assert.Single(projection.TextBoxes));
        var issue = Assert.Single(projection.SourceReferenceIssues);
        Assert.Equal(IWorkSourceReferenceIssueKind.UnexpectedTargetType, issue.Kind);
        Assert.Equal(invalidField.ToString(), issue.FieldPath);
        Assert.Equal(12ul, issue.TargetIdentifier);
        Assert.Contains(projection.Diagnostics, d => d.Code == "IWORK_PAGES_DRAWABLE_UNSUPPORTED");
        package.Position = 0;
        using var fallback = WordIWorkConverter.ConvertPagesToWordResult(package, conversionOptions: new IWorkConversionOptions { RequireCompleteVisualCoverage = false });
        Assert.True(fallback.IsVisualFallback);
        package.Position = 0;
        using var partial = WordIWorkConverter.ConvertPagesToWordResult(package,
            conversionOptions: new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
        Assert.False(partial.IsVisualFallback);
        using var saved = new MemoryStream(); partial.Value.Save(saved); saved.Position = 0;
        using var reopened = WordDocument.Load(saved);
        Assert.Equal("Recoverable text", Assert.Single(Assert.Single(reopened.TextBoxes).Paragraphs).Text);
    }

    [Fact]
    public void Pages_does_not_choose_between_two_different_valid_text_storages() {
        using var package = TextReferencePackage(IWorkDocumentKind.Pages,
            Message(ReferenceField(2, 11), ReferenceField(4, 12)),
            Message(ArchiveRecord(11, 2001, StringField(3, "First")),
                ArchiveRecord(12, 2001, StringField(3, "Second"))));
        var projection = IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages).ReadPages();
        Assert.Empty(projection.TextBoxes);
        Assert.Contains(projection.Diagnostics, d => d.Code == "IWORK_PAGES_DRAWABLE_UNSUPPORTED");
    }
}
