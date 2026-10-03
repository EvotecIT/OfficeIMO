using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData("missing", IWorkSourceReferenceIssueKind.MissingTarget)]
    [InlineData("wrong-type", IWorkSourceReferenceIssueKind.UnexpectedTargetType)]
    [InlineData("malformed", IWorkSourceReferenceIssueKind.MalformedReference)]
    [InlineData("duplicate", IWorkSourceReferenceIssueKind.RejectedReferenceSet)]
    public void Keynote_invalid_template_reference_gates_conversion_even_with_local_background(
        string kind, IWorkSourceReferenceIssueKind expected) {
        byte[] reference = kind == "malformed" ? VarintField(17, 12)
            : kind == "duplicate" ? Message(ReferenceField(17, 12), ReferenceField(17, 12))
            : ReferenceField(17, 12);
        var records = new List<byte[]> { SlideBackgroundStyle(10, FillColor(1, 0, 0)) };
        if (kind != "missing") records.Add(ArchiveRecord(12, kind == "wrong-type" ? 6004u : 5u, Message()));
        using var package = KeynoteWithBuildDeclarations(Message(ReferenceField(1, 10), reference), records.ToArray());
        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package);
        Assert.True(result.IsVisualFallback);
        Assert.Equal("FF0000", result.Projection.Slides[0].BackgroundColor?.RgbHex);
        Assert.Contains(result.Report.SourceReferenceIssues, issue => issue.FieldPath == "17" && issue.Kind == expected);
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "IWORK_KEYNOTE_TEMPLATE_UNRESOLVED");
        package.Position = 0;
        using var partial = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package,
            conversionOptions: new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
        Assert.False(partial.IsVisualFallback);
        Assert.Equal("FF0000", partial.Value.Slides[0].BackgroundColor);
        Assert.Contains(partial.Report.Diagnostics, d => d.Code == "IWORK_KEYNOTE_TEMPLATE_UNRESOLVED");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Keynote_invalid_direct_template_record_does_not_claim_complete_conversion(bool selfReference) {
        using var package = KeynoteWithBuildDeclarations(Message(ReferenceField(1, 10),
            ReferenceField(17, selfReference ? 4ul : 12ul)),
            SlideBackgroundStyle(10, FillColor(1, 0, 0)), ArchiveRecord(12, 5, new byte[] { 0x80 }));
        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package);
        Assert.True(result.IsVisualFallback);
        Assert.Equal("FF0000", result.Projection.Slides[0].BackgroundColor?.RgbHex);
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "IWORK_KEYNOTE_TEMPLATE_UNRESOLVED");
        Assert.Contains(result.Report.SourceDeclarationIssues, issue => issue.FieldPath == (selfReference ? "17" : "$"));
    }

    [Fact]
    public void Keynote_valid_template_reference_does_not_select_inactive_template_storage() {
        using var package = KeynoteWithBuildDeclarations(Message(ReferenceField(1, 10), ReferenceField(17, 12)),
            SlideBackgroundStyle(10, FillColor(1, 0, 0)), ArchiveRecord(12, 5, Message()),
            ArchiveRecord(13, 5, ReferenceField(17, 999)));
        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package);
        Assert.False(result.IsVisualFallback);
        Assert.DoesNotContain(result.Report.SourceReferenceIssues, issue => issue.FieldPath == "17");
    }
}
