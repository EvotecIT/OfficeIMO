using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData("none", false, false)]
    [InlineData("dissolve", false, true)]
    [InlineData("none", true, true)]
    public void Keynote_transition_selection_controls_static_conversion(string effect, bool automatic, bool fallback) {
        byte[] animation = Message(StringField(1, "Transition"), StringField(2, effect),
            DoubleField(3, 1), DoubleField(5, 0.5), VarintField(6, automatic ? 1UL : 0UL),
            VarintField(11, 123), VarintField(16, 0));
        using MemoryStream package = KeynoteWithBuildDeclarations(
            BytesField(4, Message(BytesField(2, Message(BytesField(8, animation))))));
        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package, conversionOptions: new IWorkConversionOptions { RequireCompleteVisualCoverage = false });
        Assert.Equal(fallback, result.IsVisualFallback);
        Assert.Equal(fallback, result.Report.Diagnostics.Any(d => d.Code == "IWORK_KEYNOTE_TRANSITION_UNSUPPORTED"));
        if (fallback) {
            Assert.Equal("4/2/8", Assert.Single(result.Report.SourceDeclarationIssues).FieldPath);
            package.Position = 0;
            using var partial = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package,
                conversionOptions: new IWorkConversionOptions { Mode = IWorkConversionMode.EditableOnly,
                    AllowPartialEditableReconstruction = true });
            Assert.False(partial.IsVisualFallback);
            Assert.Contains(partial.Report.Diagnostics, d => d.Code == "IWORK_KEYNOTE_TRANSITION_UNSUPPORTED"
                && d.LossKind == OfficeConversionLossKind.Omission);
        }
    }

    [Theory]
    [InlineData("iwork-converter/a.key")]
    [InlineData("nim-iwork/simple.key")]
    [InlineData("keynotekit/tabledeck-v15.2.1.key")]
    [InlineData("keynotekit/imagedeck-v15.2.1.key")]
    public void Keynote_native_no_effect_transitions_do_not_report_an_omission(string path) {
        IWorkKeynoteProjection projection = IWorkSourceDocument.Open(Fixture(path)).ReadKeynote();
        Assert.NotEmpty(projection.Slides);
        Assert.DoesNotContain(projection.Diagnostics, d => d.Code == "IWORK_KEYNOTE_TRANSITION_UNSUPPORTED");
    }

    [Fact]
    public void Keynote_transition_depth_limit_remains_fatal() {
        using MemoryStream package = KeynoteWithBuildDeclarations(BytesField(4, Message(BytesField(2,
            Message(BytesField(8, Message(StringField(1, "Transition"), StringField(2, "none"))))))));
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Keynote,
            new IWorkReadOptions { MaximumProtobufDepth = 2 });
        Assert.Throws<InvalidDataException>(() => source.ReadKeynote());
    }

    [Fact]
    public void Keynote_malformed_transition_retains_nested_path() {
        using MemoryStream package = KeynoteWithBuildDeclarations(
            BytesField(4, Message(BytesField(2, Message(BytesField(8, new byte[] { 0x80 }))))));
        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package, conversionOptions: new IWorkConversionOptions { RequireCompleteVisualCoverage = false });
        Assert.True(result.IsVisualFallback);
        IWorkSourceDeclarationIssue issue = Assert.Single(result.Report.SourceDeclarationIssues);
        Assert.Equal("4/2/8", issue.FieldPath);
        Assert.Equal(IWorkSourceDeclarationIssueKind.MalformedMessage, issue.Kind);
    }
}
