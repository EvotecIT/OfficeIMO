using OfficeIMO.IWork;
using OfficeIMO.Reader;
using OfficeIMO.Reader.IWork;
using System.Threading;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(43)]
    public void Keynote_selected_build_declarations_require_explicit_partial_conversion(int field) {
        using MemoryStream package = KeynoteWithBuildDeclarations(BytesField(field, Message()));
        using var fallback = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package);
        Assert.True(fallback.IsVisualFallback);
        Assert.False(fallback.Projection.HasEditableContent);
        IWorkSourceDeclarationIssue issue = Assert.Single(fallback.Report.SourceDeclarationIssues);
        Assert.Equal(field.ToString(), issue.FieldPath);
        Assert.Equal(IWorkSourceDeclarationIssueKind.UnsupportedField, issue.Kind);
        package.Position = 0;
        using var partial = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package,
            conversionOptions: new IWorkConversionOptions { Mode = IWorkConversionMode.EditableOnly,
                AllowPartialEditableReconstruction = true });
        Assert.False(partial.IsVisualFallback);
        Assert.Contains(partial.Report.Diagnostics, d => d.Code == "IWORK_KEYNOTE_BUILDS_UNSUPPORTED"
            && d.LossKind == OfficeConversionLossKind.Omission);
        package.Position = 0;
        var read = IWorkReaderAdapter.ReadDocument(package, "builds.key", new ReaderOptions(),
            new ReaderIWorkOptions(), CancellationToken.None);
        Assert.Contains(read.Diagnostics, d => d.Code == "IWORK_KEYNOTE_BUILDS_UNSUPPORTED");
    }

    [Fact]
    public void Keynote_malformed_build_envelope_retains_physical_field_evidence() {
        using MemoryStream package = KeynoteWithBuildDeclarations(
            Message(VarintField(43, 1), BytesField(43, Message())));
        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package);
        Assert.True(result.IsVisualFallback);
        IWorkSourceDeclarationIssue issue = Assert.Single(result.Report.SourceDeclarationIssues);
        Assert.Equal("43", issue.FieldPath);
        Assert.Equal(IWorkSourceDeclarationIssueKind.MalformedMessage, issue.Kind);
    }

    [Fact]
    public void Keynote_unused_slide_builds_do_not_gate_selected_content() {
        using MemoryStream package = KeynoteWithBuildDeclarations(Message(),
            ArchiveRecord(9, 5, Message(BytesField(2, Message()))));
        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package);
        Assert.False(result.IsVisualFallback);
        Assert.DoesNotContain(result.Report.Diagnostics, d => d.Code == "IWORK_KEYNOTE_BUILDS_UNSUPPORTED");
    }

    [Fact]
    public void Keynote_build_declaration_budget_is_enforced_before_reporting() {
        using MemoryStream package = KeynoteWithBuildDeclarations(
            Message(BytesField(2, Message()), BytesField(3, Message())));
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Keynote,
            new IWorkReadOptions { MaximumSourceDeclarationIssues = 1 });
        Assert.Throws<InvalidDataException>(() => source.ReadKeynote());
    }

    private static MemoryStream KeynoteWithBuildDeclarations(byte[] buildFields, params byte[][] extra) =>
        CreatePackage(("Index/Slide.iwa", FrameIwa(Message(
            new[] {
                ArchiveRecord(1, 1, Message(ReferenceField(2, 2))),
                ArchiveRecord(2, 2, KeynoteShow(Message(ReferenceField(2, 3)))),
                ArchiveRecord(3, 4, Message(ReferenceField(2, 4))),
                ArchiveRecord(4, 5, Message(buildFields, ReferenceField(5, 5))),
                ArchiveRecord(5, 2011, Message(ReferenceField(2, 6))),
                ArchiveRecord(6, 2001, Message(StringField(3, "Static slide text")))
            }.Concat(extra).ToArray()))), ("preview.png", ValidPreviewPng()));
}
