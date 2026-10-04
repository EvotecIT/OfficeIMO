using OfficeIMO.IWork;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(9, false)]
    [InlineData(10, true)]
    [InlineData(14, true)]
    public void Qualified_fixed_frames_only_make_keep_lines_inactive(int paginationField, bool fallback) {
        using var package = NativeTextFramePackage(Message(), paginationField: paginationField);
        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package,
            conversionOptions: new IWorkConversionOptions { RequireCompleteVisualCoverage = false });
        Assert.Equal(fallback, result.IsVisualFallback);
        Assert.Equal(fallback, result.Report.Diagnostics.Any(diagnostic =>
            diagnostic.Code == "IWORK_KEYNOTE_PARAGRAPH_PAGINATION_UNSUPPORTED"));
        if (!fallback) result.Report.RequireCompleteEditableReconstruction();
    }

    [Theory]
    [InlineData(1, 2)]
    [InlineData(2, 3)]
    [InlineData(8, 1)]
    public void Rejected_selected_child_frame_properties_do_not_reuse_parent_layout(int field, ulong value) {
        using var package = NativeTextFramePackage(Message(VarintField(field, value)));
        IWorkKeynoteProjection projection = IWorkSourceDocument.Open(package, IWorkDocumentKind.Keynote).ReadKeynote();
        Assert.False(projection.HasEditableContent);
        Assert.Null(Assert.Single(projection.Slides).TitleBox!.Layout);
        Assert.Contains(projection.SourceDeclarationIssues, issue => issue.FieldPath == "11/" + field);
    }

    [Fact]
    public void Text_frame_inheritance_respects_the_configured_depth_and_rejects_cycles() {
        using var limited = NativeTextFramePackage(Message());
        var source = IWorkSourceDocument.Open(limited, IWorkDocumentKind.Keynote,
            new IWorkReadOptions { MaximumTextStyleInheritanceDepth = 1 });
        Assert.Contains("inheritance", Assert.Throws<InvalidDataException>(() => source.ReadKeynote()).Message);
        using var cyclic = NativeTextFramePackage(Message(), cyclicParent: true);
        Assert.False(IWorkSourceDocument.Open(cyclic, IWorkDocumentKind.Keynote).ReadKeynote().HasEditableContent);
    }

    [Fact]
    public void Connected_text_flow_is_not_qualified_as_a_fixed_frame() {
        using var package = NativeTextFramePackage(Message(), connectedFlow: true);
        IWorkKeynoteProjection projection = IWorkSourceDocument.Open(package, IWorkDocumentKind.Keynote).ReadKeynote();
        Assert.False(projection.HasEditableContent);
        Assert.Null(Assert.Single(projection.Slides).TitleBox!.Layout);
        Assert.Contains(projection.SourceDeclarationIssues, issue => issue.FieldPath == "1/3");
    }

    private static MemoryStream NativeTextFramePackage(byte[] childProperties,
        int paginationField = 9, bool cyclicParent = false, bool connectedFlow = false) {
        byte[] geometry = Message(BytesField(1, Message(FloatField(1, 55), FloatField(2, 146))),
            BytesField(2, Message(FloatField(1, 914), FloatField(2, 260))));
        byte[] shape = Message(BytesField(1, Message(BytesField(1, geometry))), ReferenceField(2, 7));
        byte[] info = Message(BytesField(1, shape), ReferenceField(2, 6),
            connectedFlow ? ReferenceField(3, 11) : Array.Empty<byte>());
        byte[] paragraphTable = Message(BytesField(1, Message(VarintField(1, 0), ReferenceField(2, 9))));
        byte[] padding = Message(FloatField(1, 4), FloatField(2, 4), FloatField(3, 4), FloatField(4, 4));
        byte[] parentProperties = Message(VarintField(1, 1), VarintField(2, 2),
            BytesField(4, Message(BytesField(1, Message(VarintField(1, 1))))), BytesField(6, padding));
        byte[] records = Message(
            ArchiveRecord(1, 1, Message(ReferenceField(2, 2))),
            ArchiveRecord(2, 2, KeynoteShow(Message(ReferenceField(2, 3)))),
            ArchiveRecord(3, 4, Message(ReferenceField(2, 4))),
            ArchiveRecord(4, 5, Message(ReferenceField(5, 5))),
            ArchiveRecord(5, 7, Message(BytesField(1, info))),
            ArchiveRecord(6, 2001, Message(StringField(3, "Frame text"), BytesField(5, paragraphTable))),
            ArchiveRecord(7, 2025, Message(BytesField(1, Message(BytesField(1, Message(ReferenceField(3, 8))))), BytesField(11, childProperties))),
            ArchiveRecord(8, 2025, Message(BytesField(1, Message(BytesField(1, cyclicParent ? Message(ReferenceField(3, 7)) : Message()))),
                BytesField(11, parentProperties))),
            ArchiveRecord(9, 2022, Message(BytesField(12, Message(VarintField(paginationField, 1))))),
            ArchiveRecord(11, 2010, Message()));
        return CreatePackage(("Index/Slide.iwa", FrameIwa(records)), ("preview.png", ValidPreviewPng()));
    }
}
