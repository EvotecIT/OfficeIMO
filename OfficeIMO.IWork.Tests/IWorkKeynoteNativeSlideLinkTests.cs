using OfficeIMO.IWork;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    private const string NativeSecondSlideLink = "?slideid=1FB800A1-9682-4993-9868-F35E294F1088";

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void Keynote_native_targets_bind_selected_slide_nodes_without_changing_source_links(
        bool nextAction, bool missingPredecessor) {
        string target = nextAction ? "?slide=next" : NativeSecondSlideLink;
        using MemoryStream package = CreateNativeSlideLinkPackage(target,
            missingPredecessor: missingPredecessor);
        IWorkKeynoteProjection projection = IWorkSourceDocument.Open(package,
            IWorkDocumentKind.Keynote).ReadKeynote();
        IWorkKeynoteSlide sourceSlide = projection.Slides[0];
        Assert.Equal(missingPredecessor ? 2 : 1, sourceSlide.Index);
        Assert.Equal(target, sourceSlide.TitleBox!.Hyperlink);
        Assert.Equal(target, Assert.Single(Assert.Single(sourceSlide.TitleBox.Content.Paragraphs).Runs).Hyperlink);
        Assert.Equal(target, Assert.Single(Assert.Single(sourceSlide.PresenterNoteContent.Paragraphs).Runs).Hyperlink);

        package.Position = 0;
        using var conversion = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package,
            conversionOptions: new IWorkConversionOptions {
                Mode = IWorkConversionMode.EditableOnly,
                AllowPartialEditableReconstruction = missingPredecessor
            });
        AssertInternalLinks(conversion.Value);
        using var saved = new MemoryStream();
        conversion.Value.Save(saved); saved.Position = 0;
        using PowerPointDocument reopened = PowerPointDocument.Load(saved);
        AssertInternalLinks(reopened);
        Assert.Empty(reopened.ValidateDocument());
    }

    [Theory]
    [InlineData("unknown")]
    [InlineData("ambiguous")]
    [InlineData("malformed")]
    [InlineData("next-missing")]
    [InlineData("next-final")]
    public void Unresolved_native_targets_are_rejected_even_when_partial_reconstruction_is_allowed(string defect) {
        string target = defect.StartsWith("next-", StringComparison.Ordinal)
            ? "?slide=next" : defect == "unknown"
                ? "?slideid=11111111-2222-3333-4444-555555555555" : NativeSecondSlideLink;
        using MemoryStream package = CreateNativeSlideLinkPackage(target, defect: defect);
        IWorkKeynoteProjection projection = IWorkSourceDocument.Open(package,
            IWorkDocumentKind.Keynote).ReadKeynote();
        Assert.Equal(target, projection.Slides[0].TitleBox!.Hyperlink);
        if (defect == "malformed") Assert.Contains(projection.SourceDeclarationIssues,
            issue => issue.Kind == IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata);
        package.Position = 0;
        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package,
                conversionOptions: new IWorkConversionOptions {
                    Mode = IWorkConversionMode.EditableOnly,
                    AllowPartialEditableReconstruction = true
                }));
        Assert.Contains("hyperlink", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Native_uuid_catalog_obeys_the_source_wide_projection_limit() {
        using MemoryStream package = CreateNativeSlideLinkPackage(NativeSecondSlideLink, mapCount: 70);
        IWorkSourceDocument source = IWorkSourceDocument.Open(package,
            IWorkDocumentKind.Keynote, new IWorkReadOptions {
                MaximumProjectedTables = 1, MaximumProjectedImages = 1,
                MaximumProjectedTextItems = 64
            });
        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => source.ReadKeynote());
        Assert.Contains("combined projection limit of 66", exception.Message, StringComparison.Ordinal);
    }

    private static MemoryStream CreateNativeSlideLinkPackage(string target,
        bool missingPredecessor = false, string? defect = null, int mapCount = 1) {
        var tree = new List<byte[]>();
        if (missingPredecessor) tree.Add(ReferenceField(2, 99));
        tree.Add(ReferenceField(2, 3));
        if (defect == "next-missing") tree.Add(ReferenceField(2, 99));
        if (defect != "next-final") tree.Add(ReferenceField(2, 4));
        byte[] map = Message(VarintField(1, 4), BytesField(2, Message(
            VarintField(1, 10613157580333168671UL), VarintField(2, 9804423427890833560UL))));
        byte[] component = defect == "malformed" ? VarintField(11, 4)
            : Message(Enumerable.Repeat(BytesField(11, map), mapCount)
                .Concat(defect == "ambiguous" ? new[] { BytesField(11, Message(
                    VarintField(1, 3), BytesField(2, Message(VarintField(1, 10613157580333168671UL),
                        VarintField(2, 9804423427890833560UL))))) } : Array.Empty<byte[]>()).ToArray());
        byte[] hyperlinkTable = Message(BytesField(1, Message(VarintField(1, 0), ReferenceField(2, 9))));
        byte[] noteHyperlinkTable = Message(BytesField(1, Message(VarintField(1, 0), ReferenceField(2, 12))));
        byte[] records = Message(
            ArchiveRecord(1, 1, ReferenceField(2, 2)),
            ArchiveRecord(2, 2, KeynoteShow(Message(tree.ToArray()))),
            ArchiveRecord(3, 4, ReferenceField(2, 5)),
            ArchiveRecord(4, 4, ReferenceField(2, 6)),
            ArchiveRecord(5, 5, Message(ReferenceField(5, 7), ReferenceField(27, 10))),
            ArchiveRecord(6, 5, Message()),
            ArchiveRecord(7, 2011, Message(BytesField(1, Message(BytesField(1,
                StringField(4, target)))), ReferenceField(2, 8))),
            ArchiveRecord(8, 2001, Message(StringField(3, "Linked title"), BytesField(11, hyperlinkTable))),
            ArchiveRecord(9, 2032, StringField(2, target)),
            ArchiveRecord(10, 15, ReferenceField(1, 11)),
            ArchiveRecord(11, 2001, Message(StringField(3, "Linked note"), BytesField(11, noteHyperlinkTable))),
            ArchiveRecord(12, 2032, StringField(2, target)),
            ArchiveRecord(100, 11006, BytesField(3, component)));
        return CreatePackage(("Index/Document.iwa", FrameIwa(records)));
    }
}
