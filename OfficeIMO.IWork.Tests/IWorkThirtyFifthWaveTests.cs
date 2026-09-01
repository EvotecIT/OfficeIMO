using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Keynote_slide_bound_is_enforced_before_nested_references_are_parsed() {
        using MemoryStream package = CreateKeynotePackageWithMalformedSecondSlideReference();
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Keynote,
            new IWorkReadOptions { MaximumProjectedSlides = 1 });

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => source.ReadKeynote());

        Assert.Contains("slide count", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Numbers_merge_bound_is_enforced_before_nested_pairs_are_parsed() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Merges", 1, 1, 1d, malformedSecondMergePair: true)
        });
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers,
            new IWorkReadOptions { MaximumTableMergedRanges = 1 });

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => source.ReadNumbers());

        Assert.Contains("merged-range count", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Numbers_sheet_bound_is_enforced_before_nested_references_are_parsed() {
        using MemoryStream package = CreateNumbersPackage(Array.Empty<TableSpec>(),
            includeMalformedSecondSheetReference: true);
        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers,
                new IWorkReadOptions { MaximumProjectedSheets = 1 }));

        Assert.Contains("sheet count", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Formula_node_bound_is_enforced_before_nested_nodes_are_parsed() {
        byte[] nodeArray = Message(
            BytesField(1, Message(VarintField(1, 17), DoubleField(4, 1d))),
            BytesField(1, new byte[] { 0x80 }));
        IWorkWireMessage formula = IWorkProtobuf.Parse(
            Message(BytesField(1, nodeArray)), new IWorkReadOptions());

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            IWorkFormulaReader.Render(formula, 0, 0, maximumNodes: 1, maximumCharacters: 32));

        Assert.Contains("syntax-node limit", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("◦")]
    [InlineData("→")]
    [InlineData("◆")]
    public void Pages_owner_preserves_custom_bullet_glyphs(string glyph) {
        using MemoryStream package = CreatePagesPackageWithListLabel(glyph);
        using var result = WordDocument.LoadPagesWithReport(package);
        using var saved = new MemoryStream();
        result.Document.Save(saved);
        saved.Position = 0;

        using WordprocessingDocument document = WordprocessingDocument.Open(saved, false);
        Numbering numbering = document.MainDocumentPart?.NumberingDefinitionsPart?.Numbering
            ?? throw new InvalidDataException("The reconstructed DOCX has no numbering definitions.");
        Level level = Assert.Single(numbering.Elements<AbstractNum>().SelectMany(item =>
            item.Elements<Level>()));

        Assert.False(result.IsVisualFallback);
        Assert.Equal(glyph, level.LevelText?.Val?.Value);
    }

    [Fact]
    public void Keynote_visual_fallback_has_accessible_alternative_text() {
        using MemoryStream package = CreateKeynotePackageWithRepeatedSlides(1,
            rotation: float.MaxValue);

        using var result = PowerPointPresentation.LoadKeynoteWithReport(package);
        using var saved = new MemoryStream();
        result.Document.Save(saved);
        saved.Position = 0;
        using PowerPointPresentation reopened = PowerPointPresentation.Load(saved);
        PowerPointPicture picture = Assert.Single(Assert.Single(reopened.Slides).Pictures);

        Assert.True(result.IsVisualFallback);
        Assert.Equal("Visual fallback from the source Keynote package", picture.AltText);
    }

    private static MemoryStream CreateKeynotePackageWithMalformedSecondSlideReference() {
        const ulong documentId = 1;
        const ulong showId = 2;
        const ulong nodeId = 3;
        byte[] slideTree = Message(
            ReferenceField(2, nodeId),
            BytesField(2, new byte[] { 0x80 }));
        byte[] records = Message(
            ArchiveRecord(documentId, 1,
                Message(ReferenceField(2, showId)), new[] { showId }),
            ArchiveRecord(showId, 2,
                KeynoteShow(slideTree), new[] { nodeId }),
            ArchiveRecord(nodeId, 4, Message()));
        return CreatePackage(("Index/Slide.iwa", FrameIwa(records)));
    }
}
