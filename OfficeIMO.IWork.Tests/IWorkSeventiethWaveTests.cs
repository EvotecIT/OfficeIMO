using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Rejected_raster_previews_do_not_consume_the_shared_decode_budget() {
        byte[] large = CreateSizedPreviewPng(100, 100);
        byte[] malformed = Message(large[..33],
            CreatePngChunk("IDAT", new byte[] { 0 }),
            CreatePngChunk("IEND", Array.Empty<byte>()));
        byte[] records = Message(ArchiveRecord(1, 10000, Message()));
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ("preview.png", malformed),
            ("preview-web.png", ValidPreviewPng()));

        IWorkSourceDocument source = IWorkSourceDocument.Open(package,
            IWorkDocumentKind.Pages,
            new IWorkReadOptions { MaximumPackageBytes = 40_002 });

        Assert.Equal("preview-web.png", source.PreferredRasterPreview?.Path);
    }

    [Fact]
    public void Formula_boolean_literals_require_zero_or_one() {
        byte[] booleanNode = Message(VarintField(1, 18), VarintField(5, 2));
        IWorkWireMessage formula = IWorkProtobuf.Parse(
            Message(BytesField(1, Message(BytesField(1, booleanNode)))),
            new IWorkReadOptions());

        IWorkFormulaResult result = IWorkFormulaReader.Render(
            formula, 0, 0, maximumNodes: 32, maximumCharacters: 128);

        Assert.False(result.IsComplete);
    }

    [Theory]
    [InlineData("node", "IWORK_KEYNOTE_SLIDE_NODE_UNSUPPORTED")]
    [InlineData("slide", "IWORK_KEYNOTE_SLIDE_MALFORMED")]
    [InlineData("note", "IWORK_KEYNOTE_NOTES_UNSUPPORTED")]
    [InlineData("drawable", "IWORK_KEYNOTE_DRAWABLE_UNSUPPORTED")]
    [InlineData("storage", "IWORK_KEYNOTE_TEXT_STORAGE_UNSUPPORTED")]
    public void Malformed_keynote_graph_records_use_visual_fallback(
        string malformedRecord, string diagnosticCode) {
        using MemoryStream package = CreateKeynotePackageWithMalformedGraphRecord(malformedRecord);

        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics,
            diagnostic => diagnostic.Code == diagnosticCode);
    }

    private static MemoryStream CreateKeynotePackageWithMalformedGraphRecord(string malformedRecord) {
        const ulong documentId = 1;
        const ulong showId = 2;
        const ulong nodeId = 3;
        const ulong slideId = 4;
        const ulong childId = 5;
        const ulong storageId = 6;
        byte[] malformed = { 0x80 };
        byte[] slideTree = Message(ReferenceField(2, nodeId));
        var records = new List<byte[]> {
            ArchiveRecord(documentId, 1,
                Message(ReferenceField(2, showId)), new[] { showId }),
            ArchiveRecord(showId, 2,
                KeynoteShow(slideTree), new[] { nodeId })
        };
        if (malformedRecord == "node") {
            records.Add(ArchiveRecord(nodeId, 4, malformed));
        } else {
            records.Add(ArchiveRecord(nodeId, 4,
                Message(ReferenceField(2, slideId)), new[] { slideId }));
            if (malformedRecord == "slide") {
                records.Add(ArchiveRecord(slideId, 5, malformed));
            } else if (malformedRecord == "note") {
                records.Add(ArchiveRecord(slideId, 5,
                    Message(ReferenceField(27, childId)), new[] { childId }));
                records.Add(ArchiveRecord(childId, 15, malformed));
            } else {
                records.Add(ArchiveRecord(slideId, 5,
                    Message(ReferenceField(5, childId)), new[] { childId }));
                if (malformedRecord == "drawable") {
                    records.Add(ArchiveRecord(childId, 2011, malformed));
                } else {
                    records.Add(ArchiveRecord(childId, 2011,
                        Message(ReferenceField(2, storageId)), new[] { storageId }));
                    records.Add(ArchiveRecord(storageId, 2001, malformed));
                }
            }
        }
        return CreatePackage(
            ("Index/Slide.iwa", FrameIwa(Message(records.ToArray()))),
            ("preview.png", ValidPreviewPng()));
    }
}
