using OfficeIMO.IWork;
using OfficeIMO.Reader;
using OfficeIMO.Reader.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData("missing", 2)]
    [InlineData("malformed-id", 2)]
    [InlineData("wrong-type", 2)]
    [InlineData("duplicate", 3)]
    [InlineData("skipped", 2)]
    public void Keynote_source_ordinals_survive_partial_slide_tree_recovery(string predecessor, int expectedIndex) {
        byte[] firstReference = predecessor == "malformed-id"
            ? BytesField(2, VarintField(2, 9)) : ReferenceField(2, 9);
        byte[] tree = predecessor == "duplicate"
            ? Message(firstReference, firstReference, ReferenceField(2, 3))
            : Message(firstReference, ReferenceField(2, 3));
        var records = new List<byte[]> {
            ArchiveRecord(1, 1, ReferenceField(2, 2)),
            ArchiveRecord(2, 2, KeynoteShow(tree)),
            ArchiveRecord(3, 4, ReferenceField(2, 4)),
            ArchiveRecord(4, 5, ReferenceField(5, 5)),
            ArchiveRecord(5, 2011, ReferenceField(2, 6)),
            ArchiveRecord(6, 2001, StringField(3, "Surviving source slide"))
        };
        if (predecessor == "wrong-type") records.Add(ArchiveRecord(9, 9999, Message()));
        if (predecessor is "duplicate" or "skipped") {
            records.Add(ArchiveRecord(9, 4, Message(ReferenceField(2, 10),
                predecessor == "skipped" ? VarintField(4, 1) : Message())));
            records.Add(ArchiveRecord(10, 5, Message()));
        }
        using var package = CreatePackage(("Index/Document.iwa", FrameIwa(Message(records.ToArray()))));
        var projection = IWorkSourceDocument.Open(package, IWorkDocumentKind.Keynote).ReadKeynote();
        var surviving = Assert.Single(projection.Slides, slide => slide.SourceIdentity?.RecordIdentifier == 4);
        Assert.Equal("Surviving source slide", surviving.Title);
        Assert.Equal(expectedIndex, surviving.Index);
        if (predecessor != "skipped") Assert.NotEmpty(projection.Diagnostics);

        package.Position = 0;
        var read = new OfficeDocumentReaderBuilder().AddIWorkHandler().Build().ReadDocument(package, "recovery.key");
        var chunk = Assert.Single(read.Chunks, item => item.Text.Contains("Surviving source slide"));
        Assert.Equal(expectedIndex, chunk.Location.Slide);
        Assert.Equal(expectedIndex, read.Pages.Last().Number);
    }
}
