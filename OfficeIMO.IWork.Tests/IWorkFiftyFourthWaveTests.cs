using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Pages_tables_on_the_same_page_keep_their_source_order() {
        byte[] firstPage = Message(
            BytesField(2, Message(ReferenceField(1, 10))),
            BytesField(2, Message(ReferenceField(1, 20))));
        byte[] floating = Message(BytesField(1, firstPage));
        byte[] records = Message(
            ArchiveRecord(1, 10000,
                Message(ReferenceField(4, 2), ReferenceField(3, 3))),
            ArchiveRecord(2, 2001, Message(StringField(3, "Body"))),
            ArchiveRecord(3, 10020, floating),
            TableInfo(10, 11, "First"),
            TableModel(11, "First"),
            TableInfo(20, 21, "Second"),
            TableModel(21, "Second"));
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)));

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.False(result.IsVisualFallback);
        Assert.Equal(new[] { "First", "Second" },
            result.Document.Tables.Select(table => table.Description));
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(4)]
    public void Repeated_geometry_scalars_disable_editable_reconstruction(int component) {
        byte[] position = component == 1
            ? Message(FloatField(1, 10f), FloatField(1, 20f), FloatField(2, 10f))
            : Message(FloatField(1, 10f), FloatField(2, 10f));
        byte[] size = component == 2
            ? Message(FloatField(1, 40f), FloatField(1, 50f), FloatField(2, 30f))
            : Message(FloatField(1, 40f), FloatField(2, 30f));
        byte[] geometry = component == 4
            ? Message(BytesField(1, position), BytesField(2, size),
                FloatField(4, 0f), FloatField(4, 15f))
            : Message(BytesField(1, position), BytesField(2, size));
        using MemoryStream package = CreatePagesPackage(
            includeBody: true, textBox: "Shape", includePreview: true,
            textBoxDrawable: Message(BytesField(1, geometry)));

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_PAGES_DRAWABLE_UNSUPPORTED");
    }

    private static byte[] TableInfo(ulong identifier, ulong modelIdentifier,
        string description) => ArchiveRecord(identifier, 6000,
        Message(BytesField(1, Message(StringField(8, description))),
            ReferenceField(2, modelIdentifier)));

    private static byte[] TableModel(ulong identifier, string name) =>
        ArchiveRecord(identifier, 6001,
            Message(BytesField(4, Message(BytesField(3, Message()))),
                VarintField(6, 1), VarintField(7, 1), StringField(8, name)));
}
