using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Malformed_rich_text_cell_records_do_not_abort_table_decoding(bool malformedWrapper) {
        var options = new IWorkReadOptions();
        const ulong listId = 1;
        const ulong wrapperId = 2;
        const ulong storageId = 3;
        byte[] listPayload = Message(BytesField(3, Message(
            VarintField(1, 1), ReferenceField(9, wrapperId))));
        byte[] wrapperPayload = malformedWrapper
            ? new byte[] { 0x80 }
            : Message(ReferenceField(1, storageId));
        byte[] storagePayload = malformedWrapper
            ? Message(StringField(3, "Value"))
            : new byte[] { 0x80 };
        var records = new[] {
            Record(listId, 6005, listPayload),
            Record(wrapperId, 6218, wrapperPayload),
            Record(storageId, 2001, storagePayload)
        };
        var index = new IWorkObjectIndex(records, options);
        IWorkWireMessage store = IWorkProtobuf.Parse(
            Message(ReferenceField(17, listId)), options);

        IReadOnlyDictionary<uint, IWorkTextContent> strings = IWorkTableRichTextReader.Read(
            index, store, new IWorkProjectionBudget(options), options,
            options.MaximumTableCatalogEntries, out bool complete);

        Assert.False(complete);
        Assert.Empty(strings);

        static IWorkArchiveRecord Record(ulong id, uint type, byte[] payload) =>
            new(id, type, Array.Empty<uint>(), Array.Empty<ulong>(),
                Array.Empty<ulong>(), "Index/Tables/Test.iwa", 0, payload);
    }

    [Fact]
    public void Unresolved_rich_cell_style_does_not_invalidate_complete_plain_text() {
        var options = new IWorkReadOptions();
        const ulong listId = 1;
        const ulong wrapperId = 2;
        const ulong storageId = 3;
        const ulong missingStyleId = 99;
        byte[] listPayload = Message(BytesField(3, Message(
            VarintField(1, 1), ReferenceField(9, wrapperId))));
        byte[] storagePayload = Message(
            StringField(3, "Value"),
            BytesField(5, Message(BytesField(1, Message(
                VarintField(1, 0), ReferenceField(2, missingStyleId))))));
        var records = new[] {
            Record(listId, 6005, listPayload),
            Record(wrapperId, 6218, Message(ReferenceField(1, storageId))),
            Record(storageId, 2001, storagePayload)
        };
        var index = new IWorkObjectIndex(records, options);
        IWorkWireMessage store = IWorkProtobuf.Parse(
            Message(ReferenceField(17, listId)), options);

        IReadOnlyDictionary<uint, IWorkTextContent> strings = IWorkTableRichTextReader.Read(
            index, store, new IWorkProjectionBudget(options), options,
            options.MaximumTableCatalogEntries, out bool complete);

        Assert.True(complete);
        Assert.Equal("Value", strings[1].PlainText);
        Assert.True(strings[1].IsTextComplete);
        Assert.False(strings[1].IsComplete);

        static IWorkArchiveRecord Record(ulong id, uint type, byte[] payload) =>
            new(id, type, Array.Empty<uint>(), Array.Empty<ulong>(),
                Array.Empty<ulong>(), "Index/Tables/Test.iwa", 0, payload);
    }

    [Fact]
    public void Plain_text_cell_does_not_traverse_style_inheritance() {
        var options = new IWorkReadOptions { MaximumTextStyleInheritanceDepth = 2 };
        const ulong listId = 1;
        const ulong wrapperId = 2;
        const ulong storageId = 3;
        const ulong styleId = 10;
        byte[] storagePayload = Message(
            StringField(3, "Value"),
            BytesField(5, Message(BytesField(1, Message(
                VarintField(1, 0), ReferenceField(2, styleId))))));
        var records = new[] {
            Record(listId, 6005, Message(BytesField(3, Message(
                VarintField(1, 1), ReferenceField(9, wrapperId))))),
            Record(wrapperId, 6218, Message(ReferenceField(1, storageId))),
            Record(storageId, 2001, storagePayload),
            Record(styleId, 2022, Message(BytesField(1, Message(ReferenceField(3, styleId + 1))))),
            Record(styleId + 1, 2022, Message(BytesField(1, Message(ReferenceField(3, styleId + 2))))),
            Record(styleId + 2, 2022, Message())
        };
        var index = new IWorkObjectIndex(records, options);
        IWorkWireMessage store = IWorkProtobuf.Parse(
            Message(ReferenceField(17, listId)), options);

        IReadOnlyDictionary<uint, string> strings = IWorkTableRichTextReader.Read(
            index, store, new IWorkProjectionBudget(options), options,
            options.MaximumTableCatalogEntries, out bool complete);

        Assert.True(complete);
        Assert.Equal("Value", strings[1]);

        static IWorkArchiveRecord Record(ulong id, uint type, byte[] payload) =>
            new(id, type, Array.Empty<uint>(), Array.Empty<ulong>(),
                Array.Empty<ulong>(), "Index/Tables/Test.iwa", 0, payload);
    }
}
