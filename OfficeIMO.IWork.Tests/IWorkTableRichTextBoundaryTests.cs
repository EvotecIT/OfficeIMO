using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;
using System.IO.Compression;
using System.Threading;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData("nim-iwork/simple.pages")]
    [InlineData("nim-iwork/simple.numbers")]
    [InlineData("nim-iwork/simple.key")]
    public void Structural_probe_recognizes_modern_packages_without_consuming_them(string name) {
        using FileStream stream = File.OpenRead(Fixture(name));

        Assert.True(IWorkContainerProbe.HasModernIndex(stream, stream.Length,
            8192, CancellationToken.None));
        Assert.Equal(0, stream.Position);
    }

    [Theory]
    [InlineData("nim-iwork/simple.pages")]
    [InlineData("nim-iwork/simple.numbers")]
    [InlineData("nim-iwork/simple.key")]
    public void Structural_probe_respects_the_current_stream_slice(string name) {
        byte[] package = File.ReadAllBytes(Fixture(name));
        byte[] prefix = { 0x19, 0x27, 0x38, 0x44, 0x55 };
        using var stream = new MemoryStream(prefix.Concat(package).ToArray(), writable: false);
        stream.Position = prefix.Length;

        Assert.True(IWorkContainerProbe.HasModernIndex(stream, package.LongLength,
            8192, CancellationToken.None));
        Assert.Equal(prefix.Length, stream.Position);
    }

    [Fact]
    public void Structural_probe_rejects_a_generic_zip() {
        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true)) {
            using Stream entry = archive.CreateEntry("notes.txt").Open();
            entry.WriteByte(1);
        }
        stream.Position = 0;

        Assert.False(IWorkContainerProbe.HasModernIndex(stream, stream.Length,
            8192, CancellationToken.None));
        Assert.Equal(0, stream.Position);
    }

    [Fact]
    public void Structural_probe_enforces_entry_limit_before_materializing_archive() {
        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true)) {
            using (Stream entry = archive.CreateEntry("metadata.txt").Open()) entry.WriteByte(1);
            using (Stream entry = archive.CreateEntry("Index/Document.iwa").Open()) entry.WriteByte(1);
        }
        stream.Position = 0;

        Assert.False(IWorkContainerProbe.HasModernIndex(stream, stream.Length,
            1, CancellationToken.None));
        Assert.True(IWorkContainerProbe.HasModernIndex(stream, stream.Length,
            2, CancellationToken.None));
        Assert.Equal(0, stream.Position);

        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        Assert.Throws<OperationCanceledException>(() =>
            IWorkContainerProbe.HasModernIndex(stream, stream.Length, 2,
                cancelled.Token));
        Assert.Equal(0, stream.Position);
    }

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
        var catalog = CreateRichCatalog(records, options);
        Assert.Empty(catalog.Materialized);
        catalog.TryRead(1, out _);
        var strings = catalog.Materialized;
        bool complete = catalog.FullyReconstructed;

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
        var catalog = CreateRichCatalog(records, options);
        Assert.Empty(catalog.Materialized);
        catalog.TryRead(1, out _);
        var strings = catalog.Materialized;
        bool complete = catalog.FullyReconstructed;

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
        var catalog = CreateRichCatalog(records, options);
        Assert.Empty(catalog.Materialized);
        catalog.TryRead(1, out _);
        var strings = catalog.Materialized;
        bool complete = catalog.FullyReconstructed;

        Assert.True(complete);
        Assert.Equal("Value", strings[1].PlainText);
        Assert.True(strings[1].IsTextComplete);
        Assert.False(strings[1].IsComplete);

        static IWorkArchiveRecord Record(ulong id, uint type, byte[] payload) =>
            new(id, type, Array.Empty<uint>(), Array.Empty<ulong>(),
                Array.Empty<ulong>(), "Index/Tables/Test.iwa", 0, payload);
    }

    [Fact]
    public void Aliased_rich_catalog_entries_share_decoding_and_retain_partial_text_status() {
        var options = new IWorkReadOptions();
        var records = new[] {
            Record(1, 6005, Message(
                BytesField(3, Message(VarintField(1, 1), ReferenceField(9, 2))),
                BytesField(3, Message(VarintField(1, 2), ReferenceField(9, 2))))),
            Record(2, 6218, Message(ReferenceField(1, 3))),
            Record(3, 2001, Message(StringField(3, "Before\uFFFCafter")))
        };
        var catalog = CreateRichCatalog(records, options);
        Assert.Empty(catalog.Materialized);
        catalog.TryRead(1, out _);
        catalog.TryRead(2, out _);
        var strings = catalog.Materialized;
        bool complete = catalog.FullyReconstructed;

        Assert.False(complete);
        Assert.Equal("Beforeafter", strings[1].PlainText);
        Assert.False(strings[1].IsTextComplete);
        Assert.Same(strings[1], strings[2]);

        static IWorkArchiveRecord Record(ulong id, uint type, byte[] payload) =>
            new(id, type, Array.Empty<uint>(), Array.Empty<ulong>(),
                Array.Empty<ulong>(), "Index/Tables/Test.iwa", 0, payload);
    }

    [Fact]
    public void Unused_rich_catalog_aliases_do_not_consume_cell_text_budget() {
        var options = new IWorkReadOptions { MaximumProjectedTextItems = 2 };
        var records = new[] {
            Record(1, 6005, Message(
                BytesField(3, Message(VarintField(1, 1), ReferenceField(9, 2))),
                BytesField(3, Message(VarintField(1, 2), ReferenceField(9, 2))))),
            Record(2, 6218, Message(ReferenceField(1, 3))),
            Record(3, 2001, Message(StringField(3, "Value")))
        };
        var catalog = CreateRichCatalog(records, options);
        Assert.Empty(catalog.Materialized);
        catalog.TryRead(1, out _);
        var strings = catalog.Materialized;
        bool complete = catalog.FullyReconstructed;

        Assert.True(complete);
        Assert.Equal("Value", Assert.Single(strings).Value.PlainText);

        static IWorkArchiveRecord Record(ulong id, uint type, byte[] payload) =>
            new(id, type, Array.Empty<uint>(), Array.Empty<ulong>(),
                Array.Empty<ulong>(), "Index/Tables/Test.iwa", 0, payload);
    }

    private static IWorkTableRichTextCatalog CreateRichCatalog(IWorkArchiveRecord[] records, IWorkReadOptions options) {
        byte[] store = Message(ReferenceField(17, 1));
        using MemoryStream package = CreatePackage(("Index/Document.iwa", FrameIwa(Message(
            ArchiveRecord(100, 1, Message(ReferenceField(1, 101))),
            ArchiveRecord(101, 2, Message()),
            ArchiveRecord(104, 6001, Message(BytesField(4, store))),
            Message(records.Select(record => ArchiveRecord(record.Identifier, record.MessageType,
                record.Payload, record.ObjectReferences.ToArray())).ToArray())))));
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers, options);
        return IWorkTableRichTextCatalog.Create(source, IWorkProtobuf.Parse(store, source.Options),
            source.Index.Find(104)!, new IWorkProjectionBudget(source.Options), new IWorkSourceReferenceIssueCollector(source));
    }
}
