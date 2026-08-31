using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Failed_semantic_image_decodes_consume_the_shared_budget() {
        using MemoryStream package = CreatePagesImagePackage(duplicateMetadata: false,
            imageCount: 1, imageBytes: CreateInvalidAdlerPng(20, 20));
        var options = new IWorkReadOptions { MaximumPackageBytes = 700 };
        IWorkSourceDocument source = IWorkSourceDocument.Open(
            package, IWorkDocumentKind.Pages, options);
        IWorkArchiveRecord image = Assert.Single(source.Records,
            record => record.MessageType == 3005);
        var budget = new IWorkProjectionBudget(options);

        IWorkImageAsset? asset = IWorkDrawingReader.ReadImage(
            source, image, budget, out bool complete);

        Assert.Null(asset);
        Assert.False(complete);
        Assert.InRange(budget.RemainingDecodedImageBytes, 0, 699);
    }

    [Fact]
    public void Repeated_Keynote_title_placeholders_disable_editable_reconstruction() {
        byte[] records = Message(
            ArchiveRecord(1, 1, Message(ReferenceField(2, 2))),
            ArchiveRecord(2, 2, Message(BytesField(3, Message(ReferenceField(2, 3))))),
            ArchiveRecord(3, 4, Message(ReferenceField(2, 4))),
            ArchiveRecord(4, 5, Message(ReferenceField(5, 5), ReferenceField(5, 7))),
            ArchiveRecord(5, 2011, Message(ReferenceField(2, 6))),
            ArchiveRecord(6, 2001, Message(StringField(3, "First"))),
            ArchiveRecord(7, 2011, Message(ReferenceField(2, 8))),
            ArchiveRecord(8, 2001, Message(StringField(3, "Second"))));
        using MemoryStream package = CreatePackage(
            ("Index/Slide.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));

        using var result = PowerPointPresentation.LoadKeynoteWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_KEYNOTE_DRAWABLE_UNSUPPORTED");
    }

    [Theory]
    [InlineData(7)]
    [InlineData(11)]
    public void Keynote_indents_must_round_trip_through_owner_emus(int indentField) {
        const ulong styleId = 7;
        byte[] styleTable = Message(BytesField(1,
            Message(VarintField(1, 0), ReferenceField(2, styleId))));
        byte[] records = Message(
            ArchiveRecord(1, 1, Message(ReferenceField(2, 2))),
            ArchiveRecord(2, 2, Message(BytesField(3, Message(ReferenceField(2, 3))))),
            ArchiveRecord(3, 4, Message(ReferenceField(2, 4))),
            ArchiveRecord(4, 5, Message(ReferenceField(5, 5))),
            ArchiveRecord(5, 2011, Message(ReferenceField(2, 6))),
            ArchiveRecord(6, 2001,
                Message(StringField(3, "Indented"), BytesField(5, styleTable)), new[] { styleId }),
            ArchiveRecord(styleId, 2022,
                Message(BytesField(12, Message(FloatField(indentField, 0.00001f))))));
        using MemoryStream package = CreatePackage(
            ("Index/Slide.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));

        using var result = PowerPointPresentation.LoadKeynoteWithReport(package);

        Assert.True(result.Projection.HasEditableContent);
        Assert.True(result.IsVisualFallback);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(3)]
    public void Repeated_MessageInfo_scalars_are_rejected(int field) {
        byte[] payload = Message(VarintField(1, 42));
        byte[] messageInfo = field == 1
            ? Message(VarintField(1, 9999), VarintField(1, 1),
                VarintField(3, checked((ulong)payload.Length)))
            : Message(VarintField(1, 1), VarintField(3, 0),
                VarintField(3, checked((ulong)payload.Length)));
        byte[] archiveInfo = Message(VarintField(1, 1), BytesField(2, messageInfo));
        byte[] record = Message(Varint(checked((ulong)archiveInfo.Length)), archiveInfo, payload);
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(record)));

        Assert.Throws<InvalidDataException>(() =>
            IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers));
    }

    private static byte[] CreateInvalidAdlerPng(int width, int height) {
        byte[] bytes = CreateSizedPreviewPng(width, height);
        int offset = 8;
        while (offset <= bytes.Length - 12) {
            int dataLength = checked((int)((uint)bytes[offset] << 24
                | (uint)bytes[offset + 1] << 16
                | (uint)bytes[offset + 2] << 8
                | bytes[offset + 3]));
            int typeOffset = offset + 4;
            int dataOffset = typeOffset + 4;
            int crcOffset = dataOffset + dataLength;
            if (bytes[typeOffset] == (byte)'I' && bytes[typeOffset + 1] == (byte)'D'
                && bytes[typeOffset + 2] == (byte)'A' && bytes[typeOffset + 3] == (byte)'T') {
                bytes[dataOffset + dataLength - 1] ^= 1;
                WriteBigEndian32(bytes, crcOffset,
                    unchecked((int)CalculatePngCrc(bytes, typeOffset, 4 + dataLength)));
                return bytes;
            }
            offset = crcOffset + 4;
        }
        throw new InvalidDataException("The generated PNG has no IDAT chunk.");
    }
}
