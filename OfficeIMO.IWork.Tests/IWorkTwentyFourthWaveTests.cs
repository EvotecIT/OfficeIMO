using OfficeIMO.IWork;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Pages_text_boxes_follow_recovered_z_order() {
        using MemoryStream package = CreatePagesPackageWithRestackedTextBoxes();

        using var result = WordIWorkConverter.ConvertPagesToWordResult(package);

        Assert.True(result.Projection.HasEditableContent,
            string.Join("; ", result.Projection.Diagnostics.Select(diagnostic =>
                $"{diagnostic.Code}: {diagnostic.Message}")));
        Assert.False(result.IsVisualFallback);
        Assert.Equal(new[] { "Higher identifier", "Lower identifier" },
            result.Projection.TextBoxObjects.Select(textBox => textBox.Content.PlainText));
        Assert.Equal(new[] { "Higher identifier", "Lower identifier" },
            result.Value.TextBoxes.Select(textBox => textBox.Paragraphs[0].Text.TrimEnd('\n')));

        using var saved = new MemoryStream();
        result.Value.Save(saved);
        saved.Position = 0;
        using WordDocument reopened = WordDocument.Load(saved);
        Assert.Equal(new[] { "Higher identifier", "Lower identifier" },
            reopened.TextBoxes.Select(textBox => textBox.Paragraphs[0].Text.TrimEnd('\n')));
    }

    private static MemoryStream CreatePagesPackageWithRestackedTextBoxes() {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        const ulong zOrderId = 3;
        const ulong lowerShapeId = 10;
        const ulong higherShapeId = 30;
        const ulong lowerStorageId = 11;
        const ulong higherStorageId = 31;
        byte[] geometry = Message(
            BytesField(1, Message(FloatField(1, 36f), FloatField(2, 72f))),
            BytesField(2, Message(FloatField(1, 216f), FloatField(2, 108f))));
        byte[] drawable = Message(BytesField(1, geometry));
        byte[] lowerShape = Message(BytesField(1, Message(BytesField(1, drawable))),
            ReferenceField(2, lowerStorageId));
        byte[] higherShape = Message(BytesField(1, Message(BytesField(1, drawable))),
            ReferenceField(2, higherStorageId));
        byte[] records = Message(
            ArchiveRecord(documentId, 10000,
                Message(ReferenceField(4, bodyId), ReferenceField(20, zOrderId)),
                new[] { bodyId, zOrderId }),
            ArchiveRecord(bodyId, 2001, Message(StringField(3, "Body"))),
            ArchiveRecord(zOrderId, 10020,
                Message(ReferenceField(1, higherShapeId), ReferenceField(1, lowerShapeId)),
                new[] { higherShapeId, lowerShapeId }),
            ArchiveRecord(lowerShapeId, 2011, lowerShape, new[] { lowerStorageId }),
            ArchiveRecord(lowerStorageId, 2001, Message(StringField(3, "Lower identifier"))),
            ArchiveRecord(higherShapeId, 2011, higherShape, new[] { higherStorageId }),
            ArchiveRecord(higherStorageId, 2001, Message(StringField(3, "Higher identifier"))));
        return CreatePackage(("Index/Document.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));
    }
}
