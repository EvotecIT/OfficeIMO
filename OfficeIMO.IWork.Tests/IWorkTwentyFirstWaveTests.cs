using OfficeIMO.IWork;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Unsupported_pages_z_order_drawables_use_visual_fallback() {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        const ulong zOrderId = 3;
        const ulong unsupportedDrawableId = 4;
        byte[] records = Message(
            ArchiveRecord(documentId, 10000,
                Message(ReferenceField(4, bodyId), ReferenceField(20, zOrderId)),
                new[] { bodyId, zOrderId }),
            ArchiveRecord(bodyId, 2001, Message(StringField(3, "Body"))),
            ArchiveRecord(zOrderId, 10020,
                Message(ReferenceField(1, unsupportedDrawableId)),
                new[] { unsupportedDrawableId }),
            ArchiveRecord(unsupportedDrawableId, 9999, Message()));
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));

        using var result = WordIWorkConverter.LoadPagesWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.False(result.Projection.HasEditableContent);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_PAGES_DRAWABLE_UNSUPPORTED"
            && diagnostic.RecordIdentifier == unsupportedDrawableId);
    }

    [Theory]
    [InlineData(24f, 0f, 0f, 0f, 0f)]
    [InlineData(0f, 0f, 240f, 120f, 0f)]
    [InlineData(0f, 0f, 0f, 0f, 15f)]
    public void Positioned_sized_or_rotated_pages_tables_use_visual_fallback(
        float left, float top, float width, float height, float rotation) {
        using MemoryStream package = CreatePagesPackageWithTableGeometry(
            left, top, width, height, rotation, includePreview: true);

        using var result = WordIWorkConverter.LoadPagesWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.True(result.Projection.HasEditableContent);
    }

    [Fact]
    public void Zero_pages_table_geometry_remains_inline_editable() {
        using MemoryStream package = CreatePagesPackageWithTableGeometry(
            0f, 0f, 0f, 0f, 0f, includePreview: false);

        using var result = WordIWorkConverter.LoadPagesWithReport(package);

        Assert.False(result.IsVisualFallback);
        Assert.Single(result.Projection.Tables);
        Assert.Single(result.Document.Tables);
    }

    private static MemoryStream CreatePagesPackageWithTableGeometry(
        float left, float top, float width, float height, float rotation,
        bool includePreview, double? defaultRowHeight = null,
        double? defaultColumnWidth = null, int rows = 1, int columns = 1,
        string? accessibilityDescription = null) {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        const ulong tableId = 10;
        const ulong modelId = 11;
        byte[] geometry = Message(
            BytesField(1, Message(FloatField(1, left), FloatField(2, top))),
            BytesField(2, Message(FloatField(1, width), FloatField(2, height))),
            FloatField(4, rotation));
        byte[] drawable = Message(BytesField(1, geometry),
            accessibilityDescription == null
                ? Array.Empty<byte>()
                : StringField(8, accessibilityDescription));
        var modelFields = new List<byte[]> {
            BytesField(4, Message(BytesField(3, Message()))),
            VarintField(6, checked((ulong)rows)), VarintField(7, checked((ulong)columns)),
            StringField(8, "Table")
        };
        if (defaultRowHeight.HasValue) modelFields.Add(DoubleField(16, defaultRowHeight.Value));
        if (defaultColumnWidth.HasValue) modelFields.Add(DoubleField(17, defaultColumnWidth.Value));
        byte[] model = Message(modelFields.ToArray());
        byte[] records = Message(
            ArchiveRecord(documentId, 10000,
                Message(ReferenceField(4, bodyId)), new[] { bodyId, tableId }),
            ArchiveRecord(bodyId, 2001, Message(StringField(3, "Body"))),
            ArchiveRecord(tableId, 6000,
                Message(BytesField(1, drawable), ReferenceField(2, modelId)),
                new[] { modelId }),
            ArchiveRecord(modelId, 6001, model));
        return includePreview
            ? CreatePackage(
                ("Index/Document.iwa", FrameIwa(records)),
                ("preview.png", ValidPreviewPng()))
            : CreatePackage(("Index/Document.iwa", FrameIwa(records)));
    }
}
