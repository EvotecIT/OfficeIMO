using System.Text;

namespace OfficeIMO.OneNote.Tests;

public sealed class SemanticPropertyTests {
    [Fact]
    public void TextStyleMapsTrueBooleanHyperlinkAndProtectedUrl() {
        OneNoteRevisionStoreObject style = CreateObject(
            0x0012004D,
            Property(0x88001E14, booleanValue: true),
            Property(0x88001E19, booleanValue: true),
            Property(0x88003401, booleanValue: true),
            Property(0x1C001E20, data: Unicode("https://example.test/path")));
        var run = new OneNoteTextRun { Text = "Example" };

        OneNoteSemanticMapper.ApplyTextStyle(run, style);

        Assert.Equal("https://example.test/path", run.Hyperlink);
        Assert.True(run.HyperlinkProtected);
        Assert.True(run.Style.IsMath);
    }

    [Fact]
    public void NumberListMapsOrderedMarkerRestartAndFont() {
        string encodedFormat = new string(new[] { (char)3, '\uFFFD', (char)0, '.' });
        OneNoteRevisionStoreObject listNode = CreateObject(
            0x00060012,
            Property(0x1C001C1A, data: Encoding.Unicode.GetBytes(encodedFormat)),
            Property(0x1C001C52, data: Unicode("Calibri")),
            Property(0x14001CB7, scalarValue: 4));

        OneNoteListInfo list = OneNoteSemanticMapper.MapListInfo(listNode, 2);

        Assert.True(list.Ordered);
        Assert.Equal(0U, list.Format);
        Assert.Equal(2, list.Level);
        Assert.True(list.Restart);
        Assert.Equal(4, list.DisplayIndex);
        Assert.Equal("Calibri", list.FontFamily);
    }

    [Fact]
    public void TaskTagMapsStatusShapeAndTime32Fields() {
        OneNotePropertySet state = PropertySet(
            Property(0x10003463, scalarValue: 105),
            Property(0x10003464, scalarValue: 89),
            Property(0x10003470, scalarValue: 0x1D),
            Property(0x1400346B, scalarValue: 86_400),
            Property(0x1400346E, scalarValue: 60),
            Property(0x1400346F, scalarValue: 120));

        OneNoteTag tag = OneNoteSemanticMapper.MapTagState(state, _ => null);

        Assert.True(tag.IsTask);
        Assert.True(tag.IsCheckable);
        Assert.True(tag.IsCompleted);
        Assert.False(tag.IsDisabled);
        Assert.True(tag.IsUnsynchronized);
        Assert.True(tag.IsRemoved);
        Assert.Equal(105U, tag.ActionItemType);
        Assert.Equal(89U, tag.Shape);
        Assert.Equal(new DateTime(1980, 1, 2, 0, 0, 0, DateTimeKind.Utc), tag.DueUtc);
        Assert.Equal(new DateTime(1980, 1, 1, 0, 1, 0, DateTimeKind.Utc), tag.CreatedUtc);
        Assert.Equal(new DateTime(1980, 1, 1, 0, 2, 0, DateTimeKind.Utc), tag.CompletedUtc);
    }

    [Fact]
    public void NormalTagResolvesSharedDefinitionProperties() {
        OneNoteExtendedGuid definitionId = new OneNoteExtendedGuid(Guid.NewGuid(), 7, 4);
        OneNoteRevisionStoreObject definition = CreateObject(
            0x00120043,
            Property(0x10003463, scalarValue: 8),
            Property(0x10003464, scalarValue: 1),
            Property(0x1C003468, data: Unicode("Important")),
            Property(0x14003465, scalarValue: 0x0000FFFF),
            Property(0x14003466, scalarValue: 0x000000FF));
        OneNotePropertySet state = PropertySet(
            Property(0x10003470, scalarValue: 0),
            Property(0x1400346E, scalarValue: 60),
            Property(0x1400346F, scalarValue: 0),
            Property(0x20003488, referencedIds: new[] { definitionId }));

        OneNoteTag tag = OneNoteSemanticMapper.MapTagState(state, id => id.Equals(definitionId) ? definition : null);

        Assert.False(tag.IsTask);
        Assert.Equal(definitionId, tag.DefinitionId);
        Assert.Equal(8U, tag.ActionItemType);
        Assert.Equal("Important", tag.Label);
        Assert.True(tag.IsCheckable);
        Assert.False(tag.IsCompleted);
        Assert.Equal(0x0000FFFFU, tag.HighlightColorArgb);
        Assert.Equal(0x000000FFU, tag.TextColorArgb);
        Assert.Null(tag.CompletedUtc);
    }

    [Theory]
    [InlineData(1U, "meeting.mp3", OneNoteMediaKind.Audio, "audio/mpeg")]
    [InlineData(2U, "demo.wmv", OneNoteMediaKind.Video, "video/x-ms-wmv")]
    public void EmbeddedRecordingMapsToTypedMedia(uint recordMedia, string fileName, OneNoteMediaKind expectedKind, string expectedMediaType) {
        OneNoteRevisionStoreObject source = CreateObject(
            0x00060035,
            Property(0x14001D24, scalarValue: recordMedia),
            Property(0x1C001D9C, data: Unicode(fileName)),
            Property(0x1C001D9D, data: Unicode("C:\\recordings\\" + fileName)));

        OneNoteMedia media = Assert.IsType<OneNoteMedia>(OneNoteSemanticMapper.CreateEmbeddedElement(source));

        Assert.Equal(expectedKind, media.RecordingKind);
        Assert.Equal(fileName, media.FileName);
        Assert.Equal(expectedMediaType, media.MediaType);
        Assert.EndsWith(fileName, media.SourcePath, StringComparison.Ordinal);
    }

    [Fact]
    public void OpaqueObjectRetainsRawPropertyStreamAndDecodedReferences() {
        OneNoteExtendedGuid reference = new OneNoteExtendedGuid(Guid.NewGuid(), 9, 4);
        OneNoteRevisionStoreObject source = CreateObject(
            0x0006FFFF,
            Property(0x88001234, booleanValue: true),
            Property(0x24005678, referencedIds: new[] { reference }));
        source.RawPropertyData = OneNoteBinaryPayload.FromBytes(new byte[] { 1, 2, 3, 4 });

        OneNoteOpaqueObject opaque = OneNoteSemanticMapper.CreateOpaqueObject(source, 7);

        Assert.Equal(source.Id, opaque.Id);
        Assert.Equal(0x0006FFFFU, opaque.Jcid);
        Assert.Equal(7, opaque.Ordinal);
        Assert.Equal(new byte[] { 1, 2, 3, 4 }, opaque.GetRawData());
        Assert.True(opaque.Properties[0].BooleanValue);
        Assert.Equal(reference, Assert.Single(opaque.Properties[1].ReferencedIds));
    }

    private static OneNoteRevisionStoreObject CreateObject(uint jcid, params OneNotePropertyValue[] properties) {
        var node = new OneNoteFileNode(0, 4, 0, 0, OneNoteFileNodeBaseType.Inline, 0, null, Array.Empty<byte>());
        var value = new OneNoteRevisionStoreObject(
            new OneNoteExtendedGuid(Guid.NewGuid(), 1, 4),
            new OneNoteJcid(jcid),
            node) {
            PropertySet = PropertySet(properties)
        };
        return value;
    }

    private static OneNotePropertySet PropertySet(params OneNotePropertyValue[] properties) {
        return new OneNotePropertySet(properties, 0);
    }

    private static OneNotePropertyValue Property(
        uint rawId,
        bool? booleanValue = null,
        ulong? scalarValue = null,
        byte[]? data = null,
        IReadOnlyList<OneNoteExtendedGuid>? referencedIds = null) {
        var property = new OneNotePropertyValue(rawId, 0) {
            BooleanValue = booleanValue,
            ScalarValue = scalarValue,
            ReferencedIds = referencedIds ?? Array.Empty<OneNoteExtendedGuid>()
        };
        if (data != null) property.Data = OneNoteBinaryPayload.FromBytes(data);
        return property;
    }

    private static byte[] Unicode(string value) => Encoding.Unicode.GetBytes(value + "\0");
}
