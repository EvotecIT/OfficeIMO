namespace OfficeIMO.OneNote;

internal static partial class OneNoteSemanticMapper {
    private const uint OutlineElementChildLevel = 0x0C001C03;
    private const uint NumberListFormat = 0x1C001C1A;
    private const uint ListFont = 0x1C001C52;
    private const uint ListRestart = 0x14001CB7;

    private static OneNoteListInfo? BuildListInfo(OneNoteMaterializedObjectSpace space, OneNoteRevisionStoreObject outlineElement) {
        OneNoteRevisionStoreObject? listNode = GetReferences(outlineElement, ListNodes)
            .Select(space.GetObject)
            .LastOrDefault(item => item?.Jcid.Value == JcidNumberListNode);
        if (listNode == null) return null;
        int level = Math.Max(0, (ReadByte(outlineElement, OutlineElementChildLevel) ?? 1) - 1);
        return MapListInfo(listNode, level);
    }

    internal static OneNoteListInfo MapListInfo(OneNoteRevisionStoreObject listNode, int level) {
        if (listNode == null) throw new ArgumentNullException(nameof(listNode));
        if (level < 0) throw new ArgumentOutOfRangeException(nameof(level));

        string format = ReadNumberListFormat(listNode);
        int marker = format.IndexOf('\uFFFD');
        uint? numberingFormat = marker >= 0 && marker + 1 < format.Length ? format[marker + 1] : (uint?)null;
        uint? restart = ReadUInt32(listNode, ListRestart);
        return new OneNoteListInfo {
            ObjectId = listNode.Id,
            Ordered = marker >= 0,
            Format = numberingFormat,
            Level = level,
            Restart = restart.HasValue,
            DisplayIndex = restart.HasValue && restart.Value <= int.MaxValue ? (int)restart.Value : (int?)null,
            FontFamily = ReadString(listNode, ListFont) ?? ReadString(listNode, Font)
        };
    }

    private static string ReadNumberListFormat(OneNoteRevisionStoreObject listNode) {
        byte[]? data = ReadData(listNode, NumberListFormat);
        if (data == null || data.Length < 2) return string.Empty;
        string value = System.Text.Encoding.Unicode.GetString(data, 0, data.Length - data.Length % 2);
        if (value.Length == 0) return string.Empty;
        int count = value[0];
        return value.Substring(1, Math.Min(count, value.Length - 1));
    }

    private static byte? ReadByte(OneNoteRevisionStoreObject? item, uint propertyId) {
        ulong? value = FindProperty(item?.PropertySet, propertyId)?.ScalarValue;
        return value.HasValue ? (byte)value.Value : null;
    }
}
