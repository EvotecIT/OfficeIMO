namespace OfficeIMO.Pdf;

public sealed partial class PdfEmbeddedFontFamily {
    internal const int MaxTrueTypeCollectionFontsToInspect = 256;
    internal const int MaxExtractedTrueTypeCollectionFontBytes = 64 * 1024 * 1024;
    internal const int MaxExtractedTrueTypeCollectionBytes = 128 * 1024 * 1024;

    private static System.Collections.Generic.List<byte[]> ExtractTrueTypeFontPrograms(byte[] data) {
        if (!IsTrueTypeCollection(data)) {
            return new System.Collections.Generic.List<byte[]> { data };
        }

        EnsureRange(data, 0, 12);
        uint fontCount = ReadUInt32(data, 8);
        if (fontCount == 0 || fontCount > MaxTrueTypeCollectionFontsToInspect) {
            throw new System.NotSupportedException("TrueType collection font count is invalid.");
        }

        EnsureRange(data, 12, checked((int)fontCount * 4));
        var fonts = new System.Collections.Generic.List<byte[]>((int)fontCount);
        int extractedBytes = 0;
        for (int i = 0; i < fontCount; i++) {
            uint offset = ReadUInt32(data, 12 + i * 4);
            if (offset > int.MaxValue) {
                throw new System.NotSupportedException("TrueType collection font offset is too large.");
            }

            byte[] font = ExtractTrueTypeCollectionFont(data, (int)offset);
            extractedBytes = checked(extractedBytes + font.Length);
            if (extractedBytes > MaxExtractedTrueTypeCollectionBytes) {
                throw new System.NotSupportedException("TrueType collection extracted font data exceeds supported limits.");
            }

            fonts.Add(font);
        }

        return fonts;
    }

    private static bool IsTrueTypeCollection(byte[] data) =>
        data.Length >= 4 &&
        data[0] == (byte)'t' &&
        data[1] == (byte)'t' &&
        data[2] == (byte)'c' &&
        data[3] == (byte)'f';

    private static byte[] ExtractTrueTypeCollectionFont(byte[] collectionData, int fontOffset) {
        EnsureRange(collectionData, fontOffset, 12);
        ushort tableCount = ReadUInt16(collectionData, fontOffset + 4);
        if (tableCount == 0) {
            throw new System.NotSupportedException("TrueType collection font has no tables.");
        }

        int directoryLength = checked(12 + tableCount * 16);
        EnsureRange(collectionData, fontOffset, directoryLength);
        var tables = new System.Collections.Generic.List<FontTableCopyRecord>(tableCount);
        int outputOffset = Align4(directoryLength);
        for (int i = 0; i < tableCount; i++) {
            int recordOffset = fontOffset + 12 + i * 16;
            string tag = System.Text.Encoding.ASCII.GetString(collectionData, recordOffset, 4);
            uint checksum = ReadUInt32(collectionData, recordOffset + 4);
            uint sourceOffset = ReadUInt32(collectionData, recordOffset + 8);
            uint length = ReadUInt32(collectionData, recordOffset + 12);
            if (sourceOffset > int.MaxValue || length > int.MaxValue) {
                throw new System.NotSupportedException("TrueType collection table offsets are too large.");
            }

            EnsureRange(collectionData, (int)sourceOffset, (int)length);
            tables.Add(new FontTableCopyRecord(tag, checksum, (int)sourceOffset, (int)length, outputOffset));
            outputOffset = Align4(checked(outputOffset + (int)length));
        }

        if (outputOffset > MaxExtractedTrueTypeCollectionFontBytes) {
            throw new System.NotSupportedException("TrueType collection font data exceeds supported limits.");
        }

        byte[] fontData = new byte[outputOffset];
        System.Array.Copy(collectionData, fontOffset, fontData, 0, 4);
        WriteUInt16(fontData, 4, tableCount);
        WriteSearchParameters(fontData, tableCount);
        for (int i = 0; i < tables.Count; i++) {
            FontTableCopyRecord table = tables[i];
            int targetRecordOffset = 12 + i * 16;
            byte[] tagBytes = System.Text.Encoding.ASCII.GetBytes(table.Tag);
            System.Array.Copy(tagBytes, 0, fontData, targetRecordOffset, 4);
            WriteUInt32(fontData, targetRecordOffset + 4, table.Checksum);
            WriteUInt32(fontData, targetRecordOffset + 8, (uint)table.TargetOffset);
            WriteUInt32(fontData, targetRecordOffset + 12, (uint)table.Length);
            System.Array.Copy(collectionData, table.SourceOffset, fontData, table.TargetOffset, table.Length);
        }

        return fontData;
    }

    private static int Align4(int value) => checked((value + 3) & ~3);

    private static void WriteSearchParameters(byte[] data, int tableCount) {
        int maxPowerOfTwo = 1;
        int entrySelector = 0;
        while (maxPowerOfTwo * 2 <= tableCount) {
            maxPowerOfTwo *= 2;
            entrySelector++;
        }

        WriteUInt16(data, 6, (ushort)(maxPowerOfTwo * 16));
        WriteUInt16(data, 8, (ushort)entrySelector);
        WriteUInt16(data, 10, (ushort)((tableCount - maxPowerOfTwo) * 16));
    }

    private static void WriteUInt16(byte[] data, int offset, ushort value) {
        EnsureRange(data, offset, 2);
        data[offset] = (byte)(value >> 8);
        data[offset + 1] = (byte)value;
    }

    private static void WriteUInt32(byte[] data, int offset, uint value) {
        EnsureRange(data, offset, 4);
        data[offset] = (byte)(value >> 24);
        data[offset + 1] = (byte)(value >> 16);
        data[offset + 2] = (byte)(value >> 8);
        data[offset + 3] = (byte)value;
    }

    private readonly struct FontTableCopyRecord {
        public FontTableCopyRecord(string tag, uint checksum, int sourceOffset, int length, int targetOffset) {
            Tag = tag;
            Checksum = checksum;
            SourceOffset = sourceOffset;
            Length = length;
            TargetOffset = targetOffset;
        }

        public string Tag { get; }

        public uint Checksum { get; }

        public int SourceOffset { get; }

        public int Length { get; }

        public int TargetOffset { get; }
    }
}
