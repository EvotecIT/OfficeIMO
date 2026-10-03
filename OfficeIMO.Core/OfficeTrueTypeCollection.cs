namespace OfficeIMO.Drawing;

internal static class OfficeTrueTypeCollection {
    internal const int MaxTrueTypeCollectionFontsToInspect = 256;
    internal const int MaxExtractedTrueTypeCollectionFontBytes = 64 * 1024 * 1024;
    internal const int MaxExtractedTrueTypeCollectionBytes = 128 * 1024 * 1024;

    internal static System.Collections.Generic.List<byte[]> ExtractPrograms(byte[] data) {
        if (!IsTrueTypeCollection(data)) {
            return HasTrueTypeOutlines(data, 0)
                ? new System.Collections.Generic.List<byte[]> { data }
                : new System.Collections.Generic.List<byte[]>();
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

            // Installed collections can contain CFF faces. The TrueType consumer cannot
            // use them; reject them before copying potentially large shared tables.
            // ExtractFace remains the format-neutral single-face extraction boundary.
            if (!HasTrueTypeOutlines(data, (int)offset)) continue;
            byte[] font = ExtractTrueTypeCollectionFont(data, (int)offset);
            extractedBytes = checked(extractedBytes + font.Length);
            if (extractedBytes > MaxExtractedTrueTypeCollectionBytes) {
                throw new System.NotSupportedException("TrueType collection extracted font data exceeds supported limits.");
            }

            fonts.Add(font);
        }

        return fonts;
    }

    private static bool HasTrueTypeOutlines(byte[] data, int fontOffset) {
        EnsureRange(data, fontOffset, 12);
        uint scaler = ReadUInt32(data, fontOffset);
        ushort tableCount = ReadUInt16(data, fontOffset + 4);
        if (tableCount == 0) throw new System.NotSupportedException("Font has no tables.");
        EnsureRange(data, fontOffset, checked(12 + tableCount * 16));
        bool glyf = false;
        bool loca = false;
        for (int i = 0; i < tableCount; i++) {
            int record = fontOffset + 12 + i * 16;
            uint tag = ReadUInt32(data, record);
            uint offset = ReadUInt32(data, record + 8);
            uint length = ReadUInt32(data, record + 12);
            if (offset > int.MaxValue || length > int.MaxValue)
                throw new System.NotSupportedException("Font table offsets are too large.");
            // Preserve range validation even for an unsupported face; skipping its
            // payload allocation must not turn malformed collection data into a match.
            EnsureRange(data, (int)offset, (int)length);
            glyf |= tag == 0x676c7966;
            loca |= tag == 0x6c6f6361;
        }
        return (scaler == 0x00010000 || scaler == 0x74727565) && glyf && loca;
    }

    internal static bool IsTrueTypeCollection(byte[] data) =>
        data.Length >= 4 &&
        data[0] == (byte)'t' &&
        data[1] == (byte)'t' &&
        data[2] == (byte)'c' &&
        data[3] == (byte)'f';

    private static byte[] ExtractTrueTypeCollectionFont(byte[] collectionData, int fontOffset, int maximumFontBytes = MaxExtractedTrueTypeCollectionFontBytes) {
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

        if (outputOffset > System.Math.Min(maximumFontBytes, MaxExtractedTrueTypeCollectionFontBytes)) {
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

    internal static byte[] ExtractFace(byte[] data, int faceIndex, int maximumFontBytes) {
        if (!IsTrueTypeCollection(data)) {
            if (data.Length > maximumFontBytes) throw new System.NotSupportedException("Font data exceeds the supported byte limit.");
            return (byte[])data.Clone();
        }
        EnsureRange(data, 0, 12);
        uint count = ReadUInt32(data, 8);
        if (count == 0 || count > MaxTrueTypeCollectionFontsToInspect || faceIndex < 0 || faceIndex >= count)
            throw new System.NotSupportedException("TrueType collection face index is invalid.");
        EnsureRange(data, 12, checked((int)count * 4));
        uint offset = ReadUInt32(data, 12 + faceIndex * 4);
        if (offset > int.MaxValue) throw new System.NotSupportedException("TrueType collection font offset is too large.");
        return ExtractTrueTypeCollectionFont(data, (int)offset, maximumFontBytes);
    }

    private static void EnsureRange(byte[] data, int offset, int length) {
        if (offset < 0 || length < 0 || offset > data.Length - length)
            throw new System.NotSupportedException("TrueType collection data range is invalid.");
    }
    private static ushort ReadUInt16(byte[] data, int offset) {
        EnsureRange(data, offset, 2);
        return (ushort)((data[offset] << 8) | data[offset + 1]);
    }
    private static uint ReadUInt32(byte[] data, int offset) {
        EnsureRange(data, offset, 4);
        return ((uint)data[offset] << 24) | ((uint)data[offset + 1] << 16) | ((uint)data[offset + 2] << 8) | data[offset + 3];
    }
}
