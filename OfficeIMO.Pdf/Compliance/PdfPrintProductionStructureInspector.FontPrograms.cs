namespace OfficeIMO.Pdf;

internal static partial class PdfPrintProductionStructureInspector {
    private static bool IsValidFontProgram(
        string key,
        PdfStream stream,
        byte[] decoded,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth) {
        if (decoded.Length == 0) return false;
        if (string.Equals(key, "FontFile", StringComparison.Ordinal)) {
            return IsValidType1Program(decoded);
        }
        if (string.Equals(key, "FontFile2", StringComparison.Ordinal)) {
            return IsValidSfntProgram(decoded, requireTrueTypeOutlines: true);
        }
        if (!string.Equals(key, "FontFile3", StringComparison.Ordinal)) return false;

        string? subtype = ResolveName(
            stream.Dictionary.Items.TryGetValue("Subtype", out PdfObject? subtypeObject)
                ? subtypeObject
                : null,
            objects,
            maximumObjectDepth);
        if (string.Equals(subtype, "Type1C", StringComparison.Ordinal) ||
            string.Equals(subtype, "CIDFontType0C", StringComparison.Ordinal)) {
            return IsValidCff1Program(decoded);
        }
        if (string.Equals(subtype, "OpenType", StringComparison.Ordinal)) {
            return IsValidSfntProgram(decoded, requireTrueTypeOutlines: false);
        }
        return false;
    }

    private static bool IsValidType1Program(byte[] data) {
        if (data.Length >= 10 && data[0] == 0x80 && data[1] == 0x01) {
            uint segmentLength = ReadUInt32LittleEndian(data, 2);
            if (segmentLength < 2 || segmentLength > data.Length - 6) return false;
            return data[6] == (byte)'%' && data[7] == (byte)'!' && ContainsAscii(data, "eexec");
        }

        return (StartsWithAscii(data, "%!PS-AdobeFont") || StartsWithAscii(data, "%!FontType1")) &&
            ContainsAscii(data, "eexec");
    }

    private static bool IsValidSfntProgram(byte[] data, bool requireTrueTypeOutlines) {
        if (data.Length < 12) return false;
        uint scaler = ReadUInt32BigEndian(data, 0);
        bool isTrueType = scaler == 0x00010000U || scaler == 0x74727565U;
        bool isOpenTypeCff = scaler == 0x4F54544FU;
        if (requireTrueTypeOutlines ? !isTrueType : (!isTrueType && !isOpenTypeCff)) return false;

        int tableCount = ReadUInt16BigEndian(data, 4);
        if (tableCount <= 0 || tableCount > 4096 || 12L + tableCount * 16L > data.Length) return false;
        int directoryEnd = 12 + tableCount * 16;
        bool hasHead = false;
        bool hasMaxp = false;
        bool hasGlyf = false;
        bool hasLoca = false;
        bool hasCff = false;
        bool cffIsVersion2 = false;
        int cffOffset = 0;
        int cffLength = 0;
        for (int index = 0; index < tableCount; index++) {
            int recordOffset = 12 + index * 16;
            uint tableOffset = ReadUInt32BigEndian(data, recordOffset + 8);
            uint tableLength = ReadUInt32BigEndian(data, recordOffset + 12);
            if (tableOffset > data.Length || tableLength > data.Length - tableOffset ||
                (tableLength > 0 && tableOffset < directoryEnd)) return false;
            string tag = new(new[] {
                (char)data[recordOffset],
                (char)data[recordOffset + 1],
                (char)data[recordOffset + 2],
                (char)data[recordOffset + 3]
            });
            switch (tag) {
                case "head": hasHead = tableLength >= 54; break;
                case "maxp": hasMaxp = tableLength >= 6; break;
                case "glyf": hasGlyf = tableLength > 0; break;
                case "loca": hasLoca = tableLength >= 4; break;
                case "CFF ":
                case "CFF2":
                    hasCff = tableLength >= 4;
                    cffIsVersion2 = string.Equals(tag, "CFF2", StringComparison.Ordinal);
                    cffOffset = (int)tableOffset;
                    cffLength = (int)tableLength;
                    break;
            }
        }

        if (!hasHead || !hasMaxp) return false;
        if (isTrueType) return hasGlyf && hasLoca;
        if (!hasCff) return false;
        var cff = new byte[cffLength];
        Buffer.BlockCopy(data, cffOffset, cff, 0, cffLength);
        return cffIsVersion2 ? IsValidCff2Program(cff) : IsValidCff1Program(cff);
    }

    private static bool IsValidCff1Program(byte[] data) {
        if (data.Length < 4 || data[0] != 1 || data[2] < 4 || data[2] > data.Length ||
            data[3] < 1 || data[3] > 4) return false;
        int cursor = data[2];
        if (!TryReadCffIndex(data, cursor, out cursor, out int nameCount) || nameCount <= 0) return false;
        if (!TryReadCffIndex(data, cursor, out cursor, out int topDictionaryCount) ||
            topDictionaryCount != nameCount) return false;
        if (!TryReadCffIndex(data, cursor, out cursor, out _)) return false;
        return TryReadCffIndex(data, cursor, out _, out _);
    }

    private static bool IsValidCff2Program(byte[] data) {
        if (data.Length < 5 || data[0] != 2 || data[2] < 5 || data[2] > data.Length) return false;
        int topDictionaryLength = ReadUInt16BigEndian(data, 3);
        int globalSubrOffset = data[2] + topDictionaryLength;
        return topDictionaryLength > 0 && globalSubrOffset <= data.Length &&
            TryReadCffIndex(data, globalSubrOffset, out _, out _);
    }

    private static bool TryReadCffIndex(byte[] data, int offset, out int nextOffset, out int count) {
        nextOffset = offset;
        count = 0;
        if (offset < 0 || offset > data.Length - 2) return false;
        count = ReadUInt16BigEndian(data, offset);
        if (count == 0) {
            nextOffset = offset + 2;
            return true;
        }
        if (offset > data.Length - 3) return false;
        int offsetSize = data[offset + 2];
        if (offsetSize < 1 || offsetSize > 4) return false;
        long offsetsStart = offset + 3L;
        long dataStart = offsetsStart + (count + 1L) * offsetSize;
        if (dataStart > data.Length) return false;
        uint first = ReadCffOffset(data, (int)offsetsStart, offsetSize);
        uint last = ReadCffOffset(data, (int)(offsetsStart + count * (long)offsetSize), offsetSize);
        if (first != 1 || last < first || last - 1L > data.Length - dataStart) return false;
        uint previous = first;
        for (int index = 1; index <= count; index++) {
            uint current = ReadCffOffset(data, (int)(offsetsStart + index * (long)offsetSize), offsetSize);
            if (current < previous || current > last) return false;
            previous = current;
        }
        nextOffset = (int)(dataStart + last - 1L);
        return true;
    }

    private static uint ReadCffOffset(byte[] data, int offset, int size) {
        uint value = 0;
        for (int index = 0; index < size; index++) value = (value << 8) | data[offset + index];
        return value;
    }

    private static bool StartsWithAscii(byte[] data, string value) {
        if (data.Length < value.Length) return false;
        for (int index = 0; index < value.Length; index++) {
            if (data[index] != (byte)value[index]) return false;
        }
        return true;
    }

    private static bool ContainsAscii(byte[] data, string value) {
        if (value.Length == 0 || data.Length < value.Length) return false;
        for (int offset = 0; offset <= data.Length - value.Length; offset++) {
            int index = 0;
            while (index < value.Length && data[offset + index] == (byte)value[index]) index++;
            if (index == value.Length) return true;
        }
        return false;
    }

    private static ushort ReadUInt16BigEndian(byte[] data, int offset) =>
        (ushort)((data[offset] << 8) | data[offset + 1]);

    private static uint ReadUInt32BigEndian(byte[] data, int offset) =>
        ((uint)data[offset] << 24) |
        ((uint)data[offset + 1] << 16) |
        ((uint)data[offset + 2] << 8) |
        data[offset + 3];

    private static uint ReadUInt32LittleEndian(byte[] data, int offset) =>
        data[offset] |
        ((uint)data[offset + 1] << 8) |
        ((uint)data[offset + 2] << 16) |
        ((uint)data[offset + 3] << 24);
}
