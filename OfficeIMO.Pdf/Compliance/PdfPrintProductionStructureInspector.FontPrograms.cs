using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfPrintProductionStructureInspector {
    private static bool IsValidFontProgram(
        string? fontSubtype,
        string key,
        PdfStream stream,
        byte[] decoded,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumObjectDepth) {
        if (decoded.Length == 0) return false;
        if (string.Equals(key, "FontFile", StringComparison.Ordinal)) {
            return PdfFontProgramCompatibility.IsCompatible(fontSubtype, key) &&
                IsValidType1Program(decoded);
        }
        if (string.Equals(key, "FontFile2", StringComparison.Ordinal)) {
            return PdfFontProgramCompatibility.IsCompatible(fontSubtype, key) &&
                IsValidSfntProgram(decoded, fontSubtype);
        }
        if (!string.Equals(key, "FontFile3", StringComparison.Ordinal)) return false;

        string? subtype = ResolveName(
            stream.Dictionary.Items.TryGetValue("Subtype", out PdfObject? subtypeObject)
                ? subtypeObject
                : null,
            objects,
            maximumObjectDepth);
        if (!PdfFontProgramCompatibility.IsCompatible(fontSubtype, key, subtype)) return false;
        if (string.Equals(subtype, "Type1C", StringComparison.Ordinal)) {
            return IsValidCff1Program(decoded, requireCidKeyed: false);
        }
        if (string.Equals(subtype, "CIDFontType0C", StringComparison.Ordinal)) {
            return IsValidCff1Program(decoded, requireCidKeyed: true);
        }
        if (string.Equals(subtype, "OpenType", StringComparison.Ordinal)) {
            return IsValidSfntProgram(decoded, fontSubtype);
        }
        return false;
    }

    private static bool IsValidType1Program(byte[] data) {
        return data.Length >= 2 && data[0] == 0x80
            ? IsValidPfbProgram(data)
            : IsValidPfaProgram(data);
    }

    private static bool IsValidSfntProgram(byte[] data, string? fontSubtype) {
        if (data.Length < 12) return false;
        uint scaler = ReadUInt32BigEndian(data, 0);
        bool isTrueType = scaler == 0x00010000U || scaler == 0x74727565U;
        if (!PdfFontProgramCompatibility.IsCompatibleOpenTypeProgram(fontSubtype, data)) return false;

        int tableCount = ReadUInt16BigEndian(data, 4);
        if (tableCount <= 0 || tableCount > 4096 || 12L + tableCount * 16L > data.Length) return false;
        int directoryEnd = 12 + tableCount * 16;
        bool hasHead = false;
        bool hasMaxp = false;
        bool hasGlyf = false;
        bool hasLoca = false;
        bool hasCff = false;
        bool cffIsVersion2 = false;
        int headOffset = 0;
        int headLength = 0;
        int maxpOffset = 0;
        int maxpLength = 0;
        int glyfOffset = 0;
        int glyfLength = 0;
        int locaOffset = 0;
        int locaLength = 0;
        int cffOffset = 0;
        int cffLength = 0;
        var tableRanges = new List<(uint Offset, uint Length)>(tableCount);
        var tableTags = new HashSet<string>(StringComparer.Ordinal);
        for (int index = 0; index < tableCount; index++) {
            int recordOffset = 12 + index * 16;
            uint tableOffset = ReadUInt32BigEndian(data, recordOffset + 8);
            uint tableLength = ReadUInt32BigEndian(data, recordOffset + 12);
            if (tableOffset > data.Length || tableLength > data.Length - tableOffset ||
                (tableOffset & 3U) != 0 || (tableLength > 0 && tableOffset < directoryEnd)) return false;
            if (tableLength > 0) {
                foreach ((uint offset, uint length) in tableRanges) {
                    if (tableOffset < offset + length && offset < tableOffset + tableLength) return false;
                }
                tableRanges.Add((tableOffset, tableLength));
            }
            string tag = new(new[] {
                (char)data[recordOffset],
                (char)data[recordOffset + 1],
                (char)data[recordOffset + 2],
                (char)data[recordOffset + 3]
            });
            if (!tableTags.Add(tag)) return false;
            switch (tag) {
                case "head":
                    hasHead = tableLength >= 54;
                    headOffset = (int)tableOffset;
                    headLength = (int)tableLength;
                    break;
                case "maxp":
                    hasMaxp = tableLength >= 6;
                    maxpOffset = (int)tableOffset;
                    maxpLength = (int)tableLength;
                    break;
                case "glyf":
                    hasGlyf = tableLength > 0;
                    glyfOffset = (int)tableOffset;
                    glyfLength = (int)tableLength;
                    break;
                case "loca":
                    hasLoca = tableLength >= 4;
                    locaOffset = (int)tableOffset;
                    locaLength = (int)tableLength;
                    break;
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
        if (isTrueType) {
            return hasGlyf && hasLoca && IsValidTrueTypeGlyphLocations(
                data,
                headOffset,
                headLength,
                maxpOffset,
                maxpLength,
                glyfOffset,
                glyfLength,
                locaOffset,
                locaLength);
        }
        if (!hasCff) return false;
        var cff = new byte[cffLength];
        Buffer.BlockCopy(data, cffOffset, cff, 0, cffLength);
        if (cffIsVersion2) return IsValidCff2Program(cff);
        bool? requireCidKeyed = string.Equals(fontSubtype, "Type1", StringComparison.Ordinal)
            ? false
            : null;
        return IsValidCff1Program(cff, requireCidKeyed);
    }

    private static bool IsValidTrueTypeGlyphLocations(
        byte[] data,
        int headOffset,
        int headLength,
        int maxpOffset,
        int maxpLength,
        int glyfOffset,
        int glyfLength,
        int locaOffset,
        int locaLength) {
        if (headLength < 54 || maxpLength < 6 || glyfLength <= 0 || locaLength < 4 ||
            headOffset < 0 || headOffset > data.Length - headLength ||
            maxpOffset < 0 || maxpOffset > data.Length - maxpLength ||
            glyfOffset < 0 || glyfOffset > data.Length - glyfLength ||
            locaOffset < 0 || locaOffset > data.Length - locaLength) return false;

        short indexToLocFormat = unchecked((short)ReadUInt16BigEndian(data, headOffset + 50));
        if (indexToLocFormat != 0 && indexToLocFormat != 1) return false;
        int glyphCount = ReadUInt16BigEndian(data, maxpOffset + 4);
        if (glyphCount <= 0) return false;
        int entrySize = indexToLocFormat == 0 ? 2 : 4;
        int expectedLocaLength = checked((glyphCount + 1) * entrySize);
        if (locaLength != expectedLocaLength) return false;

        var locations = new uint[glyphCount + 1];
        uint previous = 0;
        for (int index = 0; index <= glyphCount; index++) {
            int entryOffset = checked(locaOffset + index * entrySize);
            uint current = indexToLocFormat == 0
                ? checked((uint)ReadUInt16BigEndian(data, entryOffset) * 2U)
                : ReadUInt32BigEndian(data, entryOffset);
            if (current < previous || current > glyfLength) return false;
            locations[index] = current;
            previous = current;
        }
        return OfficeTrueTypeGlyphData.IsStructurallyValid(data, glyfOffset, glyfLength, locations);
    }

    private static bool IsValidCff1Program(byte[] data, bool? requireCidKeyed) {
        return OfficeCffFontData.IsStructurallyValidProgram(
            data,
            isCff2: false,
            requireCidKeyed: requireCidKeyed);
    }

    private static bool IsValidCff2Program(byte[] data) {
        return OfficeCffFontData.IsStructurallyValidProgram(data, isCff2: true);
    }

    private static bool IsValidPfaProgram(byte[] data) {
        if (!StartsWithType1Header(data)) return false;
        int eexec = IndexOfAscii(data, "eexec", 0);
        int clearToMark = eexec < 0 ? -1 : IndexOfAscii(data, "cleartomark", eexec + 5);
        if (eexec < 0 || clearToMark < 0) return false;
        int offset = eexec + 5;
        while (offset < clearToMark && IsAsciiWhitespace(data[offset])) offset++;
        var encrypted = new List<byte>();
        bool hexadecimal = LooksLikeHexEexec(data, offset, clearToMark);
        if (hexadecimal) {
            int high = -1;
            for (int index = offset; index < clearToMark; index++) {
                if (IsAsciiWhitespace(data[index])) continue;
                int nibble = HexNibble(data[index]);
                if (nibble < 0) break;
                if (high < 0) high = nibble;
                else {
                    encrypted.Add((byte)((high << 4) | nibble));
                    high = -1;
                }
            }
            if (high >= 0) return false;
        } else {
            for (int index = offset; index < clearToMark; index++) encrypted.Add(data[index]);
        }
        return HasUsableType1PrivateProgram(encrypted.ToArray());
    }

    private static bool IsValidPfbProgram(byte[] data) {
        int offset = 0;
        bool hasHeader = false;
        bool hasEexec = false;
        bool hasTrailer = false;
        bool hasEof = false;
        var encrypted = new List<byte>();
        while (offset <= data.Length - 2 && data[offset] == 0x80) {
            byte type = data[offset + 1];
            if (type == 0x03) {
                hasEof = true;
                offset += 2;
                break;
            }
            if ((type != 0x01 && type != 0x02) || offset > data.Length - 6) return false;
            uint lengthValue = ReadUInt32LittleEndian(data, offset + 2);
            if (lengthValue > int.MaxValue) return false;
            int length = (int)lengthValue;
            int segment = offset + 6;
            if (segment > data.Length - length) return false;
            if (type == 0x01) {
                var ascii = new byte[length];
                Buffer.BlockCopy(data, segment, ascii, 0, length);
                if (!hasHeader) hasHeader = StartsWithType1Header(ascii);
                if (ContainsAscii(ascii, "eexec")) hasEexec = true;
                if (ContainsAscii(ascii, "cleartomark")) hasTrailer = true;
            } else {
                for (int index = 0; index < length; index++) encrypted.Add(data[segment + index]);
            }
            offset = segment + length;
        }
        return hasHeader && hasEexec && hasTrailer && hasEof && offset == data.Length &&
            HasUsableType1PrivateProgram(encrypted.ToArray());
    }

    private static bool HasUsableType1PrivateProgram(byte[] encrypted) {
        if (encrypted.Length <= 4) return false;
        var decrypted = new byte[encrypted.Length];
        ushort state = 55665;
        for (int index = 0; index < encrypted.Length; index++) {
            byte cipher = encrypted[index];
            decrypted[index] = (byte)(cipher ^ (state >> 8));
            state = unchecked((ushort)((cipher + state) * 52845 + 22719));
        }
        int privateBody = FindType1DictionaryBody(decrypted, "/Private", 4);
        int charStringsBody = FindType1DictionaryBody(decrypted, "/CharStrings", 4);
        return privateBody >= 0 && charStringsBody >= 0 && ContainsAscii(decrypted, "/.notdef", charStringsBody);
    }

    private static bool LooksLikeHexEexec(byte[] data, int startOffset, int endOffset) {
        int digits = 0;
        for (int offset = startOffset; offset < endOffset && digits < 4; offset++) {
            if (IsAsciiWhitespace(data[offset])) continue;
            if (HexNibble(data[offset]) < 0) return false;
            digits++;
        }
        return digits == 4;
    }

    private static int FindType1DictionaryBody(byte[] data, string token, int startOffset) {
        int offset = startOffset;
        while ((offset = IndexOfAscii(data, token, offset)) >= 0) {
            int cursor = offset + token.Length;
            if (cursor >= data.Length || !IsAsciiWhitespace(data[cursor])) {
                offset++;
                continue;
            }
            while (cursor < data.Length && IsAsciiWhitespace(data[cursor])) cursor++;
            int digitStart = cursor;
            while (cursor < data.Length && data[cursor] >= (byte)'0' && data[cursor] <= (byte)'9') cursor++;
            if (cursor == digitStart || cursor >= data.Length || !IsAsciiWhitespace(data[cursor])) {
                offset++;
                continue;
            }
            while (cursor < data.Length && IsAsciiWhitespace(data[cursor])) cursor++;
            if (cursor <= data.Length - 4 && data[cursor] == (byte)'d' && data[cursor + 1] == (byte)'i' &&
                data[cursor + 2] == (byte)'c' && data[cursor + 3] == (byte)'t') return cursor + 4;
            offset++;
        }
        return -1;
    }

    private static bool StartsWithType1Header(byte[] data) =>
        StartsWithAscii(data, "%!PS-AdobeFont") || StartsWithAscii(data, "%!FontType1");

    private static bool IsAsciiWhitespace(byte value) => value == 0 || value == 9 || value == 10 || value == 12 || value == 13 || value == 32;

    private static int HexNibble(byte value) {
        if (value >= (byte)'0' && value <= (byte)'9') return value - (byte)'0';
        if (value >= (byte)'A' && value <= (byte)'F') return value - (byte)'A' + 10;
        if (value >= (byte)'a' && value <= (byte)'f') return value - (byte)'a' + 10;
        return -1;
    }

    private static bool StartsWithAscii(byte[] data, string value) {
        if (data.Length < value.Length) return false;
        for (int index = 0; index < value.Length; index++) {
            if (data[index] != (byte)value[index]) return false;
        }
        return true;
    }

    private static bool ContainsAscii(byte[] data, string value) {
        return ContainsAscii(data, value, 0);
    }

    private static bool ContainsAscii(byte[] data, string value, int startOffset) =>
        IndexOfAscii(data, value, startOffset) >= 0;

    private static int IndexOfAscii(byte[] data, string value, int startOffset) {
        if (value.Length == 0 || startOffset < 0 || data.Length - startOffset < value.Length) return -1;
        for (int offset = startOffset; offset <= data.Length - value.Length; offset++) {
            int index = 0;
            while (index < value.Length && data[offset + index] == (byte)value[index]) index++;
            if (index == value.Length) return offset;
        }
        return -1;
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
