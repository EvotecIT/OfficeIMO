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
        if (data.Length >= 10 && data[0] == 0x80 && data[1] == 0x01) {
            uint segmentLength = ReadUInt32LittleEndian(data, 2);
            if (segmentLength < 2 || segmentLength > data.Length - 6) return false;
            return data[6] == (byte)'%' && data[7] == (byte)'!' && ContainsAscii(data, "eexec");
        }

        return (StartsWithAscii(data, "%!PS-AdobeFont") || StartsWithAscii(data, "%!FontType1")) &&
            ContainsAscii(data, "eexec");
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
        if (cffIsVersion2) return IsValidCff2Program(cff);
        bool? requireCidKeyed = string.Equals(fontSubtype, "Type1", StringComparison.Ordinal)
            ? false
            : null;
        return IsValidCff1Program(cff, requireCidKeyed);
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
