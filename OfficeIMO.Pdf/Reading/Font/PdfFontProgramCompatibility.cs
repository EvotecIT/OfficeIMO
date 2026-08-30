namespace OfficeIMO.Pdf;

/// <summary>Maps PDF font dictionary subtypes to compatible embedded-program containers.</summary>
internal static class PdfFontProgramCompatibility {
    internal static bool IsCompatible(string? fontSubtype, string descriptorKey, string? streamSubtype = null) {
        if (string.Equals(descriptorKey, "FontFile", StringComparison.Ordinal)) {
            return string.Equals(fontSubtype, "Type1", StringComparison.Ordinal) ||
                string.Equals(fontSubtype, "MMType1", StringComparison.Ordinal);
        }
        if (string.Equals(descriptorKey, "FontFile2", StringComparison.Ordinal)) {
            return string.Equals(fontSubtype, "TrueType", StringComparison.Ordinal) ||
                string.Equals(fontSubtype, "CIDFontType2", StringComparison.Ordinal);
        }
        if (!string.Equals(descriptorKey, "FontFile3", StringComparison.Ordinal)) return false;

        if (string.Equals(streamSubtype, "Type1C", StringComparison.Ordinal)) {
            return string.Equals(fontSubtype, "Type1", StringComparison.Ordinal) ||
                string.Equals(fontSubtype, "MMType1", StringComparison.Ordinal);
        }
        if (string.Equals(streamSubtype, "CIDFontType0C", StringComparison.Ordinal)) {
            return string.Equals(fontSubtype, "CIDFontType0", StringComparison.Ordinal);
        }
        if (!string.Equals(streamSubtype, "OpenType", StringComparison.Ordinal)) return false;
        return string.Equals(fontSubtype, "Type1", StringComparison.Ordinal) ||
            string.Equals(fontSubtype, "TrueType", StringComparison.Ordinal) ||
            string.Equals(fontSubtype, "CIDFontType0", StringComparison.Ordinal) ||
            string.Equals(fontSubtype, "CIDFontType2", StringComparison.Ordinal);
    }

    internal static bool IsCompatibleOpenTypeProgram(string? fontSubtype, byte[] data) {
        if (data == null || data.Length < 4) return false;
        uint scaler = ((uint)data[0] << 24) |
            ((uint)data[1] << 16) |
            ((uint)data[2] << 8) |
            data[3];
        bool hasTrueTypeOutlines = scaler == 0x00010000U || scaler == 0x74727565U;
        bool hasCffOutlines = scaler == 0x4F54544FU;
        if (!hasTrueTypeOutlines && !hasCffOutlines) return false;
        return hasTrueTypeOutlines
            ? RequiresTrueTypeOutlines(fontSubtype)
            : string.Equals(fontSubtype, "Type1", StringComparison.Ordinal) ||
                string.Equals(fontSubtype, "CIDFontType0", StringComparison.Ordinal);
    }

    internal static bool RequiresTrueTypeOutlines(string? fontSubtype) =>
        string.Equals(fontSubtype, "TrueType", StringComparison.Ordinal) ||
        string.Equals(fontSubtype, "CIDFontType2", StringComparison.Ordinal);
}
