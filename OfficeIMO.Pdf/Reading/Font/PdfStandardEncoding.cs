namespace OfficeIMO.Pdf;

internal static class PdfStandardEncoding {
    private static readonly string[] Map = BuildMap();

    public static string Decode(byte[] bytes) {
        return Decode(bytes, int.MaxValue);
    }

    public static string Decode(byte[] bytes, int maxOutputCharacters) {
        if (bytes is null || bytes.Length == 0) return string.Empty;
        long decodedLength = 0L;
        for (int i = 0; i < bytes.Length; i++) {
            decodedLength += Map[bytes[i]].Length;
            if (decodedLength > maxOutputCharacters) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.DecodedTextCharacters, maxOutputCharacters, decodedLength);
            }
        }

        var builder = new System.Text.StringBuilder((int)decodedLength);
        for (int i = 0; i < bytes.Length; i++) {
            builder.Append(Map[bytes[i]]);
        }

        return builder.ToString();
    }

    private static string[] BuildMap() {
        var map = new string[256];
        for (int i = 0; i < map.Length; i++) map[i] = string.Empty;

        map[32] = " "; map[33] = "!"; map[34] = "\""; map[35] = "#";
        map[36] = "$"; map[37] = "%"; map[38] = "&"; map[39] = "\u2019";
        map[40] = "("; map[41] = ")"; map[42] = "*"; map[43] = "+";
        map[44] = ","; map[45] = "-"; map[46] = "."; map[47] = "/";
        map[48] = "0"; map[49] = "1"; map[50] = "2"; map[51] = "3";
        map[52] = "4"; map[53] = "5"; map[54] = "6"; map[55] = "7";
        map[56] = "8"; map[57] = "9"; map[58] = ":"; map[59] = ";";
        map[60] = "<"; map[61] = "="; map[62] = ">"; map[63] = "?";
        map[64] = "@";
        for (int code = 65; code <= 90; code++) map[code] = ((char)code).ToString();
        map[91] = "["; map[92] = "\\"; map[93] = "]"; map[94] = "^";
        map[95] = "_"; map[96] = "\u2018";
        for (int code = 97; code <= 122; code++) map[code] = ((char)code).ToString();
        map[123] = "{"; map[124] = "|"; map[125] = "}"; map[126] = "~";

        map[161] = "\u00A1"; map[162] = "\u00A2"; map[163] = "\u00A3"; map[164] = "\u2044";
        map[165] = "\u00A5"; map[166] = "\u0192"; map[167] = "\u00A7"; map[168] = "\u00A4";
        map[169] = "'"; map[170] = "\u201C"; map[171] = "\u00AB"; map[172] = "\u2039";
        map[173] = "\u203A"; map[174] = "fi"; map[175] = "fl";
        map[177] = "\u2013"; map[178] = "\u2020"; map[179] = "\u2021"; map[180] = "\u00B7";
        map[182] = "\u00B6"; map[183] = "\u2022"; map[184] = "\u201A"; map[185] = "\u201E";
        map[186] = "\u201D"; map[187] = "\u00BB"; map[188] = "\u2026"; map[189] = "\u2030";
        map[191] = "\u00BF";
        map[193] = "`"; map[194] = "\u00B4"; map[195] = "\u02C6"; map[196] = "\u02DC";
        map[197] = "\u00AF"; map[198] = "\u02D8"; map[199] = "\u02D9"; map[200] = "\u00A8";
        map[202] = "\u02DA"; map[203] = "\u00B8"; map[205] = "\u02DD"; map[206] = "\u02DB";
        map[207] = "\u02C7"; map[208] = "\u2014";
        map[225] = "\u00C6"; map[227] = "\u00AA"; map[232] = "\u0141"; map[233] = "\u00D8";
        map[234] = "\u0152"; map[235] = "\u00BA"; map[241] = "\u00E6"; map[245] = "\u0131";
        map[248] = "\u0142"; map[249] = "\u00F8"; map[250] = "\u0153"; map[251] = "\u00DF";
        return map;
    }
}
