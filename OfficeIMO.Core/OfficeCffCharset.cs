using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeIMO.Drawing;

/// <summary>Validated CFF1 charset lookup used by Type 2 seac-compatible endchar programs.</summary>
internal static class OfficeCffCharset {
    private static readonly ushort[] StandardEncodingSids = {
        0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
        1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,
        33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50,51,52,53,54,55,56,57,58,59,60,61,62,63,64,
        65,66,67,68,69,70,71,72,73,74,75,76,77,78,79,80,81,82,83,84,85,86,87,88,89,90,91,92,93,94,95,0,
        0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
        96,97,98,99,100,101,102,103,104,105,106,107,108,109,110,0,111,112,113,114,0,115,116,117,118,119,120,121,122,0,123,0,
        124,125,126,127,128,129,130,131,0,132,133,0,134,135,136,137,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
        138,0,139,0,0,0,0,140,141,142,143,0,0,0,0,0,144,0,0,0,145,0,0,146,147,148,149,0,0,0,0
    };

    private static readonly ushort[] ExpertCharset = {
        0,1,229,230,231,232,233,234,235,236,237,238,13,14,15,99,239,240,241,242,243,244,245,246,247,248,27,28,249,250,251,252,
        253,254,255,256,257,258,259,260,261,262,263,264,265,266,109,110,267,268,269,270,271,272,273,274,275,276,277,278,279,280,
        281,282,283,284,285,286,287,288,289,290,291,292,293,294,295,296,297,298,299,300,301,302,303,304,305,306,307,308,309,310,
        311,312,313,314,315,316,317,318,158,155,163,319,320,321,322,323,324,325,326,150,164,169,327,328,329,330,331,332,333,334,
        335,336,337,338,339,340,341,342,343,344,345,346,347,348,349,350,351,352,353,354,355,356,357,358,359,360,361,362,363,364,
        365,366,367,368,369,370,371,372,373,374,375,376,377,378
    };

    private static readonly ushort[] ExpertSubsetCharset = {
        0,1,231,232,235,236,237,238,13,14,15,99,239,240,241,242,243,244,245,246,247,248,27,28,249,250,251,253,254,255,256,257,
        258,259,260,261,262,263,264,265,266,109,110,267,268,269,270,272,300,301,302,305,314,315,158,155,163,320,321,322,323,324,
        325,326,150,164,169,327,328,329,330,331,332,333,334,335,336,337,338,339,340,341,342,343,344,345,346
    };

    internal static int[] BuildStandardEncodingGlyphMap(
        byte[] data,
        int tableOffset,
        int tableEnd,
        int charsetValue,
        int glyphCount) {
        ushort[] sids = charsetValue switch {
            0 => BuildIsoAdobeCharset(glyphCount),
            1 => CopyPredefined(ExpertCharset, glyphCount),
            2 => CopyPredefined(ExpertSubsetCharset, glyphCount),
            _ => ReadCustomCharset(data, checked(tableOffset + charsetValue), tableEnd, glyphCount)
        };
        var glyphBySid = new Dictionary<ushort, int>(sids.Length);
        for (int glyph = 1; glyph < sids.Length; glyph++) {
            if (!glyphBySid.ContainsKey(sids[glyph])) glyphBySid.Add(sids[glyph], glyph);
        }
        var result = new int[256];
        for (int code = 0; code < result.Length; code++) {
            ushort sid = StandardEncodingSids[code];
            result[code] = sid != 0 && glyphBySid.TryGetValue(sid, out int glyph) ? glyph : -1;
        }
        return result;
    }

    private static ushort[] BuildIsoAdobeCharset(int glyphCount) {
        if (glyphCount > 229) throw new InvalidDataException("The predefined ISOAdobe charset is smaller than the CFF glyph set.");
        var result = new ushort[glyphCount];
        for (int glyph = 0; glyph < glyphCount; glyph++) result[glyph] = (ushort)glyph;
        return result;
    }

    private static ushort[] CopyPredefined(ushort[] source, int glyphCount) {
        if (glyphCount > source.Length) throw new InvalidDataException("The predefined CFF charset is smaller than the glyph set.");
        var result = new ushort[glyphCount];
        Array.Copy(source, result, glyphCount);
        return result;
    }

    private static ushort[] ReadCustomCharset(byte[] data, int offset, int end, int glyphCount) {
        if (offset < 0 || offset >= end) throw new InvalidDataException("The CFF charset offset is invalid.");
        var result = new ushort[glyphCount];
        int cursor = offset;
        int format = data[cursor++];
        int glyph = 1;
        if (format == 0) {
            if (cursor > end - checked((glyphCount - 1) * 2)) throw new InvalidDataException("The CFF charset is truncated.");
            while (glyph < glyphCount) {
                result[glyph++] = ReadUInt16(data, cursor);
                cursor += 2;
            }
            return result;
        }
        if (format != 1 && format != 2) throw new NotSupportedException("The CFF charset format is not supported.");
        while (glyph < glyphCount) {
            int rangeSize = format == 1 ? 3 : 4;
            if (cursor > end - rangeSize) throw new InvalidDataException("The CFF charset range is truncated.");
            int firstSid = ReadUInt16(data, cursor);
            int left = format == 1 ? data[cursor + 2] : ReadUInt16(data, cursor + 2);
            cursor += rangeSize;
            if (left >= glyphCount - glyph || firstSid > ushort.MaxValue - left) {
                throw new InvalidDataException("The CFF charset range exceeds the glyph set.");
            }
            for (int index = 0; index <= left; index++) result[glyph++] = (ushort)(firstSid + index);
        }
        return result;
    }

    private static ushort ReadUInt16(byte[] data, int offset) =>
        unchecked((ushort)((data[offset] << 8) | data[offset + 1]));
}
