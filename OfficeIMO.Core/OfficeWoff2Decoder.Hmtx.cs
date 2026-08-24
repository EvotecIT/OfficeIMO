using System;
using System.IO;

namespace OfficeIMO.Drawing;

internal static partial class OfficeWoff2Decoder {
    private static byte[] ReconstructHmtx(
        byte[] transformed,
        byte[] hhea,
        byte[] maxp,
        byte[] head,
        byte[] glyf,
        byte[] loca) {
        if (transformed.Length < 1) throw new InvalidDataException("The transformed WOFF 2 hmtx table is truncated.");
        if (hhea.Length < 36 || maxp.Length < 6 || head.Length < 54) {
            throw new InvalidDataException("A table required to reconstruct WOFF 2 hmtx is truncated.");
        }
        byte flags = transformed[0];
        if ((flags & 0xFC) != 0) throw new InvalidDataException("The transformed WOFF 2 hmtx flags are invalid.");
        bool hasProportionalBearings = (flags & 1) == 0;
        bool hasMonospaceBearings = (flags & 2) == 0;
        if (hasProportionalBearings && hasMonospaceBearings) {
            throw new InvalidDataException("A transformed WOFF 2 hmtx table must omit at least one bearing array.");
        }

        int glyphCount = ReadUInt16(maxp, 4);
        int metricCount = ReadUInt16(hhea, 34);
        if (glyphCount <= 0 || metricCount <= 0 || metricCount > glyphCount) {
            throw new InvalidDataException("The WOFF 2 hhea or maxp metrics count is invalid.");
        }
        short[] xMinimums = ReadGlyphXMinimums(glyphCount, head, glyf, loca);
        int cursor = 1;
        var advances = new ushort[metricCount];
        for (int index = 0; index < metricCount; index++) {
            advances[index] = ReadUInt16(transformed, cursor);
            cursor += 2;
        }

        var proportionalBearings = new short[metricCount];
        if (hasProportionalBearings) {
            for (int index = 0; index < metricCount; index++) {
                proportionalBearings[index] = ReadInt16(transformed, cursor);
                cursor += 2;
            }
        } else {
            Array.Copy(xMinimums, proportionalBearings, metricCount);
        }

        int monospaceCount = glyphCount - metricCount;
        var monospaceBearings = new short[monospaceCount];
        if (hasMonospaceBearings) {
            for (int index = 0; index < monospaceCount; index++) {
                monospaceBearings[index] = ReadInt16(transformed, cursor);
                cursor += 2;
            }
        } else {
            Array.Copy(xMinimums, metricCount, monospaceBearings, 0, monospaceCount);
        }
        if (cursor != transformed.Length) throw new InvalidDataException("The transformed WOFF 2 hmtx table contains trailing data.");

        var result = new ByteBuilder();
        for (int index = 0; index < metricCount; index++) {
            result.AddUInt16(advances[index]);
            result.AddInt16(proportionalBearings[index]);
        }
        foreach (short bearing in monospaceBearings) result.AddInt16(bearing);
        return result.ToArray();
    }

    private static short[] ReadGlyphXMinimums(int glyphCount, byte[] head, byte[] glyf, byte[] loca) {
        int indexFormat = ReadInt16(head, 50);
        if (indexFormat != 0 && indexFormat != 1) throw new InvalidDataException("The WOFF 2 loca index format is invalid.");
        int requiredLocaLength = checked((glyphCount + 1) * (indexFormat == 0 ? 2 : 4));
        if (loca.Length < requiredLocaLength) throw new InvalidDataException("The reconstructed WOFF 2 loca table is truncated.");
        var result = new short[glyphCount];
        for (int glyphId = 0; glyphId < glyphCount; glyphId++) {
            uint start = indexFormat == 0
                ? (uint)(ReadUInt16(loca, glyphId * 2) * 2)
                : ReadUInt32(loca, glyphId * 4);
            uint end = indexFormat == 0
                ? (uint)(ReadUInt16(loca, (glyphId + 1) * 2) * 2)
                : ReadUInt32(loca, (glyphId + 1) * 4);
            if (start > end || end > glyf.Length) throw new InvalidDataException("The reconstructed WOFF 2 loca offsets are invalid.");
            if (end == start) continue;
            if (end - start < 10) throw new InvalidDataException("A reconstructed WOFF 2 glyph header is truncated.");
            result[glyphId] = ReadInt16(glyf, checked((int)start + 2));
        }
        return result;
    }
}
