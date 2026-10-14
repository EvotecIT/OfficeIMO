using System;
using System.Collections.Generic;

namespace OfficeIMO.TestAssets;

internal static partial class ManagedTextShapingTestAssets {
    // One synthetic rectangle, translated by 256 units at either axis endpoint.
    // A positive endpoint also adds 16384 advance units, exposing normalization
    // half-steps through integer metrics as well as fractional public contours.
    internal static byte[] CreateNormalizationFont(bool hvar = true, bool avar = false, double extent = 1D,
        (string Tag, double Minimum, double Default, double Maximum)? axis = null) {
        (string Tag, double Minimum, double Default, double Maximum) selectedAxis = axis ?? ("TEST", -extent, 0D, extent);
        var fvar = new byte[36];
        WriteUInt16(fvar, 0, 1);
        WriteUInt16(fvar, 4, 16);
        WriteUInt16(fvar, 6, 2);
        WriteUInt16(fvar, 8, 1);
        WriteUInt16(fvar, 10, 20);
        WriteUInt16(fvar, 14, 8);
        WriteTag(fvar, 16, selectedAxis.Tag);
        WriteUInt32(fvar, 20, unchecked((uint)(int)(selectedAxis.Minimum * 65536D)));
        WriteUInt32(fvar, 24, unchecked((uint)(int)(selectedAxis.Default * 65536D)));
        WriteUInt32(fvar, 28, unchecked((uint)(int)(selectedAxis.Maximum * 65536D)));
        WriteUInt16(fvar, 34, 256);

        const int glyphDataOffset = 26;
        const int glyphDataLength = 52;
        var gvar = new byte[glyphDataOffset + glyphDataLength];
        WriteUInt16(gvar, 0, 1);
        WriteUInt16(gvar, 4, 1);
        WriteUInt32(gvar, 8, glyphDataOffset);
        WriteUInt16(gvar, 12, 2);
        WriteUInt32(gvar, 16, glyphDataOffset);
        WriteUInt16(gvar, 24, glyphDataLength / 2);
        WriteUInt16(gvar, glyphDataOffset, 2);
        WriteUInt16(gvar, glyphDataOffset + 2, 16);
        for (int tuple = 0; tuple < 2; tuple++) {
            int header = glyphDataOffset + 4 + tuple * 6;
            WriteUInt16(gvar, header, 18);
            WriteUInt16(gvar, header + 2, 0x8000);
            WriteInt16(gvar, header + 4, checked((short)(tuple == 0 ? -16384 : 16384)));
            int deltas = glyphDataOffset + 16 + tuple * 18;
            gvar[deltas] = 0x47; // eight signed-word x deltas, including phantom points
            for (int point = 0; point < 4; point++)
                WriteInt16(gvar, deltas + 1 + point * 2, checked((short)(tuple == 0 ? -256 : 256)));
            WriteInt16(gvar, deltas + 11, checked((short)(tuple == 0 ? 0 : 16384)));
            gvar[deltas + 17] = 0x87; // eight zero y deltas
        }

        var tables = new List<(string Tag, byte[] Data)> { ("fvar", fvar), ("gvar", gvar) };
        if (hvar) tables.Add(("HVAR", CreateNormalizationHvar()));
        if (avar) {
            var map = new byte[30];
            WriteUInt16(map, 0, 1);
            WriteUInt16(map, 6, 1);
            WriteUInt16(map, 8, 5);
            int[] from = { -16384, -8192, 0, 8192, 16384 };
            int[] to = { -16384, -12288, 0, 12288, 16384 };
            for (int point = 0; point < from.Length; point++) {
                WriteInt16(map, 10 + point * 4, checked((short)from[point]));
                WriteInt16(map, 12 + point * 4, checked((short)to[point]));
            }
            tables.Add(("avar", map));
        }
        return CreateFontFromCmap(CreateFormat12Cmap(new[] { (int)'A' }),
            includeTrailingMetric: true, additionalTables: tables);
    }

    internal static byte[] WithOpticalSizeRange(byte[] source, double minimum, double defaultValue, double maximum) {
        var font = (byte[])source.Clone();
        int fvarRecord = -1, headOffset = -1;
        for (int index = 0; index < ReadUInt16(font, 4); index++) {
            int record = 12 + index * 16;
            uint tag = ReadUInt32(font, record);
            if (tag == 0x66766172) fvarRecord = record;
            if (tag == 0x68656164) headOffset = checked((int)ReadUInt32(font, record + 8));
        }
        if (fvarRecord < 0 || headOffset < 0) throw new InvalidOperationException("The fixture requires fvar and head tables.");
        int offset = checked((int)ReadUInt32(font, fvarRecord + 8));
        int axes = offset + ReadUInt16(font, offset + 4);
        for (int index = 0; index < ReadUInt16(font, offset + 8); index++) {
            int record = axes + index * ReadUInt16(font, offset + 10);
            if (ReadUInt32(font, record) != 0x6F70737A) continue;
            WriteUInt32(font, record + 4, unchecked((uint)checked((int)(minimum * 65536D))));
            WriteUInt32(font, record + 8, unchecked((uint)checked((int)(defaultValue * 65536D))));
            WriteUInt32(font, record + 12, unchecked((uint)checked((int)(maximum * 65536D))));
            var fvar = new byte[checked((int)ReadUInt32(font, fvarRecord + 12))];
            Array.Copy(font, offset, fvar, 0, fvar.Length);
            WriteUInt32(font, fvarRecord + 4, CalculateChecksum(fvar));
            WriteUInt32(font, headOffset + 8, 0);
            WriteUInt32(font, headOffset + 8, unchecked(0xB1B0AFBAU - CalculateChecksum(font)));
            return font;
        }
        throw new InvalidOperationException("The fixture requires an opsz axis.");
    }

    private static byte[] CreateNormalizationHvar() {
        var hvar = new byte[66];
        WriteUInt16(hvar, 0, 1);
        WriteUInt32(hvar, 4, 20);
        const int store = 20;
        WriteUInt16(hvar, store, 1);
        WriteUInt32(hvar, store + 2, 12);
        WriteUInt16(hvar, store + 6, 1);
        WriteUInt32(hvar, store + 8, 28);
        WriteUInt16(hvar, store + 12, 1); // one axis
        WriteUInt16(hvar, store + 14, 2); // negative and positive regions
        WriteInt16(hvar, store + 16, -16384);
        WriteInt16(hvar, store + 18, -16384);
        WriteInt16(hvar, store + 24, 16384);
        WriteInt16(hvar, store + 26, 16384);
        const int data = store + 28;
        WriteUInt16(hvar, data, 2); // glyphs .notdef and A
        WriteUInt16(hvar, data + 2, 2);
        WriteUInt16(hvar, data + 4, 2);
        WriteUInt16(hvar, data + 8, 1);
        WriteInt16(hvar, data + 16, 16384);
        return hvar;
    }
}
