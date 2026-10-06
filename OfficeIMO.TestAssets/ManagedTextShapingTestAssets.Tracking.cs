using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeIMO.TestAssets;

internal static partial class ManagedTextShapingTestAssets {
    internal static byte[] CreateTrackingFont(byte[]? tracking, bool color = false) =>
        CreateFontFromCmap(CreateFormat12Cmap(new[] { (int)'A', (int)'B' }),
            glyphCount: color ? 4 : 2, colr: color ? CreateColrV0() : null,
            cpal: color ? CreateCpalV1() : null, tracking: tracking);

    internal static byte[] AddTrackingTable(byte[] source, byte[] tracking, bool includeStat = false) {
        int count = source[4] * 256 + source[5];
        var tables = new List<(string Tag, byte[] Data)>();
        for (int index = 0; index < count; index++) {
            int record = 12 + index * 16;
            string tag = Encoding.ASCII.GetString(source, record, 4);
            int offset = ReadTrackingU32(source, record + 8);
            int length = ReadTrackingU32(source, record + 12);
            if (tag != "trak" && (!includeStat || tag != "STAT")) tables.Add((tag, source.Skip(offset).Take(length).ToArray()));
        }
        tables.Add(("trak", tracking));
        if (includeStat) {
            var stat = new byte[28];
            WriteUInt16(stat, 0, 1);
            WriteUInt16(stat, 2, 1);
            WriteUInt16(stat, 18, 256);
            WriteUInt16(stat, 4, 8);
            WriteUInt16(stat, 6, 1);
            WriteUInt32(stat, 8, 20);
            WriteUInt32(stat, 14, 28);
            WriteTag(stat, 20, "opsz");
            WriteUInt16(stat, 24, 256);
            tables.Add(("STAT", stat));
        }
        tables.Sort((left, right) => string.CompareOrdinal(left.Tag, right.Tag));
        int position = 12 + tables.Count * 16;
        var result = new byte[position + tables.Sum(table => Align4(table.Data.Length))];
        Array.Copy(source, result, 12);
        WriteUInt16(result, 4, checked((ushort)tables.Count));
        for (int index = 0; index < tables.Count; index++) {
            int record = 12 + index * 16;
            WriteTag(result, record, tables[index].Tag);
            WriteUInt32(result, record + 8, checked((uint)position));
            WriteUInt32(result, record + 12, checked((uint)tables[index].Data.Length));
            Array.Copy(tables[index].Data, 0, result, position, tables[index].Data.Length);
            position += Align4(tables[index].Data.Length);
        }
        return result;
    }

    private static int ReadTrackingU32(byte[] data, int offset) => checked((int)(
        ((uint)data[offset] << 24) | ((uint)data[offset + 1] << 16) | ((uint)data[offset + 2] << 8) | data[offset + 3]));

    internal static byte[] CreateTrackingTable(double[] sizes, double[] tracks, short[][] values) {
        int sizesOffset = 20 + tracks.Length * 8;
        int valuesOffset = sizesOffset + sizes.Length * 4;
        var table = new byte[valuesOffset + tracks.Length * sizes.Length * 2];
        WriteUInt32(table, 0, 0x00010000);
        WriteUInt16(table, 6, 12);
        WriteUInt16(table, 12, checked((ushort)tracks.Length));
        WriteUInt16(table, 14, checked((ushort)sizes.Length));
        WriteUInt32(table, 16, checked((uint)sizesOffset));
        for (int index = 0; index < sizes.Length; index++)
            WriteUInt32(table, sizesOffset + index * 4, unchecked((uint)(int)(sizes[index] * 65536D)));
        for (int track = 0; track < tracks.Length; track++) {
            int record = 20 + track * 8;
            WriteUInt32(table, record, unchecked((uint)(int)(tracks[track] * 65536D)));
            WriteUInt16(table, record + 4, checked((ushort)(256 + track)));
            WriteUInt16(table, record + 6, checked((ushort)(valuesOffset + track * sizes.Length * 2)));
            for (int size = 0; size < sizes.Length; size++)
                WriteInt16(table, valuesOffset + (track * sizes.Length + size) * 2, values[track][size]);
        }
        return table;
    }
}
