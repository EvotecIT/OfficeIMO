using System;
using System.Collections.Generic;
using OfficeIMO.Drawing;

namespace OfficeIMO.TestAssets;

internal static class ManagedTextShapingTestAssets {
    internal const string FamilyName = "OfficeIMO Shaping Test";

    internal static byte[] CreateFont(params int[] scalars) {
        if (scalars == null || scalars.Length == 0) throw new ArgumentException("At least one scalar is required.", nameof(scalars));
        byte[] cmap = CreateFormat12Cmap(scalars);
        byte[] glyph = CreateVisibleGlyph();
        var tables = new List<(string Tag, byte[] Data)> {
            ("cmap", cmap),
            ("glyf", glyph),
            ("head", CreateHeadTable()),
            ("hhea", CreateHheaTable()),
            ("hmtx", new byte[] { 0x01, 0xF4, 0x00, 0x00 }),
            ("loca", new byte[] { 0x00, 0x00, 0x00, 0x00, 0x00, (byte)(glyph.Length / 2) }),
            ("maxp", new byte[] { 0x00, 0x01, 0x00, 0x00, 0x00, 0x02 }),
            ("name", new byte[6])
        };

        int tableDirectoryLength = 12 + (tables.Count * 16);
        var offsets = new int[tables.Count];
        int offset = tableDirectoryLength;
        for (int index = 0; index < tables.Count; index++) {
            offsets[index] = offset;
            offset += Align4(tables[index].Data.Length);
        }

        var font = new byte[offset];
        WriteUInt32(font, 0, 0x00010000);
        WriteUInt16(font, 4, (ushort)tables.Count);
        for (int index = 0; index < tables.Count; index++) {
            int record = 12 + (index * 16);
            WriteTag(font, record, tables[index].Tag);
            WriteUInt32(font, record + 8, (uint)offsets[index]);
            WriteUInt32(font, record + 12, (uint)tables[index].Data.Length);
            Array.Copy(tables[index].Data, 0, font, offsets[index], tables[index].Data.Length);
        }

        return font;
    }

    internal static byte[] CreateFontCollection(params int[] scalars) {
        byte[] first = CreateFont('A');
        byte[] second = CreateFont(scalars);
        const int headerLength = 20;
        int firstOffset = headerLength;
        int secondOffset = Align4(firstOffset + first.Length);
        var collection = new byte[secondOffset + second.Length];
        WriteTag(collection, 0, "ttcf");
        WriteUInt32(collection, 4, 0x00010000);
        WriteUInt32(collection, 8, 2);
        WriteUInt32(collection, 12, (uint)firstOffset);
        WriteUInt32(collection, 16, (uint)secondOffset);
        CopyCollectionFace(first, collection, firstOffset);
        CopyCollectionFace(second, collection, secondOffset);
        return collection;
    }

    internal sealed class RecordingProvider : IOfficeTextShapingProvider {
        private readonly object _gate = new();
        private readonly List<OfficeTextShapingRequest> _requests = new();

        internal IReadOnlyList<OfficeTextShapingRequest> Requests {
            get {
                lock (_gate) return _requests.ToArray();
            }
        }

        public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) {
            lock (_gate) _requests.Add(request);
            var glyphs = new List<OfficeShapedGlyph>();
            int textIndex = 0;
            foreach (string element in OfficeTextElements.Enumerate(request.Text)) {
                glyphs.Add(new OfficeShapedGlyph(1, element, textIndex, advanceWidth: 500));
                textIndex += element.Length;
            }
            return new OfficeTextShapingResult(glyphs);
        }
    }

    private static byte[] CreateFormat12Cmap(int[] scalars) {
        var ordered = new SortedSet<int>(scalars);
        var data = new byte[28 + (ordered.Count * 12)];
        WriteUInt16(data, 2, 1);
        WriteUInt16(data, 4, 3);
        WriteUInt16(data, 6, 10);
        WriteUInt32(data, 8, 12);
        WriteUInt16(data, 12, 12);
        WriteUInt32(data, 16, (uint)(16 + (ordered.Count * 12)));
        WriteUInt32(data, 24, (uint)ordered.Count);
        int offset = 28;
        foreach (int scalar in ordered) {
            WriteUInt32(data, offset, (uint)scalar);
            WriteUInt32(data, offset + 4, (uint)scalar);
            WriteUInt32(data, offset + 8, 1);
            offset += 12;
        }
        return data;
    }

    private static byte[] CreateVisibleGlyph() {
        var glyph = new byte[34];
        WriteUInt16(glyph, 0, 1);
        WriteUInt16(glyph, 6, 400);
        WriteUInt16(glyph, 8, 700);
        WriteUInt16(glyph, 10, 3);
        glyph[14] = 0x01;
        glyph[15] = 0x01;
        glyph[16] = 0x01;
        glyph[17] = 0x01;
        WriteUInt16(glyph, 20, 400);
        WriteUInt16(glyph, 24, unchecked((ushort)-400));
        WriteUInt16(glyph, 30, 700);
        return glyph;
    }

    private static byte[] CreateHeadTable() {
        var table = new byte[54];
        WriteUInt16(table, 18, 1000);
        return table;
    }

    private static byte[] CreateHheaTable() {
        var table = new byte[36];
        WriteUInt16(table, 4, 800);
        WriteUInt16(table, 6, unchecked((ushort)-200));
        WriteUInt16(table, 34, 1);
        return table;
    }

    private static void CopyCollectionFace(byte[] source, byte[] destination, int destinationOffset) {
        Array.Copy(source, 0, destination, destinationOffset, source.Length);
        int tableCount = ReadUInt16(source, 4);
        for (int index = 0; index < tableCount; index++) {
            int sourceRecord = 12 + (index * 16);
            int destinationRecord = destinationOffset + sourceRecord;
            uint tableOffset = ReadUInt32(source, sourceRecord + 8);
            WriteUInt32(destination, destinationRecord + 8, checked((uint)destinationOffset + tableOffset));
        }
    }

    private static int Align4(int value) => (value + 3) & ~3;

    private static void WriteTag(byte[] data, int offset, string tag) {
        for (int index = 0; index < 4; index++) data[offset + index] = (byte)tag[index];
    }

    private static void WriteUInt16(byte[] data, int offset, ushort value) {
        data[offset] = (byte)(value >> 8);
        data[offset + 1] = (byte)value;
    }

    private static void WriteUInt32(byte[] data, int offset, uint value) {
        data[offset] = (byte)(value >> 24);
        data[offset + 1] = (byte)(value >> 16);
        data[offset + 2] = (byte)(value >> 8);
        data[offset + 3] = (byte)value;
    }

    private static ushort ReadUInt16(byte[] data, int offset) =>
        (ushort)((data[offset] << 8) | data[offset + 1]);

    private static uint ReadUInt32(byte[] data, int offset) =>
        ((uint)data[offset] << 24) |
        ((uint)data[offset + 1] << 16) |
        ((uint)data[offset + 2] << 8) |
        data[offset + 3];
}
