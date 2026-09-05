using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using OfficeIMO.Drawing;

namespace OfficeIMO.TestAssets;

internal static class ManagedTextShapingTestAssets {
    internal const string FamilyName = "OfficeIMO Shaping Test";

    internal static byte[] CreateFont(params int[] scalars) {
        if (scalars == null || scalars.Length == 0) throw new ArgumentException("At least one scalar is required.", nameof(scalars));
        return CreateFontFromCmap(CreateFormat12Cmap(scalars));
    }

    internal static byte[] CreateFontWithKerning(int leftScalar, int rightScalar, short adjustment) {
        if (leftScalar == rightScalar) throw new ArgumentException("Kerning test scalars must be distinct.", nameof(rightScalar));
        return CreateFontFromCmap(
            CreateFormat12Cmap(leftScalar, 1, rightScalar, 2),
            glyphCount: 3,
            kern: CreateKernTable(1, 2, adjustment));
    }

    internal static byte[] CreateFontWithLigature(int firstScalar, int secondScalar, string featureTag = "liga") {
        if (firstScalar == secondScalar) throw new ArgumentException("Ligature test scalars must be distinct.", nameof(secondScalar));
        return CreateFontFromCmap(
            CreateFormat12Cmap(firstScalar, 1, secondScalar, 2),
            glyphCount: 4,
            gsub: CreateLigatureGsub(featureTag, 1, 2, 3));
    }

    internal static byte[] CreateColorFont(int scalar) {
        return CreateFontFromCmap(
            CreateFormat12Cmap(new[] { scalar }),
            glyphCount: 4,
            distinctSecondGlyph: true,
            colr: CreateColrV0(),
            cpal: CreateCpalV1());
    }

    internal static byte[] CreateFontWithUnicodeCmapFallback(int bmpScalar, int supplementalScalar) {
        if (bmpScalar < 0 || bmpScalar > 0xFFFF) throw new ArgumentOutOfRangeException(nameof(bmpScalar));
        if (supplementalScalar <= 0xFFFF || supplementalScalar > 0x10FFFF) {
            throw new ArgumentOutOfRangeException(nameof(supplementalScalar));
        }
        return CreateFontFromCmap(
            CreateUnicodeCmapWithFallback(bmpScalar, supplementalScalar),
            includeTrailingMetric: true);
    }

    internal static byte[] CreateFontWithConflictingUnicodeCmapFallback(int bmpScalar, int supplementalScalar) {
        if (bmpScalar < 0 || bmpScalar > 0xFFFF) throw new ArgumentOutOfRangeException(nameof(bmpScalar));
        if (supplementalScalar <= 0xFFFF || supplementalScalar > 0x10FFFF) {
            throw new ArgumentOutOfRangeException(nameof(supplementalScalar));
        }
        return CreateFontFromCmap(
            CreateConflictingUnicodeCmapFallback(bmpScalar, supplementalScalar),
            glyphCount: 3);
    }

    internal static byte[] CreateFontWithVariationSequence(int scalar, int variationSelector) {
        if (scalar < 0 || scalar > 0x10FFFF) throw new ArgumentOutOfRangeException(nameof(scalar));
        if (variationSelector < 0xFE00 || variationSelector > 0xFE0F) {
            throw new ArgumentOutOfRangeException(nameof(variationSelector));
        }
        return CreateFontFromCmap(CreateUnicodeCmapWithVariationSequence(
            scalar,
            variationSelector,
            nonDefaultGlyph: null,
            variationPlatform: 0,
            variationEncoding: 5));
    }

    internal static byte[] CreateFontWithNonDefaultVariationSequence(int scalar, int variationSelector) {
        if (scalar < 0 || scalar > 0x10FFFF) throw new ArgumentOutOfRangeException(nameof(scalar));
        if (variationSelector < 0xFE00 || variationSelector > 0xFE0F) {
            throw new ArgumentOutOfRangeException(nameof(variationSelector));
        }
        return CreateFontFromCmap(
            CreateUnicodeCmapWithVariationSequence(
                scalar,
                variationSelector,
                nonDefaultGlyph: 2,
                variationPlatform: 0,
                variationEncoding: 5),
            glyphCount: 3,
            distinctSecondGlyph: true);
    }

    internal static byte[] CreateFontWithMistypedVariationSequenceRecord(int scalar, int variationSelector) =>
        CreateFontFromCmap(CreateUnicodeCmapWithVariationSequence(
            scalar,
            variationSelector,
            nonDefaultGlyph: null,
            variationPlatform: 3,
            variationEncoding: 10));

    internal static byte[] CreateFontWithLargeNonDefaultVariationSequence(int scalar, int variationSelector) {
        const int mappingCount = 4097;
        if (scalar < mappingCount - 1) throw new ArgumentOutOfRangeException(nameof(scalar));
        return CreateFontFromCmap(
            CreateUnicodeCmapWithVariationSequence(
                scalar,
                variationSelector,
                nonDefaultGlyph: 2,
                variationPlatform: 0,
                variationEncoding: 5,
                nonDefaultMappingCount: mappingCount),
            glyphCount: 3,
            distinctSecondGlyph: true);
    }

    private static byte[] CreateFontFromCmap(
        byte[] cmap,
        bool includeTrailingMetric = false,
        int glyphCount = 2,
        byte[]? kern = null,
        byte[]? gsub = null,
        bool distinctSecondGlyph = false,
        byte[]? colr = null,
        byte[]? cpal = null) {
        byte[] glyph = CreateVisibleGlyph(400);
        var glyf = new byte[(glyphCount - 1) * glyph.Length];
        var loca = new byte[(glyphCount + 1) * 2];
        var hmtx = new byte[4 + (glyphCount - 1) * 2];
        Array.Copy(new byte[] { 0x01, 0xF4, 0x00, 0x00 }, hmtx, 4);
        for (int glyphIndex = 1; glyphIndex < glyphCount; glyphIndex++) {
            byte[] currentGlyph = distinctSecondGlyph && glyphIndex == 2 ? CreateVisibleGlyph(600) : glyph;
            Array.Copy(currentGlyph, 0, glyf, (glyphIndex - 1) * glyph.Length, glyph.Length);
            WriteUInt16(loca, (glyphIndex + 1) * 2, checked((ushort)(glyphIndex * glyph.Length / 2)));
        }
        if (!includeTrailingMetric && glyphCount == 2) hmtx = new byte[] { 0x01, 0xF4, 0x00, 0x00 };
        var maxp = new byte[] { 0x00, 0x01, 0x00, 0x00, 0x00, checked((byte)glyphCount) };
        var tables = new List<(string Tag, byte[] Data)> {
            ("cmap", cmap),
            ("glyf", glyf),
            ("head", CreateHeadTable()),
            ("hhea", CreateHheaTable()),
            ("hmtx", hmtx),
            ("loca", loca),
            ("maxp", maxp),
            ("name", new byte[6])
        };
        if (kern != null) tables.Add(("kern", kern));
        if (gsub != null) tables.Add(("GSUB", gsub));
        if (colr != null) tables.Add(("COLR", colr));
        if (cpal != null) tables.Add(("CPAL", cpal));

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

    private static byte[] CreateColrV0() {
        var table = new byte[28];
        WriteUInt16(table, 0, 0);
        WriteUInt16(table, 2, 1);
        WriteUInt32(table, 4, 14);
        WriteUInt32(table, 8, 20);
        WriteUInt16(table, 12, 2);
        WriteUInt16(table, 14, 1);
        WriteUInt16(table, 16, 0);
        WriteUInt16(table, 18, 2);
        WriteUInt16(table, 20, 2);
        WriteUInt16(table, 22, 0);
        WriteUInt16(table, 24, 3);
        WriteUInt16(table, 26, 1);
        return table;
    }

    private static byte[] CreateCpalV1() {
        var table = new byte[52];
        WriteUInt16(table, 0, 1);
        WriteUInt16(table, 2, 2);
        WriteUInt16(table, 4, 2);
        WriteUInt16(table, 6, 4);
        WriteUInt32(table, 8, 28);
        WriteUInt16(table, 12, 0);
        WriteUInt16(table, 14, 2);
        WriteUInt32(table, 16, 44);
        WriteUInt32(table, 20, 0);
        WriteUInt32(table, 24, 0);
        WriteBgra(table, 28, 255, 0, 0, 255);
        WriteBgra(table, 32, 0, 0, 255, 255);
        WriteBgra(table, 36, 255, 255, 0, 255);
        WriteBgra(table, 40, 0, 128, 0, 255);
        WriteUInt32(table, 44, 1);
        WriteUInt32(table, 48, 2);
        return table;
    }

    private static void WriteBgra(byte[] data, int offset, byte red, byte green, byte blue, byte alpha) {
        data[offset] = blue;
        data[offset + 1] = green;
        data[offset + 2] = red;
        data[offset + 3] = alpha;
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

    internal static byte[] CreateWoff(byte[] openTypeData, bool compressTables = true) {
        if (openTypeData == null || openTypeData.Length < 12) throw new ArgumentException("OpenType data is required.", nameof(openTypeData));
        int tableCount = ReadUInt16(openTypeData, 4);
        int woffDirectoryLength = 44 + tableCount * 20;
        var records = new List<(uint Tag, byte[] Data, byte[] Encoded, uint Checksum, int Offset)>(tableCount);
        int offset = woffDirectoryLength;
        for (int index = 0; index < tableCount; index++) {
            int sfntRecord = 12 + index * 16;
            uint tag = ReadUInt32(openTypeData, sfntRecord);
            int tableOffset = checked((int)ReadUInt32(openTypeData, sfntRecord + 8));
            int tableLength = checked((int)ReadUInt32(openTypeData, sfntRecord + 12));
            var table = new byte[tableLength];
            Array.Copy(openTypeData, tableOffset, table, 0, tableLength);
            byte[] compressed = compressTables ? CompressZlib(table) : table;
            byte[] encoded = compressed.Length < table.Length ? compressed : table;
            records.Add((tag, table, encoded, CalculateTableChecksum(tag, table), offset));
            offset += Align4(encoded.Length);
        }

        var woff = new byte[offset];
        WriteUInt32(woff, 0, 0x774F4646);
        WriteUInt32(woff, 4, ReadUInt32(openTypeData, 0));
        WriteUInt32(woff, 8, (uint)woff.Length);
        WriteUInt16(woff, 12, (ushort)tableCount);
        WriteUInt32(woff, 16, (uint)openTypeData.Length);
        for (int index = 0; index < records.Count; index++) {
            var record = records[index];
            int directoryRecord = 44 + index * 20;
            WriteUInt32(woff, directoryRecord, record.Tag);
            WriteUInt32(woff, directoryRecord + 4, (uint)record.Offset);
            WriteUInt32(woff, directoryRecord + 8, (uint)record.Encoded.Length);
            WriteUInt32(woff, directoryRecord + 12, (uint)record.Data.Length);
            WriteUInt32(woff, directoryRecord + 16, record.Checksum);
            Array.Copy(record.Encoded, 0, woff, record.Offset, record.Encoded.Length);
        }
        return woff;
    }

    private static byte[] CompressZlib(byte[] data) {
        using var output = new MemoryStream();
        output.WriteByte(0x78);
        output.WriteByte(0x9C);
        using (var deflate = new DeflateStream(output, CompressionLevel.Optimal, leaveOpen: true)) {
            deflate.Write(data, 0, data.Length);
        }
        uint checksum = Adler32(data);
        output.WriteByte((byte)(checksum >> 24));
        output.WriteByte((byte)(checksum >> 16));
        output.WriteByte((byte)(checksum >> 8));
        output.WriteByte((byte)checksum);
        return output.ToArray();
    }

    private static uint Adler32(byte[] data) {
        const uint modulus = 65521;
        uint a = 1;
        uint b = 0;
        foreach (byte value in data) {
            a = (a + value) % modulus;
            b = (b + a) % modulus;
        }
        return (b << 16) | a;
    }

    private static uint CalculateChecksum(byte[] data) {
        uint checksum = 0;
        for (int offset = 0; offset < data.Length; offset += 4) {
            uint value = (uint)data[offset] << 24;
            if (offset + 1 < data.Length) value |= (uint)data[offset + 1] << 16;
            if (offset + 2 < data.Length) value |= (uint)data[offset + 2] << 8;
            if (offset + 3 < data.Length) value |= data[offset + 3];
            checksum = unchecked(checksum + value);
        }
        return checksum;
    }

    private static uint CalculateTableChecksum(uint tag, byte[] data) {
        if (tag != 0x68656164 || data.Length < 12) return CalculateChecksum(data);
        var normalized = (byte[])data.Clone();
        WriteUInt32(normalized, 8, 0);
        return CalculateChecksum(normalized);
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

    private static byte[] CreateUnicodeCmapWithVariationSequence(
        int scalar,
        int variationSelector,
        int? nonDefaultGlyph,
        int variationPlatform,
        int variationEncoding,
        int nonDefaultMappingCount = 1) {
        const int cmapHeaderLength = 20;
        const int format12Length = 28;
        if (nonDefaultMappingCount <= 0) throw new ArgumentOutOfRangeException(nameof(nonDefaultMappingCount));
        int format14Length = nonDefaultGlyph.HasValue ? checked(25 + nonDefaultMappingCount * 5) : 29;
        int format12 = cmapHeaderLength;
        int format14 = format12 + format12Length;
        var data = new byte[cmapHeaderLength + format12Length + format14Length];
        WriteUInt16(data, 2, 2);

        WriteUInt16(data, 4, 3);
        WriteUInt16(data, 6, 10);
        WriteUInt32(data, 8, (uint)format12);
        WriteUInt16(data, 12, checked((ushort)variationPlatform));
        WriteUInt16(data, 14, checked((ushort)variationEncoding));
        WriteUInt32(data, 16, (uint)format14);

        WriteUInt16(data, format12, 12);
        WriteUInt32(data, format12 + 4, format12Length);
        WriteUInt32(data, format12 + 12, 1);
        WriteUInt32(data, format12 + 16, checked((uint)scalar));
        WriteUInt32(data, format12 + 20, checked((uint)scalar));
        WriteUInt32(data, format12 + 24, 1);

        WriteUInt16(data, format14, 14);
        WriteUInt32(data, format14 + 2, (uint)format14Length);
        WriteUInt32(data, format14 + 6, 1);
        WriteUInt24(data, format14 + 10, variationSelector);
        WriteUInt32(data, format14 + 13, nonDefaultGlyph.HasValue ? 0U : 21U);
        WriteUInt32(data, format14 + 17, nonDefaultGlyph.HasValue ? 21U : 0U);
        WriteUInt32(data, format14 + 21, nonDefaultGlyph.HasValue ? checked((uint)nonDefaultMappingCount) : 1U);
        if (nonDefaultGlyph.HasValue) {
            int firstScalar = checked(scalar - nonDefaultMappingCount + 1);
            for (int index = 0; index < nonDefaultMappingCount; index++) {
                int mapping = format14 + 25 + index * 5;
                WriteUInt24(data, mapping, firstScalar + index);
                WriteUInt16(data, mapping + 3, checked((ushort)nonDefaultGlyph.Value));
            }
        } else {
            WriteUInt24(data, format14 + 25, scalar);
            data[format14 + 28] = 0;
        }
        return data;
    }

    private static byte[] CreateFormat12Cmap(
        int firstScalar,
        int firstGlyph,
        int secondScalar,
        int secondGlyph) {
        var mappings = new[] {
            (Scalar: firstScalar, Glyph: firstGlyph),
            (Scalar: secondScalar, Glyph: secondGlyph)
        };
        Array.Sort(mappings, static (left, right) => left.Scalar.CompareTo(right.Scalar));
        var data = new byte[52];
        WriteUInt16(data, 2, 1);
        WriteUInt16(data, 4, 3);
        WriteUInt16(data, 6, 10);
        WriteUInt32(data, 8, 12);
        WriteUInt16(data, 12, 12);
        WriteUInt32(data, 16, 40);
        WriteUInt32(data, 24, 2);
        for (int index = 0; index < mappings.Length; index++) {
            int offset = 28 + (index * 12);
            WriteUInt32(data, offset, checked((uint)mappings[index].Scalar));
            WriteUInt32(data, offset + 4, checked((uint)mappings[index].Scalar));
            WriteUInt32(data, offset + 8, checked((uint)mappings[index].Glyph));
        }
        return data;
    }

    private static byte[] CreateKernTable(ushort leftGlyph, ushort rightGlyph, short adjustment) {
        var data = new byte[24];
        WriteUInt16(data, 2, 1);
        WriteUInt16(data, 6, 20);
        WriteUInt16(data, 8, 1);
        WriteUInt16(data, 10, 1);
        WriteUInt16(data, 12, 6);
        WriteUInt16(data, 18, leftGlyph);
        WriteUInt16(data, 20, rightGlyph);
        WriteUInt16(data, 22, unchecked((ushort)adjustment));
        return data;
    }

    private static byte[] CreateLigatureGsub(string featureTag, ushort firstGlyph, ushort secondGlyph, ushort ligatureGlyph) {
        if (featureTag == null || featureTag.Length != 4) throw new ArgumentException("Feature tags must contain four characters.", nameof(featureTag));
        var data = new byte[62];
        WriteUInt32(data, 0, 0x00010000);
        WriteUInt16(data, 4, 10);
        WriteUInt16(data, 6, 12);
        WriteUInt16(data, 8, 26);
        WriteUInt16(data, 10, 0);
        WriteUInt16(data, 12, 1);
        WriteTag(data, 14, featureTag);
        WriteUInt16(data, 18, 8);
        WriteUInt16(data, 20, 0);
        WriteUInt16(data, 22, 1);
        WriteUInt16(data, 24, 0);
        WriteUInt16(data, 26, 1);
        WriteUInt16(data, 28, 4);
        WriteUInt16(data, 30, 4);
        WriteUInt16(data, 32, 0);
        WriteUInt16(data, 34, 1);
        WriteUInt16(data, 36, 8);
        WriteUInt16(data, 38, 1);
        WriteUInt16(data, 40, 18);
        WriteUInt16(data, 42, 1);
        WriteUInt16(data, 44, 8);
        WriteUInt16(data, 46, 1);
        WriteUInt16(data, 48, 4);
        WriteUInt16(data, 50, ligatureGlyph);
        WriteUInt16(data, 52, 2);
        WriteUInt16(data, 54, secondGlyph);
        WriteUInt16(data, 56, 1);
        WriteUInt16(data, 58, 1);
        WriteUInt16(data, 60, firstGlyph);
        return data;
    }

    private static byte[] CreateUnicodeCmapWithFallback(int bmpScalar, int supplementalScalar) {
        byte[] format12 = CreateFormat12Cmap(new[] { supplementalScalar });
        const int cmapHeaderLength = 20;
        const int format4Length = 32;
        var data = new byte[cmapHeaderLength + (format12.Length - 12) + format4Length];
        WriteUInt16(data, 2, 2);
        WriteUInt16(data, 4, 3);
        WriteUInt16(data, 6, 10);
        WriteUInt32(data, 8, cmapHeaderLength);
        WriteUInt16(data, 12, 3);
        WriteUInt16(data, 14, 1);
        int format4Offset = cmapHeaderLength + format12.Length - 12;
        WriteUInt32(data, 16, (uint)format4Offset);
        Array.Copy(format12, 12, data, cmapHeaderLength, format12.Length - 12);

        WriteUInt16(data, format4Offset, 4);
        WriteUInt16(data, format4Offset + 2, format4Length);
        WriteUInt16(data, format4Offset + 6, 4);
        WriteUInt16(data, format4Offset + 8, 4);
        WriteUInt16(data, format4Offset + 10, 1);
        WriteUInt16(data, format4Offset + 14, (ushort)bmpScalar);
        WriteUInt16(data, format4Offset + 16, 0xFFFF);
        WriteUInt16(data, format4Offset + 20, (ushort)bmpScalar);
        WriteUInt16(data, format4Offset + 22, 0xFFFF);
        WriteUInt16(data, format4Offset + 24, unchecked((ushort)(1 - bmpScalar)));
        WriteUInt16(data, format4Offset + 26, 1);
        return data;
    }

    private static byte[] CreateConflictingUnicodeCmapFallback(int bmpScalar, int supplementalScalar) {
        const int recordCount = 3;
        const int headerLength = 4 + recordCount * 8;
        const int format12Length = 28;
        const int format4Length = 32;
        int selectedFormat12 = headerLength;
        int format4 = selectedFormat12 + format12Length;
        int unicodeFormat12 = format4 + format4Length;
        var data = new byte[unicodeFormat12 + format12Length];
        WriteUInt16(data, 2, recordCount);

        WriteUInt16(data, 4, 3);
        WriteUInt16(data, 6, 10);
        WriteUInt32(data, 8, (uint)selectedFormat12);
        WriteUInt16(data, 12, 3);
        WriteUInt16(data, 14, 1);
        WriteUInt32(data, 16, (uint)format4);
        WriteUInt16(data, 20, 0);
        WriteUInt16(data, 22, 4);
        WriteUInt32(data, 24, (uint)unicodeFormat12);

        WriteFormat12Subtable(data, selectedFormat12, supplementalScalar, 1);
        WriteFormat4Subtable(data, format4, bmpScalar, 1);
        WriteFormat12Subtable(data, unicodeFormat12, bmpScalar, 2);
        return data;
    }

    private static void WriteFormat12Subtable(byte[] data, int offset, int scalar, int glyph) {
        WriteUInt16(data, offset, 12);
        WriteUInt32(data, offset + 4, 28);
        WriteUInt32(data, offset + 12, 1);
        WriteUInt32(data, offset + 16, (uint)scalar);
        WriteUInt32(data, offset + 20, (uint)scalar);
        WriteUInt32(data, offset + 24, (uint)glyph);
    }

    private static void WriteFormat4Subtable(byte[] data, int offset, int scalar, int glyph) {
        WriteUInt16(data, offset, 4);
        WriteUInt16(data, offset + 2, 32);
        WriteUInt16(data, offset + 6, 4);
        WriteUInt16(data, offset + 8, 4);
        WriteUInt16(data, offset + 10, 1);
        WriteUInt16(data, offset + 14, (ushort)scalar);
        WriteUInt16(data, offset + 16, 0xFFFF);
        WriteUInt16(data, offset + 20, (ushort)scalar);
        WriteUInt16(data, offset + 22, 0xFFFF);
        WriteUInt16(data, offset + 24, unchecked((ushort)(glyph - scalar)));
        WriteUInt16(data, offset + 26, 1);
    }

    private static byte[] CreateVisibleGlyph(int width) {
        var glyph = new byte[34];
        WriteUInt16(glyph, 0, 1);
        WriteUInt16(glyph, 6, checked((ushort)width));
        WriteUInt16(glyph, 8, 700);
        WriteUInt16(glyph, 10, 3);
        glyph[14] = 0x01;
        glyph[15] = 0x01;
        glyph[16] = 0x01;
        glyph[17] = 0x01;
        WriteUInt16(glyph, 20, checked((ushort)width));
        WriteUInt16(glyph, 24, unchecked((ushort)-width));
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

    private static void WriteUInt24(byte[] data, int offset, int value) {
        data[offset] = (byte)(value >> 16);
        data[offset + 1] = (byte)(value >> 8);
        data[offset + 2] = (byte)value;
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
