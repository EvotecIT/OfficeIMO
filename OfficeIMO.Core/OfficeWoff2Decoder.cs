using System;
using System.Collections.Generic;
using System.IO;
#if NET8_0_OR_GREATER
using System.IO.Compression;
#endif

namespace OfficeIMO.Drawing;

/// <summary>
/// Bounded WOFF 2 decoder implementing the W3C table directory and transformed-table contract.
/// </summary>
internal static partial class OfficeWoff2Decoder {
    private const uint Signature = 0x774F4632;
    private const uint CollectionFlavor = 0x74746366;
    private const int HeaderLength = 48;
    private const int MaximumTableCount = 512;
    private const uint HeadTag = 0x68656164;
    private const uint HheaTag = 0x68686561;
    private const uint HmtxTag = 0x686D7478;
    private const uint MaxpTag = 0x6D617870;
    private const uint GlyfTag = 0x676C7966;
    private const uint LocaTag = 0x6C6F6361;
    private const uint ChecksumMagic = 0xB1B0AFBA;

    private static readonly uint[] KnownTags = {
        Tag("cmap"), Tag("head"), Tag("hhea"), Tag("hmtx"), Tag("maxp"), Tag("name"), Tag("OS/2"), Tag("post"),
        Tag("cvt "), Tag("fpgm"), Tag("glyf"), Tag("loca"), Tag("prep"), Tag("CFF "), Tag("VORG"), Tag("EBDT"),
        Tag("EBLC"), Tag("gasp"), Tag("hdmx"), Tag("kern"), Tag("LTSH"), Tag("PCLT"), Tag("VDMX"), Tag("vhea"),
        Tag("vmtx"), Tag("BASE"), Tag("GDEF"), Tag("GPOS"), Tag("GSUB"), Tag("EBSC"), Tag("JSTF"), Tag("MATH"),
        Tag("CBDT"), Tag("CBLC"), Tag("COLR"), Tag("CPAL"), Tag("SVG "), Tag("sbix"), Tag("acnt"), Tag("avar"),
        Tag("bdat"), Tag("bloc"), Tag("bsln"), Tag("cvar"), Tag("fdsc"), Tag("feat"), Tag("fmtx"), Tag("fvar"),
        Tag("gvar"), Tag("hsty"), Tag("just"), Tag("lcar"), Tag("mort"), Tag("morx"), Tag("opbd"), Tag("prop"),
        Tag("trak"), Tag("Zapf"), Tag("Silf"), Tag("Glat"), Tag("Gloc"), Tag("Feat"), Tag("Sill")
    };

    internal static byte[] Decode(byte[] data, int maximumDecodedBytes) {
#if !NET8_0_OR_GREATER
        throw new PlatformNotSupportedException("WOFF 2 decoding requires the .NET 8 or newer OfficeIMO.Core asset.");
#else
        if (data == null) throw new ArgumentNullException(nameof(data));
        if (maximumDecodedBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maximumDecodedBytes));
        if (data.Length < HeaderLength) throw new InvalidDataException("The WOFF 2 header is truncated.");
        if (ReadUInt32(data, 0) != Signature) throw new InvalidDataException("The font is not a WOFF 2 container.");
        uint flavor = ReadUInt32(data, 4);
        if (flavor == CollectionFlavor) {
            throw new NotSupportedException("WOFF 2 font collections are not supported; register an individual font face instead.");
        }
        uint declaredLength = ReadUInt32(data, 8);
        if (declaredLength != data.Length) throw new InvalidDataException("The WOFF 2 length does not match the supplied data.");
        int tableCount = ReadUInt16(data, 12);
        if (tableCount <= 0 || tableCount > MaximumTableCount) throw new InvalidDataException("The WOFF 2 table count is invalid.");
        // Encoders are required to write zero, but the WOFF2 decoder conformance contract
        // explicitly requires accepting a non-zero reserved header value.
        // WOFF2 totalSfntSize is reference-only. The W3C contract explicitly forbids
        // rejecting a correctly decoded font when reconstruction produces a different size.
        // Actual table and output allocations remain bounded by maximumDecodedBytes below.
        _ = ReadUInt32(data, 16);
        uint compressedSizeValue = ReadUInt32(data, 20);
        if (compressedSizeValue == 0 || compressedSizeValue > int.MaxValue) {
            throw new InvalidDataException("The WOFF 2 compressed payload length is invalid.");
        }

        int cursor = HeaderLength;
        var records = new List<TableRecord>(tableCount);
        var tags = new HashSet<uint>();
        long transformedByteCount = 0;
        for (int index = 0; index < tableCount; index++) {
            EnsureAvailable(data, cursor, 1, "The WOFF 2 table directory is truncated.");
            byte flags = data[cursor++];
            int tagIndex = flags & 0x3F;
            uint tag;
            if (tagIndex == 0x3F) {
                EnsureAvailable(data, cursor, 4, "The WOFF 2 custom table tag is truncated.");
                tag = ReadUInt32(data, cursor);
                cursor += 4;
            } else {
                tag = KnownTags[tagIndex];
            }
            if (!tags.Add(tag)) throw new InvalidDataException("The WOFF 2 table directory contains a duplicate tag.");
            uint originalLength = ReadBase128(data, ref cursor);
            if (originalLength == 0 || originalLength > int.MaxValue) {
                throw new InvalidDataException("A WOFF 2 table has an invalid original length.");
            }
            int transformVersion = flags >> 6;
            if (tag == GlyfTag || tag == LocaTag) {
                if (transformVersion != 0 && transformVersion != 3) {
                    throw new NotSupportedException("The WOFF 2 glyf/loca transform version is not supported.");
                }
            } else if (tag == HmtxTag) {
                if (transformVersion != 0 && transformVersion != 1) {
                    throw new NotSupportedException("The WOFF 2 hmtx transform version is not supported.");
                }
            } else if (transformVersion != 0) {
                throw new NotSupportedException("The WOFF 2 table uses an unknown transform version.");
            }
            bool transformed = tag == GlyfTag || tag == LocaTag
                ? transformVersion != 3
                : transformVersion != 0;
            uint payloadLength = originalLength;
            if (transformed) {
                payloadLength = ReadBase128(data, ref cursor);
                if (payloadLength > int.MaxValue) throw new InvalidDataException("A WOFF 2 transformed table is too large.");
                if (tag == LocaTag && payloadLength != 0) {
                    throw new InvalidDataException("The transformed WOFF 2 loca table must have zero payload length.");
                }
            }
            transformedByteCount = checked(transformedByteCount + payloadLength);
            if (transformedByteCount > maximumDecodedBytes) {
                throw new InvalidDataException("The expanded WOFF 2 table stream exceeds the configured byte limit.");
            }
            records.Add(new TableRecord(tag, checked((int)originalLength), checked((int)payloadLength), transformed, transformVersion));
        }
        int glyfIndex = FindRecordIndex(records, GlyfTag);
        int locaIndex = FindRecordIndex(records, LocaTag);
        if (glyfIndex >= 0 && locaIndex >= 0 && locaIndex < glyfIndex) {
            throw new InvalidDataException("The WOFF 2 loca table must follow the glyf table.");
        }

        int compressedSize = checked((int)compressedSizeValue);
        EnsureAvailable(data, cursor, compressedSize, "The WOFF 2 compressed payload is truncated.");
        ValidateTrailingBlocks(data, cursor, compressedSize);
        byte[] transformedData = Decompress(data, cursor, compressedSize, checked((int)transformedByteCount), maximumDecodedBytes);
        int transformedOffset = 0;
        var tables = new Dictionary<uint, byte[]>(records.Count);
        foreach (TableRecord record in records) {
            if (record.PayloadLength == 0) continue;
            EnsureAvailable(transformedData, transformedOffset, record.PayloadLength, "The WOFF 2 table stream is truncated.");
            var table = new byte[record.PayloadLength];
            Buffer.BlockCopy(transformedData, transformedOffset, table, 0, table.Length);
            transformedOffset += table.Length;
            tables.Add(record.Tag, table);
        }
        if (transformedOffset != transformedData.Length) {
            throw new InvalidDataException("The WOFF 2 table stream contains trailing data.");
        }

        ReconstructTransformedTables(records, tables, maximumDecodedBytes);
        byte[] sfnt = BuildSfnt(flavor, records, tables, maximumDecodedBytes);
        if (sfnt.Length > maximumDecodedBytes) throw new InvalidDataException("The decoded WOFF 2 font exceeds the configured byte limit.");
        return sfnt;
#endif
    }

    private static void ValidateTrailingBlocks(byte[] data, int compressedOffset, int compressedLength) {
        int compressedEnd = checked(compressedOffset + compressedLength);
        uint metadataOffsetValue = ReadUInt32(data, 28);
        uint metadataLengthValue = ReadUInt32(data, 32);
        uint metadataOriginalLength = ReadUInt32(data, 36);
        uint privateOffsetValue = ReadUInt32(data, 40);
        uint privateLengthValue = ReadUInt32(data, 44);
        if (metadataOffsetValue > int.MaxValue || metadataLengthValue > int.MaxValue
            || privateOffsetValue > int.MaxValue || privateLengthValue > int.MaxValue) {
            throw new InvalidDataException("A WOFF 2 trailing-block offset or length is invalid.");
        }

        int metadataOffset = (int)metadataOffsetValue;
        int metadataLength = (int)metadataLengthValue;
        int privateOffset = (int)privateOffsetValue;
        int privateLength = (int)privateLengthValue;
        bool hasMetadata = metadataOffset != 0;
        bool hasPrivateData = privateOffset != 0;
        if (hasMetadata != (metadataLength != 0) || hasMetadata != (metadataOriginalLength != 0)) {
            throw new InvalidDataException("The WOFF 2 metadata block declaration is inconsistent.");
        }
        if (hasPrivateData != (privateLength != 0)) {
            throw new InvalidDataException("The WOFF 2 private-data block declaration is inconsistent.");
        }

        int precedingEnd = compressedEnd;
        if (hasMetadata) {
            int expectedOffset = Align4(precedingEnd);
            if (metadataOffset != expectedOffset || metadataOffset > data.Length - metadataLength) {
                throw new InvalidDataException("The WOFF 2 metadata block is misplaced or truncated.");
            }
            EnsureZeroPadding(data, precedingEnd, metadataOffset);
            precedingEnd = checked(metadataOffset + metadataLength);
        }
        if (hasPrivateData) {
            int expectedOffset = Align4(precedingEnd);
            if (privateOffset != expectedOffset || privateOffset > data.Length - privateLength) {
                throw new InvalidDataException("The WOFF 2 private-data block is misplaced or truncated.");
            }
            EnsureZeroPadding(data, precedingEnd, privateOffset);
            precedingEnd = checked(privateOffset + privateLength);
        }
        if (!hasMetadata && !hasPrivateData) {
            int alignedEnd = Align4(precedingEnd);
            EnsureZeroPadding(data, precedingEnd, alignedEnd);
            precedingEnd = alignedEnd;
        }
        if (precedingEnd != data.Length) {
            throw new InvalidDataException("The WOFF 2 file contains extraneous trailing data.");
        }
    }

    private static void EnsureZeroPadding(byte[] data, int start, int end) {
        if (start < 0 || end < start || end > data.Length || end - start > 3) {
            throw new InvalidDataException("The WOFF 2 trailing-block padding is invalid.");
        }
        for (int index = start; index < end; index++) {
            if (data[index] != 0) throw new InvalidDataException("The WOFF 2 trailing-block padding must contain null bytes.");
        }
    }

#if NET8_0_OR_GREATER
    private static byte[] Decompress(
        byte[] data,
        int offset,
        int length,
        int expectedLength,
        int maximumDecodedBytes) {
        using var source = new MemoryStream(data, offset, length, writable: false);
        using var brotli = new BrotliStream(source, CompressionMode.Decompress, leaveOpen: false);
        using var output = new MemoryStream(Math.Min(expectedLength, maximumDecodedBytes));
        var buffer = new byte[81920];
        while (true) {
            int read = brotli.Read(buffer, 0, buffer.Length);
            if (read == 0) break;
            if (output.Length > maximumDecodedBytes - read || output.Length > expectedLength - read) {
                throw new InvalidDataException("The expanded WOFF 2 table stream exceeds its declared size.");
            }
            output.Write(buffer, 0, read);
        }
        if (output.Length != expectedLength) throw new InvalidDataException("The WOFF 2 Brotli stream length is invalid.");
        return output.ToArray();
    }
#endif

    private static void ReconstructTransformedTables(
        IReadOnlyList<TableRecord> records,
        Dictionary<uint, byte[]> tables,
        int maximumDecodedBytes) {
        TableRecord? glyfRecord = FindRecord(records, GlyfTag);
        TableRecord? locaRecord = FindRecord(records, LocaTag);
        if (glyfRecord.HasValue && glyfRecord.Value.Transformed) {
            if (!locaRecord.HasValue || !locaRecord.Value.Transformed) {
                throw new InvalidDataException("WOFF 2 glyf and loca tables must use matching transforms.");
            }
            if (!tables.TryGetValue(GlyfTag, out byte[]? transformedGlyf)) {
                throw new InvalidDataException("The transformed WOFF 2 glyf table is missing.");
            }
            GlyfResult result = ReconstructGlyf(transformedGlyf, maximumDecodedBytes);
            if (result.Loca.Length != locaRecord.Value.OriginalLength) {
                throw new InvalidDataException("The reconstructed WOFF 2 loca table length is invalid.");
            }
            tables[GlyfTag] = result.Glyf;
            tables[LocaTag] = result.Loca;
            if (tables.TryGetValue(HeadTag, out byte[]? head)) {
                if (head.Length < 54) throw new InvalidDataException("The WOFF 2 head table is truncated.");
                WriteUInt16(head, 50, result.IndexFormat);
            }
        } else if (glyfRecord.HasValue != locaRecord.HasValue) {
            throw new InvalidDataException("WOFF 2 glyf and loca tables must be present together.");
        }

        TableRecord? hmtxRecord = FindRecord(records, HmtxTag);
        if (hmtxRecord.HasValue && hmtxRecord.Value.Transformed) {
            if (!tables.TryGetValue(HmtxTag, out byte[]? transformedHmtx)
                || !tables.TryGetValue(HheaTag, out byte[]? hhea)
                || !tables.TryGetValue(MaxpTag, out byte[]? maxp)
                || !tables.TryGetValue(GlyfTag, out byte[]? glyf)
                || !tables.TryGetValue(LocaTag, out byte[]? loca)
                || !tables.TryGetValue(HeadTag, out byte[]? head)) {
                throw new InvalidDataException("The transformed WOFF 2 hmtx table is missing required font tables.");
            }
            byte[] hmtx = ReconstructHmtx(transformedHmtx, hhea, maxp, head, glyf, loca);
            if (hmtx.Length != hmtxRecord.Value.OriginalLength) {
                throw new InvalidDataException("The reconstructed WOFF 2 hmtx table length is invalid.");
            }
            tables[HmtxTag] = hmtx;
        }
    }

    private static byte[] BuildSfnt(
        uint flavor,
        IReadOnlyList<TableRecord> records,
        Dictionary<uint, byte[]> tables,
        int maximumDecodedBytes) {
        var ordered = new List<uint>(records.Count);
        foreach (TableRecord record in records) {
            if (!tables.ContainsKey(record.Tag)) throw new InvalidDataException("A decoded WOFF 2 table is missing.");
            ordered.Add(record.Tag);
        }
        ordered.Sort();
        int directoryLength = checked(12 + ordered.Count * 16);
        int totalLength = directoryLength;
        foreach (uint tag in ordered) {
            totalLength = checked(totalLength + Align4(tables[tag].Length));
            if (totalLength > maximumDecodedBytes) throw new InvalidDataException("The decoded WOFF 2 font exceeds the configured byte limit.");
        }

        var output = new byte[totalLength];
        WriteUInt32(output, 0, flavor);
        WriteUInt16(output, 4, checked((ushort)ordered.Count));
        int maximumPowerOfTwo = 1;
        int entrySelector = 0;
        while (maximumPowerOfTwo * 2 <= ordered.Count) {
            maximumPowerOfTwo *= 2;
            entrySelector++;
        }
        int searchRange = maximumPowerOfTwo * 16;
        WriteUInt16(output, 6, checked((ushort)searchRange));
        WriteUInt16(output, 8, checked((ushort)entrySelector));
        WriteUInt16(output, 10, checked((ushort)(ordered.Count * 16 - searchRange)));

        int tableOffset = directoryLength;
        int headOffset = -1;
        for (int index = 0; index < ordered.Count; index++) {
            uint tag = ordered[index];
            byte[] table = tables[tag];
            if (tag == HeadTag) {
                if (table.Length < 12) throw new InvalidDataException("The decoded WOFF 2 head table is truncated.");
                table = (byte[])table.Clone();
                WriteUInt32(table, 8, 0);
                headOffset = tableOffset;
            }
            int recordOffset = 12 + index * 16;
            WriteUInt32(output, recordOffset, tag);
            WriteUInt32(output, recordOffset + 4, CalculateChecksum(table));
            WriteUInt32(output, recordOffset + 8, checked((uint)tableOffset));
            WriteUInt32(output, recordOffset + 12, checked((uint)table.Length));
            Buffer.BlockCopy(table, 0, output, tableOffset, table.Length);
            tableOffset += Align4(table.Length);
        }
        if (headOffset >= 0) WriteUInt32(output, headOffset + 8, unchecked(ChecksumMagic - CalculateChecksum(output)));
        return output;
    }

    private static TableRecord? FindRecord(IReadOnlyList<TableRecord> records, uint tag) {
        for (int index = 0; index < records.Count; index++) {
            if (records[index].Tag == tag) return records[index];
        }
        return null;
    }

    private static int FindRecordIndex(IReadOnlyList<TableRecord> records, uint tag) {
        for (int index = 0; index < records.Count; index++) {
            if (records[index].Tag == tag) return index;
        }
        return -1;
    }

    private readonly struct TableRecord {
        internal TableRecord(uint tag, int originalLength, int payloadLength, bool transformed, int transformVersion) {
            Tag = tag;
            OriginalLength = originalLength;
            PayloadLength = payloadLength;
            Transformed = transformed;
            TransformVersion = transformVersion;
        }

        internal uint Tag { get; }
        internal int OriginalLength { get; }
        internal int PayloadLength { get; }
        internal bool Transformed { get; }
        internal int TransformVersion { get; }
    }
}
