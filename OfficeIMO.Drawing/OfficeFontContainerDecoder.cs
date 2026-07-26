using System;
using System.Collections.Generic;
using System.IO;
using OfficeIMO.Drawing.Internal;

namespace OfficeIMO.Drawing;

/// <summary>Detects and safely unwraps reusable web-font containers into OpenType bytes.</summary>
public static class OfficeFontContainerDecoder {
    private const uint WoffSignature = 0x774F4646;
    private const uint Woff2Signature = 0x774F4632;
    private const int WoffHeaderLength = 44;
    private const int WoffTableRecordLength = 20;
    private const int SfntHeaderLength = 12;
    private const int SfntTableRecordLength = 16;
    private const int MaximumTableCount = 512;
    private const int DefaultMaximumDecodedBytes = 128 * 1024 * 1024;
    private const uint HeadTableTag = 0x68656164;
    private const uint OpenTypeChecksumMagic = 0xB1B0AFBA;

    /// <summary>Detects a supported or known font container without decoding it.</summary>
    public static OfficeFontContainerFormat Detect(byte[]? data) {
        if (data == null || data.Length < 4) return OfficeFontContainerFormat.Unknown;
        uint signature = ReadUInt32(data, 0);
        if (signature == WoffSignature) return OfficeFontContainerFormat.Woff;
        if (signature == Woff2Signature) return OfficeFontContainerFormat.Woff2;
        if (signature == 0x00010000
            || signature == 0x74727565
            || signature == 0x4F54544F
            || signature == 0x74746366) {
            return OfficeFontContainerFormat.OpenType;
        }
        return OfficeFontContainerFormat.Unknown;
    }

    /// <summary>
    /// Attempts to decode a direct OpenType or WOFF 1 container using the default 128 MiB output limit.
    /// WOFF 2 is detected but returns a clear unsupported result until its transformed-table decoder is available.
    /// </summary>
    public static bool TryDecodeToOpenType(
        byte[]? data,
        out byte[] openTypeData,
        out OfficeFontContainerFormat format,
        out string? error) =>
        TryDecodeToOpenType(data, DefaultMaximumDecodedBytes, out openTypeData, out format, out error);

    /// <summary>Attempts to decode a direct OpenType or WOFF 1 container with an explicit output limit.</summary>
    public static bool TryDecodeToOpenType(
        byte[]? data,
        int maximumDecodedBytes,
        out byte[] openTypeData,
        out OfficeFontContainerFormat format,
        out string? error) {
        openTypeData = Array.Empty<byte>();
        format = Detect(data);
        error = null;
        if (maximumDecodedBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maximumDecodedBytes));
        if (data == null || data.Length == 0) {
            error = "Font data is empty.";
            return false;
        }
        if (format == OfficeFontContainerFormat.OpenType) {
            if (data.Length > maximumDecodedBytes) {
                error = "OpenType font data exceeds the configured decoded byte limit.";
                return false;
            }
            openTypeData = (byte[])data.Clone();
            return true;
        }
        if (format == OfficeFontContainerFormat.Woff2) {
            error = "WOFF 2 transformed-table decoding is not supported.";
            return false;
        }
        if (format != OfficeFontContainerFormat.Woff) {
            error = "Font data is not a recognized OpenType or web-font container.";
            return false;
        }

        try {
            openTypeData = DecodeWoff(data, maximumDecodedBytes);
            return true;
        } catch (Exception exception) when (exception is InvalidDataException
                                            || exception is NotSupportedException
                                            || exception is OverflowException
                                            || exception is ArgumentOutOfRangeException
                                            || exception is IndexOutOfRangeException) {
            openTypeData = Array.Empty<byte>();
            error = exception.Message;
            return false;
        }
    }

    private static byte[] DecodeWoff(byte[] data, int maximumDecodedBytes) {
        if (data.Length < WoffHeaderLength) throw new InvalidDataException("The WOFF header is truncated.");
        uint declaredLength = ReadUInt32(data, 8);
        if (declaredLength != data.Length) throw new InvalidDataException("The WOFF length does not match the supplied data.");
        int tableCount = ReadUInt16(data, 12);
        if (tableCount <= 0 || tableCount > MaximumTableCount) throw new InvalidDataException("The WOFF table count is invalid.");
        if (ReadUInt16(data, 14) != 0) throw new InvalidDataException("The WOFF reserved header field must be zero.");
        int directoryEnd = checked(WoffHeaderLength + tableCount * WoffTableRecordLength);
        if (directoryEnd > data.Length) throw new InvalidDataException("The WOFF table directory is truncated.");

        uint declaredSfntSize = ReadUInt32(data, 16);
        if (declaredSfntSize > maximumDecodedBytes || declaredSfntSize > int.MaxValue) {
            throw new InvalidDataException("The decoded WOFF font exceeds the configured byte limit.");
        }

        var records = new List<WoffTableRecord>(tableCount);
        var tags = new HashSet<uint>();
        var intervals = new List<(int Start, int End)>(tableCount);
        int decodedTableBytes = checked(SfntHeaderLength + tableCount * SfntTableRecordLength);
        uint previousTag = 0;
        for (int index = 0; index < tableCount; index++) {
            int recordOffset = WoffHeaderLength + index * WoffTableRecordLength;
            uint tag = ReadUInt32(data, recordOffset);
            uint sourceOffsetValue = ReadUInt32(data, recordOffset + 4);
            uint compressedLengthValue = ReadUInt32(data, recordOffset + 8);
            uint originalLengthValue = ReadUInt32(data, recordOffset + 12);
            uint checksum = ReadUInt32(data, recordOffset + 16);
            if (!tags.Add(tag)) throw new InvalidDataException("The WOFF table directory contains a duplicate tag.");
            if (index > 0 && tag <= previousTag) {
                throw new InvalidDataException("The WOFF table directory is not sorted by tag.");
            }
            previousTag = tag;
            if (sourceOffsetValue > int.MaxValue || compressedLengthValue > int.MaxValue || originalLengthValue > int.MaxValue) {
                throw new InvalidDataException("A WOFF table offset or length is too large.");
            }
            int sourceOffset = (int)sourceOffsetValue;
            int compressedLength = (int)compressedLengthValue;
            int originalLength = (int)originalLengthValue;
            if (sourceOffset < directoryEnd || sourceOffset % 4 != 0 || compressedLength <= 0 || originalLength <= 0
                || compressedLength > originalLength || sourceOffset > data.Length - compressedLength) {
                throw new InvalidDataException("A WOFF table record has an invalid offset or length.");
            }
            int alignedLength = Align4(originalLength);
            if (decodedTableBytes > maximumDecodedBytes - alignedLength) {
                throw new InvalidDataException("The decoded WOFF font exceeds the configured byte limit.");
            }
            records.Add(new WoffTableRecord(tag, sourceOffset, compressedLength, originalLength, checksum));
            intervals.Add((sourceOffset, sourceOffset + compressedLength));
            decodedTableBytes += alignedLength;
        }

        intervals.Sort((left, right) => left.Start.CompareTo(right.Start));
        for (int index = 1; index < intervals.Count; index++) {
            if (intervals[index].Start < intervals[index - 1].End) {
                throw new InvalidDataException("WOFF table payloads overlap.");
            }
        }
        if (decodedTableBytes != declaredSfntSize) throw new InvalidDataException("The WOFF decoded size does not match its table directory.");

        var outputOffsets = new Dictionary<uint, int>(records.Count);
        int nextOutputOffset = SfntHeaderLength + tableCount * SfntTableRecordLength;
        var physicalRecords = new List<WoffTableRecord>(records);
        physicalRecords.Sort((left, right) => left.SourceOffset.CompareTo(right.SourceOffset));
        foreach (WoffTableRecord record in physicalRecords) {
            outputOffsets[record.Tag] = nextOutputOffset;
            nextOutputOffset += Align4(record.OriginalLength);
        }

        var output = new byte[decodedTableBytes];
        WriteUInt32(output, 0, ReadUInt32(data, 4));
        WriteUInt16(output, 4, (ushort)tableCount);
        int maximumPowerOfTwo = 1;
        int entrySelector = 0;
        while (maximumPowerOfTwo * 2 <= tableCount) {
            maximumPowerOfTwo *= 2;
            entrySelector++;
        }
        int searchRange = maximumPowerOfTwo * 16;
        WriteUInt16(output, 6, (ushort)searchRange);
        WriteUInt16(output, 8, (ushort)entrySelector);
        WriteUInt16(output, 10, (ushort)(tableCount * 16 - searchRange));

        int headTableOffset = -1;
        for (int index = 0; index < records.Count; index++) {
            WoffTableRecord record = records[index];
            byte[] table;
            if (record.CompressedLength == record.OriginalLength) {
                table = new byte[record.OriginalLength];
                Buffer.BlockCopy(data, record.SourceOffset, table, 0, table.Length);
            } else {
                var compressed = new byte[record.CompressedLength];
                Buffer.BlockCopy(data, record.SourceOffset, compressed, 0, compressed.Length);
                table = OfficeZlibCodec.Decompress(compressed, record.OriginalLength, record.OriginalLength);
            }
            bool isHeadTable = record.Tag == HeadTableTag;
            if (isHeadTable && table.Length < 12) {
                throw new InvalidDataException("The WOFF head table is truncated.");
            }
            if (CalculateChecksum(table, isHeadTable) != record.Checksum) {
                throw new InvalidDataException("A WOFF table checksum is invalid.");
            }
            int outputOffset = outputOffsets[record.Tag];
            int sfntRecord = SfntHeaderLength + index * SfntTableRecordLength;
            WriteUInt32(output, sfntRecord, record.Tag);
            WriteUInt32(output, sfntRecord + 4, record.Checksum);
            WriteUInt32(output, sfntRecord + 8, (uint)outputOffset);
            WriteUInt32(output, sfntRecord + 12, (uint)record.OriginalLength);
            Buffer.BlockCopy(table, 0, output, outputOffset, table.Length);
            if (isHeadTable) {
                headTableOffset = outputOffset;
                WriteUInt32(output, headTableOffset + 8, 0);
            }
        }
        if (headTableOffset >= 0) {
            uint adjustment = unchecked(OpenTypeChecksumMagic - CalculateChecksum(output));
            WriteUInt32(output, headTableOffset + 8, adjustment);
        }
        return output;
    }

    private static uint CalculateChecksum(byte[] data, bool clearHeadChecksumAdjustment = false) {
        uint checksum = 0;
        for (int offset = 0; offset < data.Length; offset += 4) {
            uint value = clearHeadChecksumAdjustment && offset == 8
                ? 0
                : (uint)data[offset] << 24;
            if (clearHeadChecksumAdjustment && offset == 8) {
                checksum = unchecked(checksum + value);
                continue;
            }
            if (offset + 1 < data.Length) value |= (uint)data[offset + 1] << 16;
            if (offset + 2 < data.Length) value |= (uint)data[offset + 2] << 8;
            if (offset + 3 < data.Length) value |= data[offset + 3];
            checksum = unchecked(checksum + value);
        }
        return checksum;
    }

    private static int Align4(int value) => checked((value + 3) & ~3);

    private static ushort ReadUInt16(byte[] data, int offset) =>
        unchecked((ushort)((data[offset] << 8) | data[offset + 1]));

    private static uint ReadUInt32(byte[] data, int offset) =>
        unchecked(((uint)data[offset] << 24)
            | ((uint)data[offset + 1] << 16)
            | ((uint)data[offset + 2] << 8)
            | data[offset + 3]);

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

    private readonly struct WoffTableRecord {
        internal WoffTableRecord(
            uint tag,
            int sourceOffset,
            int compressedLength,
            int originalLength,
            uint checksum) {
            Tag = tag;
            SourceOffset = sourceOffset;
            CompressedLength = compressedLength;
            OriginalLength = originalLength;
            Checksum = checksum;
        }

        internal uint Tag { get; }
        internal int SourceOffset { get; }
        internal int CompressedLength { get; }
        internal int OriginalLength { get; }
        internal uint Checksum { get; }
    }
}
