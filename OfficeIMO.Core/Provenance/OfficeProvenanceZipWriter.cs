using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace OfficeIMO.Provenance;

/// <summary>
/// Rewrites provenance ZIP containers with deterministic stored-entry behavior across target frameworks.
/// </summary>
internal static class OfficeProvenanceZipWriter {
    private const uint LocalHeaderSignature = 0x04034B50;
    private const uint CentralHeaderSignature = 0x02014B50;
    private const uint EndOfCentralDirectorySignature = 0x06054B50;
    private const uint Zip64EndOfCentralDirectorySignature = 0x06064B50;
    private const uint Zip64LocatorSignature = 0x07064B50;
    private const ushort Version20 = 20;
    private const ushort Version45 = 45;
    private const ushort UnixVersion20 = 0x0314;
    private const ushort Utf8FileNameFlag = 0x0800;
    private const ushort StoredMethod = 0;
    private const ushort DeflateMethod = 8;
    private const ushort UnicodePathExtraFieldId = 0x7075;
    private const ushort UnicodeCommentExtraFieldId = 0x6375;
    private static readonly uint[] CrcTable = CreateCrcTable();

    internal static byte[] Write(IReadOnlyList<OfficeProvenanceZipWriteEntry> entries, long maximumExpandedBytes, byte[]? archiveComment = null) {
        if (entries == null) throw new ArgumentNullException(nameof(entries));
        if (maximumExpandedBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maximumExpandedBytes));
        archiveComment ??= Array.Empty<byte>();
        if (archiveComment.Length > ushort.MaxValue) throw new ArgumentOutOfRangeException(nameof(archiveComment));
        var records = new List<OfficeProvenanceZipRecord>(entries.Count);
        using var output = new MemoryStream();
        using var writer = new BinaryWriter(output, Encoding.UTF8, leaveOpen: true);
        long expandedBytes = 0;
        byte[] buffer = new byte[81920];

        foreach (OfficeProvenanceZipWriteEntry entry in entries) {
            byte[] name = Encoding.UTF8.GetBytes(entry.Name);
            if (name.Length > ushort.MaxValue) throw new InvalidDataException("ZIP entry name exceeds the supported length.");
            byte[] localExtraField = RemoveExtraField(
                RemoveExtraField(entry.LocalExtraField, UnicodePathExtraFieldId),
                UnicodeCommentExtraFieldId);
            byte[] centralExtraField = RemoveExtraField(
                RemoveExtraField(entry.CentralExtraField, UnicodePathExtraFieldId),
                UnicodeCommentExtraFieldId);
            ushort method = entry.Compress ? DeflateMethod : StoredMethod;
            GetDosTimestamp(entry.LastWriteTime, out ushort dosDate, out ushort dosTime);
            uint localOffset = ToUInt32(output.Position, "local entry offset");
            WriteLocalHeader(writer, method, dosTime, dosDate, name, localExtraField);
            writer.Write(name);
            writer.Write(localExtraField);
            writer.Flush();
            long payloadStart = output.Position;
            uint crc = uint.MaxValue;
            long uncompressedLength = 0;

            using (Stream source = entry.Open()) {
                if (entry.Compress) {
                    using var compressor = new DeflateStream(output, CompressionLevel.Optimal, leaveOpen: true);
                    Copy(source, compressor, buffer, ref crc, ref uncompressedLength, ref expandedBytes, maximumExpandedBytes);
                } else {
                    Copy(source, output, buffer, ref crc, ref uncompressedLength, ref expandedBytes, maximumExpandedBytes);
                }
            }
            if (uncompressedLength != entry.ExpectedLength) throw new InvalidDataException("ZIP entry expanded to an unexpected length.");
            long compressedLength = output.Position - payloadStart;
            long entryEnd = output.Position;
            output.Position = localOffset + 14L;
            writer.Write(~crc);
            writer.Write(ToUInt32(compressedLength, "compressed entry size"));
            writer.Write(ToUInt32(uncompressedLength, "uncompressed entry size"));
            writer.Flush();
            output.Position = entryEnd;
            records.Add(new OfficeProvenanceZipRecord(
                name,
                method,
                dosTime,
                dosDate,
                ~crc,
                ToUInt32(compressedLength, "compressed entry size"),
                ToUInt32(uncompressedLength, "uncompressed entry size"),
                localOffset,
                entry.InternalAttributes,
                entry.ExternalAttributes,
                centralExtraField,
                entry.Comment));
        }

        uint centralOffset = ToUInt32(output.Position, "central directory offset");
        foreach (OfficeProvenanceZipRecord record in records) {
            WriteCentralHeader(writer, record);
            writer.Write(record.Name);
            writer.Write(record.CentralExtraField);
            writer.Write(record.Comment);
        }
        uint centralSize = ToUInt32(output.Position - centralOffset, "central directory size");
        if (records.Count >= ushort.MaxValue) {
            ulong zip64Offset = (ulong)output.Position;
            WriteZip64EndOfCentralDirectory(writer, (ulong)records.Count, centralSize, centralOffset);
            WriteZip64Locator(writer, zip64Offset);
            WriteEndOfCentralDirectory(writer, ushort.MaxValue, centralSize, centralOffset, archiveComment);
        } else {
            WriteEndOfCentralDirectory(writer, (ushort)records.Count, centralSize, centralOffset, archiveComment);
        }
        writer.Flush();
        return output.ToArray();
    }

    private static void Copy(
        Stream source,
        Stream destination,
        byte[] buffer,
        ref uint crc,
        ref long entryBytes,
        ref long expandedBytes,
        long maximumExpandedBytes) {
        while (true) {
            int read = source.Read(buffer, 0, buffer.Length);
            if (read <= 0) break;
            if (expandedBytes > maximumExpandedBytes - read) {
                throw new InvalidDataException("ZIP package exceeds the configured expanded-container limit.");
            }
            expandedBytes += read;
            entryBytes += read;
            for (int index = 0; index < read; index++) crc = CrcTable[(crc ^ buffer[index]) & 0xFF] ^ (crc >> 8);
            destination.Write(buffer, 0, read);
        }
    }

    private static void WriteLocalHeader(BinaryWriter writer, ushort method, ushort time, ushort date, byte[] name, byte[] extraField) {
        writer.Write(LocalHeaderSignature);
        writer.Write(Version20);
        writer.Write(Utf8FileNameFlag);
        writer.Write(method);
        writer.Write(time);
        writer.Write(date);
        writer.Write(0u);
        writer.Write(0u);
        writer.Write(0u);
        writer.Write((ushort)name.Length);
        writer.Write((ushort)extraField.Length);
    }

    private static void WriteCentralHeader(BinaryWriter writer, OfficeProvenanceZipRecord record) {
        writer.Write(CentralHeaderSignature);
        writer.Write((record.ExternalAttributes & 0xFFFF0000u) != 0 ? UnixVersion20 : Version20);
        writer.Write(Version20);
        writer.Write(Utf8FileNameFlag);
        writer.Write(record.Method);
        writer.Write(record.Time);
        writer.Write(record.Date);
        writer.Write(record.Crc);
        writer.Write(record.CompressedLength);
        writer.Write(record.UncompressedLength);
        writer.Write((ushort)record.Name.Length);
        writer.Write((ushort)record.CentralExtraField.Length);
        writer.Write((ushort)record.Comment.Length);
        writer.Write((ushort)0);
        writer.Write(record.InternalAttributes);
        writer.Write(record.ExternalAttributes);
        writer.Write(record.LocalOffset);
    }

    private static void WriteZip64EndOfCentralDirectory(BinaryWriter writer, ulong count, uint centralSize, uint centralOffset) {
        writer.Write(Zip64EndOfCentralDirectorySignature);
        writer.Write(44UL);
        writer.Write(Version45);
        writer.Write(Version45);
        writer.Write(0u);
        writer.Write(0u);
        writer.Write(count);
        writer.Write(count);
        writer.Write((ulong)centralSize);
        writer.Write((ulong)centralOffset);
    }

    private static void WriteZip64Locator(BinaryWriter writer, ulong zip64Offset) {
        writer.Write(Zip64LocatorSignature);
        writer.Write(0u);
        writer.Write(zip64Offset);
        writer.Write(1u);
    }

    private static void WriteEndOfCentralDirectory(BinaryWriter writer, ushort count, uint centralSize, uint centralOffset, byte[] archiveComment) {
        writer.Write(EndOfCentralDirectorySignature);
        writer.Write((ushort)0);
        writer.Write((ushort)0);
        writer.Write(count);
        writer.Write(count);
        writer.Write(centralSize);
        writer.Write(centralOffset);
        writer.Write((ushort)archiveComment.Length);
        writer.Write(archiveComment);
    }

    private static void GetDosTimestamp(DateTimeOffset value, out ushort date, out ushort time) {
        DateTime local = value.LocalDateTime;
        if (local.Year < 1980) local = new DateTime(1980, 1, 1, 0, 0, 0, DateTimeKind.Unspecified);
        if (local.Year > 2107) local = new DateTime(2107, 12, 31, 23, 59, 58, DateTimeKind.Unspecified);
        date = (ushort)(((local.Year - 1980) << 9) | (local.Month << 5) | local.Day);
        time = (ushort)((local.Hour << 11) | (local.Minute << 5) | (local.Second / 2));
    }

    private static uint ToUInt32(long value, string description) {
        if (value < 0 || value > uint.MaxValue) throw new InvalidDataException("ZIP " + description + " exceeds the supported limit.");
        return (uint)value;
    }

    private static uint[] CreateCrcTable() {
        var table = new uint[256];
        for (uint index = 0; index < table.Length; index++) {
            uint value = index;
            for (int bit = 0; bit < 8; bit++) value = (value & 1) != 0 ? 0xEDB88320u ^ (value >> 1) : value >> 1;
            table[index] = value;
        }
        return table;
    }

    private static byte[] RemoveExtraField(byte[] extraField, ushort fieldId) {
        int cursor = 0;
        using var output = new MemoryStream(extraField.Length);
        while (cursor <= extraField.Length - 4) {
            ushort currentId = (ushort)(extraField[cursor] | extraField[cursor + 1] << 8);
            int dataLength = extraField[cursor + 2] | extraField[cursor + 3] << 8;
            int fieldLength = 4 + dataLength;
            if (fieldLength > extraField.Length - cursor) break;
            if (currentId != fieldId) output.Write(extraField, cursor, fieldLength);
            cursor += fieldLength;
        }
        if (cursor < extraField.Length) output.Write(extraField, cursor, extraField.Length - cursor);
        return output.ToArray();
    }

    private sealed class OfficeProvenanceZipRecord {
        internal OfficeProvenanceZipRecord(byte[] name, ushort method, ushort time, ushort date, uint crc,
            uint compressedLength, uint uncompressedLength, uint localOffset, ushort internalAttributes, uint externalAttributes,
            byte[] centralExtraField, byte[] comment) {
            Name = name;
            Method = method;
            Time = time;
            Date = date;
            Crc = crc;
            CompressedLength = compressedLength;
            UncompressedLength = uncompressedLength;
            LocalOffset = localOffset;
            InternalAttributes = internalAttributes;
            ExternalAttributes = externalAttributes;
            CentralExtraField = centralExtraField;
            Comment = comment;
        }

        internal byte[] Name { get; }
        internal ushort Method { get; }
        internal ushort Time { get; }
        internal ushort Date { get; }
        internal uint Crc { get; }
        internal uint CompressedLength { get; }
        internal uint UncompressedLength { get; }
        internal uint LocalOffset { get; }
        internal ushort InternalAttributes { get; }
        internal uint ExternalAttributes { get; }
        internal byte[] CentralExtraField { get; }
        internal byte[] Comment { get; }
    }
}

internal sealed class OfficeProvenanceZipWriteEntry {
    private readonly Func<Stream> _open;

    internal OfficeProvenanceZipWriteEntry(
        string name,
        long expectedLength,
        bool compress,
        DateTimeOffset lastWriteTime,
        ushort internalAttributes,
        uint externalAttributes,
        byte[] localExtraField,
        byte[] centralExtraField,
        byte[] comment,
        Func<Stream> open) {
        Name = name ?? throw new ArgumentNullException(nameof(name));
        if (expectedLength < 0) throw new ArgumentOutOfRangeException(nameof(expectedLength));
        ExpectedLength = expectedLength;
        Compress = compress;
        LastWriteTime = lastWriteTime;
        InternalAttributes = internalAttributes;
        ExternalAttributes = externalAttributes;
        LocalExtraField = localExtraField ?? throw new ArgumentNullException(nameof(localExtraField));
        CentralExtraField = centralExtraField ?? throw new ArgumentNullException(nameof(centralExtraField));
        if (localExtraField.Length > ushort.MaxValue) throw new ArgumentOutOfRangeException(nameof(localExtraField));
        if (centralExtraField.Length > ushort.MaxValue) throw new ArgumentOutOfRangeException(nameof(centralExtraField));
        Comment = comment ?? throw new ArgumentNullException(nameof(comment));
        if (comment.Length > ushort.MaxValue) throw new ArgumentOutOfRangeException(nameof(comment));
        _open = open ?? throw new ArgumentNullException(nameof(open));
    }

    internal string Name { get; }
    internal long ExpectedLength { get; }
    internal bool Compress { get; }
    internal DateTimeOffset LastWriteTime { get; }
    internal ushort InternalAttributes { get; }
    internal uint ExternalAttributes { get; }
    internal byte[] LocalExtraField { get; }
    internal byte[] CentralExtraField { get; }
    internal byte[] Comment { get; }
    internal Stream Open() => _open();
}
